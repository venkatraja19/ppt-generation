import json
import logging
import os
import re
import traceback

from dotenv import load_dotenv
load_dotenv()

from groq import Groq
from flask import Flask, jsonify, request, send_file
from flask_cors import CORS

from ppt_generator import generate_ppt

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app)

_api_key = os.environ.get("GROQ_API_KEY", "").strip()
if not _api_key:
    raise RuntimeError("GROQ_API_KEY is missing. Add it to your .env file.")

client = Groq(api_key=_api_key)

SYSTEM_PROMPT = """You are an expert AI presentation designer and content strategist.
Your job is to deeply analyze the user's prompt and generate a fully structured, visually dynamic PowerPoint presentation.

OUTPUT FORMAT (STRICT JSON ONLY - a single object):
{
  "theme": {
    "title_color": "#HEX",
    "accent_color": "#HEX",
    "background_color": "#HEX",
    "text_color": "#HEX",
    "node_colors": ["#HEX", "#HEX", "#HEX"],
    "font": "Font Name",
    "style": "academic | corporate | medical | tech | creative | minimal"
  },
  "slides": [
    {
      "title": "Concise slide heading",
      "type": "content | image-left | image-right | flow | tree | multi-flow",
      "description": "2-3 sentence context for this slide",
      "bullets": ["Specific bullet point"],
      "image": "yes or no",
      "image_query": "precise visual search phrase (only when image=yes)",
      "layout": "image-left OR image-right (only when image=yes)",
      "structure_data": {}
    }
  ]
}

THEME RULES - analyze the topic and pick colors/fonts that match:
- Medical / Pharmacy / Biology -> deep red or navy title, white background, clinical feel
- Technology / AI / Software -> dark navy or indigo accent, modern sans-serif font
- Business / Finance / Corporate -> dark blue or charcoal, professional serif or sans
- Education / Academic -> deep maroon or forest green, clean and readable
- Creative / Design / Art -> vibrant accent colors, expressive palette
- Nature / Environment -> greens and earth tones
- If user specifies colors or style -> use exactly those
- node_colors: 3-5 hex colors for diagram nodes, matching the theme
- font: choose from Times New Roman, Calibri, Arial, Georgia, Trebuchet MS based on style

SLIDE RULES:
1. SLIDE COUNT:
- If user specifies -> follow exactly
- Default = 10 slides, Maximum = 80
- Decide depth intelligently based on topic complexity

2. HEADINGS:
- Extract meaningful, concise headings - NEVER use the raw prompt as a title
- Follow storytelling arc: Intro -> Core Content -> Conclusion

3. TEXT:
- description: 2-3 lines of context per slide
- bullets: 4-6 per slide, specific and informative (1-2 lines each)

4. IMAGES:
- Add images to 30-40% of slides where visuals genuinely add value
- image_query must be highly specific to the slide content
- Alternate image-left and image-right layouts
- For non-image slides: image=no, omit image_query

5. STRUCTURES:
- flow: structure_data = {"steps": ["Step 1", ...]}
- tree: structure_data = {"root": "Main", "children": [{"label": "Branch", "children": [...]}]}
- multi-flow: structure_data = {"flows": [{"label": "Flow 1", "steps": [...]}, ...]}
- Include at least ONE flow or multi-flow if topic supports it
- Non-structure slides: structure_data = {}

6. CONTENT QUALITY:
- Professional, clear English
- No repetition across slides
- Logical narrative flow

7. STRICT OUTPUT RULE:
- Output ONLY the JSON object above
- No markdown, no explanation, no code fences, no extra text"""


def _parse_hex(hex_str: str, fallback: tuple) -> tuple:
    try:
        h = hex_str.strip().lstrip("#")
        if len(h) == 6:
            return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    except (ValueError, AttributeError):
        pass
    return fallback


def _coerce_num_slides(value, default=10) -> int:
    try:
        n = int(value)
    except (TypeError, ValueError):
        return default
    return max(3, min(80, n))


def build_prompt(topic: str, user_requirements: str, num_slides: int) -> str:
    req_text = user_requirements if user_requirements else "None - use intelligent defaults based on topic"
    return f"""Topic: {topic}
User Requirements: {req_text}
Requested slide count: {num_slides}

Analyze the topic deeply. Choose a theme (colors, font, style) that visually matches this topic.
Generate exactly {num_slides} slides.
Return ONLY the JSON object with "theme" and "slides" keys."""


def extract_json(raw: str) -> dict:
    raw = raw.strip()
    if raw.startswith("```"):
        raw = re.sub(r"^```[a-z]*\n?", "", raw)
        raw = re.sub(r"\n?```$", "", raw.strip())
        raw = raw.strip()
    try:
        parsed = json.loads(raw)
        if isinstance(parsed, dict):
            return parsed
        if isinstance(parsed, list):
            return {"theme": {}, "slides": parsed}
    except json.JSONDecodeError:
        pass
    match = re.search(r"\{.*\}", raw, re.DOTALL)
    if match:
        try:
            parsed = json.loads(match.group())
            if isinstance(parsed, dict):
                return parsed
        except json.JSONDecodeError:
            pass
    match = re.search(r"\[.*\]", raw, re.DOTALL)
    if match:
        try:
            return {"theme": {}, "slides": json.loads(match.group())}
        except json.JSONDecodeError:
            pass
    raise ValueError("Could not extract valid JSON from model response.")


def safe_filename(text: str, max_len: int = 40) -> str:
    sanitized = re.sub(r"[^\w\s-]", "", text[:max_len])
    return sanitized.strip().replace(" ", "_") or "presentation"


@app.route("/")
def home():
    return jsonify({
        "message": "AI PPT Generator Backend Running",
        "llm": "Groq - LLaMA 3.3 70B",
        "endpoints": {
            "generate": "/generate (POST)",
            "preview":  "/preview (POST)",
            "health":   "/health (GET)"
        }
    })


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "llm": "groq"}), 200


@app.route("/preview", methods=["POST"])
def preview():
    data = request.get_json(silent=True)
    if not data or "topic" not in data:
        return jsonify({"error": "Missing 'topic'"}), 400

    topic = data.get("topic", "").strip()
    user_requirements = data.get("requirements", "").strip()
    num_slides = _coerce_num_slides(data.get("num_slides", 10))

    if not topic:
        return jsonify({"error": "'topic' cannot be empty"}), 400
    if len(topic) > 5000:
        return jsonify({"error": "'topic' must be under 5000 characters"}), 400

    try:
        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user",   "content": build_prompt(topic, user_requirements, num_slides)}
            ],
            temperature=0.7,
            max_tokens=4096,
        )
        raw_text = response.choices[0].message.content
        result = extract_json(raw_text)
        slides = result.get("slides", [])
        if not isinstance(slides, list) or not slides:
            return jsonify({"error": "AI returned unexpected structure"}), 500
        return jsonify({"slides": slides, "theme": result.get("theme", {}), "count": len(slides)})
    except ValueError as e:
        logger.exception("Preview JSON error")
        return jsonify({"error": f"JSON parse error: {str(e)}"}), 500
    except Exception as e:
        logger.exception("Preview error")
        return jsonify({"error": f"Server error: {str(e)}"}), 500


@app.route("/generate", methods=["POST"])
def generate():
    data = request.get_json(silent=True)

    if not data or "topic" not in data:
        return jsonify({"error": "Missing 'topic' in request body"}), 400

    topic = data.get("topic", "").strip()
    user_requirements = data.get("requirements", "").strip()
    num_slides = _coerce_num_slides(data.get("num_slides", 10))

    logger.info("Generate request - topic_len=%d num_slides=%d", len(topic), num_slides)

    if not topic:
        return jsonify({"error": "'topic' cannot be empty"}), 400
    if len(topic) > 5000:
        return jsonify({"error": "'topic' must be under 5000 characters"}), 400

    try:
        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user",   "content": build_prompt(topic, user_requirements, num_slides)}
            ],
            temperature=0.7,
            max_tokens=4096,
        )

        raw_text = response.choices[0].message.content
        result = extract_json(raw_text)
        slides = result.get("slides", [])
        theme  = result.get("theme", {})

        if not isinstance(slides, list) or not slides:
            return jsonify({"error": "AI returned unexpected structure"}), 500

        os.makedirs("outputs", exist_ok=True)

        cover_title = slides[0].get("title") or topic.split("\n")[0].strip()[:60]
        filename = safe_filename(cover_title)
        output_path = os.path.join("outputs", f"{filename}.pptx")

        title_color  = _parse_hex(theme.get("title_color",  ""), (0x1a, 0x1a, 0x5e))
        accent_color = _parse_hex(theme.get("accent_color", ""), (0xED, 0x7D, 0x31))
        bg_color     = _parse_hex(theme.get("background_color", ""), (0xFF, 0xFF, 0xFF))
        text_color   = _parse_hex(theme.get("text_color",   ""), (0x33, 0x33, 0x33))
        font         = theme.get("font", "Calibri") or "Calibri"
        node_colors  = [
            _parse_hex(c, None)
            for c in theme.get("node_colors", [])
            if _parse_hex(c, None) is not None
        ] or None

        filepath = generate_ppt(
            slides=slides,
            topic=cover_title,
            output_path=output_path,
            title_color=title_color,
            accent_color=accent_color,
            bg_color=bg_color,
            text_color=text_color,
            font=font,
            node_colors=node_colors,
        )

        return send_file(
            filepath,
            as_attachment=True,
            download_name=f"{filename}.pptx",
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

    except ValueError as e:
        logger.exception("Generate JSON error")
        return jsonify({"error": f"JSON parse error: {str(e)}"}), 500
    except json.JSONDecodeError:
        logger.exception("Generate JSON decode error")
        return jsonify({"error": "AI returned invalid JSON. Please try again."}), 500
    except Exception as e:
        logger.exception("Generate error")
        return jsonify({"error": f"Server error: {str(e)}"}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_DEBUG", "false").lower() == "true"
    app.run(host="0.0.0.0", port=port, debug=debug)

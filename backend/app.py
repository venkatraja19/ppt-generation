import json
import logging
import os
import re

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
client = Groq(api_key=_api_key) if _api_key else None

SYSTEM_PROMPT = """You are an advanced AI presentation generator.
Your job is to create a fully structured, professional PowerPoint based on the user's full prompt.

OUTPUT FORMAT (STRICT JSON ONLY):
[
  {
    "title": "Concise slide heading (not the raw user prompt)",
    "type": "content | image-left | image-right | flow | tree | multi-flow",
    "description": "2-3 sentence context or definition for this slide",
    "bullets": [
      "Specific, informative bullet point"
    ],
    "image": "yes or no",
    "image_query": "short descriptive phrase for image search (only when image=yes)",
    "layout": "image-left OR image-right (only when image=yes)",
    "structure_data": {}
  }
]

RULES:
1. SLIDE COUNT:
- Decide intelligently based on content depth
- If user specifies → follow exactly
- Default = 10 slides, Maximum = 80

2. HEADINGS:
- Extract meaningful, concise headings from the user's prompt
- NEVER use the raw user prompt as a slide title
- Follow storytelling: Intro → Core Content → Conclusion

3. TEXT:
- description: 2-3 lines of context per slide
- bullets: 4-6 per slide, each 1-2 lines, specific and informative

4. IMAGES:
- Automatically add images to 30-40% of slides where visuals add value
- Set image=yes and provide a short, specific image_query
- Alternate between image-left and image-right layouts
- For non-image slides set image=no and omit image_query

5. STRUCTURES (use where topic supports it):
- flow: structure_data = {"steps": ["Step 1", ...]}
- tree: structure_data = {"root": "Main", "children": [{"label": "Branch", "children": [...]}]}
- multi-flow: structure_data = {"flows": [{"label": "Flow 1", "steps": [...]}, ...]}
- Include at least ONE flow or multi-flow slide if topic supports it
- For non-structure slides set structure_data = {}

6. CONTENT QUALITY:
- Clear, professional English
- No repetition across slides
- Logical narrative flow

7. STRICT RULE:
- Output ONLY a valid JSON array
- No markdown, no explanation, no code fences, no extra text"""


def build_prompt(topic: str, user_requirements: str, num_slides: int) -> str:
    return f"""Topic: {topic}
User Requirements: {user_requirements if user_requirements else "None — use intelligent defaults"}
Requested slide count: {num_slides}

Generate exactly {num_slides} slides following the system rules.
Return ONLY the JSON array."""


def extract_json(raw: str) -> list:
    raw = raw.strip()
    if raw.startswith("```"):
        raw = re.sub(r"^```[a-z]*\n?", "", raw)
        raw = re.sub(r"\n?```$", "", raw.strip())
        raw = raw.strip()
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        pass
    match = re.search(r"\[.*\]", raw, re.DOTALL)
    if match:
        try:
            return json.loads(match.group())
        except json.JSONDecodeError:
            pass
    raise ValueError("Could not extract valid JSON array from model response.")


def safe_filename(text: str, max_len: int = 40) -> str:
    sanitized = re.sub(r"[^\w\s-]", "", text[:max_len])
    return sanitized.strip().replace(" ", "_") or "presentation"


@app.route("/")
def home():
    return jsonify({
        "message": "AI PPT Generator Backend Running",
        "llm": "Groq — LLaMA 3.3 70B",
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
    num_slides = data.get("num_slides", 10)

    if not topic:
        return jsonify({"error": "'topic' cannot be empty"}), 400
    if len(topic) > 2000:
        return jsonify({"error": "'topic' must be under 2000 characters"}), 400
    if not isinstance(num_slides, int) or not (3 <= num_slides <= 80):
        return jsonify({"error": "'num_slides' must be between 3 and 80"}), 400

    try:
        if not client:
            return jsonify({"error": "GROQ_API_KEY is not configured on the server."}), 503
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
        slides = extract_json(raw_text)
        if not isinstance(slides, list) or not slides:
            return jsonify({"error": "AI returned unexpected structure"}), 500
        return jsonify({"slides": slides, "count": len(slides)})
    except ValueError as e:
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
    num_slides = data.get("num_slides", 10)

    if not topic:
        return jsonify({"error": "'topic' cannot be empty"}), 400
    if len(topic) > 2000:
        return jsonify({"error": "'topic' must be under 2000 characters"}), 400
    if not isinstance(num_slides, int) or not (3 <= num_slides <= 80):
        return jsonify({"error": "'num_slides' must be between 3 and 80"}), 400

    try:
        if not client:
            return jsonify({"error": "GROQ_API_KEY is not configured on the server."}), 503
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
        slides = extract_json(raw_text)

        if not isinstance(slides, list) or not slides:
            return jsonify({"error": "AI returned unexpected structure"}), 500

        os.makedirs("outputs", exist_ok=True)

        cover_title = slides[0].get("title") or topic.split("\n")[0].strip()[:60]
        filename = safe_filename(cover_title)
        output_path = os.path.join("outputs", f"{filename}.pptx")

        filepath = generate_ppt(
            slides=slides,
            topic=cover_title,
            output_path=output_path
        )

        return send_file(
            filepath,
            as_attachment=True,
            download_name=f"{filename}.pptx",
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

    except ValueError as e:
        return jsonify({"error": f"JSON parse error: {str(e)}"}), 500
    except json.JSONDecodeError:
        return jsonify({"error": "AI returned invalid JSON. Please try again."}), 500
    except Exception as e:
        logger.exception("Generate error")
        return jsonify({"error": f"Server error: {str(e)}"}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_DEBUG", "false").lower() == "true"
    app.run(host="0.0.0.0", port=port, debug=debug)

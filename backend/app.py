import json
import os
import re

from dotenv import load_dotenv
load_dotenv()

from groq import Groq
from flask import Flask, jsonify, request, send_file
from flask_cors import CORS

from ppt_generator import generate_ppt

app = Flask(__name__)
CORS(app)

_api_key = os.environ.get("GROQ_API_KEY", "").strip()
if not _api_key:
    raise RuntimeError("GROQ_API_KEY is missing. Add it to your .env file.")

client = Groq(api_key=_api_key)

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
- Set image=yes and provide a short, specific image_query (e.g. "solar panels on rooftop")
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
    """Robustly extract JSON array from model output."""
    raw = raw.strip()

    # Strip markdown fences
    if raw.startswith("```"):
        raw = re.sub(r"^```[a-z]*\n?", "", raw)
        raw = re.sub(r"\n?```$", "", raw.strip())
        raw = raw.strip()

    # Try direct parse
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        pass

    # Find first [ ... ] block
    match = re.search(r"\[.*\]", raw, re.DOTALL)
    if match:
        try:
            return json.loads(match.group())
        except json.JSONDecodeError:
            pass

    raise ValueError("Could not extract valid JSON array from model response.")


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
    """Returns slide JSON for frontend preview without generating PPTX."""
    data = request.get_json()
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

        if not isinstance(slides, list) or len(slides) == 0:
            return jsonify({"error": "AI returned unexpected structure"}), 500

        return jsonify({"slides": slides, "count": len(slides)})

    except ValueError as e:
        return jsonify({"error": f"JSON parse error: {str(e)}"}), 500
    except Exception as e:
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

        if not isinstance(slides, list) or len(slides) == 0:
            return jsonify({"error": "AI returned unexpected structure"}), 500

        os.makedirs("outputs", exist_ok=True)

        # Use only the first slide's title as the cover heading, not the raw prompt
        cover_title = slides[0].get("title") or topic.split("\n")[0].strip()[:60]

        safe = topic[:40].replace(' ', '_')
        filepath = generate_ppt(
            slides=slides,
            topic=cover_title,
            output_path=f"outputs/{safe}.pptx"
        )

        return send_file(
            filepath,
            as_attachment=True,
            download_name=f"{topic[:40].replace(' ', '_')}.pptx",
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

    except ValueError as e:
        return jsonify({"error": f"JSON parse error: {str(e)}"}), 500
    except json.JSONDecodeError:
        return jsonify({"error": "AI returned invalid JSON. Please try again."}), 500
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({"error": f"Server error: {str(e)}"}), 500


if __name__ == "__main__":
    print("GROQ_API_KEY loaded:", bool(_api_key))
    app.run(debug=True, port=5000)




# import json
# import os

# from dotenv import load_dotenv
# load_dotenv()

# from groq import Groq
# from flask import Flask, jsonify, request, send_file
# from flask_cors import CORS

# from ppt_generator import generate_ppt

# # ---------------- APP SETUP ----------------
# app = Flask(__name__)
# CORS(app)

# _api_key = os.environ.get("GROQ_API_KEY", "").strip()
# if not _api_key:
#     raise RuntimeError("GROQ_API_KEY is missing. Add it to your .env file.")

# client = Groq(api_key=_api_key)

# # ---------------- SYSTEM PROMPT ----------------
# SYSTEM_PROMPT = """You are an expert presentation designer.
# Your job is to generate structured PowerPoint slide content.

# Rules:
# - Return ONLY valid JSON — no markdown, no explanation, no code fences
# - Each slide must have a "heading" (max 8 words) and "bullets" (3-5 points)
# - Bullets must be concise and informative (max 15 words each)
# - Structure: intro → core content → conclusion"""

# # ---------------- PROMPT BUILDER ----------------
# def build_prompt(topic: str, num_slides: int) -> str:
#     return f"""Create a {num_slides}-slide presentation on: "{topic}"

# Return ONLY this JSON format:
# {{
#   "slides": [
#     {{
#       "heading": "Slide title here",
#       "bullets": [
#         "First key point",
#         "Second key point",
#         "Third key point"
#       ]
#     }}
#   ]
# }}

# Generate exactly {num_slides} slides with a logical flow."""

# # ---------------- ROOT ROUTE ----------------
# @app.route("/")
# def home():
#     return jsonify({
#         "message": "AI PPT Generator Backend Running 🚀",
#         "llm": "Groq — LLaMA 3.3 70B (Free)",
#         "endpoints": {
#             "generate": "/generate (POST)",
#             "health":   "/health (GET)"
#         }
#     })

# # ---------------- HEALTH ----------------
# @app.route("/health", methods=["GET"])
# def health():
#     return jsonify({"status": "ok", "llm": "groq"}), 200

# # ---------------- GENERATE ROUTE ----------------
# @app.route("/generate", methods=["POST"])
# def generate():
#     data = request.get_json()

#     if not data or "topic" not in data:
#         return jsonify({"error": "Missing 'topic' in request body"}), 400

#     topic = data.get("topic", "").strip()
#     if not topic:
#         return jsonify({"error": "'topic' cannot be empty"}), 400
#     if len(topic) > 300:
#         return jsonify({"error": "'topic' must be under 300 characters"}), 400

#     num_slides = data.get("num_slides", 6)
#     if not isinstance(num_slides, int) or not (3 <= num_slides <= 12):
#         return jsonify({"error": "'num_slides' must be between 3 and 12"}), 400

#     try:
#         # ---------------- GROQ API CALL ----------------
#         response = client.chat.completions.create(
#             model="llama-3.3-70b-versatile",
#             messages=[
#                 {"role": "system", "content": SYSTEM_PROMPT},
#                 {"role": "user",   "content": build_prompt(topic, num_slides)}
#             ],
#             temperature=0.7,
#             max_tokens=2048,
#         )

#         raw_text = response.choices[0].message.content.strip()

#         # Strip markdown fences if model wraps output
#         if raw_text.startswith("```"):
#             raw_text = raw_text.split("```")[1]
#             if raw_text.startswith("json"):
#                 raw_text = raw_text[4:]
#             raw_text = raw_text.strip()

#         slides_data = json.loads(raw_text)

#         if "slides" not in slides_data or not isinstance(slides_data["slides"], list):
#             return jsonify({"error": "Unexpected response from AI"}), 500

#         # ---------------- GENERATE PPTX ----------------
#         os.makedirs("outputs", exist_ok=True)

#         filepath = generate_pptx(
#             {
#                 "topic":  topic,
#                 "slides": slides_data["slides"]
#             },
#             output_dir="outputs"
#         )

#         # ---------------- RETURN FILE ----------------
#         return send_file(
#             filepath,
#             as_attachment=True,
#             download_name=f"{topic[:40]}.pptx",
#             mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation"
#         )

#     except json.JSONDecodeError:
#         return jsonify({"error": "AI returned invalid JSON. Please try again."}), 500

#     except Exception as e:
#         return jsonify({"error": f"Server error: {str(e)}"}), 500


# # ---------------- RUN ----------------
# if __name__ == "__main__":
#     print("✅ GROQ_API_KEY loaded:", bool(_api_key))
#     app.run(debug=True, port=5000)
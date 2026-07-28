import io
import json

from flask import Flask, jsonify, request, send_file
from flask_cors import CORS
from pptx import Presentation

app = Flask(__name__)
CORS(app)


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"}), 200


@app.route("/inject", methods=["POST"])
def inject_text():
    try:
        if "template" not in request.files or "data" not in request.form:
            return jsonify({"error": "Missing template file or data JSON"}), 400

        template_file = request.files["template"]
        data_json = json.loads(request.form["data"])
        slides_data = data_json.get("slides", [])

        prs = Presentation(template_file)

        for i, slide in enumerate(prs.slides):
            if i >= len(slides_data):
                break

            slide_info = slides_data[i]
            title_text = slide_info.get("title", "")
            bullets = slide_info.get("bullets", [])

            for shape in slide.shapes:
                if not shape.has_text_frame:
                    continue

                text_frame = shape.text_frame

                if shape == slide.shapes.title or "title" in shape.name.lower():
                    if title_text:
                        text_frame.text = title_text
                else:
                    if bullets:
                        text_frame.clear()
                        for bullet in bullets:
                            p = text_frame.add_paragraph()
                            p.text = bullet

        output = io.BytesIO()
        prs.save(output)
        output.seek(0)

        return send_file(
            output,
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name="presentation_updated.pptx",
        )

    except Exception as e:  # noqa: BLE001
        print(f"Error processing PPTX: {e!s}")
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)

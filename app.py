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

                # 1. Αντικατάσταση ΜΟΝΟ του κειμένου στον Τίτλο
                if shape == slide.shapes.title or "title" in shape.name.lower():
                    if title_text and len(text_frame.paragraphs) > 0:
                        p = text_frame.paragraphs[0]
                        if len(p.runs) > 0:
                            p.runs[0].text = title_text
                            for r in p.runs[1:]:
                                r.text = ""
                        else:
                            p.text = title_text

                # 2. Αντικατάσταση ΜΟΝΟ του κειμένου στα Bullets
                else:
                    if bullets:
                        for p_idx, paragraph in enumerate(text_frame.paragraphs):
                            if p_idx < len(bullets):
                                if len(paragraph.runs) > 0:
                                    paragraph.runs[0].text = bullets[p_idx]
                                    for r in paragraph.runs[1:]:
                                        r.text = ""
                                else:
                                    paragraph.text = bullets[p_idx]
                            else:
                                paragraph.text = ""

                        if len(bullets) > len(text_frame.paragraphs):
                            for b_idx in range(
                                len(text_frame.paragraphs), len(bullets)
                            ):
                                p = text_frame.add_paragraph()
                                p.text = bullets[b_idx]

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

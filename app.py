import io
import json
from flask import Flask, jsonify, request, send_file
from flask_cors import CORS
from pptx import Presentation

app = Flask(__name__)
CORS(app)


def delete_slide(prs, index):
    """Διαγράφει μια διαφάνεια από το Presentation βάσει του index της."""
    rId = prs.slides._sldIdLst[index].rId
    prs.part.drop_rel(rId)
    del prs.slides._sldIdLst[index]


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
        num_template_slides = len(prs.slides)
        num_gemini_slides = len(slides_data)

        # 1. ΕΝΗΜΕΡΩΣΗ ΥΠΑΡΧΟΥΣΩΝ ΔΙΑΦΑΝΕΙΩΝ
        for i in range(min(num_template_slides, num_gemini_slides)):
            slide = prs.slides[i]
            slide_info = slides_data[i]
            title_text = slide_info.get("title", "")
            bullets = slide_info.get("bullets", [])

            for shape in slide.shapes:
                if not shape.has_text_frame:
                    continue

                text_frame = shape.text_frame

                # Αλλαγή Τίτλου
                if shape == slide.shapes.title or "title" in shape.name.lower():
                    if title_text and len(text_frame.paragraphs) > 0:
                        p = text_frame.paragraphs[0]
                        if len(p.runs) > 0:
                            p.runs[0].text = title_text
                            for r in p.runs[1:]:
                                r.text = ""
                        else:
                            p.text = title_text

                # Αλλαγή Bullets
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

                        # Αν το Gemini έχει περισσότερα bullets, τα προσθέτουμε
                        if len(bullets) > len(text_frame.paragraphs):
                            for b_idx in range(
                                len(text_frame.paragraphs), len(bullets)
                            ):
                                p = text_frame.add_paragraph()
                                p.text = bullets[b_idx]

        # 2. ΑΝ ΤΟ TEMPLATE ΕΧΕΙ ΠΕΡΙΣΣΟΤΕΡΕΣ ΔΙΑΦΑΝΕΙΕΣ -> ΔΙΑΓΡΑΦΟΥΜΕ ΤΙΣ ΕΠΙΠΛΕΟΝ
        if num_template_slides > num_gemini_slides:
            for i in range(num_template_slides - 1, num_gemini_slides - 1, -1):
                delete_slide(prs, i)

        # 3. ΑΝ ΤΟ GEMINI ΕΧΕΙ ΠΕΡΙΣΣΟΤΕΡΕΣ ΔΙΑΦΑΝΕΙΕΣ -> ΔΗΜΙΟΥΡΓΟΥΜΕ ΝΕΕΣ
        elif num_gemini_slides > num_template_slides:
            default_layout = (
                prs.slides[-1].slide_layout
                if len(prs.slides) > 0
                else prs.slide_layouts[1]
            )

            for i in range(num_template_slides, num_gemini_slides):
                slide = prs.slides.add_slide(default_layout)
                slide_info = slides_data[i]
                title_text = slide_info.get("title", "")
                bullets = slide_info.get("bullets", [])

                for shape in slide.shapes:
                    if not shape.has_text_frame:
                        continue

                    text_frame = shape.text_frame

                    if shape == slide.shapes.title or "title" in shape.name.lower():
                        text_frame.text = title_text
                    else:
                        text_frame.text = ""
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

import io
import json
from flask import Flask, jsonify, request, send_file
from flask_cors import CORS
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER

app = Flask(__name__)
CORS(app)


def delete_slide(prs, index):
    """Διαγράφει μια διαφάνεια από το Presentation βάσει του index της."""
    rId = prs.slides._sldIdLst[index].rId
    prs.part.drop_rel(rId)
    del prs.slides._sldIdLst[index]


def process_slide_text(slide, title_text, bullets):
    """Εντοπίζει αυστηρά μόνο τον Τίτλο και το Κύριο Σώμα (Body) και αγνοεί footers/logos."""
    title_shape = None
    body_shape = None

    # 1. Πρώτο πέρασμα: Αναγνώριση μέσω Placeholders
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue

        if shape.is_placeholder:
            ph_type = shape.placeholder_format.type

            # ΑΓΝΟΟΥΜΕ ΥΠΟ his/FOOTERS/LOGOS/DATE/SLIDE NUMBERS
            if ph_type in [
                PP_PLACEHOLDER.FOOTER,
                PP_PLACEHOLDER.SLIDE_NUMBER,
                PP_PLACEHOLDER.HEADER,
                PP_PLACEHOLDER.DATE,
            ]:
                continue

            if ph_type in [PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE]:
                title_shape = shape
            elif ph_type in [
                PP_PLACEHOLDER.BODY,
                PP_PLACEHOLDER.SUBTITLE,
                PP_PLACEHOLDER.OBJECT,
            ]:
                body_shape = shape

    # 2. Δεύτερο πέρασμα (Fallback): Αν το template δεν χρησιμοποιεί standard placeholders
    if not title_shape or not body_shape:
        valid_text_shapes = []
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue

            # Φιλτράρουμε πολύ μικρά κουτιά ή κουτιά πολύ χαμηλά στη σελίδα (footers)
            # 1 inch = 914400 EMUs -> Y > 6.5 ίντσες είναι συνήθως footer
            if shape.top > 5943600:  # ~6.5 ίντσες
                continue

            valid_text_shapes.append(shape)

        # Ταξινομούμε κατά κατακόρυφη θέση (Top)
        valid_text_shapes.sort(key=lambda s: s.top)

        if valid_text_shapes:
            if not title_shape:
                title_shape = valid_text_shapes[0]
            if not body_shape and len(valid_text_shapes) > 1:
                body_shape = valid_text_shapes[1]

    # --- ΕΦΑΡΜΟΓΗ ΚΕΙΜΕΝΟΥ ΣΤΟΝ ΤΙΤΛΟ ---
    if title_shape and title_shape.has_text_frame:
        tf = title_shape.text_frame
        if title_text and len(tf.paragraphs) > 0:
            p = tf.paragraphs[0]
            if len(p.runs) > 0:
                p.runs[0].text = title_text
                for r in p.runs[1:]:
                    r.text = ""
            else:
                p.text = title_text

    # --- ΕΦΑΡΜΟΓΗ ΚΕΙΜΕΝΟΥ ΣΤΑ BULLETS ---
    if body_shape and body_shape.has_text_frame and bullets:
        tf = body_shape.text_frame

        for p_idx, paragraph in enumerate(tf.paragraphs):
            if p_idx < len(bullets):
                if len(paragraph.runs) > 0:
                    paragraph.runs[0].text = bullets[p_idx]
                    for r in paragraph.runs[1:]:
                        r.text = ""
                else:
                    paragraph.text = bullets[p_idx]
            else:
                paragraph.text = ""

        # Αν το Gemini έχει περισσότερα bullets, προσθέτουμε παραγράφους
        if len(bullets) > len(tf.paragraphs):
            for b_idx in range(len(tf.paragraphs), len(bullets)):
                p = tf.add_paragraph()
                p.text = bullets[b_idx]


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
            process_slide_text(
                slide,
                slide_info.get("title", ""),
                slide_info.get("bullets", []),
            )

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
                process_slide_text(
                    slide,
                    slide_info.get("title", ""),
                    slide_info.get("bullets", []),
                )

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

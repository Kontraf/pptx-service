import io
import json
from flask import Flask, jsonify, request, send_file
from flask_cors import CORS
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER

app = Flask(__name__)
CORS(app)


def delete_slide(prs, index):
    """Ασφαλής διαγραφή διαφάνειας από το Presentation χωρίς σφάλματα XML."""
    try:
        rId = prs.slides._sldIdLst[index].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[index]
    except Exception:
        try:
            # Fallback μέθοδος αν το _sldIdLst συμπεριφέρεται ως απλή λίστα
            slide = prs.slides[index]
            for rel in list(prs.part.rels.values()):
                if rel.target_part == slide.part:
                    prs.part.drop_rel(rel.rId)
                    break
            del prs.slides._sldIdLst[index]
        except Exception as e:
            print(f"Error deleting slide at index {index}: {e}")


def get_best_content_layout(prs):
    """Βρίσκει αυτόματα το layout της κύριας διαφάνειας (Main Content Slide)"""
    if len(prs.slides) > 1:
        for slide in prs.slides[1:]:
            has_title = False
            has_body = False
            for shape in slide.shapes:
                if shape.is_placeholder:
                    ph_type = shape.placeholder_format.type
                    if ph_type in [PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE]:
                        has_title = True
                    elif ph_type in [PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT]:
                        has_body = True
            if has_title and has_body:
                return slide.slide_layout

    for layout in prs.slide_layouts:
        has_title = False
        has_body = False
        for ph in layout.placeholders:
            ph_type = ph.placeholder_format.type
            if ph_type == PP_PLACEHOLDER.TITLE:
                has_title = True
            elif ph_type in [PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT]:
                has_body = True
        if has_title and has_body:
            return layout

    if len(prs.slide_layouts) > 1:
        return prs.slide_layouts[1]
    return prs.slide_layouts[0]


def process_cover_slide(slide, deck_title, deck_subtitle=""):
    """Ειδικός χειρισμός Cover Slide: Βάζει μόνο τίτλο & υπότιτλο και καθαρίζει τα υπόλοιπα."""
    title_done = False

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue

        tf = shape.text_frame

        if shape.is_placeholder:
            ph_type = shape.placeholder_format.type

            if (
                ph_type in [PP_PLACEHOLDER.CENTER_TITLE, PP_PLACEHOLDER.TITLE]
                and not title_done
            ):
                if len(tf.paragraphs) > 0:
                    p = tf.paragraphs[0]
                    if len(p.runs) > 0:
                        p.runs[0].text = deck_title
                        for r in p.runs[1:]:
                            r.text = ""
                    else:
                        p.text = deck_title
                title_done = True

            elif ph_type == PP_PLACEHOLDER.SUBTITLE:
                if deck_subtitle and len(tf.paragraphs) > 0:
                    p = tf.paragraphs[0]
                    if len(p.runs) > 0:
                        p.runs[0].text = deck_subtitle
                        for r in p.runs[1:]:
                            r.text = ""
                    else:
                        p.text = deck_subtitle
                else:
                    tf.text = ""

            elif ph_type not in [
                PP_PLACEHOLDER.FOOTER,
                PP_PLACEHOLDER.HEADER,
                PP_PLACEHOLDER.SLIDE_NUMBER,
                PP_PLACEHOLDER.DATE,
            ]:
                tf.text = ""

        else:
            if not title_done and shape.top < 4000000:
                tf.text = deck_title
                title_done = True
            else:
                tf.text = ""


def process_slide_text(slide, title_text, bullets):
    """Εντοπίζει αυστηρά μόνο τον Τίτλο και το Κύριο Σώμα (Body) στις διαφάνειες περιεχομένου."""
    title_shape = None
    body_shape = None

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue

        if shape.is_placeholder:
            ph_type = shape.placeholder_format.type

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

    if not title_shape or not body_shape:
        valid_text_shapes = []
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue

            if shape.top > 5943600:
                continue

            valid_text_shapes.append(shape)

        valid_text_shapes.sort(key=lambda s: s.top)

        if valid_text_shapes:
            if not title_shape:
                title_shape = valid_text_shapes[0]
            if not body_shape and len(valid_text_shapes) > 1:
                body_shape = valid_text_shapes[1]

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

        deck_title = data_json.get("title", "")
        deck_subtitle = data_json.get("subtitle", "")
        slides_data = data_json.get("slides", [])

        prs = Presentation(template_file)
        num_template_slides = len(prs.slides)
        num_gemini_slides = len(slides_data)

        content_layout = get_best_content_layout(prs)

        # 1. ΧΕΙΡΙΣΜΟΣ COVER SLIDE
        if num_template_slides > 0:
            process_cover_slide(prs.slides[0], deck_title, deck_subtitle)

        # 2. ΕΝΗΜΕΡΩΣΗ ΥΠΟΛΟΙΠΩΝ ΔΙΑΦΑΝΕΙΩΝ
        for i in range(1, min(num_template_slides, num_gemini_slides + 1)):
            slide = prs.slides[i]
            slide_info = slides_data[i - 1] if (i - 1) < len(slides_data) else {}

            process_slide_text(
                slide,
                slide_info.get("title", ""),
                slide_info.get("bullets", []),
            )

        # 3. ΔΙΑΓΡΑΦΗ ΠΕΡΙΣΣΕΥΟΥΜΕΝΩΝ ΔΙΑΦΑΝΕΙΩΝ
        total_target_slides = len(slides_data) + 1
        if num_template_slides > total_target_slides:
            for i in range(num_template_slides - 1, total_target_slides - 1, -1):
                delete_slide(prs, i)

        # 4. ΔΗΜΙΟΥΡΓΙΑ ΕΠΙΠΛΕΟΝ ΔΙΑΦΑΝΕΙΩΝ
        elif total_target_slides > num_template_slides:
            for i in range(num_template_slides, total_target_slides):
                slide = prs.slides.add_slide(content_layout)
                slide_info = slides_data[i - 1]
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

    except Exception as e:
        print(f"Error processing PPTX: {str(e)}")
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)

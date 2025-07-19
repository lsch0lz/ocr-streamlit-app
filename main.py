import streamlit as st
import tempfile
import os
from pathlib import Path
import io
from PIL import Image
import fitz  # PyMuPDF
from docx import Document
import easyocr
import numpy as np


def setup_page():
    """Configure the Streamlit page"""
    st.set_page_config(
        page_title="PDF OCR to DOCX Converter",
        page_icon="📄",
        layout="wide"
    )


@st.cache_resource
def load_ocr_reader():
    """Load EasyOCR reader (cached to avoid reloading)"""
    try:
        reader = easyocr.Reader(['en', 'de'])  # English and German support
        return reader
    except Exception as e:
        st.error(f"Error loading OCR reader: {str(e)}")
        return None


def pdf_to_images(pdf_bytes):
    """Convert PDF pages to images using PyMuPDF"""
    try:
        images = []
        # Open PDF from bytes
        pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")

        for page_num in range(pdf_document.page_count):
            # Get page
            page = pdf_document.load_page(page_num)

            # Convert to image (higher resolution for better OCR)
            mat = fitz.Matrix(2.0, 2.0)  # 2x zoom for better quality
            pix = page.get_pixmap(matrix=mat)

            # Convert to PIL Image
            img_data = pix.tobytes("ppm")
            img = Image.open(io.BytesIO(img_data))
            images.append(img)

        pdf_document.close()
        return images

    except Exception as e:
        st.error(f"Error converting PDF to images: {str(e)}")
        return None


def perform_ocr_easyocr(images, reader):
    """Perform OCR using EasyOCR"""
    extracted_text = ""

    progress_bar = st.progress(0)
    status_text = st.empty()

    for i, image in enumerate(images):
        try:
            # Update progress
            progress = (i + 1) / len(images)
            progress_bar.progress(progress)
            status_text.text(f"Verarbeite Seite {i + 1} von {len(images)}...")

            # Convert PIL image to numpy array
            img_array = np.array(image)

            # Perform OCR
            results = reader.readtext(img_array, paragraph=True)

            # Extract text from results
            page_text = ""
            for detection in results:
                text = detection[1]  # detection[0] is bbox, detection[1] is text, detection[2] is confidence
                page_text += text + " "

            extracted_text += f"\n--- Seite {i + 1} ---\n"
            extracted_text += page_text.strip()
            extracted_text += "\n\n"

        except Exception as e:
            st.warning(f"Error performing OCR on page {i + 1}: {str(e)}")
            continue

    progress_bar.empty()
    status_text.empty()

    return extracted_text


def create_docx(text, filename):
    """Create a DOCX document from extracted text"""
    try:
        doc = Document()

        # Add title
        title = doc.add_heading(f'OCR Result: {filename}', 0)

        # Add the extracted text
        # Split text into paragraphs for better formatting
        paragraphs = text.split('\n\n')

        for paragraph in paragraphs:
            if paragraph.strip():
                # Handle page separators
                if paragraph.strip().startswith('--- Seite'):
                    doc.add_heading(paragraph.strip(), level=1)
                else:
                    doc.add_paragraph(paragraph.strip())

        return doc
    except Exception as e:
        st.error(f"Error creating DOCX: {str(e)}")
        return None


def create_download_link(doc, filename):
    """Create a download link for the DOCX document"""
    try:
        # Save document to bytes
        doc_io = io.BytesIO()
        doc.save(doc_io)
        doc_io.seek(0)

        # Create download button
        st.download_button(
            label=f"📥 Download {filename}",
            data=doc_io.getvalue(),
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=f"download_{filename}"
        )

        return True
    except Exception as e:
        st.error(f"Error creating download link: {str(e)}")
        return False


def process_single_pdf(uploaded_file, reader):
    """Process a single PDF file"""
    try:
        # Get PDF bytes
        pdf_bytes = uploaded_file.getvalue()

        # Convert PDF to images
        with st.spinner("PDF wird zu Bildern konvertiert..."):
            images = pdf_to_images(pdf_bytes)

        if images is None:
            return False

        st.success(f"✅ {len(images)} Seiten zu Bildern konvertiert")

        # Perform OCR
        with st.spinner("OCR wird durchgeführt..."):
            extracted_text = perform_ocr_easyocr(images, reader)

        if not extracted_text.strip():
            st.warning("⚠️ Kein Text wurde aus dieser PDF extrahiert")
            return False

        # Create DOCX
        with st.spinner("DOCX-Dokument wird erstellt..."):
            original_name = Path(uploaded_file.name).stem
            docx_filename = f"{original_name}_OCR.docx"
            doc = create_docx(extracted_text, original_name)

        if doc is None:
            return False

        # Create download link
        create_download_link(doc, docx_filename)

        # Show preview of extracted text
        with st.expander(f"📖 Vorschau des extrahierten Texts aus {uploaded_file.name}"):
            st.text_area(
                "Extrahierter Text:",
                extracted_text[:2000] + ("..." if len(extracted_text) > 2000 else ""),
                height=200,
                key=f"preview_{uploaded_file.name}_{hash(extracted_text[:100])}"
            )

        return True

    except Exception as e:
        st.error(f"Error processing {uploaded_file.name}: {str(e)}")
        return False


def main():
    """Main application function"""
    setup_page()

    # Header
    st.title("📄 PDF OCR zu DOCX Konverter")
    st.markdown(
        "Lade eine oder mehrere PDF-Dateien hoch, um Text mittels OCR zu extrahieren und in das DOCX-Format zu konvertieren.")

    # Load OCR reader
    with st.spinner("OCR-Modell wird geladen... (Dies kann beim ersten Start einen Moment dauern)"):
        reader = load_ocr_reader()

    if reader is None:
        st.error(
            "❌ OCR-Modell konnte nicht geladen werden. Bitte aktualisieren Sie die Seite und versuchen Sie es erneut.")
        return

    st.success("✅ OCR-Modell erfolgreich geladen!")

    # Instructions
    with st.expander("ℹ️ Anleitung"):
        st.markdown("""
        1. **PDFs hochladen**: Zieh die PDF-Dateien per Drag & Drop hierher oder wählen Sie eine oder mehrere Dateien aus
        2. **Verarbeitung**: Die App konvertiert jede PDF-Seite in Bilder und führt OCR mit EasyOCR durch
        3. **Download**: Lade die generierten DOCX-Dateien mit dem extrahierten Text herunter

        **Funktionen**:
        - ✅ Unterstützt Deutsch und Englisch
        - ✅ Hochwertige OCR mit Deep Learning
        - ✅ Verarbeitung von mehreren Dateien
        - ✅ Textvorschau vor dem Download

        **Tipps für bessere Ergebnisse**:
        - Verwende hochwertige, klare PDF-Dateien
        - Stell sicher, dass der Text nicht zu klein oder unscharf ist
        - Die Verarbeitungszeit hängt von der Anzahl der Seiten ab
        """)

    # File uploader
    uploaded_files = st.file_uploader(
        "PDF-Dateien auswählen",
        type=['pdf'],
        accept_multiple_files=True,
        help="Es können mehrere PDF-Dateien gleichzeitig hochgeladen werden"
    )

    if uploaded_files:
        st.info(f"📁 {len(uploaded_files)} file(s) uploaded")

        # Process button
        if st.button("🚀 OCR-Verarbeitung starten", type="primary"):

            success_count = 0
            total_files = len(uploaded_files)

            # Process each file
            for i, uploaded_file in enumerate(uploaded_files):
                st.subheader(f"Verarbeitung {i + 1}/{total_files}: {uploaded_file.name}")

                if process_single_pdf(uploaded_file, reader):
                    success_count += 1
                    st.success(f"✅ {uploaded_file.name} erfolgreich verarbeitet")
                else:
                    st.error(f"❌ Fehler beim Verarbeiten von {uploaded_file.name}")

                st.divider()

            # Summary
            if success_count > 0:
                st.balloons()
                st.success(
                    f"🎉 Verarbeitung abgeschlossen! {success_count}/{total_files} Dateien erfolgreich verarbeitet.")
            else:
                st.error(
                    "❌ Keine Dateien wurden erfolgreich verarbeitet. Bitte überprüfen Sie Ihre PDF-Dateien und versuchen Sie es erneut.")

    # Footer
    st.markdown("---")
    st.markdown("""
    **Technologie**: Diese App verwendet PyMuPDF für PDF-Verarbeitung und EasyOCR für Texterkennung.

    **Unterstützte Sprachen**: Deutsch und Englisch werden automatisch erkannt und verarbeitet.
    """)


if __name__ == "__main__":
    main()
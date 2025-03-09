from flask import Flask, request, render_template, redirect, url_for, send_file
import os
import docx
import openpyxl
import fitz  # PyMuPDF for PDF editing

app = Flask(__name__)

# Ensure upload directory exists
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

@app.route("/")
def home():
    return render_template("upload.html")

@app.route("/upload", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return "No file uploaded", 400
    file = request.files["file"]
    if file.filename == "":
        return "No selected file", 400

    file_path = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
    file.save(file_path)  # Save the uploaded file

    # Process DOCX or XLSX
    extracted_text = ""
    if file.filename.endswith(".docx"):
        extracted_text = extract_text_from_docx(file_path)
    elif file.filename.endswith(".xlsx"):
        extracted_text = extract_text_from_xlsx(file_path)

    # ✅ Debugging: Print extracted text in logs
    print("Extracted Text:", extracted_text)

    # Process PDF (Fill extracted text into a PDF if uploaded)
    if file.filename.endswith(".pdf"):
        filled_pdf_path = fill_pdf(file_path, extracted_text)
        return f"PDF processed successfully! <br> <a href='/download/{filled_pdf_path}'>Download PDF</a>"

    # Show extracted text on the webpage
    return f"Extracted Text: <pre>{extracted_text}</pre>"

def extract_text_from_docx(docx_path):
    """Extract text from a .docx file"""
    doc = docx.Document(docx_path)
    text = "\n".join([para.text for para in doc.paragraphs])
    return text

def extract_text_from_xlsx(xlsx_path):
    """Extract text from an .xlsx file"""
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active
    extracted_text = ""
    for row in ws.iter_rows():
        extracted_text += " ".join(str(cell.value) for cell in row if cell.value) + "\n"
    return extracted_text

def fill_pdf(pdf_path, text):
    """Insert extracted text into the first page of a PDF"""
    output_pdf = pdf_path.replace(".pdf", "_filled.pdf")
    doc = fitz.open(pdf_path)
    
    # Modify the first page only
    page = doc[0]
    text_area = fitz.Point(100, 100)  # Position where text starts
    page.insert_text(text_area, text, fontsize=12, color=(0, 0, 0))  # Black text
    
    doc.save(output_pdf)
    doc.close()
    return os.path.basename(output_pdf)

@app.route("/download/<filename>")
def download_file(filename):
    """Serve the processed file for download"""
    return send_file(os.path.join(app.config["UPLOAD_FOLDER"], filename), as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0")

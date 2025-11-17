import os
from pdf2image import convert_from_path
import pytesseract

# 1) Tesseract yolu ve tessdata üst klasörü
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
os.environ["TESSDATA_PREFIX"] = r"C:\Program Files\Tesseract-OCR\tessdata"

# 2) Poppler yolu
POPPLER_BIN = r"C:\poppler-25.07.0\Library\bin"

# 3) PDF'den ilk 15 sayfayı OCR
pdf_path = r"C:\pdfs\Hukuk.pdf"
pages = convert_from_path(
    pdf_path,
    dpi=300,
    first_page=7,
    last_page=73,
    poppler_path=POPPLER_BIN
)

metin = []
for i, sayfa in enumerate(pages, start=1):
    # tessdata yolunu ayrıca config ile de veriyoruz (çifte garanti)
    text = pytesseract.image_to_string(
        sayfa,
        lang="tur",
    )
    metin.append(f"--- Sayfa {i} ---\n{text}\n")
    print(f"{i}. sayfa OCR tamam")

with open("cikti.txt", "w", encoding="utf-8") as f:
    f.write("".join(metin))

print("✅ İlk 15 sayfa OCR tamamlandı. cikti.txt oluşturuldu.")

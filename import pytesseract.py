# -*- coding: utf-8 -*-
import re
import unicodedata
import pandas as pd

# GİRİŞ/ÇIKIŞ DOSYALARI
INPUT_TXT = "cikti.txt"     # OCR'dan aldığın ham metin
OUTPUT_XLSX = "sozluk.xlsx"     # Çıktı Excel dosyası

# 1) Metni oku
with open(INPUT_TXT, "r", encoding="utf-8") as f:
    raw = f.read()

# 2) Normalizasyon ve temel temizlik
text = unicodedata.normalize("NFC", raw)
text = re.sub(r"(?m)^\s*---\s*Sayfa\s*\d+\s*---\s*$", "", text)   # Sayfa başlıklarını sil
text = text.replace("\u00ad", "")                                  # Soft hyphen temizle
text = re.sub(r"(\w)-\n(\w)", r"\1\2", text, flags=re.UNICODE)     # 'za-\nyıf' -> 'zayıf'

# Madde başlarını koru
PLACEHOLDER = "<<<ENTRYSEP>>>"
text = re.sub(r"\n\s*(?=[^\n—]+?\s—)", PLACEHOLDER, text)
text = re.sub(r"\n+", " ", text)
text = text.replace(PLACEHOLDER, "\n")
text = re.sub(r"\s{2,}", " ", text).strip()

# 3) "kelime — anlam" bloklarını yakala
pattern = re.compile(
    r"(?m)^\s*([^\n—]+?)\s—\s(.*?)(?=^\s*[^\n—]+?\s—\s|\Z)",
    flags=re.DOTALL
)

rows = []
for m in pattern.finditer(text):
    term = re.sub(r"\s{2,}", " ", m.group(1).strip())
    definition = re.sub(r"\s{2,}", " ", m.group(2).strip())
    definition = re.sub(r"\s+([,.;:!?])", r"\1", definition)
    if term and definition:
        rows.append({"kelime": term, "anlam": definition})

df = pd.DataFrame(rows, columns=["kelime", "anlam"])

# =========================
# 5) Türkçe alfabe ve sıralama + çoklu sheet yazımı
# =========================

# Türkçe büyük harfe çevirme (şapkalılar dâhil)
TR_UP_MAP = str.maketrans({
    "i": "İ", "ı": "I",
    "ş": "Ş", "ğ": "Ğ", "ç": "Ç", "ö": "Ö", "ü": "Ü",
    "â": "Â", "î": "Î", "û": "Û"
})
def tr_upper(s: str) -> str:
    return s.translate(TR_UP_MAP).upper()

# Türkçe alfabe (Q, W, X yok)
TR_ALPHABET = list("A B C Ç D E F G Ğ H I İ J K L M N O Ö P R S Ş T U Ü V Y Z".split())
ALPHA_INDEX = {ch: idx for idx, ch in enumerate(TR_ALPHABET)}

def tr_sort_key(word: str):
    """Türkçe harf sırasına göre sıralama anahtarı."""
    w = tr_upper(word)
    key = [ALPHA_INDEX.get(ch, 100 + ord(ch)) for ch in w]
    return key

def first_letter_bucket(term: str) -> str:
    """Kelimenin ilk harfine göre doğru sayfayı belirle (Â→A, Î→İ, Û→U)."""
    if not term:
        return "#"
    t = term.strip()
    # Baştaki tırnak, rakam vb. karakterleri temizle
    t = re.sub(r"^[^A-Za-zÇĞİIÖŞÜÂÎÛçğıiöşüâîû]+", "", t)
    if not t:
        return "#"

    first = t[0]
    # Şapkalı yönlendirme
    if first in ("Â", "â"):
        return "A"
    elif first in ("Î", "î"):
        return "İ"
    elif first in ("Û", "û"):
        return "U"

    first_up = tr_upper(first)
    return first_up if first_up in ALPHA_INDEX else "#"

# Her harfe göre gruplandır
groups = {}
for _, row in df.iterrows():
    bucket = first_letter_bucket(row["kelime"])
    groups.setdefault(bucket, []).append(row)

# Excel'e yaz: her grup kendi sayfasına
with pd.ExcelWriter(OUTPUT_XLSX, engine="xlsxwriter") as writer:
    workbook = writer.book

    # Özet sayfası
    summary_rows = []
    for ch in TR_ALPHABET + ["#"]:
        count = len(groups.get(ch, []))
        if count:
            summary_rows.append({"Harf": ch, "Kayıt Sayısı": count})
    if summary_rows:
        pd.DataFrame(summary_rows).to_excel(writer, sheet_name="Özet", index=False)
        ws_sum = writer.sheets["Özet"]
        ws_sum.freeze_panes(1, 0)
        ws_sum.set_column(0, 0, 8)
        ws_sum.set_column(1, 1, 14)

    # Harf bazlı sayfalar
    for ch in TR_ALPHABET + ["#"]:
        rows_list = groups.get(ch, [])
        if not rows_list:
            continue
        gdf = pd.DataFrame(rows_list, columns=["kelime", "anlam"])
        gdf = gdf.sort_values(by="kelime", key=lambda s: s.map(tr_sort_key), kind="stable")

        sheet_name = ch if ch != "#" else "Diger"
        gdf.to_excel(writer, sheet_name=sheet_name, index=False)

        ws = writer.sheets[sheet_name]
        ws.freeze_panes(1, 0)
        ws.set_column(0, 0, 28)
        ws.set_column(1, 1, 90)

print("✅ Şapkalı harflerle uyumlu çok sayfalı Türkçe sözlük oluşturuldu →", OUTPUT_XLSX)

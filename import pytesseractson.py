# -*- coding: utf-8 -*-
import re
import unicodedata
import pandas as pd

# =========================
# 0) Kullanıcıdan giriş/çıkış isimlerini al
# =========================

INPUT_TXT = input("Girdi TXT dosyasının adı (örn: cikti_u.txt): ").strip()
if not INPUT_TXT:
    print("❌ Girdi dosya adı boş olamaz!")
    raise SystemExit

if not INPUT_TXT.lower().endswith(".txt"):
    INPUT_TXT += ".txt"

OUTPUT_XLSX = input("Çıktı Excel dosyasının adı (örn: sozluk_u.xlsx): ").strip()
if not OUTPUT_XLSX:
    print("❌ Çıktı dosya adı boş olamaz!")
    raise SystemExit

if not OUTPUT_XLSX.lower().endswith(".xlsx"):
    OUTPUT_XLSX += ".xlsx"

print(f"\n📥 Girdi TXT: {INPUT_TXT}")
print(f"📤 Çıktı XLSX: {OUTPUT_XLSX}")

# =========================
# 1) Metni oku
# =========================
with open(INPUT_TXT, "r", encoding="utf-8") as f:
    raw = f.read()

# =========================
# 2) Normalizasyon ve temel temizlik
# =========================
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

# =========================
# 3) "kelime — anlam" bloklarını yakala
# =========================
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
print(f"\n📊 Toplam madde sayısı (ham): {len(df)}")

# =========================
# 4) Türkçe alfabe ve sıralama yardımcıları
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
    """
    Kelimenin ilk harfine göre doğru sayfayı/harfi belirle (Â→A, Î→İ, Û→U).
    Baştaki tırnak, rakam, parantez vb. çöpleri atmaya çalışır.
    """
    if not term:
        return "#"
    t = term.strip()
    # Baştaki alakasız karakterleri temizle
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

# =========================
# 5) Kullanıcıdan hangi harf için sözlük yapılacağını al
# =========================
chosen = input("\nHangi harf için sözlük oluşturulsun? (örn: U): ").strip()
if not chosen:
    print("❌ Harf boş olamaz!")
    raise SystemExit

# Kullanıcının girdiği harfi bucketa çevir (Â→A, û→U gibi)
bucket = first_letter_bucket(chosen)
if bucket == "#":
    print(f"❌ '{chosen}' için geçerli bir harf bulunamadı.")
    raise SystemExit

print(f"🔠 Seçilen harf: {chosen} → gerçek bucket: {bucket}")

# =========================
# 6) Sadece bu harfle başlayan kelimeleri filtrele
# =========================
filtered_rows = []
for _, row in df.iterrows():
    b = first_letter_bucket(row["kelime"])
    if b == bucket:
        filtered_rows.append(row)

if not filtered_rows:
    print(f"⚠️ '{bucket}' harfiyle başlayan hiç madde bulunamadı.")
    raise SystemExit

gdf = pd.DataFrame(filtered_rows, columns=["kelime", "anlam"])
print(f"✅ '{bucket}' harfiyle başlayan madde sayısı: {len(gdf)}")

# Türkçe sıralamaya göre sırala
gdf = gdf.sort_values(by="kelime", key=lambda s: s.map(tr_sort_key), kind="stable")

# =========================
# 7) Excel'e yaz (sadece seçilen harf için tek sheet)
# =========================
sheet_name = bucket  # örn: "U"
if sheet_name == "#":
    sheet_name = "Diger"

with pd.ExcelWriter(OUTPUT_XLSX, engine="xlsxwriter") as writer:
    # İsteğe bağlı: küçük bir özet sheet'i
    summary_df = pd.DataFrame(
        [{"Harf": bucket, "Kayıt Sayısı": len(gdf)}],
        columns=["Harf", "Kayıt Sayısı"]
    )
    summary_df.to_excel(writer, sheet_name="Özet", index=False)
    ws_sum = writer.sheets["Özet"]
    ws_sum.freeze_panes(1, 0)
    ws_sum.set_column(0, 0, 8)
    ws_sum.set_column(1, 1, 14)

    # Asıl harf sheet'i
    gdf.to_excel(writer, sheet_name=sheet_name, index=False)
    ws = writer.sheets[sheet_name]
    ws.freeze_panes(1, 0)
    ws.set_column(0, 0, 28)  # kelime
    ws.set_column(1, 1, 90)  # anlam

print(f"\n✅ '{bucket}' harfi için Türkçe sözlük Excel dosyası oluşturuldu → {OUTPUT_XLSX}")

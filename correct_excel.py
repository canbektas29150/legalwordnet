# correct_with_txt.py
# -*- coding: utf-8 -*-
import re, unicodedata
from pathlib import Path
import pandas as pd
from difflib import SequenceMatcher

# ================== YOLLAR (gerekirse değiştir) ==================
TXT_PATH   = Path(r"ciktiafull.txt")        # A harfi TXT kaynağı
XLSX_IN    = Path(r"sozlukafull.xlsx")      # Mevcut sözlük (A sayfası içinde)
XLSX_OUT   = Path(r"sozluk_A_corrected.xlsx")
DIFF_CSV   = Path(r"sozluk_A_corrections_report.csv")
# ================================================================

HEAD = re.compile(r"^(?P<term>[^\n:—]{1,200}?)\s(?:—|:)\s(?P<def>.*)$")

TR_UP_MAP = str.maketrans({
    "i":"İ","ı":"I","ş":"Ş","ğ":"Ğ","ç":"Ç","ö":"Ö","ü":"Ü",
    "â":"Â","î":"Î","û":"Û"
})
TR_DOWN_MAP = str.maketrans({"I":"ı","İ":"i"})

def tr_upper(s: str) -> str:
    return s.translate(TR_UP_MAP).upper()

def norm_tr(s: str) -> str:
    if s is None: return ""
    s = str(s).strip().translate(TR_DOWN_MAP)
    s = unicodedata.normalize("NFKC", s)
    return s.casefold()

def clean_txt(text: str) -> str:
    t = unicodedata.normalize("NFC", text)
    t = re.sub(r"(?mi)^\s*---\s*Sayfa\s*\d+\s*---\s*$", "", t)
    t = re.sub(r"(?mi)^\s*(sayfa\s*)?\d+\s*$", "", t)
    t = t.replace("\u00ad","")
    t = re.sub(r"(\w)-\n(\w)", r"\1\2", t)
    t = t.replace("–","—").replace("−","-")
    t = re.sub(r"\s+\-\s+"," — ", t)
    t = re.sub(r"\s+—\s+"," — ", t)
    t = re.sub(r"[ \t]+\n","\n", t)
    t = re.sub(r"\n{3,}","\n\n", t)
    t = re.sub(r"[ \t]{2,}"," ", t)
    return t

def looks_like_term_line(line: str) -> bool:
    if not line: return False
    if "—" in line or ":" in line: return False
    if len(line) > 80: return False
    if re.search(r"[^A-Za-zÇĞİIÖŞÜÂÎÛçğıiöşüâîû\s]", line): return False
    return bool(re.search(r"[A-Za-zÇĞİIÖŞÜÂÎÛçğıiöşüâîû]", line))

def first_bucket(term: str) -> str:
    s = term.strip()
    s = re.sub(r"^[^A-Za-zÇĞİIÖŞÜÂÎÛçğıiöşüâîû]+", "", s)
    if not s: return "#"
    f = s[0]
    if f in ("Â","â"): return "A"
    if f in ("Î","î"): return "İ"
    if f in ("Û","û"): return "U"
    return tr_upper(f)

def guess_pos(defn: str) -> str:
    s = "" if defn is None else str(defn).strip().lower()
    words = re.findall(r"[a-zçğıöşüâîû]+", s)
    if words and (words[-1].endswith("mak") or words[-1].endswith("mek")):
        return "VERB"
    return "NOUN"

def parse_entries_with_prefix_merge(text: str):
    entries = []
    term_prefix_buf, cur_term, cur_def_parts = [], None, []
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            if cur_term and cur_def_parts and cur_def_parts[-1] != "":
                cur_def_parts.append("")
            continue
        m = HEAD.match(line)
        if m and not re.fullmatch(r"[^\wÇĞİIÖŞÜÂÎÛçğıiöşüâîû]+", m.group("term").strip()):
            if cur_term is not None:
                entries.append((cur_term, " ".join(p for p in cur_def_parts if p != "")))
                cur_term, cur_def_parts = None, []
            term_core = re.sub(r"\s{2,}", " ", m.group("term").strip())
            if term_prefix_buf:
                prefix = " ".join(tp.strip() for tp in term_prefix_buf if tp.strip())
                full_term = (prefix + " " + term_core).strip()
            else:
                full_term = term_core
            cur_term = full_term
            cur_def_parts = [m.group("def").strip()]
            term_prefix_buf = []
            continue
        if cur_term is None:
            if looks_like_term_line(line): term_prefix_buf.append(line)
            else: term_prefix_buf = []
        else:
            cur_def_parts.append(line)
    if cur_term is not None:
        entries.append((cur_term, " ".join(p for p in cur_def_parts if p != "")))

    cleaned = []
    for term, defi in entries:
        term = re.sub(r"\s{2,}", " ", term).strip(" -–—:.;, \t")
        defi = re.sub(r"\s{2,}", " ", defi)
        defi = re.sub(r"\s+([,.;:!?])", r"\1", defi).strip()
        cleaned.append((term, defi))
    return cleaned

def load_a_sheet(xlsx_path: Path) -> pd.DataFrame:
    xls = pd.ExcelFile(xlsx_path)
    def norm_col(s): return re.sub(r"\s+", "", str(s)).strip().lower()
    # A sayfasını bul, yoksa ilk sayfa
    sheet_name = "A" if "A" in xls.sheet_names else xls.sheet_names[0]
    df = xls.parse(sheet_name)
    # kolon isimlerini normalize et
    ren = {}
    for c in df.columns:
        cn = norm_col(c)
        if cn == "kelime": ren[c] = "kelime"
        elif cn == "anlam": ren[c] = "anlam"
        elif cn == "pos": ren[c] = "POS"
        elif cn == "r": ren[c] = "R"
    df = df.rename(columns=ren)
    if "kelime" not in df.columns:
        df = df.rename(columns={df.columns[0]:"kelime"})
    if "anlam" not in df.columns:
        if len(df.columns) > 1: df = df.rename(columns={df.columns[1]:"anlam"})
        else: df["anlam"] = ""
    return df

def similarity(a: str, b: str) -> float:
    a = (a or "").strip().lower()
    b = (b or "").strip().lower()
    if not a and not b: return 1.0
    if not a or not b: return 0.0
    return SequenceMatcher(None, a, b).ratio()

def main():
    # 1) TXT'ten A maddelerini (prefix-merge) çıkar
    raw = TXT_PATH.read_text(encoding="utf-8", errors="ignore")
    t = clean_txt(raw)
    entries = [(term, defi) for term, defi in parse_entries_with_prefix_merge(t) if first_bucket(term) == "A"]

    # 2) TXT entries -> index: last_token_norm -> list of (full_term, def)
    last_map = {}
    for term, defi in entries:
        toks = [w for w in re.findall(r"[A-Za-zÇĞİIÖŞÜÂÎÛçğıiöşüâîû]+", term)]
        if not toks: continue
        last = norm_tr(toks[-1])
        last_map.setdefault(last, []).append((term, defi))

    # 3) Excel A sayfasını yükle
    df = load_a_sheet(XLSX_IN)
    if "POS" not in df.columns: df["POS"] = ""
    if "R" not in df.columns:   df["R"] = 0

    # 4) Satır bazında düzeltme
    changes = []
    for i, row in df.iterrows():
        old_term = str(row["kelime"]).strip()
        old_def  = str(row.get("anlam", "")).strip()
        key = norm_tr(old_term)
        # önce doğrudan TXT terim eşleşmesi
        direct_match = [(t, d) for (t, d) in entries if norm_tr(t) == key]
        candidate = None
        reason = ""
        if direct_match:
            candidate = max(direct_match, key=lambda x: similarity(old_def, x[1]))
            reason = "exact_term_match"
        else:
            # son kelime eşleşmesi (örn. Excel: reus, TXT: ... reus)
            cands = last_map.get(key, [])
            if cands:
                # tanım benzerliği + alt dize bonusu ile en iyisini seç
                def score(c):
                    t, d = c
                    sim = similarity(old_def, d)
                    if old_def and d and (old_def.lower() in d.lower() or d.lower() in old_def.lower()):
                        sim += 0.1
                    return sim
                candidate = max(cands, key=score)
                reason = "last_token_match"
        if candidate:
            cand_term, cand_def = candidate
            sim = similarity(old_def, cand_def)
            # eşik: ya benzerlik >= 0.55, ya da excel tanımı txt tanımının alt dizini
            ok = sim >= 0.55 or (old_def and cand_def and old_def.lower() in cand_def.lower())
            if ok and norm_tr(cand_term) != key:
                # düzelt
                df.at[i, "kelime"] = cand_term
                df.at[i, "anlam"]  = cand_def if len(cand_def) >= len(old_def) else old_def
                df.at[i, "POS"]    = guess_pos(df.at[i, "anlam"])
                changes.append({
                    "row": i+2,  # başlık satırı sonrası Excel indekslemesi
                    "reason": reason,
                    "similarity": round(sim, 3),
                    "old_term": old_term,
                    "new_term": cand_term,
                    "old_def": old_def,
                    "new_def": df.at[i, "anlam"],
                })

    # 5) Çıktılar
    with pd.ExcelWriter(XLSX_OUT, engine="xlsxwriter") as w:
        df.to_excel(w, sheet_name="A", index=False)
        ws = w.sheets["A"]
        try:
            ws.set_column(0, 0, 42)
            ws.set_column(1, 1, 100)
            ws.set_column(2, 3, 10)
        except Exception:
            pass

    pd.DataFrame(changes).to_csv(DIFF_CSV, index=False, encoding="utf-8")

    print("✅ Düzeltme tamam.")
    print("  Düzeltilen satır sayısı:", len(changes))
    print("  ->", XLSX_OUT)
    print("  ->", DIFF_CSV)
    if changes[:5]:
        print("  Örnek değişiklikler (ilk 5):")
        for c in changes[:5]:
            print(f"   - r{c['row']} [{c['reason']}, sim={c['similarity']}] {c['old_term']}  ->  {c['new_term']}")

if __name__ == "__main__":
    main()

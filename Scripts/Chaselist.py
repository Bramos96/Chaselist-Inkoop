import os
import re
import pandas as pd
import win32com.client as win32
from typing import Optional, Tuple

# 
# CONFIGURATION
# 
BASE_FOLDER = r"C:\Users\bram.gerrits\Desktop\Automations\Inkoop\Chaselist Madelon"
INPUT_FOLDER = os.path.join(BASE_FOLDER, "Input")
SUPPLIER_INFO_FILE = os.path.join(INPUT_FOLDER, "Leveranciers informatie.xlsx")
LAYOUTS_FOLDER = os.path.join(BASE_FOLDER, "Layouts")
NL_TXT_FILE = os.path.join(LAYOUTS_FOLDER, "NL.txt")
EN_TXT_FILE = os.path.join(LAYOUTS_FOLDER, "EN.txt")


TESTMODE = False
TEST_EMAIL = "bram.gerrits@vhe.nl"
STATUS_VALUES = {"n/b", "mail"}
CHASE_PREFIX = "Chase"


# 
# HELPERS
# 
def find_chase_file(folder: str, prefix: str) -> str:
    """Zoekt het nieuwste Chase*-bestand op wijzigingsdatum in de opgegeven map."""
    candidates = []
    for fname in os.listdir(folder):
        if fname.lower().startswith(prefix.lower()) and fname.lower().endswith((".xlsx", ".xlsm", ".xls")):
            full_path = os.path.join(folder, fname)
            mtime = os.path.getmtime(full_path)
            candidates.append((mtime, full_path))

    if not candidates:
        raise FileNotFoundError(f"Geen Chase*-bestand gevonden in {folder}")

    # Sorteer op nieuwste wijzigingsdatum
    newest = max(candidates, key=lambda x: x[0])[1]
    print(f" Nieuwste Chase-bestand geselecteerd: {os.path.basename(newest)}")
    return newest


def latest_wk_sheet(xl_path: str) -> str:
    xl = pd.ExcelFile(xl_path)
    wk_sheets = [(s, int(re.search(r"\d+", s).group())) for s in xl.sheet_names if re.search(r"^wk\s*\d+", s, re.I)]
    if not wk_sheets:
        raise ValueError("No 'WK ####' sheets found.")
    latest = max(wk_sheets, key=lambda x: x[1])[0]
    print(f" Using sheet: {latest}")
    return latest

def normalize_status(v) -> str:
    if pd.isna(v): return "n/b"
    s = str(v).strip().lower()
    return "n/b" if s in {"#n/b", "#n/a", "n/b", "n/a", "na"} else ("mail" if "mail" in s else s)

def load_supplier_info(path: str) -> Optional[pd.DataFrame]:
    return pd.read_excel(path) if os.path.exists(path) else None

def format_article(a) -> str:
    if pd.isna(a): return ""
    s = re.sub(r"\D", "", str(a))
    return f"{s[:4]}.{s[4:7]}.{s[7:]}" if s.startswith("4022") and len(s) > 7 else a

def parse_date_force(v) -> pd.Timestamp:
    if pd.isna(v): return pd.NaT
    s = str(v).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y"):
        try:
            return pd.to_datetime(s, format=fmt)
        except Exception:
            continue
    return pd.to_datetime(s, errors="coerce")

def sort_by_delivery(df: pd.DataFrame, col: str) -> pd.DataFrame:
    df = df.copy()
    df["__sort"] = df[col].apply(parse_date_force)
    df = df.sort_values("__sort", ascending=True, na_position="last").drop(columns="__sort")
    return df

# ─────────────────────────────────────────────
# MAIL + HTML
# ─────────────────────────────────────────────
def make_html_table(df: pd.DataFrame) -> str:
    today = pd.Timestamp.today().normalize()
    cols = [("Artikel", "Item"), ("Item leverancier", "Item supplier"),
            ("Bestelnummer", "Order number"), ("Regelnummer", "Line number"),
            ("Huidige leverdatum", "Current delivery date"),
            ("Gewenste leverdatum", "Requested delivery date")]
    head = "".join([f"<th style='background:#d8edf7;border:1px solid #999;padding:6px'>{c[1]}</th>" for c in cols])
    rows = []
    for i, row in df.iterrows():
        cells = []
        for src, disp in cols:
            val = row.get(src, "")
            if src == "Artikel": val = format_article(val)
            dt = parse_date_force(val)
            text = dt.strftime("%d-%m-%Y") if pd.notna(dt) else str(val)
            if disp == "Current delivery date" and pd.notna(dt) and dt < today:
                text = f"<b><font color='red'>{text}</font></b>"
            cells.append(f"<td style='border:1px solid #999;padding:6px'>{text}</td>")
        rows.append("<tr>" + "".join(cells) + "</tr>")
    return f"<table style='border-collapse:collapse;font-family:Arial;font-size:9pt'><tr>{head}</tr>{''.join(rows)}</table>"

def detect_lang_and_name(info: Optional[pd.DataFrame], supplier: str) -> Tuple[str, str]:
    """
    Determine language (NL/ENG) and correct greeting name:
    - If Contactnaam/Contactpersoon exists  use it (unless it's NaN/None/empty).
    - If not:
         'heer/mevrouw' for NL
         'Sir/Madam' for ENG
    """

    lang = "NL"
    name = "heer/mevrouw"

    if info is None or info.empty:
        return lang, name

    # Zoek leverancier
    mask = info.iloc[:, 0].astype(str).str.strip().str.lower() == supplier.lower()
    match = info[mask]
    if match.empty:
        return lang, name

    row = match.iloc[0]

    # Taal bepalen
    for col in info.columns:
        if str(col).strip().lower() in {"eng/nl", "taal", "language"}:
            val = str(row[col]).strip().upper()
            if "EN" in val:
                lang = "ENG"
                name = "Sir/Madam"
            break

    # Contactnaam bepalen
    for col in info.columns:
        if any(word in str(col).strip().lower() for word in ["contact", "naam", "name"]):
            val = str(row[col]).strip()
            # check op NaN, None, lege tekst of whitespace
            if val and val.lower() not in {"nan", "none", ""}:
                name = val
            else:
                name = "heer/mevrouw" if lang == "NL" else "Sir/Madam"
            break

    return lang, name



def build_mail_body(lang: str, name: str, html_table: str, nl_tmpl: str, en_tmpl: str) -> str:
    t = en_tmpl if lang == "ENG" else nl_tmpl

    # 1 Clean base text
    body = (
        t.replace("<naam>", name)
         .replace("<handtekening>", "")
         .replace("<tafel>", f"<div style='margin-top:10px;margin-bottom:4px'>{html_table}</div>")
         .replace("\r", "")
         .replace("\n", "<br>")
    )

    # 2️ Collapse extra blank lines (Outlook-safe)
    body = re.sub(r"(<br>\s*){3,}", "<br><br>", body)     # no triple spacing
    body = re.sub(r"(<br>\s*)+$", "", body)               # trim trailing breaks

    # 3️ Ensure exactly ONE blank line before table, NONE after text
    body = body.replace("<br><div", "<div")               # remove blank above table
    body = re.sub(r"</div><br>", "</div>", body)          # remove blank below table

    return body



def send_mail(to: str, subj: str, html_body: str):
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = to
    mail.Subject = subj
    mail.BodyFormat = 2
    mail.Display()
    sig = mail.HTMLBody
    mail.HTMLBody = f"<html><body>{html_body}<br>{sig}</body></html>"
    mail.Display()

# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
def main():
    chase = find_chase_file(BASE_FOLDER, CHASE_PREFIX)
    sheet = latest_wk_sheet(chase)
    df = pd.read_excel(chase, sheet_name=sheet, dtype=str)
    if df.empty:
        print(" No data found.")
        return

    status_col = next((c for c in df.columns if "status" in c.lower()), None)
    if not status_col:
        raise ValueError("No Status column found.")

    df["__status"] = df[status_col].map(normalize_status)
    df = df[df["__status"].isin(STATUS_VALUES)].drop(columns="__status")
    if df.empty:
        print(" No rows with Status == N/B or Mail.")
        return

    supplier_col = next((c for c in df.columns if "leverancier" in c.lower()), df.columns[1])
    supplier_info = load_supplier_info(SUPPLIER_INFO_FILE)

    with open(NL_TXT_FILE, encoding="utf-8") as f: nl_tmpl = f.read()
    with open(EN_TXT_FILE, encoding="utf-8") as f: en_tmpl = f.read()

    for supplier, grp in df.groupby(supplier_col):
        try:
            supplier = str(supplier).strip()
            if not supplier:
                continue

            # Controleer of leverancier in Leveranciers informatie voorkomt
            if supplier_info is not None:
                supplier_names = supplier_info.iloc[:, 0].astype(str).str.strip().str.lower()
                if supplier.lower() not in supplier_names.values:
                    print(f" {supplier} niet gevonden in Leveranciers informatie — overslaan.")
                    continue

            grp = sort_by_delivery(grp, "Huidige leverdatum")
            lang, name = detect_lang_and_name(supplier_info, supplier)
            html_table = make_html_table(grp)
            body = build_mail_body(lang, name, html_table, nl_tmpl, en_tmpl)

            subj = (
                f"Pending order(s) - VHE Industrial Automation - {supplier}"
                if lang == "ENG"
                else f"Openstaande bestelling(en) - VHE Industrial Automation - {supplier}"
            )

                        # 1) Standaard: testadres
            to_addr = TEST_EMAIL

            # 2) In productie-modus proberen we het echte e-mailadres te vinden
            if not TESTMODE and supplier_info is not None:
                mask = supplier_info.iloc[:, 0].astype(str).str.strip().str.lower() == supplier.lower()
                if mask.any():
                    row = supplier_info[mask].iloc[0]

                    # Zoek een kolom met 'mail' in de naam (bijv. 'E-mail', 'Email', 'Mailadres')
                    email_col = None
                    for col in supplier_info.columns:
                        if "mail" in str(col).lower():
                            email_col = col
                            break

                    if email_col is not None:
                        raw_mail = str(row[email_col]).strip()
                        if raw_mail and raw_mail.lower() not in {"nan", "none"}:
                            to_addr = raw_mail  # echte mail uit Excel

            print(f" Sending {lang} mail to {to_addr} for {supplier}")
            send_mail(to_addr, subj, body)


        except Exception as e:
            print(f" Fout bij verwerken van leverancier '{supplier}': {e}")
            continue  # ga gewoon door met de volgende leverancier


    print(" Done.")

if __name__ == "__main__":
    main()

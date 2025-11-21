import os
import re
from typing import Tuple
from typing import Optional, List

import pandas as pd
import win32com.client as win32

# ───────────────────────────────────────────────────────────
# INSTELLINGEN
# ───────────────────────────────────────────────────────────

BASE_FOLDER = r"C:\Users\bram.gerrits\Desktop\Automations\Chaselist Madelon"
CHASE_PREFIX = "Chase"  # bestand begint met "Chase"

NL_TXT_FILE = os.path.join(BASE_FOLDER, "NL.txt")
EN_TXT_FILE = os.path.join(BASE_FOLDER, "EN.txt")

with open(NL_TXT_FILE, "r", encoding="utf-8") as f:
    NL_TEMPLATE = f.read()

with open(EN_TXT_FILE, "r", encoding="utf-8") as f:
    EN_TEMPLATE = f.read()

SUPPLIER_INFO_FILE = os.path.join(BASE_FOLDER, "Leveranciers informatie.xlsx")

TESTMODE = True
TEST_EMAIL = "bram.gerrits@vhe.nl"

# Statuswaarden die we willen meenemen
STATUS_VALUES_NORMALIZED = {"n/b", "mail"}  # n/b representeert alle #N/B-achtige dingen

# Kolommen uit de bron + benamingen in de mail
COLUMN_MAPPING = [
    ("Artikel",            "Item"),
    ("Item leverancier",   "Item supplier"),
    ("Bestelnummer",       "Order number"),
    ("Regelnummer",        "Line number"),
    ("Huidige leverdatum", "Current delivery date"),
    ("Gewenste leverdatum","Requested delivery date"),
]

# ───────────────────────────────────────────────────────────
# HULPFUNCTIES BESTAND / SHEET
# ───────────────────────────────────────────────────────────

def find_chase_file(folder: str, prefix: str) -> str:
    for fname in os.listdir(folder):
        if fname.lower().startswith(prefix.lower()) and fname.lower().endswith((".xlsx", ".xlsm", ".xls")):
            return os.path.join(folder, fname)
    raise FileNotFoundError(f"Geen Chase*-bestand gevonden in {folder}")


def extract_wk_number(sheet_name: str) -> Optional[int]:
    s = sheet_name.strip()
    if not s.lower().startswith("wk"):
        return None
    match = re.search(r"(\d+)", s)
    if match:
        return int(match.group(1))
    return None


def find_latest_wk_sheet(xl_path: str) -> str:
    xl = pd.ExcelFile(xl_path)
    best_sheet = None
    best_num = None

    for sheet in xl.sheet_names:
        num = extract_wk_number(sheet)
        if num is not None and (best_num is None or num > best_num):
            best_num = num
            best_sheet = sheet

    if best_sheet is None:
        raise ValueError(f"Geen WK-sheets gevonden in {xl_path}")

    print(f"→ Geselecteerde WK-sheet: {best_sheet} (nummer {best_num})")
    return best_sheet


# ───────────────────────────────────────────────────────────
# KOLommen zoeken
# ───────────────────────────────────────────────────────────

def find_status_column(columns: List[str]) -> Optional[str]:
    """
    Zoek naar een kolom waarvan de naam het woord 'status' bevat.
    Case-insensitive.
    """
    for col in columns:
        if "status" in str(col).strip().lower():
            return col
    return None


def get_supplier_column(df: pd.DataFrame) -> str:
    """
    Leverancier = kolomnaam met 'leverancier' erin, anders kolom B.
    """
    for col in df.columns:
        if "leverancier" in str(col).strip().lower():
            return col

    if df.shape[1] < 2:
        raise ValueError("Er zijn minder dan 2 kolommen; kan kolom B niet gebruiken.")
    return df.columns[1]


# ───────────────────────────────────────────────────────────
# DATA FILTEREN
# ───────────────────────────────────────────────────────────

def normalize_status(val) -> str:
    """
    Normaliseer Status-waarde:
    - Excel NA / fout → 'n/b'
    - '#N/B', '#N/A', 'N/B', etc. → 'n/b'
    - 'Mail', 'MAIL', etc. → 'mail'
    - anders: lowercase string
    """
    if pd.isna(val):
        return "n/b"  # NA-achtige foutwaarden

    s = str(val).strip().lower()

    if s in {"#n/b", "#n/a", "#na", "n/b", "n/a"}:
        return "n/b"

    if s == "mail":
        return "mail"

    return s


def filter_rows_on_status(df: pd.DataFrame, status_col: str) -> pd.DataFrame:
    norm_col = df[status_col].map(normalize_status)
    mask = norm_col.isin(STATUS_VALUES_NORMALIZED)
    filtered = df[mask].copy()
    print(f"→ Gevonden {mask.sum()} rijen met Status == N/B of Mail")
    return filtered


def load_supplier_info(path: str) -> Optional[pd.DataFrame]:
    if not os.path.exists(path):
        print(f"⚠ Leveranciers informatie bestand niet gevonden: {path}")
        return None
    return pd.read_excel(path)


def format_4022_item(val) -> str:
    """
    Als artikel begint met 4022 → zet om naar 4022.xxx.xxxxx(x)
    Voorbeeld: '402243612012' → '4022.436.12012'
    """
    if pd.isna(val):
        return ""
    s = str(val)
    # haal alles behalve cijfers weg (voor het geval er spaties of andere troep inzit)
    digits = "".join(ch for ch in s if ch.isdigit())

    if not digits.startswith("4022") or len(digits) <= 7:
        # niks doen als het geen 4022-code is of te kort
        return s

    eerste = digits[:4]       # 4022
    tweede = digits[4:7]      # 3 cijfers
    rest = digits[7:]         # alles erna (5 of 6 cijfers)

    return f"{eerste}.{tweede}.{rest}"


# ───────────────────────────────────────────────────────────
# MAIL OPBOUW – TABEL
# ───────────────────────────────────────────────────────────

def make_html_table_for_group(group: pd.DataFrame) -> str:
    import html

    missing = [src for src, _ in COLUMN_MAPPING if src not in group.columns]
    if missing:
        raise KeyError(f"Kolommen ontbreken in data: {missing}")

    today = pd.Timestamp.today().normalize()

    # Build header row
    header_cells = ""
    for _, display in COLUMN_MAPPING:
        header_cells += (
            f"<th style='border:1px solid #999; background-color:#d8edf7; "
            f"padding:6px; text-align:left; font-weight:bold;'>{html.escape(display)}</th>"
        )

    # Build body rows
    rows_html = ""
    for row_idx, (_, row) in enumerate(group.iterrows()):
        bg_color = "#f9f9f9" if (row_idx % 2 == 0) else "#ffffff"
        row_cells = ""
        for src_name, display_name in COLUMN_MAPPING:
            val = row[src_name]
            if src_name == "Artikel":
                val = format_4022_item(val)

            dt = pd.to_datetime(val, dayfirst=True, errors="coerce")
            val_str = dt.strftime("%d-%m-%Y") if pd.notna(dt) else ("" if pd.isna(val) else str(val))
            val_str = html.escape(val_str)

            # Highlight overdue date
            if display_name == "Current delivery date" and pd.notna(dt) and dt.normalize() < today:
                val_str = f"<b><font color='red'>{val_str}</font></b>"

            row_cells += f"<td style='border:1px solid #999; padding:6px; vertical-align:top;'>{val_str}</td>"

        rows_html += f"<tr style='background-color:{bg_color};'>{row_cells}</tr>"

    # Return final HTML table (no thead, all clean markup)
    table_html = f"""
    <table border="1" cellspacing="0" cellpadding="4" 
           style="border-collapse:collapse; width:auto; font-family:Arial, sans-serif; font-size:9pt;">
        <tr>{header_cells}</tr>
        {rows_html}
    </table>
    """.strip()

    return table_html


# ───────────────────────────────────────────────────────────
# OUTLOOK
# ───────────────────────────────────────────────────────────

def get_outlook() -> win32.CDispatch:
    return win32.Dispatch("Outlook.Application")


def send_mail(to_addr: str, subject: str, html_body_without_signature: str):
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # olMailItem

    mail.To = to_addr
    mail.Subject = subject

    # Ensure HTML format
    mail.BodyFormat = 2  # 2 = olFormatHTML

    # Retrieve the signature properly
    mail.Display()
    signature_html = mail.HTMLBody

    # Wrap full HTML document properly
    final_html = f"""
    <html>
    <head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <style>
        table, th, td {{
            border-collapse: collapse;
            border: 1px solid #999;
            font-family: Arial, sans-serif;
            font-size: 9pt;
        }}
        th {{
            background-color: #d8edf7;
            text-align: left;
            padding: 6px;
        }}
        td {{
            padding: 6px;
        }}
    </style>
    </head>
    <body>
    {html_body_without_signature}
    <br>
    {signature_html}
    </body>
    </html>
    """

    # Assign to HTMLBody AFTER setting format
    mail.HTMLBody = final_html

    mail.Display()  # to preview before sending


# ───────────────────────────────────────────────────────────
# TEMPLATE-HELPERS
# ───────────────────────────────────────────────────────────

def detect_language_and_name(
    supplier_info: Optional[pd.DataFrame],
    supplier_str: str,
    candidates: Optional[pd.DataFrame] = None
) -> Tuple[str, str]:

    """
    Bepaal taal (NL / ENG) en aanspreeknaam (<naam>) op basis van Leveranciers informatie.
    """
    lang = "NL"
    contact_name = supplier_str

    if supplier_info is None:
        return lang, contact_name

    if candidates is None or candidates.empty:
        return lang, contact_name

    # Kolomnamen zoeken
    lang_col = None
    name_col = None
    for col in supplier_info.columns:
        cname = str(col).strip().lower()
        if lang_col is None and cname in {"eng/nl", "engnl", "taal", "language"}:
            lang_col = col
        if name_col is None and any(word in cname for word in ["contact", "naam", "name"]):
            name_col = col

    # Taal ophalen
    if lang_col is not None:
        lang_val = str(candidates.iloc[0][lang_col]).strip().upper()
        if lang_val in {"ENG", "EN"}:
            lang = "ENG"
        else:
            lang = "NL"

    # Naam ophalen
    if name_col is not None:
        nm = str(candidates.iloc[0][name_col]).strip()
        if nm:
            contact_name = nm

    return lang, contact_name


def build_body_from_template(
    lang: str,
    contact_name: str,
    html_table: str
) -> str:
    template = NL_TEMPLATE if lang != "ENG" else EN_TEMPLATE

    # 1) placeholders <naam> en <handtekening> vervangen in de TEXT
    body = template
    body = body.replace("<naam>", contact_name)
    body = body.replace("<handtekening>", "")

    # 2) linebreaks in de tekst omzetten naar <br>, ZONDER tabel erin
    body = body.replace("\r\n", "\n").replace("\r", "\n")
    body = body.replace("\n", "<br>\n")

    # 3) nu pas de HTML-tabel invoegen, ongewijzigd
    body = body.replace("<tafel>", html_table)

    return body



# ───────────────────────────────────────────────────────────
# MAIN
# ───────────────────────────────────────────────────────────

def main():
    # 1) Zoek Chase-bestand
    chase_file = find_chase_file(BASE_FOLDER, CHASE_PREFIX)
    print(f"→ Geselecteerd Chase-bestand: {chase_file}")

    # 2) Laatste WK-sheet
    wk_sheet = find_latest_wk_sheet(chase_file)

    # 3) Sheet inlezen
    # dtype=str zodat we alles als tekst hebben (incl. #N/B)
    df = pd.read_excel(chase_file, sheet_name=wk_sheet, dtype=str)

    if df.empty:
        print("⚠ De geselecteerde sheet is leeg.")
        return

    # Debug kolommen
    print("\n=== KOLomnamen in sheet ===")
    for col in df.columns:
        print(f"'{col}' → '{str(col).strip().lower()}'")
    print("===========================\n")

    # 4) Status-kolom zoeken
    status_col = find_status_column(list(df.columns))
    if status_col is None:
        raise ValueError("Geen kolom gevonden waarvan de naam 'status' bevat.")
    print(f"→ Geselecteerde Status-kolom: {status_col}")

    # 5) Filteren op Status
    df_filtered = filter_rows_on_status(df, status_col=status_col)
    if df_filtered.empty:
        print("⚠ Geen rijen gevonden met Status N/B of Mail.")
        return

    # 6) Leverancier-kolom
    supplier_col = get_supplier_column(df_filtered)
    print(f"→ Leverancier-kolom: {supplier_col}")

    # 7) Leveranciers-info
    supplier_info = load_supplier_info(SUPPLIER_INFO_FILE)

    # 8) Per leverancier mail opbouwen
    for supplier, group in df_filtered.groupby(supplier_col):
        supplier_str = str(supplier).strip()
        if not supplier_str:
            continue

        print(f"→ Verwerken leverancier: {supplier_str} (rows: {len(group)})")

        # Basis e-mailadres en kandidaat-rij in leveranciers-info
        to_email = TEST_EMAIL if TESTMODE else TEST_EMAIL  # fallback
        candidates = None

        if supplier_info is not None:
            mask = (
                supplier_info.iloc[:, 0]
                .astype(str).str.strip().str.lower()
                == supplier_str.lower()
            )
            candidates = supplier_info[mask]
            if not candidates.empty and not TESTMODE:
                # In real-mode halen we e-mail uit kolom 2
                to_email = str(candidates.iloc[0, 1])

        # Taal + aanspreeknaam bepalen
        lang, contact_name = detect_language_and_name(
            supplier_info=supplier_info,
            supplier_str=supplier_str,
            candidates=candidates
        )

        # Tabel maken
        html_table = make_html_table_for_group(group)

        # Tekst uit juiste template
        body_without_signature = build_body_from_template(
            lang=lang,
            contact_name=contact_name,
            html_table=html_table
        )

        # Subject per taal
        if lang == "ENG":
            subject = f"Chaselist – open purchase order(s) {supplier_str}"
        else:
            subject = f"Chaselist – openstaande bestelling(en) {supplier_str}"

        print(f"→ Mail ({lang}) klaarzetten naar: {to_email} | Subject: {subject}")
        send_mail(to_email, subject, body_without_signature)

    print("✔ Script afgerond.")


if __name__ == "__main__":
    main()

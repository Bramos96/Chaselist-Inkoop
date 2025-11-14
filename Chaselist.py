import os
import re
from typing import Optional, List

import pandas as pd
import win32com.client as win32

# ───────────────────────────────────────────────────────────
# INSTELLINGEN
# ───────────────────────────────────────────────────────────

BASE_FOLDER = r"C:\Users\bram.gerrits\Desktop\Automations\Chaselist Madelon"
CHASE_PREFIX = "Chase"  # bestand begint met "Chase"

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


# ───────────────────────────────────────────────────────────
# MAIL OPBOUW
# ───────────────────────────────────────────────────────────

def make_html_table_for_group(group: pd.DataFrame) -> str:
    """
    Bouwt een HTML-tabel met kolommen volgens COLUMN_MAPPING.
    'Current delivery date' wordt rood/vet als < vandaag.
    """
    # Eerst checken of alle bronkolommen bestaan
    missing = [src for src, _ in COLUMN_MAPPING if src not in group.columns]
    if missing:
        raise KeyError(f"Kolommen ontbreken in data: {missing}")

    # Voor datumvergelijking
    today = pd.Timestamp.today().normalize()

    # Bouw header
    header_cells = "".join(f"<th>{display}</th>" for _, display in COLUMN_MAPPING)
    rows_html = []

    for _, row in group.iterrows():
        cells_html = []
        for src_name, display_name in COLUMN_MAPPING:
            value = row[src_name]

            # Default: gewoon value tonen
            cell_html = f"{'' if pd.isna(value) else value}"

            # Speciale format voor 'Current delivery date'
            if display_name == "Current delivery date":
                # probeer datum te parsen (NL formaat dd-mm-jjjj)
                dt = pd.to_datetime(value, dayfirst=True, errors="coerce")
                if pd.notna(dt) and dt.normalize() < today:
                    cell_html = f'<span style="color:red; font-weight:bold;">{value}</span>'

            cells_html.append(f"<td>{cell_html}</td>")

        row_html = "<tr>" + "".join(cells_html) + "</tr>"
        rows_html.append(row_html)

    table_html = f"""
    <table border="0" cellspacing="0" cellpadding="4" style="border-collapse:collapse;">
        <thead style="background-color:#d8edf7; font-weight:bold;">
            <tr>{header_cells}</tr>
        </thead>
        <tbody>
            {''.join(rows_html)}
        </tbody>
    </table>
    """
    return table_html


def get_outlook() -> win32.CDispatch:
    return win32.Dispatch("Outlook.Application")


def send_mail(to_addr: str, subject: str, html_body: str):
    outlook = get_outlook()
    mail = outlook.CreateItem(0)  # olMailItem
    mail.To = to_addr
    mail.Subject = subject
    mail.HTMLBody = html_body
    # Tijdens testen:
    mail.Display()
    # Later eventueel: mail.Send()


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

    # 7) Leveranciers-info (voor later, nu nog niet gebruikt)
    supplier_info = load_supplier_info(SUPPLIER_INFO_FILE)

    # 8) Per leverancier mail opbouwen
    for supplier, group in df_filtered.groupby(supplier_col):
        supplier_str = str(supplier).strip()
        if not supplier_str:
            continue

        print(f"→ Verwerken leverancier: {supplier_str} (rows: {len(group)})")

        # E-mailadres
        if TESTMODE:
            to_email = TEST_EMAIL
        else:
            to_email = TEST_EMAIL  # fallback
            if supplier_info is not None:
                mask = (
                    supplier_info.iloc[:, 0]
                    .astype(str).str.strip().str.lower()
                    == supplier_str.lower()
                )
                candidates = supplier_info[mask]
                if not candidates.empty:
                    to_email = str(candidates.iloc[0, 1])

        # Alleen relevante kolommen meenemen
        html_table = make_html_table_for_group(group)

        body = f"""
        <html>
        <body style="font-family: Aptos, Aptos, sans-serif; font-size: 10pt;">
            <p>Hi,</p>
            <p>Onderstaand vind je een overzicht van de openstaande bestelling(en) voor leverancier
               <b>{supplier_str}</b> uit de chaselist.</p>
            {html_table}
            <p>Kun je laten weten wat de status is?</p>
            <p>Alvast dank!</p>
            <p>Vriendelijke groet,<br>
               Bram</p>
            <p style="font-size: 8pt; color: #777;">[Automatisch verstuurd vanuit Python-chaselist]</p>
        </body>
        </html>
        """

        subject = f"Chaselist – openstaande bestelling(en) {supplier_str}"
        print(f"→ Mail klaarzetten naar: {to_email} | Subject: {subject}")
        send_mail(to_email, subject, body)

    print("✔ Script afgerond.")


if __name__ == "__main__":
    main()

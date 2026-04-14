import streamlit as st
import pdfplumber
import re
import io
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import DataPoint
import tempfile
import os

# ─── Page config ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Journal des Appels → Excel",
    page_icon="📞",
    layout="wide",
)

# ─── Custom CSS ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main { background: #f8f9fa; }
    .stApp { font-family: 'Segoe UI', sans-serif; }
    h1 { color: #1a237e; font-weight: 700; }
    .metric-card {
        background: white;
        border-radius: 12px;
        padding: 20px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        text-align: center;
    }
    .metric-value { font-size: 2rem; font-weight: 700; color: #1a237e; }
    .metric-label { font-size: 0.85rem; color: #666; margin-top: 4px; }
    .status-ok { color: #2e7d32; font-weight: 600; }
    .status-ko { color: #c62828; font-weight: 600; }
</style>
""", unsafe_allow_html=True)


# ─── PDF Parsing ─────────────────────────────────────────────────────────────
def parse_pdf(file_bytes: bytes) -> tuple[list[dict], dict]:
    """
    Extract call records from a Logimatique 'Journal des Appels' PDF.
    Returns (records, metadata).
    """
    records = []
    metadata = {}

    # Regex for answered calls:  X  HH:MM  HH:MM  <name>  <phone>  H:MM:SS
    re_answered = re.compile(
        r'^X\s+'
        r'(\d{2}:\d{2})\s+'          # début
        r'(\d{2}:\d{2})\s+'          # fin
        r'(.+?)\s+'                  # adresse
        r'(\d[\d\s]{9,14})\s+'       # téléphone
        r'(\d+:\d{2}:\d{2})\s*$'     # durée
    )
    # Regex for unanswered calls:  HH:MM  <name>  <phone>  (no duration)
    re_unanswered = re.compile(
        r'^(\d{2}:\d{2})\s+'
        r'(.+?)\s+'
        r'(\d[\d\s]{9,14})\s*$'
    )
    # Metadata header
    re_date = re.compile(r'du:\s*(\d{2}/\d{2}/\d{4})')
    re_society = re.compile(r'Société:\s*(\S+)')

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for raw_line in text.splitlines():
                line = raw_line.strip()
                if not line:
                    continue

                # Meta
                m = re_date.search(line)
                if m:
                    metadata.setdefault("date", m.group(1))
                m = re_society.search(line)
                if m:
                    metadata.setdefault("societe", m.group(1))

                # Summary line
                m = re.search(r"Nombre d'Appels Répondus\s+(\d+)", line)
                if m:
                    metadata["repondus"] = int(m.group(1))
                m = re.search(r"Nombre d'Appels Sans Réponse\s+(\d+)", line)
                if m:
                    metadata["sans_reponse"] = int(m.group(1))
                m = re.search(r"Nombre d'Appels Total\s+(\d+)", line)
                if m:
                    metadata["total"] = int(m.group(1))

                # Answered
                m = re_answered.match(line)
                if m:
                    debut, fin, adresse, telephone, duree = m.groups()
                    telephone = telephone.strip()
                    absent = "<Absent" in adresse
                    records.append({
                        "Statut": "Répondu",
                        "Heure début": debut,
                        "Heure fin": fin,
                        "Adresse / Client": adresse.replace(" <Absent dans HY - DL>", "").strip(),
                        "Téléphone": telephone,
                        "Durée": duree,
                        "Absent répertoire": "Oui" if absent else "Non",
                    })
                    continue

                # Unanswered
                m = re_unanswered.match(line)
                if m:
                    debut, adresse, telephone = m.groups()
                    telephone = telephone.strip()
                    absent = "<Absent" in adresse
                    records.append({
                        "Statut": "Sans réponse",
                        "Heure début": debut,
                        "Heure fin": "",
                        "Adresse / Client": adresse.replace(" <Absent dans HY - DL>", "").strip(),
                        "Téléphone": telephone,
                        "Durée": "",
                        "Absent répertoire": "Oui" if absent else "Non",
                    })

    return records, metadata


# ─── Excel Builder ────────────────────────────────────────────────────────────
def build_excel(records: list[dict], metadata: dict, date_str: str) -> bytes:
    wb = Workbook()

    # ── Palette ──────────────────────────────────────────────────────────────
    NAVY      = "1A237E"
    ORANGE    = "E65100"
    GREEN     = "2E7D32"
    RED_SOFT  = "C62828"
    LIGHT_BG  = "EEF2FF"
    WHITE     = "FFFFFF"
    GREY_BG   = "F5F5F5"
    GREY_LINE = "BDBDBD"

    thin  = Side(border_style="thin",   color=GREY_LINE)
    thick = Side(border_style="medium", color=NAVY)
    border_data  = Border(left=thin, right=thin, top=thin, bottom=thin)
    border_top   = Border(left=thin, right=thin, top=thick, bottom=thin)
    border_bottom= Border(left=thin, right=thin, top=thin, bottom=thick)

    def hdr_font(size=10, bold=True, color=WHITE):
        return Font(name="Arial", bold=bold, size=size, color=color)

    def cell_font(size=9, bold=False, color="212121"):
        return Font(name="Arial", bold=bold, size=size, color=color)

    def fill(hex_color):
        return PatternFill("solid", fgColor=hex_color)

    # ── Sheet 1 – Journal complet ─────────────────────────────────────────────
    ws = wb.active
    ws.title = "Journal des appels"
    ws.sheet_view.showGridLines = False

    # Title block
    ws.merge_cells("A1:G1")
    ws["A1"] = f"📞  HYMPYR — Journal des Appels  ·  {date_str}"
    ws["A1"].font = Font(name="Arial", bold=True, size=14, color=WHITE)
    ws["A1"].fill = fill(NAVY)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36

    # KPI row
    kpi_labels = [
        ("Total appels",     metadata.get("total",        len(records))),
        ("Répondus",         metadata.get("repondus",     sum(1 for r in records if r["Statut"]=="Répondu"))),
        ("Sans réponse",     metadata.get("sans_reponse", sum(1 for r in records if r["Statut"]=="Sans réponse"))),
        ("Taux de réponse",  f"{metadata.get('repondus',0) / max(metadata.get('total',1),1)*100:.1f} %"),
    ]
    ws.merge_cells("A2:B2"); ws.merge_cells("C2:D2")
    ws.merge_cells("E2:F2"); ws.merge_cells("G2:G2")
    kpi_cols = ["A", "C", "E", "G"]
    kpi_fills = [NAVY, ORANGE, RED_SOFT, GREEN]
    for i, ((label, val), col, fc) in enumerate(zip(kpi_labels, kpi_cols, kpi_fills)):
        ws[f"{col}2"] = f"{label}: {val}"
        ws[f"{col}2"].font = Font(name="Arial", bold=True, size=10, color=WHITE)
        ws[f"{col}2"].fill = fill(fc)
        ws[f"{col}2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 28

    ws.row_dimensions[3].height = 6  # spacer

    # Column headers
    headers = ["Statut", "Heure début", "Heure fin", "Adresse / Client", "Téléphone", "Durée", "Absent répertoire"]
    col_widths = [14, 13, 11, 42, 18, 10, 18]
    for col_idx, (h, w) in enumerate(zip(headers, col_widths), start=1):
        cell = ws.cell(row=4, column=col_idx, value=h)
        cell.font = hdr_font()
        cell.fill = fill(NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_top
        ws.column_dimensions[get_column_letter(col_idx)].width = w
    ws.row_dimensions[4].height = 22
    ws.freeze_panes = "A5"

    # Data rows
    for row_idx, rec in enumerate(records, start=5):
        answered = rec["Statut"] == "Répondu"
        row_fill = fill(WHITE) if row_idx % 2 == 0 else fill(GREY_BG)
        for col_idx, key in enumerate(headers, start=1):
            val = rec.get(key, "")
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = cell_font()
            cell.fill = row_fill
            cell.alignment = Alignment(vertical="center")
            cell.border = border_data

            # Statut coloring
            if col_idx == 1:
                if answered:
                    cell.font = Font(name="Arial", size=9, bold=True, color=GREEN)
                else:
                    cell.font = Font(name="Arial", size=9, bold=True, color=RED_SOFT)
                cell.alignment = Alignment(horizontal="center", vertical="center")

            # Phone column – centered mono
            if col_idx == 5:
                cell.font = Font(name="Courier New", size=9, color="212121")
                cell.alignment = Alignment(horizontal="center", vertical="center")

            # Times – center
            if col_idx in (2, 3, 6):
                cell.alignment = Alignment(horizontal="center", vertical="center")

            # Absent – center + color
            if col_idx == 7:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if val == "Oui":
                    cell.font = Font(name="Arial", size=9, color=RED_SOFT)

        ws.row_dimensions[row_idx].height = 18

    # Total row
    total_row = len(records) + 5
    ws.merge_cells(f"A{total_row}:C{total_row}")
    ws[f"A{total_row}"] = f"TOTAL : {len(records)} appels"
    ws[f"A{total_row}"].font = Font(name="Arial", bold=True, size=9, color=WHITE)
    ws[f"A{total_row}"].fill = fill(NAVY)
    ws[f"A{total_row}"].alignment = Alignment(horizontal="center", vertical="center")
    ws[f"A{total_row}"].border = border_bottom
    for c in range(4, 8):
        ws.cell(total_row, c).fill = fill(NAVY)
        ws.cell(total_row, c).border = border_bottom
    ws.row_dimensions[total_row].height = 20

    # Auto-filter
    ws.auto_filter.ref = f"A4:G{total_row - 1}"

    # ── Sheet 2 – Synthèse ────────────────────────────────────────────────────
    ws2 = wb.create_sheet("Synthèse")
    ws2.sheet_view.showGridLines = False

    ws2.merge_cells("A1:D1")
    ws2["A1"] = "Synthèse — Journal des Appels"
    ws2["A1"].font = Font(name="Arial", bold=True, size=13, color=WHITE)
    ws2["A1"].fill = fill(NAVY)
    ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 32

    # KPIs
    kpi_rows = [
        ("Date du journal",   date_str),
        ("Total appels",      metadata.get("total", len(records))),
        ("Appels répondus",   metadata.get("repondus", sum(1 for r in records if r["Statut"]=="Répondu"))),
        ("Appels sans réponse", metadata.get("sans_reponse", sum(1 for r in records if r["Statut"]=="Sans réponse"))),
        ("Taux de réponse",   f"=C4/C3"),
        ("Durée moy. (répondus)", ""),  # placeholder
    ]

    # Compute average duration
    durations = []
    for r in records:
        d = r.get("Durée", "")
        if d:
            parts = d.split(":")
            try:
                secs = int(parts[0])*3600 + int(parts[1])*60 + int(parts[2])
                durations.append(secs)
            except Exception:
                pass
    avg_sec = int(sum(durations)/len(durations)) if durations else 0
    avg_fmt = f"{avg_sec//60}m {avg_sec%60:02d}s"
    kpi_rows[-1] = ("Durée moy. (répondus)", avg_fmt)

    ws2.column_dimensions["A"].width = 2
    ws2.column_dimensions["B"].width = 28
    ws2.column_dimensions["C"].width = 20
    ws2.column_dimensions["D"].width = 2

    for i, (label, val) in enumerate(kpi_rows, start=2):
        # label cell
        lc = ws2.cell(row=i+1, column=2, value=label)
        lc.font = Font(name="Arial", size=10, bold=True, color="212121")
        lc.fill = fill(LIGHT_BG)
        lc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        lc.border = border_data
        # value cell
        vc = ws2.cell(row=i+1, column=3, value=val)
        vc.font = Font(name="Arial", size=10, color="212121")
        vc.fill = fill(WHITE)
        vc.alignment = Alignment(horizontal="center", vertical="center")
        vc.border = border_data
        ws2.row_dimensions[i+1].height = 24
        # Percent format for taux
        if label == "Taux de réponse":
            vc.number_format = "0.0%"
            vc.value = metadata.get("repondus", 0) / max(metadata.get("total", 1), 1)
            vc.font = Font(name="Arial", size=10, bold=True, color=GREEN)

    # Hourly distribution table on ws2
    from collections import Counter
    hours_ok  = Counter()
    hours_ko  = Counter()
    for r in records:
        h = r["Heure début"][:2] if r["Heure début"] else None
        if h:
            if r["Statut"] == "Répondu":
                hours_ok[h] += 1
            else:
                hours_ko[h] += 1

    all_hours = sorted(set(list(hours_ok.keys()) + list(hours_ko.keys())))

    start_row = 11
    ws2.merge_cells(f"B{start_row}:C{start_row}")
    ws2[f"B{start_row}"] = "Répartition par heure"
    ws2[f"B{start_row}"].font = Font(name="Arial", bold=True, size=10, color=WHITE)
    ws2[f"B{start_row}"].fill = fill(ORANGE)
    ws2[f"B{start_row}"].alignment = Alignment(horizontal="center")
    ws2[f"B{start_row}"].border = border_data

    for j, hdr in enumerate(["Heure", "Répondus", "Sans réponse"], start=2):
        cell = ws2.cell(row=start_row+1, column=j, value=hdr)
        cell.font = hdr_font()
        cell.fill = fill(NAVY)
        cell.alignment = Alignment(horizontal="center")
        cell.border = border_data

    for k, h in enumerate(all_hours, start=start_row+2):
        ws2.cell(k, 2, f"{h}h").font = cell_font(bold=True)
        ws2.cell(k, 2).fill = fill(GREY_BG)
        ws2.cell(k, 2).alignment = Alignment(horizontal="center")
        ws2.cell(k, 2).border = border_data
        ws2.cell(k, 3, hours_ok.get(h, 0)).fill = fill(WHITE)
        ws2.cell(k, 3).alignment = Alignment(horizontal="center")
        ws2.cell(k, 3).border = border_data
        ws2.cell(k, 4, hours_ko.get(h, 0)).fill = fill(WHITE)
        ws2.cell(k, 4).alignment = Alignment(horizontal="center")
        ws2.cell(k, 4).border = border_data
        ws2.row_dimensions[k].height = 18

    # ── Sheet 3 – Inconnus ────────────────────────────────────────────────────
    ws3 = wb.create_sheet("Numéros inconnus")
    ws3.sheet_view.showGridLines = False

    ws3.merge_cells("A1:C1")
    ws3["A1"] = "Numéros absents du répertoire"
    ws3["A1"].font = Font(name="Arial", bold=True, size=12, color=WHITE)
    ws3["A1"].fill = fill(RED_SOFT)
    ws3["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[1].height = 30

    unknowns = [r for r in records if r["Absent répertoire"] == "Oui"]
    for j, h in enumerate(["Statut", "Heure début", "Téléphone"], start=1):
        cell = ws3.cell(row=2, column=j, value=h)
        cell.font = hdr_font()
        cell.fill = fill(NAVY)
        cell.alignment = Alignment(horizontal="center")
        cell.border = border_data

    ws3.column_dimensions["A"].width = 14
    ws3.column_dimensions["B"].width = 13
    ws3.column_dimensions["C"].width = 20

    for i, r in enumerate(unknowns, start=3):
        answered = r["Statut"] == "Répondu"
        bg = fill(WHITE) if i % 2 == 0 else fill(GREY_BG)
        ws3.cell(i, 1, r["Statut"]).font = Font(name="Arial", size=9, bold=True,
                                                 color=GREEN if answered else RED_SOFT)
        ws3.cell(i, 1).fill = bg; ws3.cell(i, 1).alignment = Alignment(horizontal="center")
        ws3.cell(i, 1).border = border_data
        ws3.cell(i, 2, r["Heure début"]).fill = bg
        ws3.cell(i, 2).alignment = Alignment(horizontal="center"); ws3.cell(i, 2).border = border_data
        ws3.cell(i, 3, r["Téléphone"]).font = Font(name="Courier New", size=9)
        ws3.cell(i, 3).fill = bg; ws3.cell(i, 3).alignment = Alignment(horizontal="center")
        ws3.cell(i, 3).border = border_data
        ws3.row_dimensions[i].height = 18

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ─── UI ──────────────────────────────────────────────────────────────────────
st.title("📞 Journal des Appels → Excel")
st.markdown("**Importez un PDF Logimatique** pour obtenir un fichier Excel professionnel structuré en 3 onglets.")

st.divider()

uploaded = st.file_uploader(
    "Déposer le fichier PDF ici",
    type=["pdf"],
    help="Journal des appels exporté depuis Logimatique (format HY / HYMPYR)"
)

if uploaded:
    with st.spinner("Analyse du PDF en cours…"):
        pdf_bytes = uploaded.read()
        records, metadata = parse_pdf(pdf_bytes)

    if not records:
        st.error("Aucun appel détecté. Vérifiez que le fichier est un Journal des Appels Logimatique.")
        st.stop()

    # ── KPI Cards ────────────────────────────────────────────────────────────
    total     = metadata.get("total", len(records))
    repondus  = metadata.get("repondus", sum(1 for r in records if r["Statut"]=="Répondu"))
    sans_rep  = metadata.get("sans_reponse", total - repondus)
    taux      = repondus / max(total, 1) * 100
    date_str  = metadata.get("date", "—")

    c1, c2, c3, c4 = st.columns(4)
    for col, label, val, color in [
        (c1, "Total appels",    total,     "#1a237e"),
        (c2, "Répondus",       repondus,  "#2e7d32"),
        (c3, "Sans réponse",   sans_rep,  "#c62828"),
        (c4, "Taux de réponse", f"{taux:.1f} %", "#e65100"),
    ]:
        col.markdown(f"""
        <div class="metric-card">
            <div class="metric-value" style="color:{color}">{val}</div>
            <div class="metric-label">{label}</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown(f"<br>**Date du journal :** {date_str} &nbsp;|&nbsp; **Société :** {metadata.get('societe','—')}", unsafe_allow_html=True)
    st.divider()

    # ── DataFrame Preview ─────────────────────────────────────────────────────
    df = pd.DataFrame(records)
    
    tab1, tab2, tab3 = st.tabs(["📋 Tous les appels", "✅ Répondus", "❌ Sans réponse"])
    with tab1:
        st.dataframe(df, use_container_width=True, height=420)
    with tab2:
        st.dataframe(df[df["Statut"]=="Répondu"], use_container_width=True, height=420)
    with tab3:
        st.dataframe(df[df["Statut"]=="Sans réponse"], use_container_width=True, height=420)

    st.divider()

    # ── Export ────────────────────────────────────────────────────────────────
    with st.spinner("Génération du fichier Excel…"):
        xlsx_bytes = build_excel(records, metadata, date_str)

    filename = f"Journal_Appels_HYMPYR_{date_str.replace('/','')}.xlsx"

    st.download_button(
        label="⬇️  Télécharger le fichier Excel",
        data=xlsx_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )

    st.caption(f"Fichier : **{filename}**  ·  3 onglets : Journal complet · Synthèse · Numéros inconnus")

else:
    st.info("👆 Importez un fichier PDF pour commencer.")
    st.markdown("""
    **Contenu du fichier Excel généré :**
    - **Onglet 1 – Journal des appels** : tableau complet avec filtre automatique, alternance de lignes, codes couleur Répondu/Sans réponse
    - **Onglet 2 – Synthèse** : KPIs, taux de réponse, répartition horaire
    - **Onglet 3 – Numéros inconnus** : liste des appelants absents du répertoire
    """)
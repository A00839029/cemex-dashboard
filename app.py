# app.py
# -*- coding: utf-8 -*-
from pathlib import Path
from datetime import datetime
import pandas as pd
import streamlit as st
import plotly.express as px
from reportlab.lib.pagesizes import landscape, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER

# ==========================
# CONFIGURACIÓN BASE
# ==========================
BASE = Path("/Users/joseluisgiadanscastellanos/Library/CloudStorage/OneDrive-CEMEX/OneDrive_Cemex")
OUT_FILE = BASE / "Codigos/out/trafos_maestro_tabla.xlsx"
LOGO_FILE = BASE / "Codigos/out/cemex_logo.png"

st.set_page_config(page_title="Dashboard Transformadores CEMEX", layout="wide", page_icon="⚡")

PALETTE = {"green": "#7BC47F", "yellow": "#FFD166", "red": "#EF476F"}

# ==========================
# LOGIN
# ==========================
def check_login():
    st.session_state["logged_in"] = st.session_state.get("logged_in", False)
    if not st.session_state["logged_in"]:
        st.title("🔒 Acceso restringido")

        username = st.text_input("Usuario", "")
        password = st.text_input("Contraseña", "", type="password")

        if st.button("Iniciar sesión"):
            if username == "cemex" and password == "1234":
                st.session_state["logged_in"] = True
                st.success("✅ Acceso concedido")
                st.rerun()
            else:
                st.error("🚫 Usuario o contraseña incorrectos")

        st.stop()

check_login()

# ==========================
# FUNCIONES AUXILIARES
# ==========================
def normalize(txt: str):
    if not isinstance(txt, str):
        return ""
    import unicodedata
    return (
        unicodedata.normalize("NFKD", txt)
        .encode("ascii", "ignore")
        .decode("utf-8")
        .lower()
        .strip()
    )

def color_alerta(txt: str):
    t = normalize(txt)
    if any(x in t for x in ["critico", "t3", "d2"]):
        return f"background-color:{PALETTE['red']};color:#101010;"
    if any(x in t for x in ["preoc", "t1", "t2", "d1", "pd", "dt", "indeter"]):
        return f"background-color:{PALETTE['yellow']};color:#101010;"
    if "normal" in t:
        return f"background-color:{PALETTE['green']};color:#101010;"
    return ""

# ==========================
# CARGA DE DATOS
# ==========================
@st.cache_data(show_spinner=False)
def load_data(path):
    """Carga las hojas del Excel y crea las que falten automáticamente, evitando duplicados."""
    try:
        xls = pd.ExcelFile(path)
        sheets = xls.sheet_names
    except Exception as e:
        st.error(f"❌ Error al abrir el archivo: {e}")
        return pd.DataFrame()

    def get_sheet(name, cols=None):
        if name in sheets:
            df = pd.read_excel(path, sheet_name=name)
            if cols:
                for c in cols:
                    if c not in df.columns:
                        df[c] = ""
            return df
        else:
            st.warning(f"⚠️ Hoja '{name}' no encontrada. Se generará vacía.")
            base_cols = ["Planta", "Transformador", "Ubicacion"]
            if cols:
                base_cols += cols
            return pd.DataFrame(columns=base_cols)

    estados = get_sheet("Estados", ["Diagnóstico IEEE"])
    diag3 = get_sheet("Diag_3Ratios", ["R1 (C2H2/C2H4)", "R2 (CH4/H2)", "R3 (C2H4/C2H6)", "Diagnóstico 3 Ratios"])
    iec = get_sheet("Diag_IEC", ["Diagnóstico IEC"])
    duval = get_sheet("Diag_Duval", ["Diagnóstico Duval"])

    key = ["Planta", "Transformador", "Ubicacion"]

    # Limpieza de duplicados
    def clean_cols(df, keep_cols):
        df = df.loc[:, ~df.columns.duplicated()]
        extras = [c for c in df.columns if c not in keep_cols]
        return df

    estados = clean_cols(estados, key + ["Fecha de Muestra", "Diagnóstico IEEE"])
    diag3 = clean_cols(diag3, key + ["R1 (C2H2/C2H4)", "R2 (CH4/H2)", "R3 (C2H4/C2H6)", "Diagnóstico 3 Ratios"])
    iec = clean_cols(iec, key + ["Diagnóstico IEC"])
    duval = clean_cols(duval, key + ["Diagnóstico Duval"])

    # Eliminar Fecha duplicada
    for df in [diag3, iec, duval]:
        if "Fecha de Muestra" in df.columns:
            df = df.drop(columns=["Fecha de Muestra"])

    df = estados.merge(diag3, on=key, how="outer", suffixes=("", "_3R"))
    df = df.merge(iec, on=key, how="outer", suffixes=("", "_IEC"))
    df = df.merge(duval, on=key, how="outer", suffixes=("", "_DUV"))

    if "Diagnóstico IEEE" not in df.columns:
        df["Diagnóstico IEEE"] = "Indeterminado"

    df.loc[
        df["Diagnóstico IEEE"].str.contains("Normal", case=False, na=False),
        ["Diagnóstico 3 Ratios", "Diagnóstico IEC", "Diagnóstico Duval"],
    ] = "Normal"

    def final_y_fia(row):
        ieee = normalize(row.get("Diagnóstico IEEE", ""))
        if "normal" in ieee:
            return "Normal (100%)", 100
        if "critico" in ieee:
            return "Crítico (100%)", 100
        if "preoc" in ieee:
            return "Preocupante (85%)", 85
        return "Indeterminado", 70

    df[["Diagnóstico Final", "Fiabilidad"]] = df.apply(final_y_fia, axis=1, result_type="expand")
    return df

# ==========================
# PDF BUILDER
# ==========================
def add_footer(canvas_doc, doc):
    canvas_doc.saveState()
    footer_text = "© CEMEX — Reporte generado automáticamente por Dashboard DGA"
    canvas_doc.setFont("Helvetica", 7)
    canvas_doc.setFillColor(colors.gray)
    canvas_doc.drawCentredString(415, 20, footer_text)
    page_num = canvas_doc.getPageNumber()
    canvas_doc.setFont("Helvetica", 8)
    canvas_doc.setFillColor(colors.black)
    canvas_doc.drawRightString(790, 560, f"Página {page_num}")
    canvas_doc.restoreState()

def build_pdf(df: pd.DataFrame, outfile: Path):
    buffer = str(outfile)
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        leftMargin=40,
        rightMargin=40,
        topMargin=50,
        bottomMargin=45,
    )
    story = []
    styles = getSampleStyleSheet()
    wrap_style = ParagraphStyle("Wrap", fontName="Helvetica", fontSize=8, alignment=TA_CENTER, leading=10)

    if LOGO_FILE.exists():
        logo = Image(str(LOGO_FILE), width=130, height=45)
        logo.hAlign = "CENTER"
        story.append(logo)

    story.append(Spacer(1, 10))
    story.append(Paragraph("<b><font size=16>Reporte de Transformadores en Riesgo</font></b>", styles["Title"]))
    story.append(Spacer(1, 8))
    story.append(Paragraph(f"<font size=10>Generado el {datetime.now().strftime('%d/%m/%Y %H:%M')}</font>", styles["Normal"]))
    story.append(Spacer(1, 20))

    for planta in sorted(df["Planta"].dropna().unique()):
        sub = df[df["Planta"] == planta].copy()
        sub = sub[sub["Diagnóstico IEEE"].isin(["Preocupante", "Crítico"])]
        if sub.empty:
            continue

        story.append(Paragraph(f"<b><font size=13>Planta: {planta}</font></b>", styles["Heading2"]))
        story.append(Spacer(1, 6))

        data = [["Transformador", "Ubicación", "IEEE", "3 Ratios", "IEC", "Duval", "Final", "Fiabilidad"]]
        for _, r in sub.iterrows():
            data.append([
                Paragraph(str(r["Transformador"]), wrap_style),
                Paragraph(str(r["Ubicacion"]), wrap_style),
                Paragraph(str(r["Diagnóstico IEEE"]), wrap_style),
                Paragraph(str(r["Diagnóstico 3 Ratios"]), wrap_style),
                Paragraph(str(r["Diagnóstico IEC"]), wrap_style),
                Paragraph(str(r["Diagnóstico Duval"]), wrap_style),
                Paragraph(str(r["Diagnóstico Final"]), wrap_style),
                Paragraph(f"{int(round(r['Fiabilidad']))}%", wrap_style),
            ])

        t = Table(data, colWidths=[100, 90, 55, 70, 70, 80, 80, 60], repeatRows=1)
        t.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.black),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E9ECEF")),
        ]))
        story.append(t)
        story.append(PageBreak())

    doc.build(story, onLaterPages=add_footer, onFirstPage=add_footer)
    return outfile

# ==========================
# INTERFAZ STREAMLIT
# ==========================
col1, col2 = st.columns([1, 6])
if LOGO_FILE.exists():
    col1.image(str(LOGO_FILE), width=130)
col2.markdown("<h1 style='margin-bottom:0'>Dashboard de Transformadores CEMEX</h1>", unsafe_allow_html=True)
col2.markdown("<p style='margin-top:0;color:gray'>Monitoreo DGA · 3 Ratios · IEC 60599 · Duval</p>", unsafe_allow_html=True)

with st.sidebar:
    st.header("Panel de control")
    if st.button("🔄 Refresh / Actualizar"):
        st.cache_data.clear()
        st.rerun()
    df = load_data(OUT_FILE)
    q = st.text_input("Buscar (Planta / Transformador / Ubicación):", "")
    plantas = sorted(df["Planta"].dropna().unique().tolist())
    sel_plants = st.multiselect("Planta(s):", plantas, default=plantas)
    estados = ["Normal", "Preocupante", "Crítico"]
    sel_ieee = st.multiselect("Diagnóstico IEEE:", estados, default=estados)

tab1, tab2 = st.tabs(["📊 Resumen general", "🔍 Diagnóstico detallado"])

# ==========================
# TAB 1 — RESUMEN GENERAL
# ==========================
with tab1:
    F = df.copy()
    if q:
        F = F[F.apply(lambda r: q.lower() in str(r).lower(), axis=1)]
    F = F[F["Planta"].isin(sel_plants) & F["Diagnóstico IEEE"].isin(sel_ieee)]

    c1, c2, c3, c4 = st.columns(4)
    total = len(F)
    crit = F["Diagnóstico IEEE"].str.contains("Crít", case=False, na=False).sum()
    prec = F["Diagnóstico IEEE"].str.contains("Preoc", case=False, na=False).sum()
    norm = F["Diagnóstico IEEE"].str.contains("Normal", case=False, na=False).sum()
    c1.metric("Total monitoreados", total)
    c2.metric("Críticos", crit)
    c3.metric("Preocupantes", prec)
    c4.metric("Normales", norm)

    col1, col2 = st.columns(2)
    fig1 = px.pie(
        F, names="Diagnóstico IEEE", title="Distribución IEEE",
        color="Diagnóstico IEEE",
        color_discrete_map={"Crítico": PALETTE["red"], "Preocupante": PALETTE["yellow"], "Normal": PALETTE["green"]},
    )
    col1.plotly_chart(fig1, use_container_width=True)

    fig2 = px.histogram(
        F, x="Planta", color="Diagnóstico IEEE", title="Diagnósticos por planta",
        color_discrete_sequence=[PALETTE["green"], PALETTE["yellow"], PALETTE["red"]],
    )
    col2.plotly_chart(fig2, use_container_width=True)

    st.subheader("Detalle de Transformadores")
    mostrar = ["Planta", "Transformador", "Ubicacion", "Diagnóstico IEEE",
               "Diagnóstico 3 Ratios", "Diagnóstico IEC", "Diagnóstico Duval", "Diagnóstico Final"]
    tabla = F[mostrar].copy()
    styled = tabla.style.map(lambda v: color_alerta(v), subset=mostrar[3:])
    st.dataframe(styled, use_container_width=True)

    st.subheader("📄 Exportar PDF profesional (todas las plantas)")
    if st.button("Generar PDF"):
        outfile = Path.cwd() / f"reporte_transformadores_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
        res = build_pdf(df, outfile)
        st.success("✅ PDF generado correctamente.")
        st.download_button("Descargar PDF", data=open(res, "rb").read(), file_name=res.name, mime="application/pdf")

# ==========================
# TAB 2 — DETALLE
# ==========================
with tab2:
    st.subheader("Diagnóstico detallado (1 transformador)")
    col1, col2 = st.columns(2)
    with col1:
        planta = st.selectbox("Planta", sorted(df["Planta"].dropna().unique()))
    with col2:
        trs = sorted(df[df["Planta"] == planta]["Transformador"].dropna().unique())
        trafo = st.selectbox("Transformador", trs)

    row = df[(df["Planta"] == planta) & (df["Transformador"] == trafo)].head(1)
    if not row.empty:
        r = row.iloc[0]
        st.markdown(f"### Diagnóstico IEEE: **{r['Diagnóstico IEEE']}** — Diagnóstico Final: **{r['Diagnóstico Final']}**")
        st.markdown("#### Método de 3 Ratios")
        ratios = pd.DataFrame({
            "Cociente": ["R1 Acetileno / Etileno", "R2 Metano / Hidrógeno", "R3 Etileno / Etano", "Resultado final"],
            "Valor": [r.get("R1 (C2H2/C2H4)"), r.get("R2 (CH4/H2)"), r.get("R3 (C2H4/C2H6)"), r.get("Diagnóstico 3 Ratios")],
        })
        st.dataframe(ratios, use_container_width=True)

# ==========================
# EJECUCIÓN EN TERMINAL
# ==========================
# Para correr:
# streamlit run /Users/joseluisgiadanscastellanos/Library/CloudStorage/OneDrive-CEMEX/Codigos/app.py
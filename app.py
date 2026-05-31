import streamlit as st
import pandas as pd
import zipfile
import re
from datetime import date
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth

# ---------------- CONFIG ----------------
st.set_page_config(
    page_title="Certificados de Retención",
    page_icon="📄"
)

# ---------------- BARRA SUPERIOR ----------------
st.markdown(
    '<div style="background-color:#2E86C1;padding:12px;border-radius:8px;">'
    '<div style="display:flex;justify-content:space-between;">'
    '<div style="color:white;font-weight:bold;font-size:18px;">MASIZO SAS</div>'
    '<div style="color:white;">Wilfredo Diaz | Soporte</div>'
    '</div>'
    '<div style="margin-top:8px;color:white;font-weight:bold;">📄 Certificados de Retención</div>'
    '</div>',
    unsafe_allow_html=True
)

# ---------------- BOTONES ARRIBA ----------------
st.write("")
col1, col2, col3, col4 = st.columns(4)

with col1:
    btn_limpiar = st.button("🔄 Limpiar", key="btn_limpiar_top")

with col2:
    btn_excel = st.button("📥 Excel", key="btn_excel_top")

with col3:
    btn_opciones = st.button("⚙️ Opciones", key="btn_opciones_top")

with col4:
    btn_generar = st.button("✅ Generar", key="btn_generar_top")

# ---------------- UI ----------------
st.markdown("<h2 style='text-align: center;'>📄 Certificados de Retención</h2>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center;'>Generación automática de certificados tributarios</p>", unsafe_allow_html=True)

st.divider()

ANIO = st.number_input("Año gravable", value=2026)
NOMBRE_EMPRESA = st.text_input("Nombre empresa", value="MASIZO SAS")

CIUDAD_CONSIGNACION = st.text_input("Ciudad de consignación", value="Bogotá").strip()

if not CIUDAD_CONSIGNACION:
    st.warning("⚠️ Debes ingresar la ciudad de consignación")
    st.stop()

fecha_emision = st.date_input("Fecha de emisión", value=date.today())
fecha_texto = fecha_emision.strftime("%d de %b de %Y")



# ---------------- FUNCIONES ----------------

def periodo_a_mes(periodo):
    meses = {
        "01": "Enero","02": "Febrero","03": "Marzo",
        "04": "Abril","05": "Mayo","06": "Junio",
        "07": "Julio","08": "Agosto","09": "Septiembre",
        "10": "Octubre","11": "Noviembre","12": "Diciembre"
    }
    texto = str(periodo)
    match = re.search(r'(20\d{2})[-]?(\d{2})', texto)
    if match:
        return meses.get(match.group(2), match.group(2))
    return texto

def escribir_parrafo(texto, y, c):
    ancho_max = 550 - (57 + 42)
    palabras = texto.split()
    linea = ""

    c.setFont("Helvetica", 8)

    for palabra in palabras:
        prueba = linea + " " + palabra if linea else palabra

        if stringWidth(prueba, "Helvetica", 8) < ancho_max:
            linea = prueba
        else:
            x = 300 - stringWidth(linea, "Helvetica", 8)/2
            c.drawString(x, y, linea)
            y -= 10
            linea = palabra

    if linea:
        x = 300 - stringWidth(linea, "Helvetica", 8)/2
        c.drawString(x, y, linea)
        y -= 10

    return y

def titulo(tipo):
    if tipo == "Retefuente":
        return "CERTIFICADO DE RETENCIONES EN LA FUENTE"
    elif tipo == "ReteICA":
        return "CERTIFICADO DE RETENCIONES DE ICA"
    elif tipo == "ReteIVA":
        return "CERTIFICADO DE RETENCIONES DE IVA"
    return "CERTIFICADO"


archivo = st.file_uploader("Sube el auxiliar", type=["xlsx"])

if archivo:
    st.success("✅ Archivo cargado correctamente")

# ---------------- PROCESO ----------------

if archivo:

    df = pd.read_excel(archivo, engine="openpyxl", skiprows=10)

    df.columns = [
        "NitEmpresa","Cuenta","Nombre","Fecha",
        "Fecha2","Nit","Tercero","Descripcion",
        "Debito","Credito","Actividad",
        "Periodo","Proyecto","Tasa",
        "Base","Concepto"
    ]

    df["Nit"] = df["Nit"].astype(str).str.strip()
    df["Tercero"] = df["Tercero"].astype(str).str.strip()
    df["Periodo"] = df["Periodo"].astype(str).str.strip()

    df = df[df["Cuenta"].astype(str).str.startswith("236")]

    tipos_sel = st.multiselect(
        "Tipos de certificados",
        ["Retefuente","ReteIVA","ReteICA"],
        default=["Retefuente","ReteICA"]
    )

    periodos = sorted(df["Periodo"].dropna().unique())

    c1, c2 = st.columns(2)
    inicio = c1.selectbox("Desde periodo", periodos, 0)
    fin = c2.selectbox("Hasta periodo", periodos, len(periodos)-1)

    df = df[(df["Periodo"] >= inicio) & (df["Periodo"] <= fin)]

# -------- FILTRO POR TERCERO --------
df_base = df.copy()

nombre_input = st.text_input("Buscar tercero")

if nombre_input:
    df_base = df_base[df_base["Tercero"].str.contains(nombre_input, case=False, na=False)]

terceros = sorted(df_base["Tercero"].dropna().unique())
tercero_sel = st.selectbox("Seleccionar tercero", ["Todos"] + terceros)

if tercero_sel != "Todos":
    df = df_base[df_base["Tercero"] == tercero_sel]
else:
    df = df_base


    df["TipoRet"] = df["Cuenta"].apply(
        lambda c: "Retefuente" if str(c).startswith("2365")
        else "ReteIVA" if str(c).startswith("2367")
        else "ReteICA" if str(c).startswith("2368")
        else "Otros"
    )

    df = df[df["TipoRet"].isin(tipos_sel)]

    df["Retencion"] = (df["Credito"] - df["Debito"]).abs()
    df = df[df["Retencion"] > 0]

    agrupado = df.groupby(["Nit","Tercero","TipoRet","Concepto"]).agg({
        "Retencion":"sum"
    }).reset_index()

    st.dataframe(agrupado)

# ---------------- ACCIONES ----------------

if btn_limpiar:
    st.rerun()

if btn_excel:
    if 'agrupado' in locals():
        import io
        buffer = io.BytesIO()
        agrupado.to_excel(buffer, index=False, engine="openpyxl")
        buffer.seek(0)

        st.download_button("Descargar Excel", buffer, "retenciones.xlsx")
    else:
        st.warning("⚠️ No hay datos")

if btn_opciones:
    st.info("Opciones en desarrollo")

if btn_generar:
    if 'agrupado' not in locals():
        st.warning("⚠️ Carga archivo primero")
    else:
        st.success("✅ Aquí conectas PDF (ya lo tienes arriba listo)")
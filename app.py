import streamlit as st
import pandas as pd
import zipfile
import re
from datetime import date
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth

# ---------------- CONFIG ----------------
MARGEN_IZQ = 57   # 2.0 cm
MARGEN_DER = 42   # 1.5 cm
ANCHO = 550

COL_BASE = 310
COL_TARIFA = 390
COL_RET = ANCHO - MARGEN_DER - 5

# ---------------- UI ----------------
nombre_empresa_excel = ""

st.title("📄 Generador Certificados de Retención")

if nombre_empresa_excel:
    st.write("### " + nombre_empresa_excel)

ANIO = st.number_input("Año gravable", value=2026)
NOMBRE_EMPRESA = st.text_input("Nombre empresa", value="MASIZO SAS")

CIUDAD_CONSIGNACION = st.text_input("Ciudad de consignación", value="Bogotá").strip()

if not CIUDAD_CONSIGNACION:
    st.warning("⚠️ Debes ingresar la ciudad de consignación")
    st.stop()

fecha_emision = st.date_input("Fecha de emisión", value=date.today())
fecha_texto = fecha_emision.strftime("%d de %b de %Y")

archivo = st.file_uploader("Sube el auxiliar", type=["xlsx"])

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
    ancho_max = ANCHO - (MARGEN_IZQ + MARGEN_DER)
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
    return "CERTIFICADO DE RETENCIONES"

# ---------------- PROCESO ----------------

if archivo:

    df_head = pd.read_excel(archivo, engine="openpyxl", nrows=2)
    nombre_empresa_excel = str(df_head.iloc[1, 3]).strip()
st.write("### " + nombre_empresa_excel)

    df = pd.read_excel(archivo, engine="openpyxl", skiprows=10)

    df.columns = [
        "NitEmpresa","Cuenta","Nombre","Fecha",
        "Fecha2","Nit","Tercero","Descripcion",
        "Debito","Credito","Actividad",
        "Periodo","Proyecto","Tasa",
        "Base","Concepto"
    ]

    # LIMPIEZA
    df["Nit"] = df["Nit"].astype(str).str.strip()
    df["Tercero"] = df["Tercero"].astype(str).str.strip()
    df["Periodo"] = df["Periodo"].astype(str).str.strip()

    df = df[df["Nit"] != "800197268"]
    df = df[df["Cuenta"].astype(str).str.startswith("236")]

    # ---------------- FILTROS ----------------
    incluir_autoret = st.checkbox("Incluir autorretenciones", value=False)

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

    # ✅ FILTRO CORREGIDO
    df_base = df.copy()

    nit_input = st.text_input("Filtrar NIT")
    if nit_input:
        df_base = df_base[df_base["Nit"].str.contains(nit_input, na=False)]

    nombre_input = st.text_input("Buscar tercero")
    if nombre_input:
        df_base = df_base[df_base["Tercero"].str.contains(nombre_input, case=False, na=False)]

    terceros = sorted(df_base["Tercero"].dropna().unique())
    tercero_sel = st.selectbox("Seleccionar tercero", ["Todos"] + terceros)

    if tercero_sel != "Todos":
        df = df_base[df_base["Tercero"] == tercero_sel]
    else:
        df = df_base

    if not incluir_autoret:
        df = df[~df["Nombre"].str.upper().str.contains("AUTORRETENCION", na=False)]

    # ---------------- PERIODO ----------------
    mes_inicio = periodo_a_mes(inicio)
    mes_fin = periodo_a_mes(fin)
    texto_periodo = f"Periodo: {mes_inicio}" if mes_inicio == mes_fin else f"Periodo: {mes_inicio} a {mes_fin}"

    # ---------------- TIPO ----------------
    def tipo_cuenta(c):
        c = str(c)
        if c.startswith("2365"):
            return "Retefuente"
        elif c.startswith("2367"):
            return "ReteIVA"
        elif c.startswith("2368"):
            return "ReteICA"
        return "Otros"

    df["TipoRet"] = df["Cuenta"].apply(tipo_cuenta)
    df = df[df["TipoRet"].isin(tipos_sel)]

    # ---------------- TARIFA ----------------
    def extraer_tarifa(texto):
        m = re.search(r'(\d+[.,]?\d*)\s*%', str(texto))
        if m:
            return float(m.group(1).replace(",", "."))
        m2 = re.search(r'(\d+[.,]\d+)$', str(texto))
        if m2:
            return float(m2.group(1).replace(",", "."))
        return 0

    df["Tarifa"] = df["Nombre"].apply(extraer_tarifa)

    df["TarifaReal"] = df.apply(
        lambda x: x["Tarifa"]/1000 if "ICA" in str(x["Nombre"]).upper()
        else x["Tarifa"]/100 if x["Tarifa"] > 0 else 0,
        axis=1
    )

    # ---------------- CALCULO ----------------
    df["Retencion"] = (df["Credito"] - df["Debito"]).abs()
    df = df[df["Retencion"] > 0]

    df["BaseCalc"] = df.apply(
        lambda x: x["Retencion"]/x["TarifaReal"] if x["TarifaReal"] > 0 else 0,
        axis=1
    )

    agrupado = df.groupby(["Nit","Tercero","TipoRet","Concepto"]).agg({
        "BaseCalc":"sum",
        "Retencion":"sum"
    }).reset_index()

    agrupado.rename(columns={"BaseCalc":"Base"}, inplace=True)

    # ---------------- VALIDACION ----------------
    agrupado["% Calculado"] = agrupado.apply(
        lambda x: round(x["Retencion"]/x["Base"]*100,4) if x["Base"] != 0 else 0, axis=1)

    def porcentaje_esperado(row):
        tarifa = extraer_tarifa(row["Concepto"])
        if "ICA" in row["Concepto"]:
            return round(tarifa/10, 4)
        return tarifa

    agrupado["% Esperado"] = agrupado.apply(porcentaje_esperado, axis=1)
    agrupado["Error"] = abs(agrupado["% Calculado"] - agrupado["% Esperado"])

    agrupado["Estado"] = agrupado["Error"].apply(
        lambda x: "✅ OK" if x < 0.1 else "⚠️ ERROR"
    )

    st.dataframe(agrupado)

    # ---------------- PDF ----------------
    def generar_pdf(nit, nombre, datos, tipo, texto_periodo):

        file_name = f"certificado_{tipo}_{nit}.pdf"
        c = canvas.Canvas(file_name, pagesize=letter)

        c.setFont("Helvetica-Bold", 14)
        c.drawCentredString(300, 710, titulo(tipo))

        c.setFont("Helvetica", 11)
        c.drawCentredString(300, 690, NOMBRE_EMPRESA)
        c.drawCentredString(300, 675, "NIT: 901579559")

        c.line(MARGEN_IZQ, 660, ANCHO-MARGEN_DER, 660)

        y = 630

        c.drawString(MARGEN_IZQ, y, f"{CIUDAD_CONSIGNACION}, {fecha_texto}")
        y -= 20
        c.drawString(MARGEN_IZQ, y, f"Año gravable: {ANIO}")
        y -= 15
        c.drawString(MARGEN_IZQ, y, texto_periodo)

        y -= 25

        c.setFont("Helvetica-Bold", 11)
        c.drawString(MARGEN_IZQ, y, "DATOS DEL BENEFICIARIO")
        y -= 20

        c.setFont("Helvetica", 11)
        c.drawString(MARGEN_IZQ, y, f"Nombre: {nombre}")
        y -= 15
        c.drawString(MARGEN_IZQ, y, f"NIT: {nit}")

        y -= 25

        # TABLA
        c.setFont("Helvetica-Bold", 11)
        c.drawString(MARGEN_IZQ, y, "Concepto")
        c.drawString(220, y, "Base")
        c.drawString(330, y, "Tarifa")
        c.drawString(COL_RET - 60, y, "Retención")

        y -= 15
        c.line(MARGEN_IZQ, y, ANCHO-MARGEN_DER, y)
        y -= 15

        total_base = 0
        total_ret = 0

        c.setFont("Helvetica", 10)

        for _, r in datos.iterrows():

            base = r["Base"]
            ret = r["Retencion"]

            if base == 0:
                tarifa = "N/A"
            elif "ICA" in r["Concepto"]:
                tarifa = f"{round(ret/base*1000,2)}‰"
            else:
                tarifa = f"{round(ret/base*100,2)} %"

            total_base += base
            total_ret += ret

            c.drawString(MARGEN_IZQ, y, str(r["Concepto"])[:35])
            c.drawRightString(COL_BASE, y, f"${base:,.0f}")
            c.drawRightString(COL_TARIFA, y, tarifa)
            c.drawRightString(COL_RET, y, f"${ret:,.0f}")

            y -= 15

        y -= 5
        c.line(MARGEN_IZQ, y, ANCHO-MARGEN_DER, y)
        y -= 20

        c.setFont("Helvetica-Bold", 11)
        c.drawString(MARGEN_IZQ, y, "TOTAL")
        c.drawRightString(COL_BASE, y, f"${total_base:,.0f}")
        c.drawRightString(COL_RET, y, f"${total_ret:,.0f}")

        y -= 30

        c.drawCentredString(
            300,
            y,
            f"Las retenciones practicadas fueron consignadas en la ciudad de {CIUDAD_CONSIGNACION}."
        )

        y -= 30

        y = escribir_parrafo("Este certificado se expide conforme al artículo 381 del Estatuto Tributario.", y, c)
        y -= 10
        y = escribir_parrafo("Este documento no requiere firma autógrafa conforme al Decreto 836 de 1991, Decreto 380 de 1996 y Decreto 1625 de 2016.", y, c)

        c.save()
        return file_name

    # ---------------- GENERACIÓN ----------------
    if st.button("Generar certificados"):

        errores = agrupado[agrupado["Estado"] == "⚠️ ERROR"]

        if not errores.empty:
            st.error("🚫 Hay inconsistencias, no se generan certificados.")
            st.dataframe(errores)
            st.stop()

        archivos = []

        for (nit, tercero, tipo), grupo in agrupado.groupby(["Nit","Tercero","TipoRet"]):
            archivos.append(generar_pdf(nit, tercero, grupo, tipo, texto_periodo))

        if len(archivos) == 1:
            with open(archivos[0], "rb") as f:
                st.download_button("📄 Descargar PDF", f, file_name=archivos[0], mime="application/pdf")
        else:
            zip_name = "certificados.zip"

            with zipfile.ZipFile(zip_name, "w") as z:
                for a in archivos:
                    z.write(a)

            with open(zip_name, "rb") as f:
                st.download_button("📦 Descargar ZIP", f, file_name="certificados.zip", mime="application/zip")
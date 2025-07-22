import streamlit as st
import pandas as pd
from io import BytesIO, StringIO
import datetime
from pathlib import Path
from num2words import num2words
import xlrd
from openpyxl import load_workbook
from pdfrw import PdfReader, PdfWriter, PdfDict, PdfObject

st.set_page_config(page_title="Generador de Nota de Venta y LP", layout="wide")
st.title("Generador Automático de Nota de Venta y Líquido Producto")

def cargar_consumos(raw_bytes, filename):
    ext = Path(filename).suffix.lower()
    buf = BytesIO(raw_bytes)
    if ext == ".csv":
        return pd.read_csv(StringIO(raw_bytes.decode("utf-8", errors="replace")))
    try:
        if ext == ".xls":
            return pd.read_excel(buf, engine="xlrd")
        return pd.read_excel(buf, engine="openpyxl")
    except:
        try:
            return pd.read_html(raw_bytes, header=0)[0]
        except:
            pass
    raise RuntimeError("Error leyendo consumos; convíertelo a .xlsx o .csv.")

# ─── Sidebar Inputs ────────────────────────────────────────────────────────────
reporte_file     = st.sidebar.file_uploader("Reporte consumos (.xls/.xlsx/.csv)",   type=["xls","xlsx","csv"])
proveedores_file = st.sidebar.file_uploader("Listado proveedores (.xls/.xlsx/.csv)", type=["xls","xlsx","csv"])
plantilla_file   = st.sidebar.file_uploader("Plantilla editable LP (PDF)",          type=["pdf"])
numero_lp        = st.sidebar.text_input("Número de LP")

# IVA fijo al 21%
iva_rate = 21.0

# ─── Fechas ───────────────────────────────────────────────────────────────────
hoy     = datetime.date.today()
lp_date = datetime.date(hoy.year, hoy.month, 1)
lp_str  = lp_date.strftime("%d/%m/%Y")
if lp_date.month == 1:
    prev_m, prev_y = 12, lp_date.year - 1
else:
    prev_m, prev_y = lp_date.month - 1, lp_date.year
periodo_liq = f"{prev_m:02d}/{prev_y}"

# ─── Workflow ─────────────────────────────────────────────────────────────────
if reporte_file and proveedores_file and plantilla_file:
    # 1) Leer consumos y extraer Subtotal (celda I11) y marca (celda B7)
    raw_cons = reporte_file.read()
    try:
        df_cons   = cargar_consumos(raw_cons, reporte_file.name)
        ext_cons  = Path(reporte_file.name).suffix.lower()
        if ext_cons == ".xls":
            wb_xls   = xlrd.open_workbook(file_contents=raw_cons)
            sheet    = wb_xls.sheet_by_index(0)
            subtotal = float(sheet.cell_value(10, 8))
            marca    = str(sheet.cell_value(6, 1)).strip()
        else:
            wb_xlsx  = load_workbook(filename=BytesIO(raw_cons), data_only=True)
            ws       = wb_xlsx.active
            subtotal = float(ws["I11"].value)
            marca    = str(ws["B7"].value).strip()
    except Exception as e:
        st.error(f"Error leyendo consumos: {e}")
        st.stop()

    # 2) Leer proveedores y preparar buscador
    raw_prov = proveedores_file.read()
    try:
        df_prov = cargar_consumos(raw_prov, proveedores_file.name)
    except Exception as e:
        st.error(f"Error leyendo proveedores: {e}")
        st.stop()

    # Limpiamos razón social y marcas
    df_prov["ProvClean"]   = df_prov["Proveedores"].str.replace(r"\s*\(.*\)", "", regex=True).str.strip()
    df_prov["MarcaClean"]  = df_prov["marcas"].astype(str).str.strip()

    # 3) Buscar proveedor por marca en B7
    matches = df_prov[df_prov["MarcaClean"].str.lower() == marca.lower()]
    if matches.empty:
        st.error(f"No se encontró proveedor para la marca: '{marca}'")
        st.stop()
    prov = matches.iloc[0]

    # Mostrar marca y proveedor detectado
    st.write(f"🔍 Marca detectada en reporte: **{marca}**")
    st.write("Proveedor seleccionado automáticamente:")
    st.write(f"- Razón social: {prov['ProvClean']}")
    st.write(f"- Dirección:    {prov['Dirección']}")
    st.write(f"- CUIT:         {prov['CUIT']}")

    # 4) Generar LP
    if st.button("Generar LP"):
        try:
            tpl = PdfReader(BytesIO(plantilla_file.read()))

            # (Opcional) Mostrar campos detectados
            if tpl.Root.AcroForm and tpl.Root.AcroForm.Fields:
                campos = [f.T[1:-1] for f in tpl.Root.AcroForm.Fields if f.T]
                st.write("📑 Campos internos detectados:", campos)

            # Calcular IVA y Total en letras
            iva_amt = round(subtotal * iva_rate / 100, 2)
            total   = round(subtotal + iva_amt, 2)
            entero  = int(total)
            decs    = int(round((total - entero) * 100))
            literal = num2words(entero, lang="es").capitalize() + f" con {decs:02d}"

            # Mapeo con los nombres reales de tu PDF
            valores = {
                "cliente":     prov["ProvClean"],    # Señor/es
                "dirección":   prov["Dirección"],   # Dirección
                "iva":         "Inscripto",         # I.V.A.
                "cuit":        prov["CUIT"],        # C.U.I.T.
                "fecha":       lp_str,              # Fecha
                "nfactura":    numero_lp,           # Nº LP
                "detalle":     "ventas por cuenta y orden",
                "subtotal":    f"{subtotal}",
                "iva insc":    f"{iva_amt}",
                "iva total":   f"{total}",
                "liquidacion": periodo_liq,
                "enpesos":     literal,             # Son Pesos
            }

            # Rellenar campos sin paréntesis
            for page in tpl.pages:
                if not page.Annots:
                    continue
                for annot in page.Annots:
                    if annot.T:
                        key = annot.T[1:-1]
                        if key in valores:
                            annot.V  = valores[key]
                            annot.AP = None

            # Forzar regenerar apariencias
            if tpl.Root.AcroForm:
                tpl.Root.AcroForm.update(
                    PdfDict(NeedAppearances=PdfObject("true"))
                )

            # Guardar y descargar
            out = BytesIO()
            PdfWriter().write(out, tpl)
            out.seek(0)
            st.download_button(
                "Descargar Nota de Venta y LP (PDF)",
                data=out,
                file_name="Nota_Venta_LP.pdf",
                mime="application/pdf"
            )
            st.success("¡LP generado correctamente!")
        except Exception as e:
            st.error(f"Error generando LP: {e}")
else:
    st.info("Carga reporte consumos, proveedores y plantilla PDF para habilitar la generación.")

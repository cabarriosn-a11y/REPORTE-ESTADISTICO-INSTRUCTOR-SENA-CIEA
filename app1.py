import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import os
from datetime import datetime, date, time, timedelta
import calendar
import re

# Librerías para PDF
from reportlab.lib.pagesizes import landscape, letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch

# ==========================================
# MOTOR DE BASE DE DATOS - GOOGLE SHEETS
# ==========================================
@st.cache_data(ttl=600)
def cargar_competencias_gsheets():
    sheet_id = "1MAIAGFEBerD3Gg-WYfOMvTXQ0uj7JO8UZBag0e8sxjw"
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
    try:
        df = pd.read_csv(url)
        df.fillna("", inplace=True) 
        base_datos = {}
        for _, fila in df.iterrows():
            comp = str(fila.iloc[0]).strip()
            rap = str(fila.iloc[1]).strip() if len(fila) > 1 else ""
            if comp != "" and comp.lower() != "nan" and "unnamed" not in comp.lower():
                if comp not in base_datos:
                    base_datos[comp] = []
                if rap != "" and rap.lower() != "nan" and "unnamed" not in rap.lower():
                    base_datos[comp].append(rap)
        base_datos["OTRA (Escribir manualmente)"] = []
        return base_datos
    except:
        return {"OTRA (Escribir manualmente)": []}

DB_SENA = cargar_competencias_gsheets()

# ==========================================
# MOTOR DE EXPORTACIÓN A PDF
# ==========================================
def crear_pdf(nombre, cedula, mes, anio, datos_formacion, datos_otras, tot_dir, tot_otr, tot_gen):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    elements = []
    
    styles = getSampleStyleSheet()
    style_center = ParagraphStyle(name='Center', parent=styles['Normal'], alignment=1, fontSize=11, leading=14)
    style_title = ParagraphStyle(name='Title', parent=styles['Heading1'], alignment=1, fontSize=14, spaceAfter=12)
    style_cell = ParagraphStyle(name='Cell', parent=styles['Normal'], fontSize=8, leading=10)
    style_subtitle = ParagraphStyle(name='Sub', parent=styles['Normal'], fontSize=10, fontName='Helvetica-Bold', spaceAfter=6, spaceBefore=10)

    if os.path.exists("logo_sena.png"):
        img = Image("logo_sena.png", width=0.7*inch, height=0.7*inch)
        img.hAlign = 'CENTER'
        elements.append(img)
    
    elements.append(Paragraph("<b>CENTRO INDUSTRIAL Y DE ENERGIAS ALTERNATIVAS - REGIONAL GUAJIRA</b>", style_center))
    elements.append(Paragraph("<b>REPORTE ESTADISTICO DE HORAS MENSUALES INSTRUCTOR</b>", style_title))
    
    info_text = f"<b>INSTRUCTOR:</b> {nombre} &nbsp;&nbsp;&nbsp;&nbsp; <b>CÉDULA:</b> {cedula} &nbsp;&nbsp;&nbsp;&nbsp; <b>MES:</b> {mes} / {anio}"
    elements.append(Paragraph(info_text, styles['Normal']))
    
    # Tabla Directa
    elements.append(Paragraph("PARTE HORAS DIRECTAS - PROGRAMAS EN FORMACIÓN TITULADA", style_subtitle))
    data_table = [["FICHA", "DESDE", "HASTA", "L", "M", "MI", "J", "V", "S", "COMPETENCIA", "RAP", "EVAL", "TERM", "HRS"]]
    for f in datos_formacion:
        row = [f['ficha'], f['h_inicio'].strftime("%H:%M"), f['h_fin'].strftime("%H:%M"), "X" if f['dias']["L"] else "", "X" if f['dias']["M"] else "", "X" if f['dias']["Mi"] else "", "X" if f['dias']["J"] else "", "X" if f['dias']["V"] else "", "X" if f['dias']["S"] else "", Paragraph(f['competencia'], style_cell), Paragraph(f['rap'], style_cell), f.get('evaluado', 'NO'), f.get('termino', 'NO'), f"{f['horas']:g}"]
        data_table.append(row)
    data_table.append(["", "", "", "", "", "", "", "", "", "", "", "", "TOTAL:", f"{tot_dir:g}"])

    t_dir = Table(data_table, colWidths=[45, 35, 35, 15, 15, 15, 15, 15, 15, 195, 205, 30, 30, 25])
    t_dir.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.lightgrey), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('GRID', (0,0), (-1,-1), 0.5, colors.black), ('FONTSIZE', (0,0), (-1,-1), 8)]))
    elements.append(t_dir)
    
    # Tabla Otras
    if datos_otras:
        elements.append(Spacer(1, 15))
        elements.append(Paragraph("PARTE OTRAS ACTIVIDADES INSTRUCTORES PLANTA", style_subtitle))
        data_otras = [["ACTIVIDAD", "FECHA DESDE", "FECHA HASTA", "CANT. DÍAS", "HRS"]]
        for of in datos_otras:
            data_otras.append([of['actividad'], of['f_desde'].strftime("%d/%m/%Y"), of['f_hasta'].strftime("%d/%m/%Y"), f"{of['dias']:g}", f"{of['horas']:g}"])
        data_otras.append(["", "", "", "TOTAL OTRAS:", f"{tot_otr:g}"])
        t_otras = Table(data_otras, colWidths=[250, 110, 110, 100, 80])
        t_otras.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.lightgrey), ('GRID', (0,0), (-1,-1), 0.5, colors.black), ('ALIGN', (0,0), (-1,-1), 'CENTER')]))
        elements.append(t_otras)
    
    elements.append(Spacer(1, 20))
    elements.append(Paragraph(f"<b>TOTAL HORAS REPORTADAS EN EL MES: &nbsp; {tot_gen:g} Horas</b>", style_center))
    
    fecha_gen = datetime.now().strftime("%d/%m/%Y %I:%M %p")
    firma_html = f"<br/><br/><br/>_________________________________<br/><b>FIRMA DEL INSTRUCTOR</b><br/>{nombre}<br/>C.C. {cedula}<br/><br/><font size=8 color=gray>Reporte generado el: {fecha_gen}</font>"
    elements.append(Paragraph(firma_html, style_center))
    
    doc.build(elements)
    return buffer.getvalue()

# ==========================================
# INTERFAZ STREAMLIT
# ==========================================
st.set_page_config(page_title="Reporte SENA CIEA", layout="wide")

# CSS: BOTÓN VERDE 3D Y BOTONES ROJOS DE BORRADO
st.markdown("""
    <style>
    /* Botón de Descarga Verde 3D */
    div.stDownloadButton > button {
        background: linear-gradient(145deg, #2ecc71, #27ae60) !important;
        color: white !important;
        height: 4.5em !important;
        width: 100% !important;
        border-radius: 15px !important;
        font-weight: bold !important;
        font-size: 24px !important;
        border: none !important;
        box-shadow: 0px 8px 0px #1e8449, 0px 10px 20px rgba(0,0,0,0.3) !important;
        transition: all 0.1s ease !important;
    }
    div.stDownloadButton > button:hover { transform: translateY(2px) !important; box-shadow: 0px 5px 0px #1e8449 !important; }
    
    /* Botones de Borrar Rojos */
    button[key^="df"], button[key^="do"] {
        background-color: #ff0000 !important;
        color: white !important;
        border-radius: 10px !important;
        font-weight: bold !important;
        border: 2px solid #8b0000 !important;
    }
    </style>
    """, unsafe_allow_html=True)

# Encabezado con Logo (Corregido)
col_logo, col_tit = st.columns([1, 6])
with col_logo:
    if os.path.exists("logo_sena.png"):
        st.image("logo_sena.png", width=120)
with col_tit:
    st.markdown("## CENTRO INDUSTRIAL Y DE ENERGIAS ALTERNATIVAS")
    st.markdown("#### REGIONAL GUAJIRA - Reporte Estadístico Mensual")

with st.container():
    c1, c2, c3, c4 = st.columns(4)
    nombre_ins = c1.text_input("Nombre del Instructor")
    cedula_ins = c2.text_input("Cédula")
    meses_str = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    mes_rep = c3.selectbox("Mes del Reporte", meses_str, index=datetime.now().month - 1)
    anio_rep = c4.text_input("Año", value="2026")

if 'filas' not in st.session_state: st.session_state.filas = []
if 'otras_filas' not in st.session_state: st.session_state.otras_filas = []

# --- CARGA AUTOMÁTICA BLINDADA (Corregido para tu Excel) ---
st.divider()
with st.expander("🚀 Cargar Horario Automáticamente desde Excel Oficial", expanded=True):
    archivo = st.file_uploader("Sube tu horario oficial", type=['xlsx', 'xls'])
    if archivo and st.button("⚙️ Procesar Horario"):
        try:
            xls = pd.ExcelFile(archivo)
            # Buscar hoja flexible (ignora si termina en -FEB-9 o similar)
            hoja = next((n for n in xls.sheet_names if "HORARIOINSTRUCTOR" in n.upper() or "INSTRUCTORHORARIO" in n.upper()), None)
            if hoja:
                df_h = pd.read_excel(xls, sheet_name=hoja, header=None)
                hora_col, start_row, grupo_cols = -1, -1, []
                # Radar de cabeceras (escanea más profundo por las celdas combinadas)
                for r in range(min(40, len(df_h))):
                    row_vals = [str(x).strip().upper() for x in df_h.iloc[r].values]
                    if 'HORA' in row_vals:
                        hora_col = row_vals.index('HORA')
                        start_row = r + 1
                    for c, val in enumerate(row_vals):
                        if 'GRUPO' in val:
                            if c not in grupo_cols: grupo_cols.append(c)
                
                if hora_col != -1 and grupo_cols:
                    dias_nom = ["L", "M", "Mi", "J", "V", "S"]
                    day_cols = {dias_nom[i]: col for i, col in enumerate(sorted(list(set(grupo_cols)))) if i < 6}
                    bloques = []
                    for dia, col_idx in day_cols.items():
                        cur_f, cur_s, cur_e = None, None, None
                        for idx in range(start_row, len(df_h)):
                            h_raw = str(df_h.iloc[idx, hora_col])
                            m = re.findall(r'\d{1,2}:\d{2}', h_raw)
                            if len(m) < 2: continue
                            f_val = str(df_h.iloc[idx, col_idx]).strip().split('.')[0]
                            f_limpia = f_val if f_val.isdigit() and len(f_val) > 3 else None
                            if f_limpia:
                                if cur_f == f_limpia and cur_e == m[0]: cur_e = m[1]
                                else:
                                    if cur_f: bloques.append({'ficha': cur_f, 'dia': dia, 'inicio': cur_s, 'fin': cur_e})
                                    cur_f, cur_s, cur_e = f_limpia, m[0], m[1]
                            else:
                                if cur_f: bloques.append({'ficha': cur_f, 'dia': dia, 'inicio': cur_s, 'fin': cur_e})
                                cur_f = None
                        if cur_f: bloques.append({'ficha': cur_f, 'dia': dia, 'inicio': cur_s, 'fin': cur_e})

                    agrupados = {}
                    for b in bloques:
                        k = (b['ficha'], b['inicio'], b['fin'])
                        if k not in agrupados: agrupados[k] = []
                        agrupados[k].append(b['dia'])
                    
                    st.session_state.filas = []
                    for (f, i, fn), d_list in agrupados.items():
                        st.session_state.filas.append({
                            "ficha": f, "h_inicio": datetime.strptime(i, "%H:%M").time(), "h_fin": datetime.strptime(fn, "%H:%M").time(), 
                            "dias": {d: (d in d_list) for d in dias_nom},
                            "competencia": list(DB_SENA.keys())[0] if DB_SENA else "OTRA", "rap": "", "horas": 0, "evaluado": "NO", "termino": "NO"
                        })
                    st.success("✅ ¡Horario reconocido!")
                    st.rerun()
            else: st.error("❌ No se encontró la hoja de horario en el archivo.")
        except Exception as e: st.error(f"Error: {e}")

# --- FORMACIÓN DIRECTA ---
total_dir = 0
st.subheader("📘 1. Formación Directa")
if st.button("➕ Agregar Ficha Manual"):
    st.session_state.filas.append({"ficha":"","h_inicio":time(8,0),"h_fin":time(12,0),"dias":{d:False for d in ["L","M","Mi","J","V","S"]},"competencia":list(DB_SENA.keys())[0],"rap":"","horas":0,"evaluado":"NO","termino":"NO"})

f_idx_del = []
for i, fila in enumerate(st.session_state.filas):
    with st.expander(f"📌 Ficha: {fila['ficha']} | {fila['h_inicio'].strftime('%H:%M')}", expanded=True):
        c_dat, c_del = st.columns([0.85, 0.15])
        if c_del.button("ELIMINAR", key=f"df{i}"): f_idx_del.append(i)
        
        c1, c2, c3, c4 = c_dat.columns([1.5, 1, 1, 1])
        fila['ficha'] = c1.text_input("Ficha", fila['ficha'], key=f"f{i}")
        fila['h_inicio'], fila['h_fin'] = c2.time_input("Inicio", fila['h_inicio'], key=f"hi{i}"), c3.time_input("Fin", fila['h_fin'], key=f"hf{i}")
        cd = st.columns(6)
        for idx, d in enumerate(["L", "M", "Mi", "J", "V", "S"]): fila['dias'][d] = cd[idx].checkbox(d, fila['dias'][d], key=f"d{d}{i}")
        
        h_dia = (datetime.combine(date.today(), fila['h_fin']) - datetime.combine(date.today(), fila['h_inicio'])).seconds / 3600
        m_idx, a_int = meses_str.index(mes_rep) + 1, int(anio_rep)
        fechas = [date(a_int, m_idx, d) for d in range(1, calendar.monthrange(a_int, m_idx)[1]+1) if date(a_int, m_idx, d).weekday() in [idx for idx, d in enumerate(["L", "M", "Mi", "J", "V", "S"]) if fila['dias'][d]]]
        excl = st.multiselect("Descontar festivos:", [f.strftime("%d/%m/%Y") for f in fechas], key=f"ex{i}")
        fila['horas'] = (len(fechas) - len(excl)) * h_dia
        total_dir += fila['horas']
        c4.metric("Subtotal", f"{fila['horas']:g} hrs")
        fila['competencia'] = st.selectbox("Competencia", list(DB_SENA.keys()), key=f"cp{i}")
        ops = DB_SENA.get(fila['competencia'], [])
        fila['rap'] = st.selectbox("RAP", ops, key=f"rp{i}") if ops else st.text_area("RAP manual", key=f"rpm{i}")

for idx in reversed(f_idx_del): st.session_state.filas.pop(idx); st.rerun()

# --- OTRAS ACTIVIDADES ---
st.divider()
st.subheader("📙 2. Otras Actividades (Instructores de Planta)")
if st.button("➕ Agregar Actividad Planta"):
    st.session_state.otras_filas.append({"actividad": "Preparación de clases", "f_desde": date.today(), "f_hasta": date.today(), "dias": 1.0, "horas": 8.5})

lista_planta = ["Preparación de clases", "Semana de confraternidad", "Actividades deportivas", "Encuentros culturales", "Día del instructor", "Permiso Sindical", "Incapacidad médica", "Permiso particular", "Permiso por estudio", "Día de la familia", "Festivo", "Otro"]
total_otr, o_idx_del = 0, []
for j, ofila in enumerate(st.session_state.otras_filas):
    with st.container():
        c1, c2, c3, c4, c5, c6 = st.columns([2, 1, 1, 0.8, 0.8, 0.4])
        ofila['actividad'] = c1.selectbox("Actividad", lista_planta, key=f"oa{j}")
        ofila['f_desde'], ofila['f_hasta'], ofila['dias'] = c2.date_input("Desde", ofila['f_desde'], key=f"ofd{j}"), c3.date_input("Hasta", ofila['f_hasta'], key=f"ofh{j}"), c4.number_input("Días", 0.0, 31.0, float(ofila['dias']), step=0.5, key=f"od{j}")
        ofila['horas'] = ofila['dias'] * 8.5
        c5.metric("Hrs", f"{ofila['horas']:g}")
        if c6.button("X", key=f"do{j}"): o_idx_del.append(j)
        total_otr += ofila['horas']

for idx in reversed(o_idx_del): st.session_state.otras_filas.pop(idx); st.rerun()

# --- BARRA LATERAL (Sidebar) ---
total_mes = total_dir + total_otr
if os.path.exists("logo_sena.png"):
    st.sidebar.image("logo_sena.png", width=100)
st.sidebar.markdown(f"### 📊 RESUMEN\n**Formación:** {total_dir:g} hrs\n**Otras:** {total_otr:g} hrs\n---\n**TOTAL MES:** {total_mes:g} hrs")

if nombre_ins and total_mes > 0:
    pdf_f = crear_pdf(nombre_ins, cedula_ins, mes_rep, anio_rep, st.session_state.filas, st.session_state.otras_filas, total_dir, total_otr, total_mes)
    st.download_button(label="📥 DESCARGAR REPORTE PDF FINAL (3D)", data=pdf_f, file_name=f"Reporte_{nombre_ins}_{mes_rep}.pdf", mime="application/pdf")

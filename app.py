import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as graph_objects
from datetime import datetime, timedelta
from PIL import Image
import os
import io

# Nombre oficial del archivo base y carpeta física del proyecto
ARCHIVO_DB = "base_matriz_mce.xlsx"
CARPETA_EVIDENCIAS = "evidencias"

if not os.path.exists(CARPETA_EVIDENCIAS):
    os.makedirs(CARPETA_EVIDENCIAS)

Image.MAX_IMAGE_PIXELS = None

# 1. MOTOR DE IMPORTACIÓN INTELIGENTE DE EXCEL
def importar_registros_excel():
    if os.path.exists(ARCHIVO_DB):
        try:
            df = pd.read_excel(ARCHIVO_DB)
            if not df.empty:
                for col_fecha in ["Fecha Inicio", "Fecha Compromiso"]:
                    if col_fecha in df.columns:
                        def corregir_fecha_serial(val):
                            try:
                                if pd.isna(val) or str(val).strip() in ["None", "nan", "NaN", ""]: return ""
                                if str(val).replace('.0', '').isdigit():
                                    dias = int(str(val).replace('.0', ''))
                                    return (datetime(1899, 12, 30) + timedelta(days=dias)).strftime("%d-%b-%y")
                                return str(val).strip()
                            except: return str(val)
                        df[col_fecha] = df[col_fecha].apply(corregir_fecha_serial)

                if "% Avance" in df.columns:
                    if df["% Avance"].max() <= 1.0 and df["% Avance"].max() > 0: df["% Avance"] = df["% Avance"] * 100
                    df["% Avance"] = pd.to_numeric(df["% Avance"], errors="coerce").fillna(0).astype(int)
                
                if "No" in df.columns:
                    df["No"] = pd.to_numeric(df["No"], errors="coerce")
                    if df["No"].isnull().any(): df["No"] = range(1, len(df) + 1)
                    df["No"] = df["No"].astype(int)
                
                columnas_texto = ["Origen", "Prioridad", "Responsable", "Area", "Descripcion", "Comentario", "Evidencia"]
                for col in columnas_texto:
                    if col in df.columns: df[col] = df[col].astype(str).replace(["None", "nan", "NaN"], "")
            else:
                df = pd.DataFrame(columns=["No", "Origen", "Fecha Inicio", "Prioridad", "Responsable", "Area", "Descripcion", "% Avance", "Fecha Compromiso", "Comentario", "Evidencia"])
            return df
        except Exception as e:
            st.error(f"Error al importar el archivo Excel: {e}")
            return pd.DataFrame(columns=["No", "Origen", "Fecha Inicio", "Prioridad", "Responsable", "Area", "Descripcion", "% Avance", "Fecha Compromiso", "Comentario", "Evidencia"])
    else:
        return pd.DataFrame(columns=["No", "Origen", "Fecha Inicio", "Prioridad", "Responsable", "Area", "Descripcion", "% Avance", "Fecha Compromiso", "Comentario", "Evidencia"])

if 'actividades' not in st.session_state:
    st.session_state.actividades = importar_registros_excel()

st.set_page_config(page_title="SIGRAMA - Matriz MCE", layout="wide")
st.markdown("""
    <style>
    .stApp { background-color: #ECEFF1 !important; }
    [data-testid="stSidebar"] { background-color: #CFD8DC !important; }
    .main-title { font-size:28px !important; font-weight: bold; color: #0C2340; text-align: left; margin-top: 0px; }
    .card-header { font-size: 20px !important; font-weight: bold; color: #0C2340; margin-bottom: 3px; }
    .card-desc { font-size: 16px !important; font-weight: 500; color: #333333; margin-bottom: 5px; }
    .card { padding: 12px 18px; border-radius: 8px; box-shadow: 0px 2px 4px rgba(0,0,0,0.05); background-color: #FFFFFF; margin-bottom: 10px; border: 1px solid #E0E0E0; }
    </style>
""", unsafe_allow_html=True)
if 'personal' not in st.session_state:
    st.session_state.personal = {
        "Jesus Morales": None, "Ing. Alfredo Hdz": None, "Ing. Lorena Hdz": None, "Jesús Alday": None,
        "Alejandra Arellano": None, "Jose Romo": None, "Jose Manuel Aldama": None, "Fernando Llanas": None,
        "Celso": None, "Cruz Carreon": None, "Luis Quintana": None, "Bryan Flores": None,
        "Jorge Sanchez": None, "Voctor Montoya": None, "Moises Hernandez": None, "Rodolfo Ferndez M": None
    }

if 'areas' not in st.session_state:
    st.session_state.areas = ["⚙️ Ingenieria", "🔍 Calidad", "📦 Almacen", "✂️ Corte", "📐 Doblez", "🎨 Pintura", "🚚 Embarquez", "🏭 Planta Rio"]

LISTA_CLASIFICACIONES = ["Acuerdos", "Programa de Actividades", "Actividades Sujeridas", "Dirección", "Problema de Calidad", "Problema de Seguridad", "Lista de Pendientes", "Auto Asignado", "Plan de Control y Monitoreo", "Mejoras", "Investigación", "Manuales", "Procesos"]

# CORRECCIÓN DE LA LÍNEA 111: Se añade el rango explícito [0, 100] para reparar el SyntaxError de tu pantalla
def crear_grafico_pareto(df, columna, titulo):
    if df.empty:
        fig = graph_objects.Figure(); fig.update_layout(title=f"{titulo} (Sin Datos)"); return fig
    df = df.copy()
    df["Estado"] = df["% Avance"].apply(lambda x: "Terminada" if x == 100 else "Pendiente")
    orden_cat = df[columna].value_counts().index.tolist()
    counts = df[columna].value_counts().reset_index()
    counts.columns = [columna, 'Cantidad']
    counts['Porcentaje Acumulado'] = (counts['Cantidad'].cumsum() / len(df)) * 100
    
    fig = px.histogram(df, x=columna, color="Estado", category_orders={columna: orden_cat}, color_discrete_map={"Terminada": "#2ECC71", "Pendiente": "#FEEA9A"}, title=titulo)
    fig.add_trace(graph_objects.Scatter(x=counts[columna], y=counts['Porcentaje Acumulado'], name="% Acumulado", yaxis="y2", line=dict(color="#FF5733", width=3)))
    fig.update_layout(yaxis=dict(title="Cantidad"), yaxis2=dict(title="% Acumulado", overlaying="y", side="right", range=[0, 100]), legend=dict(orientation="h", y=1.02, x=1, xanchor="right"), template="plotly_white", barmode="stack")
    return fig

nombre_logo = "LOGOTIPO COLOR (1).jfif"
imagen_logo = Image.open(nombre_logo) if os.path.exists(nombre_logo) else None

if imagen_logo is not None:
    col_logo_g, col_tit_g = st.columns()
    with col_logo_g: st.image(imagen_logo, width=220) 
    with col_tit_g: 
        st.markdown('<h2 style="color: #0C2340; margin-bottom: 0px; padding-top: 10px; font-weight: bold;">PLANTA METALES Y MAQUINADOS</h2>', unsafe_allow_html=True)
        st.markdown('<p class="main-title">MATRIZ DE COMUNICACIÓN EFECTIVA</p>', unsafe_allow_html=True)
else:
    st.markdown('<h2 style="color: #0C2340; text-align: center; font-weight: bold;">PLANTA METALES Y MAQUINADOS</h2>', unsafe_allow_html=True)
    st.markdown('<p class="main-title" style="text-align: center;">MATRIZ DE COMUNICACIÓN EFECTIVA</p>', unsafe_allow_html=True)

opcion_menu = st.sidebar.radio("Navegación", ["📊 Dashboard Principal", "📋 Tabla de Control", "📝 Actualizar Mis Avances", "📥 Cargar Actividades (Usuario)", "🔐 Panel Administrador"])
if opcion_menu == "📊 Dashboard Principal":
    col_f1, col_f2 = st.columns(2)
    with col_f1: area_sel = st.selectbox("Filtrar por Área", ["Todas"] + st.session_state.areas)
    with col_f2: resp_sel = st.selectbox("Filtrar por Responsable", ["Todos"] + list(st.session_state.personal.keys()))
    df_f = pd.DataFrame(st.session_state.actividades)
    if area_sel != "Todas": df_f = df_f[df_f["Area"] == area_sel]
    if resp_sel != "Todos": df_f = df_f[df_f["Responsable"] == resp_sel]
    
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Actividades", len(df_f))
    c2.metric("Terminadas", len(df_f[df_f["% Avance"] == 100]))
    c3.metric("Avance Promedio", f"{df_f['% Avance'].mean() if len(df_f)>0 else 0:.1f}%")
    st.write("---")
    g1, g2 = st.columns(2)
    with g1: st.plotly_chart(crear_grafico_pareto(df_f, "Origen", "Pareto 1: Actividades vs Cantidad"), use_container_width=True)
    with g2: st.plotly_chart(crear_grafico_pareto(df_f, "Responsable", "Pareto 2: Personas vs Cantidad"), use_container_width=True)
    
    st.write("---"); st.subheader("Pareto 3: Estado de Actividades de Líderes Principales")
    lideres_p = ["Jesus Morales", "Bryan Flores", "Cruz Carreon", "Luis Quintana"]
    df_lideres = pd.DataFrame(st.session_state.actividades)[lambda x: x["Responsable"].isin(lideres_p)].copy()
    if not df_lideres.empty:
        hoy = datetime.now()
        def clasificar_vencimientos(f):
            try:
                f_comp = datetime.strptime(str(f["Fecha Compromiso"]).strip(), "%d-%b-%y")
                return "Vencida (Retraso)" if (int(f["% Avance"]) < 100 and f_comp < hoy) else ("Terminada" if int(f["% Avance"]) == 100 else "Pendiente a Tiempo")
            except: return "Pendiente a Tiempo"
        df_lideres["Estado_Real"] = df_lideres.apply(clasificar_vencimientos, axis=1)
        df_lideres["Valor_Eje"] = df_lideres["Estado_Real"].apply(lambda x: -1 if x == "Vencida (Retraso)" else 1)
        fig_l = px.bar(df_lideres, x="Origen", y="Valor_Eje", color="Estado_Real", facet_col="Responsable", facet_col_wrap=2, color_discrete_map={"Terminada": "#2ECC71", "Pendiente a Tiempo": "#FEEA9A", "Vencida (Retraso)": "#E74C3C"}, labels={"Origen": "Clasificación", "Valor_Eje": "Tareas"})
        fig_l.update_layout(barmode="stack", template="plotly_white", height=600, legend=dict(orientation="h", y=-0.2, x=0.5, xanchor="center"))
        fig_l.for_each_annotation(lambda a: a.update(text=f"<b>👤 {a.text.split('=')[-1].upper()}</b>", font=dict(size=14, color="#0C2340")))
        st.plotly_chart(fig_l, use_container_width=True)

elif opcion_menu == "📋 Tabla de Control":
    st.subheader("Historial Completo de la Matriz de Comunicación")
    df_t = pd.DataFrame(st.session_state.actividades)
    st.write("---")
    ct1, ct2, ct3 = st.columns(3)
    with ct1: fa = st.selectbox("Filtrar por Área", ["Todas"] + st.session_state.areas, key="t_a")
    with ct2: fr = st.selectbox("Filtrar por Responsable", ["Todos"] + list(st.session_state.personal.keys()), key="t_r")
    with ct3: fp = st.selectbox("Filtrar por Prioridad", ["Todas", "Urgente", "Media", "Baja"])
    if fa != "Todas": df_t = df_t[df_t["Area"] == fa]
    if fr != "Todos": df_t = df_t[df_t["Responsable"] == fr]
    if fp != "Todas": df_t = df_t[df_t["Prioridad"] == fp]
    busq = st.text_input("🔍 Buscar por palabra clave en la descripción:")
    if busq: df_t = df_t[df_t["Descripcion"].str.contains(busq, case=False, na=False)]
    if not df_t.empty:
        df_mostrar = df_t[["No", "Origen", "Fecha Inicio", "Prioridad", "Responsable", "Area", "Descripcion", "% Avance", "Fecha Compromiso", "Comentario", "Evidencia"]].copy()
        hoy = datetime.now()
        def aplicar_colores_renglon(fila):
            try:
                num_avance = int(fila["% Avance"])
                fecha_comp_str = str(fila["Fecha Compromiso"]).strip()
                fecha_vencida = False
                if fecha_comp_str:
                    f_comp = datetime.strptime(fecha_comp_str, "%d-%b-%y")
                    if f_comp < hoy: fecha_vencida = True
                if num_avance < 100 and fecha_vencida: return ['background-color: #F8D7DA; color: #721C24; font-weight: 500;'] * len(fila)
                elif num_avance == 100: return ['background-color: #D4EDDA; color: #155724;'] * len(fila)
                elif num_avance > 0: return ['background-color: #FFF3CD; color: #856404;'] * len(fila)
            except: pass
            return [''] * len(fila)
        st.dataframe(df_mostrar.style.apply(aplicar_colores_renglon, axis=1).format({"% Avance": "{:.0f}%"}), use_container_width=True, hide_index=True)
    else: st.info("No se encontraron registros.")
elif opcion_menu == "📝 Actualizar Mis Avances":
    st.subheader("Actualización de Avances de Tareas")
    u = st.selectbox("Identifícate", list(st.session_state.personal.keys()))
    df_usuario = st.session_state.actividades[st.session_state.actividades["Responsable"] == u]
    if df_usuario.empty: st.info("No tienes actividades pendientes.")
    else:
        for idx in df_usuario.index:
            r = st.session_state.actividades.loc[idx]
            with st.container():
                st.markdown('<div class="card">', unsafe_allow_html=True)
                st.markdown(f'<p class="card-header">Actividad No. {r["No"]} | {r["Area"]}</p>', unsafe_allow_html=True)
                st.markdown(f'<p class="card-desc"><b>Descripción:</b> {r["Descripcion"]}</p>', unsafe_allow_html=True)
                col_izq, col_der = st.columns(2)
                with col_izq:
                    progreso_actual = int(r['% Avance'])
                    fig_slider = graph_objects.Figure()
                    fig_slider.add_trace(graph_objects.Bar(x=["Progreso"], y=, marker_color="#E0E0E0", showlegend=False, hoverinfo="none"))
                    fig_slider.add_trace(graph_objects.Bar(x=["Progreso"], y=[progreso_actual], marker_color="#2ECC71" if progreso_actual == 100 else "#0C2340", showlegend=False, text=f"{progreso_actual}%", textposition="inside"))
                    fig_slider.update_layout(barmode="overlay", template="plotly_white", height=140, width=90, margin=dict(l=5, r=5, t=5, b=5), xaxis=dict(visible=False), yaxis=dict(range=, visible=False))
                    st.plotly_chart(fig_slider, use_container_width=False, config={'displayModeBar': False}, key=f"plot_chart_{r['No']}")
                    nv_av = st.slider("Ajustar %:", min_value=0, max_value=100, value=progreso_actual, step=5, key=f"num_{r['No']}")
                with col_der:
                    nv_co = st.text_input("Comentarios de bitácora:", str(r['Comentario']), key=f"c_{r['No']}")
                    evidencia_guardada = str(r['Evidencia']).strip()
                    if evidencia_guardada and os.path.exists(evidencia_guardada): st.image(Image.open(evidencia_guardada), width=150)
                    foto = st.file_uploader("Subir Evidencia (Cierre 100%):", type=["jpg","png","jpeg"], key=f"i_{r['No']}") if nv_av == 100 else None
                    if st.button("Guardar Tarea", key=f"b_{r['No']}"):
                        if nv_av == 100 and not foto and not evidencia_guardada: st.error("Falta la foto.")
                        else:
                            ruta_foto_final = evidencia_guardada
                            if foto is not None:
                                try:
                                    img = Image.open(foto)
                                    if img.mode in ("RGBA", "P"): img = img.convert("RGB")
                                    img.thumbnail((800, 600), Image.Resampling.LANCZOS)
                                    ruta_foto_final = os.path.join(CARPETA_EVIDENCIAS, f"MCE-{int(r['No']):03d}_evidencia.jpg")
                                    img.save(ruta_foto_final, "JPEG", quality=65)
                                except Exception as e_img: st.error(f"Error imagen: {e_img}")
                            st.session_state.actividades.loc[idx, "% Avance"] = int(nv_av)
                            st.session_state.actividades.loc[idx, "Comentario"] = str(nv_co)
                            st.session_state.actividades.loc[idx, "Evidencia"] = str(ruta_foto_final)
                            try:
                                st.session_state.actividades.to_excel(ARCHIVO_DB, index=False)
                                st.success("¡Guardado!"); st.rerun()
                            except Exception as e_save: st.error(f"Fallo Excel: {e_save}")
                st.markdown('</div>', unsafe_allow_html=True)

elif opcion_menu == "📥 Cargar Actividades (Usuario)":
    st.subheader("Captura de Nuevas Actividades")
    if st.text_input("Contraseña de Usuario:", type="password", key="pwd_carga_usr") == "Metales":
        with st.form("form_usuario_carga"):
            o, p, r, a = st.selectbox("Origen", LISTA_CLASIFICACIONES), st.selectbox("Prioridad", ["Baja", "Media", "Urgente"]), st.selectbox("Responsable", list(st.session_state.personal.keys())), st.selectbox("Área", st.session_state.areas)
            d, f = st.text_area("Descripción"), st.date_input("Fecha Compromiso")
            if st.form_submit_button("Registrar Actividad"):
                n_id = int(st.session_state.actividades["No"].max() + 1) if not st.session_state.actividades.empty else 1
                n_f = {"No": n_id, "Origen": o, "Fecha Inicio": datetime.now().strftime("%d-%b-%y"), "Prioridad": p, "Responsable": r, "Area": a, "Descripcion": d, "% Avance": 0, "Fecha Compromiso": f.strftime("%d-%b-%y"), "Comentario": "", "Evidencia": ""}
                st.session_state.actividades = pd.concat([st.session_state.actividades, pd.DataFrame([n_f])], ignore_index=True)
                try:
                    st.session_state.actividades.to_excel(ARCHIVO_DB, index=False)
                    st.success("¡Registrada!"); st.rerun()
                except Exception as e_add: st.error(f"Fallo Excel: {e_add}")
elif opcion_menu == "🔐 Panel Administrador":
    st.markdown('### Panel de Control Máster')
    if st.text_input("Introduce la contraseña Máster:", type="password", key="pwd_admin_master") == "SigramaMetales2026":
        st.success("Acceso Máster Autorizado")
        st.write("---")
        st.markdown("### 📊 Consola de Sincronización Automática con GitHub (API REST)")
        c_adm1, c_adm2 = st.columns(2)
        
        with c_adm1:
            st.info("🔄 Descarga los registros más recientes que estén guardados actualmente en tu repositorio de GitHub.")
            if st.button("📥 IMPORTAR BASE DE DATOS DESDE GITHUB", use_container_width=True, key="btn_import_master"):
                st.session_state.actividades = importar_registros_excel()
                st.success("¡Sincronizado! Datos actualizados desde GitHub."); st.rerun()
                
        with c_adm2:
            st.warning("💾 Guarda los cambios en Excel y ejecuta un PUT automático directo a la API de GitHub.")
            if st.button("🚀 RESPALDAR AND SUBIR DIRECTO A GITHUB", type="primary", use_container_width=True, key="btn_git_api_push_master"):
                import requests
                import base64
                import json
                
                try:
                    df_guardar = pd.DataFrame(st.session_state.actividades)
                    with pd.ExcelWriter(ARCHIVO_DB, engine='openpyxl') as w:
                        df_guardar.to_excel(w, index=False, sheet_name='Base_MCE')
                        ws = w.sheets['Base_MCE']
                        anchos = {'A': 10, 'B': 25, 'C': 15, 'D': 15, 'E': 22, 'F': 18, 'G': 45, 'H': 12, 'I': 20, 'J': 25, 'K': 40}
                        for col, ancho in anchos.items(): ws.column_dimensions[col].width = ancho
                        
                        from openpyxl.styles import PatternFill, Font
                        fill_verde = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")
                        fill_amarillo = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
                        fill_rojo = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
                        font_rojo = Font(color="721C24", bold=True)
                        font_normal = Font(color="000000")
                        hoy_dt = datetime.now()
                        
                        for row_idx in range(2, ws.max_row + 1):
                            try:
                                avance_val = int(str(ws.cell(row=row_idx, column=8).value).replace('%','').strip())
                                fecha_comp_str = str(ws.cell(row=row_idx, column=9).value).strip()
                                es_vencido = False
                                if avance_val < 100 and fecha_comp_str:
                                    try:
                                        if datetime.strptime(fecha_comp_str, "%d-%b-%y") < hoy_dt: es_vencido = True
                                    except: pass
                                for col_idx in range(1, 12):
                                    cell = ws.cell(row=row_idx, column=col_idx)
                                    if avance_val < 100 and es_vencido: cell.fill = fill_rojo; cell.font = font_rojo
                                    elif avance_val == 100: cell.fill = fill_verde; cell.font = font_normal
                                    elif avance_val > 0: cell.fill = fill_amarillo; cell.font = font_normal
                            except: pass

                    token_git = "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"
                    usuario_git = "jesusalbertomoraleslopez-byte"
                    repo_git = "matriz-mce-sigrama"
                    email_git = "jesusalbertomoraleslopez@gmail.com"
                    
                    url_api = f"https://github.com{usuario_git}/{repo_git}/contents/{ARCHIVO_DB}"
                    cabeceras = {"Authorization": f"token {token_git}", "Accept": "application/vnd.github.v3+json"}
                    
                    respuesta_get = requests.get(url_api, headers=cabeceras)
                    sha_archivo = None
                    if respuesta_get.status_code == 200: sha_archivo = respuesta_get.json().get("sha")
                    
                    with open(ARCHIVO_DB, "rb") as archivo_binario:
                        excel_base64 = base64.b64encode(archivo_binario.read()).decode("utf-8")
                    
                    datos_payload = {
                        "message": f"Sincronizacion MCE Planta ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')})",
                        "content": excel_base64, "branch": "main",
                        "committer": {"name": usuario_git, "email": email_git}
                    }
                    if sha_archivo: datos_payload["sha"] = sha_archivo
                        
                    respuesta_put = requests.put(url_api, headers=cabeceras, data=json.dumps(datos_payload))
                    if respuesta_put.status_code in (200, 201):
                        st.success(f"✅ ¡Éxito Absoluto! Base respaldada directamente en GitHub.")
                        st.balloons(); st.rerun()
                    else: st.error(f"❌ Error en la API: {respuesta_put.text}")
                except Exception as e_api: st.error(f"Fallo critico HTTP REST: {e_api}")
        st.write("---")
        t1, t2, t3 = st.tabs(["➕ Altas Catálogos", "✏️ Tabla de Edición Directa", "📥 Carga Masiva Excel"])
        with t1:
            n_n = st.text_input("Nombre de Colaborador:")
            if st.button("Registrar Empleado") and n_n: st.session_state.personal[n_n] = None; st.success("Registrado."); st.rerun()
        with t2:
            st.subheader("✏️ Edición en Caliente de la Matriz MCE")
            df_editable = pd.DataFrame(st.session_state.actividades)
            if not df_editable.empty:
                configuracion_columnas = {
                    "No": st.column_config.NumberColumn("No", disabled=True, format="%d"),
                    "Origen": st.column_config.SelectboxColumn("Origen", options=LISTA_CLASIFICACIONES, required=True),
                    "Prioridad": st.column_config.SelectboxColumn("Prioridad", options=["Baja", "Media", "Urgente"], required=True),
                    "Responsable": st.column_config.SelectboxColumn("Responsable", options=list(st.session_state.personal.keys())),
                    "Area": st.column_config.SelectboxColumn("Área", options=st.session_state.areas),
                    "Fecha Inicio": st.column_config.TextColumn("Fecha Inicio"),
                    "Descripcion": st.column_config.TextColumn("Descripción", width="large"),
                    "% Avance": st.column_config.NumberColumn("% Avance", min_value=0, max_value=100, format="%d%%"),
                    "Fecha Compromiso": st.column_config.TextColumn("Fecha Compromiso"),
                    "Comentario": st.column_config.TextColumn("Comentario", width="medium"),
                    "Evidencia": st.column_config.TextColumn("Ruta Evidencia", disabled=True)
                }
                df_modificado = st.data_editor(df_editable, column_config=configuracion_columnas, use_container_width=True, hide_index=True, num_rows="dynamic", key="editor_tabla_master")
                if st.button("💾 CONFIRMAR Y GUARDAR CAMBIOS EN LA MATRIZ", type="primary", use_container_width=True):
                    st.session_state.actividades = df_modificado
                    try:
                        st.session_state.actividades.to_excel(ARCHIVO_DB, index=False)
                        st.success("✅ ¡Sincronizado!"); st.rerun()
                    except Exception as e_master: st.error(f"Error: {e_master}")
            else: st.info("No hay registros.")
        with t3:
            st.subheader("Inyección de Datos por Carga Masiva")
            columnas_p = ["Origen", "Fecha Inicio", "Prioridad", "Responsable", "Area", "Descripcion", "% Avance", "Fecha Compromiso", "Comentario"]
            ej_g = pd.DataFrame([["Programa de Actividades", datetime.now().strftime("%d-%b-%y"), "Media", "Bryan Flores", "⚙️ Ingenieria", "Descripción...", 0, "15-Jun-26", "..."]], columns=columnas_p)
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine='openpyxl') as w:
                ej_g.to_excel(w, index=False, sheet_name='Plantilla')
                ws = w.sheets['Plantilla']
                anchos = {'A': 25, 'B': 15, 'C': 15, 'D': 22, 'E': 15, 'F': 45, 'G': 12, 'H': 20, 'I': 25}
                for col, ancho in anchos.items(): ws.column_dimensions[col].width = ancho
            st.download_button(label="📥 Descargar Plantilla Oficial (.xlsx)", data=buf.getvalue(), file_name="Plantilla_MCE.xlsx", key="btn_download_template_master")
            st.write("---")
            ex = st.file_uploader("Subir Excel modificado", type=["xlsx"], key="uploader_bulk_master")
            if ex is not None:
                df_ex = pd.read_excel(ex)
                if st.button("Confirmar Importación Masiva"):
                    if "No" not in df_ex.columns: df_ex.insert(0, "No", range(st.session_state.actividades["No"].max() + 1 if not st.session_state.actividades.empty else 1, (st.session_state.actividades["No"].max() + 1 if not st.session_state.actividades.empty else 1) + len(df_ex)))
                    if "Evidencia" not in df_ex.columns: df_ex["Evidencia"] = ""
                    st.session_state.actividades = pd.concat([st.session_state.actividades, df_ex], ignore_index=True)
                    try:
                        st.session_state.actividades.to_excel(ARCHIVO_DB, index=False)
                        st.success("¡Importado!"); st.rerun()
                    except Exception as e_b: st.error(f"Error: {e_b}")


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

# Creación automática de la carpeta de evidencias físicas si no existe en Windows
if not os.path.exists(CARPETA_EVIDENCIAS):
    os.makedirs(CARPETA_EVIDENCIAS)

# Desactivamos el límite de píxeles de la librería PIL para soportar capturas pesadas del taller
Image.MAX_IMAGE_PIXELS = None

# 1. MOTOR DE IMPORTACIÓN INTELIGENTE DE EXCEL
def importar_registros_excel():
    if os.path.exists(ARCHIVO_DB):
        try:
            df = pd.read_excel(ARCHIVO_DB)
            if not df.empty:
                # CORRECCIÓN DE FECHAS: Convierte números seriales de Excel a texto legible
                for col_fecha in ["Fecha Inicio", "Fecha Compromiso"]:
                    if col_fecha in df.columns:
                        def corregir_fecha_serial(val):
                            try:
                                if pd.isna(val) or str(val).strip() in ["None", "nan", "NaN", ""]:
                                    return ""
                                if str(val).replace('.0', '').isdigit():
                                    dias = int(str(val).replace('.0', ''))
                                    fecha_dt = datetime(1899, 12, 30) + timedelta(days=dias)
                                    return fecha_dt.strftime("%d-%b-%y")
                                return str(val).strip()
                            except:
                                return str(val)
                        df[col_fecha] = df[col_fecha].apply(corregir_fecha_serial)
                # Control estricto de porcentajes y consecutivos numéricos dentro del dataframe
                if "% Avance" in df.columns:
                    if df["% Avance"].max() <= 1.0 and df["% Avance"].max() > 0:
                        df["% Avance"] = df["% Avance"] * 100
                    df["% Avance"] = pd.to_numeric(df["% Avance"], errors="coerce").fillna(0).astype(int)
                
                if "No" in df.columns:
                    df["No"] = pd.to_numeric(df["No"], errors="coerce")
                    if df["No"].isnull().any(): 
                        df["No"] = range(1, len(df) + 1)
                    df["No"] = df["No"].astype(int)
                
                columnas_texto = ["Origen", "Prioridad", "Responsable", "Area", "Descripcion", "Comentario", "Evidencia"]
                for col in columnas_texto:
                    if col in df.columns:
                        df[col] = df[col].astype(str).replace(["None", "nan", "NaN"], "")
            else:
                df = pd.DataFrame(columns=["No", "Origen", "Fecha Inicio", "Prioridad", "Responsable", "Area", "Descripcion", "% Avance", "Fecha Compromiso", "Comentario", "Evidencia"])
            return df
        except Exception as e:
            st.error(f"Error al importar el archivo Excel: {e}")
            return pd.DataFrame(columns=["No", "Origen", "Fecha Inicio", "Prioridad", "Responsable", "Area", "Descripcion", "% Avance", "Fecha Compromiso", "Comentario", "Evidencia"])
    else:
        return pd.DataFrame(columns=["No", "Origen", "Fecha Inicio", "Prioridad", "Responsable", "Area", "Descripcion", "% Avance", "Fecha Compromiso", "Comentario", "Evidencia"])

# Carga inicial de datos desde memoria persistente
if 'actividades' not in st.session_state:
    st.session_state.actividades = importar_registros_excel()

# 2. Configuración de la interfaz web corporativa (Fondo Gris Claro de Planta)
st.set_page_config(page_title="SIGRAMA - Matriz MCE", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #ECEFF1 !important; }
    [data-testid="stSidebar"] { background-color: #CFD8DC !important; }
    .main-title { font-size:28px !important; font-weight: bold; color: #0C2340; text-align: left; margin-top: 0px; }
    .card-header { font-size: 20px !important; font-weight: bold; color: #0C2340; margin-bottom: 3px; }
    .card-desc { font-size: 16px !important; font-weight: 500; color: #333333; margin-bottom: 5px; }
    .card { padding: 12px 18px; border-radius: 8px; box-shadow: 0px 2px 4px rgba(0,0,0,0.05); background-color: #FFFFFF; margin-bottom: 10px; border: 1px solid #E0E0E0; }
    div[data-testid="stForm"] { padding: 15px !important; }
    </style>
""", unsafe_allow_html=True)

# 3. Catálogos Operativos con Anclas Visuales (Iconos de Área)
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

# 4. Función de Gráficos de Pareto (Mapeo Cromático: Amarillo Claro para No Terminadas y Rango Corregido)
def crear_grafico_pareto(df, columna, titulo):
    if df.empty:
        fig = graph_objects.Figure(); fig.update_layout(title=f"{titulo} (Sin Datos)"); return fig
    df = df.copy()
    df["Estado"] = df["% Avance"].apply(lambda x: "Terminada" if x == 100 else "Pendiente")
    orden_cat = df[columna].value_counts().index.tolist()
    counts = df[columna].value_counts().reset_index()
    counts.columns = [columna, 'Cantidad']
    counts['Porcentaje Acumulado'] = (counts['Cantidad'].cumsum() / len(df)) * 100
    
    fig = px.histogram(
        df, x=columna, y=None, color="Estado", category_orders={columna: orden_cat},
        color_discrete_map={"Terminada": "#2ECC71", "Pendiente": "#FEEA9A"}, title=titulo,
        labels={columna: str(columna), "count": "Cantidad", "Estado": "Estatus"}
    )
    fig.add_trace(graph_objects.Scatter(x=counts[columna], y=counts['Porcentaje Acumulado'], name="% Acumulado", yaxis="y2", line=dict(color="#FF5733", width=3)))
    
    fig.update_layout(
        yaxis=dict(title="Cantidad de Actividades"), 
        yaxis2=dict(title="% Acumulado", overlaying="y", side="right", range=[0, 105]), 
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1), 
        template="plotly_white", barmode="stack"
    )
    return fig
# 5. Carga e Inyección del Logotipo Corporativo Fijo (Ampliados 2x y Título de Planta)
nombre_logo = "LOGOTIPO COLOR (1).jfif"
imagen_logo = Image.open(nombre_logo) if os.path.exists(nombre_logo) else None

if imagen_logo is not None:
    col_logo_g, col_tit_g = st.columns([1, 4])
    with col_logo_g: 
        st.image(imagen_logo, width=220) 
    with col_tit_g: 
        st.markdown('<h2 style="color: #0C2340; margin-bottom: 0px; padding-top: 10px; font-weight: bold;">PLANTA METALES Y MAQUINADOS</h2>', unsafe_allow_html=True)
        st.markdown('<p class="main-title">MATRIZ DE COMUNICACIÓN EFECTIVA</p>', unsafe_allow_html=True)
else:
    st.markdown('<h2 style="color: #0C2340; text-align: center; font-weight: bold;">PLANTA METALES Y MAQUINADOS</h2>', unsafe_allow_html=True)
    st.markdown('<p class="main-title" style="text-align: center;">MATRIZ DE COMUNICACIÓN EFECTIVA</p>', unsafe_allow_html=True)

opcion_menu = st.sidebar.radio("Navegación", ["📊 Dashboard Principal", "📋 Tabla de Control", "📝 Actualizar Mis Avances", "📥 Cargar Actividades (Usuario)", "🔐 Panel Administrador"])

# --- TAB 1: DASHBOARD PRINCIPAL ---
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
        def clasificar_vencimientos(fila):
            try:
                f_comp = datetime.strptime(str(fila["Fecha Compromiso"]).strip(), "%d-%b-%y")
                return "Vencida (Retraso)" if (int(fila["% Avance"]) < 100 and f_comp < hoy) else ("Terminada" if int(fila["% Avance"]) == 100 else "Pendiente a Tiempo")
            except: return "Pendiente a Tiempo"
        df_lideres["Estado_Real"] = df_lideres.apply(clasificar_vencimientos, axis=1)
        df_lideres["Valor_Eje"] = df_lideres["Estado_Real"].apply(lambda x: -1 if x == "Vencida (Retraso)" else 1)
        
        fig_l = px.bar(
            df_lideres, x="Origen", y="Valor_Eje", color="Estado_Real", facet_col="Responsable", facet_col_wrap=2, 
            color_discrete_map={"Terminada": "#2ECC71", "Pendiente a Tiempo": "#FEEA9A", "Vencida (Retraso)": "#E74C3C"}, 
            labels={"Origen": "Clasificación", "Valor_Eje": "Tareas"}
        )
        fig_l.update_layout(barmode="stack", template="plotly_white", height=600, legend=dict(orientation="h", y=-0.2, x=0.5, xanchor="center"))
        fig_l.for_each_annotation(lambda a: a.update(text=f"<b>👤 {a.text.split('=')[-1].upper()}</b>", font=dict(size=14, color="#0C2340")))
        st.plotly_chart(fig_l, use_container_width=True)
# --- TAB 2: TABLA DE CONTROL COMPLETA ---
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
        df_estilizado = df_mostrar.style.apply(aplicar_colores_renglon, axis=1).format({"% Avance": "{:.0f}%"})
        st.dataframe(df_estilizado, use_container_width=True, hide_index=True)
    else: st.info("No se encontraron registros.")
# --- TAB 3: ACTUALIZAR MIS AVANCES (CON FIX DE MUTACIÓN DIRECTA EN DATA FRAME .LOC Y EXPORT EN CALIENTE) ---
elif opcion_menu == "📝 Actualizar Mis Avances":
    st.subheader("Actualización de Avances de Tareas")
    u = st.selectbox("Identifícate (Selecciona tu nombre)", list(st.session_state.personal.keys()))
    df_usuario = st.session_state.actividades[st.session_state.actividades["Responsable"] == u]
    
    if df_usuario.empty: 
        st.info("No tienes actividades pendientes asignadas.")
    else:
        for idx in df_usuario.index:
            r = st.session_state.actividades.loc[idx]
            with st.container():
                st.markdown('<div class="card">', unsafe_allow_html=True)
                st.markdown(f'<p class="card-header">Actividad No. {r["No"]} | {r["Area"]} | Prioridad: {r["Prioridad"]}</p>', unsafe_allow_html=True)
                st.markdown(f'<p class="card-desc"><b>Descripción:</b> {r["Descripcion"]}</p>', unsafe_allow_html=True)
                
                col_izq, col_der = st.columns(2)
                with col_izq:
                    progreso_actual = int(r['% Avance'])
                    fig_slider = graph_objects.Figure()
                    fig_slider.add_trace(graph_objects.Bar(x=["Progreso"], y=[100], marker_color="#E0E0E0", showlegend=False, hoverinfo="none"))
                    color_barra = "#2ECC71" if progreso_actual == 100 else "#0C2340"
                    fig_slider.add_trace(graph_objects.Bar(x=["Progreso"], y=[progreso_actual], marker_color=color_barra, showlegend=False, text=f"{progreso_actual}%", textposition="inside", textfont=dict(size=14, color="white")))
                    fig_slider.update_layout(barmode="overlay", template="plotly_white", height=140, width=90, margin=dict(l=5, r=5, t=5, b=5), xaxis=dict(visible=False), yaxis=dict(range=[0, 100], showgrid=False, zeroline=False, visible=False))
                    st.plotly_chart(fig_slider, use_container_width=False, config={'displayModeBar': False}, key=f"plot_chart_{r['No']}")
                    nv_av = st.slider("Ajustar %:", min_value=0, max_value=100, value=progreso_actual, step=5, key=f"num_{r['No']}")
                
                with col_der:
                    nv_co = st.text_input("Comentarios de bitácora:", str(r['Comentario']), key=f"c_{r['No']}")
                    evidencia_guardada = str(r['Evidencia']).strip()
                    if evidencia_guardada and os.path.exists(evidencia_guardada):
                        st.image(Image.open(evidencia_guardada), width=150, caption="📸 Evidencia Optimizada")
                    
                    foto = st.file_uploader("Evidencia Fotográfica (Cierre 100%):", type=["jpg","png","jpeg","jfif"], key=f"i_{r['No']}") if nv_av == 100 else None
                    if foto: st.image(Image.open(foto), width=120, caption="Vista Previa")
                    
                    st.markdown('<div style="margin-top: 5px;">', unsafe_allow_html=True)
                    if st.button("Guardar Tarea", key=f"b_{r['No']}"):
                        if nv_av == 100 and not foto and not evidencia_guardada: 
                            st.error("Falta la evidencia fotográfica obligatoria para cerrar al 100%.")
                        else:
                            ruta_foto_final = evidencia_guardada
                            if foto is not None:
                                try:
                                    img_abierta = Image.open(foto)
                                    if img_abierta.mode in ("RGBA", "P"): img_abierta = img_abierta.convert("RGB")
                                    img_abierta.thumbnail((800, 600), Image.Resampling.LANCZOS)
                                    stamp = datetime.now().strftime("%Y%m%d")
                                    nombre_archivo_final = f"MCE-{int(r['No']):03d} evidencia ({stamp}).jpg"
                                    ruta_foto_final = os.path.join(CARPETA_EVIDENCIAS, nombre_archivo_final)
                                    img_abierta.save(ruta_foto_final, "JPEG", quality=65)
                                except Exception as err_img: st.error(f"Error al procesar la imagen física: {err_img}")
                            
                            st.session_state.actividades.loc[idx, "% Avance"] = int(nv_av)
                            st.session_state.actividades.loc[idx, "Comentario"] = str(nv_co)
                            st.session_state.actividades.loc[idx, "Evidencia"] = str(ruta_foto_final)
                            
                            try:
                                st.session_state.actividades.to_excel(ARCHIVO_DB, index=False)
                                st.success("¡Avance registrado exitosamente en caliente!"); st.rerun()
                            except Exception as e_save: st.error(f"Fallo al escribir en Excel: {e_save}")
                    st.markdown('</div>', unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)

# --- TAB 4: CARGAR ACTIVIDADES UNIVERSALES ---
elif opcion_menu == "📥 Cargar Actividades (Usuario)":
    st.subheader("Captura de Nuevas Actividades - Planta Metales")
    if st.text_input("Contraseña de Usuario:", type="password", key="pwd_carga_usr") == "Metales":
        with st.form("form_usuario_carga"):
            o, p, r, a = st.selectbox("Origen", LISTA_CLASIFICACIONES), st.selectbox("Prioridad", ["Baja", "Media", "Urgente"]), st.selectbox("Responsable", list(st.session_state.personal.keys())), st.selectbox("Área", st.session_state.areas)
            d, f = st.text_area("Descripción de Actividad"), st.date_input("Fecha Compromiso")
            if st.form_submit_button("Registrar Actividad"):
                n_id = int(st.session_state.actividades["No"].max() + 1) if not st.session_state.actividades.empty else 1
                n_f = {"No": n_id, "Origen": o, "Fecha Inicio": datetime.now().strftime("%d-%b-%y"), "Prioridad": p, "Responsable": r, "Area": a, "Descripcion": d, "% Avance": 0, "Fecha Compromiso": f.strftime("%d-%b-%y"), "Comentario": "", "Evidencia": ""}
                st.session_state.actividades = pd.concat([st.session_state.actividades, pd.DataFrame([n_f])], ignore_index=True)
                try:
                    st.session_state.actividades.to_excel(ARCHIVO_DB, index=False)
                    st.success("¡Registrada con éxito!"); st.rerun()
                except Exception as e_add: st.error(f"Fallo al escribir en Excel: {e_add}")
# --- TAB 5: PANEL ADMINISTRADOR MÁSTER ---
elif opcion_menu == "🔐 Panel Administrador":
    st.markdown('<p class="admin-header" style="font-size:24px; font-weight:bold; color:#0C2340; margin-bottom:15px;">Panel de Control Máster</p>', unsafe_allow_html=True)
    if st.text_input("Introduce la contraseña Máster:", type="password", key="pwd_admin_master") == "SigramaMetales2026":
        st.success("Acceso Máster Autorizado")
        
        st.write("---")
        st.markdown("### 📊 Consola de Sincronización Automática Bidireccional (Streamlit 🔄 GitHub)")
        c_adm1, c_adm2 = st.columns(2)
        
        with c_adm1:
            st.info("🔄 Descarga los registros más recientes que estén guardados actualmente en tu repositorio de GitHub.")
            if st.button("📥 IMPORTAR BASE DE DATOS DESDE GITHUB", use_container_width=True, key="btn_import_master"):
                st.session_state.actividades = importar_registros_excel()
                st.success("¡Sincronizado! Datos actualizados desde GitHub."); st.rerun()
                
        with c_adm2:
            st.warning("💾 Guarda los cambios en Excel y ejecuta un GIT PUSH automático directo a tu cuenta de GitHub.")
            if st.button("💾 RESPALDAR Y CONFIRMAR CAMBIOS EN GITHUB", type="primary", use_container_width=True, key="btn_backup_master"):
                try:
                    df_guardar = pd.DataFrame(st.session_state.actividades)
                    with pd.ExcelWriter(ARCHIVO_DB, engine='openpyxl') as w:
                        df_guardar.to_excel(w, index=False, sheet_name='Base_MCE')
                        ws = w.sheets['Base_MCE']
                        anchos = {'A': 10, 'B': 25, 'C': 15, 'D': 15, 'E': 22, 'F': 18, 'G': 45, 'H': 12, 'I': 20, 'J': 25, 'K': 40}
                        for col, ancho in anchos.items(): 
                            ws.column_dimensions[col].width = ancho
                        
                        from openpyxl.worksheet.datavalidation import DataValidation
                        str_orígenes = f'"{",".join(LISTA_CLASIFICACIONES)}"'
                        str_prioridades = '"Baja,Media,Urgente"'
                        str_personal = f'"{",".join(list(st.session_state.personal.keys()))}"'
                        str_areas = f'"{",".join(st.session_state.areas)}"'
                        
                        dv_o = DataValidation(type="list", formula1=str_orígenes, allow_blank=True)
                        dv_p = DataValidation(type="list", formula1=str_prioridades, allow_blank=True)
                        dv_r = DataValidation(type="list", formula1=str_personal, allow_blank=True)
                        dv_a = DataValidation(type="list", formula1=str_areas, allow_blank=True)
                        
                        dv_o.add("B2:B1000"); dv_p.add("D2:D1000"); dv_r.add("E2:E1000"); dv_a.add("F2:F1000")
                        ws.add_data_validation(dv_o); ws.add_data_validation(dv_p); ws.add_data_validation(dv_r); ws.add_data_validation(dv_a)
                    
                    # MOTOR AUTOMÁTICO DE CONEXIÓN ROBÓTICA A GITHUB VIA CONSOLA
                    git_token = st.secrets["github"]["token"]
                    git_user = st.secrets["github"]["usuario"]
                    git_repo = st.secrets["github"]["repo"]
                    git_email = st.secrets["github"]["email"]
                    
                    os.system(f'git config --global user.email "{git_email}"')
                    os.system(f'git config --global user.name "{git_user}"')
                    os.system(f'git add {ARCHIVO_DB}')
                    os.system('git commit -m "Persistence Automation: Sincronización forzada de Excel MCE en caliente"')
                    os.system(f'git push https://{git_user}:{git_token}@://github.com{git_user}/{git_repo}.git main')
                    
                    st.success(f"✅ ¡Éxito Absoluto! Los {len(df_guardar)} registros se guardaron en Excel y se empujaron automáticamente a tu cuenta en GitHub.")
                    st.rerun()
                except Exception as e: 
                    st.error(f"❌ Error en la sincronización robótica: {e}")
                    
        st.write("---")
        t1, t2, t3 = st.tabs(["➕ Altas Catálogos", "✏️ Tabla de Edición Directa y Bajas", "📥 Carga Masiva Excel"])
        with t1:
            n_n = st.text_input("Nombre de Colaborador:")
            if st.button("Registrar Empleado", key="btn_add_emp") and n_n: 
                st.session_state.personal[n_n] = None; st.success("Registrado."); st.rerun()
            st.write("---")
            n_a = st.text_input("Nombre de Área:")
            if st.button("Registrar Área", key="btn_add_area") and n_a: 
                st.session_state.areas.append(n_a); st.success("Añadida."); st.rerun()
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
                    "

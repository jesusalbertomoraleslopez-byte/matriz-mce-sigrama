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
                # CORRECCIÓN DE FECHAS: Detecta formatos datetime extensos y los limpia a texto DD-Mes-YY
                for col_fecha in ["Fecha Inicio", "Fecha Compromiso"]:
                    if col_fecha in df.columns:
                        def corregir_fecha_serial(val):
                            try:
                                if pd.isna(val) or str(val).strip() in ["None", "nan", "NaN", ""]: 
                                    return ""
                                if isinstance(val, (datetime, pd.Timestamp)):
                                    return val.strftime("%d-%b-%y")
                                if " " in str(val):
                                    val = str(val).split(" ")[0]
                                if str(val).replace('.0', '').isdigit():
                                    dias = int(str(val).replace('.0', ''))
                                    return (datetime(1899, 12, 30) + timedelta(days=dias)).strftime("%d-%b-%y")
                                return str(val).strip()
                            except: 
                                return str(val)
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

    
        /* CONFIGURACIÓN MÁSTER 3X PARA EL NOMBRE DEL OPERADOR */
       /* CONFIGURACIÓN MÁSTER EXCLUSIVA: AFECTA ÚNICAMENTE AL SELECTOR DE LA PESTAÑA DE OPERADORES */
    div[data-testid="stSidebar"] ~ div div[class*="stSelectbox"] div[data-baseweb="select"] {
        font-size: 36px !important; 
        font-weight: 800 !important; 
        color: #0C2340 !important;  
        height: 85px !important; /* Aumentamos a 85px para dar espacio total al texto */
        min-height: 85px !important;
        display: flex !important;
        align-items: center !important;
    }
    
    div[data-testid="stSidebar"] ~ div div[class*="stSelectbox"] [data-testid="stSelectbox-SingleValue"],
    div[data-testid="stSidebar"] ~ div div[class*="stSelectbox"] div[data-baseweb="select"] span {
        line-height: 85px !important;
        font-size: 36px !important;
        overflow: visible !important; 
    }

    /* Calibración de la lista desplegable de nombres gigante */
    div[data-baseweb="popover"] ul li {
        font-size: 28px !important;
        padding-top: 10px !important;
        padding-bottom: 10px !important;
        line-height: 1.2 !important;
    }

    /* FILTRO DE EXTINCIÓN: Mantiene todos los selectores de las tablas de datos en su tamaño normal original */
    div[data-testid="stHorizontalBlock"] div[data-testid="stSelectbox"] div[data-baseweb="select"],
    div[data-testid="element-container"] div[data-testid="stSelectbox"] div[data-baseweb="select"] {
        font-size: 14px !important;
        height: auto !important;
        min-height: auto !important;
    }
    </style>
""", unsafe_allow_html=True)

if 'personal' not in st.session_state:
    st.session_state.personal = {
        "Jesus Morales": None, "Cruz Carreon": None, "Luis Quintana": None, "Bryan Flores": None, "Rodolfo Ferndez M": None,"Ing. Alfredo Hdz": None, "Ing. Lorena Hdz": None,
        "Alejandra Arellano": None, "Jose Romo": None, "Jose Manuel Aldama": None, "Fernando Llanas": None,
        "Celso": None,  "Josue Mesta": None,
        "Jorge Sanchez": None, "Voctor Montoya": None
    }

if 'areas' not in st.session_state:
    st.session_state.areas = [
        "⚙️ Ingenieria", 
        "🔍 Calidad", 
        "📦 Almacen", 
        "✂️ Corte", 
        "📐 Doblez", 
        "🎨 Pintura", 
        "🚚 Embarquez", 
        "🏭 Planta Rio", 
        "🛠️ Lijado",
        "💼 Administracion",      # Nueva área añadida
        "👥 Recursos Humanos",    # Nueva área añadida
        "👑 Direccion"            # Nueva área añadida
    ]



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

# 5. Carga e Inyección del Banner Corporativo Superior Giant 2271x448
nombre_banner = "LIDERAZGO PLANTA METALAES IMAGEN.png"

# Verificamos si existe el archivo del banner en el repositorio
if os.path.exists(nombre_banner):
    imagen_banner = Image.open(nombre_banner)
    # Mostramos el banner expandido al ancho total de la página
    st.image(imagen_banner, use_container_width=True)
else:
    # Respaldo en texto limpio y centrado por si borran la imagen del repositorio
    st.markdown('<h2 style="color: #0C2340; text-align: center; font-weight: bold; margin-top:0px;">PLANTA METALES Y MAQUINADOS</h2>', unsafe_allow_html=True)
    st.markdown('<p class="main-title" style="text-align: center;">MATRIZ DE COMUNICACIÓN EFECTIVA</p>', unsafe_allow_html=True)



opcion_menu = st.sidebar.radio("Navegación", ["📊 Dashboard Principal", "📋 Tabla de Control", "📝 Actualizar Mis Avances", "📥 Cargar Actividades (Usuario)", "🔐 Panel Administrador", "👑 Reglas de Liderazgo", "📋 Reportes PDF"])


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
        # REGLA SOLICITADA: Semáforo Amarillo Automático para Avances Diferentes de Cero
        def aplicar_colores_renglon(fila):
            try:
                num_avance = int(fila["% Avance"])
                fecha_comp_str = str(fila["Fecha Compromiso"]).strip()
                fecha_vencida = False
                if fecha_comp_str:
                    try:
                        f_comp = datetime.strptime(fecha_comp_str, "%d-%b-%y")
                        if f_comp < hoy: fecha_vencida = True
                    except: pass
                if num_avance < 100 and fecha_vencida: 
                    return ['background-color: #F8D7DA; color: #721C24; font-weight: 500;'] * len(fila)
                elif num_avance == 100: 
                    return ['background-color: #D4EDDA; color: #155724;'] * len(fila)
                elif num_avance > 0: 
                    return ['background-color: #FFF3CD; color: #856404;'] * len(fila)
            except: pass
            return [''] * len(fila)

  
    st.dataframe(df_mostrar.style.apply(aplicar_colores_renglon, axis=1).format({"% Avance": "{:.0f}%"}), use_container_width=True, hide_index=True)
    
    import io

    def generar_excel_con_colores(df_local):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_local.to_excel(writer, index=False, sheet_name='Historial_MCE')
            workbook = writer.book
            worksheet = writer.sheets['Historial_MCE']
            
            for col in worksheet.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                # CAMBIO AQUÍ: Agregamos [0] a col
                col_letter = col[0].column_letter
                worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)
                
            from openpyxl.styles import PatternFill, Font
            fill_verde = PatternFill(start_color="D4EDDA", fill_type="solid")
            fill_amarillo = PatternFill(start_color="FFF3CD", fill_type="solid")
            fill_rojo = PatternFill(start_color="F8D7DA", fill_type="solid")
            font_rojo = Font(color="721C24", bold=True)
            
            for row_idx in range(2, worksheet.max_row + 1):
                try:
                    avance_val = int(str(worksheet.cell(row=row_idx, column=8).value).replace('%','').strip())
                    fecha_comp_str = str(worksheet.cell(row=row_idx, column=9).value).strip()
                    
                    es_vencido = False
                    if avance_val < 100 and fecha_comp_str and fecha_comp_str != "None":
                        from datetime import datetime
                        if datetime.strptime(fecha_comp_str, "%d-%b-%y") < datetime.now():
                            es_vencido = True
                    
                    for col_idx in range(1, worksheet.max_column + 1):
                        cell = worksheet.cell(row=row_idx, column=col_idx)
                        if avance_val < 100 and es_vencido:
                            cell.fill = fill_rojo; cell.font = font_rojo
                        elif avance_val == 100:
                            cell.fill = fill_verde
                        elif avance_val > 0:
                            cell.fill = fill_amarillo
                except:
                    pass
        return output.getvalue()

    excel_data = generar_excel_con_colores(df_mostrar)
    st.download_button(
        label="📥 Descargar Reporte en Excel (.xlsx con Colores)",
        data=excel_data,
        file_name=f"Reporte_Matriz_MCE_{datetime.now().strftime('%d-%b-%y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# --- TAB 3: ACTUALIZAR MIS AVANCES (CON FILTROS CORPORATIVOS, MINI DASHBOARD Y LISTADO EN CASCADA) ---
elif opcion_menu == "📝 Actualizar Mis Avances":
    st.subheader("Actualización de Avances de Tareas")
    
    # 1. ENTORNO DE ENCABEZADO: FILTRO DE PERSONA Y MINI DASHBOARD SUPERIOR DERECHA
    u = st.selectbox("Identifícate (Selecciona tu nombre)", list(st.session_state.personal.keys()))
    
    # Extraemos las actividades exclusivas de este usuario para calcular sus métricas individuales
    df_usuario = pd.DataFrame(st.session_state.actividades)
    if not df_usuario.empty and "Responsable" in df_usuario.columns:
        df_usuario = df_usuario[df_usuario["Responsable"] == u]
    
    if df_usuario.empty: 
        st.info(f"👤 {u} no tiene actividades pendientes asignadas en este momento.")
    else:
        # MINI DASHBOARD INCORPORADO EN LA PARTE SUPERIOR (Dividido en columnas limpias)
        st.markdown('<p style="font-size:16px; font-weight:bold; color:#0C2340; margin-bottom:5px;">📊 Mi Rendimiento Actual</p>', unsafe_allow_html=True)
        col_dash1, col_dash2, col_dash3 = st.columns(3)
        
        tareas_pendientes = len(df_usuario[df_usuario["% Avance"].astype(int) < 100])
        tareas_hechas = len(df_usuario[df_usuario["% Avance"].astype(int) == 100])
        promedio_avance = df_usuario["% Avance"].astype(int).mean()
        
        col_dash1.metric(label="⏳ En Proceso", value=f"{tareas_pendientes} tareas")
        col_dash2.metric(label="✅ Terminadas", value=f"{tareas_hechas} tareas")
        col_dash3.metric(label="📈 Eficiencia Total", value=f"{promedio_avance:.1f}%")
        st.write("---")
        
        # 2. FILTRO DE 3 CLASIFICACIONES SOLICITADAS
        clasificacion = st.radio(
            "Filtrar mi lista por estatus:",
            ["En proceso (0% a 99%)", "Terminadas (100%)", "Ver Todas las Asignadas"],
            horizontal=True
        )
        
        # Filtramos el DataFrame local según la clasificación seleccionada por el operador
        if clasificacion == "En proceso (0% a 99%)":
            df_filtrado = df_usuario[df_usuario["% Avance"].astype(int) < 100]
        elif clasificacion == "Terminadas (100%)":
            df_filtrado = df_usuario[df_usuario["% Avance"].astype(int) == 100]
        else:
            df_filtrado = df_usuario.copy()
            
        st.write("---")
        
        if df_filtrado.empty:
            st.info(f"📋 No hay tareas en la clasificación '{clasificacion}' para {u}.")
        else:
            # 3. FORMATO DE LISTA VERTICAL EN CASCADA COMPACTA
            for idx in df_filtrado.index:
                r = st.session_state.actividades.loc[idx]
                
                # Contenedor visual tipo lista minimalista para planta
                with st.container():
                    st.markdown('<div class="card">', unsafe_allow_html=True)
                    
                    # Encabezado compacto en una sola línea
                    st.markdown(f'<p class="card-header">📋 Actividad No. {r["No"]} | {r["Area"]} | Prioridad: {r["Prioridad"]}</p>', unsafe_allow_html=True)
                    st.markdown(f'<p class="card-desc" style="font-size:15px; margin-bottom:15px;"><b>Descripción:</b> {r["Descripcion"]}</p>', unsafe_allow_html=True)
                    
                    col_izq, col_der = st.columns(2)
                    
                    with col_izq:
                        progreso_actual = int(r['% Avance'])
                        
                        # Renderizado del medidor de barra vertical
                        fig_slider = graph_objects.Figure()
                        fig_slider.add_trace(graph_objects.Bar(x=["Progreso"], y=[100], marker_color="#E0E0E0", showlegend=False, hoverinfo="none"))
                        color_barra = "#2ECC71" if progreso_actual == 100 else "#0C2340"
                        fig_slider.add_trace(graph_objects.Bar(x=["Progreso"], y=[progreso_actual], marker_color=color_barra, showlegend=False, text=f"{progreso_actual}%", textposition="inside", textfont=dict(size=14, color="white")))
                        fig_slider.update_layout(barmode="overlay", template="plotly_white", height=140, width=90, margin=dict(l=5, r=5, t=5, b=5), xaxis=dict(visible=False), yaxis=dict(range=[0, 100], showgrid=False, zeroline=False, visible=False))
                        st.plotly_chart(fig_slider, use_container_width=False, config={'displayModeBar': False}, key=f"plot_chart_{r['No']}")
                        
                        # Deslizador interactivo alineado a la lista
                        nv_av = st.slider("Ajustar %:", min_value=0, max_value=100, value=progreso_actual, step=5, key=f"num_{r['No']}")
                    
                    with col_der:
                        # Control de comentarios limpiando textos vacíos o nulos 'nan'
                        comentario_limpio = "" if str(r['Comentario']).strip().lower() in ["nan", "none", ""] else str(r['Comentario'])
                        nv_co = st.text_input("Comentarios de bitácora:", value=comentario_limpio, key=f"c_{r['No']}")
                        
                        # Despliegue de evidencia previa si existe
                        evidencia_guardada = str(r['Evidencia']).strip()
                        if evidencia_guardada and os.path.exists(evidencia_guardada):
                            st.image(Image.open(evidencia_guardada), width=130, caption="📸 Evidencia Actual")
                        
                        # Carga de fotos obligatoria al marcar 100%
                        foto = st.file_uploader("Evidencia Fotográfica (Cierre 100%):", type=["jpg","png","jpeg","jfif"], key=f"i_{r['No']}") if nv_av == 100 else None
                        if foto: 
                            st.image(Image.open(foto), width=100, caption="Vista Previa")
                        
                        st.markdown('<div style="margin-top: 10px;">', unsafe_allow_html=True)
                        if st.button("Guardar Tarea", key=f"b_{r['No']}", use_container_width=True):
                            if nv_av == 100 and not foto and not evidencia_guardada: 
                                st.error("⚠️ Faltan las fotografías físicas obligatorias para autorizar el cierre al 100%.")
                            else:
                                ruta_foto_final = evidencia_guardada
                                if foto is not None:
                                    try:
                                        img_abierta = Image.open(foto)
                                        if img_abierta.mode in ("RGBA", "P"): 
                                            img_abierta = img_abierta.convert("RGB")
                                        img_abierta.thumbnail((800, 600), Image.Resampling.LANCZOS)
                                        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                        nombre_archivo_final = f"MCE-{int(r['No']):03d}_evidencia_{stamp}.jpg"
                                        ruta_foto_final = os.path.join(CARPETA_EVIDENCIAS, nombre_archivo_final)
                                        img_abierta.save(ruta_foto_final, "JPEG", quality=65)
                                    except Exception as err_img: 
                                        st.error(f"Fallo al procesar captura: {err_img}")
                                
                                # Inyección en caliente en el DataFrame virtual de sesión
                                st.session_state.actividades.loc[idx, "% Avance"] = int(nv_av)
                                st.session_state.actividades.loc[idx, "Comentario"] = str(nv_co)
                                st.session_state.actividades.loc[idx, "Evidencia"] = str(ruta_foto_final)
                                
                                # Persistencia inmediata en el archivo Excel físico de Planta
                                try:
                                    st.session_state.actividades.to_excel(ARCHIVO_DB, index=False)
                                    st.success("🏁 ¡Cambios registrados con éxito!"); st.rerun()
                                except Exception as e_save: 
                                    st.error(f"Fallo de escritura en base física Excel: {e_save}")
                        st.markdown('</div>', unsafe_allow_html=True)
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
            if st.button("📥 IMPORTAR BASE DE DATOS DESDE GITHUB", use_container_width=True, key="btn_import_master"):
                st.session_state.actividades = importar_registros_excel()
                st.success("¡Sincronizado! Datos actualizados desde GitHub."); st.rerun()
        with c_adm2:
            if st.button("🚀 RESPALDAR AND SUBIR DIRECTO A GITHUB", type="primary", use_container_width=True, key="btn_git_api_push_master"):
                import requests, base64, json
                try:
                    df_guardar = pd.DataFrame(st.session_state.actividades)
                    with pd.ExcelWriter(ARCHIVO_DB, engine='openpyxl') as w:
                        df_guardar.to_excel(w, index=False, sheet_name='Base_MCE')
                        ws = w.sheets['Base_MCE']
                        anchos = {'A': 10, 'B': 25, 'C': 15, 'D': 15, 'E': 22, 'F': 18, 'G': 45, 'H': 12, 'I': 20, 'J': 25, 'K': 40}
                        for col, ancho in anchos.items(): ws.column_dimensions[col].width = ancho
                        from openpyxl.styles import PatternFill, Font
                        fill_verde, fill_amarillo, fill_rojo = PatternFill(start_color="D4EDDA", fill_type="solid"), PatternFill(start_color="FFF3CD", fill_type="solid"), PatternFill(start_color="F8D7DA", fill_type="solid")
                        font_rojo, font_normal = Font(color="721C24", bold=True), Font(color="000000")
                        hoy_dt = datetime.now()
                        for row_idx in range(2, ws.max_row + 1):
                            try:
                                avance_val = int(str(ws.cell(row=row_idx, column=8).value).replace('%','').strip())
                                fecha_comp_str = str(ws.cell(row=row_idx, column=9).value).strip()
                                es_vencido = False
                                if avance_val < 100 and fecha_comp_str:
                                    if datetime.strptime(fecha_comp_str, "%d-%b-%y") < hoy_dt: es_vencido = True
                                for col_idx in range(1, 12):
                                    cell = ws.cell(row=row_idx, column=col_idx)
                                    if avance_val < 100 and es_vencido: cell.fill = fill_rojo; cell.font = font_rojo
                                    elif avance_val == 100: cell.fill = fill_verde; cell.font = font_normal
                                    elif avance_val > 0: cell.fill = fill_amarillo; cell.font = font_normal
                            except: pass
                            # --- REEMPLAZO SEGURO DESDE LA LÍNEA 285 ---
                        try:
                            token_git = st.secrets["TOKEN_GITHUB"]
                        except Exception:
                            token_git = "" # Respaldo por si realizas pruebas locales fuera de la nube
                
                        usuario_git = "jesusalbertomoraleslopez-byte"
                        repo_git = "matriz-mce-sigrama"
                        email_git = "jesusalbertomoraleslopez@gmail.com"
                
                        url_api = "https://" + "api.gi" + "thub.com" + "/repos/" + usuario_git + "/" + repo_git + "/contents/base_matriz_mce.xlsx"
                        cabeceras = {"Authorization": f"token {token_git}", "Accept": "application/vnd.github.v3+json"}
                        respuesta_get = requests.get(url_api, headers=cabeceras)
                        sha_archivo = None
                        if respuesta_get.status_code == 200:
                            try:
                                sha_archivo = respuesta_get.json().get("sha")
                            except Exception:
                                pass
                        elif respuesta_get.status_code != 404:
                            st.error(f"Fallo de comunicación con GitHub API (Código: {respuesta_get.status_code}).")
                            st.stop()

                    with open(ARCHIVO_DB, "rb") as archivo_binario: excel_base64 = base64.b64encode(archivo_binario.read()).decode("utf-8")
                    datos_payload = {"message": f"Sincronizacion MCE Planta ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')})", "content": excel_base64, "branch": "main", "committer": {"name": usuario_git, "email": email_git}}
                    if sha_archivo: datos_payload["sha"] = sha_archivo
                    respuesta_put = requests.put(url_api, headers=cabeceras, data=json.dumps(datos_payload))
                    if respuesta_put.status_code in (200, 201):
                        st.success(f"✅ ¡Éxito Absoluto! Base respaldada directamente en GitHub."); st.balloons(); st.rerun()
                    else: st.error(f"❌ Error en la API. Respuesta: {respuesta_put.text}")
                except Exception as error_global_api: st.error(f"Fallo critico HTTP REST: {error_global_api}")
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

# --- NUEVA SECCIÓN: REGLAS DE LIDERAZGO ---

elif opcion_menu == "👑 Reglas de Liderazgo":
    st.markdown('<h2 style="color: #0C2340; font-weight: bold; margin-bottom: 20px;">👑 REGLAS DE LIDERAZGO: PLANTA METALES</h2>', unsafe_allow_html=True)
    st.write("Guía oficial de comportamiento, gestión y control en piso de producción.")
    st.write("---")

    col_l1, col_l2 = st.columns(2)

    with col_l1:
        with st.expander("🪵 Liderazgo y Comportamiento en Piso", expanded=True):
            st.markdown("""
            1. **Predica con el ejemplo:** Cumple las normas de seguridad, puntualidad y vestimenta antes que nadie.
            2. **Lidera desde el Gemba:** Pasa el 80% de tu tiempo en el piso de producción, no en la oficina.
            3. **Corrige en privado, reconoce en público:** Protege la dignidad de tu personal y celebra sus éxitos frente al grupo.
            4. **Escucha antes de juzgar:** Investiga la causa raíz de un problema antes de buscar culpables.
            5. **Mantén la consistencia:** Aplica las reglas y sanciones de manera justa, equitativa y sin favoritismos.
            """, unsafe_allow_html=True)

        with st.expander("💼 Operación y Gestión del Personal", expanded=True):
            st.markdown("""
            1. **Cero tolerancia al ocio:** Asegura que cada operador tenga una tarea asignada en todo momento.
            2. **Activa el plan B de inmediato:** Si la producción se detiene, reasigna al personal a actividades secundarias (5S, capacitación, mantenimiento).
            3. **Informa desviaciones a Dirección:** Reporta al final del turno cualquier falta de actividad programada y justifica el uso del tiempo.
            4. **Capacita constantemente:** Desarrollar la polivalencia de tu equipo reduce la dependencia de personas específicas.
            5. **La seguridad es primero:** Detén cualquier operación que ponga en riesgo la integridad física del personal.
            """, unsafe_allow_html=True)

    with col_l2:
        with st.expander("📢 Comunicación y Relaciones Interdepartamentales", expanded=True):
            st.markdown("""
            1. **Arranca con juntas Tier 1:** Realiza reuniones de 5 minutos al inicio del turno para alinear metas y riesgos.
            2. **Comunícate con datos:** Al solicitar apoyo a Calidad o Mantenimiento, habla con números y hechos, no con opiniones.
            3. **Define el 'para cuándo':** Al delegar o solicitar una tarea a otra jefatura, establece siempre una fecha y hora límite.
            4. **Cierra el ciclo de comunicación:** Confirma que el receptor entendió el mensaje pidiéndole que lo explique con sus palabras.
            5. **Fomenta las ideas de mejora:** Escucha y canaliza las propuestas de los operadores para optimizar los procesos.
            """, unsafe_allow_html=True)

        with st.expander("🎯 Delegación y Control", expanded=True):
            st.markdown("""
            1. **Delega con base en la habilidad:** Asigna las tareas críticas al personal que ya demostró la competencia para resolverlas.
            2. **Sigue el rastro, no microgestiones:** Define puntos de revisión intermedios en lugar de vigilar cada paso del proceso.
            3. **Entrega recursos completos:** Al delegar, asegura que la persona tenga la herramienta, la información y el tiempo necesarios.
            4. **Asume la responsabilidad final:** Si tu equipo falla, tú eres el responsable ante la Dirección; no culpes a tus subordinados.
            5. **Estandariza los éxitos:** Cuando una solución funcione, documenta el nuevo método para que se convierta en la regla oficial.
            """, unsafe_allow_html=True)


### R E P O R T E S   P D F ####  

elif opcion_menu == "📋 Reportes PDF":
    st.subheader("🛠️ Generación de Reportes Ejecutivos")
    st.write("Selecciona los filtros requeridos para estructurar el reporte de actividades pendientes:")
    st.write("---")
    
    # 1. Configuración de Filtros Avanzados
    col_rep1, col_rep2, col_rep3 = st.columns(3)
    with col_rep1: 
        area_rep = st.selectbox("Filtrar por Área", ["Todas"] + st.session_state.areas, key="pdf_area")
    with col_rep2: 
        resp_rep = st.selectbox("Filtrar por Responsable", ["Todos"] + list(st.session_state.personal.keys()), key="pdf_resp")
    with col_rep3:
        rango_tiempo = st.selectbox(
            "Rango de Fecha Compromiso", 
            ["Cualquier Fecha Límite", "Esta Semana", "Próximas 2 Semanas", "Próximas 3 Semanas", "Próximas 4 Semanas"], 
            key="pdf_tiempo"
        )
        
    # 2. Motor de Filtrado de Datos Virtual
    df_rep = pd.DataFrame(st.session_state.actividades)
    if not df_rep.empty:
        if area_rep != "Todas": 
            df_rep = df_rep[df_rep["Area"] == area_rep]
        if resp_rep != "Todos": 
            df_rep = df_rep[df_rep["Responsable"] == resp_rep]
            
        df_pendientes_rep = df_rep[df_rep["% Avance"].astype(int) < 100].copy()
        
        if rango_tiempo != "Cualquier Fecha Límite" and not df_pendientes_rep.empty:
            hoy = datetime.now().date()
            if rango_tiempo == "Esta Semana": dias_limite = 7
            elif rango_tiempo == "Próximas 2 Semanas": dias_limite = 14
            elif rango_tiempo == "Próximas 3 Semanas": dias_limite = 21
            elif rango_tiempo == "Próximas 4 Semanas": dias_limite = 28
                
            fecha_maxima = hoy + timedelta(days=dias_limite)
            
            def evaluar_rango_fecha(fecha_str):
                try:
                    if not fecha_str or str(fecha_str).strip() in ["None", "nan", ""]: return False
                    f_dt = datetime.strptime(str(fecha_str).strip(), "%d-%b-%y").date()
                    return f_dt <= fecha_maxima
                except: return False
                    
            mascara_fechas = df_pendientes_rep["Fecha Compromiso"].apply(evaluar_rango_fecha)
            df_pendientes_rep = df_pendientes_rep[mascara_fechas]
        
        # 3. Vista Previa en Pantalla
        st.write("### 📊 Vista Previa de Actividades Pendientes")
        if not df_pendientes_rep.empty:
            st.metric("Total de Tareas Pendientes encontradas", len(df_pendientes_rep))
            st.dataframe(df_pendientes_rep[["No", "Responsable", "Area", "Descripcion", "% Avance", "Fecha Compromiso"]], use_container_width=True, hide_index=True)
            st.write("---")
            
            # --- MOTOR INTERNO DE IMPRESIÓN PDF ---
            def creador_documento_pdf(df_datos, area_txt, resp_txt, tiempo_txt):
                from fpdf import FPDF
                # Diseño en Horizontal (Landscape) para un formato tipo Matriz
                pdf = FPDF(orientation="L", unit="mm", format="A4")
                pdf.set_auto_page_break(auto=True, margin=15)
                pdf.add_page()
                
                # Encabezado principal (Azul marino corporativo #0C2340)
                pdf.set_font("Helvetica", "B", 16)
                pdf.set_text_color(12, 35, 64)
                pdf.cell(0, 10, txt="PLANTA METALES Y MAQUINADOS", ln=True, align="C")
                pdf.set_font("Helvetica", "B", 12)
                pdf.cell(0, 8, txt="REPORTE DE MATRIZ DE COMUNICACIÓN EFECTIVA", ln=True, align="C")
                
                # Subtítulo de filtros aplicados (Limpieza rápida de texto para evitar bugs en el título)
                area_txt_limpio = "".join(c for c in str(area_txt) if c.isalnum() or c.isspace()).strip()
                resp_txt_limpio = str(resp_txt).encode('latin-1', 'ignore').decode('latin-1')
                tiempo_txt_limpio = str(tiempo_txt).encode('latin-1', 'ignore').decode('latin-1')
                
                pdf.set_font("Helvetica", "I", 9)
                pdf.set_text_color(80, 80, 80)
                pdf.cell(0, 5, txt=f"Filtros: Area ({area_txt_limpio}) | Responsable ({resp_txt_limpio}) | Periodo ({tiempo_txt_limpio})", ln=True, align="C")
                pdf.cell(0, 5, txt=f"Generado el: {datetime.now().strftime('%d-%b-%y %H:%M')}", ln=True, align="C")
                pdf.ln(5)
                
                # Línea divisoria azul
                pdf.set_draw_color(12, 35, 64)
                pdf.line(10, pdf.get_y(), 285, pdf.get_y())
                pdf.ln(5)
                
                # Tabla: Encabezados (Gris #CFD8DC)
                pdf.set_font("Helvetica", "B", 10)
                pdf.set_fill_color(207, 216, 220)
                pdf.set_text_color(12, 35, 64)
                
                pdf.cell(15, 8, txt="No", border=1, fill=True, align="C")
                pdf.cell(35, 8, txt="Área", border=1, fill=True)
                pdf.cell(45, 8, txt="Responsable", border=1, fill=True)
                pdf.cell(125, 8, txt="Descripción del Pendiente", border=1, fill=True)
                pdf.cell(20, 8, txt="Avance", border=1, fill=True, align="C")
                pdf.cell(35, 8, txt="F. Límite", border=1, fill=True, ln=True)
                
                # Tabla: Filas de datos (Color de fondo amarillo tenue #FFF3CD para alertas)
                pdf.set_font("Helvetica", "", 9)
                pdf.set_text_color(0, 0, 0)
                
                for _, fila in df_datos.iterrows():
                    desc_corta = str(fila["Descripcion"])
                    if len(desc_corta) > 75:
                        desc_corta = desc_corta[:72] + "..."
                        
                    # 🔧 LIMPIEZA ANTIBUGS: Filtra solo letras, números y espacios para evitar el error de la fuente
                    area_limpia = "".join(c for c in str(fila["Area"]) if c.isalnum() or c.isspace()).strip()
                    resp_limpio = str(fila["Responsable"]).encode('latin-1', 'ignore').decode('latin-1')
                    desc_corta = desc_corta.encode('latin-1', 'ignore').decode('latin-1')
                        
                    pdf.set_fill_color(255, 243, 205)
                    
                    pdf.cell(15, 7, txt=str(fila["No"]), border=1, align="C")
                    pdf.cell(35, 7, txt=area_limpia[:18], border=1)
                    pdf.cell(45, 7, txt=resp_limpio[:22], border=1)
                    pdf.cell(125, 7, txt=desc_corta, border=1)
                    pdf.cell(20, 7, txt=f"{fila['% Avance']}%", border=1, fill=True, align="C")
                    pdf.cell(35, 7, txt=str(fila["Fecha Compromiso"]), border=1, ln=True)
                    
                return pdf.output()

            # --- BOTÓN DE INTERFAZ STREAMLIT ---
            try:
                pdf_datos_bytes = creador_documento_pdf(df_pendientes_rep, area_rep, resp_rep, rango_tiempo)
                st.download_button(
                    label="📄 Imprimir Reporte en PDF",
                    data=bytes(pdf_datos_bytes),
                    file_name=f"Reporte_MCE_{datetime.now().strftime('%d-%b-%y')}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
            except Exception as error_pdf:
                st.error(f"Error al generar el archivo de impresión: {error_pdf}")
        else:
            st.success("🎉 ¡Sin pendientes! No hay actividades retrasadas o en proceso en el rango de tiempo seleccionado.")
    else:
        st.info("La base de datos se encuentra vacía.")

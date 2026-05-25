
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as graph_objects
from datetime import datetime, timedelta
from PIL import Image
import os
import io
from github import Github, GithubException

# Nombre oficial del archivo base en la raíz de tu repositorio GitHub
ARCHIVO_DB = "base_matriz_mce.xlsx"
CARPETA_EVIDENCIAS = "evidencias"

if not os.path.exists(CARPETA_EVIDENCIAS):
    os.makedirs(CARPETA_EVIDENCIAS)

Image.MAX_IMAGE_PIXELS = None

# Configuración de la interfaz web corporativa (Fondo Gris Claro de Planta)
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
def conectar_github():
    try:
        token = st.secrets["GITHUB_TOKEN"]
        repo_name = st.secrets["GITHUB_REPO"]
        g = Github(token)
        return g.get_repo(repo_name)
    except Exception as e:
        st.error(f"Error de configuración con las credenciales de GitHub: {e}")
        return None

def importar_registros_excel():
    repo = conectar_github()
    columnas_por_defecto = ["No", "Origen", "Fecha Inicio", "Prioridad", "Responsable", "Area", "Descripcion", "% Avance", "Fecha Compromiso", "Comentario", "Evidencia"]
    
    if repo is not None:
        try:
            contenido_archivo = repo.get_contents(ARCHIVO_DB)
            datos_binarios = contenido_archivo.decoded_content
            df = pd.read_excel(io.BytesIO(datos_binarios))
            
            if df is not None and not df.empty:
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

                # Control estricto de porcentajes y consecutivos numéricos
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
                return df
        except GithubException as ge:
            if ge.status == 404:
                return pd.DataFrame(columns=columnas_por_defecto)
            else:
                st.error(f"Error de GitHub: {ge.data.get('message', ge)}")
        except Exception as e:
            st.error(f"Error al procesar el archivo Excel desde GitHub: {e}")
            
    return pd.DataFrame(columns=columnas_por_defecto)

def guardar_registros_excel(df_actualizado):
    repo = conectar_github()
    if repo is None:
        st.error("No se pudo guardar: Error de conexión con GitHub.")
        return False
        
    try:
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            df_actualizado.to_excel(writer, index=False)
        excel_binario = output_buffer.getvalue()
        
        try:
            contenido_archivo = repo.get_contents(ARCHIVO_DB)
            repo.update_file(
                path=ARCHIVO_DB,
                message=f"Sistema MCE: Actualización de matriz - {datetime.now().strftime('%d-%b-%y %H:%M')}",
                content=excel_binario,
                sha=contenido_archivo.sha,
                branch="main"
            )
        except GithubException as ge:
            if ge.status == 404:
                repo.create_file(
                    path=ARCHIVO_DB,
                    message="Sistema MCE: Creación inicial de la base de datos Excel",
                    content=excel_binario,
                    branch="main"
                )
            else:
                raise ge
        st.success("¡Base de datos sincronizada en tu repositorio de GitHub!")
        return True
    except Exception as e:
        st.error(f"Error crítico al subir los cambios a GitHub: {e}")
        return False

# Carga inicial de datos compartida en memoria
if 'actividades' not in st.session_state:
    st.session_state.actividades = importing_registros_excel() if 'importing_registros_excel' in globals() else  importar_registros_excel()
if 'personal' not in st.session_state:
    st.session_state.personal = [
        "Jesus Morales", "Ing. Alfredo Hdz", "Ing. Lorena Hdz", "Jesús Alday",
        "Alejandra Arellano", "Jose Romo", "Jose Manuel Aldama", "Fernando Llanas",
        "Celso", "Cruz Carreon", "Luis Quintana", "Bryan Flores",
        "Jorge Sanchez", "Voctor Montoya", "Moises Hernandez", "Rodolfo Ferndez M"
    ]

if 'areas' not in st.session_state:
    st.session_state.areas = ["⚙️ Ingenieria", "🔍 Calidad", "📦 Almacen", "✂️ Corte", "📐 Doblez", "🎨 Pintura", "🚚 Embarquez", "🏭 Planta Rio"]

LISTA_CLASIFICACIONES = ["Acuerdos", "Programa de Actividades", "Actividades Sujeridas", "Dirección", "Problema de Calidad", "Problema de Seguridad", "Lista de Pendientes", "Auto Asignado", "Plan de Control y Monitoreo", "Mejoras", "Investigación", "Manuales", "Procesos"]

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

# Encabezado Corporativo Fijo
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
if opcion_menu == "📊 Dashboard Principal":
    st.markdown("### Resumen Ejecutivo de la Planta")
    df_actual = st.session_state.actividades
    
    if not df_actual.empty:
        total = len(df_actual)
        terminadas = len(df_actual[df_actual["% Avance"] == 100])
        pendientes = total - terminadas
        
        m1, m2, m3 = st.columns(3)
        m1.metric("Total de Tareas", total)
        m2.metric("Terminadas ✅", terminadas, f"{(terminadas/total)*100:.1f}% del total")
        m3.metric("Pendientes ⚠️", pendientes, f"{(pendientes/total)*100:.1f}% del total", delta_color="inverse")
        
        col_g1, col_g2 = st.columns(2)
        with col_g1:
            st.plotly_chart(crear_grafico_pareto(df_actual, "Area", "Distribución de Carga por Área"), use_container_width=True)
        with col_g2:
            st.plotly_chart(crear_grafico_pareto(df_actual, "Responsable", "Tareas Pendientes vs Terminadas por Colaborador"), use_container_width=True)
    else:
        st.info("No hay actividades registradas en el repositorio. Registra una nueva tarea en los menús laterales.")
elif opcion_menu == "📋 Tabla de Control":
    st.markdown("### Repositorio Global de Actividades")
    if not st.session_state.actividades.empty:
        st.dataframe(st.session_state.actividades, use_container_width=True, hide_index=True)
    else:
        st.info("La base de datos se encuentra vacía.")
elif opcion_menu == "📝 Actualizar Mis Avances":
    st.markdown("### Modificar Porcentaje de Avance y Comentarios")
    df_actual = st.session_state.actividades
    
    if not df_actual.empty:
        id_tarea = st.selectbox("Seleccione el No. de Tarea que desea actualizar:", df_actual["No"].tolist())
        idx = df_actual[df_actual["No"] == id_tarea].index
        
        st.markdown(f"""
        <div class="card">
            <div class="card-header">Tarea No. {df_actual.at[idx, 'No'].values[0] if hasattr(df_actual.at[idx, 'No'], 'values') else df_actual.at[idx, 'No']} - {df_actual.at[idx, 'Area'].values[0] if hasattr(df_actual.at[idx, 'Area'], 'values') else df_actual.at[idx, 'Area']}</div>
            <div class="card-desc"><b>Responsable:</b> {df_actual.at[idx, 'Responsable'].values[0] if hasattr(df_actual.at[idx, 'Responsable'], 'values') else df_actual.at[idx, 'Responsable']} | <b>Prioridad:</b> {df_actual.at[idx, 'Prioridad'].values[0] if hasattr(df_actual.at[idx, 'Prioridad'], 'values') else df_actual.at[idx, 'Prioridad']}</div>
            <div class="card-desc"><b>Descripción:</b> {df_actual.at[idx, 'Descripcion'].values[0] if hasattr(df_actual.at[idx, 'Descripcion'], 'values') else df_actual.at[idx, 'Descripcion']}</div>
        </div>
        """, unsafe_allow_html=True)
        
        with st.form("form_actualizar_avance"):
            nuevo_avance = st.slider("Nuevo % de Avance", min_value=0, max_value=100, value=int(df_actual.at[idx, "% Avance"].values[0] if hasattr(df_actual.at[idx, '% Avance'], 'values') else df_actual.at[idx, '% Avance']), step=5)
            nuevo_comentario = st.text_area("Comentarios de Seguimiento", value=str(df_actual.at[idx, "Comentario"].values[0] if hasattr(df_actual.at[idx, 'Comentario'], 'values') else df_actual.at[idx, 'Comentario']))
            
            guardar_cambio = st.form_submit_button("Sincronizar Avance en GitHub")
            
            if guardar_cambio:
                df_actual.at[idx, "% Avance"] = nuevo_avance
                df_actual.at[idx, "Comentario"] = nuevo_comentario
                st.session_state.actividades = df_actual
                
                if guardar_registros_excel(st.session_state.actividades):
                    st.rerun()
    else:
        st.info("No hay tareas creadas para modificar.")
elif opcion_menu == "📥 Cargar Actividades (Usuario)":
    st.markdown("### Levantar Nueva Orden / Actividad")
    
    with st.form("formulario_alta"):
        col1, col2 = st.columns(2)
        with col1:
            origen = st.selectbox("Origen de la Actividad", LISTA_CLASIFICACIONES)
            prioridad = st.selectbox("Prioridad", ["Baja", "Media", "Alta", "Crítica"])
            responsable = st.selectbox("Responsable Asignado", st.session_state.personal)
        with col2:
            area = st.selectbox("Área Solicitante", st.session_state.areas)
            fecha_compromiso = st.date_input("Fecha Compromiso", value=datetime.now() + timedelta(days=7))
            
        descripcion = st.text_area("Descripción Detallada del Requerimiento / Acción")
        enviar_alta = st.form_submit_button("Dar de Alta e Inyectar en Repo")
        
        if enviar_alta:
            nuevo_no = int(st.session_state.actividades["No"].max() + 1) if not st.session_state.actividades.empty else 1
            
            nueva_fila = {
                "No": nuevo_no,
                "Origen": origen,
                "Fecha Inicio": datetime.now().strftime("%d-%b-%y"),
                "Prioridad": prioridad,
                "Responsable": responsable,
                "Area": area,
                "Descripcion": descripcion,
                "% Avance": 0,
                "Fecha Compromiso": fecha_compromiso.strftime("%d-%b-%y"),
                "Comentario": "",
                "Evidencia": ""
            }
            
            st.session_state.actividades = pd.concat([st.session_state.actividades, pd.DataFrame([nueva_fila])], ignore_index=True)
            
            if guardar_registros_excel(st.session_state.actividades):
                st.balloons()
                st.rerun()

elif opcion_menu == "🔐 Panel Administrador":
    st.markdown("### Configuración del Sistema Interno")
    st.warning("Zona restringida para edición y purga de la base de datos.")
    
    if st.button("Sincronizar / Forzar lectura manual desde GitHub"):
        st.session_state.actividades = importar_registros_excel()
        st.success("Sincronización forzada completada con éxito.")
        st.rerun()

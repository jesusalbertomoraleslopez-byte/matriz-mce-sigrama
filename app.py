import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as graph_objects
from datetime import datetime, timedelta
from PIL import Image
import os
import io
from github import Github

# --- CONFIGURACIÓN DE BASE DE DATOS Y REPOSITORIO ---
ARCHIVO_DB = "base_matriz_mce.xlsx"
CARPETA_EVIDENCIAS = "evidencias"
REPOSITORIO_NAME = "TU_USUARIO_GITHUB/TU_REPOSITORIO"  # <-- REEMPLAZA CON TUS DATOS REALES

# Desactivar límite de píxeles para capturas pesadas de taller
Image.MAX_IMAGE_PIXELS = None

# Inicializar conexión con el repositorio de GitHub de forma segura
repo = None
if "GITHUB_TOKEN" in st.secrets:
    try:
        g = Github(st.secrets["GITHUB_TOKEN"])
        repo = g.get_repo(REPOSITORIO_NAME)
    except Exception as e:
        st.error(f"Error crítico al conectar con el repositorio de GitHub: {e}")
else:
    st.warning("⚠️ GITHUB_TOKEN no detectado en Secrets. Los datos serán volátiles y temporales.")
def importar_registros_excel():
    # Estructura base en caso de fallo o archivo nuevo
    df_vacio = pd.DataFrame(columns=["No", "Origen", "Fecha Inicio", "Prioridad", "Responsable", "Area", "Descripcion", "% Avance", "Fecha Compromiso", "Comentario", "Evidencia"])
    
    # Intentar descargar el archivo desde GitHub
    if repo is not None:
        try:
            contenido_archivo = repo.get_contents(ARCHIVO_DB)
            datos_binarios = contenido_archivo.decoded_content
            df = pd.read_excel(io.BytesIO(datos_binarios))
        except Exception:
            return df_vacio
    elif os.path.exists(ARCHIVO_DB):  # Respaldo local si no hay red/token
        try:
            df = pd.read_excel(ARCHIVO_DB)
        except Exception:
            return df_vacio
    else:
        return df_vacio

    # --- PROCESAMIENTO Y LIMPIEZA DE DATOS ---
    if not df.empty:
        # Corrección de fechas seriales de Excel a formato legible
        for col_fecha in ["Fecha Inicio", "Fecha Compromiso"]:
            if col_fecha in df.columns:
                def corregir_fecha_serial(val):
                    try:
                        if pd.isna(val) or str(val).strip() in ["None", "nan", "NaN", ""]:
                            return ""
                        if str(val).replace('.0', '').isdigit():
                            dias = int(str(val).replace('.0', ''))
                            return (datetime(1899, 12, 30) + timedelta(days=dias)).strftime("%d-%b-%y")
                        return str(val).strip()
                    except:
                        return str(val)
                df[col_fecha] = df[col_fecha].apply(corregir_fecha_serial)

        # Control estricto de porcentajes de avance
        if "% Avance" in df.columns:
            if df["% Avance"].max() <= 1.0 and df["% Avance"].max() > 0:
                df["% Avance"] = df["% Avance"] * 100
            df["% Avance"] = pd.to_numeric(df["% Avance"], errors="coerce").fillna(0).astype(int)
        
        # Consecutivos numéricos estrictos
        if "No" in df.columns:
            df["No"] = pd.to_numeric(df["No"], errors="coerce")
            if df["No"].isnull().any(): 
                df["No"] = range(1, len(df) + 1)
            df["No"] = df["No"].astype(int)
        
        # Limpieza de textos nulos
        columnas_texto = ["Origen", "Prioridad", "Responsable", "Area", "Descripcion", "Comentario", "Evidencia"]
        for col in columnas_texto:
            if col in df.columns:
                df[col] = df[col].astype(str).replace(["None", "nan", "NaN"], "")
    else:
        df = df_vacio.copy()
        
    return df

# Carga inicial de datos en la memoria de la sesión de Streamlit
if 'actividades' not in st.session_state:
    st.session_state.actividades = importar_registros_excel()
def guardar_registros_excel(df_actualizado):
    try:
        # Empaquetar el Excel en memoria RAM antes de enviarlo
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_actualizado.to_excel(writer, index=False)
        datos_excel = output.getvalue()
        
        if repo is not None:
            try:
                # Si el archivo ya existe en GitHub, se actualiza mediante su SHA
                contents = repo.get_contents(ARCHIVO_DB)
                repo.update_file(path=ARCHIVO_DB, message="Actualización de Matriz MCE", content=datos_excel, sha=contents.sha, branch="main")
            except Exception:
                # Si el archivo fue borrado del repositorio, lo vuelve a crear de cero
                repo.create_file(path=ARCHIVO_DB, message="Creación inicial de Matriz MCE", content=datos_excel, branch="main")
            st.success("💾 ¡Cambios sincronizados y respaldados en el Repositorio de GitHub!")
        else:
            df_actualizado.to_excel(ARCHIVO_DB, index=False)
            st.warning("⚠️ Guardado local temporal (No sincronizado con la nube).")
    except Exception as e:
        st.error(f"Error crítico al guardar los registros: {e}")

def guardar_evidencia_en_github(archivo_subido, nombre_final_imagen):
    if repo is None:
        # Fallback local de emergencia
        if not os.path.exists(CARPETA_EVIDENCIAS): os.makedirs(CARPETA_EVIDENCIAS)
        with open(os.path.join(CARPETA_EVIDENCIAS, nombre_final_imagen), "wb") as f: f.write(archivo_subido.getvalue())
        return f"{CARPETA_EVIDENCIAS}/{nombre_final_imagen}"

    try:
        ruta_completa = f"{CARPETA_EVIDENCIAS}/{nombre_final_imagen}"
        contenido_imagen = archivo_subido.getvalue()
        try:
            # Reemplaza la imagen existente si coincide el nombre
            contents = repo.get_contents(ruta_completa)
            repo.update_file(ruta_completa, f"Actualizar evidencia: {nombre_final_imagen}", contenido_imagen, contents.sha, branch="main")
        except Exception:
            # Sube un nuevo archivo físico a la carpeta evidencias de tu repo
            repo.create_file(ruta_completa, f"Subir nueva evidencia: {nombre_final_imagen}", contenido_imagen, branch="main")
        return ruta_completa
    except Exception as e:
        st.error(f"Error al subir el archivo de imagen a GitHub: {e}")
        return ""
# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="SIGRAMA - Matriz MCE", layout="wide")

# Inyección de estilos CSS Corporativos
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

# --- INICIALIZACIÓN DE CATÁLOGOS EN SESSION STATE ---
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

# --- RENDERIZADO DEL ENCABEZADO CORPORATIVO ---
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
# Menú de control lateral principal
opcion_menu = st.sidebar.radio("Navegación", ["📊 Dashboard Principal", "📋 Tabla de Control", "📝 Actualizar Mis Avances", "📥 Cargar Actividades (Usuario)", "🔐 Panel Administrador"])

# --- RENDERIZADO DE VISTAS SEGÚN SELECCIÓN ---
if opcion_menu == "📊 Dashboard Principal":
    st.subheader("Cuadro de Mando Operativo")
    col_f1, col_f2 = st.columns(2)
    with col_f1: 
        area_sel = st.selectbox("Filtrar por Área Operativa", ["Todas"] + st.session_state.areas)
    
    # Aquí continuará la lógica específica de graficación y filtrado de tu pestaña 1...

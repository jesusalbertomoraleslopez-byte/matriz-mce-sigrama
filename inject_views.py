import os

app_path = "app.py"
with open(app_path, "r", encoding="utf-8") as f:
    content = f.read()

# 1. Define the new functions
new_functions = """
def renderizar_formulario_tarea(r):
    col_izq, col_der = st.columns(2)
    
    with col_izq:
        progreso_actual = int(r['% Avance'])
        
        color_progreso = "#2ECC71" if progreso_actual == 100 else "#EC2024"
        st.markdown(f'''
        <div style="font-family:'Montserrat', sans-serif; font-size:13px; font-weight:700; color:#111111; margin-bottom:4px;">
            PROGRESO: {progreso_actual}%
        </div>
        <div style="width:100%; background-color:#D2D3D5; border-radius:4px; height:18px; overflow:hidden; margin-bottom:12px; border:1px solid #D2D3D5;">
            <div style="width:{progreso_actual}%; background-color:{color_progreso}; height:100%; transition:width 0.4s ease-in-out;"></div>
        </div>
        ''', unsafe_allow_html=True)
        
        nv_av = st.slider("Ajustar %:", min_value=0, max_value=100, value=progreso_actual, step=5, key=f"num_{r['No']}")
    
    with col_der:
        comentario_limpio = "" if str(r['Comentario']).strip().lower() in ["nan", "none", ""] else str(r['Comentario'])
        nv_co = st.text_input("Comentarios de bitácora:", value=comentario_limpio, key=f"c_{r['No']}")
        evidencia_guardada = str(r['Evidencia']).strip()
        if evidencia_guardada and os.path.exists(evidencia_guardada):
            st.image(Image.open(evidencia_guardada), use_container_width=True, caption="📸 Evidencia Actual")
        
        foto = st.file_uploader("Evidencia Fotográfica (Cierre 100%):", type=["jpg","png","jpeg","jfif"], key=f"i_{r['No']}") if nv_av == 100 else None
        if foto: 
            st.image(Image.open(foto), use_container_width=True, caption="Vista Previa")
        
        st.markdown('<div style="margin-top: 10px;">', unsafe_allow_html=True)
        if nv_av == 100:
            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                guardar_clicked = st.button("Guardar Tarea", key=f"b_{r['No']}", use_container_width=True)
            with col_btn2:
                pasted_b64 = paste_clipboard(key=f"paste_{r['No']}")
        else:
            guardar_clicked = st.button("Guardar Tarea", key=f"b_{r['No']}", use_container_width=True)
            pasted_b64 = None
        
        if guardar_clicked:
            if nv_av == 100 and not foto and not pasted_b64 and not evidencia_guardada: 
                st.error("⚠️ Faltan las fotografías físicas obligatorias para autorizar el cierre al 100% (Subir archivo o pegar desde portapapeles).")
            else:
                ruta_foto_final = evidencia_guardada
                nombre_archivo_final = None
                
                if foto is not None:
                    try:
                        img_abierta = Image.open(foto)
                        if img_abierta.mode in ("RGBA", "P"): 
                            img_abierta = img_abierta.convert("RGB")
                        img_abierta.thumbnail((800, 600), Image.Resampling.LANCZOS)
                        
                        fecha_cierre = datetime.now().strftime("%Y%m%d")
                        desc_sanitizada = normalizar_texto(r['Descripcion']).replace(' ', '_')
                        import re
                        desc_sanitizada = re.sub(r'[^a-zA-Z0-9_]', '', desc_sanitizada)[:50]
                        
                        nombre_archivo_final = f"SCA-{int(r['No']):03d}_{desc_sanitizada}_{fecha_cierre}.jpg"
                        ruta_foto_final = os.path.join(CARPETA_EVIDENCIAS, nombre_archivo_final)
                        img_abierta.save(ruta_foto_final, "JPEG", quality=65)
                    except Exception as err_img: 
                        st.error(f"Fallo al procesar captura: {err_img}")
                        
                elif pasted_b64 is not None and pasted_b64.startswith("data:image"):
                    try:
                        import base64
                        format_part, base64_data = pasted_b64.split(',', 1)
                        ext = "png"
                        if "jpeg" in format_part or "jpg" in format_part:
                            ext = "jpg"
                            
                        img_bytes = base64.b64decode(base64_data)
                        
                        fecha_cierre = datetime.now().strftime("%Y%m%d")
                        desc_sanitizada = normalizar_texto(r['Descripcion']).replace(' ', '_')
                        import re
                        desc_sanitizada = re.sub(r'[^a-zA-Z0-9_]', '', desc_sanitizada)[:50]
                        
                        nombre_archivo_final = f"SCA-{int(r['No']):03d}_{desc_sanitizada}_{fecha_cierre}.{ext}"
                        ruta_foto_final = os.path.join(CARPETA_EVIDENCIAS, nombre_archivo_final)
                        
                        with open(ruta_foto_final, "wb") as f_img:
                            f_img.write(img_bytes)
                            
                        try:
                            img_abierta = Image.open(ruta_foto_final)
                            if img_abierta.mode in ("RGBA", "P") and ext == "jpg":
                                img_abierta = img_abierta.convert("RGB")
                            img_abierta.thumbnail((800, 600), Image.Resampling.LANCZOS)
                            img_abierta.save(ruta_foto_final, "JPEG" if ext == "jpg" else "PNG", quality=75)
                        except Exception as err_comp:
                            pass
                    except Exception as err_paste:
                        st.error(f"Fallo al decodificar imagen pegada: {err_paste}")
                        
                if nombre_archivo_final and os.path.exists(ruta_foto_final):
                    mensaje_commit = f"Respaldo evidencia SCA-{int(r['No']):03d} ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')})"
                    exito_git, msg_git = subir_archivo_github(ruta_foto_final, f"evidencias/{nombre_archivo_final}", mensaje_commit)
                    if exito_git:
                        st.toast(f"✅ Evidencia respaldada en GitHub: {nombre_archivo_final}")
                    else:
                        st.warning(f"⚠️ Alerta: Evidencia guardada localmente pero falló el respaldo en GitHub: {msg_git}")
                
                df_disco = importar_registros_excel()
                if not df_disco.empty:
                    row_no = int(r["No"])
                    idx_disco = df_disco[df_disco["No"] == row_no].index
                    if len(idx_disco) > 0:
                        df_disco.loc[idx_disco[0], "% Avance"] = int(nv_av)
                        df_disco.loc[idx_disco[0], "Comentario"] = str(nv_co)
                        df_disco.loc[idx_disco[0], "Evidencia"] = str(ruta_foto_final)
                        
                        try:
                            df_disco.to_excel(ARCHIVO_DB, index=False)
                            subir_archivo_github(ARCHIVO_DB, "base_matriz_mce.xlsx", f"Auto-backup Excel (SCA-{row_no:03d})")
                            st.session_state.actividades = df_disco
                            st.success("🏁 ¡Cambios registrados con éxito!")
                            st.rerun()
                        except Exception as e_save:
                            st.error(f"Fallo de escritura en base física Excel: {e_save}")
                    else:
                        st.error("No se encontró la tarea en la base de datos física.")
        st.markdown('</div>', unsafe_allow_html=True)

@st.dialog("Editar Tarea", width="large")
def dialogo_editar_tarea(r):
    renderizar_formulario_tarea(r)

if 'actividades' not in st.session_state:
"""
content = content.replace("if 'actividades' not in st.session_state:", new_functions)


# 2. Replace the UI block for 'Actualizar Mis Avances'
old_ui_block_start = """            if df_filtrado.empty:
                st.info(f"📋 No hay tareas en la clasificación '{clasificacion}' para {u}.")
            else:
                for idx in df_filtrado.index:"""

new_ui_block = """            tipo_vista = st.radio("Tipo de Vista:", ["🔲 Vista Cuadrícula (Miniaturas)", "📄 Vista Detalle (Lista)"], horizontal=True)
            st.write("---")
            
            if df_filtrado.empty:
                st.info(f"📋 No hay tareas en la clasificación '{clasificacion}' para {u}.")
            else:
                if tipo_vista == "📄 Vista Detalle (Lista)":
                    for idx in df_filtrado.index:
                        r = st.session_state.actividades.loc[idx]
                        
                        with st.container():
                            st.markdown(f'''
                            <div class="task-card">
                                <p style="font-family:'Montserrat', sans-serif; font-size:18px; font-weight:bold; color:#111111; margin:0 0 6px 0;">
                                    📋 Actividad No. {r["No"]} | {r["Area"]} | Prioridad: <span style="color:#EC2024; font-weight:700;">{r["Prioridad"]}</span>
                                </p>
                                <p style="font-family:'Questrial', sans-serif; font-size:15px; color:#333333; margin:0 0 10px 0;">
                                    <b>Descripción:</b> {r["Descripcion"]}
                                </p>
                            </div>
                            ''', unsafe_allow_html=True)
                            
                            renderizar_formulario_tarea(r)
                            st.write("---")
                else:
                    cols = st.columns(3)
                    for i, idx in enumerate(df_filtrado.index):
                        r = st.session_state.actividades.loc[idx]
                        col = cols[i % 3]
                        with col:
                            with st.container(border=True):
                                progreso_actual = int(r['% Avance'])
                                color_progreso = "#2ECC71" if progreso_actual == 100 else "#EC2024"
                                prioridad_color = "#EC2024" if str(r["Prioridad"]).lower() == "urgente" else "#F59E0B" if str(r["Prioridad"]).lower() == "media" else "#10B981"
                                
                                st.markdown(f'''
                                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
                                    <span style="font-family:'Montserrat', sans-serif; font-weight:bold; font-size:16px; color:#111111;">📋 No. {r["No"]}</span>
                                    <span style="font-size:11px; font-weight:bold; padding:2px 8px; border-radius:12px; background-color:{prioridad_color}20; color:{prioridad_color};">{r["Prioridad"]}</span>
                                </div>
                                <p style="font-family:'Questrial', sans-serif; font-size:13px; color:#475569; margin:0 0 10px 0; height:40px; overflow:hidden; text-overflow:ellipsis;">
                                    {str(r["Descripcion"])[:70]}...
                                </p>
                                <div style="width:100%; background-color:#E2E8F0; border-radius:4px; height:8px; overflow:hidden; margin-bottom:10px;">
                                    <div style="width:{progreso_actual}%; background-color:{color_progreso}; height:100%;"></div>
                                </div>
                                ''', unsafe_allow_html=True)
                                
                                evidencia_guardada = str(r['Evidencia']).strip()
                                if progreso_actual == 100 and evidencia_guardada and os.path.exists(evidencia_guardada):
                                    try:
                                        st.image(Image.open(evidencia_guardada), use_container_width=True)
                                    except:
                                        pass
                                elif progreso_actual == 100:
                                    st.markdown("<div style='text-align:center; font-size:12px; color:#94A3B8; margin-bottom:10px; border:1px dashed #94A3B8; padding:20px;'>📷 Sin evidencia</div>", unsafe_allow_html=True)
                                
                                if st.button("📝 Editar / Ver Detalle", key=f"btn_edit_{r['No']}", use_container_width=True):
                                    dialogo_editar_tarea(r)"""

start_idx = content.find(old_ui_block_start)
if start_idx != -1:
    end_str = "st.markdown('</div>', unsafe_allow_html=True)"
    end_idx = content.find(end_str, start_idx) + len(end_str)
    
    if end_idx > start_idx + len(end_str):
        content = content[:start_idx] + new_ui_block + content[end_idx:]

with open(app_path, "w", encoding="utf-8") as f:
    f.write(content)
print("Done")

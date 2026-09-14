import os

app_path = "app.py"
with open(app_path, "r", encoding="utf-8") as f:
    content = f.read()

# 1. Add `dialogo_ver_imagen` after `dialogo_editar_tarea`
new_dialogs = """@st.dialog("Editar Tarea", width="large")
def dialogo_editar_tarea(r):
    renderizar_formulario_tarea(r)

@st.dialog("🖼️ Evidencia Fotográfica", width="large")
def dialogo_ver_imagen(ruta_evidencia, num_tarea):
    st.markdown(f"<h3 style='text-align: center; color: #111111;'>Actividad No. {num_tarea}</h3>", unsafe_allow_html=True)
    st.image(Image.open(ruta_evidencia), use_container_width=True)"""

content = content.replace("""@st.dialog("Editar Tarea", width="large")
def dialogo_editar_tarea(r):
    renderizar_formulario_tarea(r)""", new_dialogs)

# 2. Replace List view block
old_list_view_start = """                if tipo_vista == "📄 Vista Detalle (Lista)":
                    for idx in df_filtrado.index:"""

new_list_view = """                if tipo_vista == "📄 Vista Detalle (Lista)":
                    st.markdown('''
                    <style>
                    .tbl-hdr { font-family:'Montserrat', sans-serif; font-size:14px; font-weight:700; color:#111; padding-bottom:8px; border-bottom:2px solid #EC2024; margin-bottom:10px; }
                    .tbl-row { font-family:'Questrial', sans-serif; font-size:13px; color:#333; display:flex; align-items:center; }
                    </style>
                    ''', unsafe_allow_html=True)
                    
                    header = st.columns([0.8, 4, 1.5, 2, 2.5])
                    header[0].markdown("<div class='tbl-hdr'>No.</div>", unsafe_allow_html=True)
                    header[1].markdown("<div class='tbl-hdr'>Descripción</div>", unsafe_allow_html=True)
                    header[2].markdown("<div class='tbl-hdr'>Prioridad</div>", unsafe_allow_html=True)
                    header[3].markdown("<div class='tbl-hdr'>Avance</div>", unsafe_allow_html=True)
                    header[4].markdown("<div class='tbl-hdr'>Acciones</div>", unsafe_allow_html=True)
                    
                    for idx in df_filtrado.index:
                        r = st.session_state.actividades.loc[idx]
                        row = st.columns([0.8, 4, 1.5, 2, 2.5], vertical_alignment="center")
                        row[0].markdown(f"<div class='tbl-row'><b>{r['No']}</b></div>", unsafe_allow_html=True)
                        row[1].markdown(f"<div class='tbl-row'>{str(r['Descripcion'])[:90]}...</div>", unsafe_allow_html=True)
                        
                        prioridad_color = "#EC2024" if str(r["Prioridad"]).lower() == "urgente" else "#F59E0B" if str(r["Prioridad"]).lower() == "media" else "#10B981"
                        row[2].markdown(f"<div class='tbl-row'><span style='color:{prioridad_color}; font-weight:bold;'>{r['Prioridad']}</span></div>", unsafe_allow_html=True)
                        
                        progreso_actual = int(r['% Avance'])
                        color_progreso = "#2ECC71" if progreso_actual == 100 else "#EC2024"
                        row[3].markdown(f'''
                        <div style="margin-top:2px; width:100%; background-color:#E2E8F0; border-radius:4px; height:10px; overflow:hidden;">
                            <div style="width:{progreso_actual}%; background-color:{color_progreso}; height:100%;"></div>
                        </div>
                        <div style="font-size:10px; text-align:right; font-weight:bold; color:#111;">{progreso_actual}%</div>
                        ''', unsafe_allow_html=True)
                        
                        with row[4]:
                            ca, cb = st.columns(2)
                            with ca:
                                if st.button("📝 Editar", key=f"btn_lst_ed_{r['No']}", use_container_width=True):
                                    dialogo_editar_tarea(r)
                            with cb:
                                ev_path = str(r.get("Evidencia", "")).strip()
                                if progreso_actual == 100 and ev_path and os.path.exists(ev_path):
                                    if st.button("🖼️ Foto", key=f"btn_lst_img_{r['No']}", use_container_width=True):
                                        dialogo_ver_imagen(ev_path, r['No'])
                        st.write("---")"""

start_idx = content.find(old_list_view_start)
if start_idx != -1:
    end_str = "st.write(\"---\")"
    end_idx = content.find(end_str, start_idx) + len(end_str)
    
    if end_idx > start_idx + len(end_str):
        content = content[:start_idx] + new_list_view + content[end_idx:]
else:
    print("Could not find start block")

with open(app_path, "w", encoding="utf-8") as f:
    f.write(content)
print("Done")

import os

app_path = "app.py"
with open(app_path, "r", encoding="utf-8") as f:
    content = f.read()

# 1. Remove the title from the top
top_title = """        else:
            u = st.selectbox("Seleccionar Responsable:", list(st.session_state.personal.keys()))
            
        st.markdown(f'<h1 style="color: #EC2024; font-family: \\'Montserrat\\', sans-serif; font-weight: 900; text-transform: uppercase; font-size: 42px; margin-top: 10px; margin-bottom: 25px; border-bottom: 2px solid #E2E8F0; padding-bottom: 15px;">👨‍🔧 {u}</h1>', unsafe_allow_html=True)"""

top_replacement = """        else:
            u = st.selectbox("Seleccionar Responsable:", list(st.session_state.personal.keys()))"""

content = content.replace(top_title, top_replacement)


# 2. Add the title at the bottom, just above the tasks
bottom_target = """            tipo_vista = st.radio("Tipo de Vista:", ["🔲 Vista Cuadrícula (Miniaturas)", "📄 Vista Detalle (Lista)"], horizontal=True)
            st.write("---")"""

bottom_replacement = """            tipo_vista = st.radio("Tipo de Vista:", ["🔲 Vista Cuadrícula (Miniaturas)", "📄 Vista Detalle (Lista)"], horizontal=True)
            st.write("---")
            st.markdown(f'<h1 style="color: #EC2024; font-family: \\'Montserrat\\', sans-serif; font-weight: 900; text-transform: uppercase; font-size: 38px; margin-top: 10px; margin-bottom: 25px; border-bottom: 2px solid #E2E8F0; padding-bottom: 15px;">👨‍🔧 Actividades de: {u}</h1>', unsafe_allow_html=True)"""

content = content.replace(bottom_target, bottom_replacement)

with open(app_path, "w", encoding="utf-8") as f:
    f.write(content)
print("Done")

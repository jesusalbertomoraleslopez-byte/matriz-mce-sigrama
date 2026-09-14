import re

app_path = "app.py"
with open(app_path, "r", encoding="utf-8") as f:
    content = f.read()

# Replace 1: app.py:212
# st.image(Image.open(evidencia_guardada), use_container_width=True, caption="?? Evidencia Actual")
content = content.replace(
    'st.image(Image.open(evidencia_guardada), use_container_width=True, caption="📸 Evidencia Actual")',
    'st.image(evidencia_guardada, width="stretch", caption="📸 Evidencia Actual")'
)

# Replace 2: app.py:216
# st.image(Image.open(foto), use_container_width=True, caption="Vista Previa")
content = content.replace(
    'st.image(Image.open(foto), use_container_width=True, caption="Vista Previa")',
    'st.image(foto, width="stretch", caption="Vista Previa")'
)

# Replace 3: app.py:322
# st.image(Image.open(ruta_evidencia), use_container_width=True)
content = content.replace(
    'st.image(Image.open(ruta_evidencia), use_container_width=True)',
    'st.image(ruta_evidencia, width="stretch")'
)

# Replace 4: app.py:511-512
# if os.path.exists(nombre_banner):
#     imagen_banner = Image.open(nombre_banner)
#     st.image(imagen_banner, use_container_width=True)
old_banner_block = """if os.path.exists(nombre_banner):
    imagen_banner = Image.open(nombre_banner)
    st.image(imagen_banner, use_container_width=True)"""

new_banner_block = """if os.path.exists(nombre_banner):
    st.image(nombre_banner, width="stretch")"""

content = content.replace(old_banner_block, new_banner_block)

# Replace 5: app.py:521
# st.sidebar.image("LOGOTIPO COLOR (1).jfif", use_container_width=True)
content = content.replace(
    'st.sidebar.image("LOGOTIPO COLOR (1).jfif", use_container_width=True)',
    'st.sidebar.image("LOGOTIPO COLOR (1).jfif", width="stretch")'
)

# Replace 6: app.py:601
# st.sidebar.image("LOGOTIPO COLOR (1).jfif", use_container_width=True)
# This is handled by replacing all occurrences of that string, but let's be safe.

# Replace 7: app.py:931
# st.image(Image.open(evidencia_guardada), use_container_width=True)
content = content.replace(
    'st.image(Image.open(evidencia_guardada), use_container_width=True)',
    'st.image(evidencia_guardada, width="stretch")'
)

with open(app_path, "w", encoding="utf-8") as f:
    f.write(content)

print("Optimization complete.")

# 🔁 INSTALACIÓN DE DEPENDENCIAS
!apt-get update > /dev/null
!apt-get install -y pandoc latexml > /dev/null

# 📁 CARGA DE ARCHIVOS
from google.colab import files
import os, re, shutil

print("📥 Subí tu archivo .tex, imágenes (.png, .jpg, .jpeg) y opcionalmente el .bib:")
uploaded = files.upload()

# 🧠 DETECTAR ARCHIVO .tex
TEX_FILE = next((f for f in uploaded if f.endswith(".tex")), None)
if TEX_FILE is None:
    raise ValueError("❌ No se subió ningún archivo .tex.")

BASE_NAME = TEX_FILE.replace(".tex", "")
DOCX_FILE = f"{BASE_NAME}.docx"

# ✨ PREPROCESAMIENTO: limpiar comandos problemáticos
with open(TEX_FILE, 'r', encoding='utf-8') as f:
    tex_content = f.read()

# 🔄 Reemplazo de comandos problemáticos comunes
tex_content = re.sub(r'\\SI\{([\d]+),([\d]+)\}\{([^}]*)\}', r'\1.\2 \3', tex_content)
tex_content = re.sub(r'\\num\{([\d]+),([\d]+)\}', r'\1.\2', tex_content)
tex_content = re.sub(r'\\boxed\{([^}]*)\}', r'\1', tex_content)
tex_content = re.sub(r'\$\$(.*?)\$\$', r'\\[\1\\]', tex_content, flags=re.DOTALL)
tex_content = re.sub(r'\\usepackage\[spanish\]\{babel\}', r'\\usepackage{babel}', tex_content)
tex_content = re.sub(r'\\usepackage\{tikz\}', '', tex_content)
tex_content = re.sub(r'\\begin\{tikzpicture\}.*?\\end\{tikzpicture\}', '', tex_content, flags=re.DOTALL)
tex_content = re.sub(r'\\bibliography\{[^}]*\}', '', tex_content)
tex_content = re.sub(r'\\cite\{[^}]*\}', '', tex_content)

# 💾 Guardar archivo limpio
CLEAN_FILE = f"{BASE_NAME}_limpio.tex"
with open(CLEAN_FILE, 'w', encoding='utf-8') as f:
    f.write(tex_content)

# 📂 CREAR CARPETA PARA MEDIA (imágenes)
media_dir = "media"
if os.path.exists(media_dir):
    shutil.rmtree(media_dir)
os.makedirs(media_dir)

# Mover imágenes subidas a carpeta 'media'
for fname in uploaded:
    if fname.lower().endswith((".png", ".jpg", ".jpeg")):
        shutil.move(fname, os.path.join(media_dir, fname))

# 🔁 Intentar CONVERSIÓN con LaTeXML + Pandoc
print("\n🔁 Intentando conversión con LaTeXML...")
exit_code = os.system(f'latexml --includestyles "{CLEAN_FILE}" --dest="{BASE_NAME}.xml"')

if exit_code == 0:
    print("🔁 Postprocesando a HTML...")
    exit_code2 = os.system(f'latexmlpost "{BASE_NAME}.xml" --dest="{BASE_NAME}.html"')

    if exit_code2 == 0:
        print("📄 Convirtiendo a Word con Pandoc...")
        os.system(f'pandoc "{BASE_NAME}.html" -o "{DOCX_FILE}" --extract-media=media')
    else:
        print("⚠️ LaTeXMLpost falló. Usando Pandoc directo...")
        os.system(f'pandoc "{CLEAN_FILE}" -o "{DOCX_FILE}" --extract-media=media')

else:
    print("⚠️ LaTeXML falló. Usando Pandoc directo...")
    os.system(f'pandoc "{CLEAN_FILE}" -o "{DOCX_FILE}" --extract-media=media')

# 📥 DESCARGA DEL DOCUMENTO
if os.path.exists(DOCX_FILE):
    print("\n✅ ¡Conversión lista! Descargando .docx...")
    files.download(DOCX_FILE)
else:
    print("❌ Error: el archivo .docx no se creó.")

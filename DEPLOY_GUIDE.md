# Guía de Despliegue Gratuito en la Nube (Streamlit Cloud)

Esta guía explica cómo publicar la aplicación web de forma **100% gratuita**, sin necesidad de tener dominio propio ni pagar servidores, para que la profesora **Araceli** pueda entrar desde cualquier computadora o celular y generar sus presentaciones.

---

## 1. Probar la aplicación en tu computadora primero

Antes de subirla a la nube, puedes verla funcionando en tu máquina:

1. Ve a la carpeta `Innova-Ara`.
2. Haz doble clic en el archivo:
   👉 **`Iniciar_App_Innova.bat`**
3. Se abrirá automáticamente una ventana en tu navegador web (`http://localhost:8501`) con la interfaz lista.
4. Puedes arrastrar cualquier archivo `.docx` (como `A1_Unit5_Session2_Instructional_Design.docx`) y presionar **"Generar Presentación PPTX Completa"** para probar la descarga directa.

---

## 2. Publicarla en la nube (Paso a paso para que Araceli la use)

### Paso A: Subir el proyecto a GitHub (Gratis)
1. Entra a [github.com](https://github.com) (crea una cuenta gratuita si no tienes una).
2. Crea un nuevo repositorio (por ejemplo: `innova-sesiones-araceli`).
   * *Nota: Puedes marcarlo como Privado o Público.*
3. Sube los siguientes archivos y carpetas de `Innova-Ara`:
   * `app.py`
   * `session_engine.py`
   * `requirements.txt`
   * `.streamlit/` (carpeta con `config.toml`)
   * `assets/` (carpeta con las imágenes, logotipos y plantilla)

### Paso B: Conectar con Streamlit Community Cloud (Gratis)
1. Entra a [streamlit.io/cloud](https://streamlit.io/cloud) e inicia sesión con tu cuenta de GitHub o Google.
2. Haz clic en el botón azul: **"Create app"** (o "New app").
3. Selecciona tu repositorio: `innova-sesiones-araceli`.
4. En **Main file path**, escribe: `app.py`.
5. En **App URL** (opcional), puedes elegir un nombre bonito como:
   `araceli-innova.streamlit.app` o `innova-sesiones.streamlit.app`.
6. Haz clic en **"Deploy!"**.

---

## 3. ¿Cómo la usará Araceli?

¡Listo! En menos de 2 minutos, la aplicación estará en vivo en internet.
Solo tienes que enviarle el enlace por WhatsApp o correo:
🔗 `https://araceli-innova.streamlit.app`

Ella verá una pantalla limpia con el logo de Innova Schools:
1. **Arrastra su archivo Word (`.docx`)**.
2. **Revisa los datos detectados**.
3. **Hace clic en "Generar Presentación PPTX"**.
4. **Descarga su presentación `.pptx` completa** con las 16 diapositivas, imágenes fotorrealistas, temporizadores y rúbrica oficial.

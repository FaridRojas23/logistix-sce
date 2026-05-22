# LOGISTIX AI — Dashboard combustible (SCE)

Dashboard Streamlit que lee el Excel en SharePoint (hojas **BdMes** y **Registro**).

## Ejecutar en local

```bash
cd dashboard
pip install -r requirements.txt
streamlit run STREAMFINAL.py
```

O doble clic en `STREAMLIT.bat`.

## Subir a GitHub

1. Crea un repositorio vacío en GitHub (ej. `logistix-dashboard`).
2. En esta carpeta:

```bash
git init
git add .
git commit -m "Dashboard Streamlit LOGISTIX para Render"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/logistix-dashboard.git
git push -u origin main
```

## Desplegar en Render

1. [render.com](https://render.com) → **New** → **Web Service**.
2. Conecta el repositorio de GitHub.
3. Configuración:
   - **Runtime:** Python 3
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:**
     ```bash
     streamlit run STREAMFINAL.py --server.port=$PORT --server.address=0.0.0.0 --server.headless=true
     ```
4. **Environment** → agrega variable:
   - `EXCEL_URL` = enlace de descarga del Excel en SharePoint (el mismo que usa la app).
5. Deploy. La URL pública será tipo `https://logistix-dashboard.onrender.com`.

Los datos se actualizan desde SharePoint cada **5 minutos** (caché) o con **Actualizar Datos (F5)** en la app.

## Notas

- El enlace de SharePoint debe permitir descarga **sin login** desde internet (enlace compartido).
- No subas `secrets.toml` ni archivos `.xlsx` al repositorio.

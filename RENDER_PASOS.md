# Render — qué poner en cada campo

## Start Command (copiar tal cual)

```bash
streamlit run STREAMFINAL.py --server.port=$PORT --server.address=0.0.0.0 --server.headless=true --server.fileWatcherType=none
```

## Build Command

```bash
pip install -r requirements.txt
```

## Health Check Path (en Settings del servicio)

```
/_stcore/health
```

Si no existe ese campo, deja `/` o vacío.

## Environment Variable

| Key | Value |
|-----|--------|
| EXCEL_URL | (enlace SharePoint del Excel) |

## Root Directory

Vacío si en GitHub ves `STREAMFINAL.py` en la raíz del repo.  
Si el código está dentro de carpeta `dashboard/`, escribe: `dashboard`

## URL correcta

Abre la URL del servicio, ejemplo:

`https://logistix-sce.onrender.com`

No uses la URL del panel de Render (eso puede decir Not Found).

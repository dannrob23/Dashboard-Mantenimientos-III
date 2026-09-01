# Dashboard Web — Mantenimientos Preventivos

Dashboard estático (HTML + **Plotly.js** + `dashboard.json`) que muestra el avance diario de los
mantenimientos preventivos, las desviaciones y el **mapa de Colombia por zonas intervenidas**.
Se publica gratis en **Vercel** (nunca se apaga) y se actualiza automáticamente cada día desde el
Excel maestro **`data/Bitacora_Diaria.xlsx`** mediante **GitHub Actions**.

> Base: patrón de dashboards estáticos Plotly.js + JSON (referencia: `plotly-dashboard-patient-data`),
> geodatos de Colombia (referencia: `colombia.geojson` de santiblanko / john-guerra).

## Estructura

```
dashboard_web/
├── data/
│   └── Bitacora_Diaria.xlsx      ← EXCEL MAESTRO (lo edita el equipo; versionado en Git)
├── public/
│   ├── index.html                ← dashboard (tarjetas, graficos, mapa)
│   ├── dashboard.json            ← datos generados por pipeline.py (NO editar a mano)
│   └── colombia.geojson          ← opcional: activa el mapa choropleth (ver abajo)
├── pipeline.py                   ← lee el Excel, recalcula KPIs y genera dashboard.json
├── descargar_geojson.py          ← descarga colombia.geojson (ejecutar una vez con internet)
├── requirements.txt
└── .github/workflows/diario.yml  ← GitHub Actions: regenera datos cada dia (06:00 CO) y al subir el Excel
```

## Puesta en marcha (una sola vez, ~10 minutos)

1. **Crear el repositorio** en GitHub (privado) y subir esta carpeta:
   ```bash
   git init
   git add .
   git commit -m "Dashboard mantenimientos preventivos"
   git branch -M main
   git remote add origin https://github.com/<TU_USUARIO>/<TU_REPO>.git
   git push -u origin main
   ```
2. **Conectar Vercel (gratis)**:
   - https://vercel.com → *Add New → Project* → importar el repo.
   - Framework Preset: **Other** · Build Command: *(vacío)* · Output Directory: **public**.
   - Deploy. El dashboard queda en `https://tu-proyecto.vercel.app` y **se redeploya solo con cada push**.
3. **(Opcional pero recomendado) Activar el mapa choropleth** — ejecutar una vez con internet:
   ```bash
   pip install -r requirements.txt
   python descargar_geojson.py     # guarda public/colombia.geojson
   git add public/colombia.geojson && git commit -m "geojson departamentos" && git push
   ```
   Sin este archivo el mapa funciona igual en **modo burbuja** (puntos por departamento).

## Flujo diario del equipo de mantenimiento

1. Abrir `data/Bitacora_Diaria.xlsx` (el mismo Excel con fórmulas y validaciones) y diligenciar
   los mantenimientos del día (estado, fechas reales, causales, actas).
2. Guardar, y subir el cambio al repo (GitHub Desktop: *commit + push*; o en la web: editar y commit).
3. GitHub Actions regenera `dashboard.json` y hace push → **Vercel publica automáticamente (~2 min)**.
4. También hay ejecución **automática diaria a las 06:00 CO** (workflow `diario.yml`) y botón
   *Run workflow* manual en la pestaña Actions.

## Probar en local

```bash
cd dashboard_web
pip install -r requirements.txt
python pipeline.py                     # genera public/dashboard.json
python -m http.server 8080 --directory public
# abrir http://localhost:8080
```

## Cambiar el periodo mostrado

Por defecto `pipeline.py` usa el mes del cierre más reciente. Para forzar:
```bash
python pipeline.py --mes 9 --ano 2026
```
(Ajuste la línea `run: python pipeline.py` del workflow si quiere un periodo fijo.)

## Seguridad y gobernanza

- Repo **privado** = control de cambios completo (historial del Excel y del código, ramas, issues).
- La URL de Vercel no se indexa; si necesita acceso con contraseña, agregue un login simple en
  `index.html` o un middleware en Vercel.
- El Excel maestro queda versionado: cualquier cambio queda registrado en Git (quién y cuándo).
- Backup implícito: cada versión del Excel vive en el historial del repo.

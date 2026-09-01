# Dashboard Mantenimientos Preventivos BAC

Dashboard corporativo moderno para visualización de mantenimientos preventivos del Banco Agrario de Colombia.

## Stack Tecnológico

- **Framework:** Next.js 14 (App Router)
- **Lenguaje:** TypeScript
- **Estilos:** Tailwind CSS
- **Gráficos:** Recharts
- **Iconos:** Lucide React
- **Procesamiento Excel:** xlsx (SheetJS)
- **Despliegue:** Vercel / Netlify

## Estructura del Proyecto

```
dashboard-bac/
├── data/
│   ├── raw/                    # Excel original (Bitacora_Final_Dashboard_BAC.xlsx)
│   └── processed/              # JSON procesado (bitacora.json)
├── public/                     # Archivos estáticos
├── scripts/
│   └── processExcel.js         # Script de procesamiento
├── src/
│   ├── app/                    # Páginas y layout
│   ├── components/             # Componentes UI
│   │   ├── cards/              # KPI Cards
│   │   ├── charts/             # Gráficos
│   │   ├── layout/             # Header, Sidebar
│   │   ├── tables/             # Tablas
│   │   └── ui/                 # Componentes base
│   ├── lib/                    # Lógica de negocio
│   │   ├── parser.ts           # Parser de datos
│   │   ├── kpis.ts             # Cálculo de KPIs
│   │   └── utils.ts            # Utilidades
│   └── types/                  # Tipos TypeScript
└── .github/
    └── workflows/              # GitHub Actions
```

## Instalación Local

```bash
# Instalar dependencias
npm install

# Procesar Excel a JSON (una vez)
node scripts/processExcel.js

# Iniciar servidor de desarrollo
npm run dev

# Abrir en navegador
# http://localhost:3000
```

## Despliegue en Vercel

1. **Sube el proyecto a GitHub:**
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/tu-usuario/dashboard-bac.git
   git push -u origin main
   ```

2. **Conecta con Vercel:**
   - Ve a [vercel.com](https://vercel.com)
   - Importa el repositorio
   - Vercel detectará automáticamente Next.js
   - Click en "Deploy"

3. **Configuración automática:**
   - Cada push a `main` redespliega automáticamente
   - Cuando se actualiza el Excel, GitHub Actions procesa los datos
   - Vercel detecta el cambio y redespliega

## KPIs del Dashboard

### Tier 1 - Esenciales
1. **% Cumplimiento General** - Finalizadas + Cerradas / Total
2. **Total Mantenimientos** - Cantidad total de registros
3. **% Avance Equipos** - Equipos intervenidos / Total
4. **Días Desviación** - Promedio de días de retraso

### Tier 2 - Operativos
5. **Cumplimiento por Región** - Barras horizontales
6. **Distribución por Estado** - Barras verticales con colores
7. **Causales de Desviación** - Dona con leyenda
8. **Sedes con Riesgo** - Tabla de top 10

## Flujo de Datos

```
Excel (Equipo Calidad)
   ↓
Sube a /data/raw/ en GitHub
   ↓
GitHub Action procesa → genera JSON
   ↓
Vercel redespliega automáticamente
   ↓
Dashboard actualizado en producción
```

## Actualizar Datos

### Opción 1: Automática (Recomendado)
1. Actualiza el archivo `Bitacora_Final_Dashboard_BAC.xlsx`
2. Súbelo a `data/raw/` en el repositorio
3. GitHub Actions procesa y redespliega automáticamente

### Opción 2: Manual
```bash
# 1. Reemplaza el archivo Excel en data/raw/
# 2. Ejecuta el procesador
node scripts/processExcel.js

# 3. Commit y push
git add .
git commit -m "Actualizar datos bitácora"
git push
```

## Personalización

### Colores
Edita `tailwind.config.ts` para cambiar la paleta corporativa.

### KPIs
Edita `src/lib/kpis.ts` para modificar los cálculos.

### Componentes
- `src/components/cards/` - KPI Cards
- `src/components/charts/` - Gráficos
- `src/app/page.tsx` - Página principal

## Licencia

Propietario - Banco Agrario de Colombia © 2026

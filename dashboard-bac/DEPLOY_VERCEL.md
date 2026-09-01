# 🚀 Guía Rápida para Deployar en Vercel

## Opción 1: Desde la Web (Más Fácil) ⭐

### Paso 1: Sube el código a GitHub

Si aún no lo has hecho, sube el proyecto a GitHub:

```bash
cd "C:\Users\darobles\Documents\Deepseek Harness\Dashboard BAC\dashboard-bac"

# Agregar todos los archivos
git add .

# Hacer commit
git commit -m "Dashboard BAC v1.0 - Listo para deploy"

# Si no tienes repositorio remoto, créalo primero en github.com
# Luego conéctalo:
git remote add origin https://github.com/TU-USUARIO/dashboard-bac.git
git branch -M main
git push -u origin main
```

### Paso 2: Conectar con Vercel

1. **Ve a:** https://vercel.com
2. **Click en:** "Sign Up" o "Log In"
3. **Selecciona:** "Continue with GitHub"
4. **Autoriza** a Vercel a acceder a tus repositorios

### Paso 3: Importar el Proyecto

1. **Click en:** "Add New Project" o "Import Project"
2. **Busca:** el repositorio `dashboard-bac`
3. **Click en:** "Import"

### Paso 4: Configurar el Deploy

Vercel detectará automáticamente que es Next.js. Verifica:

- **Framework Preset:** Next.js ✓
- **Root Directory:** `./` (dejar por defecto)
- **Build Command:** `next build` (automático)
- **Output Directory:** `.next` (automático)
- **Install Command:** `npm install` (automático)

### Paso 5: Deploy

1. **Click en:** "Deploy"
2. **Espera:** 2-3 minutos mientras construye
3. **¡Listo!** Tu dashboard estará en línea

La URL será algo como: `https://dashboard-bac-tu-usuario.vercel.app`

---

## Opción 2: Desde la Terminal con Vercel CLI

Si prefieres usar la terminal:

### Paso 1: Instalar Vercel CLI

```bash
npm install -g vercel
```

### Paso 2: Login

```bash
vercel login
```

Te pedirá tu email y enviará un link de verificación.

### Paso 3: Deploy

```bash
cd "C:\Users\darobles\Documents\Deepseek Harness\Dashboard BAC\dashboard-bac"
vercel
```

Responde las preguntas:
- **Set up and deploy?** → Y
- **Which scope?** → Selecciona tu cuenta
- **Link to existing project?** → N
- **Project name?** → dashboard-bac (o el nombre que quieras)
- **In which directory is your code located?** → ./
- **Override settings?** → N

### Paso 4: Deploy a Producción

```bash
vercel --prod
```

---

## 🔧 Verificación Post-Deploy

Una vez deployado, verifica:

1. **URL funciona:** Abre la URL en tu navegador
2. **Datos cargan:** Deberías ver los 200 registros de muestra
3. **Gráficos se muestran:** KPI cards, barras, dona
4. **Sin errores:** Abre la consola del navegador (F12) y verifica

---

## 🐛 Si el Deploy Falla

### Error: "Build failed"

**Causa común:** Dependencias no instaladas

**Solución:**
```bash
# Localmente, verifica que compile
cd dashboard-bac
npm install
npm run build
```

Si hay errores, corrígelos antes de hacer push.

### Error: "Module not found: Can't resolve 'xlsx'"

**Solución:** El parser ya no importa xlsx en el frontend. Verifica que el cambio se subió:
```bash
git add src/lib/parser.ts
git commit -m "fix: remover import xlsx del frontend"
git push
```

### Error: "Type errors"

**Solución:** Ya está configurado TypeScript. Si hay errores, revisa:
```bash
npx tsc --noEmit
```

---

## 📝 Configuración Adicional (Opcional)

### Variables de Entorno

Si necesitas variables de entorno, en Vercel:
1. Ve a tu proyecto → Settings → Environment Variables
2. Agrega las que necesites

### Dominio Personalizado

1. Ve a tu proyecto → Settings → Domains
2. Agrega tu dominio (ej: dashboard.bac.com.co)
3. Configura DNS según las instrucciones

---

## ✅ Checklist Final

- [ ] Código subido a GitHub
- [ ] Vercel conectado al repo
- [ ] Deploy completado sin errores
- [ ] URL funciona y muestra datos
- [ ] Gráficos interactivos funcionan
- [ ] Responsive (probar en móvil)

---

## 🎯 Próximos Pasos Después del Deploy

1. **Reemplazar datos de muestra:**
   - Sube el Excel real a `data/raw/Bitacora_Final_Dashboard_BAC.xlsx`
   - GitHub Action procesará automáticamente
   - Vercel redespliegará en 2-3 minutos

2. **Personalizar:**
   - Cambiar colores en `tailwind.config.ts`
   - Agregar logo en `public/`
   - Modificar textos en componentes

3. **Monitorear:**
   - Vercel Analytics (incluido gratis)
   - Logs en tiempo real

---

¿Necesitas ayuda con algún paso específico del deploy?

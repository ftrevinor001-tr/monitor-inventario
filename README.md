# Monitor de Inventario — Faltantes y Riesgo

Dashboard interactivo para analizar claves con días de faltante y riesgo en inventario.

---

## 🚀 Configuración inicial en GitHub

### 1. Crear el repositorio

1. En GitHub → **New repository**
2. Nombre: `monitor-inventario` (o el que prefieras)
3. Visibilidad: **Public** (necesario para GitHub Pages gratuito)
4. Haz clic en **Create repository**

### 2. Subir los archivos

Sube todos estos archivos al repositorio:
```
monitor-inventario/
├── index.html                  ← dashboard listo para ver
├── dashboard_template.html     ← plantilla base (no editar)
├── actualizar_dashboard.py     ← script de actualización
├── requirements.txt
├── .github/
│   └── workflows/
│       └── actualizar.yml      ← automatización con GitHub Actions
└── data/
    └── datos.xlsx              ← tu archivo Excel actual
```

### 3. Activar GitHub Pages

1. Ve a **Settings → Pages**
2. Source: **Deploy from a branch**
3. Branch: `main` / carpeta: `/ (root)`
4. Haz clic en **Save**

El dashboard estará disponible en:
`https://TU-USUARIO.github.io/monitor-inventario/`

### 4. Activar permisos de Actions

1. Ve a **Settings → Actions → General**
2. En *Workflow permissions* selecciona: **Read and write permissions**
3. Haz clic en **Save**

---

## 🔄 Actualizar datos (flujo normal)

### Opción A — Automático (recomendado)

1. Reemplaza `data/datos.xlsx` con tu nuevo archivo Excel
2. Haz commit y push a GitHub
3. GitHub Actions ejecuta automáticamente el script (~1 minuto)
4. El dashboard en GitHub Pages se actualiza solo

```bash
# Desde tu computadora con git instalado:
cp /ruta/a/tu/nuevo_archivo.xlsx data/datos.xlsx
git add data/datos.xlsx
git commit -m "Actualizar datos de inventario"
git push
```

### Opción B — Manual desde GitHub Web

1. Ve a tu repositorio en GitHub
2. Entra a la carpeta `data/`
3. Haz clic en **Add file → Upload files**
4. Sube tu nuevo `.xlsx` (reemplaza `datos.xlsx`)
5. Haz clic en **Commit changes**
6. GitHub Actions se activa automáticamente

### Opción C — Ejecutar el script localmente

```bash
# Instalar dependencias (solo la primera vez)
pip install -r requirements.txt

# Actualizar con tu Excel
python actualizar_dashboard.py --excel data/tu_archivo.xlsx

# El index.html se regenera localmente
# Luego haz commit y push manualmente
git add index.html data/
git commit -m "Datos actualizados"
git push
```

---

## 📊 Características del dashboard

- Filtros en cascada: Comprador → Marca → Clasificación → Rotación → Mes → Estatus
- Tablero de faltantes con agrupación por semanas (1-7 días, 8-14, etc.)
- Exportar tabla filtrada a Excel
- 💾 Guardar dashboard con datos actualizados embebidos
- Excluye automáticamente claves con clasificación **NUEVOS**

---

## 🛠️ Requisitos del Excel

El archivo debe tener una hoja llamada **"Historico de faltantes"** con estas columnas:

| Columna | Descripción |
|---|---|
| Día/ mes | Fecha del registro |
| Clave | Código del producto |
| Nombre | Descripción del producto |
| Marca | Marca |
| COMPRADOR | Comprador responsable |
| Clasificación | Clasificación (A, B, C…) |
| Clasificacion | Tipo de rotación (SAGRADOS, ALTA, MEDIA, LENTA) |
| Clasificacion 2 | Estatus: FALTANTE o RIESGO |

---

## ❓ Solución de problemas

**GitHub Actions no se ejecuta:**
- Verifica que el workflow tenga permisos de escritura (Settings → Actions → General)
- El archivo Excel debe estar en la carpeta `data/`

**El dashboard no se actualiza en GitHub Pages:**
- GitHub Pages puede tardar 1-3 minutos en reflejar cambios
- Verifica en la pestaña **Actions** que el workflow terminó exitosamente

**Error en el script Python:**
- El Excel debe tener la columna `Clasificacion 2` con valores `FALTANTE` o `RIESGO`
- Instala las dependencias: `pip install pandas openpyxl`

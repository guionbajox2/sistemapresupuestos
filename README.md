# 📋 Solicitud de Presupuesto — Sistema de Generación Excel

Sistema web para generar solicitudes oficiales de presupuesto en formato Excel,
desplegado en Netlify con Netlify Functions para el procesamiento backend.

---

## 📁 Estructura del Proyecto

```
mi-proyecto/
│
├── index.html                ← Formulario principal
├── style.css                 ← Estilos (diseño institucional)
├── script.js                 ← Lógica frontend (validaciones, tabla, fetch)
│
├── plantilla.xlsx            ← ⚠ DEBES PROVEER ESTE ARCHIVO (ver abajo)
│
├── package.json              ← Dependencias Node.js
├── netlify.toml              ← Configuración de Netlify
│
└── netlify/
    └── functions/
        └── generar.js        ← Netlify Function: genera el Excel
```

---

## ⚠ ARCHIVO PLANTILLA REQUERIDO

**Debes colocar tu archivo `plantilla.xlsx` en la raíz del proyecto.**

El archivo debe contener una hoja con las celdas pre-formateadas.
El backend insertará valores en estas celdas sin modificar los estilos existentes:

| Campo                          | Celda  |
|--------------------------------|--------|
| Fecha de la Solicitud          | L2     |
| Organización                   | H1     |
| Nombre del Líder               | D22    |
| Nombre del Receptor            | L8, G22|
| Tipo de Solicitud (texto)      | J9     |
| Nombre de la Actividad         | E5     |
| Fecha de la Actividad          | C6     |
| Hora de la Actividad           | H6     |
| Lugar                          | D7     |
| Cantidad de Asistencia         | B8     |
| Actividad Dirigida a           | C9     |
| Propósito                      | A12    |
| Meta: Edificar la Fe           | J18    |
| Meta: Diversión y Unidad       | J19    |
| Meta: Crecimiento Personal     | J21    |
| Meta: Fortalecer Familias      | J22    |
| Meta: Obra de Salvación        | J23    |
| Meta Otra (texto)              | K25    |
| Meta Otra (marca X)            | J25    |
| Cantidad del Producto (x13)    | A29:A41|
| Nombre del Producto (x13)      | B29:B41|
| Costo Unitario (x13)           | J29:J41|
| Costo Total (x13)              | K29:K41|
| Total General                  | K42, L7|

---

## 🚀 Instalación y Ejecución Local

### Requisitos previos
- **Node.js** v18 o superior → https://nodejs.org
- **npm** v9 o superior (incluido con Node.js)
- **Cuenta en Netlify** → https://app.netlify.com/signup (gratis)

### Paso 1 — Clonar o descargar el proyecto
```bash
# Opción A: si tienes Git
git clone <url-del-repositorio>
cd mi-proyecto

# Opción B: si lo descargaste como ZIP
# Descomprimir y abrir la carpeta en terminal
cd mi-proyecto
```

### Paso 2 — Instalar dependencias
```bash
npm install
```
Esto instalará `exceljs` y `netlify-cli`.

### Paso 3 — Colocar la plantilla
```bash
# Copia tu plantilla a la raíz del proyecto
cp /ruta/a/tu/plantilla.xlsx ./plantilla.xlsx
```

### Paso 4 — Instalar Netlify CLI globalmente (recomendado)
```bash
npm install -g netlify-cli
```

### Paso 5 — Iniciar servidor de desarrollo local
```bash
netlify dev
```
Esto iniciará el servidor en `http://localhost:8888` con las Functions activas.

> Si el puerto 8888 está en uso, Netlify usará el siguiente disponible.

---

## 🌐 Deploy en Netlify

### Opción A — Deploy con Netlify CLI (recomendado)

```bash
# 1. Autenticarse en Netlify
netlify login

# 2. Crear nuevo sitio (solo la primera vez)
netlify init

# 3. Deploy a producción
netlify deploy --prod
```

### Opción B — Deploy desde GitHub (interfaz web)

1. Subir el proyecto a un repositorio GitHub
2. Ir a https://app.netlify.com → **"Add new site"** → **"Import an existing project"**
3. Conectar con GitHub y seleccionar el repositorio
4. Configuración de build:
   - **Base directory**: (dejar vacío o `/`)
   - **Build command**: (dejar vacío)
   - **Publish directory**: `.` o `/`
5. Hacer clic en **"Deploy site"**

> ⚠ **IMPORTANTE**: No olvides subir `plantilla.xlsx` al repositorio.
> Si no quieres exponer el archivo en GitHub, usa [Netlify Large Media](https://docs.netlify.com/large-media/overview/).

### Opción C — Drag & Drop (más simple)

1. Comprimir la carpeta del proyecto (incluyendo `plantilla.xlsx`)
2. Ir a https://app.netlify.com/drop
3. Arrastrar la carpeta al área de drop
4. ✅ Listo — obtendrás una URL pública

---

## 🔧 Errores Comunes y Soluciones

### ❌ Error: `plantilla.xlsx no encontrada en el servidor`
**Causa**: El archivo `plantilla.xlsx` no está en la raíz del proyecto durante el deploy.
**Solución**: 
- Verificar que `plantilla.xlsx` existe en la misma carpeta que `index.html`
- Si usas Git, asegúrate de que el archivo NO esté en `.gitignore`
- Hacer `git add plantilla.xlsx && git commit -m "Add plantilla"`

---

### ❌ Error: `Cannot find module 'exceljs'`
**Causa**: Las dependencias no se instalaron correctamente.
**Solución**:
```bash
rm -rf node_modules package-lock.json
npm install
```

---

### ❌ Error en Netlify: `Function timeout`
**Causa**: La Function tardó más de 10 segundos (límite del plan Free).
**Solución**:
- Reducir el tamaño de `plantilla.xlsx` (eliminar imágenes pesadas, hojas innecesarias)
- Actualizar a un plan Netlify con límite de 26 segundos

---

### ❌ Error: `isBase64Encoded` — El archivo Excel descargado está corrupto
**Causa**: Problemas con la codificación del buffer en la respuesta.
**Solución**: Verificar que en `generar.js` la respuesta incluye `isBase64Encoded: true` y el body es `buffer.toString('base64')`.

---

### ❌ El formulario envía pero no se descarga el archivo
**Causa**: Posible error de CORS o de configuración.
**Solución**:
1. Abrir DevTools del navegador → pestaña **Network**
2. Buscar la solicitud a `/.netlify/functions/generar`
3. Revisar el **Status Code** y el cuerpo de la respuesta de error
4. Corregir según el mensaje de error

---

### ❌ Error: `La plantilla no contiene ninguna hoja de trabajo`
**Causa**: El archivo `plantilla.xlsx` está vacío o dañado.
**Solución**: 
- Abrir `plantilla.xlsx` en Excel/LibreOffice y verificar que tiene al menos una hoja
- Re-guardar el archivo en formato `.xlsx` (no `.xls` antiguo)

---

### ❌ En desarrollo local: `netlify dev` no funciona
**Causa**: Netlify CLI no está instalado globalmente.
**Solución**:
```bash
# Instalar globalmente
npm install -g netlify-cli

# O ejecutar con npx
npx netlify dev
```

---

## 📝 Personalización

### Cambiar el nombre de la hoja de trabajo usada
En `netlify/functions/generar.js`, línea donde se obtiene la hoja:
```js
// Por índice (0 = primera hoja):
const ws = workbook.worksheets[0];

// Por nombre de hoja:
const ws = workbook.getWorksheet('NombreDeLaHoja');
```

### Agregar más campos
1. Agregar el campo en `index.html` con el `name` correspondiente
2. Serializar el nuevo campo en `script.js` dentro de `serializeForm()`
3. Agregar la celda destino en el `MAPA_CELDAS` de `generar.js`
4. Insertar el valor con `setCellValue(ws, MAPA_CELDAS.tuCampo, valor)`

---

## 🔐 Seguridad implementada

- Sanitización de strings (prevención de inyección de fórmulas Excel)
- Validación de tipos de datos y rangos numéricos
- `JSON.parse` protegido con `try/catch`
- Método HTTP restringido a POST
- Solo acepta hasta 13 productos (límite de la plantilla)
- Headers de seguridad HTTP configurados en `netlify.toml`

---

## 📦 Dependencias

| Paquete       | Versión  | Uso                              |
|---------------|----------|----------------------------------|
| exceljs       | ^4.4.0   | Lectura y escritura de archivos Excel |
| netlify-cli   | ^17.0.0  | Desarrollo local y deploy        |

---

*Sistema de Solicitudes Oficiales — Uso Interno*

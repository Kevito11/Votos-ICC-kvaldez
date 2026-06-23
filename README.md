# 🗳️ Sistema de Votaciones - ICC

> Plataforma web de votación electrónica para la admisión de membresía en la **Iglesia Cristiana de Confraternidad (ICC)**. Permite a los miembros votar a favor o en contra de candidatos a membresía de forma digital, segura y en tiempo real.

---

## 📋 Descripción

Este sistema reemplaza el proceso manual de votación por mano alzada o papeleta física, ofreciendo una experiencia moderna, transparente y auditable. Cada miembro de la iglesia puede acceder con su código único, votar sobre cada candidato postulante, y el administrador puede monitorear los resultados en tiempo real desde un panel de control.

---

## ✨ Funcionalidades

### 👤 Panel del Votante
- Acceso mediante **código único** de miembro (validado contra Google Sheets)
- Votación por cada candidato con opciones: **Apruebo**, **No Apruebo** u **Objeción con comentario**
- Validación de respuestas: no permite comentarios vacíos ni respuestas vagas (ej. "porque no", "no sé")
- Indicador de **progreso de votación** con candidatos restantes
- Mensaje flotante al finalizar mostrando candidatos que aún faltan por votar
- **Cierre de sesión automático** por inactividad (configurable)
- Diseño responsivo optimizado para uso en móvil durante asambleas

### 🛡️ Panel de Administrador
- Acceso protegido con **contraseña segura** (configurable por variable de entorno)
- Cierre de sesión automático tras **5 minutos de inactividad**
- Dashboard en tiempo real con:
  - Total de votos emitidos vs. miembros registrados
  - Barras de progreso por candidato con colores: 🟢 Aprobados / 🔴 No aprobados
  - Tabla de desglose de votos individuales por votante
  - Gestión de candidatos (agregar, editar, eliminar)
- Exportación y visualización de resultados

### 🔗 Integración con Google Sheets
- Los datos de **votantes** y **candidatos** se leen desde Google Sheets vía Apps Script
- Los votos se registran en tiempo real en la hoja de cálculo
- No requiere base de datos adicional para una instalación básica

---

## 🛠️ Tecnologías

| Tecnología | Uso |
|---|---|
| **React 19** | Framework de UI |
| **Vite 8** | Bundler y servidor de desarrollo |
| **Vanilla CSS** | Estilos con glassmorphism y diseño oscuro |
| **Google Apps Script** | Backend serverless conectado a Google Sheets |
| **Supabase** *(opcional)* | Almacenamiento de imágenes de candidatos |

---

## 🚀 Instalación y uso local

### Prerrequisitos
- Node.js 18+
- npm

### Pasos

```bash
# 1. Clonar el repositorio
git clone https://github.com/Kevito11/Votos-ICC-kvaldez.git
cd Votos-ICC-kvaldez

# 2. Instalar dependencias
npm install

# 3. Configurar variables de entorno
cp .env.example .env
# Editar .env con tus propios valores

# 4. Iniciar servidor de desarrollo
npm run dev
```

Abre [http://localhost:5173](http://localhost:5173) en tu navegador.

---

## ⚙️ Variables de entorno

Crea un archivo `.env` en la raíz del proyecto con las siguientes variables:

```env
# URL del Google Apps Script desplegado como Web App
VITE_SHEET_URL=https://script.google.com/macros/s/TU_SCRIPT_ID/exec

# URLs específicas por hoja (opcional, si usas el mismo script usa la misma URL)
VITE_SHEET_URL_VOTERS=https://script.google.com/macros/s/TU_SCRIPT_ID/exec
VITE_SHEET_URL_CANDIDATES=https://script.google.com/macros/s/TU_SCRIPT_ID/exec

# Contraseña del panel de administrador
VITE_ADMIN_PASSWORD=tu_contraseña_segura

# Supabase (opcional - para imágenes de candidatos)
VITE_SUPABASE_URL=https://tu-proyecto.supabase.co
VITE_SUPABASE_KEY=tu_anon_key
VITE_SUPABASE_BUCKET=candidatos
```

> ⚠️ **Nunca subas el archivo `.env` a GitHub.** Ya está incluido en el `.gitignore`.

---

## 📊 Estructura del Google Sheet

El proyecto requiere un Google Sheet con las siguientes hojas:

| Hoja | Columnas requeridas |
|---|---|
| `Votantes` | `id`, `nombre`, `codigo` |
| `Candidatos` | `id`, `nombre`, `descripcion`, `foto` |
| `Votos` | `votanteId`, `candidatoId`, `voto`, `comentario`, `timestamp` |

El Apps Script debe estar desplegado como **Web App** con acceso público (`Anyone`).

---

## 🏗️ Estructura del proyecto

```
src/
├── App.jsx                  # Componente raíz, manejo de estado global
├── App.css                  # Estilos del layout principal
├── index.css                # Design system y estilos globales
├── components/
│   ├── AdminPanel.jsx       # Panel de administración y dashboard
│   ├── VoterPanel.jsx       # Flujo de votación para miembros
│   └── Tooltip.jsx          # Componente de tooltip reutilizable
└── utils/
    ├── api.js               # Integración con Google Sheets API
    └── mockData.js          # Datos de prueba para desarrollo local
```

---

## 🔒 Seguridad

- Las contraseñas de administrador se manejan por variables de entorno
- Los códigos de votante se validan en el backend (Google Sheets)
- Cada votante solo puede votar **una vez por candidato**
- La sesión de administrador expira automáticamente por inactividad

---

## 📸 Capturas

> *Sistema de votación con diseño oscuro y glassmorphism, optimizado para uso durante asambleas de membresía.*

---

## 📄 Licencia

Proyecto privado desarrollado para uso interno de la **Iglesia Cristiana de Confraternidad (ICC)**.

---

<div align="center">
  Desarrollado con ❤️ para la ICC
</div>

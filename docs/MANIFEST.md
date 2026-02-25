# 🗺️ Manifest de Estructura (Repo Público)

## 🎯 Propósito

Este documento resume cómo está organizada la extracción VBA en el repo público.

---

## 🧱 Arquitectura actual (real del repo)

La orquestación está distribuida en varias áreas:

| Ruta | Descripción |
|---|---|
| `src/vba/bootstrap/modStartup.bas` + `AppContext.cls` | Inicio del workbook y asignación de tablas/rangos. |
| `src/vba/bootstrap/modState.bas` | Estado compartido de runtime. |
| `src/vba/excel/modImport.bas` + `modImportPdf.bas` | Pipeline de importación y despacho de parsers. |
| `src/vba/validation/modValidation.bas` | Orquestación de validaciones por fila. |
| `src/vba/integrations/*.bas` | Integraciones SAP / RetailWeb / Python / auxiliares. |
| `src/vba/ui/*.frm` + `modButtons.bas` + `modProgress.bas` | UI para operadores y progreso. |
| `src/vba/core/*.bas` | Lógica determinística extraída para tests. |
| `src/vba/tests/*.bas` | Tests, runner headless y logging. |

---

## ⚠️ Nota sobre `src/vba/bootstrap/Módulo1.bas`

En el proyecto original existía un módulo “monolítico” (miles de líneas) que concentraba gran parte de la orquestación y dependencias del entorno.

En este repo público, **`src/vba/bootstrap/Módulo1.bas` es solo un placeholder (stub)** que mantiene el nombre del módulo (`Attribute VB_Name`) pero **no contiene la implementación**.  
Esto es intencional: evita publicar un monolito acoplado y guía la evaluación hacia la estructura modular, que es lo que se revisa en este portfolio.

---

## 📁 Áreas del repositorio

### 🔬 Core (`src/vba/core/`)

Funciones puras para:

- Clasificación.
- Reglas de validación.
- Armado de mensajes/textos.
- Constantes compartidas por tests del core.

### 🧪 Tests (`src/vba/tests/`)

- Runner headless (`RunCoreTests`).
- Assertions.
- Logging a archivo (`artifacts/`).
- Módulos de tests por categoría del core.

### ⚙️ Excel / Orquestación (`src/vba/excel/`, `src/vba/validation/`, `src/vba/reporting/`)

- Etapas `Importar_*`.
- Importación de PDFs vía Power Query.
- Validaciones y actualización de estado por fila.
- Reportes y salidas operativas.

### 🔗 Integraciones (`src/vba/integrations/`)

- SAP GUI Scripting.
- RetailWeb (portal web interno de retail).
- Ejecución de scripts Python auxiliares.

### 🧩 Parsers (`src/vba/parsers/`)

Parsers de PDF por formato/proveedor anonimizados como `modParserVendorNN.bas`.
Ver `docs/PARSERS.md`.

### 🖥️ UI (`src/vba/ui/`)

UserForms y módulos de botones/progreso exportados desde el workbook.
Se versionan `.frm` (texto); los `.frx` binarios se omiten del release público.

---

## 📝 Notas de extracción y publicación

- El código VBA se versiona como fuente exportada (`.bas`, `.cls`, `.frm`).
- Los `.frx` se excluyen por ser binarios y no aportar valor a la revisión técnica pública.
- Strings e identificadores sensibles fueron redactados/anonimizados.

# Migración y Validación de Facturas (AP) — Automatización end-to-end (Excel VBA + Python)

![Excel](https://img.shields.io/badge/Excel-VBA-217346)
![Python](https://img.shields.io/badge/Python-Automation-3776AB)
![Power%20Query](https://img.shields.io/badge/Power%20Query-PDF%20Parsing-00A1E0)
![Selenium](https://img.shields.io/badge/Selenium-Headless-43B02A)
![SAP](https://img.shields.io/badge/SAP-GUI%20Scripting-0FAAFF)
![Portfolio](https://img.shields.io/badge/Portfolio-Safe%20Release-6E56CF)

Este repositorio presenta una automatización real (portfolio-safe) de **Cuentas a Pagar (AP)**: orquesta el ingreso de **facturas en PDF**, extrae y normaliza datos, cruza contra un portal interno (`RetailWeb`/`RW`), aplica **validaciones fiscales y de negocio**, y prepara/verifica la contabilización y el seguimiento en SAP.

> 📈 **Resultado real:** ~x8 en productividad operativa · de ~60 a ~500 documentos/día · errores operativos reducidos al mínimo

![Pipeline](assets/demo2.png)

> ⚠️ Versión pública: prioriza **arquitectura + reglas determinísticas + evidencia reproducible** sin exponer datos sensibles. Las integraciones reales (SAP/Outlook/portal) requieren entorno corporativo.

---

![Impacto](assets/highlights/impacto.png) 

![UX/UI](assets/highlights/ux.png)

![Seguridad](assets/highlights/seguridad.png) 

---

## 🔗 Accesos rápidos

| Tipo | Recurso |
|---|---|
| 📄 Deck (PDF, 7 slides) | [Facturas-ASAP-Portfolio.pdf](assets/deck/Facturas-ASAP-Portfolio.pdf) |
| 🖥️ UI principal (capturas) | [TOOL_UI_OVERVIEW.md](docs/TOOL_UI_OVERVIEW.md) |
| 🧩 UserForms (galería) | [UI_FORMS_GALLERY.md](docs/UI_FORMS_GALLERY.md) |
| 📁 Descripción general | [PORTFOLIO_OVERVIEW.md](docs/PORTFOLIO_OVERVIEW.md) |
| 📚 Case Study técnico | [CASE_STUDY.md](docs/CASE_STUDY.md) |
| 🧪 Cómo correr evidencia | [TESTING.md](docs/TESTING.md) |
| 🗺️ Mapa del repo | [MANIFEST.md](docs/MANIFEST.md) |
| 🧩 Parsers PDF (VendorXX) | [PARSERS.md](docs/PARSERS.md) |

---

## 🧠 Qué demuestra

- **Automatización end-to-end** de un flujo AP completo: documentos → extracción → cruce → validación → decisión operativa → preparación/chequeo.
- **Integración de sistemas heterogéneos**: Excel/VBA + Power Query (PDF) + scripts Python + automatización web + SAP GUI Scripting.
- **Motor de decisiones trazable**: estados determinísticos por fila + mensajes consistentes + tolerancias configurables.
- **Diseño operator-centric**: UX interna (UserForms), controles, progreso, rollback/seguridad, reducción de intervención manual.
- **Evidencia reproducible sin VPN**: tests headless del core + scans de prepublicación + export del paquete público.

---

## 🧱 Decisiones técnicas clave

| Decisión | Por qué importa |
|---|---|
| **Arquitectura híbrida VBA + Python** | Combina la integración nativa Office/SAP (VBA) con la potencia de librerías modernas web/PDF (Python). |
| **Power Query como ETL embebido** | Uso de `Pdf.Tables` para extracción estructurada sin OCR, con transformaciones declarativas en lenguaje M. |
| **Separación Core vs Adapters** | Aísla la lógica determinística (reglas) de los side-effects (IO), permitiendo testabilidad real. |
| **Testing headless** | Runner propio en VBA que ejecuta validaciones con Excel invisible y reporta logs a archivos. |
| **Publicación segura** | Export a `dist/` sanitizado, scan de términos sensibles y normalización de encoding. |

---

## 🛠️ Stack tecnológico

### ⚙️ Backend / Scripting
![VBA](https://img.shields.io/badge/VBA-7.1-217346)
![Python](https://img.shields.io/badge/Python-3.11-3776AB)
![Pandas](https://img.shields.io/badge/Pandas-Data-150458)
![PyMuPDF](https://img.shields.io/badge/PyMuPDF-PDF-red)

### 📊 Data & ETL
![Power Query](https://img.shields.io/badge/Power%20Query-M%20Language-00A1E0)
![Analysis Services](https://img.shields.io/badge/Analysis%20Services-MDX%2FADOMD-yellow)
![SQLite](https://img.shields.io/badge/SQLite-Concept-003B57)

### 🔗 Integración
![SAP](https://img.shields.io/badge/SAP-GUI%20Scripting-0FAAFF)
![Selenium](https://img.shields.io/badge/Selenium-WebDriver-43B02A)
![Outlook](https://img.shields.io/badge/Outlook-COM-0078D4)

### 🚀 DevOps & Tooling
![PowerShell](https://img.shields.io/badge/PowerShell-Automation-5391FE)
![Git](https://img.shields.io/badge/Git-Version%20Control-F05032)
![Custom Runner](https://img.shields.io/badge/Custom-Test%20Runner-6E56CF)
![Public Release](https://img.shields.io/badge/Repo-Sanitizado-success)

---
## 🧭 Flujo end-to-end
```mermaid
flowchart TD
%% ===== ESTILOS =====
    classDef ingesta   fill:#1e3a5f,stroke:#4a9eff,color:#fff
    classDef parsing   fill:#1a472a,stroke:#4caf50,color:#fff
    classDef cruce     fill:#3d1a5f,stroke:#ce93d8,color:#fff
    classDef validacion fill:#7a3000,stroke:#ff9800,color:#fff
    classDef salida    fill:#003d4f,stroke:#26c6da,color:#fff

%% ===== INGESTA =====
    subgraph S1["📥  Ingesta de documentos"]
        A1(["📧 Outlook / Adjuntos PDF"])
        A2["🗂️ Organización & Normalización"]
        A3[("📊 Importación masiva en Excel")]
        A1 --> A2 --> A3
    end

%% ===== PARSING =====
    subgraph S2["🔍  Extracción & Parsing"]
        B1["📄 Lectura PDF - Power Query"]
        B2{"🏷️ Identificación de proveedor"}
        B3["🧩 Parser VendorXX"]
        B4[("📋 Tabla operativa estructurada")]
        B1 --> B2 --> B3 --> B4
    end

%% ===== CRUCE =====
    subgraph S3["🔗  Enriquecimiento & Cruce"]
        C1["🌐 Reporte RW headless / Cubo"]
        C2["🔎 Match referencia/remito + reglas"]
        C3[("✅ Campos RW enriquecidos")]
        C1 --> C2 --> C3
    end

%% ===== VALIDACIONES =====
    subgraph S4["⚙️  Validaciones & Decisión"]
        D1["🧮 Integridad fiscal"]
        D2["📐 Tolerancias configurables"]
        D3{"📏 Reglas determinísticas"}
        D4[["🏷️ Estado por fila + comentarios"]]
        D1 --> D2 --> D3 --> D4
    end

%% ===== SALIDA =====
    subgraph S5["🚀  Salida / Evidencia"]
        E1["📊 Reporte filtrado"]
        E2[("🏢 Acción en SAP")]
        E3["🧪 Tests + scan + export"]
        E4[("📦 dist/public_release")]
        E1 --> E2
        E3 --> E4
    end

%% ===== CONEXIONES ENTRE ETAPAS =====
    A3 --> B1
    B4 --> C1
    C3 --> D1
    D4 --> E1
    D4 --> E3

%% ===== ASIGNACIÓN DE COLORES =====
    class A1,A2,A3 ingesta
    class B1,B2,B3,B4 parsing
    class C1,C2,C3 cruce
    class D1,D2,D3,D4 validacion
    class E1,E2,E3,E4 salida
```
---

## ✅ Capacidades (resumen)

### Ingesta y preparación

- Importación masiva de PDFs (incluye multipágina).
- Normalización de referencias y metadatos operativos (sucursal/site, fechas, tipo de doc).

---

### Extracción (PDF → datos)

- Lectura por página con Power Query.
- Parsers anonimizados por proveedor (`Vendor01`..`VendorNN`) bajo contrato común.

---

### Cruce y enriquecimiento

- Cruce contra RetailWeb/RW (descarga y carga de reporte / o cubo según configuración).
- Match por referencia/remito con reglas especiales (FC/NC y variantes).

---

### Validaciones y decisión operativa

- Cálculo de integridad fiscal (totales vs componentes: IVA/II/percepciones).
- Aplicación de tolerancias y reglas determinísticas (estado + comentarios).

---

### Integraciones

- Automatización web (portal interno) y SAP GUI scripting.
- Utilidades Python para soporte operativo (redactadas en esta versión pública).

<details>
<summary><b>Ver capacidades completas (end-to-end)</b></summary>

**A) Convierte PDFs en registros estructurados listos para decidir acción**

Procesa PDFs multi-página, identifica proveedor (CUIT/heurísticas) y despacha a parser VendorXX.

Escribe en tabla operativa por fila: tipo doc, referencia/remito, fechas, CAE/CAEA, total/subtotales/IVA, impuestos internos, percepciones (IIBB multi-jurisdicción, municipal, etc.).

**B) Cruza contra RW / cubo corporativo**

Obtiene datos operativos por reporte descargado (headless) o por cubo (según configuración).

Completa por fila: estado pago, anulado, scan, fechas relevantes, totales RW, comentarios RW.

**C) Motor de validaciones determinístico**

Calcula diferencias (ej. integridad fiscal) y aplica tolerancias configurables.

Asigna estados operativos por fila (`OK / Revisar / Validar / Completar / etc.`) y genera mensajes trazables.

**D) Persistencia y recuperación operativa**

Upsert en base interna para recuperar datos sin PDF o reusar resultados.

**E) Integraciones reales (limitadas en versión pública)**

- Portal interno: búsquedas / impresión / cambio de estado.
- SAP: consultas de existencia y validaciones por transacciones.
- AFIP/ARCA: autocompletado de validación (dejando captcha al usuario).
- Outlook: descarga de adjuntos y organización por estructura.

**F) UX y robustez**

UserForms, progress, manejo de timeouts, rollback, protección de hojas, limpieza/renombrado de PDFs, reportes filtrados.

</details>

---

## 🗂️ Arquitectura de módulos

| Módulo | Ruta | Descripción |
|---|---|---|
| Core determinístico | `src/vba/core/` | Clasificación, validaciones, mensajes, tolerancias y constantes. |
| Parsers | `src/vba/parsers/` | Implementaciones VendorXX anonimizadas bajo contrato común. |
| Integraciones | `src/vba/integrations/` + `scripts/` | Automatización web, SAP, Outlook y utilidades Python (redactadas). |
| UI / Operación | `src/vba/ui/` + `src/vba/excel/` | UserForms, orquestación, interacción con hojas/tablas. |
| Tooling / Evidencia | `tools/` + `src/vba/tests/` | Harness, tests headless, scans y export del paquete público. |

---

## 🧪 Evidencia reproducible (sin VPN)

> Requiere Windows + Microsoft Excel instalado.

**1) Correr tests headless del core**

```powershell
powershell -ExecutionPolicy Bypass -File tools/run_core_tests.ps1
```

**2) Scan de prepublicación (redacción/sanitización)**

```powershell
powershell -ExecutionPolicy Bypass -File tools/prepublish_scan.ps1
```

**3) Export del paquete público (snapshot para repo público)**

```powershell
powershell -ExecutionPolicy Bypass -File tools/export_public_release.ps1
```

---

## 📌 Qué incluye / qué no incluye (versión pública)

**Incluye:**

- Reglas determinísticas del core (clasificación / validaciones / mensajes).
- Parsers anonimizados VendorXX.
- Tooling de evidencia (tests / scans / export).
- Documentación técnica.

**No incluye (ejecutable end-to-end):**

- Acceso real a SAP, Outlook y portal interno (requieren credenciales/entorno).
- Binarios `.frx` de UserForms (se versionan `.frm` en texto; ver `docs/NOTES_FORMS.md`).

---

## 🛡️ Seguridad y anonimización

Esta versión pública elimina o reemplaza:

- Proveedores reales, CUIT/CAE reales y datos sensibles.
- Rutas/URLs internas.
- Identificadores operativos.

**Convenciones:**

| Placeholder | Significado |
|---|---|
| `RetailWeb` / `RW` | Portal interno de retail. |
| `Sucursal` | Identificadores de sitio/negocio (término público). |
| `VendorXX` | Parser anonimizado por proveedor. |
| `<REDACTED>`, `<REDACTED_PATH>`, `<REDACTED_URL>`, `<REDACTED_ID_XX>` | Datos sensibles eliminados. |

> La password de protección de hojas no se versiona: se inyecta por variable de entorno `MIGRAR_PASSWORD`.

---

## 📚 Documentación

| Archivo | Descripción |
|---|---|
| `docs/CASE_STUDY.md` | Historia, enfoque y decisiones técnicas. |
| `docs/TESTING.md` | Harness y ejecución de tests. |
| `docs/PARSERS.md` | Contrato y estrategia de parsers VendorXX. |
| `docs/MANIFEST.md` | Estructura del repo. |
| `docs/NOTES_FORMS.md` | Forms y dependencias `.frx`. |
| `docs/PUBLIC_RELEASE_CHECKLIST.md` | Checklist para publicación. |

---

## ⚖️ Marcas registradas

SAP y el logo de SAP son marcas registradas de SAP SE. Este proyecto es un portfolio técnico y no está afiliado ni respaldado por SAP.

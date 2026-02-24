# Migración y Validación de Facturas (AP) — Automatización end-to-end (Excel VBA + Python)

![Excel](https://img.shields.io/badge/Excel-VBA-217346)
![Python](https://img.shields.io/badge/Python-Automation-3776AB)
![Power%20Query](https://img.shields.io/badge/Power%20Query-PDF%20Parsing-00A1E0)
![Selenium](https://img.shields.io/badge/Selenium-Headless-43B02A)
![SAP](https://img.shields.io/badge/SAP-GUI%20Scripting-0FAAFF)
![Portfolio](https://img.shields.io/badge/Portfolio-Safe%20Release-6E56CF)

Este repositorio presenta una automatización real (portfolio-safe) de **Cuentas a Pagar (AP)**: orquesta el ingreso de **facturas en PDF**, extrae y normaliza datos, cruza contra un portal interno (`RetailWeb`/`RW`), aplica **validaciones fiscales y de negocio**, y prepara/verifica la contabilización y el seguimiento en SAP.

![Pipeline](assets/demo2.png)

> ⚠️ Versión pública: prioriza **arquitectura + reglas determinísticas + evidencia reproducible** sin exponer datos sensibles. Las integraciones reales (SAP/Outlook/portal) requieren entorno corporativo.

---

![Impacto](assets/highlights/impacto.png) 

![UX/UI](assets/highlights/ux.png)

![Seguridad](assets/highlights/seguridad.png) 

---

## 🔗 Accesos rápidos

- 📄 **Deck (PDF, 7 slides):** [Migracion-Facturas-AP.pdf](assets/deck/Migracion-Facturas-AP.pdf)
- 🖥️ **UI principal (capturas):** [TOOL_UI_OVERVIEW.md](docs/TOOL_UI_OVERVIEW.md)
- 🧩 **UserForms (galería):** [UI_FORMS_GALLERY.md](docs/UI_FORMS_GALLERY.md)
- 📁 **Descripción General:** [PORTFOLIO_OVERVIEW.md](docs/PORTFOLIO_OVERVIEW.md)
- 📚 **Case Study técnico:** [CASE_STUDY.md](docs/CASE_STUDY.md)
- 🧪 **Cómo correr evidencia:** [TESTING.md](docs/TESTING.md)
- 🗺️ **Mapa del repo:** [MANIFEST.md](docs/MANIFEST.md)
- 🧩 **Parsers PDF (VendorXX):** [PARSERS.md](docs/PARSERS.md)

---

## 🧠 Qué demuestra

- **Automatización end-to-end** de un flujo AP completo: documentos → extracción → cruce → validación → decisión operativa → preparación/chequeo.
- **Integración de sistemas heterogéneos**: Excel/VBA + Power Query (PDF) + scripts Python + automatización web + SAP GUI Scripting.
- **Motor de decisiones trazable**: estados determinísticos por fila + mensajes consistentes + tolerancias configurables.
- **Diseño operator-centric**: UX interna (UserForms), controles, progreso, rollback/seguridad, reducción de intervención manual.
- **Evidencia reproducible sin VPN**: tests headless del core + scans de prepublicación + export del paquete público.

---

## 🧭 Flujo end-to-end

```mermaid
flowchart LR
  %% ===== Etapas =====
  subgraph S1[Ingesta de documentos]
    A1[Outlook / Adjuntos PDF] --> A2[Organización / Normalización]
    A2 --> A3[Importación masiva en Excel]
  end

  subgraph S2[Extracción & Parsing]
    B1[Lectura PDF por página<br/>(Power Query)] --> B2[Identificación proveedor<br/>(CUIT / heurísticas)]
    B2 --> B3[Parser por proveedor<br/>VendorXX]
    B3 --> B4[Tabla operativa<br/>filas estructuradas]
  end

  subgraph S3[Enriquecimiento & Cruce]
    C1[Reporte RW (headless)<br/>o Cubo (según config)] --> C2[Match por referencia/remito<br/>+ reglas especiales]
    C2 --> C3[Campos RW enriquecidos<br/>(pago/anulado/scan/fechas/totales)]
  end

  subgraph S4[Validaciones & Decisión]
    D1[Integridad fiscal<br/>(totales vs componentes)] --> D2[Tolerancias configurables]
    D2 --> D3[Reglas determinísticas<br/>(fiscal + negocio)]
    D3 --> D4[Estado por fila + comentarios]
  end

  subgraph S5[Salida / Evidencia]
    E1[Reporte filtrado<br/>(pendientes / revisar)] --> E2[Acción/seguimiento en SAP<br/>(según entorno)]
    E3[Tests headless + scan + export<br/>(sin VPN)] --> E4[dist/public_release]
  end

  %% ===== Conexiones entre etapas =====
  S1 --> S2 --> S3 --> S4 --> S5
  D4 --> E1
  D4 --> E2
  D4 --> E3
```

---

## ✅ Capacidades (resumen)

### Ingesta y preparación

- Importación masiva de PDFs (incluye multipágina).
- Normalización de referencias y metadatos operativos (sucursal/site, fechas, tipo de doc).

### Extracción (PDF → datos)

- Lectura por página con Power Query.
- Parsers anonimizados por proveedor (`Vendor01`..`VendorNN`) bajo contrato común.

### Cruce y enriquecimiento

- Cruce contra RetailWeb/RW (descarga y carga de reporte / o cubo según configuración).
- Match por referencia/remito con reglas especiales (FC/NC y variantes).

### Validaciones y decisión operativa

- Cálculo de integridad fiscal (totales vs componentes: IVA/II/percepciones).
- Aplicación de tolerancias y reglas determinísticas (estado + comentarios).

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

## 🧱 Arquitectura

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

## 🖼️ Material visual

- 📄 **Deck (PDF):** `assets/deck/Migracion-Facturas-AP.pdf`
- 🧩 **Imágenes / Slides:** `assets/images/`

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
# 📁 Portfolio Overview

## 🧠 Qué hace la aplicación

Aplicación VBA que orquesta la migración de facturas desde correo/PDFs hacia una tabla operativa en Excel, aplica validaciones de negocio y prepara la contabilización en SAP.

Integra:

- Entrada de documentos (Outlook + PDFs)
- Lectura/normalización de datos
- Consulta de portal web interno (`RetailWeb`)
- Verificación/contabilización en SAP GUI
- Soporte con scripts Python

---

## 🔄 Flujo operativo (resumen)

```
Outlook → PDFs → Validaciones → RetailWeb (RW) → SAP → Estado por fila / Comentarios / Trazabilidad
```

---

## ✅ Capacidades que muestra el repo

- Orquestación de procesos con múltiples dependencias.
- Reglas de negocio y tolerancias convertidas en lógica repetible.
- UI para operadores (UserForms) orientada a reducir errores.
- Manejo de estados y mensajes por fila para trazabilidad.
- Refactor incremental de core testeable + adapters con side-effects.
- Tooling de publicación segura (scan + export + checklist).

---

## 🧱 Arquitectura pública (para evaluación técnica)

| Módulo | Descripción |
|---|---|
| `src/vba/core/` | Motor de decisiones y utilidades determinísticas. |
| `src/vba/tests/` | Pruebas del core y runner headless. |
| `src/vba/excel/`, `src/vba/ui/`, `src/vba/integrations/` | Orquestación y side-effects. |
| `src/vba/parsers/` | Parsing de PDFs anonimizados por `VendorXX`. |
| `tools/` | Automatización de testing y release público. |

---

## 🧪 Cómo se demuestra sin entorno corporativo

- Tests headless del core → `tools/run_core_tests.ps1`
- Logs reproducibles → `artifacts/`
- Scan de redacción → `tools/prepublish_scan.ps1`
- Export del repo público → `tools/export_public_release.ps1`

---

## ⚠️ Limitaciones conocidas (y cómo se compensan en el portfolio)

- No se ejecutan SAP / Outlook / RetailWeb reales sin VPN y credenciales.
- Los `UserForms` se versionan como `.frm` (texto), sin `.frx` binarios.
- La UI se muestra con capturas mock/sanitizadas en el PDF/slides.

---

## 🔐 Terminología pública (obligatoria)

| Término | Significado |
|---|---|
| `RetailWeb` / `RW` | Portal web interno de retail. |
| `Sucursal` | Término usado en lugar de nombres internos legacy para sitio/negocio. |
| `VendorXX` | Parser anónimo por proveedor. |
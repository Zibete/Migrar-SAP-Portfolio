# 📚 Case Study (Resumen técnico)

## 📈 Impacto

Esta automatización permitió **multiplicar la productividad (~x8)** en un flujo crítico de Cuentas a Pagar: de un proceso manual y repetitivo a un **pipeline end-to-end** con controles, validaciones y trazabilidad.  
El resultado fue **más volumen procesado por día**, **menos errores operativos** y **menos tiempo de revisión manual**, manteniendo el criterio del negocio.

> En la versión pública no se incluyen datos corporativos ni se ejecutan integraciones reales (VPN/credenciales). La evidencia se centra en **reglas, decisiones y tooling reproducible**.

---

## ⚠️ Problema (antes)

El flujo resolvía una necesidad operativa crítica: **validar y preparar facturas** para su contabilización en **SAP**, incluyendo:

- Lectura y normalización de datos desde **PDFs** (y herramientas auxiliares).
- Validaciones de negocio (tolerancias, referencias, estados).
- Asistencia al operador mediante **UserForms** y estados por fila.
- Integraciones con **SAP GUI Scripting**, **Outlook** y un portal web interno (**RetailWeb**).

---

## 🚀 Solución (cómo se logró el x8)

La mejora de productividad no vino de "una macro grande", sino de **convertir el proceso en un pipeline controlado**, reduciendo el trabajo humano a excepciones:

**1. Orquestación end-to-end**
- Un flujo guiado por estados: importar → validar → clasificar → generar mensajes/acciones → dejar la fila lista para operar.
- Menos pasos manuales repetidos; el operador decide solo en casos dudosos.

**2. Reglas de negocio codificadas y trazables**
- Validaciones determinísticas (tolerancias, referencias, mismatch, estados).
- Mensajes y comentarios generados con criterio operativo (qué falta, qué corregir, qué anular).

**3. Integración de sistemas (cuando existe el entorno)**
- SAP GUI Scripting + Outlook + portal web (RetailWeb) + scripts Python auxiliares.
- Objetivo: evitar copiar/pegar y navegación manual constante.

**4. Robustez operativa**
- Manejo de errores y timeouts.
- Estandarización de estados por fila (qué está OK, qué revisar, qué bloquear).
- Trazabilidad: cada decisión queda reflejada en el output.

---

## 🧪 Versión pública verificable (qué se puede comprobar sin VPN)

Limitación real: era un sistema **muy dependiente del entorno** (Office, credenciales, VPN y automatizaciones externas). Eso hacía difícil:

- Revisar cambios con bajo riesgo (sin romper el flujo productivo).
- Demostrar calidad técnica fuera del entorno corporativo.
- Compartir un caso técnico verificable para portfolio.

Para portfolio se preparó un paquete **portfolio-safe** que conserva el valor técnico sin exponer información sensible:

- **Core puro** con reglas y decisiones: `src/vba/core/`
- **Tests del core**: `src/vba/tests/`
- **Ejecución headless reproducible** (Excel invisible) vía scripts.
- **Logs** y tooling de release (scan + export sin historial).
- Parsers de PDF anonimizados: `Vendor01`..`VendorNN`.

---

## 🧱 Arquitectura

```text
Entradas (Outlook / PDFs / operador) + Excel
                |
                v
     Orquestación VBA (Excel / UI / Integraciones)
                |
      +---------+---------+
      |                   |
      v                   v
 Core puro            Adapters / side-effects
 (testeable)          (SAP / RetailWeb / Outlook / PQ / Python)
      |
      v
 Estados, mensajes y decisiones operativas por fila
```

---

## 🔬 Evidencia técnica (sin VPN)

```powershell
powershell -ExecutionPolicy Bypass -File tools/run_core_tests.ps1
powershell -ExecutionPolicy Bypass -File tools/prepublish_scan.ps1
powershell -ExecutionPolicy Bypass -File tools/export_public_release.ps1
```

**Artefactos generados:**

| Archivo | Descripción |
|---|---|
| `artifacts/core-tests.txt` | Resumen de tests. |
| `artifacts/core-tests-details.txt` | Detalle de fallos. |
| `artifacts/prepublish-scan.txt` | Scan de seguridad/redacción. |

---

## 💡 Valor que demuestra

- Capacidad para ordenar una base VBA legacy sin reescritura total.
- Criterio para separar reglas de negocio de automatizaciones frágiles.
- Diseño de evidencia reproducible en un entorno históricamente manual.
- Preparación de un paquete público con controles de redacción y release.

---

## 🚧 Límites (explicitados a propósito)

La versión pública no intenta probar ejecución real de:

- SAP GUI Scripting.
- RetailWeb con credenciales corporativas.
- Outlook COM sobre buzones reales.
- Flujos end-to-end dependientes de VPN y sistemas internos.

La evidencia se concentra en arquitectura, decisiones técnicas, pruebas del core y tooling de release.
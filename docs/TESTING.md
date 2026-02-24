# 🧪 Testing (Core VBA headless)

## 🎯 Objetivo

Ejecutar tests unitarios de `src/vba/core/` desde terminal (PowerShell / IDE), con Excel invisible y sin abrir manualmente el workbook legacy.

---

## 📋 Requisitos

1. Windows con Microsoft Excel instalado.
2. Macros habilitadas para el entorno local.
3. Excel con `Trust access to the VBA project object model` habilitado (para construir el harness).

---

## ▶️ Comando principal

Desde la raíz del repo:

```powershell
powershell -ExecutionPolicy Bypass -File tools/run_core_tests.ps1
```

---

## 🔄 Qué hace el flujo

1. Verifica si existe `portfolio\CoreTestsHarness.xlsm`.
2. Si falta (o si `src/vba/core` / `src/vba/tests` cambió), reconstruye el harness.
3. Abre Excel en modo invisible.
4. Ejecuta `RunCoreTests` en el harness.
5. Lee la cantidad de fallos devuelta por el runner VBA.
6. Genera logs y devuelve exit code:
   - `0` → sin fallos.
   - `1` → fallos o error real (COM / macro / build).

---

## 📄 Logs generados

| Archivo | Descripción |
|---|---|
| `artifacts\core-tests.txt` | Resumen de ejecución. |
| `artifacts\core-tests-details.txt` | Detalle de fallos (`FAIL: ...`) y resumen final del runner. |

---

## ✅ Cobertura (qué sí prueba)

- Clasificación de documentos.
- Validaciones puras del core.
- Armado y truncado de strings.
- Mensajes y comentarios automáticos.
- Reglas determinísticas sin side-effects.

---

## 🚫 Fuera de alcance (qué no prueba)

- SAP GUI Scripting real.
- Outlook COM real.
- RetailWeb real.
- Power Query / `Pdf.Tables` end-to-end.
- Flujos con VPN, credenciales o sistemas internos.

---

## 🛠️ Troubleshooting

### Error al importar módulos (VBProject access)

**Síntoma:** el build del harness falla al importar módulos VBA.

**Acción:**

1. Excel → `Archivo > Opciones > Centro de confianza`.
2. `Configuración del Centro de confianza > Configuración de macros`.
3. Activar `Trust access to the VBA project object model`.
4. Volver a correr `tools\run_core_tests.ps1`.

---

### Macros deshabilitadas

**Síntoma:** Excel abre el harness pero no ejecuta `RunCoreTests`.

**Acción:** habilitar macros para el entorno local y volver a correr.

---

### Excel no instalado

**Síntoma:** error al crear `Excel.Application` vía COM.

**Acción:** instalar Microsoft Excel en la PC local.

---

### Excepción COM en PowerShell (HRESULT)

**Síntoma:** `tools\run_core_tests.ps1` termina en `FAIL` con excepción COM.

**Acción:**

1. Revisar `artifacts\core-tests.txt`.
2. Revisar `artifacts\core-tests-details.txt`.
3. Si el archivo de detalle no existe, el problema suele estar antes de ejecutar el runner (macros, compilación del harness, COM).

---

## 🔁 Ciclo recomendado de debugging

1. Ejecutar `tools\run_core_tests.ps1`.
2. Revisar `artifacts\core-tests.txt`.
3. Revisar `artifacts\core-tests-details.txt`.
4. Corregir `core` o tests.
5. Repetir.
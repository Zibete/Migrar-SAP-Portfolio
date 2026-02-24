# 📋 Notes Forms (dependencias `.frx` de UserForms)

## 📝 Resumen
Este repo incluye UserForms exportados como texto (`.frm`).  
En VBA, los `.frm` suelen referenciar un archivo binario asociado (`.frx`) mediante `OleObjectBlob`.

En esta versión pública, los `.frx` se omiten intencionalmente.

---

## 🚫 Motivos de la exclusión
- Los binarios **no aportan** a la revisión de la lógica VBA (handlers, flujo, integración).
- El objetivo del repo público es mostrar **arquitectura, reglas, integraciones y evidencia reproducible**.
- Incluir binarios agrega ruido y aumenta el riesgo de exponer contenido no deseado.

---

## 📊 Estado actual (scan sobre la fuente)

| Ítem | Valor |
|---|---|
| Total de `.frm` | `10` |
| Formularios con `OleObjectBlob` (requieren `.frx`) | `10` |
| Archivos `.frx` versionados en el repo | `0` |

---

## 📂 Formularios detectados en `src/vba/ui/`
- `Configuración.frm`
- `formImportar.frm`
- `FormRutaCarpeta.frm`
- `form_BloqueoB.frm`
- `Paso2.frm`
- `Paso3.frm`
- `Password.frm`
- `Password_SB.frm`
- `ProgressBar.frm`
- `VisualizarScan.frm`

---

## 🔍 Implicancia para quien revisa
Los `.frm` permiten entender:
- Handlers de eventos
- Convenciones de controles
- Pasos del flujo operativo
- Puntos de integración con Excel y sistemas externos

> Nota: Los `.frm` requieren `.frx` para compilar. En este repo público se prioriza la revisión del flujo y la lógica.

---

# 🖥️ Evidencia visual del flujo
➡️ **Galería de formularios:** [UI_FORMS_GALLERY.md](UI_FORMS_GALLERY.md)

Este anexo muestra la UI real con:
- capturas sanitizadas,
- captions por formulario (qué resuelve y cuándo se usa),
- variantes del mismo formulario cuando aplica.
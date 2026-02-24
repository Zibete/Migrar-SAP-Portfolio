# 🧩 Parsers de PDF (anonimizados)

## 🎯 Objetivo

Los parsers de PDF por proveedor fueron anonimizados en la versión pública para evitar exponer nombres comerciales o relaciones internas.

- **Antes:** módulos con nombres explícitos por proveedor.
- **Ahora:** módulos neutrales `modParserVendor01` ... `modParserVendorNN`.

La forma del sistema se conserva; solo cambia la identificación pública.

---

## 🏷️ Convención de nombres

- Archivo: `src/vba/parsers/modParserVendorNN.bas`
- `VB_Name`: `modParserVendorNN`
- Sub principal: `ParseVendorNN(hoja, y, Optional ctx As AppContext)`

> `NN` es un identificador estable dentro del repo público. No existe un mapa a nombres reales.

---

## 📐 Contrato común

Cada parser recibe:

| Parámetro | Descripción |
|---|---|
| `hoja` | Hoja temporal con datos extraídos del PDF (`Pdf.Tables` / Power Query). |
| `y` | Fila destino en la tabla de trabajo. |
| `ctx` *(opcional)* | `AppContext` con tablas, rangos nombrados y configuración. |

**Responsabilidades típicas:**

- Extraer referencia, fecha, importes, CAE/CAEA y otros campos.
- Normalizar texto, fechas y números.
- Escribir resultados en la fila `y`.
- Consultar tablas auxiliares (proveedores, sucursales, configuración) cuando aplica.

---

## 🔀 Despacho de parser

El despachador vive en `src/vba/excel/modImportPdf.bas`.

La selección del parser se realiza por `vendorId` (redactado) y ejecuta `ParseVendorNN` vía `Application.Run`.

---

## 💡 Qué demuestra técnicamente

- Diseño modular por formato de entrada.
- Encapsulamiento de variaciones por proveedor.
- Estrategia de anonimato sin destruir la arquitectura original.
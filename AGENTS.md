# AGENTS.md · Migrar SAP Portfolio

## Propósito y alcance

Este repositorio presenta una versión pública y anonimizada de una automatización histórica de Cuentas a Pagar desarrollada con Excel/VBA, Power Query, Python e integraciones empresariales.

El objetivo es **preservar y documentar la ingeniería de la herramienta original**. No reconstruir las integraciones privadas, modernizarla por preferencia ni alterar su comportamiento para que parezca un producto distinto.

## Fuentes de verdad

1. Antes de actuar, verificar rama, HEAD, `git status` y archivos relevantes.
2. Usar el código y las pruebas actuales para determinar qué comportamiento está implementado; contrastar la documentación cuando corresponda.
3. Consultar según la tarea:
   - `docs/MANIFEST.md`: estructura y responsabilidades.
   - `docs/PARSERS.md`: convenciones y contratos de parsers.
   - `docs/TESTING.md`: pruebas disponibles y requisitos del entorno.
   - `docs/CASE_STUDY.md`: contexto y decisiones técnicas históricas.
4. Si se aporta `BASELINE_001.md`, tratarlo como una fotografía de auditoría y decisiones, no como sustituto del HEAD vigente. No agregarlo al repositorio.
5. `Zibete/Portfolio` y `Zibete/Zibete` son referencias editoriales **solo de lectura** salvo autorización independiente.

## Preservación del legado

- Conservar reglas de negocio, validaciones fiscales, tolerancias, estados, mensajes, secuencias operativas y contratos existentes, salvo autorización explícita para cambiarlos.
- Distinguir un riesgo detectado por inspección estática de un defecto reproducido. No corregir automáticamente decisiones históricas por parecer mejorables.
- Mantener las responsabilidades de `src/vba/core/`, `parsers/`, `excel/`, `bootstrap/`, `integrations/` y `ui/`.
- Respetar la convención `modParserVendorNN` y la firma `ParseVendorNN(hoja, y, Optional ctx As AppContext)`.
- No asumir que todos los parsers son independientes de `gCtx`, `Hoja2` o del workbook original.
- No sustituir *stubs* públicos por código corporativo reconstruido. Evitar frameworks, dependencias o abstracciones nuevas sin necesidad demostrable.
- Priorizar ajustes pequeños, verificables y limitados a lo autorizado. Si un cambio podría alterar el resultado funcional, detenerse y consultar.

## Seguridad y publicación

- Tratar usuarios, proveedores, identificadores, datos fiscales, credenciales, rutas, endpoints, capturas y metadatos como potencialmente sensibles.
- Nunca introducir valores corporativos reales ni configuraciones privadas por defecto. Para accesos privilegiados, preferir configuración explícita y denegación segura si falta.
- No reconstruir datos privados eliminados ni reescribir el historial de Git sin análisis y aprobación específicos.
- Revisar el exportador `tools/export_public_release.ps1` y el escáner `tools/prepublish_scan.ps1` cuando se trabaje en publicación. Un escaneo textual exitoso **no garantiza** anonimización de archivos, imágenes y PDF.
- No ejecutar transacciones ni alterar estados de SAP, Outlook, RetailWeb u otros sistemas empresariales sin autorización expresa.

## Verificación y método de trabajo

1. Definir el alcance autorizado, los contratos involucrados y los riesgos de regresión.
2. Realizar cambios pequeños y revisables; conservar archivos y modificaciones ajenos a la tarea.
3. Ejecutar las comprobaciones pertinentes disponibles y revisar el diff completo, las rutas, los enlaces y los datos potencialmente sensibles.
4. Distinguir pruebas ejecutadas, fallidas y no ejecutables en el entorno. No afirmar éxito por la mera existencia de scripts o pruebas.
5. Entregar archivos afectados, verificaciones reales, limitaciones y pendientes.

El harness de `tools/run_core_tests.ps1` requiere Windows, Microsoft Excel COM, macros y acceso confiable al proyecto VBA; verifica reglas del core, **no** integraciones corporativas completas. No declarar CI verde sin una ejecución comprobada.

Mantener commits atómicos y proporcionales. **No hacer push, merge, release ni publicación sin autorización explícita.**

## Documentación y comunicación

- Conservar la estructura y los recursos visuales existentes salvo indicación contraria.
- Respaldar afirmaciones con fuentes; separar hechos observados, estimaciones, riesgos y decisiones.
- Evitar documentación redundante y enlaces que no aporten evidencia adicional.
- Responder en español de Argentina, de manera directa, profesional y técnicamente precisa.

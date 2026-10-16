# Validación y límites de la implementación

Estado de construcción: 2026-10-16.

Pruebas ejecutadas sobre una copia de la plantilla, con texto ficticio:

- Selección neutral sin pérdida de opciones originales.
- Diez controles académicos conservados por fila nueva.
- Todas las partes del DOCX distintas de document.xml se mantienen con idénticos bytes.
- Repetir ensure_area_period no agrega una fila duplicada.
- Seguimiento del mismo año conservado; borrado rechazado sin new_academic_year.
- Nuevo año permite limpiar la evaluación en una copia, conservando el original.
- Evaluación vacía no produce alerta y Elija un elemento no cuenta como apoyo completo.
- Rechazo de hash fuente incorrecto.
- Corrección dirigida de nombre repartido en runs.
- Once áreas por cuatro períodos: 44 filas, evaluaciones vacías e identificadores de controles únicos.

Actualización de renderizado del 2026-10-16: la ausencia de LibreOffice se resolvió usando Microsoft Word instalado, en modo de solo lectura y mediante exportación nativa a PDF. Se revisaron visualmente las 14 páginas del documento de prueba, incluidos sus controles añadidos. Word reconoció 44 controles; la plantilla original tiene 13 páginas y 34 controles. La entrada conservó su SHA-256. No se ha verificado selección manual de cada desplegable. Consultar [renderizado con Word](renderizado-word.md) para reproducir el procedimiento, los hallazgos heredados de la plantilla y los límites. La prueba visual de este ejemplo no sustituye la revisión de cada PIAR real completo ni la del caso estructural de 44 filas.

La cobertura de pruebas es estructural. No demuestra que cualquier archivo externo tenga la misma estructura ni que las decisiones pedagógicas sean adecuadas. La detección de áreas usa aliases exactos normalizados; un nombre institucional distinto requiere confirmar su equivalencia y añadir alias, no una coincidencia parcial arriesgada. La detección de períodos reconoce encabezados PRIMER/SEGUNDO/TERCER/CUARTO PERIODO; revisar manualmente otros esquemas.

Lecciones de los scripts aportados inicialmente: el auditor confundía placeholders con apoyos completos, podía extraer etiquetas como nombre del estudiante y alertaba por evaluación vacía. El llenador anterior reemplazaba controles al vaciar celdas, añadía períodos sin deduplicar y escribía encabezados en evaluaciones nuevas. No reutilizar esas conductas. Los scripts nuevos no generan el contenido pedagógico: el agente debe leer las referencias, preparar los cambios y comprobar el resultado.

La biblioteca incluye los PDF aportados con duplicados de texto consolidados y aliases en el catálogo, más las guías 21 y 39 descargadas del MEN. El decreto conserva su enlace oficial en curriculo.md; no se guardó un PDF oficial del decreto porque la URL de exportación no devolvió PDF. El archivo de orientaciones de 2017 no lo sustituye.

## Migración a Claude Code (2026-10-16)

La habilidad se construyó originalmente en otra herramienta (Codex, instalada en `C:\Users\carlo\.codex\skills\generador-piar`) y se recreó aquí como skill de Claude Code en `C:\Users\carlo\.claude\skills\generador-piar`, sin depender de la carpeta `Downloads\Piar 2026`. Se migraron SKILL.md, assets, references y scripts; no se copió `agents/openai.yaml` por ser metadato específico de Codex sin función en Claude Code.

Pruebas repetidas en este equipo Windows, no solo heredadas del entorno anterior:

- Dependencias instaladas: lxml 6.1.3, pypdfium2 5.14.0, Pillow 12.1.0 (Python 3.12.4). La instalación con pip requirió `--trusted-host pypi.org --trusted-host files.pythonhosted.org --trusted-host pypi.python.org` por un error de verificación SSL del entorno; documentado por si se repite en otra máquina.
- `scripts/test_piar.py` ejecutado contra una copia de la plantilla: paso completo, mismo resultado que el registrado el 2026-10-16 en el entorno original (opciones, controles, paquete, idempotencia, evaluaciones, hash, nombre repartido, 44 filas académicas).
- `scripts/render_word.ps1` falló la primera vez en este equipo: `Document.Close(0)` y `Application.Quit(0)` lanzaban `MethodException` ("Argumento: '1' debe ser PSReference") y dejaban un proceso WINWORD.exe invisible sin cerrar. Se corrigió envolviendo el argumento en `[ref]` (`$wordDocument.Close([ref]$noSave)`, `$wordApp.Quit([ref]$noSave)`). Verificado sin procesos huérfanos tras el cambio.
- Con el script corregido se exportó a PDF el documento de cobertura (44 filas, once áreas por cuatro períodos): 59 páginas, 454 controles de contenido reconocidos por Word. Se revisaron visualmente la página 1 (Anexo 1 original intacto) y una página intermedia con fila nueva de Ciencias Sociales (desplegables con la opción neutral seleccionada, texto personalizado legible, sin controles rotos).
- Se descargó el texto oficial del Decreto 1421 de 2017 (13 páginas, `references/oficiales/decreto-1421-2017.pdf`) desde el Gestor Normativo de Función Pública; antes solo existía el enlace. Ver [curriculo.md](curriculo.md).
- Las áreas de sexto, séptimo, décimo y undécimo siguen sin confirmación institucional; se dejó una propuesta razonada marcada como tal en curriculo.md, no como dato verificado.

Esta migración no repitió la revisión visual completa de un PIAR real de once áreas y cuatro períodos con contenido pedagógico; esa revisión sigue siendo responsabilidad de cada uso real de la habilidad, conforme al paso 10 de SKILL.md.

# Revisión visual con Microsoft Word en Windows

Método probado el 2026-10-16. Requiere Microsoft Word de escritorio instalado con automatización COM disponible. No requiere LibreOffice. Fuera de Windows o sin Word, resolver otro renderizador disponible; no afirmar que este método funciona en cualquier equipo.

## Procedimiento

1. Seguir primero el flujo Documents del entorno. Si el renderizador habitual falla porque no encuentra soffice.exe, comprobar el diagnóstico. No instalar suites ni reemplazar el Word del usuario sin necesidad.
2. En Windows con Word instalado, ejecutar el script incluido desde PowerShell:

```powershell
& 'RUTA_HABILIDAD/scripts/render_word.ps1' -InputDocx 'RUTA/documento.docx' -OutputPdf 'RUTA_QA/documento.pdf'
```

El script crea una instancia de Word sin ventana visible, desactiva macros en esa instancia, abre el DOCX en modo de solo lectura, repagina y exporta PDF. Cierra su documento y su instancia sin guardar. Compara SHA-256 antes y después para detectar cambios en la entrada. No usar GetActiveObject ni cerrar procesos Word ajenos. Si el entorno bloquea COM con E_ACCESSDENIED, solicitar la ejecución escalada del comando concreto conforme a sus permisos, sin desactivar protecciones.

3. Convertir todas las páginas del PDF a PNG con el Python disponible del entorno:

```text
python scripts/render_pdf_pages.py RUTA_QA/documento.pdf RUTA_QA/paginas --scale 1.5
```

Requiere pypdfium2 y Pillow. El script rechaza sobreescrituras. Abrir cada imagen, no solo una hoja de miniaturas. Buscar texto cortado, solapamientos, filas partidas difíciles de seguir, encabezados huérfanos y espacios excesivos. Si se cambia el DOCX, exportar de nuevo a otro nombre y repetir revisión.

4. Usar el PDF únicamente como comprobación. La entrega sigue siendo el DOCX con controles, sin conversión de vuelta desde PDF. Exportar PDF no verifica por sí solo que todos los desplegables se puedan seleccionar manualmente: combinar inspección OOXML, recuento de ContentControls de Word y, cuando proceda, prueba de interacción en una copia.

## Evidencia y límites de la prueba

- Plantilla original: 13 páginas, 34 controles reconocidos por Word.
- Documento estructural de prueba con Emprendimiento, segundo período: 14 páginas, 44 controles reconocidos por Word. Sus 14 páginas fueron inspeccionadas visualmente.
- Las primeras cuatro páginas renderizadas resultaron idénticas píxel a píxel entre plantilla y prueba a escala 1.4.
- La página 3 casi vacía, con Director de Grado y Docente de Apoyo, ya está en la plantilla original. No atribuirla al generador.
- La tabla académica se extiende por varias páginas. La fila nueva de Emprendimiento continúa en la página siguiente. La plantilla contiene texto vertical, encabezados altos y compromisos repetidos. Conservar el original y evaluar mejoras de paginación sobre la copia, respetando la estructura y los controles.
- El documento de prueba contiene campos vacíos y opciones neutrales porque verifica estructura; no es un PIAR listo de un estudiante real. No extrapolar su paginación a un PIAR completo de once áreas y cuatro períodos. Ese caso solo tiene prueba estructural hasta que se genere y revise su contenido real.

Referencia técnica oficial: [Document.ExportAsFixedFormat](https://learn.microsoft.com/en-us/office/vba/api/word.document.exportasfixedformat). La exportación usa el formato PDF, valor 17. La solución fue verificada localmente; no depende de la carpeta de la conversación original.

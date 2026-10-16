# Edición del Word sin perder controles

Los scripts trabajan solo con `word/document.xml`; preservan los demás componentes del paquete sin cambios de contenido. Revisar encabezados, pies, comentarios y cuadros de texto por separado si contienen datos antiguos. No afirmar que se han actualizado si no se inspeccionaron.

## Inventario y auditoría

```text
python scripts/inspect_piar.py entrada.docx --json inventario.json
python scripts/audit_piar.py entrada.docx --grado 8 --periodos 1 2 --json auditoria.json
```

El inventario incluye hash del archivo, XPath de cada control/párrafo/celda, texto esperado y hash estructural. No adivinar índices: leer el inventario real. Para vacíos del anexo 1 y datos del anexo 3, inspeccionar la celda y sus párrafos, preservando etiquetas. Para campos compuestos, usar `text_node` con XPath y hash del nodo exacto; no sustituir todo el párrafo que contiene otros campos. No generar datos que falten.

## Cambios

```json
{
  "source_sha256": "HASH_REAL_DEL_INVENTARIO",
  "operations": [
    {"op": "select", "xpath": "XPATH_REAL_CONTROL", "expected_hash": "HASH_REAL", "value": "OPCION_LITERAL_EXISTENTE"},
    {"op": "select", "xpath": "OTRO_CONTROL", "expected_hash": "HASH_REAL", "value": "Información insuficiente para determinar esta barrera", "add_neutral": true},
    {"op": "append", "xpath": "XPATH_REAL_CELDA", "expected_hash": "HASH_REAL", "lines": ["Metodología: explicación personalizada vinculada con el objetivo y la evidencia disponible."]},
    {"op": "paragraph", "xpath": "XPATH_REAL_PARRAFO", "expected_text": "texto previo completo", "value": "texto corregido"}
  ],
  "ensure_area_period": [{"area": "Emprendimiento", "periodo": 2}]
}
```

Los valores anteriores son un esquema, no un JSON ejecutable sin inspección. Cada operación requiere texto esperado o hash del elemento real. El hash fuente debe corresponder exactamente al documento de entrada. No reutilizar inventarios después de una edición.

`ensure_area_period` no redacta contenido: si existe área/período no lo duplica; si falta, clona la fila vacía de la plantilla incluida, conserva sus desplegables/opciones y les asigna una neutral. Deja evaluación vacía. Agrega una fila del mismo formato por área/período solicitado. Después se debe inspeccionar y completar el contenido de las filas nuevas. Las filas existentes con varios períodos se editan de forma localizada con `operations`, sin reescribir períodos correctos. Una fila sin período exige identificarlo primero.

En las filas originales, `select` conserva el control. El texto añadido va fuera del control, en su misma celda. `paragraph` permite corregir texto repartido en varios runs, pero rechaza párrafos con controles, imágenes o campos. Nunca vaciar una celda completa para resolver un desplegable. Para columnas corridas, trasladar párrafos concretos solo tras inspeccionar su contenido; ampliar el script o usar OOXML conservador si las operaciones disponibles no cubren el caso. No copiar a ciegas una celda con controles ni reducirla a texto plano.

Para actualizar año, incluir `"new_academic_year": true` y operaciones `reset_evaluation` en las celdas de evaluación pertinentes, cada una con XPath y hash. Admite filas académicas y las cinco dimensiones de Otras. Guardar el original y borrar solo lo histórico en la copia del año nuevo. No cambiar globalmente números de año, edades o nombres.

```text
python scripts/fill_piar.py entrada.docx cambios.json salida-nueva.docx
```

No sobrescribe el original ni una salida existente. Si falla, revisar el diagnóstico y volver a inspeccionar; no quitar las validaciones para forzar una edición.

## Verificación

1. Auditar la salida y revisar períodos solicitados, áreas, objetivos, cinco apoyos y opciones seleccionadas.
2. Comparar los controles originales: permanecen y conservan las opciones originales. Los nuevos deben tener identificadores propios. La opción visible debe pertenecer a la lista.
3. Revisar identidad y posibles nombres ajenos en contexto; el auditor no identifica personas por sí solo.
4. Evaluaciones nuevas vacías; históricos del mismo año conservados; históricos del año anterior no presentados como nuevos.
5. Renderizar y revisar todas las páginas con Documents. Si falta LibreOffice en Windows, seguir [renderizado con Word](renderizado-word.md). La validación XML no acredita maquetación correcta ni interacción manual comprobada en Word.
6. Entregar el DOCX editado sin pasar por una conversión de ida y vuelta desde PDF/LibreOffice.

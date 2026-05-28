# PathCapacitacion

Proyecto para generar una presentación ejecutiva de **path de capacitación por persona** a partir de matrices en Excel.

La solución cruza:

- tecnologías usadas por aplicaciones del banco,
- tecnologías conocidas por cada persona,
- categorías tecnológicas,

y produce una presentación PowerPoint con una lámina por persona.

---

## ¿Qué hace este proyecto?

1. Lee `AplicacionesBanco.xlsx`.
2. Toma la hoja `Matriz` para identificar tecnologías y uso en aplicaciones.
3. Toma la hoja `MtzPerson` para identificar conocimientos por persona.
4. Clasifica tecnologías por estado:
   - **Tecnologías que ya conoce** (verde)
   - **Tecnología conocida pero no identificada en banco** (gris)
   - **Tecnologías a aprender** (rojo)
5. Genera `PathCapacitacion.pptx` con resumen visual por persona.

---

## Estructura del repositorio

- `varios.py`: script principal de generación de presentación.
- `propuesta_capacitacion_ppt.py`: genera propuesta de capacitación personalizada en formato presentación (PPT).
- `AplicacionesBanco.xlsx`: fuente principal de datos.
- `PathCapacitacion.pptx`: salida generada (presentación).
- `PropuestaCapacitacion.pptx`: salida generada del plan de capacitación en presentación.
- `externos.py`: script auxiliar para procesar externos desde TXT a XLSX.
- `externos_santander.txt`: entrada de externos por empresa.
- `externos_santander.xlsx`: salida generada del script de externos.
- `.gitignore`: exclusiones de archivos temporales y entorno.
- `requirements.txt`: dependencias de Python.

---

## Requisitos

- Python 3.10+
- Entorno virtual recomendado (`.venv`)
- Dependencias:
  - `pandas`
  - `openpyxl`
  - `python-pptx`

---

## Instalación

1. Crear/activar entorno virtual.
2. Instalar dependencias desde `requirements.txt`.
3. Verificar que `AplicacionesBanco.xlsx` esté en la raíz del proyecto.

---

## Ejecución

### 1) Generar presentación principal

Ejecutar el script `varios.py`.

Resultado esperado:

- `PathCapacitacion.pptx`
- Si el archivo está abierto y bloqueado por PowerPoint, el script guarda en:
  - `PathCapacitacion_actualizado.pptx`

### 2) Generar listado de externos (opcional)

Ejecutar `externos.py` para convertir `externos_santander.txt` a `externos_santander.xlsx`.

### 3) Generar propuesta de capacitación en presentación (PPT)

Ejecutar `propuesta_capacitacion_ppt.py`.

Resultado esperado:

- `PropuestaCapacitacion.pptx`
  - Portada general.
  - Una lámina por persona.
  - Tres columnas por fase (Quick Wins, Consolidación, Expansión).
  - Bloque de "Base conocida actual" por persona para validar alineación con la sugerencia.
  - Semáforo de afinidad visible en recomendaciones/foco inmediato.
  - Foco inmediato recomendado al pie de la lámina.

El plan se construye tomando en cuenta:

- demanda de tecnología en aplicaciones del banco,
- afinidad con conocimientos ya existentes en cada persona,
- progresión de aprendizaje en 3 fases:
  - Fase 1 (Quick Wins)
  - Fase 2 (Consolidación)
  - Fase 3 (Expansión)

Si el archivo de salida está abierto y bloqueado por PowerPoint, el script guarda en:

- `PropuestaCapacitacion_actualizada.pptx`

---

## Modelo de datos esperado (Excel)

### Hoja `Matriz`

- Fila de categorías (encabezado superior).
- Fila de tecnologías.
- Filas por aplicación.
- Valores binarios esperados por tecnología (0/1).

### Hoja `MtzPerson`

- Filas por persona.
- Columnas de tecnologías (mismo universo/intersección con `Matriz`).
- Valores binarios esperados por tecnología (0/1).

> Nota: el script usa la intersección de columnas entre `Matriz` y `MtzPerson`.

---

## Lógica de presentación (resumen)

En cada lámina de persona:

- **Encabezado** con conteos:
  - conocidas
  - por aprender
  - conocidas no identificadas en banco
  - total relevante

- **Columna izquierda**:
  - ✅ Tecnologías que ya conoce
  - ℹ️ Tecnología conocida pero no identificada en banco

- **Columna derecha**:
  - ❌ Tecnologías a aprender
  - agrupadas por categoría con encabezados destacados

- **Barra de progreso**:
  - porcentaje de cobertura respecto del universo de tecnologías relevantes en banco.

---

## Troubleshooting

### El PPT no se sobreescribe

Si `PathCapacitacion.pptx` está abierto en PowerPoint, Windows lo bloquea.

Comportamiento implementado:

- El script captura `PermissionError` y guarda automáticamente en
  `PathCapacitacion_actualizado.pptx`.

### Una tecnología no aparece

Revisar:

1. Que exista en ambas hojas (`Matriz` y `MtzPerson`).
2. Que esté en la intersección de columnas.
3. Si está en `MtzPerson` pero no en uso en `Matriz`, se verá en la sección gris.

### Texto difícil de leer en brechas

La sección roja usa formato compacto con:

- agrupación por categoría,
- jerarquía visual (categoría + ítems),
- ajuste dinámico de columnas/tamaño.

---

## Mantenimiento y buenas prácticas

- Mantener nombres de tecnologías consistentes entre hojas.
- Evitar columnas duplicadas con variaciones de espacios.
- No versionar temporales de Office (`~$*.xlsx`, `~$*.pptx`).
- Regenerar el PPT después de cualquier cambio de fuente.

---

## Versionado

- Rama principal: `main`
- Tag de referencia publicado: `v1.0.0`
- Release disponible en GitHub Releases.

---

## Autoría

Repositorio: `vperezguzman66/PatchCapacitacion`

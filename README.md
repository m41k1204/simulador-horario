# Armador de horario · UTEC 2026-2

Web local para armar el horario cruzando la oferta del semestre con la carga hábil.

## Uso

Abrir **`horario.html`** con doble clic. Es un solo archivo, con los datos adentro:
no necesita servidor ni internet ni ningún archivo vecino.

- Panel izquierdo: cursos filtrados, obligatorios primero ordenados por ciclo.
- Clic en un curso → sus opciones de matrícula (sección + grupo de laboratorio).
  Pasar el mouse por una opción la dibuja en gris sobre la grilla antes de elegirla.
- Las opciones que chocan con lo ya elegido salen en rojo; los cruces en la grilla
  quedan con borde rojo y se listan en la barra superior.
- La selección se guarda sola en el navegador (localStorage).

## Archivos

| Archivo | Qué es |
|---|---|
| `horario.html` | **La app lista para usar**, generada. Un solo archivo, ábrelo con doble clic. |
| `carga_habil.tsv` | Los cursos que aún faltan llevar: código, tipo, ciclo, créditos. **Editar aquí.** |
| `build_data.py` | Cruza el Excel con `carga_habil.tsv` y genera `courses.json`, `data.js` y `horario.html` |
| `index.html` | El código fuente de la app; carga `data.js` aparte. Editar aquí y regenerar. |

## Regenerar los datos

Con un Excel nuevo de Consulta de Horario (se toma el `Consulta_Horario-*.xlsx` de
la carpeta padre) o después de editar `carga_habil.tsv`:

```sh
pip install openpyxl        # única dependencia
python build_data.py        # o: python build_data.py ruta/al/excel.xlsx
```

## Notas sobre los datos

- Un curso se lleva como *sección* (teoría fija) + *un* grupo de laboratorio de esa
  sección. La app trata cada combinación sección+lab como una opción.
- La frecuencia `Semana A` / `Semana B` se respeta: dos sesiones en semanas
  distintas no cuentan como cruce.
- Modalidad: cada opción trae una etiqueta (Presencial / Sincrónico / Mixto), las
  sesiones virtuales salen marcadas en morado en la lista y rayadas en la grilla,
  y cada bloque dice `Presencial · aula` o `Sincrónico`.
- 95 secciones del Excel vienen sin docente; se muestran como "Docente por definir".
- `carga_habil.tsv` está pegada hasta `MT5310`; si faltan cursos al final de la lista
  (MT53xx en adelante), agregarlos ahí y volver a correr `build_data.py`.

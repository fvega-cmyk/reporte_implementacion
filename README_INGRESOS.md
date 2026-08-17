# Reporte Semanal de Ingresos (Activación)

Flujo **paralelo** al de Implementación, en el mismo repo. No se modificó ningún
archivo del flujo que ya está en producción: los nuevos importan lo que
necesitan (`google_clients`, `utils`, los helpers de slides de `generar_ppt.py`).

## Qué hace

1. Lee la hoja "Actividades".
2. Deja solo las filas con **ACTIVIDAD = `Ingreso`** y **TIPO ACTIVIDAD = `Activación`**.
3. Se queda con las que caen en una semana ISO (lunes a domingo).
4. Por cada campaña genera **un PPT** (sin Excel) usando `Template_Reporte_Ingresos`.
5. Crea la carpeta `SEMANA 33 - 2026` en la Unidad Compartida y guarda ahí todos
   los PPT de esa semana.
6. Manda **un solo correo** con todos los PPT de la semana adjuntos.

## Archivos nuevos

```
src/
├── semana.py                    ← semana ISO, rangos y etiqueta "SEMANA 33 - 2026"
├── config_ingresos.py           ← toda la config de este flujo
├── leer_sheets_ingresos.py      ← filtro Ingreso + Activación por semana
├── generar_ppt_ingresos.py      ← PPT desde Template_Reporte_Ingresos
├── subir_a_drive_ingresos.py    ← carpeta por semana + subir/actualizar
├── enviar_correo_ingresos.py    ← 1 correo con N adjuntos
└── main_ingresos.py             ← orquestador
.github/workflows/
├── ingresos_actualizar.yml      ← genera y sube (sin correo)
└── ingresos_notificar.yml       ← genera, sube y manda el correo
```

## Semana ISO, no semana calendario

`SEMANA 33 - 2026` sale de `date.isocalendar()`. Coincide con la numeración que
ya usan: el 12/08/2026 es la semana 33.

Ojo con el borde de año, que es donde estas cosas se rompen: el **domingo
03/01/2027 pertenece a la `SEMANA 53 - 2026`**, no a la semana 1 de 2027. Por eso
se usa el **año ISO** y no el año calendario — si no, ese domingo caería en una
carpeta `SEMANA 53 - 2027` inexistente y la semana quedaría partida en dos.

## Setup (una sola vez)

### 1. Compartir el template con la cuenta de servicio

La carpeta `Templates_Reportes` está en **Mi unidad**. Compartila (o al menos el
archivo `Template_Reporte_Ingresos`) con el email de la cuenta de servicio, con
permiso de **Lector**. Solo se lee.

Si querés saltarte la búsqueda por nombre, pegá el ID del template en
`config_ingresos.py` → `TEMPLATE_PPT_INGRESOS_ID`. Lo saca de la URL:
`docs.google.com/presentation/d/ESTE_ES_EL_ID/edit`.

### 2. Unidad Compartida de destino

Las carpetas de semana se crean en la Unidad Compartida
`0ABVN4w8aqhhqUk9PVA` (ya está como default en `config_ingresos.py`).

> **Por qué no puede ir en Mi unidad**: la cuenta de servicio no tiene cuota de
> almacenamiento propia. Todo lo que cree en "Mi unidad" falla con
> *storage quota exceeded*. En una Unidad Compartida el dueño es la unidad, no
> la cuenta de servicio, así que funciona. Es el mismo motivo por el que los
> reportes de Implementación viven en una Unidad Compartida.

Compartí esa unidad con la cuenta de servicio como **Administrador de contenido**
(necesita crear carpetas, no solo archivos).

### 3. Secreto de GitHub

Agregá **un** secreto nuevo (los demás ya existen):

| Secreto | Valor |
|---|---|
| `DRIVE_CARPETA_INGRESOS_ID` | `0ABVN4w8aqhhqUk9PVA` |

Reutiliza `GOOGLE_CREDENTIALS_JSON`, `GMAIL_USER` y `GMAIL_APP_PASS`.

### 4. Destinatarios del correo

En `config_ingresos.py` → `CORREOS_INGRESOS`. Ahora está solo tu correo; agregá
al resto del equipo.

### 5. cron-job.org

| Job | Método | Event type | Cuándo |
|---|---|---|---|
| Actualizar ingresos | `repository_dispatch` | `ingresos_actualizar` | lu–sá, 12:00 y 18:00 Chile |
| Notificar ingresos | `repository_dispatch` | `ingresos_notificar` | viernes, 17:00 Chile |

Mismo endpoint y mismo PAT que los otros jobs:

```
POST https://api.github.com/repos/fvega-cmyk/reporte_implementacion/dispatches
Body: {"event_type":"ingresos_notificar"}
```

El notificador **no tiene cron nativo de GitHub** a propósito: el cron de GitHub
se dispara con retrasos y a veces dos veces, y acá un disparo duplicado
significa un correo duplicado al equipo entero.

## Placeholders del template

Se reemplaza lo que exista y se ignora lo que no, así que podés ajustar el
template sin tocar el código.

**Slide 1 (portada)**
`[CAMPAÑA]` `[CLIENTE]` `[SEMANA]` `[N SEMANA]` `[AÑO]` `[PERIODO]`
`[FECHA ENVIO DEL REPORTE]` `[TOTAL SALAS]` `[TOTAL REGISTROS]`

**Slide 2 (detalle, se duplica una vez por fila)**
`[CÓD.] - [NOMBRE SALA]` · `[DIRECCIÓN], [COMUNA] - [REGION]` · `[CADENA]`
`[MARCA]` `[CATEGORIA]` `[SKU]` `[MATERIAL]` `[CANTIDAD]` `[GUIA]`
`[PROCESO]` `[DETALLE]` `[OBSERVACIONES]` `[EJECUTIVO]` `[SEMANA]`
`[FECHA INICIO]` `[FECHA TERMINO]` `[FECHA ENTREGA]` `[FECHA COMPROMISO]`
`[HORA INICIO]` `[HORA TERMINO]` · `[FOTO 1]` … `[FOTO 4]`

Reglas heredadas de Implementación: las slides se ordenan por código numérico
real de sala (`J9` antes que `J633`), `Reagenda interna/externa` se normaliza a
`Reagenda`, y las fotos se insertan respetando proporción y centradas en el área
de los placeholders (con 1, 2, 3 o 4 fotos se recalcula la grilla).

## Uso manual

```bash
# semana en curso, sin correo
python src/main_ingresos.py

# semana en curso + correo
python src/main_ingresos.py --enviar

# la semana que ya cerró
python src/main_ingresos.py --enviar --semana-anterior

# regenerar una semana puntual
python src/main_ingresos.py --semana 30 --anio 2026

# prueba a un solo correo, una sola campaña
python src/main_ingresos.py --enviar --campana "Ingresos Verano 2026" \
  --destinatarios "fvega@sell-out.cl"
```

Desde GitHub: **Actions → "Ingresos - Notificar (1 correo semanal)" → Run
workflow**. Ahí tenés los mismos parámetros en el formulario.

## Detalles de implementación

- **Idempotente**: si el PPT de una campaña ya existe en la carpeta de la
  semana, se actualiza su contenido con `files().update()`. Mismo file ID, mismo
  link. El actualizador puede correr N veces por semana sin duplicar nada ni
  romper links ya enviados.
- **Peso del correo**: Gmail corta en 25 MB. Se adjunta en orden alfabético
  hasta 22 MB; los PPT que no entran van solo por link y el correo lo dice
  explícitamente.
- **Solape de fechas**: una activación de jueves a martes entra en las dos
  semanas. Es a propósito: si estuvo viva en la semana, corresponde informarla.
- **Filas sin fecha**: se omiten y se avisa en el log (no se pueden ubicar en
  ninguna semana).
- **Canceladas**: las campañas 100% en Cancelado no se notifican, igual que en
  Implementación. Se puede desactivar con `--incluir-canceladas`.

## Primera prueba recomendada

1. Compartí el template y la Unidad Compartida con la cuenta de servicio.
2. Cargá `DRIVE_CARPETA_INGRESOS_ID` en los secretos.
3. Corré **"Ingresos - Actualizar PPT"** a mano (sin correo) y revisá que la
   carpeta `SEMANA 33 - 2026` aparezca con los PPT dentro.
4. Abrí un PPT: si quedó algún `[TOKEN]` sin reemplazar, avisame con el nombre
   exacto y lo agrego a `_tokens_fila()`.
5. Recién entonces corré **"Ingresos - Notificar"** con `destinatarios` = solo
   tu correo.
6. Si sale bien, sacá el override y armá el job en cron-job.org.

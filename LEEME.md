# Generador de la Orden de Servicio del CAT

Toma las hojas `PREORDEN` / `PREORDEN FINDE` del **SISTEMA DE PREORDEN.xlsx** y arma
el Word final con todos los anexos.

**La orden semanal es una sola** e incluye, en este orden:

1. Un anexo por cada **base operativa** (Alejandra Beretta, Cochabamba, Ocampo,
   Vedia, Cinthia Choque)
2. Un anexo por cada **zona comunal** (Centro, Norte, Sur)

El índice se arma solo con todos los anexos y **cada renglón es clickeable**:
salta directo al anexo, tanto en el Word como en el PDF.

**La orden de fin de semana va aparte** (`--finde`), con su propia plantilla y
**solo con bases operativas**: las comunas no trabajan los fines de semana. Si
aparece un servicio de finde cargado con base de despliegue comunal, el
generador avisa porque es un error de carga.

Los anexos que quedan sin ningún servicio validado no se generan ni aparecen en
el índice.

---

## Uso semanal (3 pasos)

### 1. Bajar la planilla
Desde Excel Online: **Archivo → Guardar como → Descargar una copia**.
Dejá el archivo en la carpeta **`planilla\`**, acá adentro del proyecto.

El nombre puede tener sufijos: `SISTEMA DE PREORDEN (1).xlsx` sirve igual. Si
hay varios, usa el más reciente y avisa cuál eligió.

> Si en `planilla\` no hay nada, también busca en `Descargas`. Y con `--xlsx`
> se le pasa cualquier ruta.

> El libro sigue online para que CGM, Planeamiento, CEF y Comunas carguen en
> simultáneo. El generador trabaja siempre sobre la copia bajada.

### 2. Actualizar el texto legal
Abrí la plantilla que corresponda y editá lo que cambie esta semana: número de
orden, fechas, o cualquier párrafo.

| Plantilla | Para |
|---|---|
| `plantilla_OS_SEMANA.docx` | La orden semanal (bases + comunas) |
| `plantilla_OS_FINDE.docx` | La orden de fin de semana (solo bases) |

**Hay dos párrafos que no hay que tocar:**

| Marcador | Qué pone ahí el generador |
|---|---|
| `{{INDICE}}` | La tabla de anexos, con los links |
| `{{ANEXOS}}` | Todos los anexos |

Si se borra alguno, el generador avisa y no genera.

### 3. Generar

```bash
python generar_orden.py
```

Salen las dos órdenes en la carpeta `salida\`, cada una en Word y en PDF:

- `Orden de servicio SEMANA.docx` + `ODS_2006-2026-O.pdf`
- `Orden de servicio FINDE.docx` + `ODS_2006-2026-F.pdf`

El nombre del PDF sale del N° de orden que tenga el título de cada plantilla.
Ya vienen con el anexo de imágenes adentro. Listo para mandar.

---

## El anexo de imágenes

Hay un Word por orden que se completa a mano durante la semana:

| Archivo | Va en |
|---|---|
| `anexo_imagenes_SEMANA.docx` | La orden semanal |
| `anexo_imagenes_FINDE.docx` | La de fin de semana |

Se abren en Word y se pegan las imágenes abajo de los dos párrafos grises. Al
generar, **todo lo que haya en ese archivo** —fotos, textos, tablas— se agrega
al final de la orden, con su entrada clickeable en el índice.

Detalles:

- Los párrafos que contienen `{{INSTRUCCIONES}}` no se copian. Son los dos
  grises del principio: sirven para no perder de vista para qué es el archivo.
- El título "ANEXO DE IMÁGENES" lo pone el generador, no hay que escribirlo.
- El anexo vuelve a **hoja vertical**, aunque las tablas vayan apaisadas.
- Si el archivo no existe, la orden se genera igual y avisa que faltó.
- Conviene vaciarlo al empezar cada semana, porque si no arrastra las imágenes
  de la anterior.

---

## Variantes

| Comando | Qué genera |
|---|---|
| `python generar_orden.py` | **Las dos órdenes**: la semanal y la de fin de semana |
| `python generar_orden.py --semana` | Sólo la semanal (bases + comunas) |
| `python generar_orden.py --finde` | Sólo la de fin de semana (sólo bases) |
| `python generar_orden.py --ambito bases` | Sólo las bases operativas |
| `python generar_orden.py --ambito comunas` | Sólo las zonas comunales |
| `python generar_orden.py --zona CENTRO` | Un solo anexo, para revisar rápido |
| `python generar_orden.py --vertical` | Anexos verticales en vez de apaisados |
| `python generar_orden.py --xlsx "C:\ruta\otro.xlsx"` | Otra planilla |
| `python generar_orden.py --tam 10` | Letra más grande (por defecto 9pt) |
| `python generar_orden.py --max-desc 250` | Recorte distinto de DESCRIPCIÓN (`0` = sin recorte) |

---

## De dónde sale cada columna del anexo

| Columna del anexo | Origen en la planilla |
|---|---|
| DIA | `DIA` — los rangos se expanden: `MAR A VIE` → `MAR-MIE-JUE-VIE`. Tolera variantes (`MAR A VIER`, `DE LUN A VIE`, `MIERCOLES`) y respeta las fechas sueltas (`MAR 18/08`) |
| TURNO | `TURNO` |
| TAREA | `SERVICIO`. Si `TIPO` = MISION se antepone **MISIÓN** en negrita |
| DESCRIPCIÓN | `FUNCION`, recortada a 400 caracteres |
| OBSERVACIONES | `OBSERVACIONES` |
| DIRECCIÓN | `UBICACION` + `ALTURA` / `CALLE 2` / `CALLE 3` |
| HORA | `HORA`, tal como se cargó (ver abajo) |
| AGENTES POR TURNO | `AT TM` / `AT TT` / `AT TIN` / `AT TN` en una columna cada uno |
| BASE | `BASE`, abreviada (ver abajo) |

### La columna HORA

Sale **tal como se cargó en la planilla**: `8 A 19HS`, `19 A 23 HS`,
`7 A 12HS / 12 A 17HS`. El generador no la reescribe.

La excepción son las celdas donde Excel guardó una hora de verdad en vez de
texto. Excel las almacena como fracción del día (`0.5` = mediodía) y el
`FILTRAR()` las derrama a `PREORDEN` perdiendo el formato, así que llegan como
un número largo. Esas se muestran como `HH:MM`, porque el valor crudo no se
entiende.

**Se ven en el control de calidad**, con fila e ID. Conviene arreglarlas en la
planilla escribiéndolas como texto (`8 A 19HS`), porque cargadas como hora sólo
guardan un horario y se pierde el rango.

Si alguna vez se quiere volver a normalizar todo a `08:00 A 19:00`, es
`NORMALIZAR_HORA = True`.

### Abreviaturas de la columna BASE

La columna BASE se repite en cada fila, así que va abreviada. **El título del
anexo y el índice siguen usando el nombre completo.**

| En la planilla | En la tabla |
|---|---|
| `DESPLIEGUE COMUNA 1` … `15` | `BDC 1` … `BDC 15` (Base de Despliegue Comuna) |
| `ALEJANDRA BERETTA` | `A. Beretta` |
| `COCHABAMBA` | `Cocha` |
| `CINTHIA CHOQUE` | `C. Choque` |
| `ARAOZ DE LAMADRID` | `A. Lamadrid` |
| `BRD SARMIENTO` / `BRD TACUARI` | `Brd. Sarmiento` / `Brd. Tacuarí` |
| `OCAMPO`, `VEDIA`, `COUTURE` | igual, ya son cortos |

Se editan en `ABREVIATURAS_BASE`. Una base que no esté en la lista sale con su
nombre completo, así que agregar una base nueva no rompe nada. `ABREVIAR_BASES
= False` vuelve a los nombres largos.

En el finde las columnas de dotación son `FSD S`, `FSD D`, `FSI S`, `FSI D` y `FSN D`.

### Orden de los servicios

Dentro de cada anexo van agrupados **por base**, y dentro de cada base:

1. Primero los que **cubren la semana completa** (lunes a viernes)
2. Después los de días puntuales, **empezando por los lunes**
3. A igual día de arranque, primero el que cubre más días
   (`LUN-MAR-JUE-VIE` antes que `LUN-MIE`)
4. A igualdad de todo lo anterior, por turno y por horario

En los anexos de zona comunal la secuencia de días **se reinicia en cada
comuna**, porque cada una es un bloque aparte: el agente de la BDC 1 encuentra
todo lo suyo junto.

Se cambia con `ORDEN_SERVICIOS`.

Zonas comunales: **CENTRO** 1-3-4-5-6 · **NORTE** 2-12-13-14-15 · **SUR** 7-8-9-10-11

### Detalles del formato

- El texto legal queda **vertical** y los anexos **apaisados**: hay un salto de
  sección en el medio. El encabezado con los logos se conserva en las dos partes.
- La tabla se sale de los márgenes del texto pero **deja 1 cm de aire contra el
  borde de la hoja** (`MARGEN_ANEXO`). Es a propósito: ninguna impresora imprime
  hasta el borde, y con full-bleed se cortaban DIA y BASE.
- El encabezado se repite cuando la tabla corta de página, y ninguna fila se
  parte al medio.

### Densidad: por qué está así

La orden se lee pasando muchas hojas, así que cada pulgada de página cuenta.
Las decisiones tomadas y lo que cuesta cada una, medido sobre los 178 servicios
de comunas:

| Decisión | Costo | Por qué se dejó así |
|---|---|---|
| Banner del encabezado en su tamaño original | **–27 páginas** | Escalarlo al ancho apaisado lo agranda también a lo alto y se come una pulgada por hoja: 66 páginas contra 39. Queda el vacío blanco a la derecha, y está bien que quede |
| Abreviar la columna BASE | –4 páginas | `BD COMUNA 1` ocupaba tres renglones en cada fila; `BDC 1` ocupa uno |
| Letra 9pt en vez de 10pt | –8 páginas | Sigue siendo perfectamente legible en tabla |
| Encabezado repetido en cada hoja | +6 páginas | Sin esto hay hojas donde no se sabe qué columna es cuál |
| Ninguna fila cortada al medio | +3 páginas | Un servicio partido entre dos hojas se lee mal y se puede ejecutar mal |

Recortar más la descripción **no sirve**: de 400 a 200 caracteres sólo se ganan
3 páginas y se pierde contenido.

**Ojo con `REPETIR_TITULO`:** Word sólo repite filas de encabezado si forman un
bloque contiguo desde la primera fila de la tabla. Si se apaga la banda del
título, dejan de repetirse también los nombres de columna.

### Colores

Hay cuatro paletas en `PALETAS`. Se elige con `PALETA` o `--paleta`:

| Paleta | Cuándo |
|---|---|
| `institucional` | Por defecto. Toma los azules del banner de la Dirección |
| `amarillo` | Mantiene el amarillo de la orden original, pero más calmo |
| `sobrio` | Mínima tinta, para imprimir muchas copias en blanco y negro |
| `original` | Los colores exactos de la orden que venía en Word |

```bash
python generar_orden.py --paleta sobrio
```

---

## Ajustes

Todo se toca en el bloque `CONFIG` arriba de `generar_orden.py`:

| Opción | Para qué |
|---|---|
| `ORIENTACION` | `'apaisado'` o `'vertical'` |
| `PALETA` | Colores de la tabla (ver arriba) |
| `MARGEN_ANEXO` | Aire entre la tabla y el borde de la hoja. `567` = 1 cm |
| `MARGEN_SUP_ANEXO` / `MARGEN_INF_ANEXO` | Aire entre el banner y la tabla |
| `BANNER_ANCHO_COMPLETO` | `True` estira el banner a lo ancho del apaisado. Cuesta 27 páginas |
| `TAM` | Letra en medios puntos. `18` = 9pt, `20` = 10pt |
| `PADDING_CELDA` | Aire arriba y abajo dentro de cada celda |
| `REPETIR_TITULO` | Banda del anexo en cada hoja (ver advertencia arriba) |
| `FILA_ENTERA` | Que ningún servicio se corte entre dos páginas |
| `USAR_ZEBRA` | Sombreado de filas alternas |
| `ORDEN_SERVICIOS` | `'cobertura'` (el que se usa), `'base-turno-dia'`, `'base-dia-hora'`, `'base-hora'`, `'dia-hora'`, `'planilla'` |
| `CAMPOS_SEMANA` / `CAMPOS_FINDE` | **Qué columnas se muestran y su ancho.** Comentá una línea para sacar esa columna: los anchos se reajustan solos |
| `MAX_DESCRIPCION` | `400`. Poné `None` para copiar la función entera |
| `MAX_OBSERVACIONES` | Ídem para observaciones |
| `EXPANDIR_RANGO_DE_DIAS` | `False` deja `MAR A VIE` tal cual |
| `NORMALIZAR_HORA` | `True` reescribe la hora a `08:00 A 19:00` |
| `COLOR_ZEBRA` | Sombreado de filas alternas. `None` lo apaga |
| `NUMERO_ORDEN` / `ANIO_ORDEN` | Nombre del PDF (`ODS_Comunas_032_2026.pdf`) |
| `ZONAS` | Qué comunas integra cada zona |
| `FINDE_SOLO_BASES` | La orden de finde lleva solo bases operativas |
| `OMITIR_ANEXOS_VACIOS` | Saltear los anexos sin servicios validados |
| `ABREVIATURAS_BASE` | Cómo se abrevia cada base en la columna BASE |
| `INCLUIR_ANEXO_IMAGENES` | `False` genera la orden sin el anexo de imágenes |
| `ANEXO_IMAGENES_ORIENTACION` | `'vertical'` o `'apaisado'` |
| `MARCA_INSTRUCCIONES` | El texto que marca los párrafos que no se copian |
| `EXPORTAR_PDF` | Apagá si no querés el PDF |

Para agregar o sacar una columna alcanza con tocar `CAMPOS_SEMANA`. Por ejemplo,
para sumar el ID del servicio:

```python
('id', 'ID', 900, 1000, None),
```

---

## El control de calidad

Cada vez que corre avisa qué servicios validados tienen problemas de carga:
sin base, sin dotación, sin ubicación, sin horario, o con la función tan larga
que ocupa media página. **Eso se arregla en la planilla, no en el Word.**

Un servicio sin base **no entra en ningún anexo** — el aviso es para que no se
pierda.

---

## Llevarlo a otra computadora

Se copia la carpeta entera y listo. Hace falta:

```bash
pip install -r requirements.txt
```

Eso instala `openpyxl` y `python-docx`, que es todo lo que necesita para
generar los Word. Las dependencias opcionales están comentadas dentro del
mismo `requirements.txt`.

**En Windows** el PDF lo genera Word, que ya viene con Office. Si falla con
`ImportError`, falta `pip install pywin32`.

**En Mac** no existe Word por COM, así que el PDF lo genera LibreOffice:

```bash
brew install --cask libreoffice
```

Sin LibreOffice el generador **igual produce el .docx completo** —lo único que
no puede es exportar el PDF, y avisa. En ese caso se abre el Word y se hace
*Guardar como → PDF* a mano.

Lo demás es todo multiplataforma: la lectura del Excel, el armado de las
tablas, el anexo de imágenes y los links del índice funcionan igual.

Dos detalles:

- La planilla se busca en la carpeta **Descargas** del usuario. Si está en otro
  lado, se pasa con `--xlsx`.
- Si la Mac no tiene la tipografía **Roboto**, Word va a sustituirla y las
  tablas se ven distintas. Se instala gratis desde Google Fonts, o se cambia
  `FUENTE` en el config.

---

## Si algo falla

| Síntoma | Causa |
|---|---|
| `No encuentro la planilla` | El xlsx no está en Descargas o tiene otro nombre |
| `el archivo esta abierto` | Tenés el Word generado abierto. Cerralo y volvé a correr |
| `no tiene el marcador {{ANEXOS}}` | Se borró el párrafo marcador de la plantilla |
| Anexos vacíos | No hay servicios con el check `AÑADIDO` tildado |
| `ODS_... está abierto en otro programa` | Tenés el PDF abierto en un visor. Cerralo y volvé a correr |
| `no hay con qué convertir a PDF` | Falta Word (Windows) o LibreOffice (Mac). El .docx se generó igual |
| `no pude exportar el PDF` | Word estaba abierto con un documento trabado |

## Archivos

```
OrdenServicioCAT\
  generar_orden.py           <- el generador
  planilla\                   <- ACA se deja el xlsx bajado de Excel Online
  plantilla_OS_SEMANA.docx   <- texto legal de la semanal, se edita cada semana
  plantilla_OS_FINDE.docx    <- texto legal de la de fin de semana
  anexo_imagenes_SEMANA.docx <- las imágenes de la semanal, se completa a mano
  anexo_imagenes_FINDE.docx  <- las imágenes de la de fin de semana
  ver.py                     <- utilidad: pasa un docx a PDF y a imágenes
  verificar_links.py         <- chequea que los links del índice no estén rotos
  salida\                    <- lo generado
  LEEME.md                   <- esto
```

Para chequear que el índice no tenga links rotos:

```bash
python verificar_links.py "salida\Orden de servicio SEMANA.docx" "salida\ODS_Semanal_032_2026.pdf"
```

# Oficina

Suite ofimática local con procesador de textos, hoja de cálculo y presentaciones. Abre y guarda `.docx`, `.xlsx`, `.xlsm` y `.pptx`, funciona sin conexión y no envía archivos a ningún sitio.

**Para usarla:** abre el acceso directo `Oficina.lnk`, que la abre en Edge en su propia ventana, sin barra de direcciones ni pestañas. También funciona abriendo `index.html` en Edge o Chrome, pero entonces el navegador se queda con Ctrl+N y Ctrl+T.

Es código propio. No contiene código ni binarios de Microsoft; el único recurso suyo son las fuentes Aptos de la carpeta `fonts`, copiadas del Office instalado en este PC. Del análisis de Office solo toma la idea de organización: un núcleo compartido con aplicaciones encima. Los formatos se implementan a partir de su estructura pública (Office Open XML).

## En el móvil

1. En el PC, abre `Oficina para el movil.cmd`. La ventana muestra una dirección como `http://192.168.1.20:8080`. La primera vez Windows pregunta si deja pasar a Node por el cortafuegos: acepta solo para redes privadas.
2. En el móvil, conectado a la misma wifi, abre esa dirección en Chrome o Safari.
3. Para tenerla como una aplicación más: menú del navegador > **Añadir a pantalla de inicio**.

Mientras la usas, el PC tiene que estar encendido con esa ventana abierta. Guardar descarga el archivo en la carpeta de descargas del móvil y Abrir archivo deja elegir uno del teléfono (también de OneDrive o Drive si su aplicación está instalada).

En pantalla estrecha Oficina se adapta: la cinta y la barra de Cálculo son una fila que se desliza con el dedo, Textos empieza en **Vista móvil** (el texto a tamaño de lectura y al ancho de la pantalla; Vista > Diseño de impresión vuelve a las hojas) y en Presentaciones las miniaturas pasan a una tira abajo. Con el dedo: un toque selecciona, doble toque edita una celda o una forma, mantener pulsado abre el menú y las formas y gráficos se arrastran.

**Publicada en https://sergiogomez27-sys.github.io/oficina/** (repositorio público `sergiogomez27-sys/oficina`). Ábrela en el móvil, añádela a la pantalla de inicio y funciona sin el PC y sin conexión. Solo se publica la aplicación: `.gitignore` deja fuera `fonts` (las letras Aptos son de tu licencia de Office; en el móvil se ve con una letra parecida), las pruebas, el servidor y los accesos directos. Para publicar cambios: `git add -A`, `git commit` y `git push`; el móvil coge la versión nueva la siguiente vez que la abre con red. Si cambias la lista de archivos de `sw.js`, sube también el número de `CACHE`.

## Qué hace

| Aplicación | Incluye |
|---|---|
| Textos | Estilos (Normal y Título 1 a 3), fuentes instaladas, tamaños, colores y resaltado, mayúsculas y minúsculas, copiar formato, listas, sangría, interlineado, buscar y reemplazar, tablas, imágenes, vínculos, hojas separadas con saltos de página automáticos, márgenes, orientación y tamaño de papel, zoom e impresión |
| Cálculo | Varias hojas, referencias entre hojas, 90 funciones, tablas, segmentaciones, controles de formulario, asistente de funciones, formato condicional, formatos de número y copiar y pegar fórmulas con ajuste de referencias |
| Presentaciones | Diseños de diapositiva, marcadores de título y texto, fuentes, formas con relleno y contorno, imágenes, copiar y pegar formas, mover y redimensionar, miniaturas, fondos, transiciones (desvanecer, empujar, barrido, dividir y cubrir), zoom y modo de presentación |

**Aspecto de Office 2024.** Textos se maneja como Word y Presentaciones como PowerPoint: barra de título con el color de cada programa, botones de guardar, deshacer y rehacer arriba, y cinta con las pestañas Archivo, Inicio, Insertar, Disposición y Vista (en Presentaciones, Diseño, Transiciones y Presentación con diapositivas en lugar de Disposición y Vista). Inicio agrupa los botones igual que Office: Portapapeles, Fuente, Párrafo, Estilos y Edición en Word, y Portapapeles, Diapositivas, Fuente, Párrafo y Dibujo en PowerPoint. Los colores salen de la misma paleta de Office y la barra de estado muestra la página o diapositiva, las palabras y el zoom. Doble clic en una pestaña oculta la cinta.

Un documento nuevo tiene los estilos de Word 2024 (Aptos 11, 8 pt después de cada párrafo, interlineado 1,08, márgenes de 2,5 y 3 cm, títulos en Aptos Display azul) y una presentación nueva empieza con la diapositiva de título de PowerPoint. Los archivos abiertos se ven con sus propios estilos. Office instala Aptos solo para sus programas, así que el navegador no la ve; la carpeta `fonts` tiene una copia de las de este PC (Aptos, Aptos Display y Aptos Narrow) para que se vea igual que en Office. Son de Microsoft y vienen con la licencia de Office: no compartas esa carpeta con la aplicación.

Atajos de Word en español: Ctrl+G guarda, Ctrl+K cursiva, Ctrl+S subrayado, Ctrl+Q, Ctrl+D y Ctrl+J alinean, Ctrl+E selecciona todo, Ctrl+B busca, Ctrl+L reemplaza, Ctrl+Intro inserta un salto de página y Ctrl+> y Ctrl+< cambian el tamaño. Ctrl+N (negrita) y Ctrl+T (centrar) solo funcionan abriendo Oficina con `Oficina.lnk`: en una pestaña normal el navegador los usa para abrir ventanas y pestañas. En Presentaciones, Ctrl+M crea una diapositiva y F5 empieza la presentación.

El clic derecho abre un menú en Cálculo (celdas, encabezados de fila y columna, pestañas de hoja) y en Presentaciones (formas, lienzo y miniaturas). En Textos se usa el menú del navegador, que incluye el corrector.

Las fórmulas se escriben como en un Excel en español: `=SUMA(A1:A5)`, `=SI(A1>5;"alto";"bajo")`, `=BUSCARV(A1;B:C;2;FALSO)`, con `;` entre argumentos y coma decimal. También se aceptan los nombres en inglés con `,` (`=SUM(A1,A2)`). En el archivo se guardan en inglés, como hace Excel, así que se abren bien en cualquier idioma.

Para crear fórmulas hay tres ayudas:
- **fx Función** (o Mayús+F3): elige la función de una lista con su descripción y rellena sus argumentos.
- **Autocompletado:** al escribir `=PRO` aparecen las funciones que empiezan así; Tab completa la elegida y, dentro de la función, se muestra su sintaxis.
- **Σ Suma:** suma lo que hay encima o a la izquierda.

**Formato condicional** (botón de la barra): resaltar celdas mayores, menores, entre, iguales, con un texto, con una fecha o duplicadas; superiores e inferiores; por encima o por debajo del promedio; barras de datos; escalas de color; conjuntos de iconos; reglas con fórmula; borrar y administrar reglas. Se leen y guardan las reglas de Excel.

Como en Excel:
- **Combinar** (⊞ y su menú ▾): combinar y centrar, combinar horizontalmente, combinar celdas y separar.
- **Validación de datos:** las listas muestran el botón ▾ junto a la celda (o Alt+↓). Si la regla lo pide, un valor no permitido se rechaza con su mensaje.
- **Errores en español** (`#¡DIV/0!`, `#¿NOMBRE?`, `#N/D`…). Una fórmula mal escrita no se acepta y, si solo falta cerrar un paréntesis, se cierra sola. Para escribir un texto que empieza por `=`, ponle delante un apóstrofo (`'`).
- El texto largo sigue por encima de las celdas vacías de la derecha. Lo que no cabe se corta sin «…». Un número que no cabe se muestra como `###`. También se muestran el texto girado, los bordes de cualquier estilo (trazos, puntos, raya y punto, dobles) y la fuente de cada celda. Las columnas miden 64 píxeles y las filas 20, como en un libro nuevo de Excel.
- **Gráficos** (botón Gráfico ▾): columnas, barras, líneas, circular, anillo, áreas y dispersión a partir de los datos seleccionados. Los gráficos, imágenes y cuadros de texto de los archivos se ven en su sitio, se pueden mover, cambiar de tamaño y borrar, y los gráficos se actualizan al cambiar los datos.
- **Filtro** (botón ⏷ Filtro o Ctrl+Mayús+L): flechas en los encabezados con ordenar, buscar, casillas de valores, (Vacías) y filtros de condición. Se leen y guardan los filtros de Excel, también los de color y los de los diez mejores.
- **Proteger hoja y Proteger libro** (clic derecho en la pestaña): con contraseña opcional y los mismos permisos que Excel. Las celdas desbloqueadas siguen editables y las fórmulas ocultas no se ven. La contraseña se guarda cifrada igual que en Excel, así que la de un archivo de Excel sigue sirviendo.
- **Tablas** (las de Insertar > Tabla o Dar formato como tabla): se ven con su estilo (los 60 integrados de Excel y los personalizados del libro, con encabezado, bandas, primera y última columna y fila de totales) y cada una tiene su filtro. Las fórmulas admiten referencias estructuradas en español (`=[@Ventas]-[@Gastos]`, `=SUMA(Tabla1[Ventas])`, `Tabla1[#Encabezados]`) y `SUBTOTALES`. Como en Excel, escribir debajo o a la derecha amplía la tabla, una fórmula en una columna vacía la rellena entera (columna calculada) y cambiar un encabezado renombra la columna en todas las fórmulas.
- **Segmentaciones:** filtran al pulsar un valor (Ctrl+clic para elegir varios, el embudo para borrar el filtro). Las de una tabla ordenan los meses y días como Excel y dejan al final, más claros, los valores ocultos por otro filtro. Las de una tabla dinámica vuelven a calcular todas las tablas dinámicas que conectan.
- **Tablas dinámicas:** se ven con los valores que guardó Excel y Oficina las recalcula desde su caché al filtrarlas con una segmentación, igual que Excel. Lo hace con el diseño compacto (el que Excel usa por omisión): campos de fila anidados con subtotales, un campo de columna o varios valores, filtros de informe y totales generales, con suma, cuenta, promedio, máximo, mínimo, producto, desviación y varianza.
- **Controles de formulario:** casillas, botones de opción, listas desplegables, cuadros de lista, controles de número y barras de desplazamiento cambian su celda vinculada y las fórmulas que dependen de ella.
- **Controles ActiveX:** casillas, botones de opción y de alternar, cuadros combinados, de lista y de texto, controles de número y barras de desplazamiento se ven y cambian su celda vinculada como en Excel.
- **Archivos con macros (`.xlsm`):** se abren y se guardan conservando las macros, pero Oficina no las ejecuta.

## Qué no conserva

Textos reparte el texto en hojas como Word: el párrafo que no cabe pasa entero a la hoja siguiente (Word puede partirlo entre las dos) y un salto de página empieza hoja nueva. Por eso el número de páginas puede variar en uno respecto a Word en documentos largos.

Al abrir un archivo, Oficina avisa de buena parte de lo que no va a conservar, aunque no de todo lo que aparece aquí. Si después guardas, eso se pierde en el archivo nuevo, así que conviene conservar el original.

- **Textos:** encabezados y pies, comentarios, notas al pie, control de cambios, formas dibujadas e índices automáticos.
- **Cálculo:** las tablas dinámicas con diseño tabular o de esquema, campos calculados, agrupaciones o «Mostrar valores como» se conservan y se filtran al guardar, pero Oficina no las recalcula: lo hace Excel al abrir el archivo. Oficina tampoco las actualiza cuando cambian los datos de origen; como en Excel, eso se hace al actualizar, que Excel hace al abrir. Los controles ActiveX que guardan su texto en binario (lo habitual cuando los crea Excel) se ven con su nombre en lugar de su texto, y su valor se guarda en la celda vinculada pero no dentro del control. Esto se ha probado con un libro preparado según el formato, porque el Excel de las pruebas tiene los ActiveX desactivados. Los botones de formulario y ActiveX conservan su macro, que no se ejecuta. Las macros se conservan sin ejecutarse; esto solo se ha probado con archivos generados en las pruebas, no con un libro real con macros.
- **Presentaciones:** tablas, gráficos, animaciones, transiciones distintas de las cinco que reproduce (se ven como Desvanecer), conectores, notas del orador y fondos o imágenes del patrón.

Tampoco tiene corrector propio (usa el del navegador), colaboración en línea ni ejecución de macros.

## Borradores

Mientras trabajas se guarda un borrador en el navegador, y aparece en la página de inicio si cierras sin guardar. No sustituye al archivo: si borras los datos del navegador, se pierde. **Guardar** descarga el archivo en la carpeta de descargas.

## Estructura

```
index.html, styles.css   Entrada y estilos
core/                    Núcleo compartido: XML y ZIP, fórmulas, ejemplos y lanzador
formats/                 Lectura y escritura de docx, xlsx y pptx
apps/                    Interfaz de cada aplicación
vendor/jszip.min.js      JSZip 3.10.1 (licencia MIT o GPLv3), para leer y crear ZIP
tests/                   Pruebas
```

## Pruebas

Requieren Node.js, y la segunda también Office instalado.

```bash
cd tests && npm install && node run.js
```

```bash
powershell -NoProfile -ExecutionPolicy Bypass -File tests/office_check.ps1
```

`run.js` comprueba las fórmulas y la ida y vuelta de cada formato. `office_check.ps1` abre en Office lo generado por Oficina y crea en `tests/desde_office` archivos que `run.js` lee en su siguiente ejecución.

Pruebas completas, en este orden, con Edge y Office instalados:

| Script | Qué hace |
|---|---|
| `complejo_office.ps1` | Crea con Office documentos complejos y los resultados de Excel para 123 fórmulas |
| `formulas_vs_excel.js` | Compara el motor de Oficina con esos resultados. Tarda unos 5 minutos por la lentitud del DOM simulado |
| `e2e.js` | Usa la aplicación en Edge desde `file://`, con descargas y PDF reales |
| `menu_excel.js`, `menu_slides.js` | Prueban en Edge los menús de clic derecho |
| `menu_cf.js` | Prueba en Edge el formato condicional, las fórmulas en español, el asistente y el autocompletado |
| `movil_ui.js` | Prueba Oficina como en un móvil táctil (diseño estrecho, dedo, doble toque y sin conexión); antes arranca `node servidor.js 8099` |
| `paginas_ui.js` | Prueba en Edge las hojas de Textos, los saltos y la pestaña Disposición, y guarda `e2e/paginas.docx` |
| `transiciones_ui.js` | Prueba en Edge la pestaña Transiciones y la presentación, y guarda `e2e/transiciones.pptx` |
| `cinta_ui.js` | Prueba en Edge la cinta de Textos y Presentaciones y guarda `e2e/cinta.docx` y `e2e/cinta.pptx` para abrirlos en Office |
| `tablas_office.ps1` | Crea con Excel `desde_office/tablas.xlsx`: tablas con totales, filtro y estilo propio, segmentaciones, dos tablas dinámicas y controles |
| `activex_office.ps1` | Crea `desde_office/activex.xlsx`: un libro de Excel al que `activex_make.js` añade controles ActiveX |
| `tablas_ui.js` | Usa en Edge las segmentaciones y los controles de esos libros y los guarda en `salida/` |
| `validar.ps1 -Archivos "a.docx,b.xlsx"` | Valida archivos contra el esquema OOXML con el SDK instalado con Office |
| `abrir_e2e.ps1` | Abre en Office lo que descargó `e2e.js` |
| `vista_impresion.js` | Captura la vista de impresión para revisarla a ojo |

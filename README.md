# SSBB: consumos de agua faltantes

Herramienta para convertir consumos mensuales en filas y generar solamente los períodos que faltan en la base de agua, sin modificar el archivo original.

## Uso

1. Descarga este repositorio (Code → Download ZIP) y descomprímelo, o clónalo.
2. Instala **Python 3.10 o superior**. En Windows, activa la opción de agregar Python a PATH durante la instalación.
3. Coloca tu Excel `.xlsx` en `input/`. Debe contener las tres hojas descritas abajo.
4. **Windows:** haz doble clic en `ejecutar.bat`. **Linux/macOS:** abre una terminal en esta carpeta y ejecuta `bash ejecutar.sh`.
5. Encontrarás `<nombre>_faltantes_2026.xlsx` en `output/`.

El lanzador prepara un entorno Python local en `.venv/` e instala las dependencias. La primera ejecución necesita Internet; también se consulta el gestor de paquetes en ejecuciones posteriores. No es un `.exe` autónomo: requiere Python instalado. El lanzador de Windows se proporciona para ese sistema; la validación realizada en desarrollo es en Linux.

```text
SSBB/
├── input/              ← coloca aquí uno o varios Excel
├── output/             ← resultados, uno por Excel
├── ejecutar.bat        ← Windows: doble clic
├── ejecutar.sh         ← Linux/macOS
├── preparar_carga.py
├── requirements.txt
└── tests/
```

El año por defecto es **2026**. Para cambiarlo:

```bash
bash ejecutar.sh --year 2027
# Windows, desde una terminal:
ejecutar.bat --year 2027
```

Cada Excel se procesa de forma independiente contra su propia base. No se cruzan ni deduplican archivos distintos. Se ignoran los archivos temporales `~$`. Si ya existe el resultado, se informa un error y no se sobrescribe: muévelo o elimínalo antes de repetir. Actualiza la base de datos del Excel de entrada después de cada carga; repetir con la misma base producirá los mismos faltantes.

## Formato de entrada

Se comprueban los encabezados principales antes de procesar. Se espera el formato del archivo original:

- **Hoja1:** encabezados en fila 2; sucursal en C, código en D, cliente en E, consumos de enero a octubre en L:U. Datos desde fila 3. Noviembre y diciembre no se procesan en este formato.
- **BBDD N° Clientes:** encabezados en fila 1; cliente en A, código de sucursal en C, fuente de obtención de datos en G y nombre de sucursal en J.
- **Base de datos Agua Reprocesada:** encabezados en fila 1; cliente en J, fecha en K, consumo en M. Fuente de información en U. Mes y año en S:T se usan cuando no hay una fecha válida.

Se leen valores guardados de Excel, no se calculan fórmulas. Si el archivo contiene fórmulas de entrada, abre y guarda primero el libro en Excel para actualizar sus resultados.

## Reglas

- Se conservan ceros, números y textos; se omiten únicamente vacíos.
- Fecha convencional: **día 28 de cada mes** del año seleccionado. No representa la fecha real de emisión.
- Cliente + año + mes identifica un período. Si ya existe, no se vuelve a cargar, aunque cambie el consumo; no se actualizan registros existentes.
- Se normalizan espacios, tildes y signos para comparar identificadores. En la salida se conserva el identificador del catálogo.
- Si el cliente contiene `/`, se busca uno de sus componentes y se exige el código de sucursal.
- Un número sin dígito verificador puede coincidir con el número completo del catálogo si también coincide el código de sucursal.
- Para gestión interna, infraestructura, gasto común y otros identificadores especiales, se usa el código de sucursal y la familia del identificador cuando corresponde. Las coincidencias sin un cliente único se dejan pendientes, sin inventar números.
- Dos valores distintos de origen para el mismo cliente y período se dejan pendientes. Repeticiones idénticas se consolidan.
- Número de documento, límite, sobreconsumo e importes quedan vacíos: Hoja1 no aporta esos datos.

## Resultado y pegado

**Identificador para cargar:** A de `Carga J-R` siempre contiene el número de cliente exacto del catálogo `BBDD N° Clientes`, sin reemplazarlo por el valor original de Hoja1 ni por su versión normalizada. El programa verifica que cada identificador de carga exista en ese catálogo, distinguiendo texto de número. El identificador original se muestra aparte en M. Los clientes sin correspondencia segura aparecen en Pendientes y nunca en Carga J-R. En Todos J-R esos casos conservan el original en A y M, resaltados en amarillo; esa hoja es solo para comparación.

- **Carga J-R:** las primeras nueve columnas A:I equivalentes a las columnas J:R de la base. Copia desde A2 hasta I de la última fila de datos y pega en J:R del destino. Si no hay datos, no hay filas que copiar.
- **Todos J-R:** todos los datos no vacíos de Hoja1 con las mismas columnas de carga y de fuente, ordenados por cliente y fecha. Incluye períodos existentes, faltantes, repeticiones y valores conflictivos de origen, conservando ceros y textos. Es una hoja de comparación; **no debe pegarse como carga**. Los clientes sin correspondencia segura conservan el identificador original, resaltado en amarillo y con un comentario; también aparecen en Pendientes. No es una copia de todo el historial de la base: representa los registros de Hoja1.
- **Pendientes:** registros que requieren resolver una correspondencia o conflicto; no se incluyen en la carga.
- **Ya existentes:** registros excluidos y referencias a filas del historial.
- **Trazabilidad** y **Correspondencias:** origen y método de asignación de los registros.
- **Instrucciones:** criterios y recuentos.

En **Carga J-R** y **Todos J-R** se agregan cuatro columnas filtrables después de las nueve originales:

| Columna de salida | Contenido | Uso |
| --- | --- | --- |
| J | Fuente de información | Clasificación del cliente, por ejemplo `Casa Funcionarios no se considera` |
| K | Origen de la fuente | Explica qué base se usó y advierte diferencias o ambigüedad |
| L | Fuentes en historial Agua | Todas las clasificaciones distintas encontradas para ese cliente en el historial |
| M | Número de cliente original Hoja1 | Identificador recibido, solo para comparación; no se carga |

La fuente se obtiene primero de G de **BBDD N° Clientes**. Si no existe, se utiliza U de **Base de datos Agua Reprocesada** solo cuando hay una única fuente histórica. Si hay varias fuentes en el catálogo o, a falta de catálogo, varias fuentes históricas, J queda vacía y K indica que debe revisarse. No se elige una fuente arbitrariamente ni se excluyen automáticamente clientes por su clasificación. `#N/A` y otros valores de error no se tratan como fuentes válidas. Las fuentes históricas se consultan en todos los años disponibles, no solo en el año de la carga.

**Pegado:** copia A:I de la salida en J:R del destino, como antes. Si también quieres cargar la fuente, copia **J de la salida a U del destino** por separado, usando exactamente las mismas filas y filtros. K:M son auxiliares. No pegues A:M como un bloque: S:T del destino corresponden a mes y año, no a fuente.

La columna F del resultado corresponde a **O (Sobre consumo)** del destino y queda vacía. Al pegar, utiliza las opciones apropiadas para conservar las fórmulas del destino (por ejemplo, omitir blancos). No se añaden automáticamente fórmulas ni filas a la base original.

Los Excel de entrada y salida se ignoran en Git: no se suben a GitHub al guardar cambios del programa.

## Desarrollo y comprobación

Cada tarea en la nube ya está aislada: usa el checkout existente; no necesitas crear un worktree.

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python -m unittest discover -s tests -v
.venv/bin/python preparar_carga.py --year 2026
```

También se pueden seleccionar carpetas con `--input ruta` y `--output ruta`. No pueden ser la misma carpeta.

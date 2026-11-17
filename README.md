Juego de la Vida de Conway en Excel
Simulación del autómata celular de Conway desarrollada exclusivamente con fórmulas de Microsoft Excel, sin VBA ni macros.

Permite avanzar por generaciones con F9, configurar el tamaño del tablero e insertar patrones predefinidos durante la simulación.
Descripción
El proyecto implementa una cuadrícula bidimensional de células que evolucionan siguiendo las reglas del Juego de la Vida de Conway (B3/S23).

Utiliza matrices dinámicas, funciones LAMBDA y referencias circulares con cálculo iterativo para conservar y actualizar el estado entre generaciones.

Las fronteras son periódicas: los extremos opuestos del tablero están conectados, formando una topología toroidal.
Versiones
Archivo	Descripción
Ejercicio 01 Conway.xlsx	Primera versión: población inicial aleatoria, reinicio y avance por generaciones.
Ejercicio 01 Conway V2.0.xlsx	Añade inicio en blanco, catálogo de patrones, inserción por coordenadas y vistas mediante imágenes vinculadas.


Características
- Dimensiones configurables del tablero.
- Población inicial aleatoria o tablero en blanco.
- Reinicio mediante casilla de verificación.
- Avance mediante la tecla F9.
- Contador de pasos.
- Catálogo de 15 patrones predefinidos.
- Previsualización del patrón seleccionado.
- Inserción mediante coordenadas de su esquina superior izquierda.
- Una inserción por activación de la casilla, sin repetición continua al dejarla marcada.
- Fronteras periódicas.
- Representación visual mediante formato condicional.
- Imágenes vinculadas para mostrar el tablero y el patrón a una escala independiente de las celdas de configuración.
- Implementación sin macros.
Requisitos
Excel de escritorio con soporte para las funciones utilizadas:
- Matrices dinámicas y referencias a rangos desbordados (#).
- LET, LAMBDA y REDUCE.
- MAKEARRAY —puede aparecer como ARCHIVOMAKEARRAY—.
- MATRIZALEAT y SECUENCIA.
- DIVIDIRTEXTO y COINCIDIRX, utilizadas en el catálogo de la versión 2.0.
Los nombres de las funciones pueden variar según el idioma de Excel.
Configuración inicial
1. Descargar y abrir el archivo de la versión deseada.
2. Ir a Archivo → Opciones → Fórmulas.
3. Activar Habilitar cálculo iterativo.
4. Establecer el número máximo de iteraciones en 1.
5. Seleccionar el modo de cálculo Manual.
La configuración de cálculo es parte del funcionamiento del simulador. Comprueba que cada pulsación de F9 produce el avance esperado.

Uso de la versión 2.0
Iniciar o reiniciar
1. Introducir el número de filas y columnas.
2. Seleccionar el tipo de inicio: Aleatorio o En blanco.
3. Activar Reset.
4. Pulsar F9 si el tablero no se ha actualizado.
5. Desactivar Reset.
6. Pulsar F9 para avanzar.
Mientras Reset permanezca activado, se seguirá aplicando la inicialización. En modo aleatorio, un nuevo cálculo puede generar otra población.

Para cambiar las dimensiones del tablero, realizar un reinicio.
Insertar un patrón
1. Desmarcar Insertar patrón, si estaba marcada.
2. Seleccionar una figura del catálogo.
3. Consultar su previsualización, anchura y altura.
4. Introducir las coordenadas de inserción:
   - X: columna, contando desde la izquierda.
   - Y: fila, contando desde arriba.
   - Ambas comienzan en 1.
5. Marcar Insertar patrón.
6. Pulsar F9 si la inserción todavía no se ha ejecutado.
La inserción se realiza cuando coinciden el paso actual y el paso de inserción preparado por el libro.

Dejar la casilla marcada no inserta continuamente el patrón. Para preparar una nueva inserción, desmarcarla y volver a marcarla.

Desmarcar la casilla no elimina las células que ya se han insertado.
Comportamiento de la inserción
En la versión 2.0, se calcula la siguiente generación y después se coloca el patrón sobre el resultado.

La inserción reemplaza el rectángulo completo del patrón, incluidos sus ceros. Por tanto, puede borrar células existentes dentro de esa región.

Las coordenadas corresponden a la esquina superior izquierda de la matriz del patrón, incluido su margen de células muertas.
Patrones incluidos
Patrón	Tipo	Período
Bloque	Estable	1
Colmena de abejas	Estable	1
Pan	Estable	1
Bote	Estable	1
Bañera	Estable	1
Intermitente	Oscilador	2
Sapo	Oscilador	2
Faro	Oscilador	2
Púlsar	Oscilador	3
Pentadecátlon	Oscilador	15
Planeador	Nave espacial	4
Nave espacial ligera	Nave espacial	4
Nave espacial de peso medio	Nave espacial	4
Nave espacial pesada	Nave espacial	4
Cañón de planeadores de Gosper	Generador de planeadores	30


En las naves espaciales, el período indica cuándo se repite la forma en una posición desplazada.

El cañón de Gosper emite un planeador cada 30 generaciones cuando no recibe perturbaciones. El conjunto completo no vuelve al estado inicial, porque incluye los planeadores emitidos.
Organización del libro
La versión 2.0 contiene tres hojas:
Hoja	Función
Implementación	Controles, contador, selección de patrones y visualización.
Tablero	Matriz del estado de la simulación.
Atlas	Catálogo y conversión del patrón seleccionado en una matriz numérica.


Las imágenes vinculadas muestran rangos del libro y permiten redimensionar su presentación sin modificar el tamaño de las celdas de los controles.
Funcionamiento
Reglas B3/S23
Cada célula considera sus ocho vecinas:
- Una célula muerta nace si tiene exactamente 3 vecinos vivos.
- Una célula viva sobrevive si tiene 2 o 3 vecinos vivos.
- En cualquier otro caso, la célula queda muerta.
Todas las células de una generación se calculan a partir del estado anterior.
Cálculo de la siguiente generación
La función definida siguientePaso(grilla):
1. Obtiene las dimensiones de la matriz.
2. Recorre las posiciones mediante MAKEARRAY.
3. Cuenta los ocho vecinos mediante REDUCE.
4. Conecta los bordes mediante RESIDUO.
5. Aplica las reglas B3/S23.
Conservación del estado
La fórmula del tablero hace referencia a su propio rango desbordado. El cálculo iterativo permite utilizar el estado anterior para obtener el siguiente.

El contador también utiliza una referencia circular para registrar el avance.
Control de inserción
Mientras la casilla está desmarcada y la validación de posición se cumple, el libro prepara un paso de inserción.

Al marcarla, ese paso queda almacenado. El patrón se inserta cuando coincide con el paso actual. Después, el contador continúa avanzando y la coincidencia deja de cumplirse.
Formato de los patrones
Los patrones se almacenan como texto:
- 1: célula viva.
- 0: célula muerta.
- ;: separador de columnas.
- \: separador de filas.
Ejemplo de un bloque con un margen de células muertas:
0;0;0;0\0;1;1;0\0;1;1;0\0;0;0;0

La hoja Atlas convierte este texto en una matriz numérica.
Comprobaciones básicas
Para verificar el comportamiento del simulador, utilizar patrones aislados:
Prueba	Resultado esperado
Bloque	Permanece sin cambios.
Intermitente	Recupera su estado inicial tras 2 generaciones.
Planeador	Tras 4 generaciones recupera su forma, desplazado una fila y una columna.
Casilla de inserción marcada durante varios pasos	No repone continuamente el patrón.


Limitaciones y consideraciones
- El funcionamiento depende de la configuración de cálculo de Excel.
- Editar controles o realizar otras operaciones que provoquen recálculo puede afectar al avance; F9 no es necesariamente la única acción que lo desencadena.
- Los tableros grandes requieren más tiempo de cálculo y actualización visual.
- Las imágenes vinculadas son vistas del tablero; no permiten editar directamente las células.
- Los patrones pueden interactuar con células existentes y perder su comportamiento habitual.
- Al conectar los bordes, una nave puede reaparecer por el extremo opuesto y colisionar con otros patrones.
- El catálogo utiliza rangos de búsqueda fijos en esta versión. Para ampliarlo, también deben actualizarse los rangos utilizados por las fórmulas y el selector.
Referencias
- LifeWiki — Conway’s Game of Life
- LifeWiki — Glider
- LifeWiki — Gosper glider gun
- Microsoft — Función MAKEARRAY
- Microsoft — Función LAMBDA
- Microsoft — Mostrar dinámicamente un rango en un objeto gráfico
Licencia
Consultar el archivo LICENSE del repositorio.

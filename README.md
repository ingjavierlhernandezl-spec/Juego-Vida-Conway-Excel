Juego de la Vida de Conway en Excel

Simulación del autómata celular de Conway desarrollada exclusivamente con fórmulas de Microsoft Excel, sin utilizar VBA ni macros.

Descripción

El proyecto implementa una cuadrícula bidimensional de células que evolucionan siguiendo las reglas del Juego de la Vida de Conway (B3/S23).

La simulación utiliza matrices dinámicas, funciones LAMBDA y referencias circulares controladas para mantener el estado entre generaciones.

Características

Dimensiones configurables de la cuadrícula.

Generación aleatoria de la población inicial.

Reinicio mediante casilla de verificación.

Evolución mediante la tecla F9.

Fronteras periódicas (topología toroidal).

Representación visual mediante formato condicional.

Implementación sin macros.

Requisitos

Microsoft Excel con soporte para matrices dinámicas y funciones LET, LAMBDA, REDUCE y MAKEARRAY.

Instrucciones

Descargar y abrir el archivo Ejercicio 01 Conway.xlsx.

Habilitar el cálculo iterativo en Archivo → Opciones → Fórmulas.

Configurar el número máximo de iteraciones en 1.

Seleccionar el modo de cálculo Manual.

Introducir el número de filas y columnas.

Activar Reset para generar una población aleatoria.

Desactivar Reset.

Presionar F9 para avanzar la simulación.

Funcionamiento

La función SiguientePaso(grilla) calcula el estado de la siguiente generación considerando las ocho células vecinas.

Se aplican las reglas B3/S23:

Una célula muerta nace cuando tiene exactamente 3 vecinos vivos.

Una célula viva sobrevive cuando tiene 2 o 3 vecinos vivos.

En los demás casos, la célula permanece muerta o muere.

La matriz se actualiza mediante una fórmula autorreferencial con cálculo iterativo.

Limitaciones

La simulación depende de la configuración de cálculo de Excel. Se recomienda comprobar que cada pulsación de F9 produce exactamente una generación.

Referencias

Conway's Game of Life — https://conwaylife.com/wiki/Conway%27s_Game_of_Life

Microsoft Excel — https://support.microsoft.com/excel

Licencia

Consultar el archivo LICENSE del repositorio.

COMO USAR SCRIPT----

El script usa bucles para leer,escribir y operar con archivos .rwl

Para usarlo necesitamos tener 1 CARPETA DONDE SOLO TENGAMOS LOS ARCHIVOS RWL PARA CONVERTIR

El script devolverá como output:
* 1 carpeta nueva que contendrá todos los archivos rwl convertidos a excel con valores brutos con el nombre de carpeta "ArchivosExcel"
* 1 carpeta dentro de la anteriormente mencionada con el nombre "Calculado_Incremento_BAI" con todos los excel donde se ha calculado "celda_actual+celda_previa"
* 1 carpeta dentro de la anteriormente mencionada con el nombre "Incremento_Superficie_Anual" con todos los excel donde se ha calculado el incremento a partir del BAI con la fórmula (((pi*(n[[i,j]]^2))-(pi*(n[[i-1,j]])^2))/100)

Para el correcto funcionamiento tenemos que cambiar las direcciones de lecto-escritura de las matrices para que concuerden con nuestro ordenador y tenemos que asegurarnos que las carpetas se llamen exactamente como se indican aquí INCLUYENDO TILDES
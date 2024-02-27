# Licencias estados
Script para revisar estado de licencias en página milicenciamedica.cl por listados de RUT.

# Coding in times of COVID-19.
Realizado con la intención de optimizar procesos de revisión estado de licencias médicas para personal del HGGB Concepción.
Hecho por Jonathan Friz B. jfriz[@]protonmail.com

# Instrucción
Debes Reemplazar Licencias_beta.xlsx por nombre de tu archiivo .xlsx)

# Cambios y mejoras.
Mejoras desde versión 1. Se maneja errores de Key Error con excepción para que no se detenga script, 
Al igual con TypeError.
Se agregan colores y formatos para visualización en consola
Se agrega tiempos de esperas ya no fijos, sino aleatorios dentro de un rango para reposo del script 
y evitar bloqueos y sobrecargas del sitio web. 
Se agrega hora de inicio y final con marca de tiempo y se hace diferencia de la duración del proceso completo.
Se acorta a escribar en planilla de salida ahora se registrara sin el primer digito ni el guion para mejor compración, 
el sitio web lo entrega con este digito.
----------------------------------------------------------
Version 3.0 Se generan los links en python directo de lista de RUT y Folios.
Se genera validación de ruts mayores y menores a 10 millones.
----------------------------------------------------------
3.1 Se agregan índices o headers a la planilla de salida

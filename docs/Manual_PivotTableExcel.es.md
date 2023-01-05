# Tablas dinámicas
  
Módulo para trabajar e interactuar con tablas dinámicas  

*Read this in other languages: [English](Manual_PivotTableExcel.md), [Español](Mannual_PivotTableExcel.es.md).*
  
![banner](imgs/Banner_PivotTableExcel.png)
## Como instalar este módulo
  
__Descarga__ e __instala__ el contenido en la carpeta 'modules' en la ruta de Rocketbot.  



## Descripción de los comandos

### Crear
  
Crea una nueva tabla dinámica
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ingrese Rango de datos |Ingrese el rango de datos que desea utilizar para crear la tabla dinámica|Hoja1!B2:C4|
|Celda de destino |Ingrese la celda donde desea que se cree la tabla dinámica|Hoja2!C4|
|Nombre de la tabla dinámica |Ingrese el nombre de la tabla dinámica|Name: |

### Actualizar
  
Actualiza una tabla dinámica
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentra la tabla dinámica|Hoja 1|
|Nombre de la tabla dinámica |Nombre de la tabla dinámica a actualizar|Name: |

### Agregar campo
  
Agrega un campo a una tabla dinámica
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentra la tabla dinámica|Hoja 1|
|Nombre de la tabla dinámica |Nombre de la tabla dinámica|Name: |
|Selecciona una opción|Selecciona una opción para agregar un campo a la tabla dinámica|Add Data|
|Campo a agregar|Nombre del campo a agregar|Field: |

### Filtrar
  
Filtra una tabla dinámica
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentra la tabla dinámica|Hoja 1|
|Nombre de la tabla dinámica |Nombre de la tabla dinámica|Name: |
|Nombre del filtro |Nombre del campo que se quiere filtrar|Campo |
|Limpiar todos los filtros|Si se selecciona, se limpiarán todos los filtros de la tabla dinámica|True|
|Valor(es) del filtro a marcar |Nombre del valor del filtro que se quiere marcar|Name: |
|Valor(es) del filtro a desmarcar |Nombre del valor del filtro que se quiere desmarcar|Name: |

### Listar Campos
  
Lista todos los campos disponibles
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentra la tabla dinámica|Hoja 1|
|Nombre de la tabla dinámica |Nombre de la tabla dinámica|Name: |
|Asignar resultado a variable|Nombre de la variable donde se almacenará el resultado|Variable|

### Cambiar origen de datos
  
Cambia el rango de origen de los datos de una tabla dinámica
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentra la tabla dinámica|Hoja 1|
|Nombre de la tabla dinámica |Nombre de la tabla dinámica|Name: |
|Nuevo Rango|Rango de datos de la tabla dinámica|A1:R200|

### Lista items de filtro
  
Devuelve una lista con los items de un filtro
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja|Hoja 1|
|Nombre de la tabla dinámica |Nombre de la tabla dinámica|Name: |
|Nombre del filtro |Nombre del filtro|Campo |
|Asignar resultado a variable|Nombre de la variable donde se almacenará el resultado|Variable|

### Insertar Línea de tiempo
  
Crea una nueva línea de tiempo
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se insertará la línea de tiempo|Hoja 1|
|Nombre de la tabla dinámica |Nombre de la tabla dinámica que se usará para crear la línea de tiempo|TablaDinámica1: |
|Campo de la tabla dinámica |Nombre del campo de la tabla dinámica que se usará para crear la línea de tiempo|Campo |
|Rango donde posicionar|Rango donde se insertará la línea de tiempo|A1:D20|

### Filtrar Slider
  
Modifica el filtro de un slider
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentra el slider|Hoja 1|
|Nombre del slider |Nombre del slider|Name: |
|Fecha de incio|Fecha de inicio del filtro|13/12/1999: |
|Fecha final|Fecha final del filtro|13/12/2000: |

### Estado de filtro
  
Retorna True si el filtro está marcado
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentra el filtro|Hoja 1|
|Nombre de la tabla dinámica |Nombre de la tabla dinámica donde se encuentra el filtro|Name: |
|Nombre del filtro |Nombre del filtro que se desea consultar|Campo |
|Valor del filtro|Valor del filtro que se desea consultar|Name: |
|Asignar resultado a variable||Variable|

### Cambiar tabla a formato tabular
  
Altera los campos de la tabla dinamica a formato tabular.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentra la tabla dinamica.|Hoja 1|
|Nombre de la tabla |Nombre de la tabla dinamica.|Name: |
|Campos de la tabla|Campos de la tabla dinamica.|['Numero', 'Fecha', 'Hora']|

### Borrar subtotales
  
Elimina los subtotales de la tabla dinamica.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentra la tabla dinamica.|Hoja 1|
|Nombre de la tabla |Nombre de la tabla dinamica.|Name: |
|Campos de la tabla|Campos de la tabla dinamica.|['Numero', 'Fecha', 'Hora']|

### Repetir etiquetas de elementos
  
Permite a la tabla dinamica repetir etiquetas.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentra la tabla dinamica.|Hoja 1|
|Nombre de la tabla |Nombre de la tabla dinamica.|Name: |
|Campos de la tabla|Campos de la tabla dinamica que se repetiran.|['Numero', 'Fecha', 'Hora']|

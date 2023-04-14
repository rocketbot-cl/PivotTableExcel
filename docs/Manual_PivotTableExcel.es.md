



# Tablas dinámicas
  
Módulo para trabajar e interactuar con Tablas Dinámicas de Microsoft Excel.  
  
![banner](imgs/Banner_PivotTableExcel.png)
## Como instalar este módulo
  
Para instalar el módulo en Rocketbot Studio, se puede hacer de dos formas:
1. Manual: __Descargar__ el archivo .zip y descomprimirlo en la carpeta modules. El nombre de la carpeta debe ser el mismo al del módulo y dentro debe tener los siguientes archivos y carpetas: \__init__.py, package.json, docs, example y libs. Si tiene abierta la aplicación, refresca el navegador para poder utilizar el nuevo modulo.
2. Automática: Al ingresar a Rocketbot Studio sobre el margen derecho encontrara la sección de **Addons**, seleccionar **Install Mods**, buscar el modulo deseado y presionar install.



## Descripción de los comandos

### Crear
  
Crea una nueva tabla dinámica
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Ingrese Rango de datos ||Hoja1!B2:C4|
|Celda de destino ||Hoja2!C4|
|Nombre de la tabla dinámica ||Name: |

### Actualizar
  
Actualiza una tabla dinámica
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja ||Hoja 1|
|Refrescar todo ||False|
|Nombre de la tabla dinámica ||Nombre: |

### Agregar campo
  
Agrega un campo a una tabla dinámica
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja ||Hoja 1|
|Nombre de la tabla dinámica ||Name: |
|Campo a agregar||Field: |
|Selecciona una opción|||
|Selecciona una funcióm|||
|Nombre de campo ||Suma de Ventas|

### Remover campo
  
Remover un campo de una tabla dinámica
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja ||Hoja 1|
|Nombre de la tabla dinámica ||Name: |
|Campo a remover||Field: |

### Filtrar
  
Filtra una tabla dinámica
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja ||Hoja 1|
|Nombre de la tabla dinámica ||Name: |
|Nombre del filtro ||Campo |
|Limpiar todos los filtros|||
|Valor(es) del filtro a marcar ||Name: |
|Valor(es) del filtro a desmarcar ||Name: |

### Filtrar Valores
  
Filtra valores de una tabla dinámica
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja ||Hoja 1|
|Nombre de la tabla dinámica ||Nombre: |
|Nombre del campo base ||Campo base: |
|Campo al cual aplicar el filtro.||Campo: |
|Limpiar todos los filtros|||
|Seleccione filtro |Tipo de filtro a aplicar.|xlValueEquals|
|Valor(es) del filtro a marcar ||Name: |

### Listar Campos
  
Lista todos los campos disponibles
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja ||Hoja 1|
|Nombre de la tabla dinámica ||Name: |
|Asignar resultado a variable||Variable|

### Cambiar origen de datos
  
Cambia el rango de origen de los datos de una tabla dinámica
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja ||Hoja 1|
|Nombre de la tabla dinámica ||Name: |
|Nuevo Rango||A1:R200|

### Lista items de filtro
  
Devuelve una lista con los items de un filtro
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja ||Hoja 1|
|Nombre de la tabla dinámica ||Name: |
|Nombre del filtro ||Campo |
|Asignar resultado a variable||Variable|

### Insertar Línea de tiempo
  
Crea una nueva línea de tiempo
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja ||Hoja 1|
|Nombre de la tabla dinámica ||TablaDinámica1: |
|Campo de la tabla dinámica ||Campo |
|Rango donde posicionar||A1:D20|

### Filtrar Slider
  
Modifica el filtro de un slider
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja ||Hoja 1|
|Nombre del slider ||Name: |
|Fecha de incio||13/12/1999: |
|Fecha final||13/12/2000: |

### Estado de filtro
  
Retorna True si el filtro está marcadoa
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja ||Hoja 1|
|Nombre de la tabla dinámica ||Name: |
|Nombre del filtro ||Campo |
|Nombre del filtro ||Name: |
|Asignar resultado a variable||Variable|

### Cambiar tabla a formato tabular
  
Altera los campos de la tabla dinamica a formato tabular.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja ||Hoja 1|
|Nombre de la tabla ||Name: |
|Campos de la tabla||['Numero', 'Fecha', 'Hora']|

### Borrar subtotales
  
Elimina los subtotales de la tabla dinamica.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja ||Hoja 1|
|Nombre de la tabla ||Name: |
|Campos de la tabla||['Numero', 'Fecha', 'Hora']|

### Repetir etiquetas de elementos
  
Permite a la tabla dinamica repetir etiquetas.
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja ||Hoja 1|
|Nombre de la tabla ||Name: |
|Campos de la tabla||['Numero', 'Fecha', 'Hora']|

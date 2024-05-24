



# Tabelas dinâmicas
  
Modulo para trabalhar e interagir com tabelas dinâmicas do Microsoft Excel.  

![banner](imgs/Banner_PivotTableExcel.png)

## Como instalar este módulo
  
Para instalar o módulo no Rocketbot Studio, pode ser feito de duas formas:
1. Manual: __Baixe__ o arquivo .zip e descompacte-o na pasta módulos. O nome da pasta deve ser o mesmo do módulo e dentro dela devem ter os seguintes arquivos e pastas: \__init__.py, package.json, docs, example e libs. Se você tiver o aplicativo aberto, atualize seu navegador para poder usar o novo módulo.
2. Automático: Ao entrar no Rocketbot Studio na margem direita você encontrará a seção **Addons**, selecione **Install Mods**, procure o módulo desejado e aperte instalar.  


## Descrição do comando

### 
  

|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Data range |Enter the data range you want to use to create the pivot table|Sheet1!B2:C4|
|Destination Cell |Enter the cell where you want the pivot table to be created|Sheet2!C4|
|Pivot table name |Enter the pivot table name|Name: |

### 
  

|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sheet |Name of the sheet where the pivot table is located|Sheet1|
|Refresh all ||False|
|Pivote table name |Name of the pivot table to update|Name: |

### 
  

|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sheet |Name of the sheet where the pivot table is located|Sheet1|
|Pivot table name |Pivot table name|Name: |
|Field to add ||Field: |
|Select option|Select option to add a field to the pivot table|Add Data|
|Select a function|||
|Field name ||Sales Sum|

### 
  

|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sheet ||Sheet1|
|Pivot table name ||Name: |
|Field to remove |Name of the field to add|Field: |

### 
  

|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sheet |Name of the sheet where the pivot table is located|Sheet1|
|Pivot table name |Pivot table name|Name: |
|Filter name |Name of the field to be filtered|Field |
|Clear All Filters|If selected, all filters of the pivot table will be cleared|True|
|Filter(s) name to check |Name of the filter value to be checked|Name: |
|Filter(s) name to uncheck |Name of the filter value to be unchecked|Name: |

### 
  

|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sheet ||Sheet1|
|Pivot table name ||Name: |
|Base field name ||Base field: |
|Field to apply the filter to.||Field: |
|Clear All Filters|||
|Select filter |Type of filter to apply.|xlValueEquals|
|Filter(s) name to check ||Name: |

### 
  

|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sheet |Sheet name where the pivot table is located|Sheet1|
|pivot table name |Pivot table name|Name: |
|Assign result to variable |Variable name where the result will be stored|Variable|

### 
  

|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sheet |Name of the sheet where the pivot table is located|Sheet1|
|Pivot table name |Pivot table name|Name: |
|New Range |Pivot table data range|Sheet1!A1:R200|

### 
  

|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sheet |Sheet name|Sheet1|
|Pivot table name |Pivot table name|Name: |
|Filter name |Filter name|Field |
|Assign result to variable |Variable name to store the result|Variable|

### Inserir linha do tempo
  
Este comando cria uma nova linha do tempo
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Planilha|Nome da planilha onde a linha do tempo será inserida|Hoja 1|
|Nome do tabela dinâmica |Nome da tabela dinâmica que será usada para criar a linha do tempo|TabelaDinâmica1: |
|Campo do tabela dinâmica |Nome do campo do tabela dinâmica que será usado para criar a linha do tempo|Campo |
|Interválo onde posicionar|Interválo onde a linha do tempo será inserida|A1:D20|

### 
  

|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sheet |Name of the sheet where the slider is located|Sheet1|
|Slider name |Slider name|Name: |
|Start date|Start date of the filter|13/12/1999: |
|End date|End date of the filter|13/12/2000: |

### Status do filtro
  
Retorna True se o filtro estiver marcado
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sheet |Sheet name where the filter is located|Sheet1|
|Pivot table name |Pivot table name where the filter is located|Name: |
|Field name |Filter name to be consulted|Field |
|Filter element to check|Field filter value to be checked|Value: |
|Assign result to variable ||Variable|

### 
  

|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sheet |Name of the sheet where the pivot table is located.|Sheet1|
|Table name |Name of the pivot table.|Name: |
|Pivot fields|Pivot table fields.|['Number', 'Date', 'Hours']: |

### 
  

|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sheet |Name of the sheet where the pivot table is located.|Sheet1|
|Table name |Name of the pivot table.|Name: |
|Pivot fields|Pivot table fields.|['Number', 'Date', 'Hours']: |

### 
  

|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sheet |Name of the sheet where the pivot table is located.|Sheet1|
|Table name |Name of the pivot table.|Name: |
|Pivot fields|Pivot table fields that will be repeated.|['Number', 'Date', 'Hours']: |

### 
  

|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sheet |Excel Sheet name where the pivot table is located|Sheet1|
|Pivot table name |Pivot table name|MyTable|
|Field |Name of the field where the record to expand is located (Active Field)|Month|
|Item to expand|Name of the item to expand as it appears in the List items command|January|

# Pivot Table 
  
Module to work and interact with Pivot Tables  

*Read this in other languages: [English](Manual_PivotTableExcel.md), [Español](Mannual_PivotTableExcel.es.md).*
  
![banner](imgs/Banner_PivotTableExcel.png)
## How to install this module
  
__Download__ and __install__ the content in 'modules' folder in Rocketbot path  



## Description of the commands

### Create
  
Create a new pivot table
|Parameters|Description|example|
| --- | --- | --- |
|Data range |Enter the data range you want to use to create the pivot table|Sheet1!B2:C4|
|Destination Cell |Enter the cell where you want the pivot table to be created|Sheet2!C4|
|Pivot table name |Enter the pivot table name|Name: |

### Refresh
  
Refresh a pivot table
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the pivot table is located|Sheet1|
|Pivote table name |Name of the pivot table to update|Name: |

### Add field
  
Add field to a pivot table
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the pivot table is located|Sheet1|
|Pivot table name |Pivot table name|Name: |
|Select option|Select option to add a field to the pivot table|Add Data|
|Field to add |Name of the field to add|Field: |

### Filter
  
Filter a pivot table
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the pivot table is located|Sheet1|
|Pivot table name |Pivot table name|Name: |
|Filter name |Name of the field to be filtered|Field |
|Clear All Filters|If selected, all filters of the pivot table will be cleared|True|
|Filter(s) name to check |Name of the filter value to be checked|Name: |
|Filter(s) name to uncheck |Name of the filter value to be unchecked|Name: |

### List Fields
  
List all available table fields
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Sheet name where the pivot table is located|Sheet1|
|pivot table name |Pivot table name|Name: |
|Assign result to variable |Variable name where the result will be stored|Variable|

### Change pivot table data
  
Change data range of Pivot Table
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the pivot table is located|Sheet1|
|Pivot table name |Pivot table name|Name: |
|New Range |Pivot table data range|A1:R200|

### List Filter Items 
  
Return all items from filter
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Sheet name|Sheet1|
|Pivot table name |Pivot table name|Name: |
|Filter name |Filter name|Field |
|Assign result to variable |Variable name to store the result|Variable|

### Insert timeline
  
Create a new timeline.
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the timeline will be inserted|Plan1|
|Pivote table name |Name of the pivot table that will be used to create the timeline|PivotTable1: |
|Pivot table field |Name of the pivot table field that will be used to create the timeline|Field |
|Position range |Range where the timeline will be inserted|A1:D20|

### Filter slider
  
Sets the timeline's filter.
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the slider is located|Sheet1|
|Slider name |Slider name|Name: |
|Start date|Start date of the filter|13/12/1999: |
|End date|End date of the filter|13/12/2000: |

### Filter status
  
Returns True if the filter is checked
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Sheet name where the filter is located|Sheet1|
|Pivot table name |Pivot table name where the filter is located|Name: |
|Filter name |Filter name to be consulted|Field |
|Filter value|Filter value to be consulted|Name: |
|Assign result to variable ||Variable|

### Change table format to tabular
  
Changes the field of the pivot table to tabular format.
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the pivot table is located.|Sheet1|
|Table name |Name of the pivot table.|Name: |
|Pivot fields|Pivot table fields.|['Number', 'Date', 'Hours']: |

### Delete subtotals
  
Erase the subtotals from the pivot table.
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the pivot table is located.|Sheet1|
|Table name |Name of the pivot table.|Name: |
|Pivot fields|Pivot table fields.|['Number', 'Date', 'Hours']: |

### Repets label
  
Allows to the pivot table to repet labels.
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the pivot table is located.|Sheet1|
|Table name |Name of the pivot table.|Name: |
|Pivot fields|Pivot table fields that will be repeated.|['Number', 'Date', 'Hours']: |

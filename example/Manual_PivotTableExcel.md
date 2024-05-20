



# Pivot tables
  
Module to work and interact with Pivot Tables from Microsoft Excel.  

![banner](imgs/Banner_PivotTableExcel.png)

## How to install this module
  
To install the module in Rocketbot Studio, it can be done in two ways:
1. Manual: __Download__ the .zip file and unzip it in the modules folder. The folder name must be the same as the module and inside it must have the following files and folders: \__init__.py, package.json, docs, example and libs. If you have the application open, refresh your browser to be able to use the new module.
2. Automatic: When entering Rocketbot Studio on the right margin you will find the **Addons** section, select **Install Mods**, search for the desired module and press install.  


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
|Refresh all ||False|
|Pivote table name |Name of the pivot table to update|Name: |

### Add field
  
Add field to a pivot table
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the pivot table is located|Sheet1|
|Pivot table name |Pivot table name|Name: |
|Field to add ||Field: |
|Select option|Select option to add a field to the pivot table|Add Data|
|Select a function|||
|Field name ||Sales Sum|

### Remove field
  
Remove a field from a pivot table
|Parameters|Description|example|
| --- | --- | --- |
|Sheet ||Sheet1|
|Pivot table name ||Name: |
|Field to remove |Name of the field to add|Field: |

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

### Filter Values
  
Filter a pivot table values
|Parameters|Description|example|
| --- | --- | --- |
|Sheet ||Sheet1|
|Pivot table name ||Name: |
|Base field name ||Base field: |
|Field to apply the filter to.||Field: |
|Clear All Filters|||
|Select filter |Type of filter to apply.|xlValueEquals|
|Filter(s) name to check ||Name: |

### List Fields
  
List all available table fields
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Sheet name where the pivot table is located|Sheet1|
|pivot table name |Pivot table name|Name: |
|Assign result to variable |Variable name where the result will be stored|Variable|

### Change pivot table data
  
Change the source data range of a pivot table
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the pivot table is located|Sheet1|
|Pivot table name |Pivot table name|Name: |
|New Range |Pivot table data range|Sheet1!A1:R200|

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
  
Checks whether the element is marked in the field filter as visible or not visible. Return True or False respectively.
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Sheet name where the filter is located|Sheet1|
|Pivot table name |Pivot table name where the filter is located|Name: |
|Field name |Filter name to be consulted|Field |
|Filter element to check|Field filter value to be checked|Value: |
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

### Repeat label
  
Allows to the pivot table to repeat labels.
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the pivot table is located.|Sheet1|
|Table name |Name of the pivot table.|Name: |
|Pivot fields|Pivot table fields that will be repeated.|['Number', 'Date', 'Hours']: |

### Ungroup item
  
Ungroups the item indicated
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Excel Sheet name where the pivot table is located|Sheet1|
|Pivot table name |Pivot table name|MyTable|
|Field in which the item to ungroup is located.|Name of the field where the record to ungroup is located|Month|
|Item to ungroup|Name of the item to ungroup as it appears in the List items command|January|

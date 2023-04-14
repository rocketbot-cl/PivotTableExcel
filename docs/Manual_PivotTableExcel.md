



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
|Data range ||Sheet1!B2:C4|
|Destination Cell ||Sheet2!C4|
|Pivot table name ||Name: |

### Refresh
  
Refresh a pivot table
|Parameters|Description|example|
| --- | --- | --- |
|Sheet ||Sheet1|
|Refresh all ||False|
|Pivote table name ||Name: |

### Add field
  
Add field to a pivot table
|Parameters|Description|example|
| --- | --- | --- |
|Sheet ||Sheet1|
|Pivot table name ||Name: |
|Field to add ||Field: |
|Select option|||
|Select a function|||
|Field name ||Sales Sum|

### Remove field
  
Remove a field from a pivot table
|Parameters|Description|example|
| --- | --- | --- |
|Sheet ||Sheet1|
|Pivot table name ||Name: |
|Field to remove ||Field: |

### Filter
  
Filter a pivot table
|Parameters|Description|example|
| --- | --- | --- |
|Sheet ||Sheet1|
|Pivot table name ||Name: |
|Filter name ||Field |
|Clear All Filters|||
|Filter(s) name to check ||Name: |
|Filter(s) name to uncheck ||Name: |

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
|Sheet ||Sheet1|
|pivot table name ||Name: |
|Assign result to variable ||Variable|

### Change pivot table data
  
Change data range of Pivot Table
|Parameters|Description|example|
| --- | --- | --- |
|Sheet ||Sheet1|
|Pivot table name ||Name: |
|New Range ||A1:R200|

### List Filter Items 
  
Return all items from filter
|Parameters|Description|example|
| --- | --- | --- |
|Sheet ||Sheet1|
|Pivot table name ||Name: |
|Filter name ||Field |
|Assign result to variable ||Variable|

### Insert timeline
  
Create a new timeline.
|Parameters|Description|example|
| --- | --- | --- |
|Sheet ||Plan1|
|Pivote table name ||PivotTable1: |
|Pivot table field ||Field |
|Position range ||A1:D20|

### Filter slider
  
Sets the timeline's filter.
|Parameters|Description|example|
| --- | --- | --- |
|Sheet ||Sheet1|
|Slider name ||Name: |
|Start date||13/12/1999: |
|End date||13/12/2000: |

### Filter status
  
Returns True if the filter is checked
|Parameters|Description|example|
| --- | --- | --- |
|Sheet ||Sheet1|
|Pivot table name ||Name: |
|Filter name ||Field |
|Filter Name ||Name: |
|Assign result to variable ||Variable|

### Change table format to tabular
  
Changes the field of the pivot table to tabular format.
|Parameters|Description|example|
| --- | --- | --- |
|Sheet ||Sheet1|
|Table name ||Name: |
|Pivot fields||['Number', 'Date', 'Hours']: |

### Delete subtotals
  
Erase the subtotals from the pivot table.
|Parameters|Description|example|
| --- | --- | --- |
|Sheet ||Sheet1|
|Table name ||Name: |
|Pivot fields||['Number', 'Date', 'Hours']: |

### Repets label
  
Allows to the pivot table to repet labels.
|Parameters|Description|example|
| --- | --- | --- |
|Sheet ||Sheet1|
|Table name ||Name: |
|Pivot fields||['Number', 'Date', 'Hours']: |

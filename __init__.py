# coding: utf-8
"""
Base para desarrollo de modulos externos.
Para obtener el modulo/Funcion que se esta llamando:
     GetParams("module")

Para obtener las variables enviadas desde formulario/comando Rocketbot:
    var = GetParams(variable)
    Las "variable" se define en forms del archivo package.json

Para modificar la variable de Rocketbot:
    SetVar(Variable_Rocketbot, "dato")

Para obtener una variable de Rocketbot:
    var = GetVar(Variable_Rocketbot)

Para obtener la Opcion seleccionada:
    opcion = GetParams("option")


Para instalar librerias se debe ingresar por terminal a la carpeta "libs"

    pip install <package> -t .

"""
# Changing the data types of all strings in the module at once
from __future__ import unicode_literals
import os
import sys
import traceback
# from string import ascii_letters
# result32 = stringMod.ascii_letters
# print(result32)


base_path = tmp_global_obj["basepath"]
cur_path = os.path.join(base_path, 'modules', 'PivotTableExcel', 'libs')

cur_path_x64 = os.path.join(cur_path, 'Windows' + os.sep +  'x64' + os.sep)
cur_path_x86 = os.path.join(cur_path, 'Windows' + os.sep +  'x86' + os.sep)

if sys.maxsize > 2**32 and cur_path_x64 not in sys.path:
        sys.path.append(cur_path_x64)
if sys.maxsize <= 2**32 and cur_path_x86 not in sys.path:
        sys.path.append(cur_path_x86)

global ascii_letters
ascii_letters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"

def column_to_number(col):
    num = 0
    for c in col:
        if c in ascii_letters:
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num

def number_to_column(n):
    string2 = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string2 = chr(65 + remainder) + string2
    return string2

constants = {"xlRowField": 1, "xlColumnField": 2, "xlPageField": 3}
functions = {"xlSum": -4157, "xlCount": -4112, "xlAverage": -4106, "xlProduct": -4149, "xlMax": -4136, "xlMin": -4139}
abc = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v',
           'w', 'x', 'y', 'z']

excel = GetGlobals("excel")
module = GetParams("module")

try:
    if module == "createPivotTable":
        from openpyxl.utils import column_index_from_string
        import re
        
        data = GetParams("data")
        destination = GetParams("destination")
        table_name = GetParams("tableName")

        xls = excel.file_[excel.actual_id]
        wb = xls['workbook']
        # wb = xw.Book("ruta")

        data = data.replace('$', '').split("!")
        destination = destination.replace('$', '').split("!")
        sheet_1, sheet_2 = None, None
        
        if len(data) > 1:
            sheet = data[0]
            data = data[1]
        else:
            sheet = 1
            data = data[0]
            
        sheet_1 = wb.api.Worksheets(sheet)
        if len(destination) > 1:
            pivot_sheet = destination[0]
            cell = destination[1]
            if sheet != pivot_sheet:
                sheet_2 = wb.api.Worksheets(destination[0])
        else:
            cell = destination[0]
            
        source_range = sheet_1.Range(data)

        if sheet_2:
            pivotTargetRange = sheet_2.Range(cell)
        else:
            pivotTargetRange = sheet_1.Range(cell)

        pivot_table = wb.api.PivotCaches().Create(SourceType=1, SourceData=source_range)
        pivot_table.CreatePivotTable(TableDestination=pivotTargetRange, TableName=table_name)

    if module == "refreshPivot":
        sheet = GetParams("sheet")
        pivotTableName = GetParams("table")
        refresh_all = GetParams("all")
        
        xls = excel.file_[excel.actual_id]
        wb = xls['workbook']
        ws = wb.sheets[sheet]
        if pivotTableName:
            for table in ws.api.PivotTables():
                if table._inner() == pivotTableName:
                    table.RefreshTable()
                    break
        
        if refresh_all and eval(refresh_all)==True:
            for table in ws.api.PivotTables():
                table.RefreshTable()

    if module == "addField":

        sheet = GetParams("sheet")
        pivotTableName = GetParams("table")
        data = GetParams("data")
        
        option = GetParams("option_")
        name = GetParams("data_field_name")
        func = GetParams("data_field_func")
        
        xls = excel.file_[excel.actual_id]

        wb = xls['workbook']
        # wb = xw.Book("ruta")
        sht = wb.sheets[sheet].select()

        pivot_table = wb.api.ActiveSheet.PivotTables(pivotTableName)
        
        fields_names = [field.Name for field in pivot_table.PivotFields()]
        
        for d in data.split(","):
            cubeField = pivot_table.PivotFields(d)
            if option != "data":
                cubeField.Orientation = constants[option]
                cubeField.Position = 1
            else:
                name_ = name if name else option.strip('xl') + " {value}".format(value=d)
                if name_ in fields_names:
                    raise Exception("Cannot have the same name as one of the source fields, choose another...")

                field = pivot_table.PivotFields(d)
                pivot_table.AddDataField(field, name_, functions[func])
                
                # field = pivot_table.PivotFields("Suma de {value}".format(value=d))
                # field.Function = -4157
    
    if module == "removeField":
        sheet = GetParams("sheet")
        pivotTableName = GetParams("table")
        data = GetParams("data")
        
        xls = excel.file_[excel.actual_id]
        wb = xls['workbook']
        sht = wb.sheets[sheet].select()

        pivot_table = wb.api.ActiveSheet.PivotTables(pivotTableName)
        
        # Get the fields of the pivot table
        row_fields = [field.Name for field in pivot_table.RowFields]
        column_fields = [field.Name for field in pivot_table.ColumnFields]
        data_fields = [field.Name for field in pivot_table.DataFields]
        page_fields = [field.Name for field in pivot_table.PageFields]
        fields_names = row_fields + column_fields + data_fields + page_fields
                
        field = pivot_table.PivotFields(data)
        if field.Name in fields_names:
            field.Orientation = 0
        else:
            raise Exception("Can't find field...")
        
    if module == "filter":

        sheet = GetParams("sheet")
        pivotTableName = GetParams("table")
        data = GetParams("filter")
        check = GetParams("value")
        no_check = GetParams("noCheck")
        clean = GetParams("clean")

        xls = excel.file_[excel.actual_id]

        wb = xls['workbook']
        sht = wb.sheets[sheet].select()

        pivotTable = wb.api.ActiveSheet.PivotTables(pivotTableName)
        filter_ = pivotTable.PivotFields(data)

        if clean is not None:
            clean = eval(clean)
        if clean:
            filter_.ClearAllFilters()

        if check:
            check = eval(check) if check.startswith("[") else check.split(",")

            for item in check:
                filter_.PivotItems(item).Visible = True

            if not no_check:
                for item in filter_.PivotItems():
                    if item.Name not in check:
                        filter_.PivotItems(item.Name).Visible = False
                
        if no_check:
            no_check = eval(no_check) if no_check.startswith("[") else no_check.split(",")
            
            if not check:
                for item in filter_.PivotItems():
                    if item.Name not in no_check:
                        filter_.PivotItems(item.Name).Visible = True
            
            for item in no_check:
                filter_.PivotItems(item).Visible = False
    
    if module == "filter_value":

        sheet = GetParams("sheet")
        pivotTableName = GetParams("table")
        data = GetParams("filter")
        field = GetParams("field")
        check = GetParams("value")
        clean = GetParams("clean")
        filter_type = GetParams("filter_type")

        xls = excel.file_[excel.actual_id]

        wb = xls['workbook']
        sht = wb.sheets[sheet].select()
        pivotTable = wb.api.ActiveSheet.PivotTables(pivotTableName)
        filter_ = pivotTable.PivotFields(data)

        if not filter_type and clean is not None:
            clean = eval(clean)
            if clean:
                filter_.ClearAllFilters()
        else:
            if clean is not None:
                clean = eval(clean)
            if clean:
                filter_.ClearAllFilters()
            
            filter_value = eval(check)
            filter_type = eval(filter_type)
            
            data_field = wb.api.ActiveSheet.PivotTables(pivotTableName).PivotFields(field)
            
            if filter_type in [13, 14] and isinstance(filter_value, list) and len(filter_value) == 2:
                value1 = filter_value[0]
                value2 = filter_value[1]
                filter_.PivotFilters.Add2(filter_type, data_field, value1, value2)
            else:
                filter_.PivotFilters.Add2(filter_type, data_field, filter_value)
            
            # ActiveSheet.PivotTables("TablaDinámica1").PivotFields("Nombre").PivotFilters. _
            # Add2 Type:=xlValueIsGreaterThan, DataField:=ActiveSheet.PivotTables( _
            # "TablaDinámica1").PivotFields("Suma de Horas"), Value1:=100

    if module == "listFields":
        sheet = GetParams("sheet")
        pivotTableName = GetParams("table")
        result = GetParams("result")

        xls = excel.file_[excel.actual_id]

        wb = xls['workbook']
        try:
            sht = wb.sheets[sheet].select()
        except:
            pass

        pivotTable = wb.api.ActiveSheet.PivotTables(pivotTableName)

        cubeFields = pivotTable.PivotFields()

        fields = [field.Name for field in cubeFields]

        SetVar(result, fields)

    if module == "changeOrigin":
        sheet = GetParams("sheet")
        pivotTableName = GetParams("table")
        range_ = GetParams("range")

        xls = excel.file_[excel.actual_id]
        wb = xls['workbook']
        sh = wb.sheets[sheet]
        sh.select()

        range_ = range_.replace('$', '')
        
        if "!" in range_:
            sheet, range_ = range_.split("!")
            source_range = wb.sheets[sheet].api.Range(range_)
        else:
            source_range = sh.api.Range(range_)
        pivot = wb.api.ActiveSheet.PivotTables(pivotTableName)
        pivot_table = wb.api.PivotCaches().Create(SourceType=1, SourceData=source_range, Version=5)
        pivot.ChangePivotCache(pivot_table)


    if module == "getItems":

        sheet = GetParams("sheet")
        pivotTableName = GetParams("table")
        data = GetParams("filter")
        result = GetParams("result")

        xls = excel.file_[excel.actual_id]

        wb = xls['workbook']
        try:
            sht = wb.sheets[sheet].select()
        except:
            pass
        
        pivotTable = wb.sheets[sheet].api.PivotTables(pivotTableName)
        filter_ = pivotTable.PivotFields(data)
        # filter_.CurrentPage = "(All)"
        items = [item.Name for item in filter_.PivotItems()]
        if result:
            SetVar(result, items)

    if module == "filter_slider":

        sheet_name = GetParams("sheet")
        slider_name = GetParams("name")
        start_date = GetParams("start")
        end_date = GetParams("end")

        xls = excel.file_[excel.actual_id]
        wb = xls['workbook']
        sheet = wb.sheets[sheet_name]
        sheet.select()
        wb.api.SlicerCaches(slider_name).TimelineState.SetFilterDateRange(start_date, end_date)

    if module == "create_slider":
        sheet_name = GetParams("sheet")
        pivotTableName = GetParams("table")
        field = GetParams("field")
        position = GetParams("range")

        xls = excel.file_[excel.actual_id]
        wb = xls['workbook']

        if not sheet_name in [sh.name for sh in wb.sheets]:
            raise Exception(f"The name {sheet_name} does not exist in the book")

        sheet = wb.sheets[sheet_name]
        pivot_table = wb.api.ActiveSheet.PivotTables(pivotTableName)

        start = position
        end = None
        if ":" in position:
            cells = position.split(":")
            start = cells[0]
            end = cells[1]

        top = sheet.range(start).api.Cells.Top
        left = sheet.range(start).api.Cells.Left
        width = 262.5
        height = 108
        print(top, left, height, width)
        if end is not None:
            width = sheet.range(end).api.Cells.Left - left
            height = sheet.range(end).api.Cells.Top - top

        sheet.select()
        macro = f"""
Sub RocketAddSlider()
'
' RocketAddSlider Macro
'

'
ActiveWorkbook.SlicerCaches.Add2(ActiveSheet.PivotTables("{pivotTableName}"), _
        "{field}", , xlTimeline).Slicers.Add ActiveSheet, , "{field}", "{field}", {top} _
        , {left}, {width}, {height}
End Sub"""

        try:
            m = wb.macro("RocketAddSlider")
            m.run()
        except:
            tmp = wb.api.VBProject.VBComponents.Add(1)
            tmp.CodeModule.AddFromString(macro)
            m = wb.macro("RocketAddSlider")
            m.run()


    if module == "visible":

        sheet = GetParams("sheet")
        pivotTableName = GetParams("table")
        data = GetParams("filter")
        field = GetParams("value")
    
        result = GetParams("result")

        xls = excel.file_[excel.actual_id]
        wb = xls['workbook']
        try:
            sht = wb.sheets[sheet].select()
        except:
            pass

        pivotTable = wb.api.ActiveSheet.PivotTables(pivotTableName)
        filter_ = pivotTable.PivotFields(data)
        is_visible = filter_.PivotItems(field).Visible
        if result:
            SetVar(result, is_visible)

    if module == "pivot_table_tabular":


        
        sheet_name = GetParams("sheet")
        pivotTableName = GetParams("table")
        fields = ""
        try:
            fields = eval(GetParams("fields"))
        except:
            pass

        xls = excel.file_[excel.actual_id]
        wb = xls['workbook']

        if not sheet_name in [sh.name for sh in wb.sheets]:
            raise Exception(f"The name {sheet_name} does not exist in the book")

        sheet = wb.sheets[sheet_name].select()
        pivot_table = wb.api.ActiveSheet.PivotTables(pivotTableName).PivotFields()
        if fields != "":
            for cada in fields:
                for each in pivot_table:
                    if each.Name == cada:
                        each.LayoutForm = 0

    if module == "pivot_table_delete_subtotal":


        
        sheet_name = GetParams("sheet")
        pivotTableName = GetParams("table")
        fields = ""
        try:
            fields = eval(GetParams("fields"))
        except:
            pass

        xls = excel.file_[excel.actual_id]
        wb = xls['workbook']

        if not sheet_name in [sh.name for sh in wb.sheets]:
            raise Exception(f"The name {sheet_name} does not exist in the book")

        sheet = wb.sheets[sheet_name].select()
        pivot_table = wb.api.ActiveSheet.PivotTables(pivotTableName).PivotFields()
        if fields != "":
            for cada in fields:
                for each in pivot_table:
                    if each.Name == cada:
                        each.Subtotals = (False, False, False, False, False, False, False, False, False, False, False, False)

    if module == "pivot_table_repet_labels":

        sheet_name = GetParams("sheet")
        pivotTableName = GetParams("table")
        fields = ""
        try:
            fields = eval(GetParams("fields"))
        except:
            pass

        xls = excel.file_[excel.actual_id]
        wb = xls['workbook']

        if not sheet_name in [sh.name for sh in wb.sheets]:
            raise Exception(f"The name {sheet_name} does not exist in the book")

        sheet = wb.sheets[sheet_name].select()
        pivot_table = wb.api.ActiveSheet.PivotTables(pivotTableName).PivotFields()
        if fields != "":
            for cada in fields:
                for each in pivot_table:
                    if each.Name == cada:
                        each.RepeatLabels = True

except Exception as e:
    traceback.print_exc()
    print("\x1B[" + "31;40mError\x1B[" + "0m")
    PrintException()
    raise e


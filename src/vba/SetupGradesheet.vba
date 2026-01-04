Sub SetupGradesheet()
'
' SetupGradesheet Macro
'

'
    ActiveSheet.Next.Select
    Selection.End(xlToRight).Select
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Progress"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = _
        "=LET(total, AGGREGATE(3, 5, [Question]), filled, AGGREGATE(3, 5, [Deductions]), filled / total)"
    Range("Marks[Progress]").Select
    Range("I3").Activate
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.0%"
    Selection.NumberFormat = "0.00%"
    Range("Marks[Deductions]").Select
    Range("F3").Activate
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="0", Formula2:= _
        "=INDIRECT(""RC[-1]"", FALSE)"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("B2").Select
    Application.CutCopyMode = False
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Marks", Version:=8).CreatePivotTable TableDestination:="Sheet1!R3C1", _
        TableName:="PivotTable1", DefaultVersion:=8
    Sheets("Sheet1").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable1")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable1").RepeatAllLabels xlRepeatLabels
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "Grades"
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("OrgDefinedID")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Question")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Score")
        .Orientation = xlRowField
        .Position = 3
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Question")
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Score"), "Count of Score", xlCount
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Count of Score")
        .Caption = "Sum of Score"
        .Function = xlSum
    End With
    ActiveSheet.Next.Select
    Range("B2").Select
End Sub

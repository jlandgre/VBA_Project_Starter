Attribute VB_Name = "Populate"
' Version 12/3/24
Option Explicit
'-----------------------------------------------------------------------------------------
' Populate DemoStudy procs.OverallProject and procs.DemoStudy validation data
'-----------------------------------------------------------------------------------------------
'Populate .Inputs table (Assumes .InitAllTables has been run to Provision tbls)
'JDL 12/3/24
Sub PopulateInputsTbl(tbls)
    Dim lstVals() As Variant
    
    With tbls.Inputs
        If Not .rngRows Is Nothing Then Intersect(.rngRows, .rngTable).ClearContents
        
        ReDim lstVals(1 To .ncols)
        lstVals(1) = "A,B,C,D"
        lstVals(2) = "Something 1,Something 2,Something 3,Something 4"
        lstVals(3) = "10,20,30,40"
        PopulateByCols tbls.Inputs, lstVals, 4
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'Populate .RowsCols table (Assumes .InitAllTables has been run to Provision tbls)
'JDL 12/3/24
Sub PopulateRowsColsTbl(tbls)
    Dim lstVals() As Variant
    
    With tbls.RowsCols
    
        'Clear previous contents and column(s) potentially added by ExcelSteps refresh
        If Not .rngRows Is Nothing Then
            Intersect(.rngRows, .rngTable).Clear
            .wksht.Columns(3).Clear
        End If
        
        ReDim lstVals(1 To 2)
        lstVals(1) = "1,2,3,4,5,6,7,8"
        lstVals(2) = "A,A,B,B,C,C,D,D"
        PopulateByCols tbls.RowsCols, lstVals, 8
    End With
End Sub
'-----------------------------------------------------------------------------------------------
' Utility subs
'-----------------------------------------------------------------------------------------------
' Helper sub to populate the table based on LstVals array
' JDL 10/7/24; Modified 12/3/24
'
Sub PopulateByCols(tbl, lstVals, nrows)
    Dim i As Integer, iCol As Integer, rng As Range, ary As Variant
    With tbl
        For i = 1 To UBound(lstVals)
            ary = Split(lstVals(i), ",")
            iCol = .cellHome.Column + i - 1
            Set rng = Range(.wksht.Cells(.cellHome.Row, iCol), .wksht.Cells(.cellHome.Row + nrows - 1, iCol))
            rng = WorksheetFunction.Transpose(ary)
        Next i
    End With
End Sub
'-----------------------------------------------------------------------------------------
' Clear previous table data values
' JDL 10/9/24
'
Sub ClearDefaultTableValues(tbl)
    With tbl
        .wksht.Activate
        If Not .rngRows Is Nothing And Not .rngTable Is Nothing Then _
            Intersect(.rngRows, .rngTable).ClearContents
        If .wksht.UsedRange.Rows.Count > 1 Then _
            Range(.wksht.Rows(2), .wksht.Rows(.wksht.UsedRange.Rows.Count)).Clear
    End With
End Sub

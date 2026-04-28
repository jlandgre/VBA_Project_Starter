Attribute VB_Name = "Utilities_Testing"
'This module contains testing-specific utilities
'Version 4/28/26
Option Explicit


'-----------------------------------------------------------------------------
'Purpose: Build a comma-separated list of a column range's contents
'
'Created: 11/23/21 JDL
'
Function LstColContents(wkbk, sht, iCol)
    Dim rng As Range
    With wkbk.Sheets(sht)
        Set rng = Range(.Cells(1, iCol), rngLastPopCell(.Cells(1, iCol), xlDown))
    LstColContents = ListFromArray(rng)
    End With
End Function
'-----------------------------------------------------------------------------------------------
' Determine whether specified range is a row range
'
Function IsRowRange(rng) As Boolean
    IsRowRange = rng.Address = rng.EntireRow.Address
End Function
'-----------------------------------------------------------------------------------------------
' Determine whether specified range is a column range
'
Function IsColRng(rng) As Boolean
    IsColRng = rng.Address = rng.EntireColumn.Address
End Function
'-----------------------------------------------------------------------------------------------
' Determine whether specified range is cell or block of cells (not column or row)
'
Function IsCellRng(rng) As Boolean
    IsCellRng = Not IsRowRange(rng) And Not IsColRng(rng)
End Function
'-----------------------------------------------------------------------------------------------
Sub ShadeYellow(rng)
    rng.Interior.Pattern = xlSolid
    rng.Interior.Color = 65535
End Sub
'-----------------------------------------------------------------------------------------------
' Reveal all worksheet cells by clearing outline, turning off filter and unhiding
'
'Created:   12/1/21 JDL
'
Sub RevealWkshtCells(wksht)
    With wksht
        .AutoFilterMode = False
        .Cells.ClearOutline
        .Cells.EntireColumn.Hidden = False
        .Cells.EntireRow.Hidden = False
    End With
End Sub
'-----------------------------------------------------------------------------
' Set Application environment for testing
' JDL 1/4/22
Sub SetApplEnvir(IsEvents, IsScreenUpdate, xlCalc)
    With Application
        .EnableEvents = IsEvents
        .ScreenUpdating = IsScreenUpdate
        .Calculation = xlCalc
    End With
End Sub
'-----------------------------------------------------------------------------------------------
' Count and delete stray worksheets if they are empty (Sheet1 etc.)
'
'Modified: 1/20/22 Used for testing tblRowsCols
'
Function iCountAndDeleteStraySheets(wkbk) As Integer
    Dim wksht As Variant
    iCountAndDeleteStraySheets = 0
    For Each wksht In wkbk.Sheets
        If LCase(Left(wksht.name, 5)) = "sheet" Then
            If wksht.UsedRange.Address = "$A$1" And IsEmpty(wksht.Cells(1, 1)) Then
                DeleteSheet wkbk, wksht.name
                iCountAndDeleteStraySheets = iCountAndDeleteStraySheets + 1
            End If
        End If
    Next wksht
End Function
'-----------------------------------------------------------------------------------------------
' Create a uni-dimension array from a specified Range row
' JDL 7/26/23
Function AryFromRowRng(rngRow, iRow) As Variant
    Dim i As Integer, aryRow As Variant, aryTemp() As Variant
    aryRow = rngRow.Value
    ReDim aryTemp(0 To UBound(aryRow, 2) - 1)
    
    For i = LBound(aryRow, 2) To UBound(aryRow, 2)
        aryTemp(i - 1) = aryRow(1, i)
    Next i
    AryFromRowRng = aryTemp
End Function
'-----------------------------------------------------------------------------------------------
' Create a uni-dimensional array from a specified Range column
' JDL 7/26/23
Function AryFromColRng(rngcol, iCol) As Variant
    Dim i As Integer, aryCol As Variant, aryTemp() As Variant
    aryCol = rngcol.Value
    ReDim aryTemp(0 To UBound(aryCol, 1) - 1)
    
    For i = LBound(aryCol, 1) To UBound(aryCol, 1)
        aryTemp(i - 1) = aryCol(i, 1)
    Next i
    AryFromColRng = aryTemp
End Function
'-----------------------------------------------------------------------------------------------
' Get wkbk name from its VBA Project name (for setting Test.wkbkTest)
' JDL 4/14/25; Updated with error message 11/4/25
'
Function GetWorkbookByVBProjectName(VBAProjectName As String) As Workbook
    Dim wkbk As Workbook, s As String
    
    ' Loop through all open workbooks and check for a match
    For Each wkbk In Application.Workbooks
        If wkbk.VBProject.name = VBAProjectName Then
            Set GetWorkbookByVBProjectName = wkbk
            Exit Function
        End If
    Next wkbk
    s = "Error setting wkbk by VBA Project Name: Specified name not found in open workbooks"
    MsgBox s, Title:="Testing Error"
    
    Set GetWorkbookByVBProjectName = Nothing
End Function


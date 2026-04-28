VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Procedures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Version 4/28/26
'This class defines project-specific groups of tests to run in test suite driver subs

' DemoStudy class
Public OverallProject As Object
Public DemoStudy As Object
Public ProductionSetup As Object

' General attributes for all procedures in the code project
Public wkbk_testing As Workbook 'Test suite workbook aka ThisWorkbook
Public wksht_results As Worksheet
Public test_suite_name As String
Public curProcedure As String 'Name of current procedure

'Column ranges on wksht_results
Public colrngProcedure As Range
Public colrngTestName As Range
Public colrngTestResult As Range
Public rngHeader As Range
'---------------------------------------------------------------------------------------
' Instance Procedure objects for each TestGroup
' JDL 5/2/25; Updated 4/28/26
Sub Init(procs, ByVal wkbk, ByVal shtT, test_suite_name)
    With procs
    
        'instance project-specific Procedure class instances
        Set .OverallProject = New Procedure
        .OverallProject.name = "OverallProject"
        Set .DemoStudy = New Procedure
        .DemoStudy.name = "DemoStudy"
        Set .ProductionSetup = New Procedure
        .ProductionSetup.name = "ProductionSetup"
         
        ' Set attributes and initialize test results sheet
        .test_suite_name = test_suite_name
        Set .wkbk_testing = wkbk
        
        ' Initialize the Test Result sheet including creating it if it's not present
        .InitTestsSheet procs, shtT

        Set .colrngProcedure = .wksht_results.Columns(1)
        Set .colrngTestName = .wksht_results.Columns(2)
        Set .colrngTestResult = .wksht_results.Columns(3)
        Set .rngHeader = Intersect(.wksht_results.Rows(1), _
            Range(.wksht_results.Columns(1), .wksht_results.Columns(3)))
                
        ' Clear the results columns and refresh header strings
        .colrngTestResult.Clear
        .rngHeader = Split("Procedure,Test Name,Test Result", ",")
        
        ' Clear previous yellow highlighting of failed results if any
        .colrngTestResult.Interior.Pattern = xlNone
        
        ColWidthAutofit Range(.colrngProcedure, .colrngTestResult)
    End With
End Sub
'-----------------------------------------------------------------------------------------------
'Initialize sheet for reporting results (clears previous results)
'Inputs: wkbk [Workbook] contains tests sheet
'        shtT [String] tests sheet name for reporting test results
'
'Created:   11/10/21 JDL; Modified 11/8/24
'
Sub InitTestsSheet(procs, ByVal shtT)
    
    'Add tests sheet if it does not exist and set Worksheet object
    With procs.wkbk_testing
        If Not SheetExists(procs.wkbk_testing, shtT) Then _
            AddSheet procs.wkbk_testing, shtT, .Sheets(.Sheets.Count).name
        Set procs.wksht_results = .Sheets(shtT)
    End With
    
    With procs
        Set .colrngProcedure = .wksht_results.Columns(1)
        Set .colrngTestName = .wksht_results.Columns(2)
        Set .colrngTestResult = .wksht_results.Columns(3)
        Set .rngHeader = Intersect(.wksht_results.Rows(1), _
        Range(.wksht_results.Columns(1), .wksht_results.Columns(3)))
                
        'Clear the results columns and refresh header strings
        .colrngTestResult.Clear
        .rngHeader = Split("Procedure,Test Name,Test Result", ",")
        
        'Clear previous yellow highlighting of failed results if any
        .colrngTestResult.Interior.Pattern = xlNone
        
        ColWidthAutofit Range(.colrngProcedure, .colrngTestResult)
    End With

End Sub
'-----------------------------------------------------------------------------------------------
' Evaluate overall test results
'
'Inputs:  wkbk [Workbook] workbook containing test results table
'         shtT [String] worksheet name containing test results table
'         sTestSuite [String] name of test suite
'
'   JDL 11/18/21; Modified 10/1/24
Sub EvalOverall(procs)
    Dim w As Variant, IsPass As Boolean, curCell As Range, sMsg As String, rngcol As Range
    
    With procs
        IsPass = True
        Set curCell = Intersect(.wksht_results.Rows(2), .colrngTestName)
        Do While Len(curCell.Value) > 0
                
            'w = curCell.Offset(0, 1).Value
            w = Intersect(curCell.EntireRow, .colrngTestResult).Value
            If Not IsEmpty(w) Then IsPass = IsPass And w
            Set curCell = curCell.Offset(1, 0)
        Loop
        
        'Report overall result to user
        sMsg = procs.test_suite_name & " Testing Failed"
        If IsPass Then sMsg = procs.test_suite_name & " Testing Passed"
        MsgBox sMsg
        
        'Autofit columns
        For Each rngcol In Range(.colrngProcedure, .colrngTestResult)
        With rngcol
            .ColumnWidth = 120
            .AutoFit
            If .ColumnWidth < 120 Then .ColumnWidth = .ColumnWidth + 2
        End With
    Next rngcol
    End With
End Sub
Public Function WkbkFromProjectName(sProjectName As String) As Workbook
    Dim wkbk As Workbook
    For Each wkbk In Application.Workbooks
        If wkbk.VBProject.name = sProjectName Then
            Set WkbkFromProjectName = wkbk
            Exit Function
        End If
    Next wkbk
End Function



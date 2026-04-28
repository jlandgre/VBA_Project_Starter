Attribute VB_Name = "Interface"
'Version 12/2/24
Option Explicit
'Subs activated by user buttons on Info sheet

'Clear all input data (validation test data)
Sub ClearInputTablesDemo()
    DemoProj.ClearInputTables
End Sub

'Populate input data
Sub PopulateInputsDemo()
    Dim procs As New Procedures
    SetApplEnvir False, False, xlCalculationManual
    procs.Init procs, ThisWorkbook, "DemoStudy", ""
    
    'test_PopulateForDemo procs
    
    SetApplEnvir True, True, xlCalculationAutomatic
End Sub




Attribute VB_Name = "Tests_Setup"
'Version 4/28/26
Option Explicit
'---------------------------------------------------------------------------------------
' Miscellaneous Setup tests for project
' (e.g. generate test data and setup for prodn)
' 11/13/25; Updated 4/28/26
'
Sub TestingDriver_Setup()
    Dim procs As New Procedures, AllEnabled As Boolean
    
    With procs
        .Init procs, ThisWorkbook, "Setup", "Tests_Setup"
        SetApplEnvir False, False, xlCalculationManual
        '.CreateTestData.Enabled = True ' Warning: This overwrites current test data!
        .ProductionSetup.Enabled = True
    End With
        
    'Create and import mockup data
    'With procs.CreateTestData
    '    If .Enabled Or AllEnabled Then
    '        procs.curProcedure = .name
    '        'test_xyz_setup_test_data procs
    '    End If
    'End With
        
    With procs.ProductionSetup
        If .Enabled Or AllEnabled Then
            procs.curProcedure = .name
            test_ExcelStepsVersion procs
            test_ErrorHandlingEnabled procs
            test_ClearTables procs
            test_DemoStudyVersion procs
            test_HideSheets procs
        End If
    End With
    
    SetApplEnvir True, True, xlCalculationAutomatic
    procs.EvalOverall procs
End Sub
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
'procs.ProductionSetup - "Tests that set the DemoStudy workbook for production
'-------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------
' Hide non user-facing sheets
' JDL 12/3/25
'
Sub test_HideSheets(procs)
    Dim tst As New Test: tst.Init tst, "test_HideSheets"
    Dim sht As Variant
    
    With tst
        For Each sht In Array(shtParams, shtErrors)
            .wkbkTest.Sheets(sht).Visible = xlSheetHidden
        Next sht
        .Update tst, procs
    End With
End Sub
'-------------------------------------------------------------------------------------
' Check version is set on Home (based on Constant in Constants module)
' JDL 4/28/26
'
Sub test_DemoStudyVersion(procs)
    Dim tst As New Test: tst.Init tst, "test_DemoStudyVersion"
        
    With tst
        .Assert tst, .wkbkTest.Sheets(shtHome).Cells(12, 1) = VersionDemoProj
        .Update tst, procs
    End With
End Sub
'-------------------------------------------------------------------------------------
' Clear Input Tables
' JDL 4/28/26
'
Sub test_ClearTables(procs)
    Dim tst As New Test: tst.Init tst, "test_ClearTables"
    Dim tbls As Object
    With tst
        
        'Turn off dialog request for user permission to clear
        DemoProj.IsTest = True
        DemoProj.ClearInputTables
        
        'Doublecheck no errors
        .Assert tst, Len(ExcelSteps.errs.errMsg) < 1
        .Update tst, procs
    End With
End Sub

'-------------------------------------------------------------------------------------
' Check ExcelSteps version in Info/Properties Comment (from Constant in Constants module)
' JDL 12/3/25
'
Sub test_ExcelStepsVersion(procs)
    Dim tst As New Test: tst.Init tst, "test_ExcelStepsVersion"
    Dim CommentXLAM As String, wkbk As Workbook
    
    With tst
        Set wkbk = Application.Workbooks("XLSteps.xlam")
        CommentXLAM = wkbk.BuiltinDocumentProperties("Comments").Value
        .Assert tst, CommentXLAM = VersionExcelSteps
        .Update tst, procs
    End With
End Sub
'-------------------------------------------------------------------------------------
' Check that error handling is enabled
' JDL 12/3/25
'
Sub test_ErrorHandlingEnabled(procs)
    Dim tst As New Test: tst.Init tst, "test_ErrorHandlingEnabled"
    Dim IsEnabled As Boolean
    
    With tst
    
        'Call SetErrs to initialize errs
        ExcelSteps.SetErrs "driver, .wkbktest"
        
        'Check that error handling is enabled
        .Assert tst, ExcelSteps.errs.IsHandle = True
        .Update tst, procs
    End With
End Sub



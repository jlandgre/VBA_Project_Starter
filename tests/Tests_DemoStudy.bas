Attribute VB_Name = "Tests_DemoStudy"
'Version 4/28/26
Option Explicit
'---------------------------------------------------------------------------------------
' Validate DemoStudy Class in VBAProject_DemoStudy separate file
' JDL 12/3/24
Sub TestingDriver_DemoStudy()
    Dim procs As New Procedures, AllEnabled As Boolean
    With procs
        .Init procs, ThisWorkbook, "DemoStudy", "DemoStudy_test_suite"
        SetApplEnvir False, False, xlCalculationAutomatic
        
        'Enable testing of all or individual procedures
        AllEnabled = True
        .OverallProject.Enabled = True
        .DemoStudy.Enabled = False
        '.OtherProcedureBlock.Enabled = False
    End With
    
    ' Tests of overall setup and Interface module
    With procs.OverallProject
        If .Enabled Or AllEnabled Then
            procs.curProcedure = .name
            test_InitTbls procs
            test_InitMdls procs
            test_ClearTables procs
            test_PopulateInputsTbl procs
            test_PopulateRowsColsTbl procs
        End If
    End With
    
    ' Tests of DemoStudy class
    With procs.DemoStudy
        If .Enabled Or AllEnabled Then
            procs.curProcedure = .name
            test_RefreshRowsColsTblProcedure procs
        End If
    End With
    
    'With procs.OtherProcedureBlock
        'If .Enabled Or AllEnabled Then
            'procs.curProcedure = .Name
            'test_xxx procs
        'End If
    'End With

    procs.EvalOverall procs
End Sub
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
' Tests of procs.OverallProcedure (Init and Populate routines)
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
' Initialize all project tables
' JDL 12/2/24
Sub test_InitTbls(procs)
    Dim tst As New Test: tst.Init tst, "test_InitTbls"
    Dim tbls As Object
    
    tst.Assert tst, DemoProj.InitTbls(tbls)
    
    ' Check that tblRowsCols objects are provisioned correctly
    With tbls
    
        'Extra checks of setting row/column ranges
        CheckProvision tst, .Inputs, "Home", True, 6, 8
        tst.Assert tst, .Inputs.rngRows.Address = "$6:$9"
        tst.Assert tst, .Inputs.rngTable.Address = "$H$5:$J$9"
        
        CheckProvision tst, .RowsCols, "RowsCols", False, 2, 1
    End With
    
    ' Update and report test results
    tst.Update tst, procs
End Sub
'-----------------------------------------------------------------------------------------
' CheckProvision helper method for test_InitTbls
' JDL 10/7/24
Sub CheckProvision(tst As Test, tbl As Object, sht As String, IsCustomTbl As Boolean, _
        rHome As Long, cHome As Long)
    tst.Assert tst, tbl.sht = sht
    tst.Assert tst, tbl.IsCustomTbl = IsCustomTbl
    tst.Assert tst, tbl.cellHome.Address = tbl.wksht.Cells(rHome, cHome).Address
End Sub
'-----------------------------------------------------------------------------------------
' Initialize project Scenario Model(s)
' JDL 12/2/24
'
Sub test_InitMdls(procs)
    Dim tst As New Test: tst.Init tst, "test_InitMdls"
    Dim mdls As Object
    
    With tst
        .Assert tst, DemoProj.InitMdls(mdls)
    
        ' Check that Home model is provisioned correctly
        .Assert tst, mdls.Home.sht = "Home"
        .Assert tst, mdls.Home.colrngVarNames.Address = "$C:$C"
        .Assert tst, mdls.Home.rngRows.Address = "$4:$6"
    
        'Check Params (Params_ sheet) model
        .Assert tst, mdls.params.sht = "params_"
        .Assert tst, mdls.params.ScenModelLoc(mdls.params, "mm.m") = 1000#
        .Assert tst, mdls.params.rngRows.Address = "$2:$8"
        
        .Update tst, procs
    End With
End Sub
'-----------------------------------------------------------------------------------------
' Clear Study Tables
' JDL 11/7/24
Sub test_ClearTables(procs)
    Dim tst As New Test: tst.Init tst, "test_ClearTables"
    Dim tbls As Object, study As Object
    
    Set study = DemoProj.New_DemoStudy()
    tst.Assert tst, study.ClearTables(tbls)
    
    tst.Update tst, procs
End Sub
'-----------------------------------------------------------------------------------------------
'Populate .Inputs table (Assumes .InitAllTables has been run to Provision tbls)
'JDL 12/3/24
Sub test_PopulateInputsTbl(procs)
    Dim tst As New Test: tst.Init tst, "test_PopulateInputsTbl"
    Dim tbls As Object
    
    tst.Assert tst, DemoProj.InitTbls(tbls)
    PopulateInputsTbl tbls
    
    ' Check the first and last row values
    With tbls.Inputs
        tst.Assert tst, .wksht.Cells(.cellHome.Row, .cellHome.Column).Value = "A"
        tst.Assert tst, .wksht.Cells(.cellHome.Row, .cellHome.Column + 1).Value = "Something 1"
        tst.Assert tst, .wksht.Cells(.cellHome.Row, .cellHome.Column + 2).Value = 10
        
        tst.Assert tst, .wksht.Cells(.cellHome.Row + .nrows - 1, .cellHome.Column).Value = "D"
        tst.Assert tst, .wksht.Cells(.cellHome.Row + .nrows - 1, .cellHome.Column + 1).Value = "Something 4"
        tst.Assert tst, .wksht.Cells(.cellHome.Row + .nrows - 1, .cellHome.Column + 2).Value = 40
    End With
    
    tst.Update tst, procs
End Sub
'-----------------------------------------------------------------------------------------------
'Populate .RowsCols table (Assumes .InitAllTables has been run to Provision tbls)
'JDL 12/3/24
Sub test_PopulateRowsColsTbl(procs)
    Dim tst As New Test: tst.Init tst, "test_PopulateRowsColsTbl"
    Dim tbls As Object
    
    tst.Assert tst, DemoProj.InitTbls(tbls)
    PopulateRowsColsTbl tbls
    
    ' Check the first and last row values
    With tbls.RowsCols
        tst.Assert tst, .wksht.Cells(.cellHome.Row, .cellHome.Column).Value = 1
        tst.Assert tst, .wksht.Cells(.cellHome.Row, .cellHome.Column + 1).Value = "A"
        
        tst.Assert tst, .wksht.Cells(.cellHome.Row + 7, .cellHome.Column).Value = 8
        tst.Assert tst, .wksht.Cells(.cellHome.Row + 7, .cellHome.Column + 1).Value = "D"
    End With
    
    tst.Update tst, procs
End Sub
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
' Tests of procs.DemoStudy
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
' RefreshRowsColsTblProcedure - Example of instancing class and testing class method
' JDL 12/3/24; Updated 4/28/26
Sub test_RefreshRowsColsTblProcedure(procs)
    Dim tst As New Test: tst.Init tst, "test_RefreshRowsColsTblProcedure"
    Dim tbls As Object, study As Object
    
    'Instance DemoStudy class
    Set study = DemoProj.New_DemoStudy
    
    With tst
        'Initialize tbls and populate .RowsCols with mockup data
        .Assert tst, DemoProj.InitTbls(tbls)
        PopulateRowsColsTbl tbls
    
        'Method being tested
        .Assert tst, study.RefreshRowsColsTblProcedure(study, tbls)
        
        .Assert tst, .wkbkTest.Sheets(shtRowsCols).Cells(2, 3).Value2 = "Something 1"
    
        .Update tst, procs
    End With
End Sub


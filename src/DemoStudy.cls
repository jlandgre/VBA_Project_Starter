VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DemoStudy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Version 12/3/24 - Demo of mockup study
Option Explicit
'-----------------------------------------------------------------------------------------
' Demo procedure to refresh RowsCols table
'-----------------------------------------------------------------------------------------
' RefreshRowsColsTblProcedure - demo
' JDL 12/3/24
'
Public Function RefreshRowsColsTblProcedure(study, tbls)
    SetErrs RefreshRowsColsTblProcedure: If errs.IsHandle Then On Error GoTo ErrorExit
    
    'Refresh .RowsCols table
    If Not RefreshTblAPI(tbls.RowsCols.wkbk, IsReplace:=True, IsTblFormat:=True, _
                            sht:="RowsCols", IsNamePrefix:=True) Then GoTo ErrorExit

    'Add other code to do use case tasks
    
    Exit Function
    
ErrorExit:
    errs.RecordErr "RefreshRowsColsTblProcedure", RefreshRowsColsTblProcedure
End Function
'-----------------------------------------------------------------------------------------
' Utilities
'-----------------------------------------------------------------------------------------
' ClearTables (called by interface.ClearAllInputs driver sub)
' JDL 12/3/24
'
Public Function ClearTables(tbls) As Boolean
    SetErrs ClearTables: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim study As New DemoStudy

    ' Provision all tables
    If Not InitTbls(tbls) Then GoTo ErrorExit
    
    With tbls
        'Custom tables
        Intersect(.Inputs.rngRows, .Inputs.rngTable).ClearContents
        
        'Default, rowscols tables
        If Not study.ClearTbl(.RowsCols) Then GoTo ErrorExit
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "ClearTables", ClearTables
End Function
'-----------------------------------------------------------------------------------------
' Clear previous table data
' JDL 12/2/24
'
Public Function ClearTbl(tbl) As Boolean
    SetErrs ClearTbl: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim iLastRow As Integer
    
    ' Clear the previous table data if any
    With tbl
        iLastRow = .wksht.UsedRange.Rows(.wksht.UsedRange.Rows.Count).Row
        If iLastRow > 1 Then .wksht.Range(.wksht.Rows(2), .wksht.Rows(iLastRow)).Clear
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "ClearTbl", ClearTbl
End Function

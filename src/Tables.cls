VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Version 4/28/26
Option Explicit
Public Inputs As Object
Public RowsCols As Object
' Initialize project tables
' JDL 4/28/26
'-----------------------------------------------------------------------------------------
Function Init(tbls) As Boolean
    SetErrs Init: If errs.IsHandle Then On Error GoTo ErrorExit
    With tbls
        Set .Inputs = ExcelSteps.New_tbl
        Set .RowsCols = ExcelSteps.New_tbl
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "Tables.Init", Init
End Function

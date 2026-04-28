VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Models"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Version 4/28/26
Option Explicit
Public Home As Object
Public Params As Object
' Initialize project Scenario Models
' JDL 4/28/26
'-----------------------------------------------------------------------------------------
Function Init(mdls) As Boolean
    SetErrs Init: If errs.IsHandle Then On Error GoTo ErrorExit
    With mdls
        Set .Home = ExcelSteps.New_mdl
        Set .Params = ExcelSteps.New_mdl
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "Models.Init", Init
End Function



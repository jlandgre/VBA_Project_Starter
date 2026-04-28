Attribute VB_Name = "interface"
'Version 4/28/26
Option Explicit
'-----------------------------------------------------------------------------------------
' Initialize all or some project tables
' JDL 4/28/26
Public Function InitTbls(tbls, Optional IsAll As Boolean = True, _
    Optional ByVal IsInputs As Boolean = False, _
    Optional ByVal IsRowsCols As Boolean = False, _
    Optional IsRefresh = False) As Boolean
    
    SetErrs InitTbls, ThisWorkbook: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim wkbk As Workbook: Set wkbk = ThisWorkbook
    Dim defn As String, sht As String
    
    ' Initialize tbls if not already instanced (allows multiple calls)
    If tbls Is Nothing Then
        Set tbls = New Tables: If Not tbls.Init(tbls) Then GoTo ErrorExit
    End If
    
    If IsAll Then
        IsInputs = True
        IsRowsCols = True
    End If

    With tbls
        'Non-default tbls "sht:rHome,cHome:....:nrows:ncols" see .SetCustomTblParams(tbl)
        If IsInputs Then
            defn = "Home:6,8:T:T:T:F:F:T:0:-1:4:3"
            If Not .Inputs.Provision(.Inputs, wkbk, False, TblName:="Inputs", defn:=defn) _
                Then GoTo ErrorExit
                
            'No refresh for custom, non-homed table
        End If
            
        'Default table (homed, single object on sheet)
        If IsRowsCols Then
            sht = "RowsCols"
            If Not .RowsCols.Provision(.RowsCols, wkbk, False, sht:=sht, _
                IsSetColRngs:=True, IsSetColNames:=True) Then GoTo ErrorExit
              
            'Optional refresh default, homed table
            If IsRefresh Then
                If Not RefreshTblAPI(wkbk, IsReplace:=True, IsTblFormat:=True, sht:=sht) _
                    Then GoTo ErrorExit
            End If
        End If
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "InitTbls", InitTbls
End Function
'-----------------------------------------------------------------------------------------
' Initialize project Scenario Model(s)
'
' Example Definition w/o non-default sName: Process:8,31:0:T:T:T:T:T
'                        non-default sName: Process:8,31:20:T:T:T:T:T:mdlProcess
'
' "Process" = sheet name
' True/False params: IsCalc, IsSuppHeader, IsRngNames, IsMdlNmPrefix, IsLiteModel
' 8,31 = cell home row and column
' 0 / 20 = specified, fixed number of rows (0 = no limit)
' "mdlProcess" = model name (to override default model name = sheet name)
'
' JDL 12/2/24; Updated 4/28/26
Public Function InitMdls(mdls, Optional IsAll As Boolean = True, _
    Optional ByVal IsHome As Boolean = False, _
    Optional ByVal IsParams As Boolean = False, _
    Optional ByVal IsRefresh As Boolean = False) As Boolean
    
    SetErrs InitMdls, ThisWorkbook: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim wkbk As Workbook: Set wkbk = ThisWorkbook
    Dim defn As String, sht As String

    ' Initialize tbls if not already instanced (allows multiple calls)
    If mdls Is Nothing Then
        Set mdls = New Models: If Not mdls.Init(mdls) Then GoTo ErrorExit
    End If
    
    If IsAll Then
        IsHome = True
        IsParams = True
    End If

    'Set mdls = New Models: If Not mdls.Init(mdls) Then GoTo ErrorExit

    ' Initialize all needed mdlScenario objects in mdls
    With mdls
        
        'Non-default model "sht:rHome,cHome:....:nrows:ncols"
        If IsHome Then
            defn = "Home:4,1:0:T:T:T:T:T"
            If Not .Home.Provision(.Home, wkbk, defn:=defn, IsMdlNmPrefix:=True) Then GoTo ErrorExit
        End If
            
        'Default model on its own sheet
        If IsParams Then
            sht = "params_"
            If Not .Params.Provision(.Params, wkbk, sht:=sht, _
                IsCalc:=True, IsMdlNmPrefix:=False) Then GoTo ErrorExit
            If Not .Params.Refresh(.Params) Then GoTo ErrorExit
        End If
    End With
    Exit Function
    
ErrorExit:
    errs.RecordErr "InitMdls", InitMdls
End Function
'-----------------------------------------------------------------------------------------
' Clear input tables
' JDL 12/3/24
Sub ClearInputTables()
    SetErrs "driver", ThisWorkbook: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim tbls As Object, study As New DemoStudy
    
    ' Display Yes/No message box to warn the user
    If Not IsTest Then
        If errs.ShowMessage("ClearAllInputs", 1, vbYesNo + vbExclamation) = vbNo Then Exit Sub
    End If
    
    SetApplEnvir False, False, xlCalculationAutomatic
    
    If Not study.ClearTables(tbls) Then GoTo ErrorExit
    
    SetApplEnvir True, True, xlCalculationAutomatic
    Exit Sub
    
ErrorExit:
    errs.RecordErr "ClearAllInputs"
End Sub

'-----------------------------------------------------------------------------------------------------
' Hide admin and validation sheets
' JDL 12/2/24
Sub ManageSheetVisibility(wkbk, Optional IsHide = True)
    Dim sht As Variant
    
    For Each sht In Array("Params_", "Errors_")
        wkbk.Sheets(sht).Visible = Not IsHide
    Next sht
End Sub


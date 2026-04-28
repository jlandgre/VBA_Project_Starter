Attribute VB_Name = "Validation"
Option Explicit

'Global variable (default False for production) can toggle to True from tests workbook
Public IsTest As Boolean

'-----------------------------------------------------------------------------------------------------
'This module contains functions for instancing project workbook objects from a
'second workbook. To call, the second workbook's VBA Project
'needs to add a Reference to the project workbook's VBA project (Tools / References menu
'in VBA editor). The second workbook should instance by
'calling these modValidation functions as shown where "VBAProject_XYZ is the name of
'the referenced workbook's VBA Project:
'
'Dim tbl as object
'Set tbl = VBAProject_XYZ.new_tblRowsCols
'
'   <<or alternatively>>
'
'Set tbl = Application.Run(sDirPrefix_XYZ & "New_tbl")
'
'   where sDirPrefix_XYZ = "c:\dir1\dir2!" -- path to XLSteps.xlam
'
'JDL 12/15/22; updated 9/30/24
Public Function New_DemoStudy() As DemoStudy
    Set New_DemoStudy = New DemoStudy
End Function

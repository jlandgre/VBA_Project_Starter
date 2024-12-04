### VBA_Project_Starter

This repository is a demo and project starter. It is designed to interoperate with use of Github Copilot for coding. The VBA toolbox aka "modeling operating system" contains the following:

* The project workbook, **DemoStudy.xlsm**, demonstrates a user interface with structured data objects useful for a wide range of spreadsheet modeling projects. These objects are rows/columns tables and "scenario models" aka a columns-by-rows format --demonstrated here for handling single-valued inputs. These standard objects can provide both user-facing (Excel range names) and programmatic, object attribute wayfinding to access and work with their data. The "model" in this example is just a mockup, but it is represented by the **DemoStudy** class in the project workbook's VBA project.
<p align="center">
  <u>Project Workbook Home Sheet</u></br><small>Scenario Model B3:F6</br>Custom Input Table H5:J9</small></br>
  <img src=images/project1.png alt="Overall Inputs" width=500></br>
</p>

* A test suite workbook, **tests_DemoStudy.xlsm**, with tests grouped by "Procedures" to make it convenient to run all or subsets of unit tests. This workbook contains a custom **Test** class to handle pass/fail assertions testing project workbook results. It reports test results on the **DemoStudyTests** sheet in the test workbook.
* The test workbook's VBA project contains a reference to the project workbook's VBA project --allowing tests to execute project workbook code and check the results.
<p align="center">
  <u>Test Suite Results</u></br><small>Best practice unit test naming: "test_&lt;class method name>"</small></br>
  <img src=images/tests1.png alt="Overall Inputs" width=300></br>
</p>

* The project workbook contains **Tables** and **Models** class instances, **tbls** and **mdls**, to manage and provide wayfinding for the data objects within a project. Our complex consulting projects have involved 8+ such objects to handle multiple use cases of a simulator or other application. The individual objects have hard-coded names (aka "Inputs" and "RowsCols" here) making programmatic references simple and transparent. An example is **tbls.Inputs.rngRows** for referring to the range of data rows in the **Inputs** table that is on the **Home** sheet.
<div style="font-size: 9px;">

```vba
'-----------------------------------------------------------------------------------------
' Initialize all project tables
' JDL 12/2/24 (based on client project 9/24)
Public Function InitAllTables(tbls) As Boolean
    SetErrs InitAllTables: If errs.IsHandle Then On Error GoTo ErrorExit
    Dim wkbk As Workbook, defn As String
    Set wkbk = ThisWorkbook

    Set tbls = New Tables: If Not tbls.Init(tbls) Then GoTo ErrorExit

    With tbls

        'Non-default tbls "sht:rHome,cHome:....:nrows:ncols" see .SetCustomTblParams(tbl)
        defn = "Home:6,8:T:T:T:F:F:T:0:-1:4:3"
        If Not .Inputs.Provision(.Inputs, wkbk, False, TblName:="Inputs", defn:=defn) _
            Then GoTo ErrorExit
    
        'Default tables (homed, single object on sheet)
        If Not .RowsCols.Provision(.RowsCols, wkbk, False, sht:="RowsCols", _
            IsSetColRngs:=True, IsSetColNames:=True) Then GoTo ErrorExit
    End With
    Exit Function

ErrorExit:
    errs.RecordErr "InitAllTables", InitAllTables
End Function
```
</div>

* The project workbook utilizes a custom [ErrorHandling class](https://github.com/jlandgre/VBA_ErrorHandling) for error trapping and to enable giving the user advisory warnings via dialog boxes or Excel cell comments. 
* the project workbook utilizes [ExcelSteps API refresh](https://github.com/jlandgre/ExcelSteps) of structured data elements as detailed below

<u>Interplay of code objects and ExcelSteps Recipe Refresh API</u></br>
In a programmatic model such as the example, rows/columns table or Scenario Model objects can be "default" or "custom" format that allows non-default location, dimensions and other attributes. In the example, **RowsColumns** and **params_** are examples of a default table and Scenario Model, respectively. The table is homed to cell A1 and does not have fixed dimensions. Once the sheet name is specified, the program can sense the table's dimensions and work with its data. Similarly, the **params_** sheet's Scenario Model utilizes a standard template that can be refreshed based only on knowing its sheet name.

Examples of custom table and Scenario Model objects are the **Inputs** table on the Home sheet and the Scenario Model in B3:F6 on Home. Both have custom home cell locations and their dimensions are fixed to allow additional objects to be placed around them on the same sheet as needed per application.

The ExcelSteps API's recipe refresh capability is a powerful tool for storing curated and validated spreadsheet model formulas where they can be hidden from the user to prevent corruption. This eliminates the need to hard code much of the model including formatting the input and output tables and models. 

#### Calculation Model Recipe Example
In a spreadsheet model with our toolbox, inserted columns' formulas can be left in a live calculation state, or the recipe can be flagged to paste over with just values. For example, the picture shows the validated formulas for color calculations including the [CMC Delta E color difference (blog gives background)](https://datadelveengineer.com/getting-color-right-in-products/) from a cleaning study template that client lab personnel use to run and analyze their studies.
<p align="center">
  <u>Color Model Recipe for Model with Curated, Inserted Columns</u></br><small>Column F "FALSE" causes model to not retain live, Excel formulas post-calculation</small></br>
  <img src=images/color_recipe1.png alt="Overall Inputs" width=500></br>
</p>

In the case of our **DemoStudy**, to run a demo of the use case of adding a recipe-sourced calculated column to **RowsCols** sheet:
1. Open the project and test workbooks. 
2. In the test workbook's **TestingDriver_DemoStudy** subroutine, set procs.DemoStudy.Enabled = True
3. Run the subroutine to execute the **test_RefreshRowsColsTblProcedure**. This calls that DemoStudy class method in the project workbook. It uses the ExcelSteps recipe to refresh the **RowsCols** sheet --adding a calculated column and some appropriate formatting in the process.

<p align="center">
  <u>Recipe for DemoStudy Example and Resulting, Refreshed RowsCols Table</u></br>
  <img src=images/project3.png alt="Overall Inputs" width=500></br></br>
  <img src=images/project2.png alt="Overall Inputs" width=300></br>
</p>

#### Conclusion
When used with AI tool prompts for code writing, the DemoStudy template is a starting point for custom VBA projects where there is an emphasis on having a good user interface for helping users walk through use cases such as designing studies followed by lab data entry and analysis involving complex calculations. An alternate application is populating output templates from Python models to put output data in the hands of users. In such applications, the Python code can write data to the *.xlsm project template and can even populate the ExcelSteps recipe before calling the API to refresh the output to format it nicely for the end user.

J.D. Landgrebe / Data Delve LLC
December 2024
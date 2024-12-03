### VBA_Multitable

This repository is a demo and project starter based on a VBA toolbox comprised of:

* A Project workbook **DemoStudy.xlsm** demonstrates a user interface containing structured data elements useful for spreadsheet modeling projects. These are rows/columns tables and "scenario models" aka columns by rows templates for handling single-valued inputs. These standard objects provide both user-facing (range names) and programmatic wayfinding to access and work with their data. The "model" in this example is the **DemoStudy** class in the VBA project.
* A test suite workbook **tests_DemoStudy.xlsm** with tests grouped by "Procedures" to make it convenient to run all or subsets of unit tests. This workbook contains a custom **Test** class to handle pass/fail assertions testing project workbook results. It reports test results on the **DemoStudyTests** sheet in the test workbook.
* The test workbook's VBA project contains a reference to the project workbook's VBA project --allowing tests to execute project workbook code and check the results.
* The project workbook contains **Tables** and **Models** class instances, **tbls** and **mdls**, to manage and provide wayfinding for the data objects within a project. The individual objects have hard-coded names making programmatic references easy and transparent. An example is **tbls.Inputs.rngRows** for the range of data rows in the **Inputs** table that is on the **Home** sheet.
* The project workbook contains a custom [ErrorHandling class](https://github.com/jlandgre/VBA_ErrorHandling) for error trapping and to enable advisory warnings via dialog boxes or cell comments. 
* [ExcelSteps refresh](https://github.com/jlandgre/ExcelSteps) of structured data elements. This is demonstrated by a recipe on the ExcelSteps sheet that can insert or refresh formula-containing columns. The recipe also includes example formatting instructions.

The DemoStudy workbook's VBA project contains classes for managing rows/columns tables (tblRowsCols) and Scenario Models (mdlScenario). In a programmatic model such as the example, the table or Scenario Model objects can be "default" format as with the table on the **RowsColumns** sheet and the Scenario Model on the **params_** sheet. The table is default because it is homed to cell A1 and does not have fixed dimensions. Once the sheet name is specified, the program can sense the table's dimensions and work with its data. Similarly, the **params_** sheet's Scenario Model utilizes a standard template that can be refreshed based only on knowing its sheet name.

Alternatively, table and Scenario Model objects can be custom as in the example **Inputs** table on the Home sheet. It has a custom home cell location on the sheet, and its dimensions are fixed. The definition of this table is defined programmatically in the VBA code. Similarly, the Home sheet contains a custom Scenario Model useful for user-facing input values.

The DemoStudy template can be picked up as a starting point for a custom VBA project to gather user inputs or import data and then perform data transformations and calculations.

To run a demo of the use case of adding a recipe-sourced calculated column to RowsCols sheet:
1. Open the project and test workbooks. 
2. In the test workbook's **TestingDriver_DemoStudy** subroutine, set procs.DemoStudy.Enabled = True
3. Run the subroutine to execute the **test_RefreshRowsColsTblProcedure**. This calls that DemoStudy class method in the project workbook. It uses the ExcelSteps recipe to refresh the **RowsCols** sheet --adding a calculated column and some appropriate formatting in the process.


J.D. Landgrebe / Data Delve LLC
December 2024
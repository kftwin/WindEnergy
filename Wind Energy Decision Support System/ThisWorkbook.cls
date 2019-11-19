VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub Workbook_Open()
'''''''' Hide the data worksheets and leave the welcome sheet visible
Worksheets("Welcome").Visible = True
Worksheets("Average Daily Wind").Visible = False
Worksheets("Wind Turbine Specs").Visible = False
Worksheets("Result Worksheet").Visible = False
Worksheets("charts").Visible = False
''''''''
End Sub

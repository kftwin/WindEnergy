VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmmonthly 
   Caption         =   "Monthly Output"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14415
   OleObjectBlob   =   "frmmonthly.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmmonthly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdcalculate_Click()
If tbxperiods.Value = "" Then
    MsgBox ("Please input the number of periods for the analysis.")
    Exit Sub
End If
If tbxinterestrate.Value = "" Then
    MsgBox ("Please input the interest rate for the analysis.")
    Exit Sub
End If

Dim npv As Double, paybackperiod As Double, initalinvestment As Double, annualprofit As Double
Dim interestrate As Double, periods As Double, n As Double
initialinvestment = CDbl(lblinitialinvestment.Caption)
annualprofit = CDbl(lblannualprofit.Caption)
interestrate = CDbl(tbxinterestrate.Value) / 100
periods = CDbl(tbxperiods.Value)

npv = (annualprofit) * ((1 + interestrate) ^ periods - 1) / (interestrate * (interestrate + 1) ^ periods) - initialinvestment
lblnpv.Caption = FormatCurrency(npv, 2)
n = 1
Do Until (annualprofit) * ((1 + interestrate) ^ n - 1) / (interestrate * (interestrate + 1) ^ n) - initialinvestment >= 0
n = n + 1
Loop

lblppb.Caption = n & " Years"
End Sub

Private Sub cmdgoback_Click()
Unload Me
'Frmresult.Show
End Sub

Private Sub UserForm_Activate()
Dim cityName As String
'cityName = Frmresult.pin.ControlTipText
'cityName = "Gainesville"
Call add_chart(cityMap)
Dim ecostart As Range, ecocount As Double, i As Double
Set ecostart = Worksheets("Result Worksheet").Range("A6")
If ecostart.Offset(1, 0) <> "" Then
    ecocount = Range(ecostart.Offset(1, 0), ecostart.End(xlDown)).Rows.count
End If
lblinitialinvestment.Caption = FormatCurrency(Worksheets("Result Worksheet").Range("B4") * 1000000, 2)
i = 1
For i = 1 To ecocount
    If ecostart.Offset(i, 0) = cityMap Then
        lblannualprofit.Caption = FormatCurrency(ecostart.Offset(i, 4), 2)
    End If
Next i
End Sub

Sub add_chart(cityName As String)
'''''''''''''' To speed up
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
'''''''''''''''''''''
Worksheets("charts").Visible = True

Dim cityPos As Integer
Dim datastart As Range, cityRange As Range



Set datastart = Worksheets("Average Daily Wind").Range("A:A")
cityPos = Application.WorksheetFunction.Match(cityName, datastart, 0)


 Dim wsItem As Worksheet
    Dim chtObj As ChartObject
     
    For Each wsItem In ThisWorkbook.Worksheets
         
        For Each chtObj In wsItem.ChartObjects
             
            chtObj.Delete
             
        Next
         
    Next


If num_charts > 0 Then
ThisWorkbook.Charts.Delete
End If

Charts.Add

ActiveChart.ChartArea.Select
ActiveChart.name = "my_chart"
ActiveChart.Location Where:=xlLocationAsObject, name:="charts"
ActiveChart.ChartType = xlLine 'Type of graph
ActiveChart.SetSourceData Source:=Sheets("Average Daily Wind").Range(Worksheets("Average Daily Wind").Cells(cityPos, 5), Worksheets("Average Daily Wind").Cells(cityPos, 16)), PlotBy _
:=xlRows 'data source
ActiveChart.SeriesCollection(1).XValues = Sheets("Average Daily Wind").Range("E2:P2") 'naming the x-axis
ActiveChart.SeriesCollection(1).name = cityName & " Monthly Wind Efficiency " ' Heading of the graph
With ActiveChart.Axes(xlValue)
.HasMajorGridlines = False
.HasMinorGridlines = False
.HasTitle = True
 With .AxisTitle
 .Characters.Text = "% Wind Efficiency"
 End With
End With
With ActiveChart.Axes(xlCategory)
.HasTitle = True
 With .AxisTitle
 .Characters.Text = "Months"
 End With
End With
 ' Background of graph
With ActiveChart.PlotArea.Border
.ColorIndex = 16
.Weight = xlThin
.LineStyle = xlNone
End With
ActiveChart.PlotArea.Interior.ColorIndex = xlNone

Dim currentchart As Chart

Set currentchart = Sheets("charts").ChartObjects(1).Chart
    Fname = ThisWorkbook.Path & "\temp.gif"
    currentchart.Export Filename:=Fname, FilterName:="GIF"
'''''''''''speed up
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
'''''''''''''''''
graph.Picture = LoadPicture(Fname)
graph.AutoSize = True

End Sub



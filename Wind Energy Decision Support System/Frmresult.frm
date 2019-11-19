VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frmresult 
   Caption         =   "Result"
   ClientHeight    =   10875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   19350
   OleObjectBlob   =   "Frmresult.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Frmresult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim startX, startLat As Integer
Dim startY, startLong As Integer
Dim endX As Double
Dim endY As Double
Dim Labels() As New Class1
Sub GroupLabels()
    Dim LabelCount As Integer
    Dim ctl As Control

    LabelCount = 0
    For Each ctl In Frmresult.Controls
        If TypeName(ctl) = "Label" Then
            If CStr(Left(ctl.name, 4)) = "plot" Or CStr(Left(ctl.name, 4)) = "city" Then
                LabelCount = LabelCount + 1
                ReDim Preserve Labels(1 To LabelCount)
               Set Labels(LabelCount).LabelGroup = ctl
            End If
        End If
    Next ctl
End Sub
Public Sub plotLocation(locLat As Double, locLong As Double, ByRef latitude() As Double, ByRef longitude() As Double)
startX = 53.4197
startY = 90
startLat = 55
startLong = 125
Const longConvert = 0.112319
Const latConvert = 0.091
endX = startX + (startLong - Abs(locLong)) / longConvert
endY = startY + (startLat - Abs(locLat)) / latConvert
Dim stringlength As Integer
Dim cityCheck As String


pin.Left = endX
pin.Top = endY

pin.Visible = True
'=============================================Other Cities ==================================================

Dim i As Integer
i = 0

For i = 0 To UBound(latitude())
    If Left(selection(i), Len(lblcity.Caption)) <> lblcity.Caption Then
    endX = startX + (startLong - Abs(longitude(i))) / longConvert
    endY = startY + (startLat - Abs(latitude(i))) / latConvert
    
    
    Dim NewLabel As MSForms.label
    
    
    Set NewLabel = Frmresult.Controls.Add("Forms.label.1")
            With NewLabel
                .name = "plot" & i + 1
                .Caption = "O"
                .Top = endY + 11
                .Left = endX + 7
                .ForeColor = vbBlue
                .AutoSize = True
                .BackColor = vbWhite
                .BorderColor = vbBlack
                .Font.Bold = True
                .ControlTipText = selection(i)
                
            End With
    Else
    pin.ControlTipText = selection(i)
            
            End If
Next
'=======================================cities around selected radius
Dim onOff As Variant
Dim j As Integer
j = 0
Worksheets("Result Worksheet").Activate
Worksheets("Result Worksheet").Visible = True
Dim withinradiusstart As Range
Set withinradiusstart = Worksheets("Result Worksheet").Range("F29")

Dim k As Integer
k = 0


For j = 0 To Range("allonoroff").Rows.count - 1
If withinradiusstart.Offset(j + 1, 1) = 1 Then

ReDim Preserve arraywithinCity(k)
ReDim Preserve arraywithinState(k)
ReDim Preserve arraywithinDemand(k)
ReDim Preserve arraywithinPrice(k)
ReDim Preserve arraywithinDistance(k)
ReDim Preserve arraywithinLatitude(k)
ReDim Preserve arraywithinLongitude(k)

arraywithinCity(k) = withinradiusstart.Offset(j + 1, 0)
arraywithinState(k) = withinradiusstart.Offset(j + 1, -1)
arraywithinDemand(k) = withinradiusstart.Offset(j + 1, 2)
arraywithinPrice(k) = withinradiusstart.Offset(j + 1, 3)
arraywithinDistance(k) = withinradiusstart.Offset(j + 1, 5)
arraywithinLatitude(k) = withinradiusstart.Offset(j + 1, 6)
arraywithinLongitude(k) = withinradiusstart.Offset(j + 1, 7)




    endX = startX + (startLong - Abs(arraywithinLongitude(k))) / longConvert
    endY = startY + (startLat - Abs(arraywithinLatitude(k))) / latConvert

Dim l As Integer
For l = 0 To UBound(selection())
    stringlength = Len(selection(l)) - 4
    cityCheck = Mid(selection(l), 1, stringlength)

If arraywithinCity(k) <> lblcity.Caption And arraywithinCity(k) <> cityCheck Then

    Set NewLabel = Frmresult.Controls.Add("Forms.label.1")
            With NewLabel
                .name = "city" & k + 1
                .Caption = "X"
                .Top = endY + 11
                .Left = endX + 7
                .ForeColor = vbRed
                .AutoSize = True
                .BackColor = vbWhite
                .BorderColor = vbBlack
                .Font.Bold = True
                .ControlTipText = arraywithinCity(k) & ", " & arraywithinState(k)

            End With
        Else
        Exit For
        End If
            Next
End If
k = k + 1
Next
Call GroupLabels
End Sub
Private Sub pin_Click()
cityMap = lblcity.Caption
frmmonthly.Show
End Sub

Private Sub UserForm_Activate()
Call showresults
Call populatelistbox


End Sub
Public Sub showresults()
Dim checkpointresult As Range, checkpointcount As Double, i As Double
i = 1
Set checkpointresult = Worksheets("Result Worksheet").Range("A6")
If checkpointresult.Offset(1, 0) <> "" Then
    checkpointcount = Range(checkpointresult.Offset(1, 0), checkpointresult.End(xlDown)).Rows.count
End If
For i = 1 To checkpointcount
    If checkpointresult.Offset(i, 3) = Worksheets("Result Worksheet").Range("H3") Then
        lblcity.Caption = checkpointresult.Offset(i, 0)
        lbllatitude = FormatNumber(checkpointresult.Offset(i, 2), 4)
        lbllongitude = FormatNumber(checkpointresult.Offset(i, 1), 4)
        Call plotLocation(checkpointresult.Offset(i, 2), checkpointresult.Offset(i, 1), selectedlatitude(), selectedlongitude())
    End If
Next i

lbloutputcapacity = FormatNumber(Worksheets("Result Worksheet").Range("H17"), 4) & " MWh"
lblprofit = FormatCurrency(Worksheets("Result Worksheet").Range("L23"))
lblsupply = FormatNumber(Worksheets("Result Worksheet").Range("H26"), 4) & " MWh"


End Sub
Private Sub btnGraph_Click()
cityMap = lblcity.Caption
frmmonthly.Show
End Sub
Private Sub cmdback_Click()
Unload Me
frmMain.Show
End Sub
Public Sub populatelistbox()
Dim turbinestart As Range, i As Integer, n As Integer
Set turbinestart = Worksheets("Result Worksheet").Range("G8")
    lbxturbines.Clear
    lbxturbines.ColumnCount = 3
    
    lbxturbines.AddItem ("Brand")
    lbxturbines.list(0, 1) = "Rate (MW)"
    lbxturbines.list(0, 2) = "Quantity"
n = 1
For i = 0 To 4
    If turbinestart.Offset(i, 3) <> 0 Then
        lbxturbines.AddItem (turbinestart.Offset(i, 0))
        lbxturbines.list(n, 1) = turbinestart.Offset(i, 1)
        lbxturbines.list(n, 2) = turbinestart.Offset(i, 3)
        n = n + 1
    End If
Next i
End Sub



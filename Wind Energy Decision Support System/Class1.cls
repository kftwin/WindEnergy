VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents LabelGroup As MSForms.label
Attribute LabelGroup.VB_VarHelpID = -1

Private Sub LabelGroup_Click()
Dim stringlength As Integer
Dim thecity As String

    For i = 0 To UBound(selectedlatitude())
    If LabelGroup.name = "plot" & i + 1 Then
    stringlength = Len(LabelGroup.ControlTipText) - 4
    cityMap = Mid(LabelGroup.ControlTipText, 1, stringlength)
    frmmonthly.Show
    End If
    Next
    
    
     For j = 0 To UBound(arraywithinCity())
    If LabelGroup.name = "city" & j + 1 Then
    MsgBox ("City, State: " & arraywithinCity(j) & ", " & arraywithinState(j) & vbNewLine & "Revenue/Year: " & FormatCurrency(arraywithinPrice(j), 2) & vbNewLine & "Demand/Year: " & arraywithinDemand(j) & "MWh" & vbNewLine & "Distance From Facility: " & Round(arraywithinDistance(j), 2) & " miles"), , arraywithinCity(j) & ", " & arraywithinState(j)
    'stringlength = Len(LabelGroup.ControlTipText) - 4
    'cityMap = Mid(LabelGroup.ControlTipText, 1, stringlength)
    
    End If
    
    Next
End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "Wind Power Optimal Formulation"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11265
   OleObjectBlob   =   "frmMain.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public selectioncount As Double

Private Sub cmdAddAll_Click() '**
Call cmdselectall_Click
Call cmdadd_Click
End Sub

Private Sub cmdRemoveAll_Click() '**
 Call cmdselectall2_Click
 Call cmdremove_Click
End Sub
Private Sub tbxinvestment_Change()
If Not IsNumeric(tbxinvestment.Text) Then tbxinvestment.Text = ""
End Sub
Private Sub tbxradius_Change()
If Not IsNumeric(tbxradius.Text) Then tbxradius.Text = ""
End Sub
' Public rowcount As Double
Private Sub UserForm_Activate()
selectioncount = 0
ReDim selection(0)
Call fillarrays
Call fillcombobox
End Sub
Public Sub filterstatecombobox(ByVal regionname As String) '**
Dim combocount As Double, i As Double, states As String, p As Double
cmbstate.Clear
i = 0
p = 0
cmbstate.AddItem ("Regional States")
For i = 0 To rowcount - 1
    If region(i) = regionname Then
        states = state(i)
            If stateexist(existstate(), states) = False Then
            cmbstate.AddItem (state(i))
            p = p + 1
            ReDim Preserve existstate(p - 1)
            existstate(p - 1) = states
        End If
    End If
    i = i + 1
Next i
End Sub

Public Sub fillarrays() '**'''''fills the City(), State(), Annavergae(),output(), Lat(), and long() arrays
Dim i As Double, n As Double
Dim datastart As Range
Set datastart = Worksheets("Average Daily Wind").Range("A2")
If datastart.Offset(1, 0) <> "" Then
    rowcount = Range(datastart.Offset(1, 0), datastart.End(xlDown)).Rows.count
End If

ReDim city(rowcount)
ReDim state(rowcount)
ReDim region(rowcount)
ReDim annaverage(rowcount)
ReDim output(rowcount)
ReDim lat(rowcount)
ReDim longit(rowcount)
ReDim price(rowcount)
ReDim avenconsumption(rowcount)
ReDim percentdemand(rowcount)
ReDim demand(rowcount)
'Public price() As Double, avenconsumption() As Double, percentdemand() As Double, demand() As Double
n = 1
For i = 0 To rowcount - 1

    city(i) = datastart.Offset(n, 0)
    state(i) = datastart.Offset(n, 1)
    region(i) = datastart.Offset(n, 2)
    annaverage(i) = datastart.Offset(n, 16)
    output(i) = datastart.Offset(n, 17)
    lat(i) = datastart.Offset(n, 18)
    longit(i) = datastart.Offset(n, 19)
    price(i) = datastart.Offset(n, 20)
    avenconsumption(i) = datastart.Offset(n, 21)
    percentdemand(i) = datastart.Offset(n, 22)
    demand(i) = datastart.Offset(n, 23)
    n = n + 1
Next i
End Sub
Public Sub fillcombobox() '**
Dim regions As String, states As String, i As Double, n As Double, p As Double
cmbregion.AddItem ("All Regions")
ReDim existregion(0)
ReDim existstate(0)
n = 0
p = 0
For i = 0 To rowcount - 1 ' I originally had it without the substraction
    states = state(i)
    regions = region(i)
    
    If regionexist(existregion(), regions) = False Then
        cmbregion.AddItem (region(i))
        n = n + 1
        ReDim Preserve existregion(n - 1)
        existregion(n - 1) = regions
    End If
    If stateexist(existstate(), states) = False Then
        cmbstate.AddItem (state(i))
        p = p + 1
        ReDim existstate(p - 1)
        existstate(p - 1) = states
    End If
Next i
End Sub
Private Sub cmbregion_Change() '**
Dim regionselected As String, i As Double, thecity As String, thestate As String, states As String, p As Double
Dim regionname As Variant
cmbstate.Value = "Select a State"
lbxcityoptions.Clear
regionselected = cmbregion.Text
i = 0
p = 0
If regionselected = "All Regions" Then
    cmbstate.AddItem ("All States")
    ReDim existstate(0)
    For i = 0 To rowcount - 1
        states = state(i)
        If stateexist(existstate(), states) = False Then
            cmbstate.AddItem (state(i))
            p = p + 1
            ReDim Preserve existstate(p - 1)
            existstate(p - 1) = states
        End If
            thecity = city(i)
            thestate = state(i)
            Call AddWithNoDuplicates(lbxcityselected, thecity & ", " & thestate)
            'lbxcityoptions.AddItem (thecity & ", " & thestate)
    Next i
Else
    For Each regionname In region()
        If regionname = regionselected Then
            thecity = city(i)
            thestate = state(i)
            'lbxcityoptions.AddItem (thecity & ", " & thestate)
            Call AddWithNoDuplicates(lbxcityselected, thecity & ", " & thestate)
        End If
        i = i + 1
    Next regionname
    Call filterstatecombobox(regionselected)
End If
End Sub
Private Sub cmbstate_Change() '**
Dim stateselected As String, i As Double, thecity As String, thestate As String, selectedregion As String

stateselected = cmbstate.Value
selectedregion = cmbregion.Value
lbxcityoptions.Clear

If cmbstate.Value = "Regional States" Then
    For i = 0 To rowcount
        If region(i) = selectedregion Then
            thecity = city(i)
            thestate = state(i)
            'lbxcityoptions.AddItem (thecity & ", " & thestate)
            Call AddWithNoDuplicates(lbxcityselected, thecity & ", " & thestate)
        End If
    Next i
ElseIf cmbstate.Value = "All States" Then
    For i = 0 To rowcount - 1
            thecity = city(i)
            thestate = state(i)
            'lbxcityoptions.AddItem (thecity & ", " & thestate)
            Call AddWithNoDuplicates(lbxcityselected, thecity & ", " & thestate)
    Next i
Else
    For i = 0 To rowcount
        If state(i) = stateselected Then
            thecity = city(i)
            thestate = state(i)
            'lbxcityoptions.AddItem (thecity & ", " & thestate)
            Call AddWithNoDuplicates(lbxcityselected, thecity & ", " & thestate)
        End If
    Next i
End If
End Sub
Private Sub cmdadd_Click() '**
'Dim listcount As Double, i As Double, copycity As String
'listcount = lbxcityoptions.listcount

'ReDim Preserve selection(selectioncount)
'If lbxcityoptions.ListIndex <> -1 Then
'    For i = 0 To listcount - 1
'        If lbxcityoptions.Selected(i) = True Then
'            copycity = lbxcityoptions.Text
'            lbxcityselected.AddItem (copycity)
'            lbxcityoptions.RemoveItem (i)
'            selection(selectioncount) = copycity
'        End If
'    Next i
'End If
'selectioncount = selectioncount + 1
   Call AddDelValues(lbxcityoptions, lbxcityselected)
   Call SortListBox(lbxcityselected)
End Sub
Private Sub cmdremove_Click() '**
'Dim listcount As Double, i As Double, p As Double
'listcount = lbxcityselected.listcount
'    If lbxcityselected.ListIndex <> -1 Then
'        For i = 0 To listcount
'            If lbxcityselected.Selected(i) = True Then
'                For p = 0 To selectioncount
'                    If selection(p) = lbxcityselected.Text Then
'                        selection(p).Clear
'
'                    End If
'                Next
'                lbxcityselected.RemoveItem (i)
'            End If
'        Next i
'    End If

    Call AddDelValues(lbxcityselected, lbxcityoptions)
    Call SortListBox(lbxcityoptions)

End Sub
Private Sub cmdselectall_Click() '**
 Dim count As Double, i As Double
  count = lbxcityoptions.listcount
  For i = 0 To count - 1
    lbxcityoptions.selected(i) = True
  Next i
End Sub
Private Sub cmdselectall2_Click() '**
 Dim count As Double, i As Double
  count = lbxcityselected.listcount
  For i = 0 To count - 1
    lbxcityselected.selected(i) = True
  Next i
End Sub
Private Sub AddDelValues(objFromListBox As MSForms.ListBox, objToListBox As MSForms.ListBox) '**
Dim varItm As Variant
Dim aryItm() As Variant
Dim intCtr As Integer
Dim n As Integer
Dim listcount As Double
Dim i As Integer
Dim ind As Variant
Dim counter As Integer
Dim selected As Boolean
listcount = objFromListBox.listcount - 1
intCtr = 0
      
        For i = 0 To listcount
            If objFromListBox.selected(i) = True Then
                selected = True
                ReDim Preserve aryItm(intCtr)
                aryItm(intCtr) = i
                intCtr = intCtr + 1
                objToListBox.AddItem objFromListBox.list(i)
            End If
        Next i

'For each item we added
If selected = True Then
For Each ind In aryItm()
objFromListBox.RemoveItem ind - counter
counter = counter + 1
Next ind

End If
End Sub


Private Sub AddWithNoDuplicates(lst As MSForms.ListBox, str As String) '**

    Dim tf As Boolean
    Dim x As Integer

    If lst.listcount = 0 And str <> "" Then
        lbxcityoptions.AddItem str
'check to make sure no duplicate entries get into list box
    Else
        tf = False
        For x = 0 To lst.listcount - 1
            If lst.list(x) <> str Then
                tf = False
            Else
                tf = True
                Exit For
            End If
        Next x
        If tf = False And str <> "" Then
            lbxcityoptions.AddItem str
        End If

End If
End Sub

Private Sub SortListBox(objListBox As MSForms.ListBox) '**
Dim intFirst As Integer
Dim intLast As Integer
Dim intNumItems As Integer
Dim i As Integer
Dim j As Integer
Dim strTemp As String
Dim MyArray() As Variant
Dim item As Variant
Dim count As Integer

If objListBox.listcount <> 0 Then
'Re-Dim the array
ReDim MyArray(objListBox.listcount - 1)

'Get upper and lower boundary
intFirst = LBound(MyArray)
intLast = UBound(MyArray)

'Set array values
For i = LBound(MyArray) To UBound(MyArray)
MyArray(i) = objListBox.list(i)
Next i

'Loop through array values to determine sort
For i = intFirst To intLast - 1
For j = i + 1 To intLast
If CStr(Right(MyArray(i), 2)) > CStr(Right(MyArray(j), 2)) And MyArray(i) > MyArray(j) Then
strTemp = MyArray(j)
MyArray(j) = MyArray(i)
MyArray(i) = strTemp
End If
Next j
Next i

'Remove all items
For i = 0 To objListBox.listcount - 1
objListBox.RemoveItem i - count
count = count + 1
Next i

'Add all items in order
For i = 0 To UBound(MyArray)
objListBox.AddItem MyArray(i)
Next i

End If
End Sub

Function regionexist(ByRef existregion() As String, ByVal region As String) As Boolean '**
Dim gcell As Variant
regionexist = False
For Each gcell In existregion()
    If gcell = region Then
        regionexist = True
    End If
Next gcell
End Function
Function stateexist(ByRef existstate() As String, ByVal state As String) As Boolean '**
Dim rcell As Variant
stateexist = False
For Each rcell In existstate()
    If rcell = state Then
        stateexist = True
    End If
Next rcell
End Function
Private Sub cmdoptimize_Click()
'''''''''''''''''''''''''''' speed up process and free up memory
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Error Check
If tbxinvestment.Value = "" Then
MsgBox ("Please input an investment!")
Exit Sub
End If

If tbxradius.Value = "" Then
MsgBox ("Please input a radius!")
Exit Sub
End If

Dim listcount As Double, k As Integer
listcount = lbxcityselected.listcount - 1

ReDim selection(listcount)
If lbxcityselected.listcount <> 0 Then
    For k = 0 To listcount
     selection(k) = lbxcityselected.list(k)
    Next k
Else
MsgBox ("You must add at least one city!")
End If

Call setprimeinfo

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Loop for each Checkpoint
Dim checkpointstart As Range, checkpointcount As Double, i As Double, thelongit As Double, thelat As Double
Dim maxprofit As Double, bestmaxprofit As Double, bestlocation As String, bestlat As Double, bestlongit As Double
Dim bestresultindex As Double
bestmaxprofit = 0
maxprofit = 0
i = 1
Set checkpointstart = Worksheets("Result Worksheet").Range("A6")

If checkpointstart.Offset(1, 0) <> "" Then
    checkpointcount = Range(checkpointstart.Offset(1, 0), checkpointstart.End(xlDown)).Rows.count
End If

For i = 1 To checkpointcount
    Worksheets("Result Worksheet").Range("H3") = checkpointstart.Offset(i, 3)
    thelongit = checkpointstart.Offset(i, 1)
    thelat = checkpointstart.Offset(i, 2)
    Call withinradius(thelongit, thelat)
    Call solver1
    Call solver2
    maxprofit = CDbl(Worksheets("Result Worksheet").Range("maxprofit"))
    
    '''' to Check '''
    Dim writeresult As Range
    Set writeresult = Worksheets("Result Worksheet").Range("E6")
    writeresult.Offset(i, 0) = maxprofit
    
    If maxprofit > bestmaxprofit Then
        bestmaxprofit = maxprofit
        bestlocation = checkpointstart.Offset(i, 0)
        bestlongit = thelongit
        bestlat = thelat
        bestresultindex = i
    End If
Next i

    Worksheets("Result Worksheet").Range("H3") = checkpointstart.Offset(bestresultindex, 3)
    thelongit = checkpointstart.Offset(bestresultindex, 1)
    thelat = checkpointstart.Offset(bestresultindex, 2)
    Call withinradius(thelongit, thelat)
    Call solver1
    Call solver2
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Dim count As Double, i As Long, check As String, check2 As String
'count = lbxcityselected.listcount
'ReDim selection(count)
'
'For i = 0 To count - 1
'    'lbxcityselected.ListIndex = i
'    lbxcityselected.ListIndex = i
''    If lbxcityselected.Selected(i) = False Then
'
'    check = lbxcityselected.Value
''    selection(i) = lbxcityselected.Value
''    End If
'Next i
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
Unload Me
Frmresult.Show

End Sub
Public Sub setprimeinfo()
Dim radius As Double, initialinvestment As Double

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Set raidus and initial investment
radius = CDbl(tbxradius.Value)
initialinvestment = CDbl(tbxinvestment.Value)

Worksheets("Result Worksheet").Range("B3") = radius
Worksheets("Result Worksheet").Range("B4") = initialinvestment

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' Clear Checkpoint Cities
Dim checkpointstart As Range, checkpointcount As Double

Set checkpointstart = Worksheets("Result Worksheet").Range("A6")
If checkpointstart.Offset(1, 0) <> "" Then
    checkpointcount = Range(checkpointstart.Offset(1, 0), checkpointstart.End(xlDown)).Rows.count
    Range(checkpointstart.Offset(1, 0), checkpointstart.Offset(checkpointcount, 4)).Clear
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Copy Checkpoint Cities
Dim cityselectedcount As Double, i As Double, p As Double, stringlength As Double, thecity As String, n As Double
Dim cityName As String, cityselected As Variant
n = 0
cityselectedcount = lbxcityselected.listcount

For Each cityselected In selection
    
    cityName = cityselected
    stringlength = Len(cityName) - 4
    thecity = Mid(cityName, 1, stringlength)
    
    p = 0
    For p = 0 To rowcount - 1
        If thecity = city(p) Then
        
            ReDim Preserve selectedlongitude(n)
            ReDim Preserve selectedlatitude(n)
            ReDim Preserve selectedoutput(n)
            
            selectedlongitude(n) = longit(p)
            selectedlatitude(n) = lat(p)
            selectedoutput(n) = output(p)
            
            n = n + 1
            checkpointstart.Offset(n, 0) = thecity
            checkpointstart.Offset(n, 1) = longit(p)
            checkpointstart.Offset(n, 2) = lat(p)
            checkpointstart.Offset(n, 3) = output(p)
        End If
    Next p
Next cityselected
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub
Public Sub withinradius(ByVal thelongit As Double, ByVal thelat As Double)
Dim withinradiusstart As Range
Set withinradiusstart = Worksheets("Result Worksheet").Range("F29")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Clean within radius
Dim withincount As Double
If withinradiusstart.Offset(1, 0) <> "" Then
    withincount = Range(withinradiusstart.Offset(1, 0), withinradiusstart.End(xlDown)).Rows.count
    Range(withinradiusstart.Offset(1, 0), withinradiusstart.Offset(withincount, 5)).ClearContents
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Check distance to each city and copy to solver
Dim p As Double, i As Double
p = 0
i = 1
For p = 0 To rowcount - 1
    If finddistance(thelongit, thelat, longit(p), lat(p)) < CDbl(tbxradius.Text) Then
        withinradiusstart.Offset(i, 0) = city(p)
        withinradiusstart.Offset(i, 1) = 0
        withinradiusstart.Offset(i, 2) = demand(p)
        
        withinradiusstart.Offset(i, 3) = price(p)
        withinradiusstart.Offset(i, 5) = finddistance(thelongit, thelat, longit(p), lat(p))
        withinradiusstart.Offset(i, -1) = state(p)
        withinradiusstart.Offset(i, 6) = lat(p)
        withinradiusstart.Offset(i, 7) = longit(p)
        i = i + 1
    End If
Next p
Range(withinradiusstart.Offset(1, 4), withinradiusstart.Offset(i - 1, 4)).Formula = "=PRODUCT(G30,H30)"
Range(withinradiusstart.Offset(1, 4), withinradiusstart.Offset(i - 1, 4)).name = "allsupplies"
Range("totalsupply").Formula = "=SUM(allsupplies)"
Range(withinradiusstart.Offset(1, 1), withinradiusstart.Offset(i - 1, 1)).name = "allonoroff"
Range(withinradiusstart.Offset(1, 3), withinradiusstart.Offset(i - 1, 3)).name = "allprices"
Range("maxprofit").Formula = "=SUMPRODUCT(allonoroff,allprices)"

End Sub
Public Function finddistance(ByVal longit1 As Double, ByVal lat1 As Double, longit2 As Double, lat2 As Double) As Double
    finddistance = (Application.WorksheetFunction.Acos(Cos(Application.WorksheetFunction.Radians(90 - lat1)) * Cos(Application.WorksheetFunction.Radians(90 - lat2)) + Sin(Application.WorksheetFunction.Radians(90 - lat1)) * Sin(Application.WorksheetFunction.Radians(90 - lat2)) * Cos(Application.WorksheetFunction.Radians(longit1 - longit2))) * 3958.756)
End Function
Public Sub solver1()
Dim result As Integer
Worksheets("Result Worksheet").Activate

solverreset

solverok setcell:=Range("maxoutput"), maxminval:=1, bychange:=Range("turbineqty")
solveradd cellref:=Range("turbinecost"), relation:=1, formulatext:=Range("initialinvestment")
solveradd cellref:=Range("turbineqty"), relation:=4
solveroptions assumelinear:=True, assumenonneg:=True

result = solversolve(userfinish:=True)
If result = 5 Then
    MsgBox ("The solution is infeasible with the current Initial Investment.")
End If

End Sub
Public Sub solver2()
Dim result As Integer
Worksheets("Result Worksheet").Activate

solverreset

solverok setcell:=Range("maxprofit"), maxminval:=1, bychange:=Range("allonoroff")
solveradd cellref:=Range("allonoroff"), relation:=5
solveradd cellref:=Range("totalsupply"), relation:=1, formulatext:=Range("totaloutput")

result = solversolve(userfinish:=True)
If result = 5 Then
    MsgBox ("The solution is infeasible with the current inputs.")
End If
End Sub


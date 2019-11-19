Attribute VB_Name = "Module1"
Option Explicit
Public city() As String, state() As String, region() As String, output() As Double, annaverage() As Double, lat() As Double _
, longit() As Double, existregion() As String, existstate() As String, selection() As String, selectedlatitude() As Double _
, selectedlongitude() As Double, selectedoutput() As Double, arrayOptions() As String, arraySelected() As String, arraywithinCity() As String, arraywithinDistance() As Double, arraywithinState() As String, arraywithinLatitude() As Double, arraywithinLongitude() As Double, arraywithinDemand() As Double, arraywithinPrice() As Double
Public price() As Double, avenconsumption() As Double, percentdemand() As Double, demand() As Double
Public cityMap As String

Public rowcount As Double
Public selectioncount As Double
Sub appstart() ''''' show the main form after clicking the START button in the welcome page
frmMain.Show
End Sub


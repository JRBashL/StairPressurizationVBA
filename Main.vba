' VBA Regular Module Main
Option Explicit

' This is the dictionary that will hold our Doors class instances
Public DoorsDict As New Scripting.Dictionary


Public Sub DoorDebugPrinterButton()
    Dim doorToPrint As DoorClass
    Dim doorKey As Variant

    doorKey= Range("AE5").Value
    
    ' Error Handling
    If doorKey = "" Then 
        Exit Sub
    End If

    If DoorsDict.Exists(doorKey) Then
        Set doorToPrint = DoorsDict(doorKey)
        DoorDebugPrinter doorToPrint
    Else
        Exit Sub
    End If
End Sub
    
Public Sub CreateDoorsDictionaryButton()
    CreateDoorsDictionary DoorsDict
End Sub

Public Sub PopulateOpeningDoorForceData()
    Dim doorToPopulate As DoorClass
    Dim ws As Worksheet
    Dim doorCell As Variant

    Set ws = Worksheets("Opening Door Force")
    doorCell = ws.Range("B9").Value

    If doorCell = "" Then 
        Exit Sub
    ElseIf Not DoorsDict.Exists(doorCell) Then
        Exit Sub
    Else
        doorToPopulate = DoorsDict(doorCell)
    End If

    ' Print Values
    ws.Range("B13").Value = doorToPopulate.P_Width
    ws.Range("H13").Value = doorToPopulate.P_Width
    ws.Range("B14").Value = doorToPopulate.P_SingleDoorArea
    ws.Range("H14").Value = doorToPopulate.P_SingleDoorArea
    ws.Range("B16").Value = doorToPopulate.P_HandleDistance
    ws.Range("H16").Value = doorToPopulate.P_HandleDistance
End Sub

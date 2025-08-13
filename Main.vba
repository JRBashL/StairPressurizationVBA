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
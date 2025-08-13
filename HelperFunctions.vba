' Regular Module HelperFunctions
Option Explicit

Public Sub CreateDoorsDictionary()
    
    ' This is the dictionary that will hold our Doors class instances
    Dim doorsDict As Object
    Dim newDoor As DoorClass
    Dim wsDoors As Worksheet

    Dim doorCount As Long
    Dim checkCell As Variant
    Dim dictKey As Variant
    Dim outputRow As Long

    ' Set up the dictionary object and the worksheet
    Set doorsDict = CreateObject("Scripting.Dictionary")
    Set wsDoors = ThisWorkbook.Worksheets("Doors")

    ' Set up variables to pass to Doors constructor
    Dim doorName As Variant
    Dim doorType As Variant
    Dim doorWidth As Variant
    Dim doorHeight As Variant
    ' Dim doorArea As Variant not needed since DoorClass calculates area
    Dim doorHandleDistance As Variant
    Dim doorLeakageGap As Variant
    Dim doorLeakageType As Variant
    Dim doorLeakageArea As Variant
    
    doorCount = 0

    ' Define the cells to check for the TRUE value
    Dim checkCells As Variant
    checkCells = Array("F4", "L4", "R4", "X4", "F37", "L37", "R37", "X37", _
                       "F68", "L68", "R68", "X68", "F101", "L101", "R101", "X101")
    
    ' Loop through each cell in the array
    For Each checkCell In checkCells
        
        ' Check if the cell's value is TRUE
        If wsDoors.Range(checkCell).Value = True Then
            
            ' Increment the door counter
            doorCount = doorCount + 1
            
            ' Create a new instance of the Doors class
            Set newDoor = New DoorClass
            
            'Read the data off the cells
            doorName = checkCell.Offset(1, -3).Value
            doorType = checkCell.Offset(2, -3).Value
            doorWidth = checkCell.Offset(4, -3).Value
            doorHeight = checkCell.Offset(5, -3).Value
            'doorArea = checkCell.Offset() not needed since DoorClass calculates area
            doorHandleDistance = checkCell.Offset(8, -3).Value
            doorLeakageGap = checkCell.Offset(10, -3).Value
            doorLeakageType = checkCell.Offset(11, -3).Value
            doorLeakageArea = checkCell.Offset(12, -3).Value

            newDoor.Constructor checkCell.Value, doorName, doorType, doorWidth, doorHeight, doorHandleDistance, _
                                doorLeakageGap, doorLeakageType, doorLeakageArea
            
            ' Add the new door instance to the dictionary
            ' The key is the door's name, and the value is the instance itself
            ' This is just an example, a real name would be read from a cell.
            doorsDict.Add doorCount, newDoor
            
            ' Clean up the object variable for the next loop iteration
            Set newDoor = Nothing
            
        End If
        
    Next checkCell
    
    ' Output the dictionary keys and values to the worksheet

    ' Clear the output area first
    wsDoors.Range("AA5:AB" & wsDoors.Rows.Count).ClearContents

    outputRow = 5 ' Start from row 5
    For Each dictKey In doorsDict.Keys
        wsDoors.Range("AA" & outputRow).Value = dictKey
        ' For the value, we'll output the door's name from the class instance
        wsDoors.Range("AB" & outputRow).Value = doorsDict(dictKey).P_Name
        outputRow = outputRow + 1
    Next dictKey

    ' Clean up the objects after use
    Set doorsDict = Nothing
    Set wsDoors = Nothing

End Sub

Public Sub DoorDebugPrinter()
    Dim doortoPrint As DoorClass

End Sub

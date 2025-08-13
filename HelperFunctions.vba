' Regular Module HelperFunctions
Option Explicit

Public Sub CreateDoorsDictionary(ByRef a_doorsDict As Scripting.Dictionary)
    
    Dim newDoor As DoorClass
    Dim wsDoors As Worksheet

    Dim doorCount As Long
    Dim checkCell As Variant
    Dim dictKey As Variant
    Dim outputRow As Long

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
    
    ' Clean up the dictionary for refreshing in runtime
    a_doorsDict.RemoveAll

    Set wsDoors = ThisWorkbook.Worksheets("Doors")
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
            doorName = wsDoors.Range(checkCell).Offset(1, -3).Value
            doorType = wsDoors.Range(checkCell).Offset(2, -3).Value
            doorWidth = wsDoors.Range(checkCell).Offset(4, -3).Value
            doorHeight = wsDoors.Range(checkCell).Offset(5, -3).Value
            'doorArea = wsDoors.Range(checkCell).Offset() not needed since DoorClass calculates area
            doorHandleDistance = wsDoors.Range(checkCell).Offset(8, -3).Value
            doorLeakageGap = wsDoors.Range(checkCell).Offset(10, -3).Value
            doorLeakageType = wsDoors.Range(checkCell).Offset(11, -3).Value
            doorLeakageArea = wsDoors.Range(checkCell).Offset(12, -3).Value

            newDoor.Constructor wsDoors.Range(checkCell).Value, doorName, doorType, doorWidth, doorHeight, doorHandleDistance, _
                                doorLeakageGap, doorLeakageType, doorLeakageArea
            
            ' Add the new door instance to the dictionary
            ' The key is the door's name, and the value is the instance itself
            ' This is just an example, a real name would be read from a cell.
            a_doorsDict.Add newDoor.P_Name, newDoor
            
            ' Clean up the object variable for the next loop iteration
            Set newDoor = Nothing
            
        End If
        
    Next checkCell
    
    ' Output the dictionary keys and values to the worksheet

    ' Clear the output area first
    wsDoors.Range("AA6:AB22").ClearContents

    outputRow = 6 ' Start from row 6
    For Each dictKey In a_doorsDict.Keys
        wsDoors.Range("AA" & outputRow).Value = dictKey
        ' For the value, we'll output the door's name from the class instance
        wsDoors.Range("AB" & outputRow).Value = a_doorsDict(dictKey).P_Name
        outputRow = outputRow + 1
    Next dictKey

End Sub

Public Sub DoorDebugPrinter(ByVal a_doorToPrint As DoorClass)
    Dim wsDoors As Worksheet

    Set wsDoors = Worksheets("Doors")

    wsDoors.Range("AE6").Value = a_doorToPrint.P_UseDoor
    wsDoors.Range("AE7").Value = a_doorToPrint.P_Name
    wsDoors.Range("AE8").Value = a_doorToPrint.P_DoorType
    wsDoors.Range("AE9").Value = a_doorToPrint.P_Width
    wsDoors.Range("AE10").Value = a_doorToPrint.P_Height
    wsDoors.Range("AE11").Value = a_doorToPrint.P_SingleDoorArea
    wsDoors.Range("AE12").Value = a_doorToPrint.P_TotalArea
    wsDoors.Range("AE13").Value = a_doorToPrint.P_HandleDistance
    wsDoors.Range("AE14").Value = a_doorToPrint.P_LeakageGap
    wsDoors.Range("AE15").Value = a_doorToPrint.P_LeakageType
    wsDoors.Range("AE16").Value = a_doorToPrint.P_LeakageArea

End Sub

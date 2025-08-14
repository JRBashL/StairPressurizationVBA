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
        Set doorToPopulate = DoorsDict(doorCell)
    End If

    ' Print Values
    ws.Range("B13").Value = doorToPopulate.P_Width
    ws.Range("H13").Value = doorToPopulate.P_Width
    ws.Range("B14").Value = doorToPopulate.P_SingleDoorArea
    ws.Range("H14").Value = doorToPopulate.P_SingleDoorArea
    ws.Range("B16").Value = doorToPopulate.P_HandleDistance
    ws.Range("H16").Value = doorToPopulate.P_HandleDistance
End Sub

Public Sub PopulateLeakageCalcDoors()
    
    CreateDoorsDictionary DoorsDict

    Dim headCells As Variant
    Dim currentHeadCell As Variant
    Dim currentDoorRangeTopCellLeakage As Variant
    Dim currentDoorRangeTopCellTotalArea As Variant
    Dim entry As Variant

    Dim ws As Worksheet

    Dim rowOffsetLeakageArea As Integer
    Dim rowOffsetTotalArea As Integer
    Dim columnOffsetLeakageArea As Integer
    Dim columnOffsetTotalArea As Integer
    Dim rowResize As Integer
    Dim columnResize As Integer
    Dim targetRowLeakage As Integer
    Dim targetRowTotalArea As Integer
    Dim targetColumnLeakage As Integer
    Dim targetColumnTotalArea As Integer

    Dim currentDoorArrayLeakage As Variant
    Dim currentDoorArrayTotalArea As Variant

    Dim i As Variant
    Dim j As Variant

    rowOffsetLeakageArea = 3
    columnOffsetLeakageArea = 7
    rowOffsetTotalArea = 3
    columnOffsetTotalArea = 17
    rowResize = 8
    columnResize = 1

    Set ws = Worksheets("Leakage Calc")

    headCells = Array("A10", "A23", "A36", "A49", "A62", "A75", "A88", "A101", _
                       "A114", "A127", "A140", "A153", "A166", "A179", "A192", "A205")

    ' Loop through each Stairwell builder block
    For i = LBound(headCells) To UBound(headCells)

        ' Get the correct cells relative to the current top cell. the currentHeadCell is the top of the Stairwell Builder Block
        ' Set this cell as the top left of the merged cell
        Set currentHeadCell = ws.Range(headCells(i))

        ' Move to the top of the range. Set targetRows and targetColumns Doing it this way will ignore the merged cells.
        targetRowLeakage = currentHeadCell.Row + rowOffsetLeakageArea
        targetColumnLeakage = currentHeadCell.Column + columnOffsetLeakageArea
        targetRowTotalArea = currentHeadCell.Row + rowOffsetTotalArea
        targetColumnTotalArea = currentHeadCell.Column + columnOffsetTotalArea
        ' Resize to encapsulate the range
        Set currentDoorRangeTopCellLeakage = ws.Cells(targetRowLeakage, targetColumnLeakage)
        Set currentDoorRangeTopCellTotalArea = ws.Cells(targetRowTotalArea, targetColumnTotalArea)
        currentDoorArrayLeakage = currentDoorRangeTopCellLeakage.Resize(rowResize, columnResize).Value
        currentDoorArrayTotalArea = currentDoorRangeTopCellTotalArea.Resize(rowResize, columnResize).Value

        ' Scan through each in the currentDoorArray and populates the entire range if it's found
        For j = 1 To UBound(currentDoorArrayLeakage, 1)
            If DoorsDict.Exists(currentDoorArrayLeakage(j, 1)) Then
                currentDoorRangeTopCellLeakage.Offset(j - 1, 1).Value = DoorsDict(currentDoorArrayLeakage(j, 1)).P_LeakageArea
                currentDoorRangeTopCellTotalArea.Offset(j - 1, 0).Value = DoorsDict(currentDoorArrayLeakage(j, 1)).P_TotalArea
            Else
                currentDoorRangeTopCellLeakage.Offset(j - 1, 1).Value = ""
                currentDoorRangeTopCellTotalArea.Offset(j - 1, 0).Value = ""
            End If
        Next j
    Next i
End Sub

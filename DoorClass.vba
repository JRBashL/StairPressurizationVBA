' VBA Class Module: Doors

Option Explicit

' Private member variables to hold the door's properties
Private v_useDoor As Boolean
Private v_name As String
Private v_doorType As String
Private v_width As Single
Private v_height As Single
Private v_singleDoorArea As Single
Private v_totalArea As Single
Private v_handleDistance As Single
Private v_leakageGap As Single
Private v_leakageType As String
Private v_leakageArea As Single

' --- Constructor ---
Public Sub Constructor(ByVal a_useDoor As Boolean, ByVal a_name As String, ByVal a_doorType As String, _
                        ByVal a_width As Single, ByVal a_height As Single, ByVal a_handleDistance As Single, _              
                        ByVal a_leakageGap As Single, ByVal a_leakageType As String, ByVal a_leakageArea As Single)
    ' Assign the input parameters to the public properties
    v_useDoor = a_useDoor
    v_name = a_name
    v_doorType = a_doorType
    v_width = a_width
    v_height = a_height
    ' Area is calculated here and assigned directly to the private variable
    CalculateArea
    v_handleDistance = a_handleDistance
    v_leakageGap = a_leakageGap
    v_leakageType = a_leakageType
    v_leakageArea = a_leakageArea
End Sub

' --- Public Properties ---

' Property for v_usedoor (Boolean)
Public Property Let P_UseDoor(ByVal a_value As Boolean)
    v_useDoor = a_value
End Property

Public Property Get P_UseDoor() As Boolean
    P_UseDoor = v_useDoor
End Property

' Property for v_name (String)
Public Property Let P_Name(ByVal a_value As String)
    v_name = a_value
End Property

Public Property Get P_Name() As String
    P_Name = v_name
End Property

' Property for v_doorType (String)
Public Property Let P_DoorType (ByVal a_value As String) 
    v_doorType = a_value
    CalculateArea
End Property

Public Property Get P_DoorType() As String
    P_DoorType = v_doorType
End Property

' Property for v_width (Single)
Public Property Let P_Width(ByVal a_value As Single)
    v_width = a_value
    CalculateArea
End Property

Public Property Get P_Width() As Single
    P_Width = v_width
End Property

' Property for v_height (Single)
Public Property Let P_Height(ByVal a_value As Single)
    v_height = a_value
    CalculateArea
End Property

Public Property Get P_Height() As Single
    P_Height = v_height
End Property

' Property for v_singleDoorArea (Single). Private Set
Public Property Get P_SingleDoorArea() As Single
    P_SingleDoorArea = v_singleDoorArea
End Property

' Property for v_totalArea (Single). Private Set
Public Property Get P_TotalArea() As Single
    P_TotalArea = v_totalArea
End Property

' Property for v_handleDistance (Single)
Public Property Let P_HandleDistance(ByVal a_value As Single)
    v_handleDistance = a_value
End Property

Public Property Get P_HandleDistance() As Single
    P_HandleDistance = v_handleDistance
End Property

' Property for v_leakageGap (Single)
Public Property Let P_LeakageGap(ByVal a_value As Single)
    v_leakageGap = a_value
End Property

Public Property Get P_LeakageGap() As Single
    P_LeakageGap = v_leakageGap
End Property

' Property for v_leakageType (String)
Public Property Let P_LeakageType(ByVal a_value As String)
    v_leakageType = a_value
End Property

Public Property Get P_LeakageType() As String
    P_LeakageType = v_leakageType
End Property

' Property for v_leakageArea (Single)
Public Property Let P_LeakageArea(ByVal a_value As Single)
    v_leakageArea = a_value
End Property

Public Property Get P_LeakageArea() As Single
    P_LeakageArea = v_leakageArea
End Property

' --- Private Helper Methods ---

' Private method to encapsulate the area calculation logic
Private Sub CalculateArea()
    v_singleDoorArea = v_width * v_height

    Select Case v_doorType
        Case "Single"
            v_totalArea = v_singleDoorArea
        Case "Double"
            v_totalArea = 2 * v_singleDoorArea
        Case Else
            ' Handle other cases or set to 0 if the doorType is unknown
            v_singleDoorArea = 0
            v_totalArea = 0
    End Select
End Sub


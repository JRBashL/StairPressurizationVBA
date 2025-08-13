' VBA Class Module: Doors

Option Explicit

' Private member variables to hold the door's properties
Private v_useDoor As Boolean
Private v_name As String
Private v_doorType As String
Private v_width As Long
Private v_height As Long
Private v_area As Long
Private v_handleDistance As Long
Private v_leakageGap As Long
Private v_leakageType As String
Private v_leakageArea As Long

' --- Constructor ---
Public Sub Constructor(ByVal a_useDoor As Boolean, ByVal a_name As String, ByVal a_doorType As String _
                        ByVal a_width As Long, ByVal a_height As Long, ByVal a_handleDistance As Long, _              
                        ByVal a_leakageGap As Long, ByVal a_leakageType As String, ByVal a_leakageArea As Long)
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
    P_name = v_name
End Property

' Property for v_width (Long)
Public Property Let P_Width(ByVal a_value As Long)
    v_width = a_value
    CalculateArea
End Property

Public Property Get P_Width() As Long
    P_width = v_width
End Property

' Property for v_height (Long)
Public Property Let P_Height(ByVal a_value As Long)
    v_height = a_value
    CalculateArea
End Property

Public Property Get P_Height() As Long
    P_height = v_height
End Property

' Property for v_area (Long) is readonly
Public Property Get P_Area() As Long
    P_area = v_area
End Property

' Property for v_handleDistance (Long)
Public Property Let P_HandleDistance(ByVal a_value As Long)
    v_handleDistance = a_value
End Property

Public Property Get P_HandleDistance() As Long
    P_HandleDistance = v_handleDistance
End Property

' Property for v_leakageGap (Long)
Public Property Let P_LeakageGap(ByVal a_value As Long)
    v_leakageGap = a_value
End Property

Public Property Get P_LeakageGap() As Long
    P_LeakageGap = v_leakageGap
End Property

' Property for v_leakageType (String)
Public Property Let P_LeakageType(ByVal a_value As String)
    v_leakageType = a_value
End Property

Public Property Get P_LeakageType() As String
    P_LeakageType = v_leakageType
End Property

' Property for v_leakageArea (Long)
Public Property Let P_LeakageArea(ByVal a_value As Long)
    v_leakageArea = a_value
End Property

Public Property Get P_LeakageArea() As Long
    P_leakageArea = v_leakageArea
End Property

' --- Private Helper Methods ---

' Private method to encapsulate the area calculation logic
Private Sub CalculateArea()
    Select Case v_doorType
        Case "Single"
            v_area = v_width * v_height
        Case "Double"
            v_area = 2 * v_width * v_height
        Case Else
            ' Handle other cases or set to 0 if the doorType is unknown
            v_area = 0
    End Select
End Sub


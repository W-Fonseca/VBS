#If VBA7 Then
Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
#Else
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
#End If

Public Structure POINTAPI
Public x As Int64
Public y As Int64
End Structure
 
Public Sub Timer1_Timer()
    Dim PT As POINTAPI
    GetCursorPos (PT)
    System.Windows.Forms.Messagebox.Show("(" & PT.x & "," & PT.y & ")")
End Sub

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Integer, ByVal dx As Integer, ByVal dy As Integer, ByVal cButtons As Integer, ByVal dwExtraInfo As Integer)

Sub LocalMouse()
        Dim MousePosition As Point
        MousePosition = Cursor.Position
'System.Windows.Forms.Messagebox.Show(MousePosition.ToString)
EixoX = Cursor.Position.x
EixoY = Cursor.Position.y
        End Sub
        
        Sub ClickLeftMouse()
mouse_event(&H2, 0, 0, 0, 0)
mouse_event(&H4, 0, 0, 0, 0)
        End Sub
        
        Sub RightClickMouse()
mouse_event(&H8, 0, 0, 0, 0)
mouse_event(&H10, 0, 0, 0, 0)
        End Sub

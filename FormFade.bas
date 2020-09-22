Attribute VB_Name = "FormFade"
Option Explicit
Private J As Integer
Private Enum TransType
  byColor
  byValue
End Enum
Public Enum FadeSpeed
  f1 = 3    '85 transitional steps
  f2 = 5    '51
  f3 = 15   '17
  f4 = 17   '15
  f5 = 51   ' 5
  f6 = 85   ' 3
  f7 = 255  ' 1
End Enum

Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Public Sub StartFadeIn(FRM As Form, Optional fSpeed As FadeSpeed = f5)
  If Not isAPIFuncPresent("SetLayeredWindowAttributes", "User32") Then Exit Sub
  WindowTransparency FRM.hwnd, byValue, FRM, , 0
  FRM.Show
  For J = 0 To 255 Step fSpeed
    WindowTransparency FRM.hwnd, byValue, FRM, , J
    DoEvents
    FRM.Refresh
  Next J
End Sub

Public Sub StartFadeOut(FRM As Form, Optional fSpeed As FadeSpeed = f5)
  If Not isAPIFuncPresent("SetLayeredWindowAttributes", "User32") Then Exit Sub
  For J = 255 To 0 Step -(fSpeed)
    WindowTransparency FRM.hwnd, byValue, FRM, , J
    DoEvents
    FRM.Refresh
  Next J
End Sub

Private Sub CreateTransparentWindowStyle(lHwnd, FRM As Form)
  '-----------------------------------
  'this is used to create the window style needed
  'to allow transparency to be set/altered with
  'calls to SetLayeredWindowAttributes
  '-----------------------------------
  On Error GoTo Err_Handler:
  
  'VARIABLES:
  Dim Ret As Long
  'CODE:
  'Set the window style to 'Layered'
  Ret = GetWindowLong(lHwnd, GWL_EXSTYLE)
  Ret = Ret Or WS_EX_LAYERED
  SetWindowLong lHwnd, GWL_EXSTYLE, Ret
  'END CODE:
  
  Exit Sub
Err_Handler:
  Err.Source = Err.Source & "." & VarType(FRM) & ".ProcName"
  MsgBox Err.Number & vbTab & Err.Source & Err.Description
  Err.Clear
  Resume Next
End Sub

Private Sub WindowTransparency(lHwnd&, TransparencyBy As TransType, FRM As Form, _
                                      Optional Clr As Long, _
                                      Optional TransVal As Integer)
On Error GoTo Err_Handler:
'---------------------------------
'sets window transparency
'proper window style must be set first
'with call to CreateTransparentWindowStyle
'that call only has to be made once for the
'life of the form.  After that, this sub
'may be called multiple times by itself
'---------------------------------
'CODE:
  'first create the window style cabable of transparancies
  Call CreateTransparentWindowStyle(lHwnd, FRM)
  
  If TransparencyBy = byColor Then
    'the color specified in Clr becomes totally transparent
    SetLayeredWindowAttributes lHwnd, Clr, 0, LWA_COLORKEY
       
  ElseIf TransparencyBy = byValue Then
    If TransVal < 0 Or TransVal > 255 Then
      'makes sure valid transparency number chosen
      '0=totally opaque    255= totally transparent
      Err.Raise 2222, "Sub WindowTransparency", _
              "must choose number between 0-255"
    End If
    SetLayeredWindowAttributes lHwnd, 0, TransVal, LWA_ALPHA
  End If
'END CODE:
Exit Sub
Err_Handler:
  Err.Source = Err.Source & "." & VarType(FRM) & ".ProcName"
  MsgBox Err.Number & vbTab & Err.Source & Err.Description
  Err.Clear
  Resume Next
End Sub

Private Function isAPIFuncPresent(ByVal FuncName As String, ByVal DllName As String) As Boolean

  Dim lHandle As Long
  Dim lAddr  As Long

  lHandle = LoadLibrary(DllName)
  If lHandle <> 0 Then
    lAddr = GetProcAddress(lHandle, FuncName)
    FreeLibrary lHandle
  End If
  
  isAPIFuncPresent = (lAddr <> 0)

End Function


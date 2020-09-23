Attribute VB_Name = "Module_TransparentLayer"
Option Explicit


Public Const LWA_COLORKEY = 1
Public Const LWA_ALPHA = 2
Public Const LWA_BOTH = 3
Public Const WS_EX_LAYERED = &H80000
Public Const GWL_EXSTYLE = -20

Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal color As Long, ByVal X As Byte, ByVal alpha As Long) As Boolean
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long


Dim g_nTransparency As Integer
Dim ColorX As Long
Dim flag As Byte
Sub SetTranslucent(ThehWnd As Long, color As Long, nTrans As Integer, flag As Byte)
    On Error GoTo ErrorRtn
    'SetWindowLong and SetLayeredWindowAttributes are API functions, see MSDN for details
    Dim attrib As Long
    attrib = GetWindowLong(ThehWnd, GWL_EXSTYLE)
    SetWindowLong ThehWnd, GWL_EXSTYLE, attrib Or WS_EX_LAYERED
    'anything with color value color will completely disappear if flag = 1 or flag = 3
    SetLayeredWindowAttributes ThehWnd, color, nTrans, flag
    Exit Sub
ErrorRtn:
    'MsgBox Err.Description & " Source : " & Err.Source
    
End Sub
Sub TampilanTransNormal(ByVal TheObject As Long)
    'fungsi untuk membuat tampilan form normal
    
    'buat settingan attribute untuk transparansi
    flag = 0
    flag = flag Or LWA_COLORKEY
    flag = flag Or LWA_ALPHA
    g_nTransparency = 255
    SetTranslucent TheObject, vbRed, g_nTransparency, flag 'buat tampilan kartu kredit normal
    
End Sub



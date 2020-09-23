Attribute VB_Name = "Module_General"
Option Explicit

Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

Public SedangDiPake As Boolean
Public JenisTransaksi As Integer
Public TotUangTunai As Long

Sub CreatePecahanUang(ByVal PecahanUang As Long, ByVal TotalUang As Long, ByVal X As Long, ByVal Y As Long)
    Dim TmpUT As Integer
    TmpUT = CInt(TotalUang \ PecahanUang)
    'MsgBox TmpUT, , PecahanUang & " " & TotalUang
    CreateUang TmpUT, PecahanUang, X, Y
End Sub

Sub CreateUang(ByVal JumlahUT As Integer, ByVal NilaiNya As Long, ByVal Letak_X As Long, ByVal Letak_Y As Long)
    Dim xi As Integer
    Dim NewForm As New FormMoney
    
    For xi = 1 To JumlahUT
        'LetakTimbul = LetakTimbul + 300
        Set NewForm = New FormMoney
        NewForm.LValue = NilaiNya
        NewForm.Top = Letak_Y
        NewForm.Left = Letak_X
        NewForm.Height = 100
        NewForm.PicTunai.Top = -(NewForm.PicTunai.Height)
        NewForm.Show
        NewForm.TimerMundur.Enabled = True
        Do While NewForm.TimerMundur.Enabled = True
            DoEvents
        Loop
        Set NewForm = Nothing
        'MsgBox Nominalnya, , xi
    Next
End Sub


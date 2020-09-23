VERSION 5.00
Begin VB.Form FormBarang 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1590
   LinkTopic       =   "Form2"
   ScaleHeight     =   1335
   ScaleWidth      =   1590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "###"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "FormBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Top = Form1.Top + (Form1.PicOut.Top * 15) + 200
Me.Left = Form1.Left + (Form1.PicOut.Left * 15) + 100
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Call ReleaseCapture
    LngValReturn = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1_MouseDown Button, Shift, X, Y
End Sub



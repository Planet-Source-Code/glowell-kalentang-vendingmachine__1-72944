VERSION 5.00
Begin VB.Form FormCardMenu 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2295
   LinkTopic       =   "Form2"
   ScaleHeight     =   4800
   ScaleWidth      =   2295
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Timer TimerFocus 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   840
      Top             =   0
   End
   Begin VB.CommandButton CmdBtn 
      BackColor       =   &H00E0E0E0&
      Caption         =   "9"
      Height          =   615
      Index           =   9
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton CmdBtn 
      BackColor       =   &H00E0E0E0&
      Caption         =   "8"
      Height          =   615
      Index           =   8
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton CmdBtn 
      BackColor       =   &H00E0E0E0&
      Caption         =   "7"
      Height          =   615
      Index           =   7
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton CmdBtn 
      BackColor       =   &H00E0E0E0&
      Caption         =   "6"
      Height          =   615
      Index           =   6
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton CmdBtn 
      BackColor       =   &H00E0E0E0&
      Caption         =   "5"
      Height          =   615
      Index           =   5
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton CmdBtn 
      BackColor       =   &H00E0E0E0&
      Caption         =   "4"
      Height          =   615
      Index           =   4
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton CmdBtn 
      BackColor       =   &H00E0E0E0&
      Caption         =   "3"
      Height          =   615
      Index           =   3
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton CmdBtn 
      BackColor       =   &H00E0E0E0&
      Caption         =   "2"
      Height          =   615
      Index           =   2
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton CmdBtn 
      BackColor       =   &H00E0E0E0&
      Caption         =   "1"
      Height          =   615
      Index           =   1
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton CmdBtn 
      BackColor       =   &H00E0E0E0&
      Caption         =   "0"
      Height          =   615
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label LValid 
      Caption         =   "Label2"
      Height          =   255
      Left            =   1320
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Shape ShapeFocus 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   5
      Height          =   4815
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PIN:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   960
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   4800
      Left            =   0
      Picture         =   "FormCCardMenu.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "FormCardMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CmdBtn_Click(Index As Integer)

    Text1.Text = Text1.Text & Index

End Sub

Private Sub Form_Activate()
'SedangDiPake = True
End Sub

Private Sub Form_Load()
    Me.Left = FormMenuBarang.Left + FormMenuBarang.Width + 1000
    Me.Top = FormMenuBarang.Top + 100
    TimerFocus.Enabled = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    TimerFocus.Enabled = False
    Form1.ShapeFocus1.Visible = False

    FormCard.Show
    FormCard.TimerInOut = True
    'SedangDiPake = False
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim LngValReturn As Long

If Button = 1 Then
    Call ReleaseCapture
    LngValReturn = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If

End Sub


Private Sub Text1_Change()
If LCase(Text1.Text) = LCase(FormCard.LPin.Caption) Then
    LValid.Caption = "1"
    FormMenuBarang.LTrans.Caption = "Kartu Kredit"
    FormMenuBarang.LTrans.ForeColor = &HFF00&
End If
End Sub

Private Sub TimerFocus_Timer()
    If ShapeFocus.Visible = False Then
        ShapeFocus.Visible = True
        Form1.ShapeFocus1.Visible = True
    Else
        ShapeFocus.Visible = False
        Form1.ShapeFocus1.Visible = False

    End If
End Sub



VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "VendingMachine Simulation"
   ClientHeight    =   8490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10545
   ForeColor       =   &H00404040&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   566
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   703
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Coba"
      Height          =   375
      Left            =   3600
      TabIndex        =   14
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   360
      ScaleHeight     =   6255
      ScaleWidth      =   2895
      TabIndex        =   3
      Top             =   960
      Width           =   2895
      Begin Project1.ucItem ucItem1 
         Height          =   1335
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   2355
         BackColor       =   8421504
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "001"
         CaptionNama     =   "Nokia N1"
         CaptionHarga    =   "700000"
         Picture         =   "Form1.frx":000C
         Picture         =   "Form1.frx":C839
      End
      Begin Project1.ucItem ucItem1 
         Height          =   1335
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   2355
         BackColor       =   8421504
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "002"
         CaptionNama     =   "Sony Ericsson"
         CaptionHarga    =   "450000"
         Picture         =   "Form1.frx":19066
         Picture         =   "Form1.frx":1AA03
      End
      Begin Project1.ucItem ucItem1 
         Height          =   1335
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   3120
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   2355
         BackColor       =   8421504
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "003"
         CaptionNama     =   "Nokia 3310"
         CaptionHarga    =   "210000"
         Picture         =   "Form1.frx":1C3A0
         Picture         =   "Form1.frx":1EA6C
      End
      Begin Project1.ucItem ucItem1 
         Height          =   1335
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   4560
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   2355
         BackColor       =   8421504
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "004"
         CaptionNama     =   "Flash-Kngstn 4 GB"
         CaptionHarga    =   "115000"
         Picture         =   "Form1.frx":21138
         Picture         =   "Form1.frx":22D83
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   7
         Height          =   6255
         Index           =   6
         Left            =   0
         Top             =   0
         Width           =   2895
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   3600
      ScaleHeight     =   6255
      ScaleWidth      =   2895
      TabIndex        =   2
      Top             =   960
      Width           =   2895
      Begin Project1.ucItem ucItem1 
         Height          =   1335
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   2355
         BackColor       =   8421504
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "005"
         CaptionNama     =   "CD-Rom Smsung"
         CaptionHarga    =   "120000"
         Picture         =   "Form1.frx":249CE
         Picture         =   "Form1.frx":26975
      End
      Begin Project1.ucItem ucItem1 
         Height          =   1335
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   2355
         BackColor       =   8421504
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "006"
         CaptionNama     =   "DVD-Comb LG"
         CaptionHarga    =   "200000"
         Picture         =   "Form1.frx":2891C
         Picture         =   "Form1.frx":2E87A
      End
      Begin Project1.ucItem ucItem1 
         Height          =   1335
         Index           =   6
         Left            =   120
         TabIndex        =   12
         Top             =   3120
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   2355
         BackColor       =   8421504
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "007"
         CaptionNama     =   "Flash-Nexus 16 GB"
         CaptionHarga    =   "300000"
         Picture         =   "Form1.frx":347D8
         Picture         =   "Form1.frx":3654A
      End
      Begin Project1.ucItem ucItem1 
         Height          =   1335
         Index           =   7
         Left            =   120
         TabIndex        =   13
         Top             =   4560
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   2355
         BackColor       =   8421504
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "008"
         CaptionNama     =   "Mouse-Genius"
         CaptionHarga    =   "50000"
         Picture         =   "Form1.frx":382BC
         Picture         =   "Form1.frx":3A514
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   7
         Height          =   6255
         Index           =   5
         Left            =   0
         Top             =   0
         Width           =   2895
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   7080
      ScaleHeight     =   1335
      ScaleWidth      =   1455
      TabIndex        =   1
      Top             =   1800
      Width           =   1455
      Begin VB.Shape Shape1 
         BackColor       =   &H00004000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         BorderStyle     =   0  'Transparent
         Height          =   135
         Index           =   26
         Left            =   510
         Top             =   870
         Width           =   135
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00004000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         BorderStyle     =   0  'Transparent
         Height          =   135
         Index           =   25
         Left            =   375
         Top             =   870
         Width           =   135
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00004000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         BorderStyle     =   0  'Transparent
         Height          =   135
         Index           =   24
         Left            =   240
         Top             =   870
         Width           =   135
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00004000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         BorderStyle     =   0  'Transparent
         Height          =   135
         Index           =   23
         Left            =   510
         Top             =   735
         Width           =   135
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00004000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         BorderStyle     =   0  'Transparent
         Height          =   135
         Index           =   22
         Left            =   375
         Top             =   735
         Width           =   135
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00004000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         BorderStyle     =   0  'Transparent
         Height          =   135
         Index           =   21
         Left            =   240
         Top             =   735
         Width           =   135
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00004000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         BorderStyle     =   0  'Transparent
         Height          =   135
         Index           =   20
         Left            =   510
         Top             =   600
         Width           =   135
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00004000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         BorderStyle     =   0  'Transparent
         Height          =   135
         Index           =   19
         Left            =   375
         Top             =   600
         Width           =   135
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00004000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         BorderStyle     =   0  'Transparent
         Height          =   135
         Index           =   18
         Left            =   240
         Top             =   600
         Width           =   135
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00004000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         BorderStyle     =   0  'Transparent
         Height          =   135
         Index           =   17
         Left            =   240
         Top             =   1005
         Width           =   135
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00004000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   255
         Index           =   0
         Left            =   240
         Top             =   240
         Width           =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00008000&
         BorderWidth     =   3
         Index           =   1
         X1              =   1200
         X2              =   1320
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004000&
         BorderWidth     =   3
         Index           =   0
         X1              =   720
         X2              =   720
         Y1              =   120
         Y2              =   1200
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         BorderWidth     =   7
         Height          =   1335
         Index           =   4
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.PictureBox PicOut 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   7560
      ScaleHeight     =   1695
      ScaleWidth      =   1695
      TabIndex        =   0
      Top             =   5400
      Width           =   1695
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Height          =   1695
         Index           =   2
         Left            =   0
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Vending Machine Simulation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   15
      Top             =   240
      Width           =   3495
   End
   Begin VB.Shape ShapeFocus2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Height          =   1335
      Left            =   7080
      Top             =   1800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape ShapeFocus1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Height          =   1455
      Left            =   8760
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   0  'Transparent
      Height          =   135
      Index           =   15
      Left            =   8880
      Top             =   2280
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   0  'Transparent
      Height          =   135
      Index           =   14
      Left            =   8880
      Top             =   2895
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   0  'Transparent
      Height          =   135
      Index           =   13
      Left            =   8880
      Top             =   2490
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   0  'Transparent
      Height          =   135
      Index           =   12
      Left            =   9015
      Top             =   2490
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   0  'Transparent
      Height          =   135
      Index           =   11
      Left            =   9150
      Top             =   2490
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   0  'Transparent
      Height          =   135
      Index           =   10
      Left            =   8880
      Top             =   2625
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   0  'Transparent
      Height          =   135
      Index           =   9
      Left            =   9015
      Top             =   2625
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   0  'Transparent
      Height          =   135
      Index           =   8
      Left            =   9150
      Top             =   2625
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   0  'Transparent
      Height          =   135
      Index           =   7
      Left            =   8880
      Top             =   2760
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   0  'Transparent
      Height          =   135
      Index           =   3
      Left            =   9015
      Top             =   2760
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   0  'Transparent
      Height          =   135
      Index           =   0
      Left            =   9150
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label LblExit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10320
      TabIndex        =   5
      Top             =   120
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   930
      Index           =   1
      Left            =   9720
      Picture         =   "Form1.frx":3C76C
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   120
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   135
      Index           =   1
      Left            =   7320
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   330
      Index           =   9
      Left            =   7200
      Picture         =   "Form1.frx":4420F
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1320
   End
   Begin VB.Shape Shape2 
      Height          =   855
      Left            =   7560
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Barang yang dibeli  tidak bisa ditukar atau dikembalikan!"
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   7560
      TabIndex        =   4
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   3090
      Index           =   11
      Left            =   7320
      Picture         =   "Form1.frx":4BCB2
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   930
      Index           =   10
      Left            =   9600
      Picture         =   "Form1.frx":53755
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   120
   End
   Begin VB.Image Image2 
      Height          =   1215
      Index           =   0
      Left            =   9480
      Picture         =   "Form1.frx":5B1F8
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   1410
      Index           =   8
      Left            =   8760
      Picture         =   "Form1.frx":5C008
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1305
   End
   Begin VB.Image Image1 
      Height          =   6570
      Index           =   0
      Left            =   3480
      Picture         =   "Form1.frx":63AAB
      Stretch         =   -1  'True
      Top             =   840
      Width           =   3120
   End
   Begin VB.Image Image1 
      Height          =   6570
      Index           =   4
      Left            =   240
      Picture         =   "Form1.frx":6B54E
      Stretch         =   -1  'True
      Top             =   840
      Width           =   3120
   End
   Begin VB.Image Image1 
      Height          =   3330
      Index           =   6
      Left            =   7080
      Picture         =   "Form1.frx":72FF1
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   2640
   End
   Begin VB.Image Image1 
      Height          =   855
      Index           =   3
      Left            =   120
      Picture         =   "Form1.frx":7410F
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   10335
   End
   Begin VB.Image Image1 
      Height          =   7290
      Index           =   5
      Left            =   6720
      Picture         =   "Form1.frx":7485C
      Stretch         =   -1  'True
      Top             =   720
      Width           =   3720
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Vending Machine Simulation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   1
      Left            =   255
      TabIndex        =   16
      Top             =   255
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   2
      Left            =   120
      Picture         =   "Form1.frx":7597A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   10335
   End
   Begin VB.Image Image1 
      Height          =   7050
      Index           =   7
      Left            =   120
      Picture         =   "Form1.frx":75B29
      Stretch         =   -1  'True
      Top             =   600
      Width           =   6960
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
FormBarang.Image1.Picture = ucItem1(0).Picture
FormBarang.Show
End Sub

Private Sub Form_Load()
Dim NewForm As New FormMoney
Me.Left = 200
Me.Top = (Screen.Height \ 2) - (Me.Height \ 2)

'SedangDiPake = False



FormCard.Show

CreateUang 2, 50000, Screen.Width - 2000, 6000
CreateUang 3, 100000, Screen.Width - 2000, 6500
CreateUang 2, 20000, Screen.Width - 2000, 7000

TotUangTunai = 0
'FormMoney.Show
'newform

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim LngValReturn As Long

If Button = 1 Then
    Call ReleaseCapture
    LngValReturn = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblExit.ForeColor = vbBlack
End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If (Index = 2 Or Index = 3) Then
    Form_MouseDown Button, Shift, X, Y
End If

End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseMove Button, Shift, X, Y

End Sub

Private Sub LblExit_Click()
'Dim xForm As Form
End


'Unload FormCardMenu
'Unload FormMenuBarang
'Unload Form1
'Unload FormCard
'Unload FormMoney

End Sub

Private Sub LblExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LblExit.ForeColor = vbRed
End Sub



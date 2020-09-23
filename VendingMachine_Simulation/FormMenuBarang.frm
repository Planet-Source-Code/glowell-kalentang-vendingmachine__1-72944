VERSION 5.00
Begin VB.Form FormMenuBarang 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   5520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6765
   LinkTopic       =   "Form2"
   ScaleHeight     =   5520
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4080
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   480
      ScaleHeight     =   313
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   385
      TabIndex        =   0
      Top             =   480
      Width           =   5775
      Begin VB.TextBox LNominal 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "- - - - - - - "
         Top             =   2160
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   615
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label CmdBatal 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         Caption         =   "BATAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   3600
         TabIndex        =   24
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label LNamaBrng 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "- - - - - - - "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   360
         Left            =   3000
         TabIndex        =   23
         Top             =   3000
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Barang:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   360
         Index           =   5
         Left            =   3000
         TabIndex        =   22
         Top             =   2640
         Width           =   1995
      End
      Begin VB.Label LBtn 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Index           =   10
         Left            =   840
         TabIndex        =   21
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label LHargaBrng 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "- - - - - - - "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   360
         Left            =   3000
         TabIndex        =   20
         Top             =   3840
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Barang:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   360
         Index           =   4
         Left            =   3000
         TabIndex        =   19
         Top             =   3480
         Width           =   2025
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00004000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   660
         Left            =   105
         Top             =   1425
         Width           =   2220
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nominal:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   360
         Index           =   3
         Left            =   3000
         TabIndex        =   18
         Top             =   1800
         Width           =   1260
      End
      Begin VB.Label LTanggal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "- - - - - - - "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   360
         Left            =   3000
         TabIndex        =   17
         Top             =   1320
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. Transaksi:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   360
         Index           =   2
         Left            =   3000
         TabIndex        =   16
         Top             =   960
         Width           =   2070
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaksi:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   360
         Index           =   1
         Left            =   3000
         TabIndex        =   15
         Top             =   120
         Width           =   1440
      End
      Begin VB.Label LTrans 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "- - - - - - - "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   360
         Left            =   3000
         TabIndex        =   14
         Top             =   480
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Barang:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   1680
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00004000&
         BorderWidth     =   5
         X1              =   184
         X2              =   184
         Y1              =   304
         Y2              =   8
      End
      Begin VB.Label LBtn 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Index           =   9
         Left            =   1560
         TabIndex        =   11
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label LBtn 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Index           =   8
         Left            =   840
         TabIndex        =   10
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label LBtn 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Index           =   7
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label LBtn 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Index           =   6
         Left            =   1560
         TabIndex        =   8
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label LBtn 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Index           =   5
         Left            =   840
         TabIndex        =   7
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label LBtn 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Index           =   4
         Left            =   120
         TabIndex        =   6
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label LBtn 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Index           =   3
         Left            =   1560
         TabIndex        =   5
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label LBtn 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Index           =   2
         Left            =   840
         TabIndex        =   4
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label LBtn 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label LBtn 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   4080
         Width           =   615
      End
      Begin VB.Label CmdOK 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   4920
         TabIndex        =   1
         Top             =   4320
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Shape ShapeFocus 
      BorderColor     =   &H000000FF&
      BorderWidth     =   7
      Height          =   5535
      Left            =   0
      Top             =   0
      Width           =   6735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   37
      Height          =   5535
      Index           =   4
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "FormMenuBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IndexItemX As Integer
Private Sub CmdOK_Click()
    Dim NewFormX As New FormBarang
    
    Form1.ShapeFocus2.Visible = False
    
    Dim LetakTimbul As Long
    Dim LetakTimbulx As Long
    

    If JenisTransaksi = 0 Then
        If IsNumeric(LHargaBrng.Caption) = True Then
            
            TotUangTunai = CLng(LNominal.Text) - CLng(LHargaBrng.Caption)
            Set NewFormX = New FormBarang
            NewFormX.Image1.Picture = Form1.ucItem1(IndexItemX).Picture
            NewFormX.Label1.Caption = Form1.ucItem1(IndexItemX).CaptionNama
            NewFormX.Show
            Set NewFormX = Nothing
        End If
        
        If TotUangTunai > 1000 Then
                LetakTimbul = Form1.Top + (Form1.Image1(9).Top * 15) + 300
            LetakTimbulx = Form1.Left + (Form1.Image1(9).Left * 15) + 200
            
            CreatePecahanUang 100000, TotUangTunai, LetakTimbulx, LetakTimbul
            TotUangTunai = TotUangTunai Mod 100000
            
            CreatePecahanUang 50000, TotUangTunai, LetakTimbulx, LetakTimbul
            TotUangTunai = TotUangTunai Mod 50000
            
            CreatePecahanUang 20000, TotUangTunai, LetakTimbulx, LetakTimbul
            TotUangTunai = TotUangTunai Mod 20000
            
            CreatePecahanUang 10000, TotUangTunai, LetakTimbulx, LetakTimbul
            TotUangTunai = TotUangTunai Mod 10000
            
            CreatePecahanUang 5000, TotUangTunai, LetakTimbulx, LetakTimbul
            TotUangTunai = TotUangTunai Mod 5000
        
            CreatePecahanUang 1000, TotUangTunai, LetakTimbulx, LetakTimbul
            'TotUangTunai = TotUangTunai Mod 1000
        End If
    Else
        
        If FormCardMenu.LValid.Caption <> "1" Then
            LTrans.ForeColor = vbRed
            LTrans.Caption = "Ditolak!"
            Exit Sub
        Else
            
 
            Set NewFormX = New FormBarang
            NewFormX.Image1.Picture = Form1.ucItem1(IndexItemX).Picture
            NewFormX.Label1.Caption = Form1.ucItem1(IndexItemX).CaptionNama
            NewFormX.Show
            Set NewFormX = Nothing
        End If
        
    
    End If
FinishNya:
    Form1.ShapeFocus2.Visible = False
    Unload FormCardMenu
    Unload FormMenuBarang
    
    TotUangTunai = 0
End Sub

Private Sub Form_Activate()
SedangDiPake = True
End Sub

Private Sub Form_Load()
    Me.Left = Form1.Left + 3000
    Me.Top = Form1.Top + (Form1.Height \ 2) - 1000
    Timer1.Enabled = True
    LTanggal.Caption = Format(Month(Now) & ":" & Day(Now) & ":" & Year(Now), "MM-DD-YYYY") 'Month(Now) & ":" & Day(Now) & ":" & Year(Now)

    If JenisTransaksi = 0 Then
        LTrans.Caption = "Uang Tuani"
        LNominal.Text = TotUangTunai
    ElseIf JenisTransaksi = 1 Then
        LTrans.Caption = "Kartu Kredit"
    End If
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim LngValReturn As Long

If Button = 1 Then
    Call ReleaseCapture
    LngValReturn = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub

Private Sub CmdBatal_Click()

    Dim LetakTimbul As Long
    Dim LetakTimbulx As Long
    
    Form1.ShapeFocus2.Visible = False
    Unload FormCardMenu
    Unload FormMenuBarang
    
    LetakTimbul = Form1.Top + (Form1.Image1(9).Top * 15) + 300
    LetakTimbulx = Form1.Left + (Form1.Image1(9).Left * 15) + 200
    
    If TotUangTunai > 10000 Then
        CreatePecahanUang 100000, TotUangTunai, LetakTimbulx, LetakTimbul
        TotUangTunai = TotUangTunai Mod 100000
        
        CreatePecahanUang 50000, TotUangTunai, LetakTimbulx, LetakTimbul
        TotUangTunai = TotUangTunai Mod 50000
        
        CreatePecahanUang 20000, TotUangTunai, LetakTimbulx, LetakTimbul
        TotUangTunai = TotUangTunai Mod 20000
        
        CreatePecahanUang 10000, TotUangTunai, LetakTimbulx, LetakTimbul
        TotUangTunai = TotUangTunai Mod 10000
        
        CreatePecahanUang 5000, TotUangTunai, LetakTimbulx, LetakTimbul
        TotUangTunai = TotUangTunai Mod 5000
    
        CreatePecahanUang 1000, TotUangTunai, LetakTimbulx, LetakTimbul
        'TotUangTunai = TotUangTunai Mod 1000
    End If
    
    TotUangTunai = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
SedangDiPake = False
End Sub

Private Sub LBtn_Click(Index As Integer)

If Index = 10 Then
    Text1.Text = ""
Else
    Text1.Text = Text1.Text & Index
End If
End Sub


Private Sub LNominal_Change()
Dim xi As Integer
For xi = 0 To Form1.ucItem1.UBound
    If Form1.ucItem1(xi).Caption = Text1.Text Then
        LHargaBrng.Caption = Form1.ucItem1(xi).CaptionHarga
        LNamaBrng.Caption = Form1.ucItem1(xi).CaptionNama
        If JenisTransaksi = 0 Then
            If CLng(LNominal.Text) < CLng(LHargaBrng.Caption) Then
                LHargaBrng.ForeColor = vbRed
                LNominal.ForeColor = vbRed
                CmdOK.Visible = False
                IndexItemX = -1
            Else
                LHargaBrng.ForeColor = &HFF00&
                LNominal.ForeColor = &HFF00&
                CmdOK.Visible = True
                IndexItemX = xi
                'MsgBox IndexItemX
                Exit Sub
            End If
        End If
        Exit Sub
    End If
Next

LHargaBrng.Caption = "- - - - - - - "

CmdOK.Visible = False
IndexItemX = -1

End Sub


Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseDown Button, Shift, X, Y
End Sub

Private Sub Text1_Change()
Dim xi As Integer
For xi = 0 To Form1.ucItem1.UBound
    If Form1.ucItem1(xi).Caption = Text1.Text Then
        LHargaBrng.Caption = Form1.ucItem1(xi).CaptionHarga
        LNamaBrng.Caption = Form1.ucItem1(xi).CaptionNama
        If JenisTransaksi = 0 Then
            If CLng(LNominal.Text) < CLng(LHargaBrng.Caption) Then
                LHargaBrng.ForeColor = vbRed
                LNominal.ForeColor = vbRed
                CmdOK.Visible = False
                IndexItemX = -1
            Else
                LHargaBrng.ForeColor = &HFF00&
                LNominal.ForeColor = &HFF00&
                CmdOK.Visible = True
                IndexItemX = xi
                'MsgBox IndexItemX
                Exit Sub
            End If
        Else
            LHargaBrng.ForeColor = &HFF00&
            LNominal.ForeColor = &HFF00&
            CmdOK.Visible = True
            IndexItemX = xi
            Exit Sub
        End If
        Exit Sub
    End If
Next

LHargaBrng.Caption = "- - - - - - - "

CmdOK.Visible = False
IndexItemX = -1


End Sub

Private Sub Timer1_Timer()
    If FormMenuBarang.ShapeFocus.Visible = False Then
        ShapeFocus.Visible = True
        Form1.ShapeFocus2.Visible = True
    Else
        ShapeFocus.Visible = False
        Form1.ShapeFocus2.Visible = False
    End If
End Sub

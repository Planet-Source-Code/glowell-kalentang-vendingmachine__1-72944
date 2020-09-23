VERSION 5.00
Begin VB.Form FormMoney 
   BorderStyle     =   0  'None
   Caption         =   "Kartu Kredit"
   ClientHeight    =   1620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   855
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   108
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   57
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerInOut 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   600
      Top             =   720
   End
   Begin VB.Timer TimerLetak 
      Interval        =   15
      Left            =   480
      Top             =   1200
   End
   Begin VB.Timer TimerTrns 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   600
      Top             =   0
   End
   Begin VB.PictureBox PicTunai 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   0
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   0
      Top             =   120
      Width           =   855
      Begin VB.Timer TimerMundur 
         Enabled         =   0   'False
         Interval        =   15
         Left            =   600
         Top             =   120
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "RP."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   615
      End
      Begin VB.Label LValue 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100000"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   600
         Width           =   855
      End
   End
End
Attribute VB_Name = "FormMoney"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim g_nTransparency As Integer
Dim ColorX As Long
Dim flag As Byte
Dim TambahAlph As Boolean

'Dim CCreditIn As Boolean
Private Sub Form_Load()
g_nTransparency = 255
'CCreditIn = False

Me.Left = Screen.Width - (Me.Width + 1000)
Me.Top = (Screen.Height \ 2) - (Me.Height \ 2) + 2000

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim LngValReturn As Long
'If FormCardMenu.Visible = True Then Exit Sub
If Button = 1 And JenisTransaksi = 0 Then
    TimerLetak.Enabled = True
    Call ReleaseCapture
    LngValReturn = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    
    If TimerTrns.Enabled = True Then 'cek jika kartu berkedip
        'jika true, berarti berada di area pemasukkan kartu kredit
        
        TimerTrns.Enabled = False 'hentikan kedipan(fade-in/out) kartu kredit
        TimerLetak.Enabled = False
        
        TampilanTransNormal Me.hwnd  'normalkan tampilan transparansi formCard
  
        TimerInOut.Enabled = True 'jalankan animasi memasukkan/mengeluarkan kartu
        
        'sesuaikan letak kartu kredit dengan tempat masuk kartu
        Me.Top = (Form1.Top + (Form1.Image1(9).Top * 15) + 200)
        Me.Left = (Form1.Left + (Form1.Image1(9).Left * 15) + 200)
        
    End If

End If


End Sub



Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    TimerLetak.Enabled = False
    TimerTrns.Enabled = False
    

End Sub


Private Sub PicKredit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseDown Button, Shift, X, Y
End Sub


Private Sub PicKredit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseUp Button, Shift, X, Y
End Sub


Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseDown Button, Shift, X, Y
End Sub


Private Sub LValue_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseDown Button, Shift, X, Y
End Sub


Private Sub PicTunai_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseDown Button, Shift, X, Y
End Sub


Private Sub PicTunai_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseDown Button, Shift, X, Y
End Sub


Private Sub TimerInOut_Timer()
    'If CCreditIn = True Then
        If Me.Height > 100 Then
            Form1.Enabled = False
            PicTunai.Top = PicTunai.Top - 3
            Me.Height = Me.Height - 45
        Else
            'FormMoney.Hide
            TimerInOut.Enabled = False
            Form1.Enabled = True
            Form1.Show
            JenisTransaksi = 0
            TotUangTunai = TotUangTunai + CLng(LValue.Caption)
            
            FormMenuBarang.Show
            FormMenuBarang.LNominal.Text = TotUangTunai
            'StatusTunai = True
            'CCreditIn = False
            Unload Me
            
        End If
    'Else
    '    If FormMoney.Height < 1620 Then
    '        FormMoney.Show
    '        Form1.Enabled = False
    '        PicTunai.Top = PicTunai.Top + 3
    '        FormMoney.Height = FormMoney.Height + 45
    '    Else
    '        TimerInOut.Enabled = False
    '        Form1.Enabled = True
    '        'Form1.Show
    '    End If
    'End If
End Sub

Private Sub TimerLetak_Timer()
    Dim BatasAtas As Long
    Dim BatasBawah As Long
    Dim BatasX As Long
    
    'ambil koordinat area kedipan kartu kredit atau area memasukkan kartu kredit  ke variable
    BatasAtas = (Form1.Top + (Form1.Image1(9).Top * 15) - 100)
    BatasBawah = BatasAtas + ((Form1.Image1(9).Height \ 2) * 15) + 100
    BatasX = (Form1.Left + (Form1.Image1(9).Left * 15))
    
    'cek jika kartu berada di area pemasukkan kartu kredit
    If (Me.Top > BatasAtas And Me.Top < (BatasBawah)) And (Me.Left > BatasX And Me.Left < BatasX + (88 * 15)) Then
       TimerTrns.Enabled = True 'kedipkan kartu kredit
       'CCreditIn = True
    Else
        'CCreditIn = False
        TimerTrns.Enabled = False 'hentikan kedipan kartu kredit
        TampilanTransNormal Me.hwnd  'normalkan tampilan transparansi formCard
    End If


End Sub

Private Sub TimerMundur_Timer()
        If Me.Height < 1620 Then
            'FormMoney.Show
            'Form1.Enabled = False
            PicTunai.Top = PicTunai.Top + 3
            Me.Height = Me.Height + 45
        Else
            TimerMundur.Enabled = False
            'Form1.Enabled = True
            'Form1.Show
        End If

End Sub

Private Sub TimerTrns_Timer()
    'ColorX = &HFF00FF   '16714752 ' 16583680 '16713728 '16713728 'vbRed 'Me.Point((Me.Width) - 3, (Me.Height) - 3)   ' RGB(0, 0, 255)
    
    'cek jika transparansi lebih dari 255, benarkan untuk looping pengurangan transparansi
    If g_nTransparency >= 255 Then TambahAlph = False
    
    'cek jika transparansi kurang dari 255, benarkan untuk looping penambahan transparansi
    If g_nTransparency <= 75 Then TambahAlph = True
    
    If TambahAlph = False Then g_nTransparency = g_nTransparency - 20 'tambahkan transparansi dengan 20
    If TambahAlph = True Then g_nTransparency = g_nTransparency + 20 'kurangkan transparansi dengan 20
    
    'settingan attribut transparansi
    flag = 0
    flag = flag Or LWA_COLORKEY
    flag = flag Or LWA_ALPHA
    'g_nTransparency = 200 '125
    SetTranslucent Me.hwnd, vbRed, g_nTransparency, flag 'set transparansi formCard
End Sub



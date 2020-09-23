VERSION 5.00
Begin VB.Form FormCard 
   BorderStyle     =   0  'None
   Caption         =   "Kartu Kredit"
   ClientHeight    =   840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1200
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   56
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   80
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerInOut 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   480
      Top             =   600
   End
   Begin VB.Timer TimerLetak 
      Interval        =   15
      Left            =   840
      Top             =   600
   End
   Begin VB.Timer TimerTrns 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   960
      Top             =   360
   End
   Begin VB.PictureBox PicKredit 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   0
      Top             =   120
      Width           =   1215
      Begin VB.Label LPin 
         Caption         =   "0431"
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Card"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FormCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim g_nTransparency As Integer
Dim ColorX As Long
Dim flag As Byte
Dim TambahAlph As Boolean

Dim CCreditIn As Boolean

Private Sub Form_Activate()
    If SedangDiPake = True And JenisTransaksi = 1 Then
        Me.Top = (Form1.Top + (Form1.Image2(0).Top * 15) + 200)
        Me.Left = (Form1.Left + (Form1.Image2(0).Left * 15) + 300)
    
    End If
End Sub

Private Sub Form_Load()
g_nTransparency = 255
CCreditIn = False

Me.Left = Screen.Width - (Me.Width + 1000)
Me.Top = ((Screen.Height \ 2) - (Me.Height \ 2)) - 1000

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim LngValReturn As Long
If SedangDiPake = True Then Exit Sub

If Button = 1 Then 'and SedangDiPake = False Then ' SedangDiPake = False
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
        Me.Top = (Form1.Top + (Form1.Image2(0).Top * 15) + 200)
        Me.Left = (Form1.Left + (Form1.Image2(0).Left * 15) + 300)
        
    End If

End If


End Sub



Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    TimerLetak.Enabled = False
    TimerTrns.Enabled = False
    

End Sub


Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseDown Button, Shift, X, Y
End Sub


Private Sub PicKredit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseDown Button, Shift, X, Y
End Sub


Private Sub PicKredit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseUp Button, Shift, X, Y
End Sub


Private Sub TimerInOut_Timer()
    If CCreditIn = True Then
        If FormCard.Width > 100 Then
            Form1.Enabled = False
            PicKredit.Left = PicKredit.Left - 3
            FormCard.Width = FormCard.Width - 45
        Else
            FormCard.Visible = False
            TimerInOut.Enabled = False
            Form1.Enabled = True
            Form1.Show
            JenisTransaksi = 1
            FormMenuBarang.Show
            FormCardMenu.Show
            'SedangDiPake = True
            CCreditIn = False
            
        End If
    Else
        If FormCard.Width < 1245 Then
            FormCard.Visible = True
            Form1.Enabled = False
            PicKredit.Left = PicKredit.Left + 3
            FormCard.Width = FormCard.Width + 45
        Else
            'SedangDiPake = False
            TimerInOut.Enabled = False
            Form1.Enabled = True
            JenisTransaksi = 2
            'Form1.Show
        End If
    End If
End Sub

Private Sub TimerLetak_Timer()
    Dim BatasAtas As Long
    Dim BatasBawah As Long
    Dim BatasX As Long
    
    'ambil koordinat area kedipan kartu kredit atau area memasukkan kartu kredit  ke variable
    BatasAtas = (Form1.Top + (Form1.Image2(0).Top * 15) - 100)
    BatasBawah = BatasAtas + ((Form1.Image2(0).Height \ 2) * 15)
    BatasX = (Form1.Left + (Form1.Image2(0).Left * 15))
    
    'cek jika kartu berada di area pemasukkan kartu kredit
    If (Me.Top > BatasAtas And Me.Top < (BatasBawah)) And (Me.Left > BatasX And Me.Left < BatasX + 600) Then
       TimerTrns.Enabled = True 'kedipkan kartu kredit
       CCreditIn = True
    Else
        CCreditIn = False
        TimerTrns.Enabled = False 'hentikan kedipan kartu kredit
        TampilanTransNormal Me.hwnd  'normalkan tampilan transparansi formCard
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



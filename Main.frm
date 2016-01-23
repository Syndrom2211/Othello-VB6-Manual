VERSION 5.00
Begin VB.Form Main 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Othello Classic"
   ClientHeight    =   5970
   ClientLeft      =   8490
   ClientTop       =   3435
   ClientWidth     =   8280
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Main.frx":014A
   MousePointer    =   99  'Custom
   ScaleHeight     =   5970
   ScaleWidth      =   8280
   Begin VB.PictureBox BoardAktif 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4170
      Left            =   720
      ScaleHeight     =   278
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   290
      TabIndex        =   1
      Top             =   1320
      Width           =   4350
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Caption         =   "Status"
      ForeColor       =   &H8000000E&
      Height          =   1335
      Index           =   1
      Left            =   5280
      TabIndex        =   6
      Top             =   3240
      Width           =   2295
      Begin VB.Label PesanGiliran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   645
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label PesanGiliran 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.CommandButton Keluar 
      BackColor       =   &H8000000B&
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2400
      Width           =   2295
   End
   Begin VB.CommandButton PilihWarna 
      BackColor       =   &H8000000B&
      Caption         =   "Pilih Warna"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Timer Timer2 
      Left            =   7680
      Top             =   1560
   End
   Begin VB.Timer Timer1 
      Left            =   7680
      Top             =   2400
   End
   Begin VB.PictureBox BoardCom 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4170
      Left            =   720
      ScaleHeight     =   278
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   290
      TabIndex        =   0
      Top             =   1320
      Width           =   4350
   End
   Begin VB.PictureBox BoardBG 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   720
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   8
      Top             =   3360
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   5280
      Width           =   4350
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "v1.0"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   7800
      TabIndex        =   4
      Top             =   5760
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   915
      Left            =   720
      Picture         =   "Main.frx":029C
      Top             =   120
      Width           =   6915
   End
   Begin VB.Image Image2 
      Height          =   4785
      Left            =   0
      Picture         =   "Main.frx":8496
      Top             =   1200
      Width           =   8295
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================================================
'GAME OTHELLO
'
'UNTUK MEMENUHI SYARAT TUGAS BESAR ANALISIS ALGORITMA 8
'KELOMPOK 12
'
'FIRDAMDAM SASMITA, FAJAR, GUNGUN ABDULLAH
'=======================================================================
Option Explicit

' Inisialisasi Definisi Variable dari A-Z
DefInt A-Z

' Inisialisasi Variable bertipe Boolean
Dim Info As Boolean 'Inisialisasi Variable Informasi untuk pengujian program
Dim Udahan As Boolean 'Inisialisasi Variable untuk Mengakhiri Game
Dim akunungguinkamu As Boolean 'Inisialisasi Variable menunggu player

Dim TransparanMain As Integer 'Inisialisasi Variable Untuk Fade Window
Dim Petak   'Inisialisasi Variable untuk 1 petak
Dim C(9, 9) 'Inisialisasi Variable Putih, Hitam, atau Hijau <-- C itu CPU
Dim P(9, 9) 'Inisialisasi Variable prioritas posisi pada potongan <-- P itu Player
Dim KolomMusuh(9), BarisMusuh(9), Jumlah 'Inisialisasi Variable Untuk Menyimpan kemungkinan si musuh menemukan petak
Dim BarisIndexX(8), BarisIndexY(8) 'Inisialisasi Variable mencari nilai arah dari x dan y
Dim warnamu, warnaku 'Inisialisasi Variable untuk memilih warna petak
Dim barisku, kolomku 'Inisialisasi Variable untuk baris dan kolom si player
Dim barismu, kolommu 'Inisialisasi Variable untuk baris dan kolom si musuh
Dim Giliran 'Inisialisasi Variable Giliran setiap pemain (Maksimal 60x)
Dim R1, K1, R2, K2 'Inisialisasi Variable pencarian batas frame selama permainan

'Deklarasi Fungsi BitBLT & GDI32
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'Deklarasi Fungsi Fade
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal Color As Long, ByVal X As Byte, ByVal alpha As Long) As Boolean

'Inisialisasi Konstanta Skala Petak menggunakan bit DWORD
Const SRCCOPY = &HCC0020

'Inisialisasi Kondisi petak di board main
Const petakputih = 0
Const petakhitam = 3
Const petakhijau = 6

'Inisialisasi Konstanta Warna
Const Putih = 16777215 'Putih
Const Hitam = 0 'Hitam
Const Hijau = 2186785 'Hijau
 
'Inisialisasi Konstanta Untuk Fade
Const LWA_BOTH = 3
Const LWA_ALPHA = 2
Const LWA_COLORKEY = 1
Const GWL_EXSTYLE = -20
Const WS_EX_LAYERED = &H80000

'===================================BAGIANCODINGUMUM===================================
Private Sub Keluar_Click()
Unload Me
End Sub

Private Sub PilihWarna_Click()
PilihWarnanya
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    TransparanMain = TransparanMain + 5
    If TransparanMain > 255 Then
        TransparanMain = 255
        Timer1.Enabled = False
    End If
      BuatTrasparan Me.hWnd, TransparanMain
    Me.Show
End Sub

Private Sub Timer2_Timer()
    On Error Resume Next
    TransparanMain = TransparanMain - 5
    If TransparanMain < 0 Then
        TransparanMain = 0
        Timer2.Enabled = False
        End
    End If
    BuatTrasparan Me.hWnd, TransparanMain
End Sub

Public Sub BuatTrasparan(hWndBro As Long, iTransp As Integer)
    On Error Resume Next
 
    Dim ret As Long
    ret = GetWindowLong(hWndBro, GWL_EXSTYLE)
 
    SetWindowLong hWndBro, GWL_EXSTYLE, ret Or WS_EX_LAYERED
    SetLayeredWindowAttributes hWndBro, RGB(255, 255, 0), iTransp, LWA_ALPHA
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    Timer1.Enabled = False
    Timer2.Enabled = True
    End
End Sub

Private Sub Form_Load()
    Timer1.Enabled = False
    Timer1.Interval = 1

    Timer2.Enabled = False
    Timer2.Interval = 1

    Me.Visible = False
    Timer1.Enabled = True

    Petak = Int(BoardAktif.ScaleWidth / 9.5)
    SetKonstanta
    TampilField
    BoardAktif.BorderStyle = 0
    BoardCom.BorderStyle = 0
    PesanGiliran(0).BorderStyle = 0
    PesanGiliran(1).BorderStyle = 0
    BoardDesain
    Show
    PilihWarna.SetFocus
End Sub
'===================================BAGIANCODINGUMUM===================================

'==================================BAGIANCODINGUTAMA1===================================
Private Sub BoardDesain()
   Dim W As Long, H As Long, Wa As Long, Ha As Long
   Dim X As Long, Y As Long, dX As Long, dY As Long
   Dim I As Long, txt As String
   Static RandomSelesai As Boolean
   
   'Fungsi Random
   If RandomSelesai = False Then
      Randomize
      dX = BoardBG.ScaleWidth - 1
      dY = BoardBG.ScaleHeight - 1
      BoardBG.Cls
      For I = 0 To 500
         X = Int(dX * Rnd + 1)
         Y = Int(dY * Rnd + 1)
         BoardBG.PSet (X, Y), QBColor(8)
         BoardBG.PSet (X + 1, Y + 1), QBColor(15)
      Next I
      RandomSelesai = True
      End If
   
   'Petak
   W = 128: H = 128
   Wa = (Me.ScaleWidth \ W) + 1
   Ha = (Me.ScaleHeight \ H) + 1
   For Y = Ha To 0 Step -1
      For X = 0 To Wa
         BitBlt Me.hDC, X * W, Y * H, W, H, BoardBG.hDC, 0, 0, SRCCOPY
      Next X
   Next Y
End Sub

Private Sub SetKonstanta()
   Dim I, Baris, Kolom
   Dim txt As String
   Dim BarisDanKolom(1 To 8) As String
   
   'Bagian kosong diluar garis
   For I = 0 To 9
     C(I, 0) = petakhijau
     C(0, I) = petakhijau
     C(9, I) = petakhijau
     C(I, 9) = petakhijau
   Next I
   
   'Inisialisasi Pencarian
   R1 = 2: K1 = 2: R2 = 7: K2 = 7
   
   'X dan Y
   For I = 1 To 8
     BarisIndexX(I) = Choose(I, 1, 1, 0, -1, -1, -1, 0, 1)
     BarisIndexY(I) = Choose(I, 0, 1, 1, 1, 0, -1, -1, -1)
   Next I
   
   'Nilai Prioritas Awal
   BarisDanKolom(1) = "30 01 20 10 10 20 01 30"
   BarisDanKolom(2) = "01 01 03 03 03 03 01 01"
   BarisDanKolom(3) = "20 03 05 05 05 05 03 20"
   BarisDanKolom(4) = "10 03 05 00 00 05 03 10"
   BarisDanKolom(5) = "10 03 05 00 00 05 03 10"
   BarisDanKolom(6) = "20 03 05 05 05 05 03 20"
   BarisDanKolom(7) = "01 01 03 03 03 03 01 01"
   BarisDanKolom(8) = "30 01 20 10 10 20 01 30"
   For Baris = 1 To 8: For Kolom = 1 To 8
      P(Baris, Kolom) = Val(Mid(BarisDanKolom(Baris), (Kolom - 1) * 3 + 1, 3))
      C(Baris, Kolom) = petakhijau
   Next Kolom: Next Baris
      
   'Peletakan 4 petak pertama
   C(4, 4) = 3: C(4, 5) = 0
   C(5, 4) = 0: C(5, 5) = 3
   Giliran = 0
End Sub

Private Sub TampilField()
   Dim Baris, Kolom
   Dim X, Y
   Dim txt As String
   
   With BoardAktif
      .Cls
      ' BG
      BitBlt .hDC, 0, 0, .ScaleWidth, .ScaleHeight, Me.hDC, .Left, .Top, SRCCOPY
      
      .FontSize = 18 'Ukuran No
      For Kolom = 1 To 8
         txt = Format(Kolom)
         X = (Kolom * Petak + Petak / 2) - .TextWidth(txt) / 2
         Y = (Petak - .TextHeight(txt)) / 2
         PrintAt BoardAktif, txt, X + 1, Y + 1, Putih
         PrintAt BoardAktif, txt, X, Y, RGB(240, 248, 255) 'Warna Text Di Board
      Next Kolom
      
      For Baris = 1 To 8
         txt = Format(Baris)
         X = (Petak - .TextWidth(txt)) / 2
         Y = (Baris * Petak + Petak / 2) - .TextHeight(txt) / 2
         PrintAt BoardAktif, txt, X + 1, Y + 1, Putih
         PrintAt BoardAktif, txt, X, Y, RGB(240, 248, 255) 'Warna Text Di Board
      Next Baris
      
      'Petak - Petak
      For Baris = 1 To 8: For Kolom = 1 To 8
         Select Case C(Baris, Kolom)
           Case petakhitam: .FillColor = Hitam
           Case petakputih: .FillColor = Putih
           Case petakhijau: .FillColor = Hijau
         End Select
         BoardAktif.Line (Petak * Kolom, Petak * Baris)-Step(Petak, Petak), .FillColor, BF
         BoardAktif.Line (Petak * Kolom, Petak * Baris)-Step(Petak, Petak), QBColor(7), B
      Next Kolom, Baris
      
      If Info Then
         .FontSize = 12
         .ForeColor = RGB(255, 0, 0)
         For Baris = 1 To 8: For Kolom = 1 To 8
            txt = Format(P(Baris, Kolom))
            Select Case C(Baris, Kolom)
              Case petakhitam: .ForeColor = Putih
              Case petakputih: .ForeColor = Hitam
              Case petakhijau: .ForeColor = QBColor(8)
            End Select
            .CurrentX = Kolom * Petak + (Petak - .TextWidth(txt)) / 2
            .CurrentY = Baris * Petak + (Petak - .TextHeight(txt)) / 2 + 1
            BoardAktif.Print txt
         Next Kolom, Baris
       End If
         
   End With
End Sub

Private Sub PrintAt(pic As PictureBox, txt As String, X, Y, Color As Long)
   With pic
      .ForeColor = Color: .CurrentX = X: .CurrentY = Y
   End With
   pic.Print txt 'Output No
End Sub
'==================================BAGIANCODINGUTAMA1===================================

'==================================BAGIANCODINGUTAMA2===================================
Private Sub PilihWarnanya()
   If PilihWarna.Caption = "Pilih Warna" Then
      
      Dialog.Show 1, Me
      If Dialog.TombolOK = False Then Exit Sub
      If Dialog.Pilihanmu = 1 Then
         warnamu = petakputih: warnaku = petakhitam
         Else
         warnamu = petakhitam: warnaku = petakputih
         End If
      PilihWarna.Caption = "Udahan"
      Udahan = False
      Mainkan
      
      Else
      If MsgBox("Kamu yakin mau udahan ?", vbOKCancel, "Udahan") = vbCancel Then Exit Sub
      Reset
      Pesan "Udahan", Putih
      
      End If

End Sub

' Giliran Musuh Memilih Petak yang paling terbaik
Private Sub CariPetakTerbaik(Max, MaxSetiapKolom)
   Dim I, Baris, Kolom
   Dim SetiapKolom, PerKolom  ' angulars, counter or total of the cell <--
   Dim CariKolom, CariBaris     ' Mulai dari Baris dan Kolom, pencarian secara melingkar
   
   ' replace search frame ?
   If Not (R1 * K1 = 1 And R2 * K2 = 64) Then
      For I = 2 To 7
         If C(2, I) <> petakhijau Then R1 = 1
         If C(7, I) <> petakhijau Then R2 = 8
         If C(I, 2) <> petakhijau Then K1 = 1
         If C(I, 7) <> petakhijau Then K2 = 8
      Next I
      If Info Then BoardAktif.Line (K1 * Petak, R1 * Petak)-(K2 * Petak + Petak, R2 * Petak + Petak), RGB(255, 0, 0), B '<--
      End If

   ' go over all cells, select empty one's and examin them
   PerKolom = 0
   Max = 0: MaxSetiapKolom = 0
   For Baris = R1 To R2: For Kolom = K1 To K2
      If C(Baris, Kolom) = petakhijau Then
         If P(Baris, Kolom) < Max Then GoTo PencarianBerikutnya:
         PerKolom = 0
         For I = 0 To 8                    ' all directions
            SetiapKolom = 0
            CariKolom = Kolom: CariBaris = Baris
            CariKolom = CariKolom + BarisIndexX(I)
            CariBaris = CariBaris + BarisIndexY(I)
            While C(CariBaris, CariKolom) = warnamu    ' count cells in this
               SetiapKolom = SetiapKolom + 1                 ' direction
               CariKolom = CariKolom + BarisIndexX(I)
               CariBaris = CariBaris + BarisIndexY(I)
            Wend
            If C(CariBaris, CariKolom) <> petakhijau And SetiapKolom <> 0 Then PerKolom = PerKolom + SetiapKolom
         Next I
         If PerKolom <> 0 Then
            If P(Baris, Kolom) > Max Then
               Max = P(Baris, Kolom)
               Jumlah = 0              ' higher priority
               MaxSetiapKolom = PerKolom               ' so restart
               KolomMusuh(0) = Kolom: BarisMusuh(0) = Baris
               GoTo PencarianBerikutnya:
               End If
            If MaxSetiapKolom > SetiapKolom Then GoTo PencarianBerikutnya:
            If MaxSetiapKolom < SetiapKolom Then
               Jumlah = 0              ' more gain
               MaxSetiapKolom = SetiapKolom               ' so restart
               KolomMusuh(0) = Kolom
               BarisMusuh(0) = Baris
               Else
               Jumlah = Jumlah + 1 ' same priority
               KolomMusuh(Jumlah) = Kolom          ' same gain
               BarisMusuh(Jumlah) = Baris
               End If
            End If
         End If
            
PencarianBerikutnya:
      DoEvents
   Next Kolom: Next Baris
   If Info Then For I = 0 To Jumlah: LingkaranKecil BarisMusuh(I), KolomMusuh(I), Putih: Next I
      
End Sub

' Menetapkan Prioritas
Private Sub UpdatePrioritas(prior)
   Dim I, J
   
   If C(2, 2) = warnamu And C(3, 1) = warnaku Or C(1, 3) = warnaku Then P(3, 1) = 1: P(1, 3) = 1 ': Stop
   If C(7, 7) = warnamu And (C(8, 6) = warnaku Or C(6, 8) = warnaku) Then P(8, 6) = 1: P(6, 8) = 1 ': Stop
   If C(2, 7) = warnamu And (C(1, 6) = warnaku Or C(3, 8) = warnaku) Then P(1, 6) = 1: P(3, 8) = 1 ': Stop
   If C(7, 2) = warnamu And (C(6, 1) = warnaku Or C(8, 3) = warnaku) Then P(6, 1) = 1: P(8, 3) = 1 ': Stop
   
   ' examin if chosen cell is on the border
   If Not (Batas(barisku, kolomku) Or Batas(barismu, kolommu)) Then Exit Sub
   For J = 1 To 8 Step 7
      For I = 2 To 7
         If C(I, J) = warnaku Then P(I + 1, J) = 21: P(I - 1, J) = 21
         If C(J, I) = warnaku Then P(J, I + 1) = 21: P(J, I - 1) = 21
      Next I
      For I = 2 To 7
         If C(I, J) = warnamu Then P(I + 1, J) = 2: P(I - 1, J) = 2
         If C(J, I) = warnamu Then P(J, I + 1) = prior: P(J, I - 1) = 2
      Next I
   Next J
   P(1, 2) = 1: P(1, 7) = 1: P(2, 1) = 1: P(7, 1) = 1
   P(2, 8) = 1: P(7, 8) = 1: P(8, 2) = 1: P(8, 7) = 1
   For I = 2 To 7
      If C(1, I - 1) = warnamu And C(1, I + 1) = warnamu Then P(1, I) = 2
      If C(8, I - 1) = warnamu And C(8, I + 1) = warnamu Then P(8, I) = 2
      If C(I - 1, 1) = warnamu And C(I + 1, 1) = warnamu Then P(I, 1) = 2
      If C(I - 1, 8) = warnamu And C(I + 1, 8) = warnamu Then P(I, 8) = 2
   Next I
   Dim K
   For J = 1 To 8 Step 7
     For I = 4 To 8
       If C(J, I) = warnaku Then
         K = I - 1
         If C(J, K) <> petakhijau Then
           While C(J, K) = warnamu: K = K - 1: Wend
           If K > 0 Then
              If C(J, K) = petakhijau And K <> 0 And Not (C(J, I + 1) = warnamu And C(J, K - 1) = petakhijau) Then P(J, K) = 26
              End If
           End If
         End If
       If C(I, J) = warnaku Then
         K = I - 1
         If C(K, J) <> petakhijau Then
           While C(K, J) = warnamu: K = K - 1: Wend
           If K > 0 Then
              If C(K, J) = petakhijau And K <> 0 And Not (C(I + 1, J) = warnamu And C(K - 1, J) = petakhijau) Then P(K, J) = 26
              End If
           End If
         End If
     Next I
     For I = 1 To 5
       If C(J, I) = warnaku Then
         K = I + 1
         If C(J, K) <> petakhijau Then
           While C(J, K) = warnamu: K = K + 1: Wend
           If K < 9 Then
              If C(J, K) = petakhijau And Not (C(J, I - 1) = warnamu And C(J, K + 1) = petakhijau) Then P(J, K) = 26
              End If
           End If
         End If
       If C(I, J) = warnaku Then
         K = I + 1
         If C(K, J) = petakhijau Then GoTo L2440
         While C(K, J) = warnamu: K = K + 1: Wend
         If K < 9 Then
            If C(K, J) = petakhijau And Not (C(I - 1, J) = warnamu And C(K + 1, J) = petakhijau) Then P(K, J) = 26
            End If
       End If
     Next I
   Next J
'-----
L2440:
   If C(1, 1) = warnaku Then For I = 2 To 6: P(1, I) = 20: P(I, 1) = 20: Next I
   If C(1, 8) = warnaku Then For I = 2 To 6: P(I, 8) = 20: P(1, 9 - I) = 20: Next I
   If C(8, 1) = warnaku Then For I = 2 To 6: P(9 - I, 1) = 20: P(8, I) = 20: Next I
   If C(8, 8) = warnaku Then For I = 3 To 6: P(I, 8) = 20: P(8, I) = 20: Next I
   If C(1, 1) <> petakhijau Then P(2, 2) = 5
   If C(1, 8) <> petakhijau Then P(2, 7) = 5
   If C(8, 1) <> petakhijau Then P(7, 2) = 5
   If C(8, 8) <> petakhijau Then P(7, 7) = 5
   P(1, 1) = 30: P(1, 8) = 30: P(8, 1) = 30: P(8, 8) = 30
   For I = 3 To 6
      If C(1, I) = warnaku Then P(2, I) = 4
      If C(8, I) = warnaku Then P(7, 1) = 4
      If C(I, 1) = warnaku Then P(1, 2) = 4
      If C(I, 8) = warnaku Then P(1, 7) = 4
   Next I
   If C(7, 1) = warnamu And C(4, 1) = warnaku And C(6, 1) = petakhijau And C(5, 1) = petakhijau Then P(6, 1) = 26
   If C(1, 7) = warnamu And C(1, 4) = warnaku And C(1, 6) = petakhijau And C(1, 5) = petakhijau Then P(1, 6) = 26
   If C(2, 1) = warnamu And C(5, 1) = warnaku And C(3, 1) = petakhijau And C(4, 1) = petakhijau Then P(3, 1) = 26
   If C(1, 2) = warnamu And C(1, 5) = warnaku And C(1, 3) = petakhijau And C(1, 4) = petakhijau Then P(1, 3) = 26
   If C(8, 2) = warnamu And C(8, 5) = warnaku And C(8, 3) = petakhijau And C(5, 1) = petakhijau Then P(6, 1) = 26
   If C(2, 8) = warnamu And C(5, 8) = warnaku And C(3, 8) = petakhijau And C(1, 5) = petakhijau Then P(1, 6) = 26
   If C(8, 7) = warnamu And C(8, 4) = warnaku And C(8, 5) = petakhijau And C(8, 6) = petakhijau Then P(8, 6) = 26
   If C(7, 8) = warnamu And C(4, 8) = warnaku And C(8, 8) = petakhijau And C(6, 8) = petakhijau Then P(6, 8) = 26
End Sub

' Dapat Petak
Private Function DapatPetak(Baris, Kolom, CariWarna)
   Dim SetiapKolom
   Dim TotalKolom
   Dim CariKolom, CariBaris
   Dim I
   Dim UbahWarna
   
   UbahWarna = IIf(CariWarna = petakhitam, petakputih, petakhitam)
   
   ' Memeriksa 8 arah
   For I = 1 To 8
      SetiapKolom = 0
      CariKolom = Kolom: CariBaris = Baris    ' always start from this particular cell
      CariKolom = CariKolom + BarisIndexX(I)
      CariBaris = CariBaris + BarisIndexY(I)
      While C(CariBaris, CariKolom) = UbahWarna
         SetiapKolom = SetiapKolom + 1
         CariKolom = CariKolom + BarisIndexX(I)
         CariBaris = CariBaris + BarisIndexY(I)
      Wend
      If C(CariBaris, CariKolom) <> petakhijau And SetiapKolom <> 0 Then
         TotalKolom = TotalKolom + SetiapKolom
         CariKolom = CariKolom - BarisIndexX(I)
         CariBaris = CariBaris - BarisIndexY(I)
         While C(CariBaris, CariKolom) <> petakhijau
            C(CariBaris, CariKolom) = CariWarna   ' turn over
            CariKolom = CariKolom - BarisIndexX(I)
            CariBaris = CariBaris - BarisIndexY(I)
         Wend
         End If
   Next I
   DapatPetak = TotalKolom
End Function

' Tanda Silang
Private Sub TandaSilang(Baris, Kolom)
   Dim txt As String, Warna As Long
   Dim X, Y
   Warna = RGB(0, 255, 0)
   BoardAktif.Line (Petak * Kolom, Petak * Baris)-Step(Petak, Petak), Warna
   BoardAktif.Line (Petak * Kolom, Petak * Baris + Petak)-Step(Petak, -Petak), Warna
   BoardAktif.FontSize = 18
   BoardAktif.ForeColor = Warna
   
   ' Kolom Horizontal
   txt = Format(Kolom)
   X = (Kolom * Petak + Petak / 2) - BoardAktif.TextWidth(txt) / 2
   Y = (Petak - BoardAktif.TextHeight(txt)) / 2
   PrintAt BoardAktif, txt, X + 1, Y + 1, Hitam
   PrintAt BoardAktif, txt, X, Y, Warna
   
   ' Kolom Vertikal
   txt = Format(Baris)
   X = (Petak - BoardAktif.TextWidth(txt)) / 2
   Y = (Baris * Petak + Petak / 2) - BoardAktif.TextHeight(txt) / 2
   PrintAt BoardAktif, txt, X + 1, Y + 1, Hitam
   PrintAt BoardAktif, txt, X, Y, Warna
   SabarNunggu 500
End Sub

' Kondisi Menang
Private Sub KondisiMenang(pesanpesan As String)
   Dim Baris, Kolom
   Dim totaldia
   Dim totalaku
   Dim kondisi As String
   Dim kondisi2 As String
   
   Udahan = True
   totaldia = 0: totalaku = 0
   For Baris = 1 To 8: For Kolom = 1 To 8
      If C(Baris, Kolom) = warnaku Then totaldia = totaldia + 1 Else totalaku = totalaku + 1
   Next Kolom, Baris
   
   If totaldia = totalaku Then
      kondisi = "DRAW, keduanya memiliki jumlah petak yang sama."
      kondisi2 = "Tidak ada pilihan lain"
      End If
   If totaldia < totalaku Then
      kondisi = "Jumlah petak aku " & totalaku & " petak, dan dia " & totaldia & "."
      kondisi2 = "Player Menang !"
      End If
   If totaldia > totalaku Then
      kondisi = "Jumlah petak aku " & totalaku & " petak, dan dia " & totaldia & "."
      kondisi2 = "Musuh Menang!"
      End If
   MsgBox pesanpesan & vbCrLf & vbCrLf & kondisi & vbCrLf & vbCrLf & kondisi2
   Pesan kondisi2, Putih
      
End Sub

' Kondisi Dilewat
Private Function Dilewat() As Boolean
   Dim SetiapKolom, I, Baris, Kolom, CariBaris, CariKolom
   
   Dilewat = True
   For Baris = 1 To 8: For Kolom = 1 To 8   ' of all cells
      If C(Baris, Kolom) = petakhijau Then       ' look for empty one's
         For I = 1 To 8             ' in all directions
            SetiapKolom = 0: CariKolom = Kolom: CariBaris = Baris  ' remember RR,KK startcell
            Do
               CariBaris = CariBaris + BarisIndexY(I): CariKolom = CariKolom + BarisIndexX(I)
               If KeluarBatas(CariBaris, CariKolom) Then Exit Do
               If C(CariBaris, CariKolom) = warnaku Then SetiapKolom = SetiapKolom + 1
            Loop Until C(CariBaris, CariKolom) <> warnaku
            If C(CariBaris, CariKolom) = warnamu And SetiapKolom > 0 Then
               Dilewat = False
               If Info Then
                  LingkaranKecil Baris, Kolom, RGB(255, 0, 0)
                  Else
                  Exit Function
                  End If
               End If
         Next I
         End If
   Next Kolom, Baris
End Function

' Tampil Pesan Giliran MouseDown
Private Sub PesanGiliran_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   BoardAktif.Visible = False
End Sub

' Tampil Pesan Giliran MouseUp
Private Sub PesanGiliran_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   BoardAktif.Visible = True
End Sub

' Kondisi MouseUp untuk Board Aktif
Private Sub BoardAktif_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If akunungguinkamu = True Then
      kolomku = Int(X / Petak)
      barisku = Int(Y / Petak)
      akunungguinkamu = False
      End If
End Sub

' Lingkaran Kecil
Private Sub LingkaranKecil(Baris, Kolom, Warna As Long)
   BoardAktif.Circle (Kolom * Petak + Petak / 2, Baris * Petak + Petak / 2), Petak / 2 - 2, Warna
End Sub

' Batas Board
Private Function Batas(Baris, Kolom) As Boolean
   Batas = IIf((Baris = 1 Or Baris = 8 Or Kolom = 1 Or Kolom = 8), True, False)
End Function

' Sabar Nungguin
Private Sub SabarNunggu(sn)
   Dim waktu As Variant
   waktu = Timer
   While Timer - waktu < sn / 1000: DoEvents: Wend
End Sub

' Giliran Bermain
Private Function GiliranMain() As Boolean
   Giliran = Giliran + 1
   If Giliran = 60 Then
      KondisiMenang "Permainan Berakhir."
      Reset
      GiliranMain = False
      Else
      GiliranMain = True
      End If
End Function

' Reset Game
Private Sub Reset()
   PilihWarna.Caption = "Pilih Warna"
End Sub

' Keluar Batas
Private Function KeluarBatas(Baris, Kolom) As Boolean
   KeluarBatas = IIf((Baris < 1 Or Baris > 8 Or Kolom < 1 Or Kolom > 8), True, False)
End Function

' Luar Board
Private Function LuarBoard(Baris, Kolom) As Boolean
   LuarBoard = IIf((Baris < 1 Or Baris > 8 Or Kolom < 1 Or Kolom > 8), True, False)
End Function

' Nunggu Musuh
Private Sub NungguMusuh()
   BoardCom.Picture = BoardAktif.Image
End Sub

'Nunggu Yang Main
Private Sub Nungguin()
Nunggulagi:
   barisku = 0: kolomku = 0
   akunungguinkamu = True
   While akunungguinkamu = True And Udahan = False: DoEvents: Wend
   If Udahan = True Then Exit Sub
   
   'Tampil Pesan karena memilih bukan pada petak
   If LuarBoard(barisku, kolomku) Then MsgBox "Kamu gak milih petak!": GoTo Nunggulagi:
End Sub

' Pesan
Private Sub Pesan(txt As String, Warna As Long)
   PesanGiliran(0).ForeColor = Warna
   PesanGiliran(0).Caption = txt
   PesanGiliran(1).Caption = txt
End Sub
'==================================BAGIANCODINGUTAMA2===================================

'==================================BAGIANCODINGMAIN===================================
Private Sub Mainkan()
   Dim Max
   Dim Untung
   Dim Lewat As Boolean
   Dim RandomIndex
   
   SetKonstanta
   TampilField

   If warnaku = petakhitam Then GoTo kamu:
   
aku:
   Pesan "Giliran Player", Hitam
   Do
      Nungguin
      If Udahan = True Then Exit Sub
      If C(barisku, kolomku) <> petakhijau Then Pesan "Petak penuh!", Putih: SabarNunggu 1000: Pesan "Giliran si Hitam !", Hitam
   Loop Until C(barisku, kolomku) = petakhijau
   
   TandaSilang barisku, kolomku
   Untung = DapatPetak(barisku, kolomku, warnamu)
   
   If Untung = 0 Then Pesan "Salah petak!", Putih: SabarNunggu 1000: TampilField: GoTo aku:
        C(barisku, kolomku) = warnamu
        NungguMusuh
        TampilField
   If GiliranMain() = False Then Exit Sub
   
kamu:
   Pesan "Giliran Musuh", Hitam
   UpdatePrioritas Max
   CariPetakTerbaik Max, Untung
   
   ' Kondisi menemukan petak terbaik
   If Max > 0 Then
      If Jumlah > 0 Then
         Jumlah = Jumlah + 1
         RandomIndex = Int(Rnd(1) * Jumlah)
         kolommu = KolomMusuh(RandomIndex): barismu = BarisMusuh(RandomIndex)
      Else
         kolommu = KolomMusuh(0): barismu = BarisMusuh(0)
      End If
      
      TandaSilang barismu, kolommu
      Untung = DapatPetak(barismu, kolommu, warnaku)
      C(barismu, kolommu) = warnaku
      NungguMusuh
      TampilField
      
      If GiliranMain() = False Then Exit Sub
      Lewat = Dilewat()
      If Lewat = True Then Pesan "Dilewat aja", Putih: SabarNunggu 2000: GoTo kamu:
      GoTo aku:
      End If
      
   ' Kondisi tidak menemukan satupun
   If Lewat = True Then
      KondisiMenang "Permainan Berakhir!"
      Reset
      Exit Sub
   End If
   Pesan "Lewat!", Putih: SabarNunggu 2000
   
   ' Kondisi Masih bisa main
   Lewat = Dilewat()
   If Lewat = True Then
      KondisiMenang "Permainan Berakhir!"
      Reset
      Exit Sub
   End If
   GoTo aku:
   
End Sub

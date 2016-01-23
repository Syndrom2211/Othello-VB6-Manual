VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Depan 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Othello Classic"
   ClientHeight    =   4650
   ClientLeft      =   3735
   ClientTop       =   3435
   ClientWidth     =   4665
   FillColor       =   &H0080FF80&
   Icon            =   "Depan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Depan.frx":014A
   MousePointer    =   99  'Custom
   Picture         =   "Depan.frx":029C
   ScaleHeight     =   4650
   ScaleWidth      =   4665
   Begin VB.CommandButton CaraMain 
      BackColor       =   &H80000007&
      Height          =   375
      Index           =   1
      Left            =   2520
      MaskColor       =   &H80000001&
      Picture         =   "Depan.frx":57C1
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Left            =   6000
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Left            =   5280
      Top             =   720
   End
   Begin VB.CommandButton Keluar 
      BackColor       =   &H80000001&
      Caption         =   "Keluar"
      Height          =   375
      Left            =   2520
      MaskColor       =   &H80000001&
      Picture         =   "Depan.frx":5C95
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Tentang 
      BackColor       =   &H80000007&
      Height          =   375
      Index           =   0
      Left            =   1080
      MaskColor       =   &H80000001&
      Picture         =   "Depan.frx":610D
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Mulai 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Niagara Engraved"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      MaskColor       =   &H80000001&
      Picture         =   "Depan.frx":65EA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.Image StopM 
      Height          =   240
      Index           =   2
      Left            =   3360
      Picture         =   "Depan.frx":6AE1
      Top             =   2280
      Width           =   255
   End
   Begin VB.Image PauseM 
      Height          =   240
      Index           =   1
      Left            =   3000
      Picture         =   "Depan.frx":9AAE
      Top             =   2280
      Width           =   255
   End
   Begin VB.Image PlayM 
      Height          =   240
      Index           =   0
      Left            =   2640
      Picture         =   "Depan.frx":C8AB
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "P e t a k  V e r s i o n"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Image Info 
      Height          =   240
      Left            =   4320
      MousePointer    =   14  'Arrow and Question
      Picture         =   "Depan.frx":F73C
      Top             =   3960
      Width           =   240
   End
   Begin WMPLibCtl.WindowsMediaPlayer Wmp 
      Height          =   135
      Left            =   6480
      TabIndex        =   0
      Top             =   360
      Width           =   135
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   238
      _cy             =   238
   End
End
Attribute VB_Name = "Depan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==========================Deklarasi Fade & Timer==========================
Option Explicit

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal Color As Long, ByVal X As Byte, ByVal alpha As Long) As Boolean
 
Const LWA_BOTH = 3
Const LWA_ALPHA = 2
Const LWA_COLORKEY = 1
Const GWL_EXSTYLE = -20
Const WS_EX_LAYERED = &H80000
 
Dim TransparanDepan As Integer

Public Sub BuatTrasparan(hWndBro As Long, iTransp As Integer)
    On Error Resume Next
 
    Dim ret As Long
    ret = GetWindowLong(hWndBro, GWL_EXSTYLE)
 
    SetWindowLong hWndBro, GWL_EXSTYLE, ret Or WS_EX_LAYERED
    SetLayeredWindowAttributes hWndBro, RGB(255, 255, 0), iTransp, LWA_ALPHA
    Exit Sub
End Sub

Private Sub CaraMain_Click(Index As Integer)
MsgBox "Cara Main : " & vbCrLf & vbCrLf & "1. Permainan dimainkan menggunakan mouse" & vbCrLf & vbCrLf & "2. Tekan tombol Pilih warna untuk memilih warna petak" & vbCrLf & vbCrLf & "3. Untuk menghentikan permainan, tekan tombol Udahan" & vbCrLf & vbCrLf & "3. Ketika giliran kamu untuk bermain tiba, kamu dapat memilih posisi tertentu dengan klik kiri mouse " & vbCrLf & vbCrLf & Chr(169) & " Othello Classic", vbInformation, "Cara Bermain"
End Sub

Private Sub Info_Click()
MsgBox "Info : " & vbCrLf & vbCrLf & "Kelompok 12 - Analisis Algoritma 8" & vbCrLf & "DFS dan Minimax" & vbCrLf & vbCrLf & "1. Firdamdam.Sasmita (10114175)" & vbCrLf & "2. Fajar (10114495)" & vbCrLf & "3. GunGun Abdullah (10114197)" & vbCrLf & vbCrLf & "Universitas Komputer Indonesia" & vbCrLf & Chr(169) & " Othello Classic", vbInformation, "Info"
End Sub

Private Sub Mulai_Click()
Main.Show
End Sub

Private Sub Tentang_Click(Index As Integer)
MsgBox "Tentang Game : " & vbCrLf & vbCrLf & "Othello adalah permainan tradisional berbentuk papan untuk dimainkan oleh 2 pemain, yaitu hitam dan putih. Salah satu aturan dalam game ini pada awal permainan kamu akan diminta memilih warna hitam atau putih. Kondisi ketika salah seorang pemain menang, adalah dengan banyak nya kepingan yang dimilikinya ketika semua papan sudah terpenuhi, siapa yang banyak maka dia yang menang. Menaklukan pemain hanya dilakukan ketika lawan sudah terhimpit oleh kedua warna yang menyerang." & vbCrLf & vbCrLf & Chr(169) & " Othello Classic", vbInformation, "Tentang Game"
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    TransparanDepan = TransparanDepan + 5
    If TransparanDepan > 255 Then
        TransparanDepan = 255
        Timer1.Enabled = False
    End If
      BuatTrasparan Me.hWnd, TransparanDepan
    Me.Show
End Sub

Private Sub Timer2_Timer()
    On Error Resume Next
    TransparanDepan = TransparanDepan - 5
    If TransparanDepan < 0 Then
        TransparanDepan = 0
        Timer2.Enabled = False
        End
    End If
    BuatTrasparan Me.hWnd, TransparanDepan
End Sub

Private Sub Form_Load()
'==========================Timer Load Form==========================
Timer1.Enabled = False
Timer1.Interval = 1

Timer2.Enabled = False
Timer2.Interval = 1

Me.Visible = False
Timer1.Enabled = True
'==========================MUSIC==========================
Dim strBuff As String
Dim strFile As String
    'Membuat nama temp file
    strFile = App.Path & "\Othello.mp3"
    
    'Extrak File dari Resource File
    strBuff = StrConv(LoadResData(102, "CUSTOM"), vbUnicode)
    
    'Menghapus attribut Read-Only sebelum membuka file untuk output
    If Len(Dir(strFile, vbHidden)) > 0 Then SetAttr strFile, vbNormal
    
    'Save the string as a file
    Open strFile For Output As #1
        Print #1, strBuff
    Close #1
    
    'Menempatkan atrribut lagi setelah menutupnya
    SetAttr strFile, vbArchive + vbHidden
    
    Wmp.URL = App.Path & "\Othello.mp3" 'Load a Music
    Wmp.Controls.play 'Mainkan
'==========================MUSIC==========================
End Sub

Private Sub PlayM_Click(Index As Integer)
Wmp.Controls.play
End Sub

Private Sub StopM_Click(Index As Integer)
Wmp.Controls.stop
End Sub

Private Sub PauseM_Click(Index As Integer)
Wmp.Controls.pause
End Sub

Private Sub Form_Unload(Cancel As Integer)
'==========================Timer UnLoad Form==========================
    Cancel = 1
    Timer1.Enabled = False
    Timer2.Enabled = True
End Sub

Private Sub Keluar_Click()
'==========================Keluar==========================
Dim Keluar As String
Keluar = MsgBox("Keluar dari Games ?", vbExclamation + vbYesNo, Chr(169) & " Othello Classic")
If Keluar = vbYes Then
    Unload Me
End If
End Sub

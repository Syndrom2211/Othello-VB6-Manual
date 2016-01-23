VERSION 5.00
Begin VB.Form Dialog 
   BackColor       =   &H80000008&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Othello Classic"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ForeColor       =   &H00000000&
   Icon            =   "DialogPilih.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "DialogPilih.frx":014A
   MousePointer    =   99  'Custom
   ScaleHeight     =   2790
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton DialogKembali 
      BackColor       =   &H8000000B&
      Caption         =   "Kembali"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton DialogOk 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.OptionButton Pilihan 
      BackColor       =   &H80000012&
      Caption         =   "Putih"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.OptionButton Pilihan 
      BackColor       =   &H00000000&
      Caption         =   "Hitam"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   1920
      Left            =   0
      Picture         =   "DialogPilih.frx":029C
      Top             =   960
      Width           =   4680
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "M a u  p i l i h  w a r n a  a p a  ?"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Pilihanmu As Integer
Public TombolOK As Boolean

Private Sub DialogKembali_Click()
    TombolOK = False: Hide
End Sub

Private Sub DialogOk_Click()
    TombolOK = True: Hide
End Sub

Private Sub Form_Load()
    Pilihanmu = 0
End Sub

Private Sub Pilihan_Click(Index As Integer)
    Pilihanmu = Index
End Sub

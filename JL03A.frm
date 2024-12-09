VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form JL03A 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CETAK NOTA PENJUALAN"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   7035
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton TblSave 
      Caption         =   "&FAKTUR JUAL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   142
      TabIndex        =   0
      Top             =   990
      Width           =   1635
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&CLOSE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5512
      TabIndex        =   1
      Top             =   990
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   1410
      Left            =   -90
      ScaleHeight     =   1350
      ScaleWidth      =   7230
      TabIndex        =   2
      Top             =   810
      Width           =   7290
   End
   Begin Crystal.CrystalReport CRPT 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3577
      TabIndex        =   4
      Top             =   225
      Width           =   1320
   End
   Begin VB.Label Label2 
      Caption         =   "Nomor Transaksi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2137
      TabIndex        =   3
      Top             =   225
      Width           =   1365
   End
End
Attribute VB_Name = "JL03A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command4_Click()
Unload Me
JL03.Show 1
End Sub

Private Sub TblSave_Click()
    CRPT.ReportFileName = App.Path & "\Report\JL03A.rpt"
    CRPT.SelectionFormula = "{JL03_NOTA.NO_TRANS} = '" + Trim(Label3) + "'"
    CRPT.WindowState = crptMaximized
    CRPT.WindowTitle = "Report"
    'CRPT.WindowMaxButton = False
    'CRPT.WindowMinButton = False
    CRPT.Action = 1
    CRPT.Reset
End Sub

Private Sub Command1_Click()
    CRPT.ReportFileName = App.Path & "\Report\JL03B.rpt"
    CRPT.SelectionFormula = "{JL03_NOTA.NO_TRANS} = '" + Trim(Label3) + "'"
    CRPT.WindowState = crptMaximized
    CRPT.WindowTitle = "Report"
    'CRPT.WindowMaxButton = False
    'CRPT.WindowMinButton = False
    CRPT.Action = 1
    CRPT.Reset
End Sub

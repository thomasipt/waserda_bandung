VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#13.2#0"; "Codejock.SkinFramework.v13.2.1.ocx"
Begin VB.Form MAIN_MENU 
   BackColor       =   &H00404040&
   Caption         =   "MAIN MENU"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   585
   ClientWidth     =   13155
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MAIN_MENU.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9150
   ScaleWidth      =   13155
   Begin WASERDA.net_Resize net_Resize1 
      Left            =   840
      Top             =   1410
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   8730
      Width           =   13155
      _ExtentX        =   23204
      _ExtentY        =   741
      SimpleText      =   ""
      ShowTips        =   0   'False
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   7541
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   7541
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7541
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework 
      Left            =   0
      Top             =   0
      _Version        =   851970
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Image Image1 
      Height          =   8745
      Left            =   0
      Picture         =   "MAIN_MENU.frx":0A02
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13200
   End
   Begin VB.Menu TBS1 
      Caption         =   "TABEL SYSTEM"
      Index           =   10
      NegotiatePosition=   2  'Middle
      Begin VB.Menu TBS 
         Caption         =   "KODE BARANG"
         Index           =   11
         Shortcut        =   {F1}
      End
      Begin VB.Menu TBS 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu TBS 
         Caption         =   "DATA PETANI"
         Index           =   13
         Shortcut        =   {F2}
      End
      Begin VB.Menu TBS 
         Caption         =   "-"
         Index           =   14
      End
      Begin VB.Menu TBS 
         Caption         =   "DATABASE"
         Index           =   15
      End
   End
   Begin VB.Menu TRJ 
      Caption         =   "PENJUALAN"
      Index           =   40
      Begin VB.Menu TJ 
         Caption         =   "BARANG"
         Index           =   41
         Shortcut        =   {F3}
      End
      Begin VB.Menu TJ 
         Caption         =   "-"
         Index           =   43
         Visible         =   0   'False
      End
      Begin VB.Menu TJ 
         Caption         =   "RETUR PENJUALAN"
         Index           =   44
         Visible         =   0   'False
      End
   End
   Begin VB.Menu LPR1 
      Caption         =   "LAPORAN"
      Index           =   80
      Begin VB.Menu LPR 
         Caption         =   "PENJUALAN PERTANGGAL"
         Index           =   81
      End
      Begin VB.Menu LPR 
         Caption         =   "PENJUALAN PERBARANG"
         Index           =   82
      End
      Begin VB.Menu LPR 
         Caption         =   "-"
         Index           =   83
      End
      Begin VB.Menu LPR 
         Caption         =   "PENJUALAN PETANI"
         Index           =   84
      End
   End
   Begin VB.Menu E 
      Caption         =   "EXIT"
      Index           =   400
   End
End
Attribute VB_Name = "MAIN_MENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub E_Click(Index As Integer)
End
End Sub

Private Sub Form_Activate()
Me.WindowState = 2
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub Form_Load()
SkinFramework.LoadSkin App.Path + "\Vista.cjstyles", ""
SkinFramework.ApplyWindow Me.hWnd
SkinFramework.ApplyOptions = SkinFramework.ApplyOptions Or xtpSkinApplyMetrics

Me.Top = 0
Me.Left = 0

With StatusBar1.Panels
    .Item(1).Style = sbrText
    .Item(1).Text = "USER : " & Operator
    
    .Item(2).Style = sbrText
    .Item(2).Text = "TANGGAL SISTEM : " & Tanggal
    
    .Item(3).Text = "Copyrighted® edp_ipt2014"
    .Item(3).Style = sbrText
End With

Me.Caption = ">>" & Me.Caption & "               >>Ver." & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub INF_Click(Index As Integer)
Select Case Index
    Case 71
        V_C012.Show 1
End Select
End Sub

Private Sub LPR_Click(Index As Integer)
Select Case Index
    Case 81
        StsCetak = 1
        REPORT.Show 1
    Case 82
        StsCetak = 2
        REPORT.Show 1
    Case 84
        StsCetak = 3
        REPORT.Show 1
End Select
End Sub

Private Sub TBS_Click(Index As Integer)
Select Case Index
    Case 11
        B003.Show 1 'TABEL KODE BARANG
    Case 13
        C012.Show 1 'TABEL KODE PETANI
    Case 15
        DATABASE.Show 1
End Select
End Sub

Private Sub TJ_Click(Index As Integer)
Select Case Index
    Case 41
        JL03.Show 1  'TRANSAKSI PENJUALAN TUNAI
    Case 44
        JL04.Show 1  'RETUR PENJUALAN
End Select
End Sub

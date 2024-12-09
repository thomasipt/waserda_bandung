VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form REPORT 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "LAPORAN"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4845
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbFilter 
      Height          =   315
      Left            =   690
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   975
      Width           =   3500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&PRINT"
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
      Left            =   352
      TabIndex        =   5
      Top             =   1890
      Width           =   1380
   End
   Begin VB.CommandButton Command2 
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
      Left            =   3112
      TabIndex        =   4
      Top             =   1890
      Width           =   1380
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   2625
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   150
      Width           =   1410
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   810
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   150
      Width           =   1410
   End
   Begin Crystal.CrystalReport CRPT 
      Left            =   105
      Top             =   2745
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   1365
      Left            =   -885
      ScaleHeight     =   1305
      ScaleWidth      =   6660
      TabIndex        =   0
      Top             =   1770
      Width           =   6720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "s/d"
      Height          =   195
      Left            =   2302
      TabIndex        =   3
      Top             =   195
      Width           =   240
   End
End
Attribute VB_Name = "REPORT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RCari As New ADODB.Recordset
Private SCari As String

Private Sub Command1_Click()
D1 = Day(Text1)
M1 = Month(Text1)
Y1 = Year(Text1)

D2 = Day(Text2)
M2 = Month(Text2)
Y2 = Year(Text2)

If StsCetak = 1 Then
    CRPT.ReportFileName = App.Path & "\Report\B005_HIS_1.rpt"
    CRPT.SelectionFormula = "{B005_HIS_1.TGL_TRANS} in date (" & Y1 & "," & M1 & "," & D1 & ") to date (" & Y2 & "," & M2 & "," & D2 & ")"
    CRPT.WindowState = crptMaximized
    CRPT.WindowTitle = "LAPORAN PENJUALAN PERTANGGAL"
    CRPT.Action = 1
    CRPT.Reset
ElseIf StsCetak = 2 Then
    CRPT.ReportFileName = App.Path & "\Report\B003_HIS_1.rpt"
    If cmbFilter.Text = "ALL" Then
        CRPT.SelectionFormula = "{B003_HIS_1.TGL_TRANS} in date (" & Y1 & "," & M1 & "," & D1 & ") to date (" & Y2 & "," & M2 & "," & D2 & ")"
    Else
        CRPT.SelectionFormula = "{B003_HIS_1.TGL_TRANS} in date (" & Y1 & "," & M1 & "," & D1 & ") to date (" & Y2 & "," & M2 & "," & D2 & ") AND {B003_HIS_1.NAMA_JNS} = '" + Trim(cmbFilter) + "'"
    End If
    CRPT.WindowState = crptMaximized
    CRPT.WindowTitle = "LAPORAN PENJUALAN PERBARANG"
    CRPT.Action = 1
    CRPT.Reset
ElseIf StsCetak = 3 Then
    CRPT.ReportFileName = App.Path & "\Report\C012_HIS_1.rpt"
    If cmbFilter.Text = "ALL" Then
        CRPT.SelectionFormula = "{C012_HIS_1.TGL_TRANS} in date (" & Y1 & "," & M1 & "," & D1 & ") to date (" & Y2 & "," & M2 & "," & D2 & ")"
    Else
        CRPT.SelectionFormula = "{C012_HIS_1.TGL_TRANS} in date (" & Y1 & "," & M1 & "," & D1 & ") to date (" & Y2 & "," & M2 & "," & D2 & ") AND {C012_HIS_1.Nama} = '" + Trim(cmbFilter) + "'"
    End If
    CRPT.WindowState = crptMaximized
    CRPT.WindowTitle = "LAPORAN PETANI"
    CRPT.Action = 1
    CRPT.Reset
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text1 = Format(DateSerial(Year(Tanggal), Month(Tanggal), 1), "DD/MM/YYYY")
Text2 = Tanggal

cmbFilter.Clear
If StsCetak = 1 Then
    cmbFilter.Visible = False
ElseIf StsCetak = 2 Then
    cmbFilter.AddItem "ALL"
    SCari = "Select * From B003 order by NAMA_JNS"
    Set RCari = New ADODB.Recordset
    RCari.Open SCari, CN, adOpenKeyset
    If RCari.RecordCount <> 0 Then
        RCari.MoveFirst
        Do While Not RCari.EOF
            cmbFilter.AddItem RCari("NAMA_JNS")
        RCari.MoveNext
        Loop
    End If
    RCari.Close
    Set RCari = Nothing
    cmbFilter.ListIndex = 0
ElseIf StsCetak = 3 Then
    cmbFilter.AddItem "ALL"
    SCari = "Select * From C012 order by Nama"
    Set RCari = New ADODB.Recordset
    RCari.Open SCari, CN, adOpenKeyset
    If RCari.RecordCount <> 0 Then
        RCari.MoveFirst
        Do While Not RCari.EOF
            cmbFilter.AddItem RCari("Nama")
        RCari.MoveNext
        Loop
    End If
    RCari.Close
    Set RCari = Nothing
    cmbFilter.ListIndex = 0
End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text1_LostFocus()
If Text1 = "" Then
    Text1.SetFocus
    MsgBox "TANGGAL TRANSAKSI TIDAK BOLEH KOSONG", vbCritical, "DATA MASIH KOSONG"
    Exit Sub
End If
If Not IsDate(Text1) Then
    Text1.SetFocus
    MsgBox "TANGGAL TRANSAKSI HARUS TYPE DD/MM/YYYY", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text2_LostFocus()
If Text2 = "" Then
    Text2.SetFocus
    MsgBox "TANGGAL TRANSAKSI TIDAK BOLEH KOSONG", vbCritical, "DATA MASIH KOSONG"
    Exit Sub
End If
If Not IsDate(Text2) Then
    Text2.SetFocus
    MsgBox "TANGGAL TRANSAKSI HARUS TYPE DD/MM/YYYY", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
End Sub

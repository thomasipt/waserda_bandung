VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form C012 
   Caption         =   "ENTRI DATA PETANI"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "C012.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPDPT 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   2085
      MaxLength       =   35
      TabIndex        =   9
      Text            =   "txtPDPT"
      Top             =   3735
      Width           =   1830
   End
   Begin VB.TextBox txtPRODUKSI 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   2085
      MaxLength       =   35
      TabIndex        =   8
      Text            =   "txtPRODUKSI"
      Top             =   3270
      Width           =   1830
   End
   Begin VB.TextBox txtPERSEN 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   2085
      MaxLength       =   35
      TabIndex        =   7
      Text            =   "txtPERSEN"
      Top             =   2835
      Width           =   1830
   End
   Begin VB.TextBox txtDaerah 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   360
      Left            =   8220
      MaxLength       =   50
      TabIndex        =   29
      Text            =   "txtDaerah"
      Top             =   1050
      Width           =   2500
   End
   Begin VB.TextBox txtJML 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   2085
      MaxLength       =   35
      TabIndex        =   6
      Text            =   "txtJML"
      Top             =   2385
      Width           =   1830
   End
   Begin VB.TextBox txtHA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   4275
      MaxLength       =   35
      TabIndex        =   5
      Text            =   "txtHA"
      Top             =   1935
      Width           =   1830
   End
   Begin VB.ComboBox cmbKel 
      Height          =   360
      Left            =   2085
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1050
      Width           =   3500
   End
   Begin VB.CommandButton TmbDel 
      Caption         =   "&DELETE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2933
      TabIndex        =   12
      Top             =   6795
      Width           =   1260
   End
   Begin VB.Frame Frame4 
      Caption         =   "Pencarian :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   750
      Left            =   6960
      TabIndex        =   20
      Top             =   3345
      Width           =   4980
      Begin VB.ComboBox cmbCari 
         Height          =   360
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   300
         Width           =   1320
      End
      Begin VB.TextBox txtCari 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1455
         TabIndex        =   22
         Text            =   "txtCari"
         Top             =   300
         Width           =   2550
      End
      Begin VB.CommandButton TblCari 
         Caption         =   "&Cari"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4080
         TabIndex        =   23
         Top             =   300
         Width           =   795
      End
   End
   Begin VB.TextBox txtNo 
      Height          =   360
      Left            =   2085
      TabIndex        =   0
      Text            =   "txtNo"
      Top             =   120
      Width           =   3500
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&PRINT"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5363
      TabIndex        =   13
      Top             =   6795
      Width           =   1380
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&CLOSE"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   10028
      TabIndex        =   11
      Top             =   6795
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&SAVE"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   473
      TabIndex        =   10
      Top             =   6795
      Width           =   1380
   End
   Begin VB.TextBox txtPatok 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   2085
      MaxLength       =   35
      TabIndex        =   4
      Text            =   "txtPatok"
      Top             =   1935
      Width           =   1830
   End
   Begin VB.TextBox txtM2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   2085
      MaxLength       =   35
      TabIndex        =   3
      Text            =   "txtM2"
      Top             =   1515
      Width           =   1830
   End
   Begin VB.TextBox txtKetua 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   360
      Left            =   5640
      MaxLength       =   50
      TabIndex        =   14
      Text            =   "txtKetua"
      Top             =   1050
      Width           =   2500
   End
   Begin VB.TextBox txtNama 
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   2085
      MaxLength       =   35
      TabIndex        =   1
      Text            =   "txtNama"
      Top             =   600
      Width           =   6060
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   1185
      Left            =   -60
      ScaleHeight     =   1125
      ScaleWidth      =   11940
      TabIndex        =   19
      Top             =   6645
      Width           =   12000
   End
   Begin Crystal.CrystalReport CRPT 
      Left            =   5730
      Top             =   3532
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
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   2385
      Left            =   90
      TabIndex        =   24
      Top             =   4170
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   4207
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   -2147483633
      WordWrap        =   -1  'True
      Redraw          =   -1  'True
      AllowUserResizing=   3
      Appearance      =   0
   End
   Begin VB.Label Label12 
      Caption         =   "PENDAPATAN PERBLN"
      Height          =   330
      Left            =   135
      TabIndex        =   32
      Top             =   3735
      Width           =   1935
   End
   Begin VB.Label Label11 
      Caption         =   "ESTIMASI PRODUKSI"
      Height          =   330
      Left            =   135
      TabIndex        =   31
      Top             =   3270
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "PROSENTASE"
      Height          =   330
      Left            =   135
      TabIndex        =   30
      Top             =   2835
      Width           =   1935
   End
   Begin VB.Label Label10 
      Caption         =   "JUMLAH POPULASI"
      Height          =   330
      Left            =   135
      TabIndex        =   28
      Top             =   2385
      Width           =   1935
   End
   Begin VB.Label Label9 
      Caption         =   "Ha"
      Height          =   330
      Left            =   4020
      TabIndex        =   27
      Top             =   1935
      Width           =   240
   End
   Begin VB.Label Label8 
      Caption         =   "Patok"
      Height          =   330
      Left            =   1650
      TabIndex        =   26
      Top             =   1935
      Width           =   420
   End
   Begin VB.Label Label7 
      Caption         =   "KELOMPOK"
      Height          =   330
      Left            =   135
      TabIndex        =   25
      Top             =   1065
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "LUAS AREAL"
      Height          =   330
      Left            =   135
      TabIndex        =   18
      Top             =   1935
      Width           =   1125
   End
   Begin VB.Label Label5 
      Caption         =   "LUAS AREAL SPPT M2"
      Height          =   330
      Left            =   135
      TabIndex        =   17
      Top             =   1515
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "N A M A"
      Height          =   330
      Left            =   135
      TabIndex        =   16
      Top             =   615
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "NO REG"
      Height          =   330
      Left            =   135
      TabIndex        =   15
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "C012"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RCari As New ADODB.Recordset
Private RSimpan As New ADODB.Recordset
Private RKode As New ADODB.Recordset
Private RDel As New ADODB.Recordset

Private SCari, SSimpan, SKode, SDel As String
Private Pesan As Boolean
Private h() As String

Private NoEdit As String

Private Sub cmbKel_Click()
SCari = "Select * from C011 where KELOMPOK = '" + Trim(cmbKel.Text) + "'"
Set RCari = New ADODB.Recordset
RCari.Open SCari, CN, adOpenKeyset
If RCari.RecordCount <> 0 Then
    txtKetua = RCari("KETUA")
    txtDaerah = RCari("DAERAH")
Else
    MsgBox "KODE KELOMPOK", vbInformation, "KODE BLM TERDAFTAR"
    cmbKel.SetFocus
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub cmbKel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Command1_Click()
If txtNo = "" Or txtNama = "" Or cmbKel = "" Or txtM2 = "" Or txtPatok = "" Or txtHA = "" Or txtJML = "" Or txtPERSEN = "" Or txtPRODUKSI = "" Or txtPDPT = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "DATA TIDAK BOLEH KOSONG"
    Exit Sub
End If

SKode = "Select * from C012 where NoNas = '" + Trim(txtNo) + "'"
Set RKode = New ADODB.Recordset
RKode.Open SKode, CN, adOpenKeyset
If RKode.RecordCount <> 0 Then
    MsgBox "KODE PETANI SUDAH TERDAFTAR", vbCritical, "WARNING"
    Exit Sub
End If
RKode.Close
Set RKode = Nothing

If StatusEdit = 1 Then
    SEdit = "Select * From C012 where NoNas = '" + Trim(NoEdit) + "'"
    Set REdit = New ADODB.Recordset
    REdit.Open SEdit, CN, adOpenDynamic, adLockOptimistic
    If REdit.RecordCount <> 0 Then
        REdit("NoNas") = txtNo
        REdit("Nama") = txtNama
        REdit("Alamat1") = txtDaerah
        REdit("M2") = txtM2
        REdit("Patok") = txtPatok
        REdit("HA") = txtHA
        REdit("Hektar") = txtJML
        REdit("Persen") = txtPERSEN
        REdit("Produksi") = txtPRODUKSI
        REdit("Pendapatan") = txtPDPT
        REdit("Sisa") = txtPDPT / 2
        REdit("KetuaKel") = txtKetua
        REdit("NamaKel") = cmbKel
        REdit.Update
    End If
    REdit.Close
    Set REdit = Nothing
    
    SDel = "UPDATE B005 SET KODE_SPL = '" + Trim(txtNo) + "' WHERE KODE_SPL = '" + Trim(NoEdit) + "'"
    Set RDel = New ADODB.Recordset
    RDel.Open SDel, CN, adOpenKeyset
ElseIf StatusEdit = 0 Then
    SSimpan = "Select * From C012"
    Set RSimpan = New ADODB.Recordset
    RSimpan.Open SSimpan, CN, adOpenKeyset, adLockBatchOptimistic
    RSimpan.AddNew
        RSimpan("NoNas") = txtNo
        RSimpan("Nama") = txtNama
        RSimpan("Alamat1") = txtDaerah
        RSimpan("M2") = txtM2
        RSimpan("Patok") = txtPatok
        RSimpan("HA") = txtHA
        RSimpan("Hektar") = txtJML
        RSimpan("Persen") = txtPERSEN
        RSimpan("Produksi") = txtPRODUKSI
        RSimpan("Pendapatan") = txtPDPT
        RSimpan("Sisa") = txtPDPT / 2
        RSimpan("KetuaKel") = txtKetua
        RSimpan("NamaKel") = cmbKel
    RSimpan.UpdateBatch adAffectAllChapters
    RSimpan.Close
    Set RSimpan = Nothing
End If

Pesan = True
Unload Me
Me.Show 1
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
CRPT.ReportFileName = App.Path & "\Report\C012.rpt"
CRPT.WindowState = crptMaximized
CRPT.WindowTitle = "Klik NO REG untuk melihat history perbulan"
'CRPT.WindowMaxButton = False
'CRPT.WindowMinButton = False
CRPT.WindowAllowDrillDown = True
CRPT.Action = 1
CRPT.Reset
End Sub

Private Sub Form_Load()
ClearTextBoxes Me
'Call CariNomor
Call Kosong
Call DaftarKel
Pesan = False

SCari = "Select * From C012 Order By NoNas"
Call Siap
Call IsiGrid

cmbCari.AddItem "Semua"
cmbCari.AddItem "Nama"
cmbCari.AddItem "Kelompok"
cmbCari.ListIndex = 0

StatusEdit = 0
End Sub

Private Sub DaftarKel()
cmbKel.Clear

SCari = "Select * From C011 order by No_Urut"
Set RCari = New ADODB.Recordset
RCari.Open SCari, CN, adOpenKeyset
If RCari.RecordCount <> 0 Then
    RCari.MoveFirst
    Do While Not RCari.EOF
        cmbKel.AddItem RCari("Kelompok")
    RCari.MoveNext
    Loop
End If
RCari.Close
Set RCari = Nothing
cmbKel.ListIndex = 0
End Sub

Private Sub Siap()
With Grid
     .Cols = 11
     .Row = 0
     .RowHeight(0) = 400
     .Col = 0: .ColWidth(0) = 1250: .Text = "NO": .CellAlignment = 4: .CellFontBold = True
     .Col = 1: .ColWidth(1) = 1750: .Text = "NAMA": .CellAlignment = 4: .CellFontBold = True
     .Col = 2: .ColWidth(2) = 1750: .Text = "ALAMAT": .CellAlignment = 4: .CellFontBold = True
     .Col = 3: .ColWidth(3) = 750: .Text = "LUAS": .CellAlignment = 4: .CellFontBold = True
     .Col = 4: .ColWidth(4) = 750: .Text = "PATOK": .CellAlignment = 4: .CellFontBold = True
     .Col = 5: .ColWidth(5) = 750: .Text = "HA": .CellAlignment = 4: .CellFontBold = True
     .Col = 6: .ColWidth(6) = 750: .Text = "PPLS": .CellAlignment = 4: .CellFontBold = True
     .Col = 7: .ColWidth(7) = 500: .Text = "%": .CellAlignment = 4: .CellFontBold = True
     .Col = 8: .ColWidth(8) = 750: .Text = "EST": .CellAlignment = 4: .CellFontBold = True
     .Col = 9: .ColWidth(9) = 1000: .Text = "PDPT": .CellAlignment = 4: .CellFontBold = True
     .Col = 10: .ColWidth(10) = 1250: .Text = "KEL": .CellAlignment = 4: .CellFontBold = True
End With
End Sub

Private Sub IsiGrid()
Set RCari = New ADODB.Recordset
RCari.Open SCari, CN, adOpenKeyset, adLockReadOnly
If RCari.RecordCount <> 0 Then
RCari.MoveFirst
Brs = 1
Do Until RCari.EOF
       With Grid
        .Rows = Brs + 1
        .Row = Brs
        .Col = 0: .Text = RCari("NoNas"): .CellAlignment = 4: .CellFontSize = 8
        .Col = 1: .Text = RCari("Nama"): .CellAlignment = 1: .CellFontSize = 8
        .Col = 2: .Text = RCari("Alamat1"): .CellAlignment = 1: .CellFontSize = 8
        .Col = 3: .Text = RCari("M2"): .CellAlignment = 7: .CellFontSize = 8
        .Col = 4: .Text = RCari("Patok"): .CellAlignment = 7: .CellFontSize = 8
        .Col = 5: .Text = RCari("HA"): .CellAlignment = 7: .CellFontSize = 8
        .Col = 6: .Text = RCari("Hektar"): .CellAlignment = 7: .CellFontSize = 8
        .Col = 7: .Text = RCari("Persen"): .CellAlignment = 4: .CellFontSize = 8
        .Col = 8: .Text = RCari("Produksi"): .CellAlignment = 7: .CellFontSize = 8
        .Col = 9: .Text = RCari("Pendapatan"): .CellAlignment = 7: .CellFontSize = 8
        .Col = 10: .Text = RCari("NamaKel"): .CellAlignment = 1: .CellFontSize = 8
      End With
      RCari.MoveNext
      Brs = Brs + 1
Loop
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub CariNomor()
Dim Nomor As Double
Dim InfoNomor As String
SCari = "Select Top 1 NoNas From C012 order by NoNas Desc"
Set RCari = New ADODB.Recordset
RCari.Open SCari, CN, adOpenKeyset
If RCari.RecordCount <> 0 Then
    Nomor = Val(Right(RCari("NoNas"), 4)) + 1
    'InfoNomor = Trim(RCari("Nonas"))
    'If Pesan = True Then
    'MsgBox "NOMOR TERSIMPAN TERAKHIR" + Trim(InfoNomor), vbOKOnly, "DATA TERSIMPAN"
    'End If
    Label1 = "C." & Digit(4, Nomor)
Else
    Label1 = "C.0001"
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Kosong()
txtM2 = 0
txtPatok = 0
txtHA = 0
txtJML = 0
txtPERSEN = 0
txtPRODUKSI = 0
txtPDPT = 0
End Sub

Private Sub Label1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys vbTab
    Label1 = Format(Label1, ">")
End If
End Sub

Private Sub Label1_LostFocus()
Dim Info
Dim Sales

If Label1 = "" Then Exit Sub

SCari = "Select * from V_C012 where NoNas = '" + Trim(Label1) + "'"
Set RCari = New ADODB.Recordset
RCari.Open SCari, CN, adOpenKeyset
If RCari.RecordCount <> 0 Then
    Info = MsgBox("NAMA CUSTOMER SUDAH TERDAFTAR, EDIT DATA ?     ", vbOKCancel, "WARNING")
    If Info = vbOK Then
        StatusEdit = 1
        Label1.Enabled = False
        Text1 = RCari("Nama")
        Text2 = RCari("Alamat1")
        Text3 = RCari("Kota")
        Text4 = RCari("Telpon")
        Sales = RCari("KodeSLS")
        Text2.SetFocus
    ElseIf Info = vbCancel Then
        StatusEdit = 2
    Else
        Text1.SetFocus
        'ClearTextBoxes Me
    End If
End If
RCari.Close
Set RCari = Nothing

SCari = "Select * From SL001 where Kode ='" + Trim(Sales) + "'"
Set RCari = New ADODB.Recordset
RCari.Open SCari, CN, adOpenKeyset
If RCari.RecordCount <> 0 Then
    cmbKel = RCari("Nama") & " | " & RCari("Kode")
End If
RCari.Close
Set RCari = Nothing

If StatusEdit = 2 Then
    Unload Me
    Me.Show 1
End If
End Sub

Private Sub TblCari_Click()
If cmbCari.Text = "Semua" Then
    txtCari = ""
    SCari = "Select * From C012 order by NoNas Asc"
ElseIf cmbCari.Text = "Nama" Then
    SCari = "Select * From C012 where Nama like '" + Trim(txtCari) + "%' order by NoNas Asc"
ElseIf cmbCari.Text = "Kelompok" Then
    SCari = "Select * From C012 where NamaKel like '" + Trim(txtCari) + "%' order by NoNas Asc"
End If

Call IsiGrid
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys vbTab
    Text1 = Format(Text1, ">")
End If
End Sub

Private Sub Text1_LostFocus()
Dim Info
Dim Sales

If Text1 = "" Then Exit Sub

SCari = "Select * from V_C012 where Nama = '" + Trim(Text1) + "'"
Set RCari = New ADODB.Recordset
RCari.Open SCari, CN, adOpenKeyset
If RCari.RecordCount <> 0 Then
    Info = MsgBox("NAMA CUSTOMER SUDAH TERDAFTAR, EDIT DATA ?     ", vbOKCancel, "WARNING")
    If Info = vbOK Then
        StatusEdit = 1
        Label1.Enabled = False
        Label1 = RCari("NoNas")
        Text1 = RCari("Nama")
        Text2 = RCari("Alamat1")
        Text3 = RCari("Kota")
        Text4 = RCari("Telpon")
        Sales = RCari("KodeSLS")
        Text2.SetFocus
    ElseIf Info = vbCancel Then
        StatusEdit = 2
    Else
        Text1.SetFocus
        'ClearTextBoxes Me
    End If
End If
RCari.Close
Set RCari = Nothing

If Sales = "" Then Exit Sub
SCari = "Select * From SL001 where Kode ='" + Trim(Sales) + "'"
Set RCari = New ADODB.Recordset
RCari.Open SCari, CN, adOpenKeyset
If RCari.RecordCount <> 0 Then
    cmbKel = RCari("Nama") & " | " & RCari("Kode")
End If
RCari.Close
Set RCari = Nothing

If StatusEdit = 2 Then
    Unload Me
    Me.Show 1
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys vbTab
End If
End Sub

Private Sub Text2_LostFocus()
If Text2 = "" Then Text2 = "-"
Text2 = Format(Text2, ">")
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys vbTab
End If
End Sub

Private Sub Text3_LostFocus()
If Text3 = "" Then Text3 = "-"
Text3 = Format(Text3, ">")
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys vbTab
End If
End Sub

Private Sub Text4_LostFocus()
If Text4 = "" Then Text4 = "-"
Text4 = Format(Text4, ">")
End Sub

Private Sub TmbDel_Click()
Dim Tanya
Tanya = MsgBox("YAKIN AKAN HAPUS NAMA PETANI " + Trim(txtNama), vbOKCancel, "YAKIN HAPUS KODE CUSTOMER ?")
If Tanya = vbCancel Then Exit Sub
If txtNo = "" Then Exit Sub
SDel = "Delete From C012 where NoNas = '" + Trim(txtNo) + "'"
Set RDel = New ADODB.Recordset
RDel.Open SDel, CN, adOpenKeyset

Unload Me
Me.Show 1
End Sub

Private Sub txtCari_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub txtCari_LostFocus()
txtCari = Format(txtCari, ">")
End Sub

Private Sub txtHA_GotFocus()
If txtHA = 0 Then txtHA = ""
End Sub

Private Sub txtHA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub txtHA_LostFocus()
If txtHA = "" Then txtHA = 0
If Not IsNumeric(txtHA) Then
    txtHA.SetFocus
    MsgBox "HARUS ANGKA", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
txtHA = Format(txtHA, "##,###.00")
End Sub

Private Sub txtJML_GotFocus()
If txtJML = 0 Then txtJML = ""
End Sub

Private Sub txtJML_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub txtJML_LostFocus()
If txtJML = "" Then txtJML = 0
If Not IsNumeric(txtJML) Then
    txtJML.SetFocus
    MsgBox "HARUS ANGKA", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
txtJML = Format(txtJML, "##,###.00")
End Sub

Private Sub txtM2_GotFocus()
If txtM2 = 0 Then txtM2 = ""
End Sub

Private Sub txtM2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub txtM2_LostFocus()
If txtM2 = "" Then txtM2 = 0
If Not IsNumeric(txtM2) Then
    txtM2.SetFocus
    MsgBox "HARUS ANGKA", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
txtM2 = Format(txtM2, "##,###.00")
End Sub

Private Sub txtNama_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys vbTab
End If
End Sub

Private Sub txtNama_LostFocus()
If txtNama = "" Then txtNama = "-"
txtNama = Format(txtNama, ">")
End Sub

Private Sub txtNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub txtNo_LostFocus()
If txtNo = "" Then Exit Sub

If StatusEdit = 1 Then Exit Sub
SKode = "Select * from C012 where NoNas = '" + Trim(txtNo) + "'"
Set RKode = New ADODB.Recordset
RKode.Open SKode, CN, adOpenKeyset
If RKode.RecordCount <> 0 Then
    Info = MsgBox("KODE PETANI SUDAH TERDAFTAR, AKAN DILAKUKAN EDIT", vbOKCancel, "KODE PETANI SUDAH TERDAFTAR")
    If Info = vbOK Then
        StatusEdit = 1
        NoEdit = RKode("NoNas")
        txtNama = RKode("Nama")
        cmbKel = RKode("NamaKel")
        txtM2 = RKode("M2")
        txtPatok = RKode("Patok")
        txtHA = RKode("HA")
        txtJML = RKode("Hektar")
        txtPERSEN = RKode("Persen")
        txtPRODUKSI = RKode("Produksi")
        txtPDPT = RKode("Pendapatan")
        TmbDel.Enabled = True
        txtNo.SetFocus
    Else
        StatusEdit = 0
        ClearTextBoxes Me
        Call Kosong
        TmbDel.Enabled = False
    End If
End If
RKode.Close
Set RKode = Nothing
Text1 = Format(Text1, ">")
End Sub

Private Sub txtPatok_GotFocus()
If txtPatok = 0 Then txtPatok = ""
End Sub

Private Sub txtPatok_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub txtPatok_LostFocus()
If txtPatok = "" Then txtPatok = 0
If Not IsNumeric(txtPatok) Then
    txtPatok.SetFocus
    MsgBox "HARUS ANGKA", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
txtPatok = Format(txtPatok, "##,###.00")
End Sub

Private Sub txtPDPT_GotFocus()
If txtPDPT = 0 Then txtPDPT = ""
End Sub

Private Sub txtPDPT_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub txtPDPT_LostFocus()
If txtPDPT = "" Then txtPDPT = 0
If Not IsNumeric(txtPDPT) Then
    txtPDPT.SetFocus
    MsgBox "HARUS ANGKA", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
txtPDPT = Format(txtPDPT, "##,###.00")
End Sub

Private Sub txtPERSEN_GotFocus()
If txtPERSEN = 0 Then txtPERSEN = ""
End Sub

Private Sub txtPERSEN_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub txtPERSEN_LostFocus()
If txtPERSEN = "" Then txtPERSEN = 0
If Not IsNumeric(txtPERSEN) Then
    txtPERSEN.SetFocus
    MsgBox "HARUS ANGKA", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
txtPERSEN = Format(txtPERSEN, "##,###.00")
End Sub

Private Sub txtPRODUKSI_GotFocus()
If txtPRODUKSI = 0 Then txtPRODUKSI = ""
End Sub

Private Sub txtPRODUKSI_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub txtPRODUKSI_LostFocus()
If txtPRODUKSI = "" Then txtPRODUKSI = 0
If Not IsNumeric(txtPRODUKSI) Then
    txtPRODUKSI.SetFocus
    MsgBox "HARUS ANGKA", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
txtPRODUKSI = Format(txtPRODUKSI, "##,###.00")
End Sub

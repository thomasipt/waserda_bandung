VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form B003 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ENTRI KODE BARANG"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11850
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton TmbEdit 
      Caption         =   "&EDIT"
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
      Left            =   12420
      TabIndex        =   35
      Top             =   1050
      Width           =   1380
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
      Left            =   6803
      TabIndex        =   31
      Top             =   1110
      Width           =   4980
      Begin VB.ComboBox cmbCari 
         Height          =   360
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   300
         Width           =   1320
      End
      Begin VB.TextBox txtCari 
         BackColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   1455
         TabIndex        =   32
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
         TabIndex        =   33
         Top             =   300
         Width           =   795
      End
   End
   Begin VB.CommandButton TmbSave 
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
      Left            =   289
      TabIndex        =   4
      Top             =   6270
      Width           =   1380
   End
   Begin VB.CommandButton TmbClose 
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
      Left            =   10181
      TabIndex        =   6
      Top             =   6270
      Width           =   1380
   End
   Begin VB.CommandButton TmbDel 
      Caption         =   "&DELETE"
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
      Left            =   2820
      TabIndex        =   5
      Top             =   6270
      Width           =   1260
   End
   Begin VB.CommandButton Command1 
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
      Left            =   5235
      TabIndex        =   7
      Top             =   6270
      Width           =   1380
   End
   Begin VB.CommandButton CtkBarkode 
      Caption         =   "&CETAK"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   10725
      TabIndex        =   16
      Top             =   7530
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   1650
      TabIndex        =   3
      Text            =   "Text6"
      Top             =   1365
      Width           =   2820
   End
   Begin VB.TextBox Text5 
      Height          =   360
      Left            =   5325
      TabIndex        =   11
      Text            =   "Text5"
      Top             =   8790
      Width           =   1860
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   5325
      TabIndex        =   9
      Text            =   "Text4"
      Top             =   8310
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.ComboBox Combo4 
      Height          =   360
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   8310
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.ComboBox Combo2 
      Height          =   360
      Left            =   1650
      TabIndex        =   2
      Text            =   "Combo2"
      Top             =   915
      Width           =   2820
   End
   Begin VB.CheckBox Check3 
      Caption         =   "EVERAGE"
      Height          =   360
      Left            =   6630
      TabIndex        =   13
      Top             =   7725
      Width           =   1095
   End
   Begin VB.CheckBox Check2 
      Caption         =   "LIFO (Last-in Last-out)"
      Height          =   360
      Left            =   4410
      TabIndex        =   12
      Top             =   7725
      Width           =   2040
   End
   Begin VB.CheckBox Check1 
      Caption         =   "FIFO (First-in First-out)"
      Height          =   360
      Left            =   2220
      TabIndex        =   10
      Top             =   7725
      Width           =   2040
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   1650
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   60
      Width           =   2820
   End
   Begin VB.TextBox Text2 
      Height          =   360
      Left            =   1650
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   495
      Width           =   5070
   End
   Begin VB.Frame Frame1 
      Caption         =   "CETAK BARCODE"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3750
      Left            =   11310
      TabIndex        =   26
      Top             =   7335
      Visible         =   0   'False
      Width           =   1905
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   225
         TabIndex        =   15
         Text            =   "Text8"
         Top             =   1845
         Width           =   1455
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   225
         TabIndex        =   14
         Text            =   "Text7"
         Top             =   810
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "JML KOLOM KOSONG"
         Height          =   375
         Left            =   90
         TabIndex        =   28
         Top             =   1530
         Width           =   1725
      End
      Begin VB.Label Label14 
         Caption         =   "JUMLAH COPY"
         Height          =   330
         Left            =   270
         TabIndex        =   27
         Top             =   540
         Width           =   1365
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4170
      Left            =   75
      TabIndex        =   19
      Top             =   1875
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   7355
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   -2147483633
      AllowUserResizing=   3
      Appearance      =   0
   End
   Begin Crystal.CrystalReport CRPT 
      Left            =   10740
      Top             =   -60
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
      Height          =   1020
      Left            =   -135
      ScaleHeight     =   960
      ScaleWidth      =   12675
      TabIndex        =   30
      Top             =   6120
      Width           =   12735
   End
   Begin VB.Label Label19 
      Caption         =   "METODE STOCK"
      Height          =   360
      Left            =   660
      TabIndex        =   29
      Top             =   7725
      Width           =   1500
   End
   Begin VB.Label Label13 
      Caption         =   "HJUAL SATUAN"
      Height          =   360
      Left            =   75
      TabIndex        =   25
      Top             =   1365
      Width           =   1500
   End
   Begin VB.Label Label12 
      Caption         =   "BENANG"
      Height          =   285
      Left            =   3765
      TabIndex        =   24
      Top             =   8835
      Width           =   1500
   End
   Begin VB.Label Label10 
      Caption         =   "%"
      Height          =   285
      Left            =   6045
      TabIndex        =   23
      Top             =   8355
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "NOMINAL"
      Height          =   330
      Left            =   4245
      TabIndex        =   22
      Top             =   8325
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label8 
      Caption         =   "HARGA JUAL"
      Height          =   330
      Left            =   540
      TabIndex        =   21
      Top             =   8325
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label6 
      Caption         =   "SATUAN"
      Height          =   315
      Left            =   75
      TabIndex        =   20
      Top             =   945
      Width           =   1500
   End
   Begin VB.Label Label3 
      Caption         =   "NAMA"
      Height          =   315
      Left            =   75
      TabIndex        =   18
      Top             =   525
      Width           =   1500
   End
   Begin VB.Label Label2 
      Caption         =   "KODE BARANG"
      Height          =   315
      Left            =   75
      TabIndex        =   17
      Top             =   90
      Width           =   1500
   End
End
Attribute VB_Name = "B003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RGol As New ADODB.Recordset
Private RCari As New ADODB.Recordset
Private RKode As New ADODB.Recordset
Private RDel As New ADODB.Recordset
Private RDelBar As New ADODB.Recordset
Private RSim As New ADODB.Recordset
Private RSave As New ADODB.Recordset
Private RDist As New ADODB.Recordset
Private SDelBar, SDist As String
Private SGol, SCari, Metode, SKode, SDel, SSim, SSave As String
Private Brs, MetodLaba, Ganti
Private h() As String

Private Sub Check1_Click()
If Check1.Value = 1 Then
    Check2.Value = 0
    Check2.Enabled = False
    Check3.Value = 0
    Check3.Enabled = False
    Metode = "1"
Else
    Check1.Value = 0
    Check1.Enabled = True
    Check2.Value = 0
    Check2.Enabled = True
    Check3.Value = 0
    Check3.Enabled = True
    Metode = ""
End If
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    Check1.Value = 0
    Check1.Enabled = False
    Check3.Value = 0
    Check3.Enabled = False
    Metode = "2"
Else
    Check1.Value = 0
    Check1.Enabled = True
    Check2.Value = 0
    Check2.Enabled = True
    Check3.Value = 0
    Check3.Enabled = True
    Metode = ""
End If
End Sub

Private Sub Check2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
    Check1.Value = 0
    Check1.Enabled = False
    Check2.Value = 0
    Check2.Enabled = False
    Metode = "3"
    If Mid(Text1, 1, 3) = "002" Then
        Check3.Left = 2385
    Else
        Check3.Left = 6345
    End If
Else
    Check1.Value = 0
    Check1.Enabled = True
    Check2.Value = 0
    Check2.Enabled = True
    Check3.Value = 0
    Check3.Enabled = True
    Metode = ""
    If Mid(Text1, 1, 3) = "002" Then
        Check3.Left = 2385
    Else
        Check3.Left = 6795
    End If
End If
End Sub

Private Sub Check3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Combo1_GotFocus()
    SendKeys "{F4}"
End Sub

Private Sub combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Combo1_LostFocus()
If Combo1 = "" Then Exit Sub
h() = Split(Trim(Combo1), " | ")

SGol = "Select Keterangan from B001 where Kode_ind = '" + Trim(h(1)) + "'"
Set RGol = New ADODB.Recordset
RGol.Open SGol, CN, adOpenKeyset
If RGol.RecordCount <> 0 Then
    Label11 = RGol("Keterangan")
Else
    MsgBox "KODE GOLONGAN BELUM TERDAFTAR", vbInformation, "KODE BLM TERDAFTAR"
    Combo1.SetFocus
End If
RGol.Close
Set RGol = Nothing
End Sub

Private Sub combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Combo3_GotFocus()
SendKeys "{F4}"
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Combo3_LostFocus()
If Combo3 = "" Then Exit Sub
h() = Split(Trim(Combo3), " | ")

SDist = "Select NAMA_DISTB from C007 where KODE_DISTB = '" + Trim(h(1)) + "'"
Set RDist = New ADODB.Recordset
RDist.Open SDist, CN, adOpenKeyset
If RDist.RecordCount <> 0 Then
    Label16 = RDist("NAMA_DISTB")
Else
    MsgBox "KODE DISTRIBUTOR BELUM TERDAFTAR", vbInformation, "KODE BLM TERDAFTAR"
    Combo3.SetFocus
End If
RDist.Close
Set RDist = Nothing
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Combo4_LostFocus()
MetodLaba = Combo4.ListIndex + 1
End Sub

Private Sub Command1_Click()
CRPT.ReportFileName = App.Path & "\Report\B003.rpt"
CRPT.WindowState = crptMaximized
CRPT.WindowTitle = "Report"
'CRPT.WindowMaxButton = False
'CRPT.WindowMinButton = False
CRPT.Action = 1
CRPT.Reset
End Sub

Private Sub CtkBarkode_Click()
Dim BrsKosong, JmlCopy, NomorUrut, Tanya

If Combo1 = "" Or Text1 = "" Or Text2 = "" Or Metode = "" Or Combo2 = "" Or Combo4 = "" Or Text4 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "DATA TIDAK BOLEH KOSONG"
    Exit Sub
End If

SDelBar = "Delete From RP11"
Set RDelBar = New ADODB.Recordset
RDelBar.Open SDelBar, CN, adOpenKeyset

NomorUrut = 1
If Text8 > 0 Then
    For BrsKosong = 1 To Text8
        SSim = "Select * from RP11"
        Set RSim = New ADODB.Recordset
        RSim.Open SSim, CN, adOpenKeyset, adLockBatchOptimistic
        RSim.AddNew
            RSim("Nomor") = NomorUrut
            RSim("Harga") = 0
            RSim.UpdateBatch adAffectAllChapters
        RSim.Close
        Set RSim = Nothing
        NomorUrut = NomorUrut + 1
    Next BrsKosong
End If

For JmlCopy = 1 To Text7
    SSim = "Select * from RP11"
    Set RSim = New ADODB.Recordset
    RSim.Open SSim, CN, adOpenKeyset, adLockBatchOptimistic
    RSim.AddNew
        RSim("Nomor") = NomorUrut
        RSim("Kode_Jns") = Trim("*") + Trim(Text1) + Trim("*")
        RSim("Nama_Jns") = Text2
        RSim("Style") = Text3
        RSim("Warna") = Text5
        RSim("Harga") = Text6
        RSim.UpdateBatch adAffectAllChapters
    RSim.Close
    Set RSim = Nothing
    NomorUrut = NomorUrut + 1
Next JmlCopy

Tanya = MsgBox("AKAN MENCETAK BARCODE KE PRINTER, DAN MENYIMPAN KODE BARANG BARU", vbOKCancel, "SIAPKAN KERTAS BARCODE KE PRINTER")
If Tanya = vbCancel Then Exit Sub

Call TmbSave_Click

CRPT.ReportFileName = App.Path & "\Report\CBARCODE.rpt"
CRPT.WindowState = crptMaximized
CRPT.WindowTitle = "Report"
'CRPT.WindowMaxButton = False
'CRPT.WindowMinButton = False
CRPT.Action = 1
CRPT.Reset
End Sub

Private Sub Form_Load()
ClearTextBoxes Me
Call Kosong
Call Siap

SCari = "Select * From B003 order by NAMA_JNS Asc"
Call IsiGrid

Call IsiCombo2

Combo4.AddItem "PERSENTASE HARGA TERTINGGI HPP  (1)", 0
Combo4.ListIndex = 0
Text3 = ""
Text4 = ""
Check1.Value = 1

cmbCari.AddItem "Semua"
cmbCari.AddItem "Nama"
cmbCari.ListIndex = 0
End Sub

Private Sub IsiCombo2()
SGol = "Select * From B003A order by NO_URUT Asc"
Set RGol = New ADODB.Recordset
RGol.Open SGol, CN, adOpenKeyset
If RGol.RecordCount <> 0 Then
    RGol.MoveFirst
    Do While Not RGol.EOF
        Combo2.AddItem RGol("SATUAN")
    RGol.MoveNext
    Loop
End If
RGol.Close
Set RGol = Nothing
Combo2.ListIndex = 0
End Sub

Private Sub Kosong()
Text1 = ""
Text2 = ""
Text3 = "-"
Text4 = 0
Text5 = "-"
Text6 = 0
Text7 = 1
Text8 = 0
Text9 = ""
Text10 = "TR"
Check1.Value = 1
Check2.Value = 0
Check3.Value = 0
Metode = ""
Label11 = ""
Label16 = ""
TmbDel.Enabled = False
TmbEdit.Enabled = True
Ganti = 0
End Sub

Private Sub TblCari_Click()
If cmbCari.Text = "Semua" Then
    SCari = "Select * From B003 order by KODE_JNS Asc"
ElseIf cmbCari.Text = "Nama" Then
    SCari = "Select * from B003 where NAMA_JNS like '%" + Trim(txtCari) + "%' order by KODE_JNS"
ElseIf cmbCari.Text = "Merk" Then
    SCari = "Select * from B003 where STYLE like '%" + Trim(txtCari) + "%' order by KODE_JNS"
ElseIf cmbCari.Text = "Seri" Then
    SCari = "Select * from B003 where KETERANGAN like '%" + Trim(txtCari) + "%' order by KODE_JNS"
End If

Call IsiGrid
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text1_LostFocus()
Dim Info, Awal, Akhir

If Text1 = "" Then Exit Sub

SCari = "Select *  from B003 where Kode_Jns = '" + Text1 + "'"
Set RCari = New ADODB.Recordset
RCari.Open SCari, CN, adOpenKeyset
If RCari.RecordCount <> 0 Then
    Info = MsgBox("KODE JENIS BARANG SUDAH TERDAFTAR, AKAN DILAKUKAN EDIT", vbOKCancel, "KODE JENIS BARANG SUDAH TERDAFTAR")
    If Info = vbOK Then
        Text2 = RCari("Nama_Jns")
        Combo2 = RCari("satuan")
        Text6 = Format(RCari("HJUAL_PCS"), "##,###.00")
        TmbDel.Enabled = True
    Else
        Call Kosong
    End If
Else
    Text2 = ""
    Text3 = ""
    Text5 = ""
    Metode = ""
    Check1.Value = 1
    Check2.Value = 0
    Check3.Value = 0
End If
RCari.Close
Set RCari = Nothing
Text1 = Format(Text1, ">")
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
        .Col = 0: .Text = RCari("NO_URUT"): .CellAlignment = 4
        .Col = 1: .Text = RCari("KODE_JNS"): .CellAlignment = 1
        .Col = 2: .Text = RCari("NAMA_JNS"): .CellAlignment = 1
        .Col = 3: .Text = RCari("SATUAN"): .CellAlignment = 4
        .Col = 4: .Text = Format(RCari("HJUAL_PCS"), "##,###.00"): .CellAlignment = 7
      End With
      RCari.MoveNext
      Brs = Brs + 1
Loop
End If
RCari.Close
Set RCari = Nothing
'If Brs > 1 Then
'    Grid.TopRow = Brs - 1
'End If
End Sub

Private Sub Siap()
With Grid
     .Cols = 5
     .Row = 0
     .RowHeight(0) = 400
     .Col = 0: .ColWidth(0) = 0
     .Col = 1: .ColWidth(1) = 2000: .Text = "KODE": .CellAlignment = 4: .CellFontBold = True
     .Col = 2: .ColWidth(2) = 6500: .Text = "NAMA": .CellAlignment = 4: .CellFontBold = True
     .Col = 3: .ColWidth(3) = 1000: .Text = "STN": .CellAlignment = 4: .CellFontBold = True
     .Col = 4: .ColWidth(4) = 1500: .Text = "H.JUAL": .CellAlignment = 4: .CellFontBold = True
End With
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys vbTab
    Check1.Value = 1
End If
End Sub

Private Sub Text2_LostFocus()
If Ganti = 1 Then Exit Sub
SCari = "Select Nama_JNS from B003 where Nama_JNS = '" + Text2 + "' and KODE_JNS <> '" + Text1 + "'"
Set RCari = New ADODB.Recordset
RCari.Open SCari, CN, adOpenKeyset
If RCari.RecordCount <> 0 Then
    Text2.SetFocus
    MsgBox "NAMA JENIS BARANG SUDAH DIGUNAKAN", vbInformation, "NAMA JENIS BARANG SUDAH TERDAFTAR"
    Exit Sub
End If
RCari.Close
Set RCari = Nothing
Text2 = Format(Text2, ">")
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text3_LostFocus()
If Text3 = "" Then Text3 = "-"
Text3 = Format(Text3, ">")
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text4_LostFocus()
If Text4 = "" Then Exit Sub
If Not IsNumeric(Text4) Then
    Text4.SetFocus
    MsgBox "NOMINAL PERSENTASE LABA HARUS ANGKA", vbCritical, "TYPE DATA SALAH"
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text5_LostFocus()
Text5 = Format(Text5, ">")
End Sub

Private Sub Text6_GotFocus()
If Text6 = 0 Then Text6 = ""
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text6_LostFocus()
If Text6 = "" Then Text6 = 0
If Not IsNumeric(Text6) Then
    Text6.SetFocus
    MsgBox "HARGA JUAL HARUS ANGKA", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
Text6 = Format(Text6, "##,###.00")
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text7_LostFocus()
If Text7 = "" Then Exit Sub
If Not IsNumeric(Text7) Then
    Text7.SetFocus
    MsgBox "JUMLAH BARCODE HARUS ANGKA", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text8_LostFocus()
If Text8 = "" Then Exit Sub
If Not IsNumeric(Text9) Then
    Text8.SetFocus
    MsgBox "JUMLAH KOLOM KOSONG HARUS ANGKA", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text9_LostFocus()
If Text9 = "" Then Text9 = "-"
Text9 = Format(Text9, ">")
End Sub

Private Sub TmbClose_Click()
Unload Me
End Sub

Private Sub TmbDel_Click()
Dim Tanya
Tanya = MsgBox("YAKIN AKAN HAPUS DATA BARANG " + Trim(Text1), vbOKCancel, "YAKIN HAPUS KODE BARANG ?")
If Tanya = vbCancel Then Exit Sub
If Text1 = "" Then Exit Sub
SDel = "Delete From B003 where Kode_JNS = '" + Trim(Text1) + "'"
Set RDel = New ADODB.Recordset
RDel.Open SDel, CN, adOpenKeyset

Unload Me
Me.Show 1
End Sub

Private Sub TmbEdit_Click()
TmbDel.Enabled = True
MsgBox "SILAHKAN TULIS KODE JENIS BARANG YANG AKAN DIEDIT", vbInformation, "EDIT DATA BARANG"
Text1 = ""
Text1.SetFocus
Ganti = 1
TmbEdit.Enabled = False
End Sub

Private Sub TmbSave_Click()
Dim Params0, Params1, Params2, Params3, Params4, Params5, Params6 As Parameter
Dim Params7, Params8, Params9, Params10, Params11, Params12, Params13 As Parameter
Dim Kirim As New ADODB.Command

If Combo2 = "" Or Text1 = "" Or Text2 = "" Or Text6 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "DATA TIDAK BOLEH KOSONG"
    Exit Sub
End If


With Kirim
  .CommandType = adCmdStoredProc
  .CommandText = "ENTRIBRG"
  .ActiveConnection = CN
End With

If Val(Text6) < 1 Then
    MsgBox "HARGA JUAL MASIH KOSONG", vbCritical, "WARNING"
    Text6 = 0
    Text6.SetFocus
    Exit Sub
End If

Set Params0 = Kirim.CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
Set Params1 = Kirim.CreateParameter("@KodeJns", adVarChar, adParamInput, 50, Text1)
Set Params2 = Kirim.CreateParameter("@NamaJns", adVarChar, adParamInput, 50, Text2)
Set Params3 = Kirim.CreateParameter("@Satuan", adVarChar, adParamInput, 50, Combo2)
Set Params4 = Kirim.CreateParameter("@HargaJuaL", adCurrency, adParamInput, , Text6)

Kirim.Parameters.Append Params0
Kirim.Parameters.Append Params1
Kirim.Parameters.Append Params2
Kirim.Parameters.Append Params3
Kirim.Parameters.Append Params4

Kirim.Execute
If Not Kirim.Parameters("RETURN_VALUE") = 0 Then
    MsgBox "TRANSAKSI ANDA TIDAK BERHASIL", vbCritical, "TRANSAKSI GAGAL"
    Text1.SetFocus
Exit Sub
End If
Set Kirim = Nothing

Unload Me

Me.Show 1
End Sub

Private Sub txtCari_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub txtCari_LostFocus()
txtCari = Format(txtCari, ">")
End Sub

Private Sub txtSLS_GotFocus()
If txtSLS = "" Then txtSLS = 0
End Sub

Private Sub txtSLS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub txtSLS_LostFocus()
If txtSLS = "" Then txtSLS = 0
If Not IsNumeric(txtSLS) Then
    txtSLS.SetFocus
    MsgBox "PROSENTASE SALES HARUS ANGKA", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
txtSLS = Format(txtSLS, "##,###.00")
End Sub

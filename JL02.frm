VERSION 5.00
Begin VB.Form JL02 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "KOREKSI DATA PENJUALAN"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5415
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSLS 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1125
      TabIndex        =   20
      Text            =   "txtSLS"
      Top             =   6255
      Width           =   915
   End
   Begin VB.TextBox txtSB 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1215
      TabIndex        =   19
      Text            =   "txtSB"
      Top             =   5175
      Width           =   915
   End
   Begin VB.TextBox txtSK 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2655
      TabIndex        =   18
      Text            =   "txtSK"
      Top             =   5460
      Width           =   915
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   1575
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   2340
      Width           =   2040
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   1575
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1890
      Width           =   1410
   End
   Begin VB.CommandButton CmdClose 
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
      Left            =   3862
      TabIndex        =   6
      Top             =   3525
      Width           =   1380
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "&HAPUS"
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
      Left            =   2017
      TabIndex        =   5
      Top             =   3525
      Width           =   1380
   End
   Begin VB.CommandButton CmdEdit 
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
      Left            =   172
      TabIndex        =   4
      Top             =   3525
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Enabled         =   0   'False
      Height          =   360
      Left            =   1575
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1035
      Width           =   1410
   End
   Begin WASERDA.AutoCompleteCombo AutoCompleteCombo1 
      Height          =   315
      Left            =   1590
      TabIndex        =   1
      Top             =   1485
      Width           =   3780
      _ExtentX        =   6668
      _ExtentY        =   556
      BackColor       =   -2147483643
      FontItalic      =   0   'False
      FontName        =   "Tahoma"
      FontSize        =   8.25
      ForeColor       =   -2147483640
      Text            =   ""
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   1185
      Left            =   -1440
      ScaleHeight     =   1125
      ScaleWidth      =   7200
      TabIndex        =   17
      Top             =   3375
      Width           =   7260
   End
   Begin VB.Label Label10 
      Caption         =   "NOMINAL"
      Height          =   330
      Left            =   90
      TabIndex        =   16
      Top             =   2385
      Width           =   1410
   End
   Begin VB.Label Label8 
      Caption         =   "JML SATUAN"
      Height          =   285
      Left            =   90
      TabIndex        =   15
      Top             =   1935
      Width           =   1410
   End
   Begin VB.Label Label7 
      Caption         =   "BARANG"
      Height          =   330
      Left            =   90
      TabIndex        =   14
      Top             =   1485
      Width           =   1410
   End
   Begin VB.Label Label6 
      Caption         =   "KODE"
      Height          =   330
      Left            =   90
      TabIndex        =   13
      Top             =   1035
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "NO. URUT"
      Height          =   330
      Left            =   90
      TabIndex        =   12
      Top             =   585
      Width           =   1410
   End
   Begin VB.Label Label4 
      Caption         =   "NO. TRANSAKSI"
      Height          =   330
      Left            =   90
      TabIndex        =   11
      Top             =   135
      Width           =   1410
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label2"
      Height          =   330
      Left            =   1590
      TabIndex        =   10
      Top             =   585
      Width           =   645
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      Height          =   330
      Left            =   1590
      TabIndex        =   9
      Top             =   135
      Width           =   1410
   End
   Begin VB.Label Label9 
      Caption         =   "HARGA POKOK"
      Height          =   330
      Left            =   90
      TabIndex        =   8
      Top             =   2835
      Width           =   1410
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label11"
      Height          =   330
      Left            =   1590
      TabIndex        =   7
      Top             =   2790
      Width           =   2040
   End
End
Attribute VB_Name = "JL02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RJenis As New ADODB.Recordset
Private RCari As New ADODB.Recordset
Private RHarga As New ADODB.Recordset
Private SJenis, SCari, SHarga, MTStock, Editing As String

Private Sub AutoCompleteCombo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub AutoCompleteCombo1_LostFocus()
If AutoCompleteCombo1.Text = "" Then Exit Sub
SJenis = "Select KODE_JNS, HJUAL_PCS from B003 where NAMA_JNS = '" + Trim(AutoCompleteCombo1.Text) + "'"
Set RJenis = New ADODB.Recordset
RJenis.Open SJenis, CN, adOpenKeyset
If RJenis.RecordCount <> 0 Then
    Text1 = RJenis("KODE_JNS")
    Text3 = Format(RJenis("HJUAL_PCS"), "##,###.00")
    Text2 = 0
Else
    AutoCompleteCombo1.SetFocus
    MsgBox "NAMA JENIS BARANG BELUM TERDAFTAR. DAFTARKAN DAHULU LEWAT MENU SYSTEM", vbInformation, "NAMA JENIS BARANG BELUM TERDAFTAR"
End If
RJenis.Close
Set RJenis = Nothing
End Sub

Private Sub CmdClose_Click()
Unload Me
End Sub

Private Sub CmdDelete_Click()
Editing = ""
Dim Tanya
Tanya = MsgBox("YAKIN HAPUS 1 DATA TRANSAKSI PENJUALAN", vbOKCancel, "HAPUS DATA TRANSAKSI PENJUALAN")
If Tanya = vbCancel Then Exit Sub
Editing = "2"
Call SimpanData
Unload Me
End Sub

Private Sub CmdEdit_Click()
Dim Tanya
Editing = ""
Tanya = MsgBox("YAKIN AKAN MERUBAH DATA PENJUALAN ?", vbOKCancel, "EDIT DATA PENJUALAN")
If Tanya = vbCancel Then Exit Sub
Editing = "1"
Call SimpanData
Unload Me
End Sub

Private Sub SimpanData()
Dim Tanya
Dim Param1, Param2, Param3, Param4, Param5, Param6, Param7, Param8, Param9 As Parameter
Dim KirimData As New ADODB.Command

If Editing = "1" Then
    If AutoCompleteCombo1.Text = "" Or Label1 = "" Or Label2 = "" Or Text3 = 0 Or Text1 = "" Or Text2 = 0 Then
        MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "DATA TIDAK BOLEH KOSONG"
        Exit Sub
    End If
ElseIf Editing = "2" Then
    If AutoCompleteCombo1.Text = "" Or Label1 = "" Or Label2 = "" Or Text3 = "" Or Text1 = "" Or Text2 = "" Then
        MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "DATA TIDAK BOLEH KOSONG"
        Exit Sub
    End If
End If
With KirimData
  .CommandType = adCmdStoredProc
  .CommandText = "JUAL_2"
  .ActiveConnection = CN
End With

Set Params0 = KirimData.CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
Set Params1 = KirimData.CreateParameter("@NOTRANS", adVarChar, adParamInput, 12, Label1)
Set Params2 = KirimData.CreateParameter("@NOURUT", adInteger, adParamInput, , Label2)
Set Params3 = KirimData.CreateParameter("@KODEJNS", adVarChar, adParamInput, 17, Text1)
Set Params4 = KirimData.CreateParameter("@NAMAJNS", adVarChar, adParamInput, 50, AutoCompleteCombo1.Text)
Set Params5 = KirimData.CreateParameter("@JMLBAHAN", adCurrency, adParamInput, , Text2)
Set Params6 = KirimData.CreateParameter("@HJUALPCS", adCurrency, adParamInput, , Text3)
Set Params7 = KirimData.CreateParameter("@HARGAJUAL", adCurrency, adParamInput, , Label11)
Set Params8 = KirimData.CreateParameter("@Editing", adVarChar, adParamInput, 1, Trim(Editing))
Set Params9 = KirimData.CreateParameter("@NOMSLS", adCurrency, adParamInput, , CCur(Label11) * CCur(txtSLS) / 100)

KirimData.Parameters.Append Params0
KirimData.Parameters.Append Params1
KirimData.Parameters.Append Params2
KirimData.Parameters.Append Params3
KirimData.Parameters.Append Params4
KirimData.Parameters.Append Params5
KirimData.Parameters.Append Params6
KirimData.Parameters.Append Params7
KirimData.Parameters.Append Params8
KirimData.Parameters.Append Params9

KirimData.Execute
If Not KirimData.Parameters("RETURN_VALUE") = 0 Then
    MsgBox "TRANSAKSI ANDA TIDAK BERHASIL", vbCritical, "TRANSAKSI GAGAL"
    Text1.SetFocus
Exit Sub
End If
Set KirimData = Nothing
End Sub

Private Sub Form_Load()
Call Kosong
Label1 = NoTrans
Label2 = NoUrutTrans
End Sub

Private Sub Kosong()
AutoCompleteCombo1.Text = ""
Label1 = ""
Label2 = ""
Text3 = 0
Label11 = ""
Text1 = ""
Text2 = 0
End Sub

Private Sub Label2_Change()
If Label2 = "" Then Exit Sub
SCari = "Select * from JL01 where No_Trans = '" + Trim(Label1) + "' and No_urut Like  '" + Label2 + "'"
Set RCari = New ADODB.Recordset
RCari.Open SCari, CN, adOpenKeyset
If RCari.RecordCount <> 0 Then
    Text1 = RCari("Kode_Jns")
    AutoCompleteCombo1.Text = RCari("Nama_Jns")
    Text2 = RCari("JML_BAHAN")
    Text3 = Format(RCari("HJUAL_PCS"), "##,###.00")
    Label11 = Format(RCari("HARGA_JUAL"), "##,###.00")
    txtSLS = CCur(RCari("SLS"))
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text1_LostFocus()
If Text1 = "" Then Exit Sub
SJenis = "Select NAMA_JNS, HJUAL_PCS from B003 where KODE_JNS = '" + Trim(Text1.Text) + "'"
Set RJenis = New ADODB.Recordset
RJenis.Open SJenis, CN, adOpenKeyset
If RJenis.RecordCount <> 0 Then
    AutoCompleteCombo1.Text = RJenis("NAMA_JNS")
    Text3 = Format(RJenis("HJUAL_PCS"), "##,###.00")
    Text2 = 0
    Text2.SetFocus
Else
    Text1.SetFocus
    MsgBox "KODE JENIS BARANG BELUM TERDAFTAR. DAFTARKAN DAHULU LEWAT MENU SYSTEM", vbInformation, "KODE JENIS BARANG BELUM TERDAFTAR"
End If
RJenis.Close
Set RJenis = Nothing
End Sub

Private Sub Text2_GotFocus()
TextBoxSelected Text2
If Text2 = 0 Then Text2 = ""
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Or KeyAscii = 13 Or KeyAscii = Asc(".")) Then KeyAscii = 0
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text2_LostFocus()
If Text1 = "" Or AutoCompleteCombo1.Text = "" Then Exit Sub
If Text2 = "" Then
    Text2 = 0
    Exit Sub
End If
If Not IsNumeric(Text2) Then
    Text2 = 0
    Text2.SetFocus
    MsgBox "JUMLAH BARANG DIJUAL HARUS ANGKA", vbCritical, "TIPE DATA SALAH"
    Exit Sub
End If
Label11 = Format(CCur(Text2) * CCur(Text3), "##,###.00")
End Sub

Private Sub Text3_GotFocus()
TextBoxSelected Text3
If Text3 = 0 Then Text3 = ""
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text3_LostFocus()
If Text3 = "" Then Text3 = 0
If Not IsNumeric(Text3) Then
    Text3 = 0
    Text3.SetFocus
    MsgBox "HARGA SATUAN HARUS ANGKA", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
Text3 = Format(Text3, "##,###.00")
Label11 = Format(CCur(Text2) * CCur(Text3), "##,###.00")
End Sub

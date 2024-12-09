VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form JL03 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PENJUALAN BARANG"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12255
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   12255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   360
      Left            =   15
      TabIndex        =   33
      Top             =   6765
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   635
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483633
      BackColorFixed  =   16761087
      ForeColorFixed  =   0
      BackColorSel    =   12632064
      BackColorBkg    =   -2147483633
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame FrameTip 
      BorderStyle     =   0  'None
      Caption         =   "Frame7"
      Height          =   330
      Left            =   0
      TabIndex        =   34
      Top             =   2550
      Width           =   4605
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ">Double Click pada  baris transaksi untuk edit jumlah penjualan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   45
         TabIndex        =   35
         Top             =   60
         Width           =   4530
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H80000018&
         BackStyle       =   1  'Opaque
         Height          =   330
         Left            =   0
         Top             =   0
         Width           =   4605
      End
   End
   Begin WASERDA.AutoCompleteCombo AutoCompleteCombo1 
      Height          =   315
      Left            =   1980
      TabIndex        =   2
      Top             =   2190
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   556
      BackColor       =   -2147483643
      FontItalic      =   0   'False
      FontName        =   "Tahoma"
      FontSize        =   8.25
      ForeColor       =   -2147483640
      Text            =   ""
   End
   Begin VB.CommandButton TblOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   370
      Left            =   11415
      TabIndex        =   6
      Top             =   2160
      Width           =   350
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   75
      TabIndex        =   1
      Text            =   "Text7"
      Top             =   2205
      Width           =   960
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   5535
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   2220
      Width           =   700
   End
   Begin VB.TextBox Text16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   8370
      TabIndex        =   5
      Text            =   "Text16"
      Top             =   2220
      Width           =   700
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   6555
      TabIndex        =   4
      Text            =   "Text9"
      Top             =   2220
      Width           =   1000
   End
   Begin Crystal.CrystalReport CRPT 
      Left            =   5745
      Top             =   8025
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
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      Left            =   6187
      TabIndex        =   17
      Text            =   "Text8"
      Top             =   30
      Width           =   1320
   End
   Begin VB.TextBox txtKel 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   360
      Left            =   1050
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   16
      Text            =   "txtKel"
      Top             =   1260
      Width           =   1770
   End
   Begin VB.TextBox txtNama 
      BackColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   1050
      Locked          =   -1  'True
      MaxLength       =   35
      TabIndex        =   12
      Text            =   "txtNama"
      Top             =   870
      Width           =   3555
   End
   Begin VB.TextBox txtKetua 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   360
      Left            =   2835
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   11
      Text            =   "txtKetua"
      Top             =   1260
      Width           =   1770
   End
   Begin VB.TextBox txtNo 
      Height          =   360
      Left            =   1050
      TabIndex        =   0
      Text            =   "txtNo"
      Top             =   480
      Width           =   1770
   End
   Begin VB.TextBox txtDaerah 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   360
      Left            =   4620
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   10
      Text            =   "txtDaerah"
      Top             =   1260
      Width           =   1770
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
      Left            =   660
      TabIndex        =   9
      Top             =   7965
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
      Left            =   10215
      TabIndex        =   7
      Top             =   7965
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   1185
      Left            =   -120
      ScaleHeight     =   1125
      ScaleWidth      =   12555
      TabIndex        =   8
      Top             =   7815
      Width           =   12615
   End
   Begin MSFlexGridLib.MSFlexGrid Grid3 
      Height          =   375
      Left            =   -30
      TabIndex        =   32
      Top             =   2160
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   661
      _Version        =   393216
      Rows            =   1
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   16777152
      BackColorBkg    =   16777152
      GridColor       =   0
      ScrollBars      =   2
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4665
      Left            =   0
      TabIndex        =   31
      Top             =   1740
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   8229
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483633
      BackColorFixed  =   16761087
      ForeColorFixed  =   0
      BackColorSel    =   12632064
      BackColorBkg    =   -2147483633
      GridColor       =   0
      MergeCells      =   2
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblHIS 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "999,999,999.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Left            =   9780
      TabIndex        =   29
      Top             =   885
      Width           =   2445
   End
   Begin VB.Label lblPAGU 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "999,999,999.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Left            =   9780
      TabIndex        =   28
      Top             =   540
      Width           =   2445
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "SISA :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Left            =   8970
      TabIndex        =   27
      Top             =   1215
      Width           =   885
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL TRANS :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Left            =   7635
      TabIndex        =   26
      Top             =   885
      Width           =   2220
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "PAGU :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Left            =   8865
      TabIndex        =   25
      Top             =   540
      Width           =   990
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   765
      Left            =   2550
      TabIndex        =   24
      Top             =   7020
      Width           =   9645
   End
   Begin VB.Label Label42 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "SUB TOTAL :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   150
      TabIndex        =   23
      Top             =   7245
      Width           =   2310
   End
   Begin VB.Label Label31 
      Caption         =   "Tanggal Penjualan"
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
      Left            =   4792
      TabIndex        =   22
      Top             =   30
      Width           =   1365
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
      Left            =   2632
      TabIndex        =   21
      Top             =   30
      Width           =   1320
   End
   Begin VB.Label Label9 
      Caption         =   "Tanggal System"
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
      Left            =   8347
      TabIndex        =   19
      Top             =   30
      Width           =   1230
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label10"
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
      Left            =   9735
      TabIndex        =   18
      Top             =   30
      Width           =   1320
   End
   Begin VB.Label lblREG 
      BackStyle       =   0  'Transparent
      Caption         =   "NO REG"
      Height          =   210
      Left            =   90
      TabIndex        =   15
      Top             =   555
      Width           =   1035
   End
   Begin VB.Label lblNama 
      BackStyle       =   0  'Transparent
      Caption         =   "N A M A"
      Height          =   210
      Left            =   90
      TabIndex        =   14
      Top             =   945
      Width           =   1035
   End
   Begin VB.Label lblKEL 
      BackStyle       =   0  'Transparent
      Caption         =   "KELOMPOK"
      Height          =   210
      Left            =   90
      TabIndex        =   13
      Top             =   1335
      Width           =   1035
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      Height          =   675
      Left            =   -135
      Top             =   7125
      Width           =   12615
   End
   Begin VB.Label lblSISA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "999,999,999.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Left            =   9780
      TabIndex        =   30
      Top             =   1215
      Width           =   2445
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Left            =   -165
      Top             =   390
      Width           =   12585
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
      Left            =   1365
      TabIndex        =   20
      Top             =   30
      Width           =   1365
   End
End
Attribute VB_Name = "JL03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RCari As New ADODB.Recordset
Private RJenis As New ADODB.Recordset
Private RTrans As New ADODB.Recordset
Private RNo As New ADODB.Recordset
Private RGrid As New ADODB.Recordset
Private RDel, RDel2 As New ADODB.Recordset
Private RSPL As New ADODB.Recordset
Private RSGL As New ADODB.Recordset
Private RPin As New ADODB.Recordset

Private SCari, SJenis, STrans, SNo, SGrid, SDel  As String
Private StatusData, SDel2, SSGL, SSPL, SPin As String
Private MTStock As String
Private Pesan As Boolean
Private HJualPCS As Currency
Private SBayar, JmlAkhir, HargaBeli, HargaJual
Private Hasil, Barkod, TotalJual

Private h() As String
Private TtlSLS As Currency

Private Sub autocompletecombo1_GotFocus()
SendKeys "{F4}"
End Sub

Private Sub AutoCompleteCombo1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF5
        SBayar = 1
        FrameBayar.Visible = True
        Text17.SetFocus
    Case vbKeyF1
        If AutoCompleteCombo1.Text <> "" Then AutoCompleteCombo1.Text = ""
        KodeJenis = ""
        NamaJenis = ""
        IB03.Show 1
        If KodeJenis = "" Or NamaJenis = "" Then Exit Sub
        Text7 = KodeJenis
        AutoCompleteCombo1.Text = NamaJenis
        Call Text7_LostFocus
End Select
End Sub

Private Sub Command1_Click()
Dim Tanya, NoTran
Dim ProsesData As New ADODB.Command
Dim KirimData As New ADODB.Command
Dim Kirim0, Kirim1, Kirim2, Kirim3, Kirim4, Kirim5 As Parameter
Dim Kirim6, Kirim7, Kirim8, Kirim9, Kirim10, Kirim11 As Parameter
Dim Kirim12, Kirim13, Kirim14, Kirim15, Kirim16, Kirim17 As Parameter
Dim Kirim18 As Parameter
Dim Kirim19 As Parameter
Dim Kirim20, Kirim21, Kirim22, Kirim23, Kirim24, Kirim25, Kirim26 As Parameter

Dim Nomor As Double
Dim KDPiutang, NoPiutang, NoNas, NamaNas, TglMulai, TglJatuh, SyaratByr, SglLain
Dim KDCustomer, KDSales

IPT1 = ""
IPT2 = ""
IPT3 = ""

NoTran = Label3

Tanya = MsgBox("YAKIN PROSES TRANSAKSI PENJUALAN NO. " + Label3 + " ?", vbOKCancel, "SIMPAN TRANSAKSI PENJUALAN")
If Tanya = vbCancel Then Exit Sub

With ProsesData
  .CommandType = adCmdStoredProc
  .CommandText = "JUAL_3"
  .ActiveConnection = CN
End With

Set Kirim0 = ProsesData.CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
Set Kirim1 = ProsesData.CreateParameter("@NOTRANS", adVarChar, adParamInput, 50, Label3)
Set Kirim2 = ProsesData.CreateParameter("@OPERATOR", adVarChar, adParamInput, 50, Operator)
Set Kirim3 = ProsesData.CreateParameter("@TGLJUAL", adDate, adParamInput, , Text8)
Set Kirim4 = ProsesData.CreateParameter("@NOCUSTO", adVarChar, adParamInput, 50, txtNo)
Set Kirim5 = ProsesData.CreateParameter("@NMCUSTO", adVarChar, adParamInput, 50, txtNama)
Set Kirim6 = ProsesData.CreateParameter("@TOTALJUAL", adCurrency, adParamInput, , TotalJual)


ProsesData.Parameters.Append Kirim0
ProsesData.Parameters.Append Kirim1
ProsesData.Parameters.Append Kirim2
ProsesData.Parameters.Append Kirim3
ProsesData.Parameters.Append Kirim4
ProsesData.Parameters.Append Kirim5
ProsesData.Parameters.Append Kirim6

ProsesData.Execute
If Not ProsesData.Parameters("RETURN_VALUE") = 0 Then
    MsgBox "TRANSAKSI ANDA TIDAK BERHASIL", vbCritical, "TRANSAKSI GAGAL"
    Exit Sub
End If
Set ProsesData = Nothing

IPT1 = Trim(Label3)
IPT2 = Trim(Text8)
IPT3 = Trim(label0)

Call SimpanNota
JL03A.Label3 = Trim(Label3)

Grid.Clear
Call Siap
Nomor = Val(Right(Label3, 7)) + 1
Call Kosong
Text7.SetFocus
Label3 = Trim("3.") + Digit(2, Trim(NoUser)) + "." + Digit(7, Nomor)

Unload Me
JL03A.Show 1
End Sub

Private Sub SimpanNota()
SSave = "Select * from JL03_NOTA"
Set RSave = New ADODB.Recordset
RSave.Open SSave, CN, adOpenKeyset, adLockBatchOptimistic
RSave.AddNew
    RSave("NO_TRANS") = Label3
    RSave("TGL_JUAL") = Text8
    RSave("TGL_S") = Label10
    RSave("T_JUAL") = CCur(Label4)
    RSave("T_DISKON") = 0
    RSave("SUB_T") = CCur(Label4)
    RSave("S_DISKON") = 0
    RSave("S_DISKON_PERSEN") = 0
    RSave("TOTAL") = CCur(Label4)
    RSave("TUNAI") = CCur(Label4)
    RSave("PIUTANG") = 0
    RSave("NO_PIUTANG") = 0
    RSave("ABA") = 0
    RSave("NO_ABA") = 0
    RSave("TOTAL_BAYAR") = CCur(Label4)
    RSave("KEMBALI") = 0
    RSave("TERBILANG") = Terbilang(Label4)
    RSave("USER_CODE") = Operator
RSave.UpdateBatch adAffectAllChapters
RSave.Close
Set RSave = Nothing

SSave = "Select * From C012 where NoNas = '" + Trim(txtNo) + "'"
Set RSave = New ADODB.Recordset
RSave.Open SSave, CN, adOpenDynamic, adLockOptimistic
If RSave.RecordCount <> 0 Then
    RSave("TotalTrans") = RSave("TotalTrans") + CCur(Label4)
    RSave("Sisa") = RSave("Sisa") - CCur(Label4)
    RSave.Update
End If
RSave.Close
Set RSave = Nothing
End Sub

Private Sub Grid_DblClick()
Call Batal
If Grid.Rows = 1 Then
    MsgBox "TIDAK ADA DATA BARANG BAKU PRODUKSI YANG DAPAT DIEDIT", vbCritical, "DATA KOSONG"
    Exit Sub
End If
NoUrutTrans = Grid.TextMatrix(Grid.Row, 0)
NoTrans = Label3
JL02.Show 1
Grid.Clear
Call Siap
Call Isi_Grid
Text7.SetFocus
End Sub

Private Sub TblOK_Click()
Text2_LostFocus
Text9_LostFocus
Text16_LostFocus

If CCur(Label4) + CCur(Grid3.TextMatrix(0, 9)) > lblSISA Then
    MsgBox "TRANSAKSI LEBIH BESAR DARI SISA PAGU", vbCritical, "WARNING"
    Exit Sub
End If

Dim Param1, Param2, Param3, Param4, Param5, Param6, Param7 As Parameter
Dim Param8, Param9 As Parameter
Dim KirimData As New ADODB.Command

If AutoCompleteCombo1.Text = "" Or Val(Text2) = 0 Or Text7 = "" Or Val(Text9) = 0 Or Trim(Grid3.TextMatrix(0, 9)) = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "DATA TIDAK BOLEH KOSONG"
    Text2.SetFocus
    Exit Sub
End If
With KirimData
  .CommandType = adCmdStoredProc
  .CommandText = "JUAL_1"
  .ActiveConnection = CN
End With

Set Param0 = KirimData.CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
Set Param1 = KirimData.CreateParameter("@NOTRANS", adVarChar, adParamInput, 50, Label3)
Set Param2 = KirimData.CreateParameter("@KODEJNS", adVarChar, adParamInput, 50, Text7)
Set Param3 = KirimData.CreateParameter("@JMLBAHAN", adCurrency, adParamInput, , Text2)
Set Param4 = KirimData.CreateParameter("@HJUALPCS", adCurrency, adParamInput, , Text9)
Set Param5 = KirimData.CreateParameter("@HARGAJUAL", adCurrency, adParamInput, , CCur(Grid3.TextMatrix(0, 9)))
Set Param6 = KirimData.CreateParameter("@OPERATOR", adVarChar, adParamInput, 50, Operator)
Set Param7 = KirimData.CreateParameter("@DISCOUNT", adCurrency, adParamInput, , CCur(Text16))

KirimData.Parameters.Append Param0
KirimData.Parameters.Append Param1
KirimData.Parameters.Append Param2
KirimData.Parameters.Append Param3
KirimData.Parameters.Append Param4
KirimData.Parameters.Append Param5
KirimData.Parameters.Append Param6
KirimData.Parameters.Append Param7
KirimData.Execute

If Not KirimData.Parameters("RETURN_VALUE") = 0 Then
    MsgBox "TRANSAKSI ANDA TIDAK BERHASIL", vbCritical, "TRANSAKSI GAGAL"
    Text7.SetFocus
Exit Sub
End If
Set KirimData = Nothing
    
Call Isi_Grid
Call Batal
Text7.SetFocus

End Sub

Private Sub Text16_GotFocus()
If Text16 = 0 Then Text16 = ""
End Sub

Private Sub Text16_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF5
        SBayar = 1
        FrameBayar.Visible = True
        Text17.SetFocus
End Select
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Or KeyAscii = 13 Or KeyAscii = Asc(".")) Then KeyAscii = 0
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text16_LostFocus()
If Text16 = "" Then Text16 = 0
If Not IsNumeric(Text16) Then
    Text16.SetFocus
    MsgBox "NOMINAL DISCOUNT HARUS ANGKA", vbCritical, "TYPE DATA SALAH"
    Exit Sub
End If
With Grid3
    If Text9 = "" Then Text9 = 0
    .Col = 9: .Text = Format((CCur(Text2) * CCur(Text9)) - CCur(Text16), "##,###.00"): .CellAlignment = 7
End With

SendKeys "{home}+{end}"
Text16 = Format(Text16, "##,###.00")
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF5
        SBayar = 1
        FrameBayar.Visible = True
        Text17.SetFocus
End Select
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Or KeyAscii = 13 Or KeyAscii = Asc(".")) Then KeyAscii = 0
If KeyAscii = 13 Then Text9.SetFocus
TextBoxSelected Text9
End Sub

Private Sub Text2_LostFocus()
If Text2 = "" Then
    Text2 = 0
    Exit Sub
End If
If Text7 = "" Or AutoCompleteCombo1.Text = "" Then Exit Sub
If CCur(JmlAkhir) < CCur(Text2) Then
    On Error Resume Next
    Text2.SetFocus
    MsgBox "STOCK BARANG TIDAK CUKUP", vbCritical, "JUMLAH TIDAK CUKUP"
    Text2 = 0
    Exit Sub
End If
If Len(Text2) > 6 Then
    Dim Hasil, Barkod
    Barkod = Right(Text2, 13)
    Hasil = Len(Text2) - Len(Barkod)
    Text2 = Mid(Text2, 1, Hasil)
    With Grid3
        .Col = 9: .Text = Format(CCur(Text2) * CCur(Text9), "##,###.00"): .CellAlignment = 7
    End With
    Call TblOK_Click
    
    Text7 = Barkod
    Call Text7_LostFocus
End If

If Not IsNumeric(Text2) Then
    Text2.SetFocus
    MsgBox "JUMLAH PENJUALAN HARUS ANGKA", vbCritical, "TIPE DATA SALAH"
    Exit Sub
End If
If Text9 = "" Then Exit Sub
With Grid3
    .Col = 8: .Text = Format(CCur(Text2) * CCur(Text9), "##,###.00"): .CellAlignment = 7
End With
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    On Error Resume Next
    SendKeys vbTab
End If
End Sub

Private Sub Text7_LostFocus()
If Text7 = "" Then Exit Sub

SJenis = "Select NAMA_JNS, JML_AKHIR, HJUAL_PCS, DISCOUNT, SATUAN_PCS, SLS, SATUAN, HBELI from B003 where KODE_JNS = '" + Trim(Text7.Text) + "'"
Set RJenis = New ADODB.Recordset
RJenis.Open SJenis, CN, adOpenKeyset
If RJenis.RecordCount <> 0 Then
    If Val(RJenis("JML_AKHIR")) > 0 Then
        AutoCompleteCombo1.Text = RJenis("NAMA_JNS")
        JmlAkhir = RJenis("JML_AKHIR")
        HargaBeli = RJenis("HBELI")
        Text9 = Format(RJenis("HJUAL_PCS"), "##,###.00")
        Text16 = RJenis("DISCOUNT")
        txtSLS = RJenis("SLS")
        Text2.SetFocus
        SendKeys "{end}"
        Text2 = 1
        With Grid3
            Row = 0
            .Col = 3: .Text = RJenis("SATUAN"): .CellAlignment = 4
        End With
    Else
        Text7 = ""
        Text7.SetFocus
        MsgBox "STOCK BARANG KOSONG", vbInformation, "WARNING"
    End If
Else
    Text7 = ""
    'Text7.SetFocus
    MsgBox "KODE BARANG BELUM TERDAFTAR. DAFTARKAN DAHULU LEWAT MENU SYSTEM", vbInformation, "KODE BARANG BELUM TERDAFTAR"
End If
RJenis.Close
Set RJenis = Nothing

End Sub

Private Sub AutoCompleteCombo1_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then Text2.SetFocus
TextBoxSelected Text2
End Sub

Private Sub AutoCompleteCombo1_LostFocus()
If AutoCompleteCombo1.Text = "" Then Exit Sub
SJenis = "Select KODE_JNS, JML_AKHIR, HJUAL_PCS, DISCOUNT, SATUAN_PCS, SLS, SATUAN, HBELI from B003 where NAMA_JNS = '" + Trim(AutoCompleteCombo1.Text) + "'"
Set RJenis = New ADODB.Recordset
RJenis.Open SJenis, CN, adOpenKeyset
If RJenis.RecordCount <> 0 Then
    If Val(RJenis("JML_AKHIR")) > 0 Then
        Text7 = RJenis("KODE_JNS")
        JmlAkhir = RJenis("JML_AKHIR")
        HargaBeli = RJenis("HBELI")
        HargaJual = RJenis("HJUAL_PCS")
        Text9 = Format(RJenis("HJUAL_PCS"), "##,###.00")
        Text16 = RJenis("DISCOUNT")
        txtSLS = RJenis("SLS")
        Text2.SetFocus
        Text2 = 1
        With Grid3
            Row = 0
            .Col = 3: .Text = RJenis("SATUAN"): .CellAlignment = 4
        End With
    Else
        AutoCompleteCombo1.SetFocus
        MsgBox "STOCK BARANG KOSONG. LAKUKAN PROSES PEMBELIAN BARANG", vbInformation, "WARNING"
        SendKeys "{F4}"
    End If
Else
    On Error Resume Next
    'AutoCompleteCombo1.SetFocus
    MsgBox "NAMA BARANG BELUM TERDAFTAR. DAFTARKAN DAHULU LEWAT MENU SYSTEM", vbInformation, "NAMA BARANG BELUM TERDAFTAR"
    SendKeys "{F4}"
End If
RJenis.Close
Set RJenis = Nothing
End Sub

Private Sub Command2_Click()
SDel = "Delete from JL01 Where NO_TRANS = '" + Trim(Label3) + "'"
Set RDel = New ADODB.Recordset
RDel.Open SDel, CN, adOpenKeyset

Unload Me
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = (Screen.Width - Width) / 2

SDel = "Delete from JL01 Where NO_TRANS = '" + Trim(Label3) + "'"
Set RDel = New ADODB.Recordset
RDel.Open SDel, CN, adOpenKeyset

ClearTextBoxes Me

Text7.Width = 1700
Text2.Width = 755
Text9.Width = 1040
Text16.Width = 735
AutoCompleteCombo1.Width = 3485

Call DaftarBrg
Call Kosong
Call Cari_Trans

End Sub

Private Sub Cari_Trans()
Dim Nomor As Double
STrans = "Select No_Trans from JL01 where user_code = '" + Operator + "'"
Set RTrans = New ADODB.Recordset
RTrans.Open STrans, CN, adOpenKeyset
If RTrans.RecordCount <> 0 Then
    If Pesan = True Then MsgBox "MASIH ADA TRANSAKSI PENJUALAN YANG BELUM SELESAI", vbInformation, "DAFTAR TRANSAKSI PENJUALAN TERSIMPAN"
    Label3 = RTrans("No_Trans")
    Call Isi_Grid
Else
    SNo = "Select NoJual from C013 where usercode = '" + Operator + "'"
    Set RNo = New ADODB.Recordset
    RNo.Open SNo, CN, adOpenKeyset
    If RNo.RecordCount <> 0 Then
        Nomor = Val(RNo("NoJual"))
        Label3 = Trim("3.") + Digit(2, Trim(NoUser)) + "." + Digit(7, Nomor)
    End If
    RNo.Close
    Set RNo = Nothing
End If
RTrans.Close
Set RTrans = Nothing
End Sub

Private Sub DaftarBrg()
SCari = "Select Nama_JNS From B003 WHERE JML_AKHIR > 0 order by Nama_JNS"
Set RCari = New ADODB.Recordset
RCari.Open SCari, CN, adOpenKeyset
If RCari.RecordCount <> 0 Then
    RCari.MoveFirst
    Do While Not RCari.EOF
        AutoCompleteCombo1.AddItem RCari("Nama_JNS")
    RCari.MoveNext
    Loop
Else
    AutoCompleteCombo1.Visible = False
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Kosong()
Call Batal
Call Siap
Label3 = ""
Label4 = 0
Label14 = 0
Label10 = Tanggal
Label29 = ""
Label35 = 0
Label37 = 0
Label41 = 0
Label42 = "SUB TOTAL :"
Label43 = 0
Label46 = 0
Text1 = ""
Text3 = 0
Text4 = 0
Text5 = 0
Text8 = Tanggal
Text10 = ""
Text15 = 0
Text14 = 0
Text17 = 0
Text18 = 0
Text19 = "-"
Text13 = "-"
Text20 = ""
Text25 = "-"
Text24 = "-"

lblPAGU = "0.00"
lblHIS = "0.00"
lblSISA = "0.00"

End Sub

Private Sub Batal()
AutoCompleteCombo1.Text = ""
Text2 = 1
Text7 = ""
Text9 = ""
Text16 = 0
End Sub

Private Sub Siap()
With Grid
    .Rows = 1
    .Row = 0
    .RowHeight(0) = 510
    .Cols = 10
    .Col = 0: .ColWidth(0) = 500: .Text = "No.": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 8: .CellFontName = "Arial Narrow"
    .Col = 1: .ColWidth(1) = 1750: .Text = "KODE": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 8: .CellFontName = "Arial Narrow"
    .Col = 2: .ColWidth(2) = 3500: .Text = "BARANG": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 8: .CellFontName = "Arial Narrow"
    .Col = 3: .ColWidth(3) = 750: .Text = "STN": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 8: .CellFontName = "Arial Narrow"
    .Col = 4: .ColWidth(4) = 800: .Text = "JML": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 8: .CellFontName = "Arial Narrow"
    .Col = 5: .ColWidth(5) = 1100: .Text = "HRG PCS": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 8: .CellFontName = "Arial Narrow"
    .Col = 6: .ColWidth(6) = 800: .Text = "DISC(Rp)": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 8: .CellFontName = "Arial Narrow"
    .Col = 7: .ColWidth(7) = 35: .CellFontBold = True: .CellFontSize = 10: .CellFontBold = True: .CellFontSize = 8: .CellFontName = "Arial Narrow"
    .Col = 8: .ColWidth(8) = 1100: .Text = "JML HARGA": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 8: .CellFontName = "Arial Narrow"
    .Col = 9: .ColWidth(9) = 1100: .Text = "HARGA NETT": .CellAlignment = 4: .CellFontBold = True: .CellFontSize = 8: .CellFontName = "Arial Narrow"
End With

Grid3.Width = 11400
With Grid3
    .RowHeight(0) = 400
    .Row = 0
    .Cols = 10
    .Col = 0: .ColWidth(0) = 475
    .Col = 1: .ColWidth(1) = 1750
    .Col = 2: .ColWidth(2) = 3500
    .Col = 3: .ColWidth(3) = 750
    .Col = 4: .ColWidth(4) = 800
    .Col = 5: .ColWidth(5) = 1100
    .Col = 6: .ColWidth(6) = 800
    .Col = 7: .ColWidth(7) = 35
    .Col = 8: .ColWidth(8) = 1100
    .Col = 9: .ColWidth(9) = 1100
End With

With Grid2
    .Row = 0
    .Cols = 9
    .Col = 0: .ColWidth(0) = 475
    .Col = 1: .ColWidth(1) = 5250
    .Col = 2: .ColWidth(2) = 750
    .Col = 3: .ColWidth(3) = 800
    .Col = 4: .ColWidth(4) = 1100
    .Col = 5: .ColWidth(5) = 800
    .Col = 6: .ColWidth(6) = 35
    .Col = 7: .ColWidth(7) = 1100
    .Col = 8: .ColWidth(8) = 1100
    .RowHeight(0) = 400
End With

Call PindahText
End Sub

Private Sub PindahText()
If Grid.Rows = 1 Then
    Grid3.Move Grid.Left + 57, Grid.CellTop + 2270
Else
    Grid3.Move Grid.Left + 57, Grid.CellTop + 2100
End If

Text7.Left = Grid3.Left + 465
Text7.Top = Grid3.Top + 65
AutoCompleteCombo1.Left = Grid3.Left + 2200
AutoCompleteCombo1.Top = Grid3.Top + 15
Text2.Left = Grid3.Left + 6475
Text2.Top = Grid3.Top + 65
Text9.Left = Text2.Left + Text2.Width + 50
Text9.Top = Grid3.Top + 65
Text16.Left = Text9.Left + Text9.Width + 50
Text16.Top = Grid3.Top + 65
TblOK.Left = Grid3.Left + Grid3.Width - 50
TblOK.Top = Grid3.Top
Grid3.Left = Grid3.Left - 55
FrameTip.Left = Grid3.Left
FrameTip.Top = Grid3.Top + Grid3.Height + 25
With Grid3
    .Row = 0: .Col = 0: .Text = Grid.Rows: .CellAlignment = 4
End With
End Sub

Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF5
        SBayar = 1
        FrameBayar.Visible = True
        Text17.SetFocus
End Select
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn Or KeyAscii = 13 Or KeyAscii = Asc(".")) Then KeyAscii = 0
On Error Resume Next
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text9_LostFocus()
If Text7 = "" Or AutoCompleteCombo1.Text = "" Then Exit Sub
If Text9 = "" Then
    Text9 = 0
    Exit Sub
End If
If CCur(HargaBeli) > CCur(Text9) Then
    On Error Resume Next
    MsgBox "HARGA JUAL LEBIH RENDAH DARI HARGA BELI", vbCritical, "HARGA JUAL KURANG"
    Text9 = HargaJual
    'Exit Sub
End If
If Not IsNumeric(Text9) Then
    Text9.SetFocus
    MsgBox "HARGA JUAL HARUS ANGKA", vbCritical, "TIPE DATA SALAH"
    Exit Sub
End If
With Grid3
    .Col = 9: .Text = Format(CCur(Text2) * CCur(Text9), "##,###.00"): .CellAlignment = 7
End With
'Text9 = Format(Text9, "##,###.00")
End Sub

Private Sub txtNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub txtNo_LostFocus()
If txtNo = "" Then Exit Sub

SKode = "Select * from C012 where NoNas = '" + Trim(txtNo) + "'"
Set RKode = New ADODB.Recordset
RKode.Open SKode, CN, adOpenKeyset
If RKode.RecordCount <> 0 Then
    txtNama = RKode("Nama")
    txtKel = RKode("NamaKel")
    txtKetua = RKode("KetuaKel")
    txtDaerah = RKode("Alamat1")
    lblPAGU = Format(RKode("Pendapatan") * 50 / 100, "##,###.00")
    lblHIS = Format(RKode("TotalTrans"), "##,###.00")
    lblSISA = Format(RKode("Sisa"), "##,###.00")
End If
RKode.Close
Set RKode = Nothing
Text1 = Format(Text1, ">")

If Val(lblSISA) = 0 Then
    MsgBox "SISA PAGU TIDAK CUKUP UNTUK MELAKUKAN TRANSAKSI", vbCritical, "WARNING"
    txtNo = ""
    txtNama = ""
    txtKel = ""
    txtKetua = ""
    txtDaerah = ""
    lblPAGU = "0.00"
    lblHIS = "0.00"
    lblSISA = "0.00"
    txtNo.SetFocus
End If
End Sub

Private Sub Isi_Grid()
Dim Brs, NomDiskon, TtlDiskon
Brs = 1
Label4 = 0
Text15 = 0
Label41 = 0
TotalJual = 0
NomDiskon = 0
TtlDiskon = 0
SGrid = "Select * from JL01 where NO_TRANS = '" + Label3 + "' order by no_urut"
Set RGrid = New ADODB.Recordset
RGrid.Open SGrid, CN, adOpenKeyset, adLockReadOnly
If RGrid.RecordCount <> 0 Then
    RGrid.MoveFirst
    Do While Not RGrid.EOF
        With Grid
        .Rows = Brs + 1
        .Row = Brs
        .RowHeight(Brs) = 350
        .MergeRow(Brs) = False
        .MergeCol(7) = False
        .Col = 0: .Text = RGrid("NO_URUT"): .CellAlignment = 4
        .Col = 1: .Text = RGrid("KODE_JNS"): .CellAlignment = 4
        .Col = 2: .Text = RGrid("NAMA_JNS")
        
        SJenis = "Select SATUAN from B003 where KODE_JNS = '" + Trim(RGrid("KODE_JNS")) + "'"
        Set RJenis = New ADODB.Recordset
        RJenis.Open SJenis, CN, adOpenKeyset
        If RJenis.RecordCount <> 0 Then
            .Col = 3: .Text = RJenis("SATUAN"): .CellAlignment = 4
        End If
        RJenis.Close
        Set RJenis = Nothing
        
        .Col = 4: .Text = RGrid("JML_BAHAN"): .CellAlignment = 4
        .Col = 5: .Text = Format(RGrid("HJUAL_PCS"), "##,###.00")
        .Col = 6: .Text = RGrid("DISCOUNT"): .CellAlignment = 4
        .Col = 8: .Text = Format(RGrid("HARGA_JUAL"), "##,###.00")
            NomDiskon = CCur(Grid.TextMatrix(Brs, 8)) - CCur(RGrid("NOMDISC"))
        .Col = 9: .Text = Format(NomDiskon, "##,###.00")
            Label4 = Format(CCur(Label4) + CCur(RGrid("HARGA_JUAL")), "##,###.00")
            Text15 = Format(CCur(Text15) + CCur(RGrid("NOMDISC")), "##,###.00")
            TotalJual = CCur(TotalJual) + CCur(RGrid("JML_BAHAN"))
            TtlDiskon = CCur(TtlDiskon) + CCur(NomDiskon)
        End With
        Brs = Brs + 1
    RGrid.MoveNext
    Loop
    Call PindahText
Else
    Unload Me
    JL03.Show
End If
RGrid.Close
Set RGrid = Nothing

With Grid2
    .Row = 0:
    .Col = 1: .Text = "TOTAL PENJUALAN": .CellAlignment = 4
    .Col = 3: .Text = Format(TotalJual, "##,###"): .CellAlignment = 4
    .Col = 7: .Text = Label4
    .Col = 8: .Text = Format(TtlDiskon, "##,###.00")
End With
Label4 = Format(Label4, "##,###.00")
Label35 = Label4
With Grid3
    .Row = 0
    .Col = 0: .Text = Grid.Rows: .CellAlignment = 4
    .Col = 3: .Text = "": .CellAlignment = 4
    .Col = 5: .Text = "": .Col = 6: .Text = "": .Col = 8: .Text = "": .Col = 9
End With
Label41 = Format(CCur(Label35) - CCur(Text15), "##,###.00")
Label46 = Format(CCur(Label41) + CCur(Label43), "##,###.00")

If Val(Label4) > Val(lblSISA) Then
    Label11.ForeColor = &HFF&
    lblSISA.ForeColor = &HFF&
    Command1.Enabled = False
Else
    Label11.ForeColor = &H404040
    lblSISA.ForeColor = &H404040
    Command1.Enabled = True
End If
End Sub

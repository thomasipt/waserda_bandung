VERSION 5.00
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#13.2#0"; "Codejock.SkinFramework.v13.2.1.ocx"
Begin VB.Form LOGON 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "LOGIN"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   180
   ClientWidth     =   4665
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Logon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3019
      TabIndex        =   5
      Top             =   1575
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&LOG IN"
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
      Left            =   266
      TabIndex        =   4
      Top             =   1575
      Width           =   1380
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2625
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   750
      Width           =   1950
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2625
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   315
      Width           =   1950
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   1365
      Left            =   -135
      ScaleHeight     =   1305
      ScaleWidth      =   6660
      TabIndex        =   6
      Top             =   1455
      Width           =   6720
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework 
      Left            =   2280
      Top             =   990
      _Version        =   851970
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   1425
      TabIndex        =   3
      Top             =   750
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "USER CODE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   1425
      TabIndex        =   2
      Top             =   315
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   -240
      Picture         =   "Logon.frx":0A02
      Stretch         =   -1  'True
      Top             =   -90
      Width           =   1890
   End
End
Attribute VB_Name = "LOGON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CONEC As String
Private RST As New ADODB.Recordset
Private RCari As New ADODB.Recordset
Private RSQL As New ADODB.Recordset
Private RDel As New ADODB.Recordset

Private sSQL, SCari, SDel As String
Private Sts As String
Private Lolos, NomorIP

Private Sub Command1_Click()
Lolos = 0
Set Con = New ADODB.Connection
Con.ConnectionString = "Provider= SQLOLEDB.1;" & Trim(KONEKSI)
Con.ConnectionTimeout = 10
Con.CursorLocation = adUseServer
Con.Open

sSQL = "Select * From C013 where UserCode = '" + Trim(Text1) + "'and password = '" + Trim(Text2) + "'"
Set RST = New ADODB.Recordset
RST.Open sSQL, Con, adOpenKeyset, adLockReadOnly
If RST.RecordCount <> 0 Then
    If RST("Status") = 2 Then
        MsgBox "ANDA HARUS MENGGANTI PASSWORD ", vbCritical, "G1.ANTI PASSWORD"
        User = Text1
        Me.Hide
        E001.Show
        Exit Sub
    End If
    
  If RST("Status") = 1 Then
        MsgBox "USER ANDA NON AKTIF HUBUNGI ADMINISTRATOR", vbCritical, "NON AKTIF"
        Text2 = ""
        Text2.SetFocus
        Exit Sub
  Else
        CCab = RST("CodeCab")
        Status = RST("Main")
        Operator = Trim(RST("UserCode"))
        NoUser = RST("NoUrut")
        CodeBagian = RST("CodeBag")
        Call Kosong
        Call Cek_TGL
    If Lolos = 1 Then Exit Sub
        Me.Hide
        Unload Me
  End If
Else
    Text2 = ""
    Text1 = ""
    Text1.SetFocus
    MsgBox "AKSES DITOLAK!           ", vbCritical, "PASSWORD"
Exit Sub
End If
RST.Close
Set RST = Nothing

Con.Close
Set Con = Nothing
Call KoneksiData
End Sub

Private Sub Cek_TGL()
SCari = "Select Tanggal from A001 "
Set RCari = New ADODB.Recordset
RCari.Open SCari, Con, adOpenDynamic, adLockOptimistic
If RCari.RecordCount <> 0 Then
    If DateValue(Date) > DateValue(RCari("Tanggal")) Then
        If Month(Date) <> Month(DateValue(RCari("Tanggal"))) Then
            SDel = "Update C012 Set TotalTrans = 0, Sisa = Pendapatan"
            Set RDel = New ADODB.Recordset
            RDel.Open SDel, Con, adOpenKeyset
        End If
        RCari("Tanggal") = DateValue(Date)
        RCari.Update
    ElseIf DateValue(Date) < DateValue(RCari("Tanggal")) Then
        MsgBox "TANGGAL SISTEM TERAKHIR = " + Format(RCari("Tanggal"), "DD/MM/YYYY") + " CEK SETTING TANGGAL PADA KOMPUTER", vbCritical, "WARNING"
        End
    End If
Tanggal = RCari("Tanggal")
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Command2_Click()
Dim i
i = MsgBox("ANDA YAKIN AKAN KELUAR DARI SYSTEM ?", vbQuestion + vbOKCancel, "INTEGRATED STORE SYSTEM")
If i = vbOK Then
    Unload Me
Else
    Exit Sub
End If
End Sub

Private Sub Form_Load()
Open App.Path + "\KONEKSI.DAT" For Input As #1
    Input #1, KONEKSI
    Input #1, NAMA_KOMPUTER
    Input #1, FOLDER_BACKUP
Close #1

SkinFramework.LoadSkin App.Path + "\Vista.cjstyles", ""
SkinFramework.ApplyWindow Me.hWnd
SkinFramework.ApplyOptions = SkinFramework.ApplyOptions Or xtpSkinApplyMetrics

Call Kosong
End Sub

Private Sub Kosong()
Text1 = ""
Text2 = "********"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text1_LostFocus()
Text1 = Format(Text1, ">")
End Sub

Private Sub Text2_GotFocus()
Text2 = ""
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub



VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form V_C012 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "INFORMASI PETANI"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11835
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   11835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   10320
      TabIndex        =   0
      Top             =   7980
      Width           =   1380
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   1185
      Left            =   -180
      ScaleHeight     =   1125
      ScaleWidth      =   12555
      TabIndex        =   1
      Top             =   7830
      Width           =   12615
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   7470
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   13176
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   -2147483633
      AllowUserResizing=   3
      Appearance      =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   ">Double Click pada  baris untuk cetak informasi"
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
      Left            =   60
      TabIndex        =   3
      Top             =   7560
      Width           =   3390
   End
End
Attribute VB_Name = "V_C012"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RCari As New ADODB.Recordset
Private RSimpan As New ADODB.Recordset
Private RKode As New ADODB.Recordset
Private SCari, SSimpan, SKode As String

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call Siap
Call IsiGrid
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
SCari = "Select * From C012 Order By NoNas"
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

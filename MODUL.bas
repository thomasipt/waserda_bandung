Attribute VB_Name = "MODUL"
Public CN As New ADODB.Connection
Public Con As New ADODB.Connection
Public Operator As String
Public Tanggal As String
Public Super As String
Public NoNas, NamaNas As String
Public NomorRek As String
Public CCab, NCab As String
Public NoUrut As Boolean
Public Tglinput As String
Public NoPinjaman As String
Public CodePinjaman As String
Public NomTrans As Currency
Public Status, User As String
Public CodeSl, NamaSl As String
Public Sale, ByBeli As Currency
Public NoUser, StatusEdit As String
Public BAkhir As Integer
Public KodeJenis, NoHutang, Cbiaya As String
Public NoUrutTrans, NoTrans, SGanti, NoSIN As String
Public TotalRetur, JmlRetur, NamaJenis As String
Public KodeGudang, NamaGudang, NamaForm As String
Public IPT1, IPT2, IPT3 As String
Public A As Integer
Public AA As Integer
Public CodeBagian As String
Public NoBG As String
Public StsFormBeli As String
Public StsCetak As String

Public KONEKSI As String
Public NAMA_KOMPUTER As String
Public NAMA_DIVISI As String
Public STS_VALIDASI As String
Public FOLDER_BACKUP As String

Sub Main()
LOGON.Show
End Sub

Public Sub KoneksiData()

Set CN = New ADODB.Connection
CN.ConnectionString = "Provider= SQLOLEDB.1;" & Trim(KONEKSI)
CN.ConnectionTimeout = 30
CN.CursorLocation = adUseServer
CN.Open

MAIN_MENU.Show
End Sub

Public Function Digit(Panjang, Nilai As Double) As String
Dim Y, NilaiP As Double
Dim Kar, NilaiS As String

If Panjang <= 0 Then Panjang = 1

NilaiS = Trim(Str(Nilai))
NilaiP = Len(NilaiS)
If NilaiP >= Panjang Then Panjang = NilaiP

Kar = " "
For Y = 1 To Panjang
    Kar = Trim(Kar) + "0"
Next
If (Panjang - NilaiP) <= 0 Then
    Digit = NilaiS
Else
    Digit = Mid(Kar, 1, (Panjang - NilaiP)) + NilaiS
End If
End Function

Public Function Satuan(ByVal Nilai As Currency) As String
Select Case Nilai
    Case 1: Satuan = "SATU "
    Case 2: Satuan = "DUA "
    Case 3: Satuan = "TIGA "
    Case 4: Satuan = "EMPAT "
    Case 5: Satuan = "LIMA "
    Case 6: Satuan = "ENAM "
    Case 7: Satuan = "TUJUH "
    Case 8: Satuan = "DELAPAN "
    Case 9: Satuan = "SEMBILAN "
End Select
End Function

Public Function Ribuan(ByVal Bilangan As Currency) As String
Dim A, b As Currency
Dim C As String

C = ""
A = Bilangan \ 1000
b = Bilangan Mod 1000
If A > 1 Then C = C + Satuan(A) + "RIBU "
If A = 1 Then C = C + "SERIBU "

A = b \ 100
b = b Mod 100
If A > 1 Then C = C + Satuan(A) + "RATUS "
If A = 1 Then C = C + "SERATUS "

A = b \ 10
b = b Mod 10
If A > 1 Then C = C + Satuan(A) + "PULUH "
If A = 1 Then
    If b = 0 Then Ribuan = C + "SEPULUH "
    If b = 1 Then Ribuan = C + "SEBELAS "
    If b > 1 Then Ribuan = C + Satuan(b) + "BELAS "
Else
    Ribuan = C + Satuan(b)
End If
End Function

Public Function Terbilang(ByVal Bilangan As Currency) As String

Dim A, b As Currency
Dim C As String


A = Bilangan \ 1000000000
b = Bilangan Mod 1000000000
C = "#"
If A > 0 Then C = Ribuan(A) + "MILYAR "

A = b \ 1000000
b = b Mod 1000000
If A > 0 Then C = C + Ribuan(A) + "JUTA "

A = b \ 1000
b = b Mod 1000
If A > 1 Then C = C + Ribuan(A) + "RIBU "
If A = 1 Then C = C + "Seribu "
Terbilang = C + Ribuan(b) + "RUPIAH#"
End Function

Public Function SumHari(Dari, Ke As Date) As Integer
If Ke - Dari <= 1 Then
    SumHari = 1
Else
    SumHari = Ke - Dari
End If
End Function

Public Function Sisip(Kar As String, Posisi As Integer, Kar2 As String) As String
Dim PJ As Integer
Dim Akhir As String
Dim depan As String
PJ = Len(Kar)
If Len(Kar) < Len(Kar2) Then
    Sisip = Kar2
Else
    If Posisi = 1 Then Sisip = Kar2 + Mid(Kar, 2, PJ - 1)
    If Posisi > 1 And Posisi < PJ Then
        depan = Mid(Kar, 1, Posisi - 1)
        Akhir = Mid(Kar, Posisi + 1, PJ - Posisi)
        Sisip = depan + Kar2 + Akhir
    End If
    If Posisi = PJ Then Sisip = Mid(Kar, 1, Posisi - 1) + Kar2
End If
End Function

Public Function RKanan(NData, CFormat) As String
RKanan = Format(NData, CFormat)
RKanan = Space(Len(CFormat) - Len(RKanan)) + RKanan
End Function

Public Function BulanStr(ByVal CBulan As Currency) As String
Select Case CBulan
    Case 1: BulanStr = " Jan. "
    Case 2: BulanStr = " Feb. "
    Case 3: BulanStr = " Mar. "
    Case 4: BulanStr = " Apr. "
    Case 5: BulanStr = " Mei "
    Case 6: BulanStr = " Juni "
    Case 7: BulanStr = " Juli "
    Case 8: BulanStr = " Agt. "
    Case 9: BulanStr = " Sept. "
    Case 10: BulanStr = " Okt. "
    Case 11: BulanStr = " Nov. "
    Case 12: BulanStr = " Des. "
End Select
BulanStr = BulanStr
End Function

Public Function BulanString(ByVal CBulan As Currency) As String
Select Case CBulan
    Case 1: BulanString = " JANUARI "
    Case 2: BulanString = " FEBRUARI "
    Case 3: BulanString = " MARET "
    Case 4: BulanString = " APRIL "
    Case 5: BulanString = " MEI "
    Case 6: BulanString = " JUNI "
    Case 7: BulanString = " JULI "
    Case 8: BulanString = " AGUSTUS "
    Case 9: BulanString = " SEPTEMBER "
    Case 10: BulanString = " OKTOBER "
    Case 11: BulanString = " NOVEMBER "
    Case 12: BulanString = " DESEMBER "
End Select
BulanString = BulanString
End Function

Public Function BulanInt(IBulan As String) As String
Select Case IBulan
    Case " JANUARI ": BulanInt = 1
    Case " FEBRUARI ": BulanInt = 2
    Case " MARET ": BulanInt = 3
    Case " APRIL ": BulanInt = 4
    Case " MEI ": BulanInt = 5
    Case " JUNI ": BulanInt = 6
    Case " JULI ": BulanInt = 7
    Case " AGUSTUS ": BulanInt = 8
    Case " SEPTEMBER ": BulanInt = 9
    Case " OKTOBER ": BulanInt = 10
    Case " NOVEMBER ": BulanInt = 11
    Case " DESEMBER ": BulanInt = 12
End Select
BulanInt = BulanInt
End Function

Public Function HariStr(ByVal CHari As Currency) As String
Select Case CHari
    Case 1: HariStr = " MINGGU "
    Case 2: HariStr = " SENIN "
    Case 3: HariStr = " SELASA "
    Case 4: HariStr = " RABU "
    Case 5: HariStr = " KAMIS "
    Case 6: HariStr = " JUMAT "
    Case 7: HariStr = " SABTU "
End Select
HariStr = HariStr
End Function

Public Function BlkKoma(Bilangan As Double) As String
Dim A, D As Double
Dim b, E, F As Double
Dim C As String
If Bilangan > 2000000000 Then
    C = ""
    
    D = Mid(Bilangan, 1, 7)
    A = D \ 1000000
    b = D Mod 1000000
    If A > 0 Then C = Ribuan(A) + "Milyar "
    
    E = Mid(Bilangan, 2, 10)
    A = E \ 1000000
    b = E Mod 1000000
    If A > 0 Then C = C + Ribuan(A) + "Juta "
    
    F = Mid(Bilangan, 5, 10)
    A = F \ 1000
    b = F Mod 1000
    If A > 0 Then C = C + Ribuan(A) + "Ribu "
    If A = 1 Then C = C + "Seribu "
BlkKoma = C + Ribuan(b)
Else
    C = ""
    A = Bilangan \ 1000000000
    b = Bilangan Mod 1000000000
    If A > 0 Then C = Ribuan(A) + "Milyar "
    
    A = b \ 1000000
    b = b Mod 1000000
    If A > 0 Then C = C + Ribuan(A) + "Juta "
    
    A = b \ 1000
    b = b Mod 1000
    If A > 1 Then C = C + Ribuan(A) + "Ribu "
    If A = 1 Then C = C + "Seribu "
BlkKoma = C + Ribuan(b)
End If
End Function

Public Sub ClearTextBoxes(frmClearMe As Form)
Dim txt As Control
For Each txt In frmClearMe
  If TypeOf txt Is TextBox Then txt.Text = ""
Next
End Sub

Public Sub TextBoxSelected(textBoxSel As Object)
textBoxSel.SelStart = 0
textBoxSel.SelLength = Len(textBoxSel.Text)
End Sub

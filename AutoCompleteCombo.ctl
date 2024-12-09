VERSION 5.00
Begin VB.UserControl AutoCompleteCombo 
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3120
   ScaleHeight     =   330
   ScaleWidth      =   3120
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   -15
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   0
      Width           =   3150
   End
End
Attribute VB_Name = "AutoCompleteCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' this is free software, made by papadogiannakis vagelis
' in Heraclion, Crete, HELLAS
' you can use and modify it at will
' but credits should be given where appropriate
' I am not resonsible for anything your computer may suffer with this code
' (just kidding)
' it needs a little work to be perfect
'
'
' if you like this code and use it, send me a postcard
' (snail mail will be provided by email)
'
' if you have suggestions/flames/whatever email me at papas@rocketmail.
' this is version 1 of the control



Option Explicit
Option Compare Text

Private ItemsArray() As String

Public Event Change()
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub Combo1_Change()
    RaiseEvent Change
End Sub

Private Sub UserControl_Initialize()
    ReDim ItemsArray(0) As String
End Sub

Private Sub UserControl_InitProperties()
    Text = Ambient.DisplayName
    UserControl.Width = 1215
    UserControl.Height = 495
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    Combo1.Top = 0
    Combo1.Left = 0
    With UserControl
        Combo1.Width = .Width
        .Height = 315
    End With
End Sub

Private Sub combo1_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub combo1_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub combo1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode < 65 Then RaiseEvent KeyUp(KeyCode, Shift): Exit Sub
    Dim iCtr As Long
    Dim MySelStart As Integer

    MySelStart = Combo1.SelStart
    If MySelStart = 0 Then Exit Sub 'MySelStart = 1
    
    For iCtr = 0 To UBound(ItemsArray)
         
        If Left$(ItemsArray(iCtr), Len(Combo1.Text)) = Combo1.Text Then
            Combo1.Text = ItemsArray(iCtr)
            Combo1.SelStart = MySelStart
            Combo1.SelLength = 255
            RaiseEvent KeyUp(KeyCode, Shift)
            Exit Sub
        End If
    Next
    Combo1.SelStart = MySelStart
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub


Public Property Get BackColor() As OLE_COLOR
    BackColor = Combo1.BackColor
End Property
Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
    Combo1.BackColor = NewValue
    PropertyChanged "BackColor"
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal NewValue As Boolean)
    UserControl.Enabled = NewValue
    Combo1.Enabled = NewValue
End Property



Public Property Get Font() As StdFont
    Set Font = Combo1.Font
End Property
Public Property Set Font(ByVal NewValue As StdFont)
    Set Combo1.Font = NewValue
    PropertyChanged "Font"
End Property


Public Property Get FontName() As String
    FontName = Combo1.FontName
End Property
Public Property Let FontName(ByVal NewValue As String)
    Combo1.FontName = NewValue
    PropertyChanged "FontName"
End Property


Public Property Get FontBold() As Boolean
    FontBold = Combo1.FontBold
End Property
Public Property Let FontBold(ByVal NewValue As Boolean)
    Combo1.FontBold = NewValue
    PropertyChanged "FontBold"
End Property


Public Property Get FontItalic() As Boolean
    FontItalic = Combo1.FontItalic
End Property
Public Property Let FontItalic(ByVal NewValue As Boolean)
    Combo1.FontItalic = NewValue
    PropertyChanged "FontItalic"
End Property


Public Property Get FontUnderline() As Boolean
    FontUnderline = Combo1.FontUnderline
End Property
Public Property Let FontUnderline(ByVal NewValue As Boolean)
    Combo1.FontUnderline = NewValue
    PropertyChanged "FontUnderline"
End Property


Public Property Get FontStrikethru() As Boolean
    FontStrikethru = Combo1.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal NewValue As Boolean)
    Combo1.FontStrikethru = NewValue
    PropertyChanged "FontStrikethru"
End Property


Public Property Get FontSize() As Single
    FontSize = Combo1.FontSize
End Property
Public Property Let FontSize(NewValue As Single)
    Combo1.FontSize = NewValue
    PropertyChanged "FontSize"
End Property


Public Property Get ForeColor() As OLE_COLOR
    ForeColor = Combo1.ForeColor
End Property
Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
    Combo1.ForeColor = NewValue
    PropertyChanged "ForeColor"
End Property


Public Property Get Text() As String
    Text = Combo1.Text
End Property

Public Property Let Text(ByVal NewValue As String)
    Combo1.Text = NewValue
    PropertyChanged "Text"
End Property

Public Property Get SelStart() As Long
    SelStart = Combo1.SelStart
End Property
Public Property Let SelStart(ByVal NewValue As Long)
    Combo1.SelStart = NewValue
End Property


Public Property Get SelLength() As Long
    SelLength = Combo1.SelLength
End Property
Public Property Let SelLength(ByVal NewValue As Long)
    Combo1.SelLength = NewValue
End Property


Public Property Get SelText() As String
    SelText = Combo1.SelText
End Property
Public Property Let SelText(ByVal NewValue As String)
    Combo1.SelText = NewValue
End Property


Public Property Get ItemData(Index As Integer) As Long
    If Index < 0 Or Index > Combo1.ListCount - 1 Then
        Err.Raise 381
    Else
        ItemData = Combo1.ItemData(Index)
    End If
End Property

Public Property Let ItemData(Index As Integer, ByVal NewValue As Long)
    If Index < 0 Or Index > Combo1.ListCount Then
        Err.Raise 381
    Else
        Combo1.ItemData(Index) = NewValue
    End If
End Property


Public Property Get list(Index As Integer) As String
If Index < 0 Or Index > UBound(ItemsArray) - 1 Then
    Err.Raise 381
Else
    list = Combo1.list(Index)
End If
End Property


Public Property Get ListCount() As Integer
    ListCount = Combo1.ListCount
End Property


Public Property Get SelectedItem() As Integer
Dim iAns As Integer
Dim sText As String
Dim iCtr As Long

iAns = -1
sText = Combo1.Text

For iCtr = 0 To UBound(ItemsArray)
    If sText = ItemsArray(iCtr) Then iAns = iCtr: iCtr = UBound(ItemsArray)
Next

SelectedItem = iAns
End Property


Public Sub Clear()
    ReDim ItemsArray(0) As String
    Combo1.Clear
End Sub

Public Sub RemoveItem(Index As Integer)
    ArrayRemoveItem ItemsArray, Index
    Combo1.RemoveItem (Index)
End Sub

Public Sub AddItem(Item As String)
    If Item = "" Then Exit Sub
    
    Combo1.AddItem UCase(Item)
    
    If ItemsArray(0) = "" Then
        ItemsArray(0) = Item
    Else
        ReDim Preserve ItemsArray(UBound(ItemsArray) + 1) As String
        ItemsArray(UBound(ItemsArray)) = Item
    End If
End Sub

Public Sub AddItems(ParamArray Items() As Variant)
    Dim iCtr As Integer
    
    For iCtr = 0 To UBound(Items)
        AddItem UCase(CStr(Items(iCtr)))
    Next
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        BackColor = .ReadProperty("BackColor", Combo1.BackColor)
        Enabled = .ReadProperty("Enabled", True)
        FontBold = .ReadProperty("FontBold", False)
        FontItalic = .ReadProperty("FontItalic", False)
        FontName = .ReadProperty("FontName", "Tahoma")
        FontSize = .ReadProperty("FontSize", 8)
        FontStrikethru = .ReadProperty("FontStrikethru", False)
        FontUnderline = .ReadProperty("FontUnderline", False)
        ForeColor = .ReadProperty("ForeColor", Combo1.ForeColor)
        Text = .ReadProperty("Text", "")
    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "BackColor", BackColor
        .WriteProperty "Enabled", Enabled, True
        .WriteProperty "FontBold", FontBold, False
        .WriteProperty "FontItalic", FontItalic
        .WriteProperty "FontName", FontName, "Ms Sans Serif"
        .WriteProperty "FontSize", FontSize, 8
        .WriteProperty "FontStrikethru", FontStrikethru, False
        .WriteProperty "FontUnderline", FontUnderline, False
        .WriteProperty "ForeColor", ForeColor
        .WriteProperty "Text", Text
    End With
End Sub

Private Sub ArrayRemoveItem(ItemArray As Variant, ByVal ItemElement As Long)
    Dim lCtr As Long
    Dim lTop As Long
    Dim lBottom As Long
    
    lTop = UBound(ItemArray)
    lBottom = LBound(ItemArray)
    
    For lCtr = ItemElement To lTop - 1
        ItemArray(lCtr) = ItemArray(lCtr + 1)
    Next
    
    ReDim Preserve ItemArray(lBottom To lTop - 1)
End Sub


Public Property Get ListIndex() As Integer
ListIndex = Combo1.ListIndex

End Property

Public Property Let ListIndex(ByVal NewValue As Integer)
Combo1.ListIndex = NewValue
PropertyChanged "ListIndex"

End Property

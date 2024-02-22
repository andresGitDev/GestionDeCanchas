VERSION 5.00
Begin VB.UserControl uNum 
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1650
   ScaleHeight     =   330
   ScaleWidth      =   1650
   Begin VB.TextBox txtNumero 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Text            =   "0"
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "uNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' control numBox

' .num      es tipo double,  **DEFAULT**
' .ssNum    es string con PUNTO, para armar sql o algo asi
' .suNum    es string con formato regional, el que vera el usuario, sera COMA si es español.


' nota
' la tecla "." del teclado numerico la toma como punto decimal, no importa la configuracion regional
' las teclas "." y "," del teclado no numerico las tomara SIN cambios

Private mKeyCode As Integer
Private mNumero As Double
Private mDecimales As Long
Private mDecimalesCalculo As Long
Private mMaximo 'As Double variant asi puedo preguntar isempty()
Private mMinimo 'As Double

Public Event cambio(numero As Double)

Public Property Get ssNum() As String
    ssNum = x2s(mNumero)
End Property
Public Property Let ssNum(cual As String)
    On Error GoTo ufa
    mNumero = Round(CDbl(cual), mDecimalesCalculo)
    PropertyChanged "Num"
    rever
    Exit Property
ufa:
    mNumero = 0
    PropertyChanged "Num"
    rever
End Property
Public Property Get suNum() As String
    suNum = CStr(mNumero)
End Property
Public Property Let suNum(cual As String)
    On Error GoTo ufa
    'mNumero = Round(CDbl(cual), mDecimalesCalculo)
    mNumero = s2n(cual, mDecimalesCalculo)
    PropertyChanged "Num"
    rever
    Exit Property
ufa:
    mNumero = 0
    PropertyChanged "Num"
    rever
End Property

Public Property Get num() As Double
Attribute num.VB_UserMemId = 0
Attribute num.VB_MemberFlags = "200"
    num = Round(mNumero, mDecimalesCalculo)
End Property
Public Property Let num(cual As Double)
    mNumero = Round(cual, mDecimalesCalculo)
    PropertyChanged "Num"
    rever
End Property

Public Property Get Decimales() As Long
    Decimales = mDecimales
End Property
Public Property Let Decimales(cuantos As Long)
    mDecimales = cuantos
    PropertyChanged "Decimales"
End Property
Public Property Get DecimalesCalculo() As Long
    DecimalesCalculo = mDecimalesCalculo
    PropertyChanged "DecimalesCalculo"
End Property
Public Property Let DecimalesCalculo(cuantos As Long)
    mDecimalesCalculo = cuantos
    PropertyChanged "DecimalesCalculo"
End Property

Public Property Get enabled() As Boolean
    enabled = UserControl.enabled
End Property
Public Property Let enabled(como As Boolean)
    UserControl.enabled = como
    PropertyChanged "Enabled"
End Property
'Public Property Get Visible() As Boolean
'    Visible = txtNumero.Visible
'    PropertyChanged "Visible"
'End Property
'
'Public Property Let Visible(como As Boolean)
'    txtNumero.Visible = como
'End Property
Public Property Get Locked() As Boolean
    Locked = txtNumero.Locked
    PropertyChanged "Locked"
End Property

Public Property Let Locked(como As Boolean)
    txtNumero.Locked = como
End Property

'Public Property Get Maximo() 'As Double
'    Maximo = mMaximo
'    PropertyChanged "Maximo"
'End Property
'
'Public Property Let Maximo(cual As Double)
'    mMaximo = cual
'End Property
'Public Property Get Minimo() 'As Double
'    Minimo = mMinimo
'    PropertyChanged "Minimo"
'End Property
'
'Public Property Let Minimo(cual As Double)
'    mMinimo = cual
'End Property
'Public Property Get MaxLen() 'As Double
'    MaxLen = mMaxLen
'    PropertyChanged "MaxLen"
'End Property
'
'Public Property Let MaxLen(cual As Long)
'    mMaxLen = cual
'End Property
'*******************************************************

Private Sub rever()
    Dim x
    x = Format(mNumero, IIf(mDecimales = 0, "#,#", "#,0." & Left("00000000", mDecimales))) 'Format(mNumero)
    txtNumero = IIf(IsNumeric(x), x, txtNumero)
End Sub


Private Sub txtNumero_Change()
    On Error GoTo ufa
    Dim x As Double
     x = Round(CDbl(txtNumero), mDecimalesCalculo)
     mNumero = x
    Exit Sub
ufa:
    On Error Resume Next
    'mNumero = 0
    'txtNumero = mNumero
'    rever
End Sub

Private Sub txtNumero_GotFocus()
'    PintoFocoActivo
'    txtNumero = IIf(mNumero = 0, "", mNumero)
    txtNumero.SelStart = 0
    txtNumero.SelLength = Len(txtNumero.Text)
End Sub

Private Sub txtNumero_KeyDown(KeyCode As Integer, Shift As Integer)
    mKeyCode = KeyCode
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
    On Error GoTo ufa
    Dim sep As String
    
    sep = SeparadorDecimal()
    If mKeyCode = 110 And sep <> "." Then
        KeyAscii = 0
        If InStr(txtNumero, sep) = 0 Then SendKeys sep
    End If
'    If KeyAscii = vbKeyBack Then
''        mNumero = CDbl(Left(txtNumero, Len(txtNumero) - 1))
''    ElseIf KeyAscii = 45 Then
''        mNumero = -CDbl(txtNumero)
''        KeyAscii = 0
''    Else
''        mNumero = CDbl(txtNumero & Chr(KeyAscii))
'    End If
'    Exit Sub
ufa:
'    'If KeyAscii <> 45 Then KeyAscii = 0
'    KeyAscii = 0
End Sub

Private Sub txtNumero_LostFocus()
    'txtNumero = Round(mNumero, mDecimales)
    rever
    RaiseEvent cambio(mNumero)
End Sub

Private Sub UserControl_GotFocus()
    On Error Resume Next
    PintoFocoActivo
    txtNumero.SetFocus
End Sub

Private Sub UserControl_Initialize()
    txtNumero = ""
End Sub
Private Sub UserControl_Resize()
    txtNumero.Height = UserControl.Height
    txtNumero.Width = UserControl.Width
End Sub

'*******************************************************

'maximo, minimo, maxlen, minlen
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim abc As Boolean
    On Error Resume Next
     mDecimales = PropBag.ReadProperty("Decimales", 2)
     mDecimalesCalculo = PropBag.ReadProperty("DecimalesCalculo", 4)
     UserControl.enabled = PropBag.ReadProperty("Enabled", True)
     txtNumero.Locked = PropBag.ReadProperty("Locked", False)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next
    PropBag.WriteProperty "Decimales", mDecimales, 2
    PropBag.WriteProperty "DecimalesCalculo", mDecimalesCalculo, 4
    PropBag.WriteProperty "Enabled", UserControl.enabled, True
    PropBag.WriteProperty "Locked", txtNumero.Locked, False
End Sub


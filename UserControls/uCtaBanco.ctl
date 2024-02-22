VERSION 5.00
Begin VB.UserControl uCtaBanco 
   BackStyle       =   0  'Transparent
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6570
   ScaleHeight     =   345
   ScaleWidth      =   6570
   Begin Gestion.ucCoDe uCoDe 
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   30
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   529
      CodigoWidth     =   1000
   End
   Begin VB.TextBox txtNumero 
      Height          =   330
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1485
   End
End
Attribute VB_Name = "uCtaBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event cambio(codigo)

Private mTipo As String
Private mCuentaContable As String

Private Sub uCoDe_cambio(codigo As Variant)
    Dim tempo
    If codigo > 0 Then
        
        tempo = obtenerDeSQL("select numero, tipo,  cuenta_con from ctasbank where codigo = '" & codigo & "' ")
        
        If IsEmpty(tempo) Then
            Vacio
        Else
            txtNumero = sSinNull(tempo(0))
            mTipo = sSinNull(obtenerDeSQL("select descripcion from TipoCtas where codigo = " & tempo(1)))
            mCuentaContable = sSinNull(tempo(2))
        End If
    Else
        Vacio
     End If
     
    RaiseEvent cambio(codigo)
End Sub

Private Sub UserControl_Initialize()
    On Error GoTo ufa
    uCoDe.ini "select bg.descripcion from ctasBank CB inner join BancosGrales BG on cb.banco = bg.codigo where cb.codigo = '###' and bg.activo = 1 and cb.activo = 1 ", "select cb.codigo, bg.descripcion as [ Banco                       ], Numero as [ Numero       ]  from ctasBank CB inner join BancosGrales BG on cb.banco = bg.codigo  where cb.activo = 1 and bg.activo = 1 ", False, False, DataEnvironment1.Sistema

ufa:
End Sub

Public Property Get NroCuenta() As String
    NroCuenta = Trim(txtNumero)
End Property

Public Property Get codigo() As Long
    codigo = uCoDe.codigo
End Property
Public Property Let codigo(cual As Long)
    uCoDe.codigo = cual
End Property

Public Property Get DESCRIPCION() As String
    DESCRIPCION = uCoDe.DESCRIPCION
End Property

Public Property Get tipo() As String
    tipo = mTipo
End Property

Public Property Get CuentaContable() As String
    CuentaContable = mCuentaContable
End Property

Public Property Let enabled(como As Boolean)
    uCoDe.enabled = como
End Property
Public Property Get enabled() As Boolean
    enabled = uCoDe.enabled
End Property
'Public Sub ini()
'    uCoDe.ini "select "
'End Sub

'Private Sub UserControl_Resize()
'    uCoDe.Width = UserControl.Width
'    uCoDe.Height = UserControl.Height
'End Sub
Private Sub Vacio()
    mTipo = ""
    txtNumero = ""
    mCuentaContable = ""
End Sub

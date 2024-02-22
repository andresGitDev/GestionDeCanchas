VERSION 5.00
Begin VB.UserControl ucCoDe 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7170
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   330
   ScaleWidth      =   7170
   Begin VB.CommandButton cmdBuscar 
      DisabledPicture =   "ucCoDe.ctx":0000
      Height          =   285
      Left            =   1455
      Picture         =   "ucCoDe.ctx":058A
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   345
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   300
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "Descripcion"
      Top             =   0
      Width           =   5295
   End
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Text            =   "Codigo"
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "ucCoDe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit 'mod 17/12/4
' Lito Explicit
'start 24/11/4 end 25/11/4 -end 25/11/4

'OJO: asume conexion    daTaenvironment1.Sistema = adodb.conection
'
'son 3 Parametros: 1 solo obligatorio, 1 boton busqueda opcional, 1 bool predeterminados
'
' 1) string sql 'trucho' trucho por los 3 numerales ###
'   hay q darle el sql para buscar descripcion para un codigo representado por ### (3 NUMERALES)
'   (el prg reemplaza el ### por el valor q se intenta poner en el txtBox "codigo")
'
' 2) string sql para el frmBuscar
'
' 3) TRUE si el codigo es string, (false default)
'
'EJs:
'
'   Producto.ini _
'       "select descripcion from producto where codigo = '###' and activo = 1", _
'       " select codigo as [Codigo   ], descripcion as [Desc del Producto     ] from producto where activo = 1 order by codigo", _
'       True
'
'    factu.ini _
'       "select razonsocial from FacturaVenta where activo = 1 and NroFactura = ### ", _
'       "select NroFactura, RazonSocial from facturaVenta where activo = 1 order by NroFactura  "
'
'Propiedades
'
'   .CodigoWidth, ( el cmd y Descripcion se acomodan al usercontrol)
'   .Codigo
'   .descripcion
'
'Eventos
'
'   _Cambio()       'solo se dispara si el cod es distinto al ultimo
'   _Buscar()       'dispara ANTES de efectuar la busqueda
'                   'util p cambiar la propiedad .SqlBuscar (ej para cambiar rango fechas: where fecha  =  & dtFecha )

Private mUltimoCodigo As Variant
Private mUltimaDescripcion As String
Private mEditaDescripcion As Boolean
Private mCodigoEsString As Boolean
Private mSqlBuscar As String
Private mViaRs As Boolean
'Private mSqlCodFromDes As String
Private mSqlDesFromCod As String

Public Event cambio(codigo) ' As Integer)
Public Event CambioPreview(codigo)
Public Event buscar()
Public Event NoExisteCodigo(codigo)
'
Public Enum uCoDeInvalido
    ucodeMSGyULTIMO
    ucodeMSGyCERO
    ucodeCERO
    ucodeULTIMO_VALOR
End Enum

Private mCodigoInvalido  As uCoDeInvalido

Private mCnx As New ADODB.Connection
'

'***************************************************************************************
Public Property Get CodigoInvalido() As uCoDeInvalido
    On Error Resume Next
    CodigoInvalido = mCodigoInvalido
End Property
Public Property Let CodigoInvalido(que As uCoDeInvalido)
    On Error Resume Next
    mCodigoInvalido = que
    PropertyChanged "CodigoInvalido"
End Property

Public Property Get enabled() As Boolean
    enabled = txtCodigo.enabled
End Property
Public Property Let enabled(como As Boolean)
    txtCodigo.enabled = como
    cmdbuscar.enabled = como
    'Txtdescripcion.enabled = como
    'EditaDescripcion = como
End Property
'
Public Property Get CodigoWidth()
    CodigoWidth = txtCodigo.Width
End Property
Public Property Let CodigoWidth(ancho)
    On Error GoTo fin
    txtCodigo.Width = ancho
    PropertyChanged "CodigoWidth"
    RevisarAnchos
fin:
End Property
'
Public Property Let EditaDescripcion(tf As Boolean)
    txtDescripcion.enabled = tf
    mEditaDescripcion = tf
    txtDescripcion.Locked = Not tf
    PropertyChanged "EditaDescripcion"
    
    txtDescripcion.TabStop = tf
End Property
Public Property Get EditaDescripcion() As Boolean
    EditaDescripcion = mEditaDescripcion
End Property
'
Public Property Get codigo()
    codigo = IIf(mCodigoEsString, txtCodigo, s2n(txtCodigo))
End Property
Public Property Let codigo(cual)
    If IsNull(cual) Then cual = ""
    RevisarCodigo (cual)
    RevisarCambio
End Property
'
Public Property Get DESCRIPCION() As String
    DESCRIPCION = txtDescripcion
End Property
Public Property Let DESCRIPCION(cual As String)
    If EditaDescripcion Then
        txtDescripcion = cual
    End If
End Property

Public Property Let strSqlBuscar(que As String)
    mSqlBuscar = que
End Property

'***************************************************************************************

'Public Sub ini(strSqlGetCodFrom_DesNUMERAL As String, strSqlGet_Des_From_CodNUMERAL, Optional strSqlBuscar As String, Optional bCodigoEsString As Boolean, Optional bPermiteEditarDes As Boolean)
Public Sub ini(strSqlGet_Des_From_CodNUMERAL, Optional strSqlBuscar As String, Optional bCodigoEsString As Boolean, Optional ViaRS As Boolean, Optional queConex As ADODB.Connection) ', Optional bPermiteEditarDes As Boolean)
    mSqlBuscar = strSqlBuscar
    mViaRs = ViaRS
'    mSqlCodFromDes = strSqlGetCodFrom_DesNUMERAL
    mSqlDesFromCod = strSqlGet_Des_From_CodNUMERAL
    mCodigoEsString = bCodigoEsString
     
    cmdbuscar.Visible = mSqlBuscar > ""
    'txtDescripcion.Locked = Not bPermiteEditarDes
     
    mUltimoCodigo = CodNoS(0, "")
    mUltimaDescripcion = ""
    txtCodigo = ""
    txtDescripcion = ""
    
    If queConex Is Nothing Then
        Set queConex = DataEnvironment1.Sistema
    End If
    Set mCnx = queConex
    
End Sub

Public Sub clear()
    RevisarCodigo (IIf(mCodigoEsString, "", 0))
    RevisarCambio
End Sub

'***************************************************************************************
Private Sub cmdBuscar_Click()
    Dim re As String
    RaiseEvent buscar
    
    re = frmBuscar.MostrarSql(mSqlBuscar, , , , , , mViaRs, mCnx)
    If re > "" Then
        codigo = IIf(mCodigoEsString, re, s2n(re))
    End If
End Sub

Private Sub txtcodigo_GotFocus()
    GotFocusPinto txtCodigo
End Sub
Private Sub txtcodigo_LostFocus()
    RevisarCodigo (txtCodigo)
    RevisarCambio
End Sub


Private Function RevisarCodigo(cual As String) As Boolean
    On Error GoTo fin
    Dim s As String
    Dim Vacio As Boolean

    Vacio = CodNoS(s2n(cual) = 0, cual = "")
    If Vacio Then
        txtCodigo = CodNoS(0, "")
        txtDescripcion = ""
        RevisarCodigo = False
    Else
        s = sSinNull(getDesNumeral(cual))
        If s > "" Then
            txtCodigo = cual
            txtDescripcion = s
            RevisarCodigo = True
        Else
            RaiseEvent NoExisteCodigo(cual)
            Select Case mCodigoInvalido
            Case ucodeMSGyULTIMO
                MsgBox "Codigo Inexistente.", vbCritical
                Restaura
            Case ucodeMSGyCERO
                MsgBox "Codigo Inexistente.", vbCritical
                txtCodigo = CodNoS(0, "")
                txtDescripcion = ""
                RevisarCodigo = False
            Case ucodeCERO
                txtCodigo = CodNoS(0, "")
                txtDescripcion = ""
                RevisarCodigo = False
            Case ucodeULTIMO_VALOR
                Restaura
            End Select
        End If
    End If
fin:
End Function
'Private Sub RevisarDescripcion(desc As String)
'    Dim X As Variant
'
''    x = obtenerDeSQL("select codigo from " & mTabla & " where descripcion = '" & desc & "'")
'    X = getCodNumeral(desc)
'    If X > CodNoS(0, "") Or mEditDescripcion Then
'        txtCodigo = X
'        txtdescripcion = desc
'    Else
'        txtCodigo = CodNoS(0, "")
'        Restaura
'    End If
'End Sub
Private Sub RevisarCambio()
    If nCod() <> mUltimoCodigo Then
        RaiseEvent CambioPreview(nCod())
        mUltimoCodigo = nCod()
        mUltimaDescripcion = txtDescripcion
        RaiseEvent cambio(nCod())
    End If
End Sub

Private Sub Restaura()
    On Error Resume Next
    txtCodigo = mUltimoCodigo
    txtDescripcion = mUltimaDescripcion
End Sub

Private Function nCod()
    nCod = CodNoS(s2n(txtCodigo), txtCodigo)
End Function


Private Function CodNoS(n, s)
    CodNoS = IIf(mCodigoEsString, s, n)
End Function

Private Function getDesNumeral(dato As String)
    getDesNumeral = obtenerDeSQL(Replace(mSqlDesFromCod, "###", dato))
End Function

Private Sub txtDescripcion_GotFocus()
    Dim sincod As Boolean
    sincod = IIf(mCodigoEsString, (Trim(txtCodigo) = ""), (s2n(txtCodigo) = 0))

     
    If mEditaDescripcion And sincod Then
        txtDescripcion.Locked = False
    Else
        txtDescripcion.Locked = False 'True
        'SendKeys "{tab}"
    End If
End Sub


Private Sub UserControl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub


Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not enabled Then Exit Sub
    If KeyCode = vbKeyF1 Then cmdBuscar_Click
End Sub

'Private Sub UserControl_LostFocus()
'    CmdBuscar.caption = "?"
'End Sub

Private Sub UserControl_Resize()
    
    txtCodigo.Height = UserControl.Height
    txtDescripcion.Height = UserControl.Height
    RevisarAnchos
End Sub

Private Sub RevisarAnchos()
    On Error GoTo fin
    cmdbuscar.Left = txtCodigo.Width + 20
    txtDescripcion.Left = txtCodigo.Width + 20 + cmdbuscar.Width + 20
    txtDescripcion.Width = UserControl.Width - txtDescripcion.Left
fin:
End Sub

Private Function obtenerDeSQL(SelectCampo As String) As Variant
    On Error GoTo errUFA
    Dim rs As New ADODB.Recordset, ssql  As String
    With rs
        '.Open SelectCampo, DataEnvironment1.Sistema, adOpenDynamic, adLockReadOnly
        .Open SelectCampo, mCnx, adOpenDynamic, adLockReadOnly
        If Not .BOF Then
            .MoveFirst
            obtenerDeSQL = .Fields(0)
        End If
    End With
fin:
    Set rs = Nothing
    Exit Function
errUFA:
    MsgBox "Error interno (ucCoDe : " & SelectCampo & ")"
    Resume fin
End Function


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    mCodigoInvalido = PropBag.ReadProperty("CodigoInvalido", ucodeCERO)
    txtCodigo.Width = PropBag.ReadProperty("CodigoWidth", 1000)
    mEditaDescripcion = PropBag.ReadProperty("EditaDescripcion", False)
    txtDescripcion.Locked = Not mEditaDescripcion
End Sub

Private Sub UserControl_Terminate()
    Set mCnx = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "CodigoWidth", txtCodigo.Width
    PropBag.WriteProperty "EditaDescripcion", mEditaDescripcion, False
    PropBag.WriteProperty "CodigoInvalido", mCodigoInvalido, ucodeCERO
End Sub

'17/12/4
'   mod EditaDescripcion propbag
'25/1/5
'   Opciones cdo usuario pone codigo invalido
'7/7/6
'   buscar via rs
'21-7-6
'   no deja editar descripcion si tiene codigo
'26-7-6
'   fix a: no deja editar descripcion si tiene codigo

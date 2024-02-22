VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LiCodigo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'OJO no compatible previos al 25/10/4 , arreglar  evento Cambio(codigo) 'no longer integer

' mejorar cbo, cuando cambia? verificar kdown y mouseclic?

Option Explicit
' Lito Explicit
' mod 26/4/5
'

'/*
'   Combinacion  txtBox  y  ComboBox    para tablas con
'                codigo     Descripcion
'
' requiere ModuloLi
'*/

Private WithEvents mCbo As ComboBox
Attribute mCbo.VB_VarHelpID = -1
Private WithEvents mTxt As TextBox
Attribute mTxt.VB_VarHelpID = -1
Private WithEvents mCmd As CommandButton
Attribute mCmd.VB_VarHelpID = -1

Private mEnabled As Boolean
Private mUltimoCodigo As Variant
Private mUltimaDescripcion As String
Private mTabla As String
Private mEditDescripcion As Boolean
Private mCodigoEsString As Boolean
Private mWhere As String
Private mViaRs As Boolean

Public Event cambio(codigo) ' As Integer)
'

' /******************** Publico *************************************
 
Public Property Let EditaDescripcion(tf As Boolean)
    mEditDescripcion = tf
End Property
Public Property Get EditaDescripcion() As Boolean
    EditaDescripcion = mEditDescripcion
End Property
 
Public Property Get enabled() As Boolean
    enabled = mEnabled
End Property
Public Property Let enabled(tf As Boolean)
    On Error Resume Next
    mEnabled = tf
    mCbo.enabled = tf
    mTxt.enabled = tf
    mCmd.enabled = tf
End Property


Public Property Get codigo()
    codigo = IIf(mCodigoEsString, mTxt, s2n(mTxt))
End Property
Public Property Let codigo(cual)
    RevisarCodigo (cual)
    RevisarCambio
End Property


Public Property Get DESCRIPCION() As String
    DESCRIPCION = mCbo.Text
End Property
Public Property Let DESCRIPCION(cual As String)
    RevisarDescripcion (cual)
    RevisarCambio
End Property


Public Sub init(cmb As ComboBox, txt As TextBox, sTabla As String, Optional PermiteEditarDescripcion As Boolean, Optional CodigoEsString As Boolean, Optional cmd As CommandButton, Optional sWhere As String, Optional ViaRS As Boolean)
    Dim sw As String
    Set mCbo = cmb
    Set mTxt = txt
    Set mCmd = cmd
    
    mViaRs = ViaRS
    mTabla = sTabla
    mEnabled = True
    mCodigoEsString = CodigoEsString
    
    mUltimoCodigo = CodNoS(0, "")
    mUltimaDescripcion = ""
    
    mEditDescripcion = PermiteEditarDescripcion
    
    mWhere = sWhere
    If mWhere > "" Then sw = " Where " & mWhere
    
    comboSql mCbo, "select descripcion from " & mTabla & sw & " order by descripcion "
    RevisarDescripcion mCbo
    RevisarCambio
End Sub


' ******************** Privado *****************************************

Private Sub mCbo_LostFocus()
    RevisarDescripcion (mCbo.Text)
    RevisarCambio
End Sub
Private Sub mCmd_Click()
    Dim re As String, sw As String
    If mWhere > "" Then sw = " where " & mWhere
    're = frmBuscar.MostrarCodigoDescripcionActivo(mTabla)
    re = frmBuscar.MostrarSql("select codigo as [Codigo   _], descripcion as [Descripcion                         _] from " & mTabla & "  " & sw & "  order by codigo ", , , , , , mViaRs)
    If re > "" Then
        codigo = IIf(mCodigoEsString, re, s2n(re))
    End If
End Sub
Private Sub mTxt_GotFocus()
    GotFocusPinto mTxt
End Sub
Private Sub mTxt_LostFocus()
    RevisarCodigo (mTxt)
    RevisarCambio
End Sub

Private Function RevisarCodigo(cual As String) As Boolean
    On Error GoTo UfaRevisar
    Dim s As String, sqlN As String, sqlS As String, Vacio As Boolean, sAnd As String
    
    sAnd = IIf(mWhere > "", " AND " & mWhere, "")
    
    sqlN = "select descripcion from " & mTabla & " where codigo = " & x2s(cual) & sAnd
    sqlS = "select descripcion from " & mTabla & " where codigo = '" & cual & "'" & sAnd
    
    s = obtenerDeSQL(CodNoS(sqlN, sqlS))
    Vacio = DatoSN(cual = "", s2n(cual) = 0) 'IIf(mCodigoEsString, (cual = ""), (s2n(cual) = 0))
    If Vacio Then
        mTxt = DatoSN("", 0)
        mCbo.ListIndex = -1 ' = ""
        RevisarCodigo = False
    ElseIf s > "" Then
        mTxt = cual
        ComboSetIndex mCbo, s, True
        RevisarCodigo = True
    Else
        Restaura
    End If
fin:
    Exit Function
UfaRevisar:
    Restaura
    Resume fin
End Function

Private Sub RevisarDescripcion(desc)
    On Error Resume Next
    Dim x As Variant
   
    x = obtenerDeSQL("select codigo from " & mTabla & " where (descripcion = '" & desc & "') and " & mWhere)
    If x > CodNoS(0, "") Or mEditDescripcion Then
        mTxt = x
        mCbo.Text = desc
    Else
        mTxt = CodNoS(0, "")
        Restaura
    End If
End Sub
Private Sub RevisarCambio()
    If nCod() <> mUltimoCodigo Then
        mUltimoCodigo = nCod()
        mUltimaDescripcion = mCbo.Text
        RaiseEvent cambio(nCod())
    End If
End Sub

Private Sub Restaura()
    On Error Resume Next
    mTxt = mUltimoCodigo
    mCbo = mUltimaDescripcion
End Sub

Private Function nCod()
    nCod = CodNoS(s2n(mTxt), mTxt)
End Function
Private Function CodNoS(n, s)
    CodNoS = IIf(mCodigoEsString, s, n)
End Function

Private Function ComboSetIndex(Combo As ComboBox, Valor As Variant, Optional EnList As Boolean = False) As Long
    Dim i As Long
    
    If val(Valor) > -1 Then
        If Not EnList Then
            While Combo.ItemData(i) <> Valor And i < Combo.ListCount
                i = i + 1
            Wend
        Else
            While Combo.List(i) <> Valor And i < Combo.ListCount
                i = i + 1
            Wend
        End If
        ComboSetIndex = IIf(i = Combo.ListCount, -1, i)
    Else
        ComboSetIndex = -1

    End If
    Combo.ListIndex = ComboSetIndex
End Function

Private Function DatoSN(datoS, datoN)
    DatoSN = IIf(mCodigoEsString, datoS, datoN)
End Function

' 23/8/4 start
' 24/8/4 end, probado
' 26/8/4 acepta codigo string, (avisarle con parametro en init())
' 7/10/4 fix cambio x focus txtbox
'25/10/4 fix evento cambio: codigo puede ser string: INCOMPATIBLE con previas versiones
'28/10/4 optional cmd para frmBuscar
'        Permito codigo 0 como borrado
'        GotFocusPinto
'11/11/4 debo pasar parametro where pal combo
'        frmBuscar compatible access
'        fix enabled
'15-11-4 seteo cbo.listindex en lugar de cbo.text
'22/11/4 fix limpiar control codigo = ""
'26/4/5 fix codigo numerico con coma

' *** Notas ****
' no me anime a hacer UC, seria facil.
' para uc, pondria .Orientacion Vert u Horz, uno arriba de otro o al lado
' V
' controles.h = uc.h /2
' cmb.w = uc.w
' txt.w = ?  lo expongo
' H
' controles.h = uc.h
' txt.w = ?  lo expongo, cmb.w = uc.w -txt.w - 60


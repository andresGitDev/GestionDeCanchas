VERSION 5.00
Begin VB.UserControl uTextConAyuda 
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4725
   ScaleHeight     =   315
   ScaleWidth      =   4725
   Begin VB.TextBox txtDato 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "ver que tiene oculto un menu para acceder con el popupmenu"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnuMenu2 
         Caption         =   "Menu2"
         Index           =   0
      End
   End
End
Attribute VB_Name = "uTextConAyuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub mnuMenu2_Click(Index As Integer)
    Dim str As String
    
    str = mnuMenu2(Index).caption
    If str <> "" Then
        txtDato = str
    End If
End Sub

Private Sub txtDato_Change()
    Dim rs As New ADODB.Recordset
    Dim sql As String
    Dim i As Long
    Dim j As Long
    
    If Len(txtDato.Text) >= 3 Then
        sql = "select descriclie from remitoventa where anulado=0 and descriClie like '%" & Trim(txtDato) & "%'"
        rs.Open sql, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        
        j = 1
        While j < mnuMenu2.Count
            Unload mnuMenu2(j)
            j = j + 1
        Wend
        i = 1
        If (rs.EOF = True And rs.BOF) Or IsNull(rs!descriclie) Or IsEmpty(rs!descriclie) Then
            mnuMenu2(0).caption = " "
        Else
            mnuMenu2(0).caption = rs!descriclie
            rs.MoveNext
        End If
        While Not rs.EOF
            'mnuMenu(0).caption = "Menu"
            'Load mnuMenu(1)
            'mnuMenu(1).caption = "Salir"
            
            Load mnuMenu2(i)
            mnuMenu2(i).caption = rs!descriclie
            
            rs.MoveNext
        Wend
        
        PopupMenu mnuMenu, , 1500, 1350
    End If
End Sub

'***********************************************************************************


Public Property Get enabled() As Boolean
    enabled = txtDato.enabled
End Property
Public Property Let enabled(como As Boolean)
    txtDato.enabled = como
End Property
'
Public Property Get CodigoWidth()
    CodigoWidth = txtDato.Width
End Property
Public Property Let CodigoWidth(ancho)
    On Error GoTo fin
    txtDato.Width = ancho
    PropertyChanged "CodigoWidth"
    RevisarAnchos
fin:
End Property
'
Public Property Let EditaDescripcion(tf As Boolean)
    txtDato.enabled = tf
    mEditaDescripcion = tf
    txtDato.Locked = Not tf
    PropertyChanged "EditaDescripcion"
    
    txtDato.TabStop = tf
End Property
Public Property Get EditaDescripcion() As Boolean
    EditaDescripcion = mEditaDescripcion
End Property

Public Property Get DESCRIPCION() As String
    DESCRIPCION = txtDato
End Property
Public Property Let DESCRIPCION(cual As String)
    If EditaDescripcion Then
        txtDato = cual
    End If
End Property

Public Property Let strSqlBuscar(que As String)
    mSqlBuscar = que
End Property

'***************************************************************************************

Public Sub ini(strSqlGet_Des_From_CodNUMERAL, Optional strSqlBuscar As String, Optional ViaRS As Boolean, Optional queConex As ADODB.Connection) ', Optional bPermiteEditarDes As Boolean)
    mSqlBuscar = strSqlBuscar
    mViaRs = ViaRS
    mSqlDesFromCod = strSqlGet_Des_From_CodNUMERAL
    
    cmdbuscar.Visible = mSqlBuscar > ""
     
    'mUltimoCodigo = CodNoS(0, "")
    mUltimaDescripcion = ""
    txtDato = ""
    
    If queConex Is Nothing Then
        Set queConex = DataEnvironment1.Sistema
    End If
    Set mCnx = queConex
    
End Sub

Public Sub clear()
    txtDato = ""
End Sub

'***************************************************************************************

Private Sub txtdato_GotFocus()
    GotFocusPinto txtDato
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub UserControl_Resize()
    
    txtDato.Height = UserControl.Height
    RevisarAnchos
End Sub

Private Sub RevisarAnchos()
    On Error GoTo fin
    txtDato.Width = UserControl.Width
fin:
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    txtDato.Width = PropBag.ReadProperty("CodigoWidth", 1000)
    mEditaDescripcion = PropBag.ReadProperty("EditaDescripcion", False)
    txtDato.Locked = Not mEditaDescripcion
End Sub

Private Sub UserControl_Terminate()
    Set mCnx = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "CodigoWidth", txtDato.Width
    PropBag.WriteProperty "EditaDescripcion", mEditaDescripcion, False
    PropBag.WriteProperty "CodigoInvalido", mCodigoInvalido, ucodeCERO
End Sub

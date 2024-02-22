VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmRemitosCompraPendientes 
   Appearance      =   0  'Flat
   Caption         =   "Remitos Proveedor Pendientes de Facturar"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8415
   Icon            =   "frmRemitosCompraPendientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraFoot 
      Height          =   1140
      Left            =   60
      TabIndex        =   7
      Top             =   6210
      Width           =   8265
      Begin VB.CommandButton cmdCarga 
         Caption         =   "A Factura"
         Height          =   840
         Left            =   6330
         Picture         =   "frmRemitosCompraPendientes.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   210
         Width           =   930
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   855
         Left            =   7290
         Picture         =   "frmRemitosCompraPendientes.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   195
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "* ParaItems individuales, cargar solamente CANTIDAD y PRECIO."
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   4935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "* Para marcar todos los items de un remito, marque el campo con la 'X'"
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   180
         Width           =   5115
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Height          =   4935
      Left            =   60
      TabIndex        =   3
      Top             =   840
      Width           =   8235
      _cx             =   14526
      _cy             =   8705
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Frame fra 
      Caption         =   "Elija el proveedor"
      Height          =   795
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8310
      Begin VB.TextBox txtCodProv 
         Height          =   315
         Left            =   1950
         TabIndex        =   0
         Top             =   270
         Width           =   1455
      End
      Begin VB.CommandButton cmdProv 
         Height          =   465
         Left            =   1425
         Picture         =   "frmRemitosCompraPendientes.frx":1A5E
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   210
         Width           =   510
      End
      Begin VB.ComboBox cboProveedor 
         Height          =   315
         Left            =   3450
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   270
         Width           =   4635
      End
   End
   Begin VB.Label txtTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   6870
      TabIndex        =   6
      Top             =   5880
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "Monto Acumulado para la Factura : "
      Height          =   285
      Left            =   4350
      TabIndex        =   5
      Top             =   5925
      Width           =   2730
   End
End
Attribute VB_Name = "frmRemitosCompraPendientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'falta:
'   Verificar q cant sea <= pendiente
'   Quizas deba hacer cargar el monto y despues ver si coincide
'   Poner boton p que reabra este form para modificar pero sin cargarlo de nuevo
'
' Usar este mismo form, 3 grillas  rem  , fac    y rel abajo (filtro fecha?) (ultimos 20?)
'    y otra func mostrar o poner parametro.
'

'19/11/4
Option Explicit

Private rAsiento As Collection

Private mOK As Boolean
Private WithEvents g As LiGrilla
Attribute g.VB_VarHelpID = -1
Private WithEvents prov As LiCodigo
Attribute prov.VB_VarHelpID = -1


Private gPROV As Long
Private gprod As Long
Private gPEND As Long
Private gPREU As Long
Private gREMI As Long
Private gCHKR As Long
Private gCANT As Long
Private gPREA As Long
Private gITEM As Long
Private gDESC As Long
'Private gTOTR As Long
'Private gTOTI As Long
'Private gCHKI As Long
'Private gFact As Long
'


Private Sub cmdCancelar_Click()
    mOK = False
    Set rAsiento = New Collection
    Me.Hide
End Sub
Private Sub cmdCarga_Click()
    'If s2n(txtImporteFactura) = 0 Then
    '    che "No se cargo el Importe de la Factura"
    'Else
    Dim i As Long, p As Long, v As Double, d
    If s2n(txttotal) = 0 Then
        che "No se ha asignado monto"
'    ElseIf s2n(s2n(txtImporteFactura) - s2n(txtTotal)) <> 0 Then
'        che "El monto acumulado no coincide con el de la factura"
    ElseIf prov.codigo = 0 Then
        che "Falta cargar proveedor"
    ElseIf aMedias() Then
        'ya te avise...
    Else ' sin objecion
    p = 0
    Set rAsiento = New Collection
        For i = 1 To Grilla.rows - 1
            If Grilla.TextMatrix(i, gCHKR) = "X" Then
                d = obtenerDeSQL("select TIENE_CUENTA,CUENTA from producto where codigo=" & ssTexto(Grilla.TextMatrix(i, gprod)))
                If d(0) = 1 Then
                    v = s2n(Grilla.TextMatrix(i, gCANT) * Grilla.TextMatrix(i, gPREA)) 'cantidad * precio
                    If v > 0 Then
                        rAsiento.Add Grilla.TextMatrix(i, gprod), "CodProd" & p
                        rAsiento.Add d(1), "Cuenta" & p
                        rAsiento.Add v, "Valor" & p
                    
                     
                        p = p + 1
                    End If
                End If
            End If
        Next
    
        mOK = True
        Me.Hide
    End If
End Sub

Private Function aMedias() As Boolean
    Dim i As Long
    aMedias = False
    For i = 1 To g.rows - 1
        If s2n(g.tx(i, gCANT)) <> 0 And s2n(g.tx(i, gPREA)) = 0 Then
'            che "Item con cantidad y sin precio"
'            aMedias = True
'            Exit Function
        ElseIf s2n(g.tx(i, gCANT)) = 0 And s2n(g.tx(i, gPREA)) <> 0 Then
            che "Item con precio y sin cantidad "
            aMedias = True
            Exit Function
        End If
    Next i
End Function

Private Sub cmdProv_Click()
    If frmBuscar.MostrarSql("SELECT DISTINCT Prov.codigo as [ Codigo ], Prov.descripcion as [ Proveedor           ] FROM RemitoCompraDetalle INNER JOIN  RemitoCompra ON RemitoCompraDetalle.CodigoRemito = RemitoCompra.codigo INNER JOIN Prov ON Prov.codigo = RemitoCompra.Proveedor Where (RemitoCompraDetalle.cantidad_a_facturar > 0)", , "Proveedores con Remitos Pendientes de Factura") = "" Then Exit Sub
    prov.codigo = s2n(frmBuscar.resultado())
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True ', True, True
End Sub


Public Function mostrar() As Variant    '(Optional Proveedor As Integer)
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    If g Is Nothing Then inigrilla
    Set prov = New LiCodigo
    
    'prov.init cboProveedor, txtcodprov, "prov", False, False, cmdprov, "activo = 1 "
    prov.init cboProveedor, TxtCodProv, "prov", False, False, , "activo = 1 "
    
    prov.codigo = 0

    Me.Show vbModal

fin:
    'mostrar = total()
    'mostrar = Array(total(), ProdAsiento)
    Set mostrar = rAsiento
    Exit Function
ufaErr:
    ufa "", "mostrar" & Me.Name ', Err
    Resume fin
End Function
Public Property Get Total() As Double
    Total = IIf(mOK, s2n(txttotal), 0)
End Property
Public Property Get ProveedorCod() As Long
    If Not prov Is Nothing Then ProveedorCod = IIf(mOK, s2n(prov.codigo), 0)
End Property
Public Property Get ProveedorNombre() As String
    If Not prov Is Nothing Then ProveedorNombre = prov.DESCRIPCION
End Property


Private Sub ActualizarGrilla()
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    Dim rs As New ADODB.Recordset, ssql As String, i As Long
   
    g.Borrar
    If prov.codigo = 0 Then Exit Sub
   
    ssql = "SELECT isnull(ItemOrdenCompra.costo,0) AS TotOC , " _
        & " RemitoCompra.Proveedor, RemitoCompra.NroRemito, " _
        & " RemitoCompraDetalle.cantidad_a_facturar AS pend, RemitoCompraDetalle.producto, " _
        & " RemitoCompraDetalle.codigo " _
        & " FROM dbo.RemitoCompra " _
        & " INNER JOIN RemitoCompraDetalle ON RemitoCompra.Codigo = RemitoCompraDetalle.CodigoRemito " _
        & " LEFT OUTER JOIN ItemOrdenCompra ON RemitoCompraDetalle.ordencompra = ItemOrdenCompra.ordencompra " _
        & " AND RemitoCompraDetalle.producto = ItemOrdenCompra.producto " _
        & " WHERE (RemitoCompra.activo = 1) AND (RemitoCompraDetalle.cantidad_a_facturar > 0) " _
        & " and proveedor = " & prov.codigo

    rs.Open ssql, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    If rs.EOF Then
'        MsgBox " No figuran remitos saldo pendiente"
    Else
        While Not rs.EOF
            i = g.addRow()
            g.tx i, gPROV, rs!Proveedor
            g.tx i, gprod, rs!Producto
            g.tx i, gDESC, obtenerDeSQL("select descripcion from producto where codigo = '" & rs!Producto & "'")
            g.tx i, gPEND, s2n(rs!pend)
            g.tx i, gPREU, s2n(IIf(s2n(rs!totoc) = 0, obtenerDatoS("producto", rs!Producto, "CostoBase"), s2n(rs!totoc)))  '/ s2n(rs!pend))
            g.tx i, gREMI, rs!NroRemito
            g.tx i, gCHKR, ""
            g.tx i, gCANT, ""
            g.tx i, gPREA, ""
            g.tx i, gITEM, rs!codigo
            rs.MoveNext
        Wend
    End If
fin:
    Set rs = Nothing
    Exit Sub
ufaErr:
    ufa "Err cargando remitos", Me.Name ', Err
    Resume fin
End Sub

Public Function Item(Index As Long) As Long
    Dim i As Long, C As Long
    If Not mOK Then Exit Function
    
    C = 0
    For i = 1 To g.rows - 1
        If s2n(g.tx(i, gCANT)) > 0 Then
            C = C + 1
            If C = Index Then
                Item = g.tx(i, gITEM)
                Exit Function
            End If
        End If
    Next i
End Function
Public Function cant(Index As Long) As Double
    Dim i As Long, C As Long
    If Not mOK Then Exit Function
    C = 0
    For i = 1 To g.rows
        If s2n(g.tx(i, gCANT)) > 0 Then
            C = C + 1
            If C = Index Then
                cant = g.tx(i, gCANT)
                Exit Function
            End If
        End If
    Next i
End Function
Public Function PrecioU(Index As Long) As Double
    Dim i As Long, C As Long
    If Not mOK Then Exit Function
    C = 0
    For i = 1 To g.rows
        If s2n(g.tx(i, gCANT)) > 0 Then
            C = C + 1
            If C = Index Then
                PrecioU = s2n(g.tx(i, gPREA))
                Exit Function
            End If
        End If
    Next i
End Function

Private Sub Form_Load()
Set rAsiento = New Collection
End Sub

'Private Sub Form_Load()
'    CentrarMe Me
'End Sub

Private Sub Form_Resize()
'    encajar grilla, Me, fra.Height + 120, 120, 120 + fraFoot.Height, 120 + txtTotal.Width
'    encajar fraFoot, Me, , 60, 0
'    encajar txtTotal, Me, 1860, , , 120
'    encajar Label1, Me, 840, , , 120
'    encajar cmdCancelar, Me, , , 60, 1500
'    encajar cmdCarga, Me, , , 60, 3000
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set g = Nothing
    Set prov = Nothing
End Sub
Private Sub inigrilla()
    Set g = New LiGrilla
    g.init Grilla
    gPROV = g.AddCol(" Proveedor", "H")
    gprod = g.AddCol(" Producto        ")
    gDESC = g.AddCol(" Descripcion                           ")
    gPEND = g.AddCol(" Pendientes ")
    gPREU = g.AddCol(" Precio  ", "9")
    gREMI = g.AddCol(" Remito ")
            g.AddCol "  "
    gCHKR = g.AddCol(" X ")
    gCANT = g.AddCol("Cantidad  ", "N")
    gPREA = g.AddCol("Precio Unitario", "N")
    gITEM = g.AddCol(" item ", "H")
    g.Borrar
 
'    gTOTI = g.AddCol(" Total Item ")
'    gFact = g.AddCol(" Facturar ", "N")
    
End Sub


Private Sub g_cambio(ByVal Row As Long, ByVal Col As Long, txt As String)
    Dim i As Long, t As Double
    
    'cambia cant, modifico precio
    If Col = gCANT And s2n(g.tx(Row, gCANT)) > 0 And s2n(g.tx(Row, gPREA)) > 0 Then g.tx Row, gPREA, s2n(g.tx(Row, gPREU))
    If Col = gCANT And s2n(g.tx(Row, gCANT)) = 0 Then g.tx Row, gPREA, ""
    
    'total p'factura
    t = 0
    For i = 1 To g.rows - 1
        t = t + s2n(g.tx(i, gCANT)) * s2n(g.tx(i, gPREA))
    Next i
    txttotal = s2n(t)
End Sub

Private Sub g_Click()
    Dim i As Long, x As Boolean

    If g.Row > 0 Then
        If g.Col = gCHKR Then
            x = (g.Text = "")
            g.Text = IIf(x, "X", "")
            
            For i = 1 To g.rows - 1
                If g.tx(i, gREMI) = g.tx(g.Row, gREMI) Then
                    g.tx i, gCANT, IIf(x, g.tx(i, gPEND), "")
                    g.tx i, gPREA, IIf(x, g.tx(i, gPREU), "")
                    g.tx i, gCHKR, g.Text
                End If
            Next i
        End If
        If g.Col = gCANT Then
            If g.Text = "" Then g.Text = g.tx(g.Row, gPEND)
        End If
        If g.Col = gPREA Then
            If g.Text = "" Then g.Text = g.tx(g.Row, gPREU)
        End If
    End If
End Sub

Private Sub Prov_cambio(codigo As Variant)
    ActualizarGrilla
End Sub

'15/11/4    s2n en precioU()
'19/11/4 adapt licodigo, +cmd, +where
'24/11/4    verificaciones,
'       cmdProv lo saque d liCodigo, lo hice filtrado x pendientes
'15/4/5 tamaño


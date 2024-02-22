VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPedidosCancelacion 
   Caption         =   "Cancelacion de Pedidos"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8640
   Icon            =   "frmPedidosCancelacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   8640
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMain 
      Height          =   6195
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8595
      Begin VB.OptionButton optPropio 
         Alignment       =   1  'Right Justify
         Caption         =   "Cliente"
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1440
         Width           =   855
      End
      Begin VB.OptionButton optPropio 
         Alignment       =   1  'Right Justify
         Caption         =   "Propio"
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   7440
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1140
         Value           =   -1  'True
         Width           =   855
      End
      Begin Gestion.ucCoDe uPedido 
         Height          =   315
         Left            =   840
         TabIndex        =   1
         Top             =   1260
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   556
         CodigoWidth     =   1000
      End
      Begin Gestion.ucFecha uFecha 
         Height          =   315
         Left            =   7440
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   720
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         FechaInit       =   0
      End
      Begin Gestion.ucCoDe uCliente 
         Height          =   315
         Left            =   840
         TabIndex        =   0
         Top             =   780
         Width           =   5655
         _ExtentX        =   8705
         _ExtentY        =   556
         CodigoWidth     =   1000
      End
      Begin TabDlg.SSTab tabMain 
         Height          =   4275
         Left            =   60
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1800
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   7541
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabHeight       =   520
         BackColor       =   -2147483644
         ForeColor       =   -2147483630
         TabCaption(0)   =   "Items"
         TabPicture(0)   =   "frmPedidosCancelacion.frx":08CA
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "grillaCancelacion"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Series"
         TabPicture(1)   =   "frmPedidosCancelacion.frx":08E6
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "lblErrorSeries"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label10"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "grillaSeries"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "chkSinSeries"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "cmdLlenaSerie"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).ControlCount=   5
         Begin VB.CommandButton cmdLlenaSerie 
            Caption         =   "Llenar Serie"
            Height          =   315
            Left            =   3360
            TabIndex        =   12
            ToolTipText     =   "Seleccione filas a llenar"
            Top             =   540
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CheckBox chkSinSeries 
            Alignment       =   1  'Right Justify
            Caption         =   "Sin Series"
            Height          =   315
            Left            =   7140
            TabIndex        =   11
            Top             =   1020
            Width           =   1095
         End
         Begin VSFlex7LCtl.VSFlexGrid grillaSeries 
            Height          =   2895
            Left            =   180
            TabIndex        =   10
            Top             =   1020
            Width           =   6795
            _cx             =   11986
            _cy             =   5106
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
         Begin VSFlex7LCtl.VSFlexGrid grillaCancelacion 
            Height          =   3615
            Left            =   -74820
            TabIndex        =   19
            Top             =   480
            Width           =   7875
            _cx             =   13891
            _cy             =   6376
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
         Begin VB.Label Label10 
            Caption         =   "Puede hacer 'Doble Clic' en el campo  Nro.Serie"
            Height          =   495
            Left            =   600
            TabIndex        =   14
            Top             =   540
            Width           =   2235
         End
         Begin VB.Label lblErrorSeries 
            Caption         =   "--------------"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   5040
            TabIndex        =   13
            Top             =   540
            Visible         =   0   'False
            Width           =   2475
         End
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Relacion es:  **MovimientoInterno-Numero**"
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   3120
         TabIndex        =   23
         Top             =   300
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "Numero :"
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblId 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   315
         Left            =   7440
         TabIndex        =   21
         Top             =   300
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label lblNumero 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   315
         Left            =   900
         TabIndex        =   20
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha :"
         Height          =   315
         Index           =   3
         Left            =   6660
         TabIndex        =   18
         Top             =   780
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo :"
         Height          =   255
         Index           =   2
         Left            =   6720
         TabIndex        =   17
         Top             =   1200
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Pedido:"
         Height          =   315
         Index           =   1
         Left            =   180
         TabIndex        =   6
         Top             =   1260
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente :"
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   5
         Top             =   840
         Width           =   735
      End
   End
   Begin Gestion.ucBotonera uMenu 
      Height          =   1530
      Left            =   0
      TabIndex        =   3
      Top             =   6315
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   2699
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   ""
      Begin Gestion.ucFecha uFechaBuscar 
         Height          =   315
         Left            =   6300
         TabIndex        =   7
         Top             =   0
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         FechaInit       =   0
      End
      Begin VB.Label Label2 
         Caption         =   "Buscar Pedidos Desde"
         Height          =   195
         Left            =   4500
         TabIndex        =   8
         Top             =   60
         Width           =   1755
      End
   End
End
Attribute VB_Name = "frmPedidosCancelacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   Relacion es R.MovimientoInterno = D.Numero
'   NroComprobante puede ser vacio, estar repetido, se usa solo a veces



'FALTA
'OBLIGAR A Q PONGA SERIES o sinseries
'ayuda series 2clic
'TipoComprobante_CANCELACIONPEDIDO


Option Explicit

Private Const CONCEPTO_CANCELACION = 50
Private Const CTE_SERIE_AGREGAR = "Registrar"

Private gITEM As Long
Private gprod As Long
Private gDESC As Long
Private gCANT As Long
Private gSALD As Long
Private gPREC As Long
Private gFORM As Long

Private g3ITEM As Long
Private g3PROD As Long
Private g3NSER As Long
Private g3HIDD As Long
Private g3ALTA As Long

Private WithEvents g As LiGrilla
Attribute g.VB_VarHelpID = -1
Private WithEvents g3 As LiGrilla
Attribute g3.VB_VarHelpID = -1
'

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True, True
End Sub

Private Sub Form_Load()
    inigrilla
    uMenu.init True, True, False, True, True, , , True
    uCliente.ini "select descripcion from clientes where codigo = '###'", "select codigo as [ Codigo ], descripcion as [ Nombre                        ] from clientes where activo = 1", False
    uCliente.enabled = False
    uPedido.ini "select descripcion from clientes inner join pedidos_clientes on pedidos_Clientes.cliente = Clientes.codigo where numero = ###", "select numero, fecha, cliente from pedidos_clientes where activo = 1 order by numero desc "
    tabMain.Tab = 0
End Sub

Private Sub Form_Unload(cancel As Integer)
    Set g = Nothing
End Sub

Private Sub g_Validar(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
    'cancela edicion si cant ingresada > saldo item
    If (Col = gCANT) And (s2n(g.EditText) > s2n(g.tx(Row, gSALD))) Then
        che "cantidad no puede ser mayor que el saldo"
        cancel = True
    End If
End Sub


Private Sub grillaSeries_dblClick()
    Dim r As Long, prod As String, resu As String
    
    r = g3.Row
    If r < 1 Then Exit Sub
    prod = VerProductoMio(g3.tx(r, g3PROD), Propio())
    
    resu = Buscar_SeriesEnStock(prod)
    If resu > "" Then grillaSeries.TextMatrix(r, g3NSER) = resu
End Sub


Private Sub tabMain_Click(PreviousTab As Integer)
    If tabMain.Tab = 1 And PreviousTab <> 1 Then LlenoGrillaSeries
End Sub


Private Sub LlenoGrillaSeries()
    Dim i As Long, j As Long, prod As String, cant As Long, cons

        'borrosinserie
        i = 1
        While i < g3.rows
            If g3.tx(i, g3NSER) = "" Then
                g3.delRow i
                i = i - 1
            End If
            i = i + 1
        Wend

        'borro marcas
        For i = 1 To g3.rows - 1
            grillaSeries.TextMatrix(i, g3HIDD) = ""
        Next i

        'marco o agrego en grilla series
        For i = 1 To g.rows - 1
            prod = Trim(g.tx(i, gprod))
            cant = s2n(g.tx(i, gCANT))
'            cons = s2n(g.tx(i, gCONS))
            If ProductoConSerie(prod, Propio()) Then
                For j = 1 To cant
                    If marcoG3(prod, cons) Then
 '                       cons = cons - 1
                    End If
                Next j
            End If
        Next i

        'borro no marcadas
        i = 1
        While i < g3.rows
            If g3.tx(i, g3HIDD) = "" Then
                g3.delRow i
                i = i - 1
            End If
            i = i + 1
        Wend

'    Dim i As Long, i3 As Long, j As Integer, prod As String, cant As Integer
'
'    g3.Borrar
'    'marco o agrego en grilla series
'    For i = 1 To g.rows - 1
'        prod = Trim(g.tx(i, gprod))
'        cant = s2n(g.tx(i, gCANT))
'        If ProductoConSerie(prod, Propio()) Then
'            For j = 1 To cant
'                i3 = g3.addRow()
'                g3.tx i3, g3PROD, prod
'            Next j
'        End If
'    Next i
End Sub
Private Function marcoG3(codi, ByVal cons) As Boolean '
    Dim i As Long
    
    For i = 1 To g3.rows - 1
        If g3.tx(i, g3PROD) = codi And g3.tx(i, g3HIDD) = "" Then
            grillaSeries.TextMatrix(i, g3HIDD) = "X"
            Exit Function
        End If
    Next i
    i = g3.addRow()
    grillaSeries.TextMatrix(i, g3PROD) = codi
    grillaSeries.TextMatrix(i, g3HIDD) = "X"
    'grillaSeries.TextMatrix(i, g3CONS) = IIf(cons > 0, "-1", "0")
    
    'marcoG3 = (g3.tx(i, g3CONS) = "-1")
End Function


Private Sub uCliente_cambio(codigo As Variant)
    If ClientePedido(uPedido.codigo) <> codigo Then
        uPedido.clear
        g.Borrar
    End If
End Sub


Private Sub uMenu_BuscarYa(que As Variant)
    If CargaCancelacion(CLng(que)) Then uMenu.BuscarOK
End Sub

Private Sub uMenu_eliminar()
    If MODO_ON_ERROR_ABM_ON Then On Error GoTo UFAalta
    
    Dim rs As New ADODB.Recordset
        
'    lblId = idDifStock
    If confirma("Elimina Cancelacion Numero" & vbCrLf & lblNumero) Then
        '*******************************************************
        DE_BeginTrans
            
            If lblNumero > 0 Then
                'actualiz pedido
                If obtenerDeSQL("select nropedido from remitodiferenciastock where movimientointerno=" & lblNumero) > 0 Then ' si 0, es componente de formula virtual, no figura en pedido
                    rs.Open "select * from remitodiferenciastock r inner join itemremitodiferenciastock i on i.numero=r.movimientointerno where movimientointerno = " & lblNumero.caption, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                    While Not rs.EOF
                        If rs!NroPedido > 0 Then
                            DataEnvironment1.Sistema.Execute "update itemPedidoCliente set facturar = facturar + " & x2s(rs!cantidad) & ", saldo = saldo + " & x2s(rs!cantidad) & " where pedido = " & rs!NroPedido & " and producto='" & rs!producto & "'"
                        End If
                        rs.MoveNext
                    Wend
                End If
                rs.MoveFirst
                While Not rs.EOF
                    If IsNull(rs!producto) Or IsNull(rs!numero) Then
                    Else
                        DataEnvironment1.Sistema.Execute "update series set activo=0,fecha_baja=" & ssFecha(Date) & ",usuario_baja=" & UsuarioActual & " where producto='" & Trim(rs!producto) & "' and nrocomprobante = " & rs!numero & " and essalida=1"
                    End If
                    rs.MoveNext
                Wend
                
                rs.MoveFirst
                'resto stock, grabo item
                If obtenerDeSQL("select nropedido from remitodiferenciastock where movimientointerno=" & lblNumero) = 0 Then ' producto virtual no baja stock
                    'DataEnvironment1.dbo_ITEMDIFSTOCK "F", prod, cant, idDifStock, prec
                    DataEnvironment1.Sistema.Execute "delete from itemremitodiferenciastock where numero = " & lblNumero
                    DataEnvironment1.Sistema.Execute "delete from remitodiferenciastock where movimientointerno= " & lblNumero
                Else
                    'DataEnvironment1.dbo_ITEMDIFSTOCK "R", prod, cant, idDifStock, prec
                    While Not rs.EOF
                        DataEnvironment1.Sistema.Execute "Update producto SET EXISTENCIA= EXISTENCIA + " & x2s(rs!cantidad) & " WHERE CODIGO='" & Trim(rs!producto) & "'"
                        rs.MoveNext
                    Wend
                    rs.MoveFirst
                    DataEnvironment1.Sistema.Execute "delete from itemremitodiferenciastock where numero = " & lblNumero
                    DataEnvironment1.Sistema.Execute "delete from remitodiferenciastock where movimientointerno= " & lblNumero
                End If
            End If
            
'            For i = 1 To g3.rows - 1 'series
'                prod = VerProductoMio(g3.tx(i, g3PROD), Propio())
'                Serie = g3.tx(i, g3NSER)
'                If Serie <> "" Then
'                    'DataEnvironment1.dbo_SERIE "A", 0, prod, serie, TipoComprobante_CANCELACIONPEDIDO, idDifStock, 0, 0, "", 0, Date, UsuarioActual(), 0, 0
'                    DataEnvironment1.dbo_abmSERIEs "A", 0, prod, Serie, TipoComprobante_CANCELACIONPEDIDO, idDifStock, 0, 0, "", 0, uFecha.dtfecha, True, Date, UsuarioActual()
'                End If
'            Next i
'            GrabaAlta = idDifStock
        DE_CommitTrans
        
        MsgBox "Se ha eliminado con exito.", , "ATENCION"
        uMenu.AceptarOk
        '*******************************************************
    End If
    
fin:
    Set rs = Nothing
    Exit Sub
    
UFAalta:
    DE_RollbackTrans
    ufa "err grabando cancelacion", "alta"
    Resume fin
End Sub

Private Sub uMenu_Imprimir()
    CancelacionPedido lblNumero
End Sub

Private Sub uPedido_Buscar()
    Dim s As String
'    If uCliente.Codigo = 0 Then
        s = "SELECT DISTINCT p.numero as [ Pedido ], p.fecha as [ Fecha        ], p.cliente, c.descripcion as [ Nombre                               ] FROM Clientes AS c INNER JOIN (Pedidos_Clientes AS p INNER JOIN ItemPedidoCliente AS i ON p.numero = i.PEDIDO) ON c.codigo = p.cliente Where p.activo = 1  And fecha > " & uFechaBuscar.ConvertFecha & " and i.Saldo > 0 ORDER BY p.numero DESC "
'    Else
'        s = "SELECT DISTINCT p.numero, p.cliente, p.fecha, c.descripcion FROM Clientes AS c INNER JOIN (Pedidos_Clientes AS p INNER JOIN ItemPedidoCliente AS i ON p.numero = i.PEDIDO) ON c.codigo = p.cliente Where p.activo = 1  And fecha > " & uFechaBuscar.ConvertFecha & " and i.Saldo > 0 and p.cliente  = " & uCliente.Codigo & " ORDER BY p.numero DESC "
'    End If
    uPedido.strSqlBuscar = s
End Sub

Private Sub uPedido_cambio(codigo As Variant)
    Dim Propio As Boolean

    If uPedido.codigo > 0 Then
        'If uCliente.Codigo = 0 Then
        uCliente.codigo = ClientePedido(codigo)
        'endif
        cargaGrilla
        Propio = obtenerDeSQL("select CodigoPropio from Pedidos_Clientes where numero = " & uPedido.codigo)
        optPropio(0).Value = Propio
        optPropio(1).Value = Not Propio
    End If
End Sub

Private Function Propio()
    Propio = optPropio(0).Value
End Function

Private Sub cargaGrilla()
    If ON_ERROR_HABILITADO Then On Error GoTo UfaCarga
    Dim rs As New ADODB.Recordset, ssql As String, i As Long

    g.Borrar
    ssql = "select codigo, producto, saldo, precio from ItemPedidoCliente where pedido = " & uPedido.codigo & " and saldo > 0 "

    With rs
        .Open ssql, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        While Not .EOF
            AgregarEnGrilla !codigo, sSinNull(!producto), !saldo, !precio
            .MoveNext
        Wend
    End With

fin:
    Set rs = Nothing
    Exit Sub
UfaCarga:
    ufa "err cargando datos", ""
    Resume fin
End Sub

Private Sub AgregarEnGrilla(Item, producto As String, saldo, precio)
    Dim rs As New ADODB.Recordset
    With g
        If gEMPR_FormulaEsVirtual Then
            Set rs = rsFormulaComponentes(producto)
            If rs.EOF Then
                ItemGrilla Item, producto, saldo, precio, ""
            Else
                ItemGrilla Item, producto, saldo, precio, "V"
            End If
        Else
            ItemGrilla Item, producto, saldo, precio, ""
        End If
    End With
    Set rs = Nothing
End Sub

Private Sub ItemGrilla(Item, producto As String, saldo, precio, formula)
    Dim i As Long
    With g
        i = .addRow()
        .tx i, gITEM, Item
        .tx i, gprod, VerProductoCliente(producto, Propio, uCliente.codigo)
        .tx i, gDESC, DescripcionProducto(producto)
        .tx i, gCANT, saldo
        .tx i, gSALD, saldo
        .tx i, gPREC, precio
        .tx i, gFORM, formula
    End With
End Sub

Private Sub inigrilla()
    Set g = New LiGrilla
    Set g3 = New LiGrilla
    
    g.init grillaCancelacion
    g3.init grillaSeries
    
    With g
        gITEM = .AddCol("-item- ", "H")         ' item pedido
        gprod = .AddCol(" Codigo                     ")
        gDESC = .AddCol(" Descripcion                                       ")
        gSALD = .AddCol(" Saldo       ", "9")   ' saldo original (maximo para control)
        gCANT = .AddCol(" Cantidad    ", "N")
        gPREC = .AddCol(" Precio Unitario ", "N")
        gFORM = .AddCol(" Formula            ")
    End With
    With g3
        g3ITEM = .AddCol("  -  ", "A")
        g3PROD = .AddCol(" Producto                      ")
        g3NSER = .AddCol(" Numero de Serie            ", "S") ' editable
        g3HIDD = .AddCol("", "H")
        g3ALTA = g3.AddCol("                ")
        g3.Borrar
        grillaSeries.SelectionMode = flexSelectionListBox
    End With
End Sub

Private Function ClientePedido(QueNumeroPedido) As Long
    ClientePedido = s2n(obtenerDeSQL("select cliente from pedidos_Clientes where numero = " & QueNumeroPedido))
End Function

Private Function GrabaAlta() As Long
    If ON_ERROR_HABILITADO Then On Error GoTo UFAalta
    Dim idDifStock As Long, i As Long, nComprobante As Long, numero As Long
    Dim rs As New ADODB.Recordset, Serie As String
    Dim Item As Long, prod As String, cant As Double, prec As Double, formu As String
    
    idDifStock = nuevoCodigo("RemitoDiferenciaStock", "MovimientoInterno")

    
    
    nComprobante = ComprobanteConcepto(CONCEPTO_CANCELACION)
    numero = nuevoCodigo("RemitoDiferenciaStock", "NroComprobante", "comprobante = " & nComprobante)
    lblNumero = idDifStock 'numero
    lblId = idDifStock
    If confirma("Graba Cancelacion Numero" & vbCrLf & lblNumero) Then
        '*******************************************************
        DE_BeginTrans
            DataEnvironment1.dbo_DIFSTOCK uFecha.dtFecha, idDifStock, nComprobante, numero, CONCEPTO_CANCELACION, uPedido.codigo, 0
            For i = 1 To g.rows - 1
                Item = s2n(g.tx(i, gITEM))
                prod = VerProductoMio(Trim$(g.tx(i, gprod)), Propio)
                cant = s2n(g.tx(i, gCANT))
                prec = s2n(g.tx(i, gPREC))
                formu = Trim$(g.tx(i, gFORM))
                
                If cant > 0 Then
                    'actualiz pedido
                    If Item > 0 Then ' si 0, es componente de formula virtual, no figura en pedido
                        DataEnvironment1.Sistema.Execute "update itemPedidoCliente set facturar = facturar - " & x2s(cant) & ", saldo = saldo - " & x2s(cant) & " where codigo = " & Item
                    End If
                    
                    'resto stock, grabo item
                    If formu = "V" Then  ' producto virtual no baja stock
                        DataEnvironment1.dbo_ITEMDIFSTOCK "F", prod, cant, idDifStock, prec
                    Else
                        DataEnvironment1.dbo_ITEMDIFSTOCK "R", prod, cant, idDifStock, prec
                    End If
                End If
            Next i
            For i = 1 To g3.rows - 1 'series
                prod = VerProductoMio(g3.tx(i, g3PROD), Propio())
                Serie = g3.tx(i, g3NSER)
                If Serie <> "" Then
                    'DataEnvironment1.dbo_SERIE "A", 0, prod, serie, TipoComprobante_CANCELACIONPEDIDO, idDifStock, 0, 0, "", 0, Date, UsuarioActual(), 0, 0
                    DataEnvironment1.dbo_abmSERIEs "A", 0, prod, Serie, TipoComprobante_CANCELACIONPEDIDO, idDifStock, 0, 0, "", 0, uFecha.dtFecha, True, Date, UsuarioActual()
                End If
            Next i
            GrabaAlta = idDifStock
        DE_CommitTrans
        '*******************************************************
    End If
    
fin:
    Set rs = Nothing
    Exit Function
UFAalta:
    GrabaAlta = 0
    DE_RollbackTrans
    ufa "err grabando cancelacion", "alta"
    Resume fin
End Function

Private Function FaltaSeries() As Boolean
    Dim r As Long, i As Long, ns As String
    Dim seri As String, prod As String
    
    FaltaSeries = False
    lblErrorSeries.Visible = False
    
    If chkSinSeries.Value = vbChecked Then Exit Function
    
    LlenoGrillaSeries
    r = g3.rows
    
    'vacio
    If r > 1 And g3.buscar(g3NSER, "") > 0 Then
        
        tabMain.Tab = 1
        grillaSeries.SetFocus
        grillaSeries.Select g3.PrimerVacio(g3NSER), g3NSER
        
        FaltaSeries = True
        Exit Function
    End If
    
    'existe serie ?
    For i = 1 To r - 1
        seri = g3.tx(i, g3NSER)
        prod = g3.tx(i, g3PROD)
        If Not SerieEnStock(seri, prod) Then
            If g3.tx(i, g3ALTA) <> CTE_SERIE_AGREGAR Then
                
                che "No figura en stock " & prod & "  " & seri
                tabMain.Tab = 1
                grillaSeries.SetFocus
                grillaSeries.Select i, g3NSER
                
                If confirma("Desea registrarlo ahora") Then
                    g3.tx i, g3ALTA, CTE_SERIE_AGREGAR
                Else
                    FaltaSeries = True
                    Exit Function
                End If
            End If
        End If
    Next i
     
    If r > 1 Then
        For i = 1 To r - 2
            ns = g3.tx(i, g3NSER)
            If ns <> "" And g3.buscar(g3NSER, ns, i + 1) > 0 Then
                tabMain.Tab = 1
                grillaSeries.SetFocus
                grillaSeries.Select i, g3NSER, g3.buscar(g3NSER, ns, i + 1), g3NSER
                
                'grillaSeries.Select g3.Buscar(g3NSER, ns, i + 1), g3NSER
                FaltaSeries = True
                Exit Function
            End If
        Next i
    End If
End Function

Private Function CargaCancelacion(numero As Long) As Boolean
    If ON_ERROR_HABILITADO Then On Error GoTo UfaCarga
    
    Dim tempo, rs As New ADODB.Recordset, i As Long
    
    tempo = obtenerDeSQL(" select " _
        & " MovimientoInterno, Comprobante, NroComprobante, Concepto, Fecha, NroPedido from RemitoDiferenciaStock where MovimientoInterno  = " & numero) 'nrocomprobante  = " & numero)
    If IsEmpty(tempo) Then Exit Function
    
    lblId = s2n(tempo(0))
    uPedido.codigo = s2n(tempo(5))
    uFecha.dtFecha CDate(tempo(4))
    'lblNumero = tempo(2)
    lblNumero = tempo(0)
    
    With rs
        g.Borrar
        .Open "select   * from ItemRemitoDiferenciaStock where numero = " & numero, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        While Not .EOF
            i = g.addRow
            g.tx i, gCANT, !cantidad
            g.tx i, gprod, !producto
            g.tx i, gDESC, ObtenerDescripcionS("producto", !producto)
            .MoveNext
        Wend
    End With
    CargaCancelacion = True
    
fin:
    Set rs = Nothing
    Exit Function
UfaCarga:
    CargaCancelacion = False
    Resume fin
End Function


'**********************************************************************
'-------------------------------- MENU --------------------------------
Private Sub uMenu_AceptarAlta()
    Dim idDif  As Long
    
    If g.suma(gCANT) = 0 Then
        che "sin cantidades a bajar "
        Exit Sub
    End If
    If uCliente.codigo = 0 Or uPedido.codigo = 0 Then
        che "faltan datos, cliente pedido"
        Exit Sub
    End If
    If FaltaSeries() Then
        'che ya te avise
        Exit Sub
    End If
    
    idDif = GrabaAlta()
    If idDif > 0 Then
        che "Diferencia grabada." & vbCrLf & " Numero interno : " & lblNumero
'        ImprimeDiferenciaStock idDif
        uMenu.AceptarOk
    End If
End Sub
Private Sub uMenu_Buscar()
    Dim resu As String, s As String
    s = "select  MovimientoInterno as Numero, NroPedido as Pedido, fecha, MovimientoInterno as [cod] from RemitoDiferenciaStock where concepto = " & CONCEPTO_CANCELACION & " and fecha >= " & uFechaBuscar.ConvertFecha & " order by MovimientoInterno desc"
    resu = frmBuscar.MostrarSql(s)
    If resu > "" Then
        CargaCancelacion s2n(resu)
        uMenu.BuscarOK
    End If
End Sub
Private Sub uMenu_BorrarControles()
    lblNumero = ""
    lblId = ""
    uCliente.clear
    chkSinSeries.Value = vbUnchecked
    g.Borrar
    g3.Borrar
End Sub
Private Sub uMenu_HabilitarEdicion(sino As Boolean)
    fraMain.enabled = sino
End Sub
Private Sub uMenu_SALIR()
    Unload Me
End Sub
'-------------------------------- MENU --------------------------------
'**********************************************************************

'
'21/3/5 codigo desglose formula pasa a pedido
'27/4/5 grilla series fixes varios: ayuda, verificacion
'11/5/5 pregunta antes de grabar
'



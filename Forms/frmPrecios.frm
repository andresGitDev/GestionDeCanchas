VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmPrecios 
   Caption         =   "Precios "
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9510
   Icon            =   "frmPrecios.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   9510
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   930
      Left            =   7650
      Picture         =   "frmPrecios.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6330
      Width           =   870
   End
   Begin Gestion.ucXls ucXls1 
      Height          =   930
      Left            =   6780
      TabIndex        =   7
      Top             =   6345
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1640
   End
   Begin VB.Frame fraBoton 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   4950
      TabIndex        =   3
      Top             =   6210
      Width           =   4545
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   930
         Left            =   3585
         Picture         =   "frmPrecios.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   870
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Recargar Grilla"
         Height          =   915
         Left            =   930
         Picture         =   "frmPrecios.frx":1A5E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   135
         Width           =   885
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "Editar"
         Height          =   915
         Left            =   120
         Picture         =   "frmPrecios.frx":2328
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   135
         Width           =   795
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Height          =   5070
      Left            =   90
      TabIndex        =   2
      Top             =   1005
      Width           =   9300
      _cx             =   16404
      _cy             =   8943
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
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
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
      AutoSearch      =   1
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
   Begin Gestion.ucCoDe uCliente 
      Height          =   300
      Left            =   2880
      TabIndex        =   0
      Top             =   195
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   529
      CodigoWidth     =   1000
   End
   Begin Gestion.ucCoDe uProv 
      Height          =   300
      Left            =   2895
      TabIndex        =   8
      Top             =   600
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   529
      CodigoWidth     =   1000
   End
   Begin VB.Label lblConIva 
      Caption         =   "Precio con IVA"
      Height          =   315
      Left            =   165
      TabIndex        =   12
      Top             =   585
      Width           =   1635
   End
   Begin VB.Label Label1 
      Caption         =   "Proveedor"
      Height          =   315
      Index           =   1
      Left            =   1830
      TabIndex        =   10
      Top             =   600
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente"
      Height          =   315
      Index           =   0
      Left            =   1830
      TabIndex        =   9
      Top             =   240
      Width           =   1005
   End
   Begin VB.Label lblEditando 
      Caption         =   "EDITANDO"
      Height          =   330
      Left            =   165
      TabIndex        =   1
      Top             =   180
      Visible         =   0   'False
      Width           =   1635
   End
End
Attribute VB_Name = "frmPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mPRECIOCONIVA   As Boolean

Private mColPrecio1 As Long
Private mColPrecio2 As Long

Private Sub cmdCancelar_Click()
    llenar
End Sub

Private Sub cmdImprimir_Click()
    If Grilla.rows < 2 Then Exit Sub
    
    Grilla.GridLines = flexGridNone
    Grilla.GridLinesFixed = flexGridNone
    
    FrmImpresiones.VSPrinter.Orientation = orPortrait
    FrmImpresiones.VSPrinter.PaperSize = pprA4
    FrmImpresiones.VSPrinter.Preview = True
    FrmImpresiones.VSPrinter.Font.Name = Grilla.Font.Name
    FrmImpresiones.VSPrinter.FontSize = 12
    FrmImpresiones.VSPrinter.Header = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
    FrmImpresiones.VSPrinter.FontSize = 8
    
    FrmImpresiones.VSPrinter.StartDoc
    FrmImpresiones.VSPrinter.Paragraph = Date
    FrmImpresiones.VSPrinter.Paragraph = titulito()
    
    FrmImpresiones.VSPrinter.TextAlign = taLeftBottom
    FrmImpresiones.VSPrinter.RenderControl = Grilla.hWnd

    FrmImpresiones.VSPrinter.Footer = "||Pagina %d  "
    FrmImpresiones.VSPrinter.Zoom = 100
    FrmImpresiones.VSPrinter.EndDoc
    
    FrmImpresiones.Show
    Grilla.GridLines = flexGridFlat

End Sub

Private Function titulito()
    Dim s
    
    If uCliente.codigo > 0 Then
        s = "Precios para cliente " & uCliente.DESCRIPCION
    ElseIf UpROV.codigo > 0 Then
        s = "Precios para cliente " & UpROV.DESCRIPCION
    Else
        s = "precios productos"
    End If
    
    titulito = s
End Function

Private Sub cmdmodificar_Click()
    Habilita True
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    uCliente.ini "select Descripcion from clientes where codigo = '###' ", "select distinct codigo, Descripcion as [ Nombre                    ] from clientes inner join relacion_producto_cliente on clientes.codigo = relacion_producto_cliente.cliente where clientes.activo = 1 and relacion_producto_cliente.activo  = 1"
    UpROV.ini "select Descripcion from prov where codigo = '###' ", "select distinct codigo, Descripcion as [ Nombre                    ] from prov inner join relacion_producto_proveedor on prov.codigo = relacion_producto_proveedor.proveedor where prov.activo = 1 "
    ucXls1.ini Grilla, "c:\Precios_" & Format(Date, "yyyy-mm-dd") & "_"
    
    'PERDON
    If gEMPR_idEmpresa = 7 Then    ' MAMYBLUE
        mPRECIOCONIVA = True
        lblConIva.Visible = True
    End If
    
    Form_Resize
    llenar
    Grilla.ExplorerBar = flexExSortShowAndMove
End Sub
Private Sub Form_Resize()
    Anclar Grilla, Me, anclarLadosTodos
    Anclar fraBoton, Me, anclarAbajo + anclarIzquierda
End Sub

Private Sub grilla_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> mColPrecio1 And Col <> mColPrecio2 Then Cancel = True
End Sub

Private Sub grilla_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim pVie As Double, pNue As Double, prod As String
    
    If Col = mColPrecio1 Or mColPrecio2 Then  ' por ahora solo dejo modificar precio
        pVie = s2n(Grilla.Text, 4)
        pNue = s2n(Grilla.EditText, 4)
        prod = Grilla.TextMatrix(Row, 0)
        
        Grilla.EditText = pNue
        If pNue <> pVie Then
            If confirma("de " & pVie & "  a  " & pNue & vbCrLf & "acepta el nuevo precio") Then
                
                If uCliente.codigo = 0 And UpROV.codigo = 0 Then
                    cambiaPrecioProducto prod, pNue, Col
                ElseIf uCliente.codigo > 0 Then
                    cambiaPrecioCliente prod, uCliente.codigo, pNue
                ElseIf UpROV.codigo > 0 Then
                    cambiaPrecioProv prod, UpROV.codigo, pNue, s2n(Grilla.TextMatrix(Row, Col + 1))
                End If
            
            Else
                Cancel = True
                Grilla.EditText = Grilla.Text ' no hace falta
            End If
        End If
    Else
        '
        Cancel = True
    End If
End Sub

Private Sub cambiaPrecioProducto(prod As String, cuanto As Double, cual As Long)
    Dim sCuanto
    
    If mPRECIOCONIVA Then
        sCuanto = " Round( " & ssNum(cuanto) & " / ( 1 + iva) , 4) "
    Else
        sCuanto = ssNum(cuanto)
    End If


    If cual = mColPrecio1 Then
        DataEnvironment1.Sistema.Execute "update producto set precio  = " & sCuanto & " where codigo = '" & ssStr(prod) & "' "
    ElseIf cual = mColPrecio2 Then
        DataEnvironment1.Sistema.Execute "update producto set precio2 = " & sCuanto & " where codigo = '" & ssStr(prod) & "' "
    End If

    che "cambiado"
End Sub


Private Sub cambiaPrecioCliente(prod As String, clie As Long, cuanto As Double) ', quien As Long)
    DataEnvironment1.Sistema.Execute _
        " update relacion_producto_cliente " & _
        " set precio = " & ssNum(cuanto) & _
        " where producto = '" & ssStr(prod) & "' " & _
        " and cliente = " & clie
    che "cambiado"
End Sub
Private Sub cambiaPrecioProv(prod As String, prov As Long, cuanto As Double, cotizacion As Double)
    Dim cotiz As Double
    
    cotiz = s2n(InputBox("Cotizacion ", , cotizacion))
    DataEnvironment1.Sistema.Execute _
        " update relacion_producto_proveedor " & _
        "   set precio = " & ssNum(cuanto) & _
        "   , fechaCarga = " & ssFecha(Date) & _
        "   , cotizacion = " & ssNum(cotiz) & _
        "   where producto = '" & ssStr(prod) & "' " & _
        "   and proveedor = " & UpROV.codigo
    che "cambiado"
End Sub


Private Sub uCliente_cambio(codigo As Variant)
    If codigo > 0 Then
        UpROV.clear
        llenar
    End If
End Sub

Private Sub llenar()
    If uCliente.codigo = 0 And UpROV.codigo = 0 Then
        If mPRECIOCONIVA Then
           LlenarGrilla Grilla, "select codigo, descripcion, round(precio *( 1 + iva),2) , round(precio2 *( 1 + iva),2) from producto order by codigo", True
        Else
           LlenarGrilla Grilla, "select codigo, descripcion, precio, precio2 from producto order by codigo", True
        End If
        mColPrecio1 = 2
        mColPrecio2 = 3
        grillaWidth Grilla, Array(1300, 3500, 1000)
    ElseIf uCliente.codigo > 0 Then
        LlenarGrilla Grilla, "select producto, productocliente, descripcion, r.precio from relacion_producto_cliente as r inner join producto as p on p.codigo = r.producto where r.activo = 1 and cliente = " & uCliente.codigo, True
        mColPrecio1 = 3
        mColPrecio2 = -1
        grillaWidth Grilla, Array(1300, 1300, 3500, 1100, 1000, 1400)
    ElseIf UpROV.codigo > 0 Then
        LlenarGrilla Grilla, "select producto, CodigoProveedor as [Producto Prov], Descripcion, r.precio, Cotizacion, FechaCarga from Relacion_producto_proveedor as r inner join producto as p on p.codigo = r.producto where  proveedor = " & UpROV.codigo, True
        mColPrecio1 = 3
        mColPrecio2 = -1
        grillaWidth Grilla, Array(1300, 1300, 3500, 1100, 1000, 1400)
    Else
        ufa "err", "prov y clie al mmo tiempo"
        mColPrecio1 = -1
        mColPrecio2 = -1
    End If
    Habilita False
    Grilla.ColAlignment(0) = flexAlignLeftCenter
    'grilla.e
End Sub

Private Sub Habilita(sino As Boolean)
    Grilla.Editable = IIf(sino, flexEDKbdMouse, flexEDNone)
    Grilla.AutoSearch = IIf(sino, flexSearchNone, flexSearchFromCursor)
    lblEditando.Visible = sino
    uCliente.enabled = Not sino
    cmdcancelar.enabled = sino
    cmdmodificar.enabled = Not sino
End Sub


Private Sub uProv_cambio(codigo As Variant)
    If codigo > 0 Then
        uCliente.clear
        llenar
    End If
End Sub

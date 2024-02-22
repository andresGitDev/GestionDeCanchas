VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmPedidosPendientesBIG 
   Caption         =   "Pedidos pendientes"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10350
   Icon            =   "frmPedidosPendientesBIG.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   10350
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   345
      Left            =   2730
      TabIndex        =   10
      Top             =   1065
      Width           =   1305
   End
   Begin VB.OptionButton optDetalle 
      Caption         =   "Con Detalle"
      Height          =   330
      Index           =   1
      Left            =   6060
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Value           =   -1  'True
      Width           =   1515
   End
   Begin VB.OptionButton optDetalle 
      Caption         =   "Sin Detalle"
      Height          =   330
      Index           =   0
      Left            =   4545
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1515
   End
   Begin Gestion.ucCoDe uCliente 
      Height          =   330
      Left            =   1455
      TabIndex        =   3
      Top             =   570
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   582
      CodigoWidth     =   1000
   End
   Begin Gestion.ucFecha uFecha 
      Height          =   315
      Left            =   1470
      TabIndex        =   0
      Top             =   120
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
      FechaInit       =   5
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4035
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1065
      Width           =   1050
   End
   Begin VB.CommandButton cmdMostrar 
      Caption         =   "&Mostrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1410
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1065
      Width           =   1305
   End
   Begin Gestion.ucXls ucXls1 
      Height          =   825
      Left            =   7740
      TabIndex        =   5
      Top             =   120
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   1455
   End
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Height          =   4650
      Left            =   15
      TabIndex        =   6
      Top             =   1500
      Width           =   10290
      _cx             =   18150
      _cy             =   8202
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
   Begin VB.Label Label1 
      Caption         =   "Cliente (0 = todos)"
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   9
      Top             =   615
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "Buscar desde :"
      Height          =   330
      Index           =   0
      Left            =   60
      TabIndex        =   8
      Top             =   150
      Width           =   1485
   End
End
Attribute VB_Name = "frmPedidosPendientesBIG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdImprimir_Click()
    If grilla.rows < 2 Then Exit Sub
    
    grilla.GridLines = flexGridNone
    grilla.GridLinesFixed = flexGridNone
    
    FrmImpresiones.VSPrinter.Orientation = orPortrait
    FrmImpresiones.VSPrinter.PaperSize = pprA4
    FrmImpresiones.VSPrinter.Preview = True
    FrmImpresiones.VSPrinter.Font.Name = grilla.Font.Name
    FrmImpresiones.VSPrinter.FontSize = 12
    FrmImpresiones.VSPrinter.Header = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
    FrmImpresiones.VSPrinter.FontSize = 8
    FrmImpresiones.VSPrinter.Footer = "||Pagina %d " 'de " & FrmImpresiones.VSPrinter.PageCount
    
    FrmImpresiones.VSPrinter.StartDoc
    FrmImpresiones.VSPrinter.Paragraph = "Pedidos Pendientes " & Date
    FrmImpresiones.VSPrinter.Paragraph = "desde : " & uFecha.strFecha
    FrmImpresiones.VSPrinter.Paragraph = " "
    
    FrmImpresiones.VSPrinter.TextAlign = taLeftBottom
    FrmImpresiones.VSPrinter.RenderControl = grilla.hWnd

    FrmImpresiones.VSPrinter.Footer = "||Pagina %d " 'de " & FrmImpresiones.VSPrinter.PageCount
    
    FrmImpresiones.VSPrinter.EndDoc
    
    FrmImpresiones.VSPrinter.Zoom = 80
    FrmImpresiones.Show
    grilla.GridLines = flexGridFlat
End Sub

Private Sub cmdMostrar_Click()
    Dim ss As String, ssW As String
    
    'esto es usando saldo para remitos y facturar para lo facturado
    'pero como no se manejan con remitos lo dejo comentado
    'ssW = " where (saldo > 0 or facturar >0 ) and fecha >=  " & uFecha.ssFecha & " and pc.activo = 1 "
    ssW = " where (facturar >0 ) and fecha >=  " & uFecha.ssFecha & " and pc.activo = 1 "
    
    If uCliente.codigo > 0 Then ssW = ssW & " and cliente = " & uCliente.codigo & " "
    
    If optDetalle(0) Then
'        ss = "select Numero, Pedido_cli as [ Pedido Cliente ], fecha, descripcion as [ Cliente     ] from Pedidos_Clientes left join clientes where Pedidos_Clientes.activo = 1 "
        
        ss = "SELECT distinct numero as Numero, Pedido_cli as [ Pedido Cliente ], Fecha, C.descripcion as [ Cliente           ] " & _
            " FROM (Pedidos_Clientes as PC INNER JOIN ItemPedidoCliente as IPC " & _
            " ON  PC.Numero = IPC.Pedido) " & _
            " INNER JOIN Clientes C ON C.Codigo = PC.Cliente " & _
            ssW & " ORDER BY  c.descripcion, fecha "
        
        LlenarGrilla grilla, ss, False, 3
        grillaWidth grilla, Array(800, 1010, 1200, 3000)
        
    Else
        ss = "SELECT distinct numero as Numero, Pedido_cli as [ Pedido Cliente ], Fecha, C.descripcion as [ Cliente           ] " & _
            " , IPC.Producto, IPC.Cantidad, IPC.Facturar as Saldo, IPC.fechaEntrega " & _
            " FROM (Pedidos_Clientes as PC INNER JOIN ItemPedidoCliente as IPC " & _
            " ON  PC.Numero = IPC.Pedido) " & _
            " INNER JOIN Clientes C ON C.Codigo = PC.Cliente " & _
            ssW & " ORDER BY  c.descripcion, fecha "

        LlenarGrilla grilla, ss, False, 0
        grillaWidth grilla, Array(800, 1010, 1200, 2000, 1200, 800, 800, 1200)
    End If
    
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    uCliente.ini "select descripcion from clientes where codigo = '###' ", "select codigo, descripcion as [ Nombre                         ] from clientes ", False
    ucXls1.ini grilla, "C:\PedidosPendientes " & uFecha.ssFecha, "PedidosPendientes " & uFecha.strFecha
    Form_Resize
    cmdMostrar_Click
End Sub
Private Sub Form_Resize()
    Anclar grilla, Me, anclarLadosTodos
End Sub


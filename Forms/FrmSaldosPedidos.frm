VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmSaldosPedidos 
   Caption         =   "Saldos de pedidos"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8640
   Icon            =   "FrmSaldosPedidos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   8640
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraBotones 
      BorderStyle     =   0  'None
      Caption         =   "fra"
      Height          =   1005
      Left            =   60
      TabIndex        =   4
      Top             =   5940
      Width           =   8700
      Begin Gestion.ucXls uXls 
         Height          =   840
         Left            =   6000
         TabIndex        =   6
         Top             =   45
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   1482
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   345
         Left            =   7095
         TabIndex        =   5
         Top             =   45
         Width           =   1395
      End
   End
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "Buscar"
      Height          =   360
      Left            =   2145
      TabIndex        =   3
      Top             =   150
      Width           =   1200
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   330
      Left            =   345
      TabIndex        =   2
      Top             =   165
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   582
      _Version        =   393216
      Format          =   67305473
      CurrentDate     =   38686
   End
   Begin VSFlex7LCtl.VSFlexGrid GridDetalle 
      Height          =   2190
      Left            =   240
      TabIndex        =   1
      Top             =   3690
      Width           =   8325
      _cx             =   14684
      _cy             =   3863
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
      BackColor       =   -2147483633
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483633
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
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmSaldosPedidos.frx":08CA
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
   Begin VSFlex7LCtl.VSFlexGrid Grid 
      Height          =   2925
      Left            =   180
      TabIndex        =   0
      Top             =   690
      Width           =   8295
      _cx             =   14631
      _cy             =   5159
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
      BackColor       =   -2147483633
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483633
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
End
Attribute VB_Name = "FrmSaldosPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rs As New ADODB.Recordset
Dim ssql  As String
Dim a As Integer

Private Sub cmdBuscar_Click()
Call CargarPedido
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
    uXls.ini Grid, "C:\temp", "Pedidos pendientes"
DTPicker1.Value = CDate("1/1/" & Year(Date))
Call CargarPedido
End Sub

Private Sub Form_Resize()
    Anclar Grid, Me, anclarLadosTodos
    Anclar GridDetalle, Me, anclarLadosAncho + anclarAbajo
    Anclar fraBotones, Me, anclarAbajo
End Sub

Private Sub Grid_Click()
Dim rsdetalle As New ADODB.Recordset

If Trim(Grid.TextMatrix(Grid.Row, 0)) <> 0 Then

   rsdetalle.Open "SELECT ItemPedidoCliente.pedido, Producto.descripcion,ItemPedidoCliente.saldo,ItemPedidoCliente.FechaEntrega " & _
   " FROM ItemPedidoCliente left JOIN Producto ON " & _
   " ItemPedidoCliente.producto =  Producto.codigo " & _
   " WHERE ItemPedidoCliente.pedido = '" & Trim(Grid.TextMatrix(Grid.Row, 0)) & "' ", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
   
   GridDetalle.cols = rsdetalle.Fields.Count
   GridDetalle.rows = 0
   
   If Not rsdetalle.EOF Then
      Do While Not rsdetalle.EOF
        GridDetalle.AddItem rsdetalle!PEDIDO & Chr(9) & rsdetalle!DESCRIPCION & Chr(9) & rsdetalle!saldo & Chr(9) & rsdetalle!fechaEntrega
         rsdetalle.MoveNext
      Loop
      
          GridDetalle.AddItem "", 0
          GridDetalle.FixedCols = 1

      For a = 0 To GridDetalle.cols - 1
            GridDetalle.TextMatrix(0, a) = rsdetalle.Fields(a).Name
      Next
      GridDetalle.FixedRows = 1
      
  End If
  rsdetalle.Close
  Set rsdetalle = Nothing
End If
End Sub

Private Sub CargarPedido()
ssql = "SELECT distinct numero as Numero,fecha,cliente,C.descripcion as Razon_Social FROM " & _
           "(Pedidos_Clientes as PC INNER JOIN ItemPedidoCliente as IPC " & _
           " ON  PC.Numero = IPC.Pedido) INNER JOIN Clientes C ON C.Codigo = PC.Cliente " & _
           " where IPC.saldo >  0 and fecha >= " & ssFecha(DTPicker1) & "  " & _
           "ORDER BY  fecha "
rs.Open (ssql), DataEnvironment1.Sistema, adOpenForwardOnly, adLockOptimistic
 
If rs.EOF = True Then
   Grid.clear
   MsgBox "No se encuentran Pedidos", vbInformation, Me.caption
   rs.Close
   Set rs = Nothing
   Exit Sub
End If

Grid.cols = rs.Fields.Count
Grid.rows = 0

Do While rs.EOF = False
  Grid.rows = Grid.rows + 1
        
     For a = 0 To rs.Fields.Count - 1
             Grid.TextMatrix(Grid.rows - 1, a) = sSinNull(rs.Fields(a))
     Next
 rs.MoveNext
Loop

Grid.AddItem "", 0
Grid.FixedCols = 1
For a = 0 To Grid.cols - 1
    Grid.TextMatrix(0, a) = rs.Fields(a).Name
Next

Grid.FixedRows = 1

If Grid.rows > 0 Then
    Grid.ColWidth(0) = Grid.Width * 0.1
    Grid.ColWidth(1) = Grid.Width * 0.2
    Grid.ColWidth(2) = Grid.Width * 0.1
    Grid.ColWidth(3) = Grid.Width * 0.4
End If
rs.Close
End Sub

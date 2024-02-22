VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form FrmLisStockSimplificado 
   Caption         =   "Listado de Stock Comprometido"
   ClientHeight    =   8940
   ClientLeft      =   1620
   ClientTop       =   345
   ClientWidth     =   10170
   Icon            =   "FrmLisStockSimplificado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   10170
   StartUpPosition =   2  'CenterScreen
   Begin Gestion.ucXls ucXls1 
      Height          =   810
      Left            =   9330
      TabIndex        =   7
      Top             =   8055
      Width           =   765
      _extentx        =   1349
      _extenty        =   1429
   End
   Begin VB.CheckBox chkPrecio 
      Caption         =   "Ver Precio"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   8430
      TabIndex        =   6
      Top             =   60
      Value           =   1  'Checked
      Width           =   1140
   End
   Begin VB.TextBox txtbuscar 
      Height          =   315
      Left            =   645
      TabIndex        =   4
      Top             =   8190
      Width           =   4005
   End
   Begin VB.CommandButton cmdact 
      Caption         =   "Actualizar"
      DisabledPicture =   "FrmLisStockSimplificado.frx":08CA
      DownPicture     =   "FrmLisStockSimplificado.frx":0D57
      Height          =   780
      Left            =   7695
      Picture         =   "FrmLisStockSimplificado.frx":13B6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8070
      Width           =   870
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir"
      DownPicture     =   "FrmLisStockSimplificado.frx":1C80
      Height          =   795
      Left            =   8565
      Picture         =   "FrmLisStockSimplificado.frx":2260
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8055
      Width           =   765
   End
   Begin VSFlex7LCtl.VSFlexGrid grillastock 
      Height          =   5550
      Left            =   30
      TabIndex        =   0
      Top             =   315
      Width           =   10110
      _cx             =   17833
      _cy             =   9790
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
      BackColorFixed  =   14737632
      ForeColorFixed  =   0
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   12640511
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   0
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
      Rows            =   1
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmLisStockSimplificado.frx":2B2A
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   7
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
   Begin VSFlex7LCtl.VSFlexGrid grillaped 
      Height          =   2070
      Left            =   30
      TabIndex        =   2
      Top             =   5910
      Width           =   10110
      _cx             =   17833
      _cy             =   3651
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
      BackColorFixed  =   14737632
      ForeColorFixed  =   0
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   12640511
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   0
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
      Rows            =   1
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmLisStockSimplificado.frx":2C2A
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
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Escribir y apretar ENTER para buscar la palabra"
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   4815
      TabIndex        =   5
      Top             =   8145
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   90
      Picture         =   "FrmLisStockSimplificado.frx":2CAE
      Stretch         =   -1  'True
      Top             =   8145
      Width           =   465
   End
End
Attribute VB_Name = "FrmLisStockSimplificado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mRow As Integer
Dim mCol As Integer


Private Sub chkPrecio_Click()
    If chkPrecio.Value = 1 Then
        grillastock.ColHidden(5) = False
    Else
        grillastock.ColHidden(5) = True
    End If
End Sub

Private Sub cmdact_Click()
    Cargogrilla
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Cargogrilla
    ucXls1.ini grillastock, "C:\ResumenStockComprometido.xls"
End Sub
Sub Cargogrilla()
    Dim titi
    titi = Timer

    Dim rsr As New ADODB.Recordset, s As String
    s = "SELECT p.codigo as cod, Sum(pi.saldo) AS sal FROM Pedidos_Clientes AS pe RIGHT JOIN (Producto AS p LEFT JOIN ItemPedidoCliente AS pi ON p.codigo = pi.Producto) ON pe.numero = pi.PEDIDO Where (((p.activo) = 1) And ((pe.activo) = 1) And ((pe.cancelado) = 0)) GROUP BY p.codigo, p.existencia, p.formula order by cod"
    rsr.Open s, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    If rsr.EOF And rsr.BOF Then
        MsgBox "No hay pedidos.", vbExclamation
        Exit Sub
    End If
    
    Dim existencia As Double, reserva As Double, rs As New ADODB.Recordset
    grillastock.rows = 1
    grillaped.rows = 1

    rs.Open "Select p.codigo, p.descripcion, p.existencia, p.formula,p.precio,s.descripcion as subg from producto p inner join subgruposproducto s on p.subgrupo=s.codigo where p.activo=1 order by p.descripcion", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    grillastock.rows = rs.RecordCount + 1
    Dim i As Long
    While Not rs.EOF
        'If rs!formula Then
            
        '    existencia = ExistenciaCalculada(rs!codigo)
        'Else

            existencia = rs!existencia
        'End If

        reserva = 0
        rsr.MoveFirst
        rsr.Find "cod = '" & rs!codigo & "'"  ', , , 0
        If Not rsr.EOF Then reserva = s2n(rsr!sal)

        i = i + 1
        grillastock.TextMatrix(i, 0) = sSinNull(rs!DESCRIPCION)
        grillastock.TextMatrix(i, 1) = rs!codigo
        grillastock.TextMatrix(i, 2) = existencia
        grillastock.TextMatrix(i, 3) = reserva
        grillastock.TextMatrix(i, 4) = existencia - reserva
        grillastock.TextMatrix(i, 5) = s2n(rs!precio)
        grillastock.TextMatrix(i, 6) = 0 's2n(rs!PORBULTO)
        grillastock.TextMatrix(i, 7) = (rs!SubG)
        
        rs.MoveNext
    Wend
    rs.Close
    
  
    Set rs = Nothing
    Set rsr = Nothing

End Sub

Private Sub grillastock_Click()
    Dim rsped As New ADODB.Recordset
    
    mRow = grillastock.Row
    mCol = grillastock.Col
    grillaped.rows = 1
    If Trim(grillastock.TextMatrix(grillastock.Row, 3)) <> "0" Then
        rsped.Open "SELECT Pedidos_Clientes.numero,Pedidos_Clientes.fecha, Clientes.descripcion, ItemPedidoCliente.Saldo" _
        & " FROM (Pedidos_Clientes INNER JOIN ItemPedidoCliente ON Pedidos_Clientes.numero = ItemPedidoCliente.PEDIDO) " _
        & "INNER JOIN Clientes ON Pedidos_Clientes.cliente = Clientes.codigo " _
        & "WHERE Pedidos_Clientes.activo=1 AND ItemPedidoCliente.Producto='" & Trim(grillastock.TextMatrix(grillastock.Row, 1)) & "' and itempedidocliente.saldo > 0", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rsped.EOF Then
            Do While Not rsped.EOF
                grillaped.AddItem rsped!numero & Chr(9) & rsped!fecha & Chr(9) & rsped!DESCRIPCION & Chr(9) & rsped!saldo
                rsped.MoveNext
            Loop
        End If
    
        rsped.Close
        Set rsped = Nothing
    End If
End Sub

Private Sub txtBuscar_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim r As Integer
    Dim bRow As Long, bCol As Long
    
    r = mRow + 1
    grillastock.Col = mCol
    grillastock.Row = mRow
    grillastock.CellBackColor = vbWhite
    If Shift = 0 Then
        If KeyCode = 114 Then
            If r < grillastock.rows Then
                Do While InStr(1, LCase(grillastock.TextMatrix(r, mCol)), Trim(txtbuscar), 1) = 0 And r + 1 < grillastock.rows
                    r = r + 1
                Loop
                If InStr(1, LCase(grillastock.TextMatrix(mRow, mCol)), Trim(txtbuscar), 1) <> 0 Then
                        grillastock.Row = r '- 1
                         
                        grillastock.CellBackColor = vbMagenta 'vbRed
                        mCol = grillastock.Col
                        mRow = grillastock.Row
                        
                        grillastock.TopRow = maximo(1, r - 1)
                Else
                    grillastock.Row = bRow
                    grillastock.Col = bCol
                    grillastock.CellBackColor = vbCyan 'vbMagenta
                End If
            End If
        End If
    End If
End Sub

Private Sub txtbuscar_KeyPress(KeyAscii As Integer)

    On Error GoTo fin
    
    Dim r As Integer, i As Integer
    Dim bRow As Long, bCol As Long
    Dim lena As Long
    Dim lenbus As Long
    Dim lcabus As String
    
    If KeyAscii = 13 Then
        lenbus = Len(Trim(txtbuscar))
        lcabus = LCase(txtbuscar)
        With grillastock
            
            bRow = .Row
            bCol = .Col
            If bRow > 0 Then
                .Row = bRow
                .Col = bCol
                .CellBackColor = vbWhite
            End If
            
            r = 1
            If lenbus > 0 Then
                While InStr(1, LCase(.TextMatrix(r, bCol)), lcabus, 1) = 0 And r + 1 < .rows
                    r = r + 1
                Wend

                If InStr(1, LCase(.TextMatrix(r, bCol)), lcabus, 1) <> 0 Then
                    .Row = r '- 1
                    .CellBackColor = vbMagenta
                    mCol = .Col
                    mRow = .Row
                    .TopRow = maximo(1, r - 1)
                Else
                    .Row = bRow
                    .Col = bCol
                    .CellBackColor = vbCyan
                End If
            End If
        End With
    End If
fin:
End Sub

Private Function maximo(a, b)
    maximo = IIf(a > b, a, b)
End Function
Private Function minimo(a, b)
    minimo = IIf(a < b, a, b)
End Function


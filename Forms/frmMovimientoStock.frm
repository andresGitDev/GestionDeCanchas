VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMovimientoStock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimiento de Stock Con Parte de Produccion"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9810
   Icon            =   "frmMovimientoStock.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   9810
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCodigoProd2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1620
      TabIndex        =   18
      Top             =   675
      Width           =   1815
   End
   Begin VB.TextBox txtDescripcionProd2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4140
      TabIndex        =   17
      Top             =   675
      Width           =   5550
   End
   Begin VB.CommandButton cmdayudaprod2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3540
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   675
      Width           =   375
   End
   Begin Gestion.ucXls ucXls1 
      Height          =   810
      Left            =   4740
      TabIndex        =   15
      Top             =   7035
      Width           =   915
      _extentx        =   1614
      _extenty        =   661
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
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
      Height          =   375
      Left            =   6750
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7035
      Width           =   975
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7035
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   7020
      Left            =   75
      TabIndex        =   0
      Top             =   -15
      Width           =   9705
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6585
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1080
         Width           =   975
      End
      Begin VSFlex7LCtl.VSFlexGrid GrillaDetalle 
         Height          =   2955
         Left            =   60
         TabIndex        =   12
         Top             =   1500
         Width           =   9555
         _cx             =   16854
         _cy             =   5212
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
         Rows            =   2
         Cols            =   2
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
      Begin VB.CommandButton cmdayudaprod 
         BackColor       =   &H00FFFFFF&
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   300
         Width           =   375
      End
      Begin VB.TextBox txtDescripcionProd 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4080
         TabIndex        =   3
         Top             =   270
         Width           =   5535
      End
      Begin VB.TextBox txtCodigoProd 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   300
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker dtfechad 
         Height          =   255
         Left            =   1545
         TabIndex        =   4
         Top             =   1155
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   58785793
         CurrentDate     =   38252
      End
      Begin MSComCtl2.DTPicker dtfechah 
         Height          =   255
         Left            =   4485
         TabIndex        =   5
         Top             =   1155
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   58785793
         CurrentDate     =   38252
      End
      Begin VSFlex7LCtl.VSFlexGrid GrillaDetalleFactura 
         Height          =   2055
         Left            =   60
         TabIndex        =   13
         Top             =   4845
         Width           =   9525
         _cx             =   16801
         _cy             =   3625
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
         Rows            =   2
         Cols            =   2
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Producto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   120
         TabIndex        =   19
         Top             =   645
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Detalle de la Factura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   75
         TabIndex        =   14
         Top             =   4560
         Width           =   1860
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Producto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   810
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha Hasta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3045
         TabIndex        =   10
         Top             =   1155
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha Desde"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   105
         TabIndex        =   9
         Top             =   1155
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmMovimientoStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CONST_AJUSTE_CLI_DEBITO = "ACD"
Private Const CONST_AJUSTE_CLI_CREDITO = "ACC"
Private Const CONST_FACTURAS_A = "FAA"
Private Const CONST_FACTURAS_B = "FAB"
Private Const CONST_NOTAS_DEBITOS_A = "NDA"
Private Const CONST_NOTAS_CREDITOS_A = "NCA"
Private Const CONST_NOTAS_CREDITOS_B = "NCB"
Private Const CONST_RECIBOS = "RAA"
Private Const CONST_RECIBOS_IMPUTADOS = "REC"


Private Const const_REMITO_VENTA = "RMV"
Private Const const_REMITO_COMPRA = "RMC"
Private Const const_MOVIMIENTO_MANUAL = "MVM"
Private Const const_SALDO_INICIAL = ">Inicial"
Private Const const_PARTE_PRODUCCION_P = "PPP"
Private Const const_PARTE_PRODUCCION_C = "PPC"

Private Sub AgregarEnGrilla(fecha As Date, TipoComprobante As String, NroComprobante As Long, _
                            cantidad As Double, saldo As Double)
    With GrillaDetalle
    .Row = 1
    .Col = 0
    If .Text = "" Then
        .Col = 0
        .Text = fecha
        .Col = 1
        .Text = TipoComprobante
        .Col = 2
        .Text = NroComprobante
        .Col = 3
        .Text = cantidad
        .Col = 4
        .Text = saldo

    Else
        .AddItem fecha & Chr(9) & _
                TipoComprobante & Chr(9) & _
                NroComprobante & Chr(9) & _
                cantidad & Chr(9) & _
                saldo
    End If
    End With
End Sub


Private Function CalcularSaldoAnterior(CodigoProducto As String, fechahasta As Date) As Double
    Dim rsAux As New ADODB.Recordset
    Dim Consulta As String
    Dim cantidad As Double

    cantidad = 0

    'TABLA DE PARTE DE PRODUCCION
    rsAux.Open "select * from itempartesproduccion where producto='" & CodigoProducto & "'", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If rsAux.EOF = True And rsAux.BOF = True Then
            'es componente
            Set rsAux = Nothing
            rsAux.Open "select * from formulasdetalle where componente='" & CodigoProducto & "'", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            If rsAux.EOF = True And rsAux.BOF = True Then
                'no existe el prod en partes de produccion
            Else
                'es componente
                Consulta = "select sum(f.cantidad) from formulasdetalle as f " & _
                        "inner join itempartesproduccion as i on i.parte=f.codigo_parte " & _
                        "INNER JOIN partesproduccion AS p ON i.parte=p.nro " & _
                        "where p.ACTIVO = 1 and f.activo=1 and f.componente = '" & CodigoProducto & "' and p.confirmacion <=" & ssFecha(fechahasta)
                        
                'Consulta = "select sum(f.cantidad) from itempartesproduccion as i " & _
                        "inner join formulasdetalle as f on i.parte=f.codigo_parte " & _
                        "INNER JOIN partesproduccion AS p ON i.parte=p.nro " & _
                        "where p.ACTIVO = 1 and f.activo=1 and f.componente = '" & CodigoProducto & "' and p.confirmacion <=" & ssFecha(FechaHasta)
                        
                cantidad = cantidad - s2n(obtenerDeSQL(Consulta))
                
            End If
            Set rsAux = Nothing
        Else
            'es producto
            Consulta = "select sum(i.cantidad) from itempartesproduccion as i " & _
                        "INNER JOIN partesproduccion AS p ON i.parte=p.nro " & _
                        "where p.ACTIVO = 1 and i.PRODUCTO = '" & CodigoProducto & "' and p.confirmacion <=" & ssFecha(fechahasta)
            
            cantidad = cantidad + s2n(obtenerDeSQL(Consulta))
            
            Set rsAux = Nothing
        End If
    
    'TABLA REMITO COMPRA
    Consulta = "Select sum(D.CANTIDAD) as CantidadTotal " & _
                "from REMITOCOMPRADETALLE as D INNER JOIN REMITOCOMPRA AS R ON R.CODIGO = D.CODIGOREMITO " & _
                "where R.ACTIVO = 1 and D.PRODUCTO = '" & CodigoProducto & "'  and " & _
                "R.FECHA <= " & ssFecha(fechahasta)
    'rsAux.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockReadOnly
    'If Not rsAux.EOF Then
    '    rsAux.MoveFirst
    '    If Not IsNull(rsAux!cantidadtotal) Then cantidad = cantidad + rsAux!cantidadtotal
    'End If
    'rsAux.Close
    'Set rsAux = Nothing
    cantidad = cantidad + s2n(obtenerDeSQL(Consulta))
    
    
    'TABLA REMITO DIFERENCIA STOCK
    Consulta = "select sum(D.CANTIDAD) as CantidadTotal " & _
                "from ITEMREMITODIFERENCIASTOCK as D INNER JOIN REMITODIFERENCIASTOCK AS R ON R.MOVIMIENTOINTERNO = D.NUMERO " & _
                "where  D.PRODUCTO = '" & CodigoProducto & "' and FECHA <= " & ssFecha(fechahasta)
'    rsAux.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
'    If Not rsAux.EOF Then
'        rsAux.MoveFirst
'        If Not IsNull(rsAux!cantidadtotal) Then cantidad = cantidad + rsAux!cantidadtotal
'    End If
'    rsAux.Close
'    Set rsAux = Nothing
    cantidad = cantidad - s2n(obtenerDeSQL(Consulta))

    
    'TABLA REMITO VENTA
    Consulta = "Select sum(D.CANTIDAD) as CantidadTotal " & _
                "from REMITOVENTADETALLE as D INNER JOIN REMITOVENTA AS R ON R.numero = D.numero " & _
                "where D.PRODUCTO = '" & CodigoProducto & "' and R.FECHA <= " & ssFecha(fechahasta) & " and r.anulado = 0 and (D.FACTURAR > 0 or r.factura = 0)"
'    rsAux.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
'    If Not rsAux.EOF Then
'        rsAux.MoveFirst
'        If Not IsNull(rsAux!cantidadtotal) Then cantidad = cantidad - rsAux!cantidadtotal
'    End If
'    rsAux.Close
'    Set rsAux = Nothing
    cantidad = cantidad - s2n(obtenerDeSQL(Consulta))

    
    'TABLA FACTURA VENTA
    Consulta = "Select sum(D.CANTIDAD) as CantidadTotal " & _
                "from FACTURAVENTADETALLE as D INNER JOIN FACTURAVENTA AS R ON R.CODIGO = D.CODIGOFACTURA " & _
                "where (R.ACTUALIZASTOCK = 1 or d.nroremito<>0) and D.PRODUCTO = '" & CodigoProducto & "' and R.FECHA <= " & ssFecha(fechahasta) & " and activo = 1"
'    rsAux.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
'    If Not rsAux.EOF Then
'        rsAux.MoveFirst
'        If Not IsNull(rsAux!cantidadtotal) Then cantidad = cantidad - rsAux!cantidadtotal
'    End If
'    rsAux.Close
'    Set rsAux = Nothing
    cantidad = cantidad - s2n(obtenerDeSQL(Consulta))
    
    CalcularSaldoAnterior = cantidad
        
End Function

Private Sub CalcularSaldo()
Dim rsAux As New ADODB.Recordset
Dim Consulta As String
Dim cantidad As Double
Dim a As String
Dim CodigoProd As String

'    cantidad = 0
    Consulta = "Select * From MOVIMIENTO_STOCK_TEMP order by producto , fecha"
    rsAux.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic

    If rsAux.EOF And rsAux.BOF Then
    Else
    rsAux.MoveFirst
        While Not rsAux.EOF
            cantidad = cantidad + rsAux!cantidad
            '----------abrir adodb y actualizar con sql me parece al reverendo
            'Consulta = "Update MOVIMIENTO_STOCK_TEMP Set SALDO = '" & cantidad & "' Where ID = " & rsAux!id
            'DataEnvironment1.Sistema.Execute Consulta
            a = cantidad
            rsAux!saldo = a
            rsAux.Update
            If rsAux!NroComprobante = "-" Then cantidad = 0
            rsAux.MoveNext
        Wend
    End If
    Set rsAux = Nothing
    
End Sub

Private Sub AgregarEnTabla(fecha As Date, TipoComprobante As String, NroComprobante As Variant, _
                            cantidad As Double, saldo As Double, producto As String)
Dim Consulta As String

    Consulta = "Insert into MOVIMIENTO_STOCK_TEMP (FECHA, TIPOCOMPROBANTE, NROCOMPROBANTE, CANTIDAD, SALDO, PRODUCTO) " & _
                    "values (" & ssFecha(fecha) & ", '" & TipoComprobante & "', '" & NroComprobante & "', ' " & _
                    cantidad & "', '" & saldo & "', '" & producto & "')"
    DataEnvironment1.Sistema.Execute Consulta
    
End Sub

Private Sub cmdayudaprod_Click()
    frmBuscar.MostrarSql "Select CODIGO as [ Codigo                       ], DESCRIPCION  as [ Descripcion                                                        ]From PRODUCTO Where ACTIVO = 1 Order By CODIGO"
    txtCodigoProd.Text = frmBuscar.resultado
    txtDescripcionProd.Text = frmBuscar.resultado(2)
End Sub

Private Sub cmdayudaprod2_Click()
    frmBuscar.MostrarSql "Select CODIGO as [ Codigo                       ], DESCRIPCION  as [ Descripcion                                                        ]From PRODUCTO Where ACTIVO = 1 Order By CODIGO"
    txtCodigoProd2.Text = frmBuscar.resultado
    txtDescripcionProd2.Text = frmBuscar.resultado(2)
End Sub

Private Sub cmdBuscar_Click()
Dim cPro As String, rsPro As New ADODB.Recordset, i As Long, redim_r As Long
Dim codigoD As String, codigoH As String
If txtCodigoProd.Text <> "" And txtCodigoProd2.Text <> "" Then
    DataEnvironment1.Sistema.Execute "Delete From MOVIMIENTO_STOCK_TEMP"
    
    codigoD = "'" & Trim(txtCodigoProd.Text) & "'"
    codigoH = "'" & Trim(txtCodigoProd2.Text) & "'"
    cPro = "select * from producto where activo=1 and codigo >=" & codigoD & " and codigo <=" & codigoH & " order by codigo"
    rsPro.Open cPro, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
    redim_r = 0
    ReDim revoke_producto(redim_r)
    With rsPro
        If .EOF And .BOF Then
            MsgBox "Sin productos cargados.", vbCritical
            Exit Sub
        Else
            .MoveFirst
            For i = 0 To .RecordCount - 1
                If cmdCargarEnTemp(!codigo) = False Then
                    ReDim Preserve revoke_producto(redim_r)
                    revoke_producto(redim_r) = !codigo
                    redim_r = redim_r + 1
                End If
                .MoveNext
            Next
        End If
    End With
    
     revoke_vacios
    
    CalcularSaldo
    LlenarGrilla GrillaDetalle, "Select FECHA, TIPOCOMPROBANTE AS 'TIPO DE COMPROBANTE', NROCOMPROBANTE AS 'NRO', CANTIDAD, SALDO " & _
                                    "From MOVIMIENTO_STOCK_TEMP order by producto, fecha", True   'Order By NROCOMPROBANTE
                                    
    If GrillaDetalle.rows > 1 Then
        With GrillaDetalle
            .ColWidth(0) = 1000
            .ColWidth(1) = 4000
            .ColWidth(2) = 1000
            .ColWidth(3) = 1000
            .ColWidth(4) = 1000
            'marco iniciales
            For i = 1 To .rows - 1
                If Mid(.TextMatrix(i, 1), 1, 1) = ">" Then
                    .cell(flexcpFontBold, i, 1) = True
                End If
            Next
        End With
    End If
    
Else
    MsgBox "Llenar con minimo dos codigos de productos.", vbCritical
End If

End Sub

Private Function revoke_vacios()
Dim p As Long, cadena_kill As String
    For p = 0 To UBound(revoke_producto)
        If revoke_producto(p) <> 0 Then
            cadena_kill = "delete from MOVIMIENTO_STOCK_TEMP where producto = '" & revoke_producto(p) & "'"
            DataEnvironment1.Sistema.Execute cadena_kill
        End If
    Next
End Function


Private Function cmdCargarEnTemp(en_vista As String) As Boolean

'select D.CODIGOREMITO, R.FECHA, D.PRODUCTO, D.CANTIDAD
'from REMITOCOMPRADETALLE as D INNER JOIN REMITOCOMPRA AS R ON R.CODIGO = D.CODIGOREMITO
'where R.ACTIVO = 1 and
'    D.PRODUCTO = '0010011355' and
'    R.FECHA BETWEEN CONVERT (DATETIME, '11-04-2004') AND CONVERT(DATETIME, '11-24-2004')
'
'
'
'select R.COMPROBANTE, R.NROCOMPROBANTE, R.FECHA, D.PRODUCTO, D.CANTIDAD
'from ITEMREMITODIFERENCIASTOCK as D INNER JOIN REMITODIFERENCIASTOCK AS R ON R.MOVIMIENTOINTERNO = D.NUMERO
'where   D.PRODUCTO = '0020011906' and
'    R.FECHA BETWEEN CONVERT (DATETIME, '1-1-2002') AND CONVERT(DATETIME, '11-1-2004')
'
'
'
'select R.NUMERO, R.FECHA, D.PRODUCTO, D.FACTURAR
'from REMITOVENTADETALLE as D INNER JOIN REMITOVENTA AS R ON R.NUMERO = D.NUMERO
'where D.PRODUCTO = '001001USAC72H5D'
'    and R.FECHA  between convert(datetime , '01-15-00', 1) AND convert(datetime , '01-15-05', 1)
'    AND D.FACTURAR > 0  --> HAGO ESTA PREGUNTA PORQUE SI ES 0 EL MOVIMIENTO LO ESTOY MOSTRANDO EN LA FACTURA
'
'
'select R.TIPODOC, R.NROFACTURA, R.FECHA, D.PRODUCTO, D.CANTIDAD
'from FACTURAVENTADETALLE as D INNER JOIN FACTURAVENTA AS R ON R.CODIGO = D.CODIGOFACTURA
'where (R.ACTUALIZASTOCK = 1 or D.NROREMITO > 0) and
'    D.PRODUCTO = '00300170406' and
'    R.FECHA BETWEEN CONVERT (DATETIME, '1-1-2002') AND CONVERT(DATETIME, '11-1-2003')
Dim tiene_algo As Integer
Dim Consulta As String
Dim CodProd As String, CodProd2 As String
Dim rsmov As New ADODB.Recordset
Dim rsAux As New ADODB.Recordset
Dim SaldoProd As Double, SaldoProd2 As Double
Dim cant As Double

    If en_vista <> "" Then
        tiene_algo = 0
        'DataEnvironment1.Sistema.Execute "Delete From MOVIMIENTO_STOCK_TEMP"
        
        CodProd = Trim(en_vista)
             
        SaldoProd = CalcularSaldoAnterior(CodProd, dtfechad.Value - 1)
        
        AgregarEnTabla dtfechad.Value - 1, const_SALDO_INICIAL & " " & ObtenerDescripcionS("PRODUCTO", Trim(CodProd)), 0, SaldoProd, 0, CodProd
        
        CodProd2 = "select * from partesproduccion p where p.producido='" & CodProd & "'" _
            & " and p.ACTIVO = 1 and p.confirmacion " & ssBetween(dtfechad.Value, dtfechah.Value)
            
        rsAux.Open CodProd2, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If rsAux.EOF = True And rsAux.BOF = True Then
        Else
            With rsAux
                .MoveFirst
                tiene_algo = 2
                While Not .EOF
                    AgregarEnTabla !Confirmacion, const_PARTE_PRODUCCION_P, !Nro, CDbl(!cantidad), 0, CodProd
                    .MoveNext
                Wend
            End With
        End If
        Set rsAux = Nothing
        
        CodProd2 = "select f.cantidad as c,p.* from formulasdetalle f inner join partesproduccion p on p.nro=f.codigo_parte where componente='" & CodProd & "'" _
            & " and p.ACTIVO = 1 and f.activo=1 and p.confirmacion " & ssBetween(dtfechad.Value, dtfechah.Value)
        rsAux.Open CodProd2, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If rsAux.EOF = True And rsAux.BOF = True Then
        Else
            With rsAux
                .MoveFirst
                tiene_algo = 2
                While Not .EOF
                    AgregarEnTabla !Confirmacion, const_PARTE_PRODUCCION_C, !Nro, -CDbl(!C), 0, CodProd
                    .MoveNext
                Wend
            End With
        End If
        Set rsAux = Nothing
               
'        'TABLA DE PARTE DE PRODUCCION
'        CodProd2 = "select * from partesproduccion where producido='" & CodProd & "'"
'        rsAux.Open CodProd2, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
'        If rsAux.EOF = True And rsAux.BOF = True Then
'            'es componente
'            Set rsAux = Nothing
'            CodProd2 = "select * from formulasdetalle where componente='" & CodProd & "'"
'            rsAux.Open CodProd2, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
'            If rsAux.EOF = True And rsAux.BOF = True Then
'                'no existe el prod en partes de produccion
'            Else
'                'es componente
'                Consulta = "select p.confirmacion as confir,i.parte as parte,f.cantidad as cant from formulasdetalle as f " & _
'                            "INNER JOIN itempartesproduccion as i on i.parte=f.codigo_parte " & _
'                            "inner join partesproduccion AS p ON i.parte=p.nro " & _
'                            "where p.ACTIVO = 1 and f.componente='" & CodProd & "' and f.activo=1 and p.confirmacion " & ssBetween(dtfechad.Value, dtfechah.Value) & " group by p.confirmacion, i.parte , f.cantidad"
'                rsmov.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
'                If Not rsmov.EOF Then
'                    rsmov.MoveFirst
'                    tiene_algo = 1
'                    While Not rsmov.EOF
'                        AgregarEnTabla rsmov!Confir, const_PARTE_PRODUCCION_C, rsmov!parte, CDbl(-rsmov!cant), 0, CodProd
'                        rsmov.MoveNext
'                    Wend
'                End If
'            End If
'        Else
'            'es producto
'            Consulta = "select p.confirmacion as confir,i.codigo_parte as parte,i.cantidad as cant from formulasdetalle as i " & _
'                        "INNER JOIN partesproduccion AS p ON i.codigo_parte=p.nro " & _
'                        "where p.ACTIVO = 1 and i.activo=1 and i.componente = '" & CodProd & "' and p.confirmacion " & ssBetween(dtfechad.Value, dtfechah.Value)
'            rsmov.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
'            If Not rsmov.EOF Then
'                rsmov.MoveFirst
'                tiene_algo = 2
'                While Not rsmov.EOF
'                    AgregarEnTabla rsmov!Confir, const_PARTE_PRODUCCION_P, rsmov!parte, CDbl(rsmov!cant), 0, CodProd
'                    rsmov.MoveNext
'                Wend
'            End If
'        End If
'            Set rsmov = Nothing
'            Set rsAux = Nothing

        'TABLA REMITO COMPRA
        Consulta = "select R.NroRemito, R.FECHA, D.PRODUCTO, D.CANTIDAD " & _
                    "from REMITOCOMPRADETALLE as D INNER JOIN REMITOCOMPRA AS R ON R.CODIGO = D.CODIGOREMITO " & _
                    "where R.ACTIVO = 1 and " & _
                    "D.PRODUCTO = '" & CodProd & "' and " & _
                    "R.FECHA " & ssBetween(dtfechad.Value, dtfechah.Value)
        rsmov.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If rsmov.EOF And rsmov.BOF Then
        Else
        rsmov.MoveFirst
        tiene_algo = 3
            While Not rsmov.EOF
                AgregarEnTabla rsmov!fecha, const_REMITO_COMPRA, rsmov!NroRemito, CDbl(rsmov!cantidad), 0, CodProd
                rsmov.MoveNext
            Wend
        End If
        Set rsmov = Nothing
        'TABLA REMITO DIFERENCIA STOCK
        Consulta = "select R.COMPROBANTE, R.MovimientoInterno , R.NROCOMPROBANTE, R.FECHA, D.PRODUCTO, D.CANTIDAD " _
            & " From ITEMREMITODIFERENCIASTOCK as D INNER JOIN REMITODIFERENCIASTOCK AS R " _
            & " ON R.movimientointerno = D.NUMERO " _
            & " Where r.activo=1 and D.PRODUCTO = '" & CodProd & "' and " _
            & " R.FECHA " & ssBetween(dtfechad.Value, dtfechah.Value)
        rsmov.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockReadOnly
        If rsmov.EOF And rsmov.BOF Then
        Else
        rsmov.MoveFirst
        tiene_algo = 4
            While Not rsmov.EOF
                AgregarEnTabla rsmov!fecha, const_MOVIMIENTO_MANUAL, CLng(rsmov!MovimientoInterno), CDbl(rsmov!cantidad), 0, CodProd
                rsmov.MoveNext
            Wend
        End If
        Set rsmov = Nothing
            
        
        
' ESTO ESTABA COMENTADO, YO LO SAQUE PORQUE CREO QUE PARA GREEN OIL FUNCIONARA MEJOR ASI (LAURA)
'        '******************************************************
'        ' revisando........
'        'TABLA REMITO VENTA
        Consulta = "select R.NUMERO, R.FECHA, D.PRODUCTO, D.FACTURAR, r.factura, d.cantidad  " & _
                    " from REMITOVENTADETALLE as D INNER JOIN REMITOVENTA AS R ON R.NUMERO = D.NUMERO " & _
                    " where D.PRODUCTO = '" & CodProd & "' and (D.FACTURAR > 0 or r.factura = 0)  And " & _
                    " R.FECHA " & ssBetween(dtfechad.Value, dtfechah.Value) & " and R.anulado=0"
        rsmov.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockReadOnly
        If rsmov.EOF And rsmov.BOF Then
        Else
        rsmov.MoveFirst
        tiene_algo = 5
            While Not rsmov.EOF
                cant = CDbl(rsmov!cantidad) - CDbl(rsmov!cantidad) * 2
                If rsmov!Factura = 0 Then cant = -rsmov!cantidad
                AgregarEnTabla rsmov!fecha, const_REMITO_VENTA, CLng(rsmov!numero), cant, 0, CodProd
                rsmov.MoveNext
            Wend
        End If
        Set rsmov = Nothing

        'TABLA FACTURA VENTA
        Consulta = "select R.TIPODOC, R.NROFACTURA, R.FECHA, D.PRODUCTO, D.CANTIDAD " & _
                    "from FACTURAVENTADETALLE as D INNER JOIN FACTURAVENTA AS R ON R.CODIGO = D.CODIGOFACTURA " & _
                    "where (R.ACTUALIZASTOCK = 1 OR D.NROREMITO > 0) and D.PRODUCTO = '" & CodProd & "' and " & _
                    "R.FECHA " & ssBetween(dtfechad.Value, dtfechah.Value) & " and activo=1"
        rsmov.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If rsmov.EOF And rsmov.BOF Then
        Else
        rsmov.MoveFirst
        tiene_algo = 6
            While rsmov.EOF = False
                If Trim(rsmov!TIPODOC) = CONST_FACTURAS_A Or Trim(rsmov!TIPODOC) = CONST_FACTURAS_B Then
                    cant = CDbl(rsmov!cantidad) - CDbl(rsmov!cantidad) * 2
                ElseIf CONST_NOTAS_CREDITOS_A = Trim(rsmov!TIPODOC) Or CONST_NOTAS_CREDITOS_B = Trim(rsmov!TIPODOC) Then
                    cant = CDbl(rsmov!cantidad) ''- CDbl(rsmov!cantidad) * 2
                Else
                    cant = CDbl(rsmov!cantidad) - CDbl(rsmov!cantidad) * 2
                End If
                AgregarEnTabla rsmov!fecha, rsmov!TIPODOC, CLng(rsmov!nrofactura), cant, 0, CodProd
                rsmov.MoveNext
            Wend
        End If
        Set rsmov = Nothing
        ' revisando........
'        '******************************************************
        'remitos no facturados
'
'        Dim ss As String, tempo
'        Consulta = "select R.numero, R.FECHA, D.PRODUCTO, d.cantidad, d.codigo  " & _
'                    " from REMITOVENTADETALLE as D INNER JOIN REMITOVENTA AS R ON R.NUMERO = D.NUMERO " & _
'                    " where D.PRODUCTO = '" & CodProd & "' And R.FECHA " & ssBetween(dtfechad.Value, dtfechah.Value) & " and r.anulado = 0"
'        With rsmov
'            .Open Consulta, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'            While Not .EOF
'                ss = "select fecha, f.tipodoc, f.nrofactura,  cantidad, producto from facturaventadetalle as d inner join facturaventa as f  on f.codigo = d.CodigoFactura " _
'                    & " where d.item_p_r = " & !codigo & " and f.activo = 1 "
'                tempo = obtenerDeSQL(ss)
'                If IsEmpty(tempo) Then
'                    'remitos sin facura
'                    AgregarEnTabla rsmov!fecha, const_REMITO_VENTA, CLng(rsmov!numero), -!cantidad, 0
'                Else
'                    'factura sobre remito
'                    AgregarEnTabla CDate(tempo(0)), CStr(tempo(1)), tempo(2), -!cantidad, 0
'''                    AgregarEnTabla CDate(tempo(0)), CStr(tempo(1)), tempo(2), -tempo(3), 0
'''                    If tempo(3) <> !cantidad Then
'''                        'difeerncia por si facturo menos
'''                        AgregarEnTabla rsmov!fecha, const_REMITO_VENTA, CLng(rsmov!numero), -(!cantidad - tempo(3)), 0
'''                    End If
'                End If
'                .MoveNext
'            Wend
'            .Close
'            'factura sin remito
'            Consulta = "select R.TIPODOC, R.NROFACTURA, R.FECHA, D.PRODUCTO, D.CANTIDAD " _
'                & "from FACTURAVENTADETALLE as D INNER JOIN FACTURAVENTA AS R ON R.CODIGO = D.CODIGOFACTURA " _
'                & "where R.ACTUALIZASTOCK = 1  and D.PRODUCTO = '" & CodProd & "' and " _
'                & "R.FECHA " & ssBetween(dtfechad.Value, dtfechah.Value) & " and R.Activo = 1"
'            .Open Consulta, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'            While Not .EOF
'                AgregarEnTabla !fecha, !TIPODOC, CLng(!nrofactura), -s2n(!cantidad), 0
'                .MoveNext
'            Wend
'            .Close
'        End With
'        '******************************************************
'
'
        
        
        'esto estaba antes de que pida una consulta de prod desde hasta'raul 24/8
         'CalcularSaldo
'        LlenarGrilla GrillaDetalle, "Select FECHA, TIPOCOMPROBANTE AS 'TIPO DE COMPROBANTE', NROCOMPROBANTE AS 'NRO DE COMPROBANTE', CANTIDAD, SALDO " & _
                                    "From MOVIMIENTO_STOCK_TEMP ", True  'Order By NROCOMPROBANTE
'        With GrillaDetalle
'            .ColWidth(0) = 1000
'            .ColWidth(1) = 2000
'            .ColWidth(2) = 2000
'            .ColWidth(3) = 1000
'            .ColWidth(4) = 1000
'        End With
        
'        Consulta = "Select * From MOVIMIENTO_STOCK_TEMP Order By FECHA"
'        With rptMovimientoStock2
'            .Data.Connection = daTaenvironment1.Sistema
'            .Data.Source = Consulta
'            .lblCodigo.Caption = CodProd
'            .lblDescripcion.Caption = ObtenerDescripcionS("PRODUCTO", CodProd)
'
'            .fieFecha.DataField = "FECHA"
'            .fieTipoComprobante.DataField = "TIPOCOMPROBANTE"
'            .fieNroComprobante.DataField = "NROCOMPROBANTE"
'            .fieCantidad.DataField = "CANTIDAD"
'            .fieSaldo.DataField = "SALDO"
'
'            .Show
'
'        End With
    Else
        'MsgBox "Debe ingresar un codigo de producto", vbOKOnly, "Atencion"
    End If
     
    AgregarEnTabla Date + 1, "->>SALDO FINAL:" & ObtenerDescripcionS("PRODUCTO", Trim(CodProd)), "-", 0, 0, CodProd
    If tiene_algo = 0 Then
        cmdCargarEnTemp = False
    Else
        cmdCargarEnTemp = True
    End If
End Function



Private Sub cmdImprimir_Click()
Dim Consulta As String
    Consulta = "Select * From MOVIMIENTO_STOCK_TEMP order by id" 'Order By FECHA
    With rptMovimientoStock2
        .Data.Connection = DataEnvironment1.Sistema
        .Data.Source = Consulta
        .lblFechaReporte.caption = Date
        
        .fieFecha.DataField = "FECHA"
        .fieTipoComprobante.DataField = "TIPOCOMPROBANTE"
        .fieNroComprobante.DataField = "NROCOMPROBANTE"
        .fieCantidad.DataField = "CANTIDAD"
        .fieSaldo.DataField = "SALDO"
        
        .Show
    
    End With

End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub

Private Sub Form_Load()
    dtfechad.Value = "01/01/" & Year(Date)
    dtfechah.Value = Date
    
'    CentrarMe frmMovimientoStock
    ucXls1.ini GrillaDetalle, "bsMovStock.xls"
End Sub

Private Sub GrillaDetalle_Click()
Dim TIPODOC As String
Dim NroDoc As String
Dim CodInt As Long

    With GrillaDetalle
        If .TextMatrix(.Row, 0) <> "" And .Row <> 0 Then
            TIPODOC = Trim(.TextMatrix(.Row, 1))
            If .TextMatrix(.Row, 2) <> "" Then NroDoc = .TextMatrix(.Row, 2)
         
            LimpiarGrilla GrillaDetalleFactura
            
            Select Case TIPODOC
                Case CONST_FACTURAS_A, CONST_FACTURAS_B
                    LlenarGrilla GrillaDetalleFactura, "Select D.CANTIDAD, D.PRODUCTO, D.DESCRIPCION, D.PRECIOUNITARIO AS 'PRECIO UNITARIO', " & _
                                                                    "D.PRECIOTOTAL AS 'PRECIO TOTAL', S.SERIE " & _
                                         "From FACTURAVENTADETALLE AS D " & _
                                            "left Join SERIES as S ON S.PRODUCTO=D.PRODUCTO AND S.NROCOMPROBANTE=D.NROFACTURA " & _
                                         "Where D.TIPODOC = '" & TIPODOC & "' AND D.NROFACTURA = " & NroDoc & _
                                         " Order By ID", False
                Case CONST_RECIBOS, CONST_RECIBOS_IMPUTADOS
                    CodInt = ObtenerDatoDB("RECIBOS", "NUMERO", NroDoc, "CODIGO")
                    LlenarGrilla GrillaDetalleFactura, "Select F.TIPODOC AS 'Tipo', F.NROFACTURA AS 'Numero', Importe " & _
                                                "From RECIBOSDETALLE AS R " & _
                                                    "Inner Join FACTURAVENTA AS F on F.CODIGO = R.FACTURAVENTA " & _
                                                "Where R.CODRECIBO = " & CodInt & " Order By R.CODIGO", True
                Case const_REMITO_COMPRA
                    LlenarGrilla GrillaDetalleFactura, "SELECT D.PRODUCTO, D.CANTIDAD, D.COSTO, D.ORDENCOMPRA " & _
                                                        "FROM REMITOCOMPRA AS R " & _
                                                            "INNER JOIN REMITOCOMPRADETALLE AS D ON R.CODIGO = D.CODIGOREMITO " & _
                                                        "WHERE R.codigo = '" & NroDoc & "'", True
                Case const_REMITO_VENTA
                    LlenarGrilla GrillaDetalleFactura, "Select PRODUCTO, CANTIDAD, PRECIO, PEDIDO " & _
                                                        "From REMITOVENTADETALLE " & _
                                                        "Where NUMERO = " & NroDoc, True
                
                Case CONST_NOTAS_DEBITOS_A, CONST_NOTAS_CREDITOS_A, CONST_NOTAS_CREDITOS_B
                
                Case CONST_AJUSTE_CLI_DEBITO, CONST_AJUSTE_CLI_CREDITO
                
                Case const_PARTE_PRODUCCION_P
                    'LlenarGrilla GrillaDetalleFactura, "SELECT D.PRODUCTO, D.CANTIDAD, D.COSTO, D.ORDENCOMPRA " & _
                    '                                    "FROM REMITOCOMPRA AS R " & _
                    '                                        "INNER JOIN REMITOCOMPRADETALLE AS D ON R.CODIGO = D.CODIGOREMITO " & _
                    '                                    "WHERE R.codigo = '" & NroDoc & "'", True
                    
                    '"select *,p.confirmacion as confir,i.parte as parte,i.cantidad as cant from itempartesproduccion as i " & _
                    '                                    "INNER JOIN partesproduccion AS p ON i.parte=p.nro " & _
                    '                                    "inner join formulasdetalle as f on i.parte=f.codigo_parte " & _
                    '                                    "where i.parte=" & NroDoc & " and p.ACTIVO = 1 and i.PRODUCTO = '" & CodProd & "' and f.codigo_articulo='1' and f.activo=1 and p.confirmacion " & ssBetween(dtfechad.Value, dtfechah.Value), True
                    LlenarGrilla GrillaDetalleFactura, "select p.confirmacion as CONFIRMACION,i.codigo_parte as PARTE,i.componente as PRODUCTO,i.cantidad as CANTIDAD from formulasdetalle as i " & _
                                                        "INNER JOIN partesproduccion AS p ON i.codigo_parte=p.nro " & _
                                                        "where p.nro=" & NroDoc & " and p.ACTIVO = 1 group by p.confirmacion, i.codigo_parte,  i.componente ,i.cantidad ", True

                Case const_PARTE_PRODUCCION_C
                    LlenarGrilla GrillaDetalleFactura, "select p.confirmacion as CONFIRMACION,p.nro as PARTE,f.componente as PRODUCTO,t.descripcion as DESCRIPCION,f.cantidad as CANTIDAD from partesproduccion as p " & _
                                                        "inner join formulasdetalle as f on p.nro=f.codigo_parte " & _
                                                        "inner join producto as t on t.codigo=f.componente " & _
                                                        "where p.nro=" & NroDoc & " and p.ACTIVO = 1 and f.activo=1 group by p.confirmacion, p.nro,  f.componente ,t.descripcion,f.cantidad ", True
                                    
            End Select
        End If
    End With
End Sub

Private Sub txtCodigoProd_LostFocus()
    If Trim(txtCodigoProd) <> "" Then
        txtDescripcionProd = ObtenerDescripcionS("PRODUCTO", Trim(txtCodigoProd.Text))
        If Trim(txtDescripcionProd) = "" Then
            MsgBox "El codigo del producto es inexistente", 48, "Atencion"
            txtCodigoProd.SetFocus
        End If
    End If
End Sub
'20050531 fix remito anulado

Private Sub txtCodigoProd2_LostFocus()
    If Trim(txtCodigoProd) <> "" Then
        txtDescripcionProd = ObtenerDescripcionS("PRODUCTO", Trim(txtCodigoProd.Text))
        If Trim(txtDescripcionProd) = "" Then
            MsgBox "El codigo del producto es inexistente", 48, "Atencion"
            txtCodigoProd.SetFocus
        End If
    End If
End Sub




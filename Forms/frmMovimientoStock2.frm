VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMovimientoStock2 
   Caption         =   "Movimiento de Stock"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   8295
   StartUpPosition =   3  'Windows Default
   Begin Gestion.ucXls ucXls1 
      Height          =   735
      Left            =   4560
      TabIndex        =   19
      Top             =   7320
      Width           =   855
      _extentx        =   1508
      _extenty        =   1296
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   7140
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8055
      Begin VB.TextBox txtCodigoProd 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   10
         Top             =   300
         Width           =   1815
      End
      Begin VB.TextBox txtDescripcionProd 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4080
         TabIndex        =   9
         Top             =   300
         Width           =   3495
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
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   300
         Width           =   375
      End
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
         Height          =   315
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1065
         Width           =   975
      End
      Begin VB.CommandButton CmdAyudaProdHasta 
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
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox TxtDescripcionProdHasta 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4080
         TabIndex        =   4
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox TxtCodigoProdHasta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Top             =   720
         Width           =   1815
      End
      Begin VSFlex7LCtl.VSFlexGrid GrillaDetalle 
         Height          =   2955
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   7755
         _cx             =   13679
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
      Begin MSComCtl2.DTPicker dtfechad 
         Height          =   255
         Left            =   1560
         TabIndex        =   11
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   58785793
         CurrentDate     =   38252
      End
      Begin MSComCtl2.DTPicker dtfechah 
         Height          =   255
         Left            =   4440
         TabIndex        =   12
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   58785793
         CurrentDate     =   38252
      End
      Begin VSFlex7LCtl.VSFlexGrid GrillaDetalleFactura 
         Height          =   2175
         Left            =   135
         TabIndex        =   13
         Top             =   4800
         Width           =   7755
         _cx             =   13679
         _cy             =   3836
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
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   1335
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
         Left            =   3000
         TabIndex        =   17
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Producto Desde"
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
         TabIndex        =   16
         Top             =   300
         Width           =   1470
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
         Left            =   135
         TabIndex        =   15
         Top             =   4485
         Width           =   1860
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Producto Hasta"
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
         TabIndex        =   14
         Top             =   720
         Width           =   1395
      End
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
      Left            =   5700
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7470
      Width           =   975
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
      Left            =   7020
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7470
      Width           =   975
   End
End
Attribute VB_Name = "frmMovimientoStock2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private ttMovStk As String


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
Private Const const_SALDO_INICIAL = "SI"


'Private rsTempMov As New ADODB.Recordset
'Private Sub AgregarEnGrilla(fecha As Date, TipoComprobante As String, NroComprobante As Long, _
'                            cantidad As Double, Saldo As Double, Optional Concepto As String)
'    With GrillaDetalle
'    .Row = 1
'    .Col = 0
'    If .Text = "" Then
'        .Col = 0
'        .Text = fecha
'        .Col = 1
'        .Text = TipoComprobante
'        .Col = 2
'        .Text = NroComprobante
'        .Col = 3
'        .Text = cantidad
'        .Col = 4
'        .Text = Saldo
'        .Col = 5
'        .Text = Concepto
'    Else
'        .AddItem fecha & Chr(9) & _
'                TipoComprobante & Chr(9) & _
'                NroComprobante & Chr(9) & _
'                cantidad & Chr(9) & _
'                Saldo & Chr(9) & _
'                Concepto
'    End If
'    End With
'End Sub


Private Function CalcularSaldoAnterior(CodigoProducto As String, fechahasta As Date) As Double
'    Dim rsAux As New ADODB.Recordset
    Dim Consulta As String
    Dim cantidad As Double

    'cantidad = 0
    
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
                "where D.PRODUCTO = '" & CodigoProducto & "' and FECHA <= " & ssFecha(fechahasta)
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
                "where D.PRODUCTO = '" & CodigoProducto & "' and R.FECHA <= " & ssFecha(fechahasta) & " and r.anulado = 0 "
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
                "where R.ACTUALIZASTOCK = 1 and D.PRODUCTO = '" & CodigoProducto & "' and R.FECHA <= " & ssFecha(fechahasta) & " and activo = 1"
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

Private Sub CalcularSaldo(CodProd)
    Dim rsAux As New ADODB.Recordset
    Dim Consulta As String
    Dim cantidad As Double
    Dim CodigoProd As String

'    cantidad = 0
    Consulta = "Select * From  " & ttMovStk & " where codigo= '" & CodProd & "' Order by FECHA"
    rsAux.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not rsAux.EOF Then rsAux.MoveFirst
    While Not rsAux.EOF
        cantidad = cantidad + rsAux!cantidad
        '----------abrir adodb y actualizar con sql me parece al reverendo
        'Consulta = "Update MOVIMIENTO_STOCK_TEMP Set SALDO = '" & cantidad & "' Where ID = " & rsAux!id
        'DataEnvironment1.Sistema.Execute Consulta
        rsAux!saldo = cantidad
        rsAux.Update
        
        rsAux.MoveNext
    Wend
    rsAux.Close
    Set rsAux = Nothing
End Sub

Private Sub AgregarEnTabla(CodProd As String, fecha As Date, TipoComprobante As String, NroComprobante As Variant, _
                            cantidad As Double, saldo As Double, Optional concepto As String, Optional desc As String)
    Dim Consulta As String

'    desc = sSinNull(obtenerDeSQL("select descripcion from producto where activo = 1 and codigo = '" & codprod & "'"))
    Consulta = "Insert into  " & ttMovStk & "  (CODIGO, FECHA, TIPOCOMPROBANTE, NROCOMPROBANTE, CANTIDAD, SALDO, concepto, Descripcion) " & _
                    "values ('| " & CodProd & "', " & ssFecha(fecha) & ", '" & TipoComprobante & "', '" & NroComprobante & "', ' " & _
                    cantidad & "', '" & saldo & "','" & concepto & "', '" & Replace(desc, "'", "´") & "'  )"
    DataEnvironment1.Sistema.Execute Consulta
    
'    With rsTempMov
'        .AddNew
'        !codigo = codprod
'        !fecha = fecha
'        !TipoComprobante = TipoComprobante
'        !NroComprobante = NroComprobante
'        !cantidad = cantidad
'        !Saldo = Saldo
'        !concepto = concepto
'        !descripcion = desc
'        .Update
'    End With
End Sub

Private Sub cmdayudaprod_Click()
    frmBuscar.MostrarSql "Select CODIGO as [ Codigo                       ], DESCRIPCION  as [ Descripcion                                                        ]From PRODUCTO Where ACTIVO = 1 Order By CODIGO"
    txtCodigoProd.Text = frmBuscar.resultado
    txtDescripcionProd.Text = frmBuscar.resultado(2)
End Sub

Private Sub CmdAyudaProdHasta_Click()
    frmBuscar.MostrarSql "Select CODIGO as [ Codigo                           ], DESCRIPCION  as [ Descripcion                                                        ]From PRODUCTO Where ACTIVO = 1 Order By CODIGO"
    TxtCodigoProdHasta.Text = frmBuscar.resultado
    TxtDescripcionProdHasta.Text = frmBuscar.resultado(2)
End Sub

Private Sub cmdBuscar_Click()

Dim Consulta As String
Dim CodProd As String, desProd As String
Dim rsmov As New ADODB.Recordset
Dim SaldoProd As Double
Dim cant As Double
Dim rsProductos As New ADODB.Recordset
'
'Dim ttt
'ttt = Timer

    
    If txtCodigoProd.Text <> "" Then

        ttMovStk = TablaTempCrear("([ID] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ,[CODIGO] [varchar](50) NULL,[FECHA] [datetime] NULL , [TIPOCOMPROBANTE] [nvarchar] (50) NULL ,[NROCOMPROBANTE] [nvarchar] (50) NULL, [CANTIDAD] [nvarchar] (50) NULL, [SALDO] [nvarchar] (50) NULL, [Concepto]  [nvarchar] (50) NULL, [Descripcion] [varchar] (60) ) ON [PRIMARY]")
'        rsTempMov.Open "select * from " & ttMovStk, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        
        rsProductos.Open "Select codigo, descripcion from Producto where activo=1 and codigo between '" & txtCodigoProd & "'  and '" & TxtCodigoProdHasta & "'", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        
        
        While Not rsProductos.EOF
                
            CodProd = Trim(rsProductos!codigo)
            desProd = sSinNull(Trim(rsProductos!DESCRIPCION))
                    
            SaldoProd = CalcularSaldoAnterior(rsProductos!codigo, dtfechad.Value)
            AgregarEnTabla CodProd, dtfechad.Value, const_SALDO_INICIAL, 0, SaldoProd, 0, "", desProd
            
            
            'TABLA REMITO COMPRA
            Consulta = "select R.codigo, R.FECHA, D.PRODUCTO, D.CANTIDAD " & _
                        "from REMITOCOMPRADETALLE as D INNER JOIN REMITOCOMPRA AS R ON R.CODIGO = D.CODIGOREMITO " & _
                        "where R.ACTIVO = 1 and " & _
                        "D.PRODUCTO = '" & CodProd & "' and " & _
                        "R.FECHA " & ssBetween(dtfechad.Value, dtfechah.Value)
            rsmov.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            If Not rsmov.EOF Then rsmov.MoveFirst
            While Not rsmov.EOF
                AgregarEnTabla CodProd, rsmov!fecha, const_REMITO_COMPRA, rsmov!codigo, CDbl(rsmov!cantidad), 0, "", desProd
                rsmov.MoveNext
            Wend
            rsmov.Close
            Set rsmov = Nothing
            
            'TABLA REMITO DIFERENCIA STOCK
            Consulta = "select R.COMPROBANTE, R.MovimientoInterno , R.NROCOMPROBANTE, R.FECHA, D.PRODUCTO, D.CANTIDAD, r.Concepto " _
                & " From ITEMREMITODIFERENCIASTOCK as D INNER JOIN REMITODIFERENCIASTOCK AS R " _
                & " ON R.movimientointerno = D.NUMERO " _
                & " Where D.PRODUCTO = '" & CodProd & "' and " _
                & " R.FECHA " & ssBetween(dtfechad.Value, dtfechah.Value)
            rsmov.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockReadOnly
    
            While Not rsmov.EOF
                AgregarEnTabla CodProd, rsmov!fecha, const_MOVIMIENTO_MANUAL, CLng(rsmov!MovimientoInterno), -CDbl(rsmov!cantidad), 0, sSinNull(obtenerDeSQL("select descripcion from conceptos where codigo = '" & rsmov!concepto & "'")), desProd
                rsmov.MoveNext
            Wend
            rsmov.Close
            Set rsmov = Nothing
            
            'remitos no facturados
            Dim ss As String, tempo
            Consulta = "select R.numero, R.FECHA, D.PRODUCTO, d.cantidad, d.codigo  " & _
                        " from REMITOVENTADETALLE as D INNER JOIN REMITOVENTA AS R ON R.NUMERO = D.NUMERO " & _
                        " where D.PRODUCTO = '" & CodProd & "' And R.FECHA " & ssBetween(dtfechad.Value, dtfechah.Value) & " and r.anulado = 0"
            With rsmov
                .Open Consulta, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                While Not .EOF
                    ss = "select fecha, f.tipodoc, f.nrofactura,  cantidad, producto from facturaventadetalle as d inner join facturaventa as f  on f.codigo = d.CodigoFactura " _
                        & " where d.item_p_r = " & !codigo & " and f.activo = 1 and producto='" & CodProd & "' "
                    tempo = obtenerDeSQL(ss)
                    If IsEmpty(tempo) Then
                        'remitos sin facura
                        AgregarEnTabla CodProd, rsmov!fecha, const_REMITO_VENTA, CLng(rsmov!numero), -!cantidad, 0, "", desProd
                    Else
                        'factura sobre remito
                        AgregarEnTabla CodProd, CDate(tempo(0)), CStr(tempo(1)), tempo(2), -!cantidad, 0, "", desProd
                    End If
                    .MoveNext
                Wend
                .Close
                'factura sin remito
                Consulta = "select R.TIPODOC, R.NROFACTURA, R.FECHA, D.PRODUCTO, D.CANTIDAD " _
                    & "from FACTURAVENTADETALLE as D INNER JOIN FACTURAVENTA AS R ON R.CODIGO = D.CODIGOFACTURA " _
                    & "where R.ACTUALIZASTOCK = 1  and D.PRODUCTO = '" & CodProd & "' and " _
                    & "R.FECHA " & ssBetween(dtfechad.Value, dtfechah.Value) & " and R.Activo = 1"
                .Open Consulta, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                While Not .EOF
                    AgregarEnTabla CodProd, !fecha, !TIPODOC, CLng(!nrofactura), -s2n(!cantidad), 0, "", desProd
                    .MoveNext
                Wend
                .Close
            End With
            '******************************************************
            
            CalcularSaldo (CodProd)
            rsProductos.MoveNext
        Wend
        LlenarGrilla GrillaDetalle, "Select Codigo, Fecha, TIPOCOMPROBANTE AS 'Comprobante', NROCOMPROBANTE AS 'Numero', CANTIDAD as 'Cant', SALDO as 'Saldo', Concepto, Descripcion  " & _
                                    "From " & ttMovStk & " Order By CODIGO, FECHA", True
        With GrillaDetalle
            .ColWidth(0) = 2000
            .ColWidth(1) = 1100
            .ColWidth(2) = 700 '2000
            .ColWidth(3) = 700 '2000
            .ColWidth(4) = 500 '1000
            .ColWidth(5) = 500 '1000
            .ColWidth(6) = 3000
            .ColWidth(7) = 4000
            .ColAlignment(0) = flexAlignLeftCenter
        End With
        
'        rsTempMov.Close
    Else
        MsgBox "Debe ingresar un codigo de producto", vbOKOnly, "Atencion"
    End If
    
'    MsgBox Timer - ttt
End Sub

Private Sub cmdImprimir_Click()
Dim Consulta As String
    
    Consulta = "Select * From  " & ttMovStk & "  Order By CODIGO, FECHA"
    With rptMovimientoStock
        .Data.Connection = DataEnvironment1.Sistema
        .Data.Source = Consulta
        '.lblCodigo.caption = Trim(txtCodigoProd.Text)
        '.lblDescripcion.caption = ObtenerDescripcionS("PRODUCTO", Trim(txtCodigoProd.Text))
        .lblFechaReporte.caption = Date
        .EncCodigo.DataField = "CODIGO"
        '.EncDescripcion.DataField = "DESCRIPCION"
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
    dtfechad.Value = Date
    dtfechah.Value = Date
    CentrarMe frmMovimientoStock
    ucXls1.ini GrillaDetalle, "bsMovStock.xls"
End Sub

'Private Sub Form_Unload(Cancel As Integer)
'    Set rsTempMov = Nothing
'End Sub

Private Sub GrillaDetalle_Click()
Dim TIPODOC As String
Dim NroDoc As String
Dim CodInt As Long

    With GrillaDetalle
        If .TextMatrix(.Row, 0) <> "" And .Row <> 0 Then
            TIPODOC = Trim(.TextMatrix(.Row, 2))
            If .TextMatrix(.Row, 3) <> "" Then NroDoc = .TextMatrix(.Row, 3)
            
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
'
Private Sub TxtCodigoProdHasta_Change()

End Sub

Private Sub TxtCodigoProdHasta_LostFocus()
If Trim(TxtCodigoProdHasta) <> "" Then
        TxtDescripcionProdHasta = ObtenerDescripcionS("PRODUCTO", Trim(TxtCodigoProdHasta.Text))
        If Trim(TxtDescripcionProdHasta) = "" Then
            MsgBox "El codigo del producto es inexistente", 48, "Atencion"
            TxtCodigoProdHasta.SetFocus
        End If
    End If
End Sub



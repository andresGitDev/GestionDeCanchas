VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmVerOPdetalle 
   Caption         =   "Detalle  de Ordenes de Pago"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10590
   Icon            =   "frmVerOPdetalle.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   10590
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCabecera 
      Height          =   1155
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   10440
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   345
         Left            =   4035
         TabIndex        =   3
         Top             =   225
         Width           =   1260
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   345
         Left            =   9045
         TabIndex        =   2
         Top             =   240
         Width           =   1260
      End
      Begin VB.CommandButton cmdVer 
         Caption         =   "&Mostrar"
         Height          =   345
         Left            =   2940
         TabIndex        =   1
         Top             =   225
         Width           =   1080
      End
      Begin Gestion.ucXls ucXls1 
         Height          =   810
         Left            =   5400
         TabIndex        =   4
         Top             =   210
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   1429
      End
      Begin Gestion.ucFecha uFechaH 
         Height          =   345
         Left            =   1605
         TabIndex        =   5
         Top             =   240
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   609
         FechaInit       =   3
      End
      Begin Gestion.ucFecha uFechaD 
         Height          =   345
         Left            =   615
         TabIndex        =   6
         Top             =   240
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   609
         FechaInit       =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Entre"
         Height          =   285
         Left            =   45
         TabIndex        =   7
         Top             =   300
         Width           =   1020
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Height          =   4755
      Left            =   60
      TabIndex        =   8
      Top             =   1260
      Width           =   10455
      _cx             =   18441
      _cy             =   8387
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
      Cols            =   30
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmVerOPdetalle.frx":08CA
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
Attribute VB_Name = "frmVerOPdetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mSelGrilla As String
Private mColCorte As Long
Private mColSum As Variant
Private mAnchos As Variant
Private mTitulo As String

Private ttOP As String
'Private Const ttpagos = " ( [id] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ,  [idDoc] [numeric](18, 0) NOT NULL, [Fecha] [datetime] NOT NULL , [TipoDoc] [char] (3)  NOT NULL ,  [NroDoc] [numeric](18, 0) NOT NULL , [NroOP] [numeric](18, 0) NOT NULL , [CodProv] [numeric](18, 0) NOT NULL , [cuit] [char] (13),  [RazonSocial] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[Total] [Float], [Neto] [Float], [RetGan] [Float],[RetIB] [Float],[CertifGan] [numeric],[CertifIIBB] [numeric], activo [bit] ) ON [PRIMARY]"
Private Const ttpagos = " ( [id] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ,  [idDoc] [numeric](18, 0) NOT NULL, [Fecha] [datetime] NOT NULL , [TipoDoc] [char] (3)  NOT NULL ,  [NroDoc] [numeric](18, 0) NOT NULL , [NroOP] [numeric](18, 0) NOT NULL , [CodProv] [numeric](18, 0)   , [cuit] [char] (13),  [RazonSocial] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[Total] [Float], [Neto] [Float], [RetGan] [Float],[RetIB] [Float],[CertifGan] [numeric],[CertifIIBB] [numeric], activo [bit], detalle [bit], Facturas [float], monto [float] ) ON [PRIMARY]"

'

Private Sub cmdVer_Click()
    Dim s1 As String, s2 As String
    
    'titulo xls
    ucXls1.aTitulo = Format(Date, "dd/mm/yy") & "Ordenes de pago entre " & uFechaD.dtFecha & " y " & uFechaH.dtFecha
    
    'tabla temp
    DataEnvironment1.Sistema.Execute "delete from " & ttOP
    
    'campos a insertar
    s1 = "insert into " & ttOP & _
                " ( iddoc, fecha, tipodoc, nrodoc, nroop, codprov,razonsocial, total, neto, RetGan, RetIB, CertifGan, CertifIIBB, cuit, activo ) "

    'op
    s2 = " SELECT op.idDoc, op.FECHA, 'REC' AS tdoc, op.NRO, d.NumeroDePago, op.CODPR, p.descripcion, total, neto,  retganPago, ibPago, NroCertifGan, nroCertifIIBB, cuit, op.activo " & _
                    " FROM PROV AS p RIGHT JOIN (REC_COMP AS op LEFT JOIN RegistroDocumentos AS d ON op.idDoc = d.idDoc) ON p.codigo = op.CODPR " & _
                    " where op.fecha between " & uFechaD.ssFecha & " and " & uFechaH.ssFecha
    DataEnvironment1.Sistema.Execute s1 & s2
    
'    'compras fac cdo y rac
'    s2 = " SELECT c.idDoc, c.FECHA, c.tipodoc, c.NROdoc, d.NumeroDePago, c.CODPR, c.razonsocialProv, total, neto, retganPago, ibPago , NroCertifGan, nroCertifIIBB, cuitprov, c.ACTIVO " & _
'                    " FROM COMPRAS  AS c LEFT JOIN RegistroDocumentos AS d ON c.idDoc = d.idDoc  " & _
'                    " WHERE c.fecha between " & uFechaD.ssFecha & " and " & uFechaH.ssFecha & _
'                    " and ((c.tipodoc = 'FAC' and c.contado = 1) or  (c.tipodoc = 'RAC')) "
'    DataEnvironment1.Sistema.Execute s1 & s2
'
'    'transcom rac
'    s2 = " SELECT c.idDoc, c.FECHA, c.tipodoc, c.NROdoc, d.NumeroDePago, c.CODPR, c.razonsocialProv, total, neto, retganPago, ibPago, NroCertifGan, nroCertifIIBB, cuitprov, c.ACTIVO  " & _
'                    " FROM transcom  AS c LEFT JOIN RegistroDocumentos AS d ON c.idDoc = d.idDoc  " & _
'                    " WHERE c.fecha between " & uFechaD.ssFecha & " and " & uFechaH.ssFecha & " and (  (c.tipodoc = 'RAC'))"
'    DataEnvironment1.Sistema.Execute s1 & s2
'
'    'caja egreso
'    s2 = " SELECT c.movimiento, c.FECHA, d.tipodoc, c.movimiento, d.NumeroDePago,  c.CODPROV, prov.descripcion, importe,0 , 0, 0, 0, 0, prov.cuit, d.activo  " & _
'                    " FROM movicaja AS c LEFT JOIN RegistroDocumentos AS d ON c.idDoc = d.idDoc  left join prov on c.codprov = prov.codigo " & _
'                    " WHERE c.fecha between " & uFechaD.ssFecha & " and " & uFechaH.ssFecha & " and d.tipodoc = 'EFE' and ing_egr = 'E' "
'    DataEnvironment1.Sistema.Execute s1 & s2
'

    'Borrados
    s2 = " select iddoc, fecha_alta, tipodoc, nrodoc, NumeroDePago, codproveedor, ' ----ANULADO---- ', 0,0,0,0,0,0, '', 0  " & _
                    " from registrodocumentos " & _
                    " where Fecha_alta between " & uFechaD.ssFecha & " and " & uFechaH.ssFecha & " and Numerodepago > 0 and activo = 0" & _
                    " and tipodoc = 'REC' " 'TIPODOC AGREGADO PARA ESTE DETALLE
    DataEnvironment1.Sistema.Execute s1 & s2
    
    'Detalle OP
    s1 = " insert into " & ttOP & _
                " ( iddoc, fecha, tipodoc, nrodoc, nroop, activo, detalle, facturas, monto ) "
    
    
    s2 = "SELECT     op.idDoc, op.FECHA, 'REC' AS tdoc, op.NRO, d.NumeroDePago, op.ACTIVO, 1, r.FACT, r.IMPOR " & _
        " FROM        REC_COMP op LEFT OUTER JOIN " & _
        " RegistroDocumentos d ON op.idDoc = d.idDoc INNER JOIN " & _
        " RELFNR_C r ON op.idDoc = r.iddoc " & _
        " WHERE     (op.idDoc > 0) AND (op.ACTIVO = 1) AND (r.TFAC = 'FAC' or r.tfac = 'N/D' or r.tfac = 'APD' )  and op.fecha between " & uFechaD.ssFecha & " and " & uFechaH.ssFecha
    DataEnvironment1.Sistema.Execute s1 & s2
    
    s2 = "SELECT     op.idDoc, op.FECHA, 'REC' AS tdoc, op.NRO, d.NumeroDePago, op.ACTIVO, 1, r.FACT, -(r.IMPOR)" & _
        " FROM        REC_COMP op LEFT OUTER JOIN " & _
        " RegistroDocumentos d ON op.idDoc = d.idDoc INNER JOIN " & _
        " RELFNR_C r ON op.idDoc = r.iddoc " & _
        " WHERE     (op.idDoc > 0) AND (op.ACTIVO = 1) AND (r.TFAC = 'RAC' or r.tfac = 'N/C' or r.tfac = 'APC' )  and op.fecha between " & uFechaD.ssFecha & " and " & uFechaH.ssFecha
    DataEnvironment1.Sistema.Execute s1 & s2
    
    'bug, me aseguro que los borrados no sumen en la grilla
    DataEnvironment1.Sistema.Execute "update " & ttOP & " set total = 0 , neto = 0 , retgan = 0, retib = 0 , razonsocial = ' ***** ANULADO *****'  where activo = 0"
    
'    optMostrar(0).Value = True
    optMostrar_Click (0)
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub
Private Sub Form_Load()
    Form_Resize
    ucXls1.ini grilla, "c:\ListadoPagos-Facturas ", "Rel Ordenes de pago/Facturas "
    HacerTabla
End Sub
Private Sub Form_Resize()
    Anclar grilla, Me, anclarLadosTodos
End Sub

Private Sub HacerTabla()
'    Dim t1 As String, t2 As String, t3 As String
'    t1 =
    ttOP = TablaTempCrear(ttpagos)
'    t2 = "A LTER TABLE " & ttOP & "  WITH NOCHECK ADD CONSTRAINT [PK_kk_LisOP] PRIMARY KEY  CLUSTERED"
'    (
'        [ID]
'    )  ON [PRIMARY]
'GO
'ALTER TABLE [dbo].[_kk_LisOP] ADD
'    CONSTRAINT [DF_kk_LisOP_RazonSocial] DEFAULT ('') FOR [RazonSocial],
'    CONSTRAINT [DF_kk_LisOP_Importe] DEFAULT (0) FOR [Importe]
'GO
End Sub

Private Sub cmdImprimir_Click()
    If grilla.rows < 2 Then Exit Sub
    
    grilla.GridLines = flexGridNone
    grilla.GridLinesFixed = flexGridNone
    
    FrmImpresiones.VSPrinter.Orientation = orLandscape
    FrmImpresiones.VSPrinter.PaperSize = pprA4
    FrmImpresiones.VSPrinter.Preview = True
    FrmImpresiones.VSPrinter.Font.Name = grilla.Font.Name
    FrmImpresiones.VSPrinter.FontSize = 12
    FrmImpresiones.VSPrinter.Header = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
    FrmImpresiones.VSPrinter.FontSize = 8
'    FrmImpresiones.VSPrinter.Footer = "||Pagina %d de " & FrmImpresiones.VSPrinter.PageCount ' & " de " & "%d"
    
    FrmImpresiones.VSPrinter.StartDoc
    'FrmImpresiones.VSPrinter.Paragraph = "Listado Mayor al " & Format$(Date, "dd / mm / yyyy")
    FrmImpresiones.VSPrinter.Paragraph = mTitulo
    FrmImpresiones.VSPrinter.Paragraph = "Entre fechas : " & uFechaD.dtFecha & " - " & uFechaH.dtFecha  '& "     Rango de Cuentas : " & CmbCtaD & "  -  " & CmbCtaH
    FrmImpresiones.VSPrinter.Paragraph = " "
    
       FrmImpresiones.VSPrinter.TextAlign = taLeftBottom
       FrmImpresiones.VSPrinter.RenderControl = grilla.hWnd

    FrmImpresiones.VSPrinter.Footer = "||Pagina %d de " & FrmImpresiones.VSPrinter.PageCount ' & " de " & "%d"
    FrmImpresiones.VSPrinter.Zoom = 100
    FrmImpresiones.VSPrinter.EndDoc
    
    FrmImpresiones.Show
    grilla.GridLines = flexGridFlat
End Sub


Private Sub optMostrar_Click(Index As Integer)
        
        mSelGrilla = "select Fecha, month(fecha) as [_H_mes] , NroOP, CodProv as Prov, CUIT, RazonSocial, CertifGan, CertifIIBB, Total, Neto, RetGan, RetIB,  Facturas, Monto" & _
                " from " & ttOP & " where nroop >0 order by fecha, NroOP, tipodoc, nrodoc, id"
        mColCorte = 2
        mColSum = Array(8, 9, 10, 11, 13)
        mAnchos = Array(1200, 0, 1000, 500, 1400, 2000, 850, 850, 1200, 1200, 1200, 1200, 975, 1400)
        mTitulo = "Listado de pagos"
    
    mostrar
End Sub

Private Sub mostrar()
    Dim i As Long
    
    With grilla
        LlenarGrilla grilla, mSelGrilla, False, mColCorte '  , mColSum     'corte por mes
        
        For i = 0 To .cols - 1: .ColWidth(i) = 1400: Next
        For i = 0 To UBound(mAnchos) - 1: .ColWidth(i) = mAnchos(i): Next
        
        .SubtotalPosition = flexSTBelow
        .subtotal flexSTClear
        sumarizo
    End With
End Sub

Private Sub sumarizo()
    Dim i As Long
    With grilla
        For i = 0 To UBound(mColSum):        .subtotal flexSTSum, -1, mColSum(i), , , , True, , , True: Next
        .TextMatrix(.rows - 1, 0) = " Totales"
    End With
End Sub



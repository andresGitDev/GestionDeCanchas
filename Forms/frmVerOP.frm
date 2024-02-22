VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmVerOP 
   Caption         =   "Listado de Pagos"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10560
   Icon            =   "frmVerOP.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   10560
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCabecera 
      Height          =   1095
      Left            =   45
      TabIndex        =   5
      Top             =   -60
      Width           =   9735
      Begin VB.CommandButton cmdVer 
         Caption         =   "&Mostrar"
         Height          =   345
         Left            =   3000
         TabIndex        =   9
         Top             =   240
         Width           =   1260
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   345
         Left            =   8340
         TabIndex        =   7
         Top             =   240
         Width           =   1260
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   345
         Left            =   4290
         TabIndex        =   6
         Top             =   240
         Width           =   1260
      End
      Begin Gestion.ucXls ucXls1 
         Height          =   795
         Left            =   5640
         TabIndex        =   8
         Top             =   240
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1402
      End
      Begin Gestion.ucFecha uFechaH 
         Height          =   345
         Left            =   1620
         TabIndex        =   10
         Top             =   240
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   609
         FechaInit       =   3
      End
      Begin Gestion.ucFecha uFechaD 
         Height          =   345
         Left            =   645
         TabIndex        =   11
         Top             =   240
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   609
         FechaInit       =   0
      End
      Begin VB.Label Label1 
         Caption         =   "Entre"
         Height          =   285
         Left            =   45
         TabIndex        =   12
         Top             =   300
         Width           =   1020
      End
   End
   Begin VB.OptionButton optMostrar 
      Caption         =   "Ret IB"
      Height          =   330
      Index           =   3
      Left            =   3885
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1455
      Width           =   1245
   End
   Begin VB.OptionButton optMostrar 
      Caption         =   "Ret Gan"
      Height          =   330
      Index           =   2
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1455
      Width           =   1245
   End
   Begin VB.OptionButton optMostrar 
      Caption         =   "Pagos"
      Height          =   330
      Index           =   1
      Left            =   1380
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1455
      Width           =   1245
   End
   Begin VB.OptionButton optMostrar 
      Caption         =   "Todo"
      Height          =   330
      Index           =   0
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1455
      Value           =   -1  'True
      Width           =   1245
   End
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Height          =   4755
      Left            =   45
      TabIndex        =   4
      Top             =   1935
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
      FormatString    =   $"frmVerOP.frx":08CA
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
Attribute VB_Name = "frmVerOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'falta evitar los activo = 0, sino le va a duplicar los anulados
    
    'Para compras, agregue los campos NumIIBB, regimenes IB GAN  de la misma tabla
    'Para transcom, agregue NumIIBB de transcom , pero regimenes falta agregarlos (en la tabla tambien estan)
    '               si agrego esto, puedo borrar el update al final ? no.

Private mSelGrilla As String
Private mColCorte As Long
Private mColSum As Variant
Private mAnchos As Variant
Private mTitulo As String

Private ttOP As String
Private Const ttpagos = " ( [id] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ,  [idDoc] [numeric](18, 0) NOT NULL, [Fecha] [datetime] NOT NULL , [TipoDoc] [char] (3)  NOT NULL ,  [NroDoc] [numeric](18, 0) NOT NULL , [NroOP] [numeric](18, 0) NOT NULL , [CodProv] [numeric](18, 0) NOT NULL , [cuit] [char] (13),  [RazonSocial] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,[Total] [Float], [Neto] [Float], [Exento] [Float], [RetGan] [Float],[RetIB] [Float],[CertifGan] [numeric],[CertifIIBB] [numeric], NumIBprov char (20), RegIB char(20), regGan char(20),  activo [bit] ) ON [PRIMARY]"
'
'

Private Sub cmdVer_Click()
    Dim s1 As String, s2 As String
    
    'titulo xls
    ucXls1.aTitulo = Format(Date, "dd/mm/yy") & "Ordenes de pago entre " & uFechaD.dtFecha & " y " & uFechaH.dtFecha
    
    'tabla temp
    DataEnvironment1.Sistema.Execute "delete from " & ttOP
    
    'campos a insertar
    s1 = "insert into " & ttOP & _
                " ( iddoc, fecha, tipodoc, nrodoc, nroop, codprov, razonsocial, total, neto, exento, RetGan, RetIB, CertifGan, CertifIIBB, cuit, numIBprov, regib, reggan, activo ) "
    
    'op
    s2 = " SELECT op.idDoc, op.FECHA, 'REC' AS tdoc, op.NRO, d.NumeroDePago, op.CODPR, p.descripcion, total, neto, 0, retganPago, ibPago, NroCertifGan, nroCertifIIBB, cuit, p.numiibb,'','', op.activo " & _
                    " FROM PROV AS p RIGHT JOIN (REC_COMP AS op LEFT JOIN RegistroDocumentos AS d ON op.idDoc = d.idDoc) ON p.codigo = op.CODPR " & _
                    " where op.fecha between " & uFechaD.ssFecha & " and " & uFechaH.ssFecha
    DataEnvironment1.Sistema.Execute s1 & s2, "a", 1 'ASI ESTABA ANTES PERO SALTABA ERROR AL PASAR S1 Y S2 ANEXADOS
    'DataEnvironment1.Sistema.Execute s2, "a", 1  '
    
    'compras fac cdo y rac       ' le agregue Num IIBB de COMPRAS
    s2 = " SELECT   c.idDoc, c.FECHA, c.TIPODOC, c.NRODOC, d.NumeroDePago, c.CODPR, c.RAZONSOCIALPROV, c.TOTAL, c.NETO, c.exento, c.RetGanPago, c.IBpago, d.NroCertifGan, d.NroCertifIIBB, c.CUITPROV, c.NroIIBB, left(ProvTipoRetIB.Descripcion,20) AS ribDes, left(ProvTipoRetGan.Descripcion,020) AS rgaDes, c.ACTIVO " & _
            " FROM         ProvTipoRetIB RIGHT OUTER JOIN " & _
            " COMPRAS c ON ProvTipoRetIB.Codigo = c.TipoRetIIBB LEFT OUTER JOIN " & _
            " ProvTipoRetGan ON c.TipoRetGan = ProvTipoRetGan.Codigo LEFT OUTER JOIN " & _
            " Prov RIGHT OUTER JOIN " & _
            " RegistroDocumentos d ON Prov.codigo = d.CodProveedor ON c.idDoc = d.idDoc " & _
            " WHERE c.fecha between " & uFechaD.ssFecha & " and " & uFechaH.ssFecha & _
            " and ((c.tipodoc = 'FAC' and c.contado = 1) or  (c.tipodoc = 'RAC')) "
    DataEnvironment1.Sistema.Execute s1 & s2
    
    'transcom rac
    s2 = " SELECT c.idDoc, c.FECHA, c.tipodoc, c.NROdoc, d.NumeroDePago, c.CODPR, c.razonsocialProv, total, neto, exento, retganPago, ibPago, NroCertifGan, nroCertifIIBB, cuitprov, c.nroiibb, '', '', c.ACTIVO  " & _
                    " FROM transcom  AS c LEFT JOIN RegistroDocumentos AS d ON c.idDoc = d.idDoc  " & _
                    " left join prov on prov.codigo = d.CodProveedor " & _
                    " WHERE c.fecha between " & uFechaD.ssFecha & " and " & uFechaH.ssFecha & " and (  (c.tipodoc = 'RAC'))"
    DataEnvironment1.Sistema.Execute s1 & s2
    
    'caja egreso
    s2 = " SELECT c.movimiento, c.FECHA, d.tipodoc, c.movimiento, d.NumeroDePago,  c.CODPROV, prov.descripcion, importe, 0, 0, 0, 0, 0, 0, prov.cuit, prov.numiibb, '', '', d.activo  " & _
                    " FROM movicaja AS c LEFT JOIN RegistroDocumentos AS d ON c.idDoc = d.idDoc  left join prov on c.codprov = prov.codigo " & _
                    " WHERE c.fecha between " & uFechaD.ssFecha & " and " & uFechaH.ssFecha & " and d.tipodoc = 'EFE' and ing_egr = 'E' "
    DataEnvironment1.Sistema.Execute s1 & s2
   
    'Borrados
    's2 = " Select iddoc, fecha_alta, tipodoc, nrodoc, NumeroDePago, codproveedor, ' ----ANULADO---- ',0,0,0,0,0,NroCertifGan,NroCertifIIBB, '', '', '', '', 0  " & _
                    " from registrodocumentos " & _
                    " where Fecha_alta between " & uFechaD.ssFecha & " and " & uFechaH.ssFecha & " and Numerodepago > 0 and activo = 0"
    'DataEnvironment1.Sistema.Execute s1 & s2
    
    
    'bug, me aseguro que los borrados no sumen en la grilla
    DataEnvironment1.Sistema.Execute "update " & ttOP & " set total = 0 , neto = 0 , retgan = 0, retib = 0 , razonsocial = ' ***** ANULADO *****' ,cuit='' where activo = 0"
    
    
'    ' y perdon pero lo prefiero al inner join
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT t.id, t.RegIB, t.RegGan, t.certifiibb, t.certifgan, r.Descripcion AS RegimenGan, i.Descripcion AS RegimenIIBB, tiporetgan, tiporetiibb " & _
            " FROM " & ttOP & " t LEFT OUTER JOIN " & _
            " Prov p ON p.codigo = t.CodProv LEFT OUTER JOIN " & _
            " ProvTipoRetGan r ON r.Codigo = p.TipoRetGan LEFT OUTER JOIN " & _
            " ProvTipoRetIB i ON i.Codigo = p.TipoRetIIBB " _
                , DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        
        While Not .EOF
            If (!certifiibb > 0 And !TipoRetIIBB > 0) Or (!certifgan > 0 And !TipoRetGan > 0) Then
                '!RegIB = ssstr(Left(!regimeniibb, 20)
                '!RegGan = Left(!regimengan, 20)
                '.Update
                DataEnvironment1.Sistema.Execute "update " & ttOP & " set " & _
                    " regib  = '" & ssStr(Left(!regimeniibb, 20)) & "' , " & _
                    " reggan  = '" & ssStr(Left(!regimengan, 20)) & "'  " & _
                    " where id = " & !ID
            End If
            .MoveNext
        Wend
        
        
    End With
    
    DataEnvironment1.Sistema.Execute "update " & ttOP & " set tipodoc='PAC'  where tipodoc='RAC'"
    DataEnvironment1.Sistema.Execute "update " & ttOP & " set tipodoc='O/P'  where tipodoc='REC'"
    
    Set rs = Nothing
    
    optMostrar(0).Value = True
    optMostrar_Click (0)
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub
Private Sub Form_Load()
    Form_Resize
    ucXls1.ini grilla, "c:\ListadoPagos ", "Ordenes de pago"
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
'        x As VSPrinter8LibCtl.OrientationSettings
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
    Select Case Index
    Case 0 ' todo
        mSelGrilla = "select Fecha, month(fecha) as [_H_mes] ,TipoDoc as Tipo, NroDoc as Nro, NroOP, CodProv as Prov, CUIT, RazonSocial, CertifGan, CertifIIBB, Total, Neto, Exento, RetGan, RetIB, NumIBprov " & _
                " from " & ttOP & " order by fecha, NroOP, tipodoc, nrodoc"
        mColCorte = 1
        mColSum = Array(10, 11, 12, 13, 14)
        mAnchos = Array(1000, 0, 500, 800, 800, 500, 1200, 2000, 800, 800, 1200, 1200, 1200, 1200, 1200, 1200, 1000)
        mTitulo = "Listado de pagos"
    Case 1 'pagos
        mSelGrilla = "select Fecha, month(fecha) as [_H_mes] ,TipoDoc as Tipo, NroDoc as Nro, NroOP, CodProv as Prov, CUIT, RazonSocial, total, neto, exento " & _
                " from " & ttOP & " order by fecha, NroOP, tipodoc, nrodoc"
        mColCorte = 1
        mColSum = Array(8, 9, 10)
        mAnchos = Array(1200, 0, 500, 1000, 1000, 500, 1400, 2000, 2000)
        mTitulo = "Listado de pagos"
    Case 2 ' r gan      ' Quizas pidan que vuelva a poner NroDoc...
        mSelGrilla = "select Fecha, month(fecha) as [_H_mes] ,TipoDoc as Tipo, NroDoc as [_H_Nro], NroOP, CodProv as Prov, CUIT, RazonSocial, certifgan as certificado, total, neto, exento, retgan as Retenciones, RegGan as Regimen " & _
                " from " & ttOP & _
                " where certifgan > 0 " & _
                " order by fecha, NroOP, tipodoc, nrodoc "
        mColCorte = 1
        mColSum = Array(9, 10, 11, 12)
        mAnchos = Array(1200, 0, 500, 1000, 1000, 500, 1400, 2000, 1000, 1000)
        mTitulo = "Listado de retenciones de Ganancias"
    Case 3 ' r ib       ' Quizas pidan que vuelva a poner NroDoc...
        mSelGrilla = "select Fecha, month(fecha) as [_H_mes] ,TipoDoc as Tipo, NroDoc as [_H_Nro], NroOP, CodProv as Prov, CUIT, RazonSocial,certifiibb as certificado,  Total, Neto, exento, retib  as Retenciones, NumIbProv as [Nro IIBB prov], RegIB as regimen" & _
                " from " & ttOP & _
                " where certifiibb > 0 " & _
                " order by fecha, NroOP, tipodoc, nrodoc"
        mTitulo = "Listado de retenciones de Ingresos Brutos"
        mColCorte = 1
        mColSum = Array(9, 10, 11, 12)
        mAnchos = Array(1200, 0, 500, 1000, 1000, 500, 1400, 2000, 1000, 1000, 1000)
    End Select
    
    mostrar
End Sub

Private Sub mostrar()
    ' tabla a grilla
'    On Error Resume Next
    Dim i As Long
    
    With grilla
        .AutoResize = True ' ??? no anda
        LlenarGrilla grilla, mSelGrilla, False   ',mColCorte  , mColSum     'corte por mes
        
        For i = 0 To .cols - 1: .ColWidth(i) = 1400: Next
        For i = 0 To UBound(mAnchos) - 1: .ColWidth(i) = mAnchos(i): Next
        
        .SubtotalPosition = flexSTBelow
        .subtotal flexSTClear
        sumarizo
    End With
End Sub

Private Sub sumarizo()
    On Error GoTo ufa
    Dim i As Long
    With grilla
        For i = 0 To UBound(mColSum)
            .subtotal flexSTSum, -1, mColSum(i), , , , True, , , True
        Next
        .TextMatrix(.rows - 1, 0) = " Totales"
    End With
ufa:
End Sub

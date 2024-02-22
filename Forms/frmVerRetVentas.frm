VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmVerRetVentas 
   Caption         =   "Retenciones ventas"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11490
   Icon            =   "frmVerRetVentas.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   11490
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraBoton 
      Height          =   990
      Left            =   90
      TabIndex        =   6
      Top             =   6585
      Width           =   11295
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir"
         Height          =   345
         Left            =   8370
         TabIndex        =   8
         Top             =   210
         Width           =   1260
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   345
         Left            =   9885
         TabIndex        =   7
         Top             =   195
         Width           =   1260
      End
      Begin Gestion.ucXls ucXls1 
         Height          =   750
         Left            =   7290
         TabIndex        =   9
         Top             =   195
         Width           =   960
         _extentx        =   1693
         _extenty        =   1323
      End
   End
   Begin Gestion.ucCoDe uRe 
      Height          =   330
      Left            =   3090
      TabIndex        =   2
      Top             =   60
      Width           =   5100
      _extentx        =   10583
      _extenty        =   582
      codigowidth     =   1000
   End
   Begin VB.CommandButton cmdVer 
      Caption         =   "Mostrar"
      Height          =   345
      Left            =   8310
      TabIndex        =   3
      Top             =   45
      Width           =   1875
   End
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Height          =   6015
      Left            =   30
      TabIndex        =   4
      Top             =   510
      Width           =   11355
      _cx             =   20029
      _cy             =   10610
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
      FormatString    =   $"frmVerRetVentas.frx":08CA
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
   Begin Gestion.ucFecha uFechaH 
      Height          =   345
      Left            =   1815
      TabIndex        =   1
      Top             =   15
      Width           =   1080
      _extentx        =   1905
      _extenty        =   609
      fechainit       =   4
   End
   Begin Gestion.ucFecha uFechaD 
      Height          =   345
      Left            =   660
      TabIndex        =   0
      Top             =   30
      Width           =   1080
      _extentx        =   1905
      _extenty        =   609
      fechainit       =   5
   End
   Begin VB.Label Label1 
      Caption         =   "Entre"
      Height          =   285
      Left            =   90
      TabIndex        =   5
      Top             =   75
      Width           =   1020
   End
End
Attribute VB_Name = "frmVerRetVentas"
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

Private Sub cmdVer_Click()
    ucXls1.aTitulo = Format(Date, "dd/mm/yy") & "Retenciones entre " & uFechaD.dtFecha & " y " & uFechaH.dtFecha
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
    ucXls1.ini grilla, "c:\ListadoRetVenta ", "Retenciones venta"
    uRe.ini "select descripcion from cuentasparam where codigo = '###' and usocuenta = " & ID_UsoCuenta_RETVTA, " select codigo, descripcion from cuentasparam where usocuenta = " & ID_UsoCuenta_RETVTA, True
End Sub
Private Sub Form_Resize()
    Anclar fraBoton, Me, anclarIzquierda + anclarAbajo
    Anclar grilla, Me, anclarLadosTodos
End Sub


Private Sub optMostrar_Click(Index As Integer)
    Dim i As Long
    Dim ss As String, so As String, sw As String
    
    ss = " SELECT rr.idRetencion as [_H_idRet], rd.TipoDoc, rd.NroDoc, rr.Numero, rr.Fecha, rr.Importe, cl.descripcion AS Cliente, cl.CUIT, cp.descripcion AS TipoRetencion,cp.codigo,rr.IDDOC as 'Doc y Nro' " & _
                     " FROM CuentasParam AS cp RIGHT JOIN ((Recibos AS rc RIGHT JOIN (RegistroDocumentos AS rd RIGHT JOIN RecibosRetenciones AS rr ON rd.idDoc = rr.iddoc) ON rc.idDoc = rd.idDoc) LEFT JOIN Clientes AS cl ON rc.Cliente = cl.codigo) ON cp.id = rr.idCuentasParam " & _
                     " Where rd.Activo = 1 and rr.fecha between " & uFechaD.ssFecha & " and " & uFechaH.ssFecha
    If uRe.codigo > 0 Then sw = " and cp.codigo = '" & uRe.codigo & "' "
    so = " order by rr.fecha, cp.codigo, rr.numero "
    
    
    mSelGrilla = ss & sw & so
'    mColCorte = 1
    mColSum = Array(5)
    mAnchos = Array(540, 735, 735, 735, 1050, 1050, 2100, 660, 3000)
    mTitulo = "Listado de retenciones"

    mostrar
    i = 1
    While i < grilla.rows
        grilla.TextMatrix(i, 5) = s2n(grilla.TextMatrix(i, 5), 2, True)
        grilla.TextMatrix(i, 10) = obtenerDeSQL("select tipodoc + ' - ' + cast(numero as varchar) from recibos where iddoc =" & s2n(grilla.TextMatrix(i, 10)))
        i = i + 1
    Wend
End Sub

Private Sub cmdImprimir_Click()

    If grilla.rows < 2 Then Exit Sub
    
    grilla.GridLines = flexGridNone
    grilla.GridLinesFixed = flexGridNone
    
    FrmImpresiones.VSPrinter.Orientation = orPortrait ' orLandscape
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



Private Sub mostrar()
'    On Error Resume Next
    Dim i As Long

    With grilla
        .AutoResize = True ' ??? no anda
        LlenarGrilla grilla, mSelGrilla, False   ',mColCorte  , mColSum     'corte por mes

        For i = 0 To .cols - 1: .ColWidth(i) = 1400: Next
        For i = 0 To UBound(mAnchos): .ColWidth(i) = mAnchos(i): Next

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


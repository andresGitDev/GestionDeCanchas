VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmLisVentas 
   Caption         =   "Listado Ventas Clientes"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10980
   Icon            =   "frmLisVentas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   10980
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCabecera 
      Height          =   1140
      Left            =   45
      TabIndex        =   7
      Top             =   465
      Width           =   8175
      Begin VB.CommandButton cmdVer 
         Caption         =   "&Mostrar"
         Height          =   345
         Left            =   3075
         TabIndex        =   3
         Top             =   195
         Width           =   1260
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   345
         Left            =   5625
         TabIndex        =   8
         Top             =   195
         Width           =   1260
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   345
         Left            =   4365
         TabIndex        =   5
         Top             =   195
         Width           =   1260
      End
      Begin Gestion.ucXls ucXls1 
         Height          =   900
         Left            =   6900
         TabIndex        =   4
         Top             =   180
         Width           =   960
         _extentx        =   1693
         _extenty        =   1588
      End
      Begin Gestion.ucFecha uFechaD 
         Height          =   345
         Left            =   645
         TabIndex        =   1
         Top             =   195
         Width           =   1080
         _extentx        =   1905
         _extenty        =   609
         fechainit       =   0
      End
      Begin Gestion.ucFecha uFechaH 
         Height          =   345
         Left            =   1785
         TabIndex        =   2
         Top             =   195
         Width           =   1080
         _extentx        =   1905
         _extenty        =   609
         fechainit       =   4
      End
      Begin VB.Label Label1 
         Caption         =   "Entre"
         Height          =   285
         Index           =   0
         Left            =   75
         TabIndex        =   9
         Top             =   240
         Width           =   1020
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Height          =   5580
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1710
      Width           =   10815
      _cx             =   19076
      _cy             =   9842
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
      FormatString    =   $"frmLisVentas.frx":08CA
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
   Begin Gestion.ucCoDe uCliente 
      Height          =   330
      Left            =   645
      TabIndex        =   0
      Top             =   75
      Width           =   6690
      _extentx        =   11800
      _extenty        =   582
      codigowidth     =   1000
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente"
      Height          =   285
      Index           =   1
      Left            =   60
      TabIndex        =   10
      Top             =   150
      Width           =   810
   End
End
Attribute VB_Name = "frmLisVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum grigri
    griDOC = 0
    griPROD = 4
    griTOT = 7
End Enum
'

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
    
    FrmImpresiones.VSPrinter.StartDoc
    FrmImpresiones.VSPrinter.Paragraph = "Listado ventas para " & uCliente.DESCRIPCION
    FrmImpresiones.VSPrinter.Paragraph = "Entre fechas : " & uFechaD.dtFecha & " - " & uFechaH.dtFecha
    FrmImpresiones.VSPrinter.Paragraph = " "
    FrmImpresiones.VSPrinter.TextAlign = taLeftBottom
    
    FrmImpresiones.VSPrinter.RenderControl = grilla.hWnd

    FrmImpresiones.VSPrinter.Footer = "||Pagina %d "
    FrmImpresiones.VSPrinter.Zoom = 100
    FrmImpresiones.VSPrinter.EndDoc
    
    FrmImpresiones.Show
    grilla.GridLines = flexGridFlat
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdVer_Click()
    If ON_ERROR_HABILITADO Then On Error GoTo ufaChe
    Dim ss As String
    
    relojito True
    
    ss = " SELECT F.TipoDoc AS Doc, F.NroFactura AS Numero, F.Fecha AS Fecha, " & _
        " D.Cantidad AS Cant, D.Producto, P.descripcion, " & _
        " D.PrecioUnitario AS [   P Unitario], D.PrecioTotal AS [     P Total] " & _
        " FROM  FacturaVenta F INNER JOIN " & _
        " FacturaVentaDetalle D ON F.Codigo = D.CodigoFactura INNER JOIN " & _
        " Producto P ON D.Producto = P.codigo " & _
        " WHERE F.Activo = 1 AND F.TipoDoc in  ('FAA','FAB','FEA','FEB', 'CEA','CEB','DEA','DEB','NCA','NCB','NDA','NDB') "

    If uCliente.codigo > 0 Then ss = ss & " and f.cliente = " & uCliente.codigo
    
    ss = ss & " and f.fecha between " & uFechaD.ssFecha & " and " & uFechaH.ssFecha
    
    ss = ss & " order by numero"

    LlenarGrilla grilla, ss, False, , griTOT
    cambio_Signo_FormatoProd
    
    grillaSumarizo grilla, Array(griTOT)
    grillaWidth grilla, Array(700, 730, 1200, 700, 1200, 2600, 1400, 1400)

    ucXls1.aTitulo = "Listado ventas para " & uCliente.DESCRIPCION & " entre fechas " & uFechaD.dtFecha & " " & uFechaH.dtFecha
    
fin:
    relojito False
    Exit Sub
ufaChe:
    ufa "err en la consulta ", " lisventas " & uCliente.codigo
    Resume fin
End Sub

Private Sub cambio_Signo_FormatoProd()
    Dim i As Long, s As String, ss As String
    With grilla
        For i = 1 To .rows - 1
            If Left(.TextMatrix(i, griDOC), 2) = "NC" Or Left(.TextMatrix(i, griDOC), 2) = "CE" Then
                .TextMatrix(i, griTOT) = " -" & Trim(.TextMatrix(i, griTOT))
            End If
            .TextMatrix(i, griPROD) = reformateoProducto(.TextMatrix(i, griPROD))
        Next i
    End With
End Sub
Private Sub Form_Load()
    uCliente.ini "select descripcion from clientes where codigo = '###' and activo = 1", "Select codigo, descripcion as [Cliente                      ] from clientes where activo = 1 order by codigo ", False
    ucXls1.ini grilla, "c:\LisVentas"
    Form_Resize
End Sub

Private Sub Form_Resize()
    Anclar grilla, Me, anclarLadosTodos
End Sub
'Public Function reformateoProducto(cual As String) As String
'    If Trim(cual) = "" Then Exit Function
'
'    reformateoProducto = Left(cual, 3) & "-" & Mid(cual, 4, 3) & "-" & Mid(cual, 7)
'End Function

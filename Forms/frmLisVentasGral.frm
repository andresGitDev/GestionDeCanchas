VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmLisVentasGral 
   Caption         =   "Informacion Ventas Gral"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10230
   Icon            =   "frmLisVentasGral.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      Caption         =   "Agregar NC sin referenciar"
      Height          =   255
      Left            =   1440
      TabIndex        =   17
      Top             =   945
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      Caption         =   "agregar ND"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   945
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Mostrar"
      Height          =   375
      Left            =   2235
      TabIndex        =   2
      Top             =   75
      Width           =   975
   End
   Begin VB.Frame fraOpc 
      Height          =   720
      Left            =   3075
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   2160
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
         Height          =   375
         Index           =   7
         Left            =   4530
         TabIndex        =   15
         Top             =   225
         Width           =   1245
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
         Height          =   375
         Index           =   6
         Left            =   3645
         TabIndex        =   14
         Top             =   180
         Width           =   1245
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
         Height          =   375
         Index           =   5
         Left            =   2910
         TabIndex        =   13
         Top             =   195
         Width           =   1245
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
         Height          =   375
         Index           =   4
         Left            =   2250
         TabIndex        =   12
         Top             =   210
         Width           =   1245
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
         Height          =   375
         Index           =   3
         Left            =   1605
         TabIndex        =   11
         Top             =   255
         Width           =   1245
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
         Height          =   375
         Index           =   2
         Left            =   975
         TabIndex        =   10
         Top             =   225
         Width           =   1245
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
         Height          =   375
         Index           =   1
         Left            =   570
         TabIndex        =   9
         Top             =   180
         Width           =   1245
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
         Height          =   375
         Index           =   0
         Left            =   90
         TabIndex        =   8
         Top             =   240
         Width           =   1245
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4425
      TabIndex        =   5
      Top             =   75
      Width           =   975
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   3225
      TabIndex        =   4
      Top             =   75
      Width           =   1185
   End
   Begin Gestion.ucXls ucXls1 
      Height          =   750
      Left            =   5430
      TabIndex        =   3
      Top             =   90
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   1323
   End
   Begin Gestion.ucFecha uFeD 
      Height          =   360
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   635
      FechaInit       =   0
   End
   Begin Gestion.ucFecha uFeH 
      Height          =   360
      Left            =   1065
      TabIndex        =   1
      Top             =   75
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   635
      FechaInit       =   4
   End
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Align           =   2  'Align Bottom
      Height          =   5640
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1650
      Width           =   10230
      _cx             =   18045
      _cy             =   9948
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
      FormatString    =   $"frmLisVentasGral.frx":08CA
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
Attribute VB_Name = "frmLisVentasGral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mSql As String
Private Const griPRO = 1
Private Const griTOT = 3
Private Const griCAN = 4

'Private Sub cmdMostrar_Click()
'    s = " SELECT f.TipoDoc, f.NroFactura AS Nro, f.Fecha, f.RazonSocial, " & _
'        " f.Neto, f.Total, f.Saldo, " & _
'        " d.Producto AS CodProducto, " & _
'        " p.descripcion AS DesProducto, p.grupo, p.SubGrupo , d.cantidad " & _
'        " FROM FacturaVenta f INNER JOIN " & _
'        " FacturaVentaDetalle d ON f.Codigo = d.CodigoFactura INNER JOIN " & _
'        " Producto p ON p.codigo = d.Producto " & _
'        " WHERE     (f.Activo = 1) AND (f.TipoDoc LIKE 'FA%%') OR " & _
'        " (f.Activo = 1) AND (f.TipoDoc LIKE 'NC%%') "
'
'    s = s & " and f.fecha  " & ssBetween(uFeD.dtfecha, uFeH.dtfecha)
'
'    s = s & " order by nro "
'
    'LlenarGrilla grilla, s, False
'End Sub

Private Sub cmdImprimir_Click()

    If GRILLA.rows < 2 Then Exit Sub
    
        GRILLA.GridLines = flexGridNone
        GRILLA.GridLinesFixed = flexGridNone
        
        FrmImpresiones.VSPrinter.Orientation = orPortrait ' orLandscape
        FrmImpresiones.VSPrinter.PaperSize = pprA4
        FrmImpresiones.VSPrinter.Preview = True
        FrmImpresiones.VSPrinter.Font.Name = GRILLA.Font.Name
        FrmImpresiones.VSPrinter.FontSize = 12
        FrmImpresiones.VSPrinter.Header = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
        FrmImpresiones.VSPrinter.FontSize = 8
        
        FrmImpresiones.VSPrinter.StartDoc
        FrmImpresiones.VSPrinter.Paragraph = "Listado de ventas General x Producto" ' & uCliente.descripcion
        FrmImpresiones.VSPrinter.Paragraph = "Entre fechas : " & uFeD.strFecha & " - " & uFeH.strFecha
        FrmImpresiones.VSPrinter.Paragraph = " "
        FrmImpresiones.VSPrinter.TextAlign = taLeftBottom
        
        FrmImpresiones.VSPrinter.RenderControl = GRILLA.hWnd
    
        FrmImpresiones.VSPrinter.Footer = "||Pagina %d "
        FrmImpresiones.VSPrinter.Zoom = 100
        FrmImpresiones.VSPrinter.EndDoc
        
        FrmImpresiones.Show
        GRILLA.GridLines = flexGridFlat

End Sub

Private Sub cmdOk_Click()
    Dim s As String
    Dim s2 As String
    Dim s3 As String
    
    s = " SELECT f.TipoDoc tdoc, d.Producto, p.descripcion AS ProdDescripcion, " & _
        " SUM(d.PrecioTotal) AS TotPrecio, SUM(d.Cantidad) AS TotCant " & _
        " FROM         FacturaVenta f INNER JOIN " & _
        " FacturaVentaDetalle d ON f.Codigo = d.CodigoFactura INNER JOIN " & _
        " Producto p ON p.codigo = d.Producto " & _
        " WHERE (f.Activo = 1) AND (f.TipoDoc LIKE 'FA%%' OR f.TipoDoc LIKE 'NC%%' OR f.TipoDoc LIKE 'FE%%' OR f.TipoDoc LIKE 'CE%%') "
        
    s = s & AndFechas
    
    s = s & " GROUP BY f.TipoDoc, d.Producto, p.descripcion "
    
    s = s & " order by totcant desc "
    
    s2 = ""
    s3 = ""
    If Check1.Value = 1 Then
        s2 = " SELECT f.TipoDoc tdoc,' ','NOTAS DE DEBITOS' AS proddescripcion, " & _
            " SUM(f.Total) AS TotPrecio,count(f.tipodoc) as TotCant FROM  " & _
            " FacturaVenta f WHERE (f.Activo = 1) AND (f.TipoDoc LIKE 'ND%%') OR (f.TipoDoc LIKE 'DE%%')  " & _
            " " & AndFechas & " GROUP BY f.TipoDoc "
    End If
    If Check2.Value = 1 Then
        s3 = " SELECT f.TipoDoc tdoc,' ','NOTAS DE CREDITOS SIN ASIG.' AS proddescripcion," & _
            " SUM(f.Total) AS TotPrecio,count(f.tipodoc) as TotCant" & _
            " FROM         FacturaVenta f WHERE f.nrofactura not in (select facturaventadetalle.nrofactura " & _
            " from facturaventadetalle where facturaventadetalle.nrofactura=f.nrofactura) and " & _
            " (f.Activo = 1) AND (f.TipoDoc LIKE 'NC%%' OR f.TipoDoc LIKE 'CE%%') " & AndFechas & " GROUP BY f.TipoDoc "
    End If

    LlenarGrilla2 GRILLA, s, s2, s3, False
    
    cambioSigno
    
    grillaWidth GRILLA, Array(1200, 1755, 3570, 1395, 1380)
    
    grillaSumarizo GRILLA, Array(3)
End Sub

Private Function AndFechas() As String
    AndFechas = " and f.fecha  " & ssBetween(uFeD.dtFecha, uFeH.dtFecha)
End Function

Private Sub cambioSigno()
    Dim i As Long
    With GRILLA
        For i = 1 To .rows - 1
            If Left(.TextMatrix(i, 0), 2) = "NC" Or Left(.TextMatrix(i, 0), 2) = "CE" Then
                .TextMatrix(i, griTOT) = " -" & Trim(.TextMatrix(i, griTOT))
                .TextMatrix(i, griCAN) = " -" & Trim(.TextMatrix(i, griCAN))
            End If
            
            .TextMatrix(i, griPRO) = reformateoProducto(.TextMatrix(i, griPRO))
                
        Next i
    End With
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ucXls1.ini GRILLA, "c:\LisVentasProd"
    Form_Resize
'    AcomodarArrayEnFrame fraOpc, cmdMostrar, Array("Producto", "fecha", "otro")
End Sub

Private Sub Form_Resize()
    Anclar GRILLA, Me, anclarArriba
End Sub


Private Sub uFeD_LostFocus()
    'uFeH.setUltDiaMes uFeD.Mes, uFeD.Anio
End Sub


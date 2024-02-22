VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmVentasProductoCliente 
   Caption         =   "Ventas Producto Cliente"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18705
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   18705
   StartUpPosition =   2  'CenterScreen
   Begin Gestion.ucXls ucXls1 
      Height          =   855
      Left            =   7950
      TabIndex        =   8
      Top             =   30
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   1508
   End
   Begin VSFlex7LCtl.VSFlexGrid gVentas 
      Height          =   6990
      Left            =   105
      TabIndex        =   7
      Top             =   1455
      Width           =   18405
      _cx             =   32464
      _cy             =   12330
      _ConvInfo       =   -1
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
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
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
   Begin VB.CommandButton cmdMostrar 
      Caption         =   "Mostrar"
      Height          =   360
      Left            =   9150
      TabIndex        =   6
      Top             =   525
      Width           =   1395
   End
   Begin Gestion.ucCoDe uProd 
      Height          =   300
      Left            =   960
      TabIndex        =   4
      Top             =   1095
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   529
      CodigoWidth     =   1455
      CodigoInvalido  =   0
   End
   Begin MSComCtl2.DTPicker dtDesde 
      Height          =   375
      Left            =   945
      TabIndex        =   0
      Top             =   135
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   661
      _Version        =   393216
      Format          =   153747457
      CurrentDate     =   44070
   End
   Begin MSComCtl2.DTPicker dtHasta 
      Height          =   375
      Left            =   945
      TabIndex        =   1
      Top             =   600
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   661
      _Version        =   393216
      Format          =   153747457
      CurrentDate     =   44070
   End
   Begin VB.Label Label3 
      Caption         =   "Producto"
      Height          =   330
      Left            =   75
      TabIndex        =   5
      Top             =   1110
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Hasta"
      Height          =   210
      Left            =   90
      TabIndex        =   3
      Top             =   615
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "Desde"
      Height          =   210
      Left            =   75
      TabIndex        =   2
      Top             =   180
      Width           =   885
   End
End
Attribute VB_Name = "frmVentasProductoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdMostrar_Click()
Dim ss As String, ssW As String
ss = "select  DISTINCT F.TIPODOC,  F.PUNTOVENTA, F.NROFACTURA ,F.FECHA,F.CLIENTE,F.CUIT,F.RAZONSOCIAL,D.CANTIDAD,d.precioTOTAL as TOTAL,C.MAIL,C.TELEFONO FROM (FACTURAVENTA AS F INNER JOIN FACTURAVENTADETALLE AS D ON F.CODIGO=D.CODIGOFACTURA) inner join CLIENTES AS C ON F.CLIENTE=C.CODIGO WHERE F.FECHA>=" & ssFecha(dtDesde) & " AND F.FECHA<=" & ssFecha(dtHasta) & " "
ssW = ""
If Trim(uProd.codigo) > "" Then
    ssW = " and d.producto=" & ssTexto(uProd.codigo)
End If
ss = ss & ssW
LlenarGrilla gVentas, ss, False
If gVentas.rows > 1 Then
    With gVentas
        .ColWidth(0) = 1200
        .ColWidth(1) = 1300
        .ColWidth(2) = 1300
        .ColWidth(3) = 1300
        .ColWidth(4) = 1300
        .ColWidth(5) = 1500
        .ColWidth(6) = 2500
        .ColWidth(7) = 1000
        .ColWidth(8) = 1500
        .ColWidth(9) = 2700
        .ColWidth(10) = 2700
    End With
End If
End Sub

Private Sub dtHasta_Change()
If dtHasta < dtDesde Then dtHasta = dtDesde
End Sub

Private Sub Form_Load()
dtDesde = CDate("01/01/" & Year(Date))
dtHasta = Date
uProd.ini "select descripcion from producto where codigo = '###' ", "select codigo as [ Codigo                 ],  descripcion as [ Descripcion                                                 ] from producto where activo = 1 and facturable=1 order by codigo ", True
gVentas.rows = 1
ucXls1.ini gVentas, "VentasProductoCliente.xls"
End Sub

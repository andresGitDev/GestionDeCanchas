VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmLisVentasGral 
   Caption         =   "Informacion Ventas Gral"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   10410
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraOpc 
      Height          =   720
      Left            =   1065
      TabIndex        =   6
      Top             =   90
      Width           =   6000
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
         Height          =   375
         Index           =   7
         Left            =   4530
         TabIndex        =   14
         Top             =   225
         Width           =   1245
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
         Height          =   375
         Index           =   6
         Left            =   3645
         TabIndex        =   13
         Top             =   180
         Width           =   1245
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
         Height          =   375
         Index           =   5
         Left            =   2910
         TabIndex        =   12
         Top             =   195
         Width           =   1245
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
         Height          =   375
         Index           =   4
         Left            =   2250
         TabIndex        =   11
         Top             =   210
         Width           =   1245
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
         Height          =   375
         Index           =   3
         Left            =   1605
         TabIndex        =   10
         Top             =   255
         Width           =   1245
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
         Height          =   375
         Index           =   2
         Left            =   975
         TabIndex        =   9
         Top             =   225
         Width           =   1245
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
         Height          =   375
         Index           =   1
         Left            =   570
         TabIndex        =   8
         Top             =   180
         Width           =   1245
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
         Height          =   375
         Index           =   0
         Left            =   90
         TabIndex        =   7
         Top             =   240
         Width           =   1245
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8970
      TabIndex        =   2
      Top             =   660
      Width           =   975
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   360
      Left            =   7245
      TabIndex        =   0
      Top             =   645
      Width           =   1185
   End
   Begin GestionTonka.ucXls ucXls1 
      Height          =   360
      Left            =   6030
      TabIndex        =   1
      Top             =   645
      Width           =   1170
      _extentx        =   2064
      _extenty        =   635
   End
   Begin GestionTonka.ucFecha uFeD 
      Height          =   360
      Left            =   60
      TabIndex        =   3
      Top             =   105
      Width           =   960
      _extentx        =   1693
      _extenty        =   635
      fechainit       =   1
   End
   Begin GestionTonka.ucFecha uFeH 
      Height          =   360
      Left            =   60
      TabIndex        =   4
      Top             =   495
      Width           =   945
      _extentx        =   1667
      _extenty        =   635
      fechainit       =   3
   End
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Height          =   5220
      Left            =   75
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1080
      Width           =   9855
      _cx             =   17383
      _cy             =   9208
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
      FormatString    =   $"frmInfVentasGral.frx":0000
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

'Private Sub cmdMostrar_Click()
'    Dim s As String
'
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
'    s = s & " and f.fecha between " & ssBetween(uFeD.dtfecha, uFeH.dtfecha)
'
'    s = s & " order by nro "
'
'    LlenarGrilla grilla, s, False
'
'
'End Sub

Private Sub uFeD_LostFocus()
    uFeH.setUltDiaMes uFeD.mes, uFeD.Anio
End Sub

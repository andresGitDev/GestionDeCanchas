VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmLisProductos2 
   Caption         =   "Informe de Productos Activos"
   ClientHeight    =   11340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14580
   Icon            =   "frmLisProductos2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11340
   ScaleWidth      =   14580
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex7LCtl.VSFlexGrid gExcel 
      Height          =   1995
      Left            =   120
      TabIndex        =   20
      Top             =   11355
      Visible         =   0   'False
      Width           =   14295
      _cx             =   25215
      _cy             =   3519
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
   Begin VB.Frame Frame2 
      Caption         =   "Ordenar"
      Height          =   1125
      Left            =   11040
      TabIndex        =   11
      Top             =   630
      Width           =   2235
      Begin VB.OptionButton Option5 
         Caption         =   "por Descripcion"
         Height          =   330
         Left            =   315
         TabIndex        =   13
         Top             =   675
         Width           =   1605
      End
      Begin VB.OptionButton Option4 
         Caption         =   "por Codigo"
         Height          =   300
         Left            =   315
         TabIndex        =   12
         Top             =   255
         Value           =   -1  'True
         Width           =   1725
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtros"
      Height          =   1395
      Left            =   5655
      TabIndex        =   7
      Top             =   465
      Width           =   5235
      Begin VB.CommandButton cmdFiltrar 
         Caption         =   "Filtrar"
         Height          =   330
         Left            =   165
         TabIndex        =   16
         Top             =   585
         Width           =   1185
      End
      Begin VB.ComboBox cboSubGrupo 
         Height          =   315
         Left            =   1515
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   990
         Width           =   3555
      End
      Begin VB.ComboBox cboGrupo 
         Height          =   315
         Left            =   1515
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   420
         Width           =   3555
      End
      Begin VB.OptionButton Option3 
         Caption         =   "SubGrupo"
         Height          =   375
         Left            =   1935
         TabIndex        =   10
         Top             =   690
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Grupo"
         Height          =   375
         Left            =   1935
         TabIndex        =   9
         Top             =   105
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Todos"
         Height          =   285
         Left            =   150
         TabIndex        =   8
         Top             =   195
         Value           =   -1  'True
         Width           =   1260
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   795
      Left            =   2115
      Picture         =   "frmLisProductos2.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   75
      Width           =   855
   End
   Begin Gestion.ucXls ucXls1 
      Height          =   795
      Left            =   1065
      TabIndex        =   3
      Top             =   75
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   1402
   End
   Begin VSFlex7LCtl.VSFlexGrid gDetalle 
      Height          =   4350
      Left            =   120
      TabIndex        =   2
      Top             =   6915
      Width           =   5550
      _cx             =   9790
      _cy             =   7673
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
   Begin VSFlex7LCtl.VSFlexGrid gProductos 
      Height          =   4650
      Left            =   120
      TabIndex        =   1
      Top             =   1950
      Width           =   14370
      _cx             =   25347
      _cy             =   8202
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
      Cols            =   10
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
      AutoSearch      =   1
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
   Begin VB.CommandButton cmdVer 
      Caption         =   "Mostrar"
      Height          =   795
      Left            =   120
      Picture         =   "frmLisProductos2.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   60
      Width           =   765
   End
   Begin Gestion.ucCoDe uProdD 
      Height          =   360
      Left            =   150
      TabIndex        =   5
      Top             =   975
      Width           =   5100
      _ExtentX        =   16007
      _ExtentY        =   635
      CodigoWidth     =   1000
   End
   Begin Gestion.ucCoDe uProdH 
      Height          =   390
      Left            =   150
      TabIndex        =   6
      Top             =   1455
      Width           =   5085
      _ExtentX        =   16034
      _ExtentY        =   688
      CodigoWidth     =   1000
   End
   Begin VSFlex7LCtl.VSFlexGrid gDetalle2 
      Height          =   4350
      Left            =   5745
      TabIndex        =   17
      Top             =   6900
      Width           =   8730
      _cx             =   15399
      _cy             =   7673
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
      AutoSearch      =   1
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
   Begin VB.Label lbltm 
      Caption         =   "Movimiento"
      Height          =   285
      Left            =   5775
      TabIndex        =   19
      Top             =   6660
      Width           =   5340
   End
   Begin VB.Label lbltf 
      Caption         =   "Formula"
      Height          =   255
      Left            =   150
      TabIndex        =   18
      Top             =   6675
      Width           =   4170
   End
End
Attribute VB_Name = "frmLisProductos2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const tt_Productos_Temp = "( [CODPRO] [varchar] (4000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL  , [DESCRIPCION] [varchar] (4000) NULL , [GRUPO] [VARCHAR] (4000) NULL ,[SUBGRUPO] [VARCHAR] (4000) NULL,[PRECIO] [VARCHAR] (4000) NULL, [ALIAS] [VARCHAR] (4000) NULL,[TIENE_CUENTA] [VARCHAR] (4000) NULL,[CUENTA] [VARCHAR] (4000) NULL,[UNIDADMEDIDA] [VARCHAR] (4000) NULL,[FACTOR_STOCK] [VARCHAR] (4000) NULL,[FACTOR_PARTE] [VARCHAR] (4000) NULL )"
Private Const tt_Productos_Excel = "( [C1] [varchar] (4000) NULL  , [C2] [varchar] (4000) NULL , [C3] [VARCHAR] (4000) NULL ,[C4] [VARCHAR] (4000) NULL,[C5] [VARCHAR] (4000) NULL, [C6] [VARCHAR] (4000) NULL,[C7] [VARCHAR] (4000) NULL,[C8] [VARCHAR] (4000) NULL,[C9] [VARCHAR] (4000) NULL,[C10] [VARCHAR] (4000) NULL)"
Private tProd As String

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



Private Sub cboGrupo_Click()
Option2.Value = True
cmdFiltrar_Click
End Sub

Private Sub cboGrupo_DblClick()
Option2.Value = True
cmdFiltrar_Click
End Sub


Private Sub cbosubGrupo_Click()
Option3.Value = True
cmdFiltrar_Click
End Sub

Private Sub cbosubGrupo_DblClick()
Option3.Value = True
cmdFiltrar_Click
End Sub

Private Sub cmdFiltrar_Click()
If tProd = "" Then Exit Sub
Dim sWhere As String, sOrder As String, sConsul As String

If Option1 Then
    sWhere = ""
ElseIf Option2 Then
    sWhere = " WHERE GRUPO=" & sstexto(cboGrupo.Text) & "  "
ElseIf Option3 Then
    sWhere = " WHERE SUBGRUPO=" & sstexto(cboSubGrupo.Text) & "  "
End If

If Option4 Then
    sOrder = " ORDER BY CODPRO"
ElseIf Option5 Then
    sOrder = " ORDER BY DESCRIPCION"
End If

gProductos.rows = 1
sConsul = "Select CODPRO AS CODIGO,DESCRIPCION,GRUPO,SUBGRUPO,ALIAS,TIENE_CUENTA AS USA_CUENTA,CUENTA,UNIDADMEDIDA AS UNIDAD_MEDIDA,FACTOR_STOCK,FACTOR_PARTE from " & tProd & " "
LlenarGrilla gProductos, sConsul & sWhere & sOrder, True

With gProductos
    .ColWidth(0) = 1200
    .ColWidth(1) = 3000
    .ColWidth(2) = 3000
    .ColWidth(3) = 3000
    .ColWidth(4) = 3000
    .ColWidth(5) = 1500
    .ColWidth(6) = 1200
    .ColWidth(7) = 2000
    .ColWidth(8) = 3500
    .ColWidth(9) = 3500
End With



End Sub

Private Sub cmdVer_Click()
    Ver
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    uProd_ini
    Ver
    ucXls1.ini gProductos, "C:\PRODUCTOS2.XLS"
    
End Sub

Private Sub uProd_ini()
Dim sqlbuscar As String, sqldesc As String

sqldesc = "select descripcion from producto where codigo = '###' "
sqlbuscar = "select p.codigo as [ Codigo                 ],  p.descripcion as [ Descripcion                                                 ],m.abreviatura as [  Medida  ] from producto as p inner join unidadesmedida as m on p.umedida=m.umcodigo where p.activo = 1 and formula=1 order by p.codigo "

uProdD.ini sqldesc, sqlbuscar, True
uProdH.ini sqldesc, sqlbuscar, True

Dim MAXMIN
MAXMIN = obtenerDeSQL("SELECT MIN(CODIGO),MAX(CODIGO) FROM PRODUCTO WHERE ACTIVO=1 ")
uProdD.codigo = MAXMIN(0)
uProdH.codigo = MAXMIN(1)
End Sub



Private Function Ver()
Dim cver As String, dProd As Long
Dim rsPro As New ADODB.Recordset, i As Long, X As Long
Dim rsParte As New ADODB.Recordset, d
Dim rsDatos As New ADODB.Recordset

Dim grupo As String, SubGrupo As String, Tiene As String, UMedida As String, Factor1 As String, Factor2 As String
tProd = TablaTempCrear(tt_Productos_Temp)
If uProdD.codigo = 0 And uProdH.codigo = 0 Then Exit Function
With rsPro
    .Open "select * from producto where activo=1 and codigo >=" & sstexto(uProdD.codigo) & " and codigo<=" & sstexto(uProdH.codigo) & " order by codigo", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
    If .EOF And .BOF Then
    Else
        .MoveFirst
        For i = 0 To .RecordCount - 1
        
            grupo = UCase(obtenerDeSQL("select descripcion from gruposproducto where codigo=" & sstexto(!grupo)))
            SubGrupo = UCase(obtenerDeSQL("select descripcion from subgruposproducto where codigo=" & sstexto(!SubGrupo)))
            Tiene = IIf(!tiene_Cuenta = 1, "Si", "No")
            UMedida = UCase(obtenerDeSQL("select descripcion from unidadesmedida where umcodigo=" & nSinNull(!UMedida)))
            Factor1 = UCase(obtenerDeSQL("select caracteristica + ' (' +   cast(factor as varchar)   + ')'   from umfactor where ufcodigo=" & nSinNull(!uFactor)))
            Factor2 = UCase(obtenerDeSQL("select caracteristica + ' (' +   cast(factor as varchar)   + ')'   from umfactor where ufcodigo=" & nSinNull(!Uparte)))
            insertTabla tProd, !codigo, sSinNull(!DESCRIPCION), grupo, SubGrupo, !precio, sSinNull(!alias), Tiene, sSinNull(!CUENTA), UMedida, Factor1, Factor2
            .MoveNext
        Next
    End If
End With


    gProductos.rows = 1
    cver = "Select CODPRO AS CODIGO,DESCRIPCION,GRUPO,SUBGRUPO,PRECIO,ALIAS,TIENE_CUENTA AS USA_CUENTA,CUENTA,UNIDADMEDIDA AS UNIDAD_MEDIDA,FACTOR_STOCK,FACTOR_PARTE from " & tProd & " "
    LlenarGrilla gProductos, cver, True
    gDetalle.rows = 1
    gDetalle2.rows = 1
    gExcel.rows = 1
    'For dProd = 1 To gProductos.rows - 1
    '    d = nSinNull(obtenerDeSQL("select f.factor from producto p inner join umfactor f on p.ufactor=f.ufcodigo where p.codigo=" & sstexto(gProductos.TextMatrix(dProd, 2))))
    '    If d = 0 Then d = 1
    '    gProductos.TextMatrix(dProd, 4) = s2n(s2n(gProductos.TextMatrix(dProd, 5)) / d)
    'Next

    With gProductos
        .ColWidth(0) = 1200
        .ColWidth(1) = 3000
        .ColWidth(2) = 3000
        .ColWidth(3) = 3000
        .ColWidth(4) = 1500
        .ColWidth(5) = 3000
        .ColWidth(6) = 1500
        .ColWidth(7) = 1200
        .ColWidth(8) = 2000
        .ColWidth(9) = 3500
        .ColWidth(10) = 3500
        
    End With
    
cboGrupo.clear
rsDatos.Open "select * from gruposproducto where activo=1 order by descripcion", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
With rsDatos
    If .EOF And .BOF Then
        cboGrupo.AddItem "NO HAY GRUPOS CARGADOS"
    Else
        .MoveFirst
        For i = 0 To .RecordCount - 1
            cboGrupo.AddItem !DESCRIPCION
            .MoveNext
        Next
    End If
End With
Set rsDatos = Nothing

cboSubGrupo.clear
rsDatos.Open "select * from subgruposproducto where activo=1 order by descripcion", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
With rsDatos
    If .EOF And .BOF Then
        cboSubGrupo.AddItem "NO HAY SUBGRUPOS CARGADOS"
    Else
        .MoveFirst
        For i = 0 To .RecordCount - 1
            cboSubGrupo.AddItem !DESCRIPCION
            .MoveNext
        Next
    End If
End With
Set rsDatos = Nothing

End Function

Private Sub AgregarEnTabla(fecha As Date, TipoComprobante As String, NroComprobante As Variant, _
                            cantidad As Double, saldo As Double, producto As String)
Dim Consulta As String

    Consulta = "Insert into MOVIMIENTO_STOCK_TEMP (FECHA, TIPOCOMPROBANTE, NROCOMPROBANTE, CANTIDAD, SALDO, PRODUCTO) " & _
                    "values (" & ssFecha(fecha) & ", '" & TipoComprobante & "', '" & NroComprobante & "', ' " & _
                    cantidad & "', '" & saldo & "', '" & producto & "')"
    DataEnvironment1.Sistema.Execute Consulta
    
End Sub

Private Function cmdCargarEnTemp(en_vista As String, Optional gVer As Boolean = True) As Boolean
Dim cantidad As Double
Dim a As String
Dim CodigoProd As String
Dim tiene_algo As Integer
Dim Consulta As String
Dim CodProd As String, CodProd2 As String
Dim rsmov As New ADODB.Recordset
Dim rsAux As New ADODB.Recordset
Dim SaldoProd As Double, SaldoProd2 As Double
Dim cant As Double, i As Long

    If en_vista <> "" And en_vista <> "0" Then
        tiene_algo = 0
        DataEnvironment1.Sistema.Execute "Delete From MOVIMIENTO_STOCK_TEMP"
        
        CodProd = Trim(en_vista)
             
        'SaldoProd = CalcularSaldoAnterior(CodProd, dtfechad.Value - 1)
        
        'AgregarEnTabla dtfechad.Value - 1, const_SALDO_INICIAL & " " & ObtenerDescripcionS("PRODUCTO", Trim(CodProd)), 0, SaldoProd, 0, CodProd
        
        CodProd2 = "select * from partesproduccion p where p.producido='" & CodProd & "'" _
            & " and p.ACTIVO = 1 and p.confirmacion " & ssBetween(CDate("01/01/1900"), Date)
            
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
            & " and p.ACTIVO = 1 and f.activo=1 and p.confirmacion " & ssBetween(CDate("01/01/1900"), Date)
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
               

        'TABLA REMITO COMPRA
        Consulta = "select R.NroRemito, R.FECHA, D.PRODUCTO, D.CANTIDAD " & _
                    "from REMITOCOMPRADETALLE as D INNER JOIN REMITOCOMPRA AS R ON R.CODIGO = D.CODIGOREMITO " & _
                    "where R.ACTIVO = 1 and " & _
                    "D.PRODUCTO = '" & CodProd & "' and " & _
                    "R.FECHA " & ssBetween(CDate("01/01/1900"), Date)
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
            & " R.FECHA " & ssBetween(CDate("01/01/1900"), Date)
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
                    " R.FECHA " & ssBetween(CDate("01/01/1900"), Date) & " and R.anulado=0"
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
                    "R.FECHA " & ssBetween(CDate("01/01/1900"), Date) & " and activo=1"
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


    Else
'        AgregarEnTabla Date, "S/N", 0, 0, 0, CodProd
    End If
    
    

    
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
                rsAux!saldo = x2s(s2n(a))
                rsAux.Update
                If rsAux!NroComprobante = "-" Then cantidad = 0
                rsAux.MoveNext
            Wend
        End If
        Set rsAux = Nothing
     
    'AgregarEnTabla Date + 1, "->>FINAL:" & ObtenerDescripcionS("PRODUCTO", Trim(CodProd)), "-", 0, 0, CodProd
    If tiene_algo = 0 Then
        cmdCargarEnTemp = False
        AgregarEnTabla Date, "S/N", 0, 0, 0, CodProd
    Else
        cmdCargarEnTemp = True
    End If
    
    If gVer Then
    
        LlenarGrilla gDetalle2, "Select FECHA, TIPOCOMPROBANTE AS 'TIPO DE COMPROBANTE', NROCOMPROBANTE AS 'NRO', CANTIDAD, SALDO " & _
                                        "From MOVIMIENTO_STOCK_TEMP order by producto, fecha", True   'Order By NROCOMPROBANTE
                                        
        If gDetalle2.rows > 1 Then
            With gDetalle2
                .ColWidth(0) = 1000
                .ColWidth(1) = 4000
                .ColWidth(2) = 1000
                .ColWidth(3) = 1000
                .ColWidth(4) = 1000
                'marco iniciales
                'For i = 1 To .rows - 1
                '    If Mid(.TextMatrix(i, 1), 1, 1) = "-" Then
                '        .cell(flexcpFontBold, i, 1) = True
                '    End If
                'Next
                .Row = .rows - 1
                .Select .rows - 1, 4
                .cell(flexcpFontBold, .rows - 1, 4) = True
            End With
        End If
    End If
End Function


Private Function insertTabla(pTABLA As String, pCODPRO As String, pDescripcion As String, pGrupo As String, pSubgrupo As String, pPRECIO As String, pAlias As String, pTIENE As String, pCuenta As String, pMEDIDA As String, pFACTOR_STOCK As String, pFACTOR_PARTE As String)
Dim cad As String
If Trim(pTABLA) = "" Then Exit Function

cad = " insert into " & pTABLA & " ( CODPRO,DESCRIPCION,GRUPO,SUBGRUPO,PRECIO,ALIAS,TIENE_CUENTA,CUENTA,UNIDADMEDIDA,FACTOR_STOCK,FACTOR_PARTE ) VALUES (" & _
        " " & sstexto(pCODPRO) & "," & sstexto(pDescripcion) & "," & sstexto(pGrupo) & "," & sstexto(pSubgrupo) & "," & sstexto(x2s(pPRECIO)) & "," & sstexto(pAlias) & "," & sstexto(pTIENE) & "," & sstexto(pCuenta) & "," & sstexto(pMEDIDA) & "," & sstexto(pFACTOR_STOCK) & "," & sstexto(pFACTOR_PARTE) & ")"

DataEnvironment1.Sistema.Execute cad
End Function

Private Function VerD()
Dim dVer As String, dPRODUCTO As String, dComp As Long, d
    gDetalle.rows = 1
    If gProductos.Row = 0 Then
        dPRODUCTO = 0
    Else
        dPRODUCTO = gProductos.TextMatrix(gProductos.Row, 0)
    End If
    dVer = "Select F.COMPONENTE, P.DESCRIPCION,F.CANTIDAD AS VOLUMEN from Formulas F INNER JOIN PRODUCTO P ON F.COMPONENTE=P.CODIGO where F.activo=1 AND F.CODIGO=" & sstexto(dPRODUCTO)
    LlenarGrilla gDetalle, dVer, False
    cmdCargarEnTemp dPRODUCTO
    
    With gDetalle
        .ColWidth(0) = 1500
        .ColWidth(1) = 2500
        .ColWidth(2) = 1000
    End With
    
    If dPRODUCTO <> "0" Then
        lbltf = "Formula " & dPRODUCTO
        lbltm = "Movimiento " & dPRODUCTO & " , Stock : " & gDetalle2.TextMatrix(gDetalle2.rows - 1, gDetalle2.cols - 1)
    Else
        lbltf = "Formula "
        lbltm = "Movimiento "
    End If
    
    'For dComp = 1 To gDetalle.rows - 1
    '    d = nSinNull(obtenerDeSQL("select f.factor from producto p inner join umfactor f on p.ufactor=f.ufcodigo where p.codigo=" & sstexto(gDetalle.TextMatrix(dComp, 0))))
    '    If d = 0 Then d = 1
    '    gDetalle.TextMatrix(dComp, 3) = s2n(s2n(gDetalle.TextMatrix(dComp, 2)) / d)
    'Next
    
End Function

Private Sub gProductos_ChangeEdit()
VerD
End Sub

Private Sub gProductos_Click()
VerD
End Sub

Private Sub gProductos_SelChange()
VerD
End Sub

Private Sub Option1_Click()
cmdFiltrar_Click
End Sub

Private Sub Option1_DblClick()
cmdFiltrar_Click
End Sub

Private Sub Option4_Click()
cmdFiltrar_Click
End Sub

Private Sub Option4_DblClick()
cmdFiltrar_Click
End Sub

Private Sub Option5_Click()
cmdFiltrar_Click
End Sub

Private Sub Option5_DblClick()
cmdFiltrar_Click
End Sub

Private Sub ucXls1_Clic(Cancel As Boolean)
Dim tEXCEL As String, i As Long, rsExcel As New ADODB.Recordset, X As Long
Dim dFormula As Boolean, dMovimientos As Boolean
dFormula = False
dMovimientos = False

If MsgBox("¿Incluir detalle de formula?", vbYesNo + vbInformation) = vbYes Then
    dFormula = True
End If

If MsgBox("¿Incluir detalle de movimietos?", vbYesNo + vbInformation) = vbYes Then
    dMovimientos = True
End If

If dFormula = False And dMovimientos = False Then
    ucXls1.ini gProductos, "C:\PRODUCTOS_SIMPLE.XLS", "PRODUCTOS EXPORTACION SIMPLE"
Else
    tEXCEL = TablaTempCrear(tt_Productos_Excel)
    
    With gProductos
        For i = 1 To .rows - 1
            If i <> 1 Then
                iExcel tEXCEL, "", "", "", "", "", "", "", "", "", ""
            End If
            iExcel tEXCEL, .TextMatrix(i, 0), .TextMatrix(i, 1), .TextMatrix(i, 2), .TextMatrix(i, 3), x2s(.TextMatrix(i, 4)), .TextMatrix(i, 5), .TextMatrix(i, 6), .TextMatrix(i, 7), .TextMatrix(i, 8), .TextMatrix(i, 9)
            If dFormula Then
                rsExcel.Open "Select F.COMPONENTE, P.DESCRIPCION,F.CANTIDAD AS VOLUMEN from Formulas F INNER JOIN PRODUCTO P ON F.COMPONENTE=P.CODIGO where F.activo=1 AND F.CODIGO=" & sstexto(.TextMatrix(i, 0)), DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
                If rsExcel.EOF And rsExcel.BOF Then
                    iExcel tEXCEL, "", "SIN FORMULA", "", "", "", "", "", "", "", ""
                Else
                    iExcel tEXCEL, "", "", "", "", "", "", "", "", "", ""
                    iExcel tEXCEL, "Detalle de Formula", .TextMatrix(i, 0), "", "", "", "", "", "", "", ""
                    rsExcel.MoveFirst
                    For X = 0 To rsExcel.RecordCount - 1
                        iExcel tEXCEL, "", CStr(rsExcel!Componente), CStr(rsExcel!DESCRIPCION), CStr(x2s(rsExcel!volumen)), "", "", "", "", "", ""
                        rsExcel.MoveNext
                    Next
                    
                End If
            End If
            Set rsExcel = Nothing
            
            If dMovimientos Then
                cmdCargarEnTemp .TextMatrix(i, 0), False
                rsExcel.Open "Select FECHA, TIPOCOMPROBANTE , NROCOMPROBANTE, CANTIDAD, SALDO " & _
                                    "From MOVIMIENTO_STOCK_TEMP order by producto, fecha", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
                If rsExcel.EOF And rsExcel.BOF Then
                    iExcel tEXCEL, "", "SIN MOVIMIENTOS", "", "", "", "", "", "", "", ""
                Else
                    iExcel tEXCEL, "", "", "", "", "", "", "", "", "", ""
                    iExcel tEXCEL, "Detalle de Movimiento", .TextMatrix(i, 0), "", "", "", "", "", "", "", ""
                    rsExcel.MoveFirst
                    For X = 0 To rsExcel.RecordCount - 1
                        iExcel tEXCEL, "", CStr(rsExcel!fecha), CStr(rsExcel!TipoComprobante), CStr(rsExcel!NroComprobante), CStr(x2s(rsExcel!cantidad)), CStr(x2s(rsExcel!saldo)), "", "", "", ""
                        rsExcel.MoveNext
                    Next
                    
                End If
            End If
            Set rsExcel = Nothing
            
        Next
    End With
    LlenarGrilla gExcel, "select * from " & tEXCEL, True
    ucXls1.ini gExcel, "C:\PRODUCTOS_COMPLEJO.XLS", "PRODUCTOS EXPORTACION COMPLEJA"
End If
End Sub

Private Function iExcel(CT As String, C1 As String, C2 As String, C3 As String, C4 As String, C5 As String, C6 As String, C7 As String, C8 As String, C9 As String, C10 As String)
Dim C As String
C = "insert into " & CT & " (C1,C2,C3,C4,C5,C6,C7,C8,C9,C10) values (" & _
     " " & sstexto(C1) & "," & sstexto(C2) & "," & sstexto(C3) & "," & sstexto(C4) & "," & sstexto(C5) & "," & sstexto(C6) & "," & sstexto(C7) & "," & sstexto(C8) & "," & sstexto(C9) & "," & sstexto(C10) & ")"

DataEnvironment1.Sistema.Execute C
End Function

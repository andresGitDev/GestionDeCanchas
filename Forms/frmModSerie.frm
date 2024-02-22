VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmModSerie 
   Caption         =   "Ver Series"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8355
   Icon            =   "frmModSerie.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   8355
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Height          =   3855
      Left            =   60
      TabIndex        =   25
      Top             =   1860
      Width           =   8115
      _cx             =   14314
      _cy             =   6800
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmModSerie.frx":08CA
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
   Begin Gestion.ucXls ucXls1 
      Height          =   855
      Left            =   7080
      TabIndex        =   39
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1508
   End
   Begin Gestion.ucCoDe uProducto 
      Height          =   315
      Left            =   2130
      TabIndex        =   38
      Top             =   720
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   556
      CodigoWidth     =   1000
   End
   Begin Gestion.ucFecha uFecha 
      Height          =   315
      Left            =   2175
      TabIndex        =   37
      Top             =   0
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      FechaInit       =   0
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   6840
      TabIndex        =   24
      Top             =   6120
      Width           =   1035
   End
   Begin Gestion.ucBotonera uMenu 
      Height          =   1575
      Left            =   120
      TabIndex        =   36
      Top             =   5160
      Visible         =   0   'False
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   1296
      MsgConfirmaSalir=   ""
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   ""
      CaptionEliminar =   "&Eliminar"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Excel"
      Height          =   375
      Left            =   7080
      TabIndex        =   35
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtDocumento 
      Height          =   315
      Left            =   2130
      TabIndex        =   23
      Top             =   360
      Width           =   1035
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Mostrar"
      Height          =   375
      Left            =   2115
      TabIndex        =   22
      Top             =   1440
      Width           =   990
   End
   Begin VB.CommandButton cmdBuscarDoc 
      Caption         =   "?"
      Height          =   315
      Left            =   3240
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   360
      Width           =   435
   End
   Begin VB.TextBox txtNumero 
      Height          =   315
      Left            =   5760
      TabIndex        =   20
      Top             =   360
      Width           =   1035
   End
   Begin VB.TextBox txtSerie 
      Height          =   315
      Left            =   2130
      TabIndex        =   19
      Top             =   1080
      Width           =   1935
   End
   Begin VB.ListBox lstDetalle 
      Height          =   3765
      Left            =   8340
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1860
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Frame fraCambioSerie 
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   8280
      TabIndex        =   2
      Top             =   1860
      Visible         =   0   'False
      Width           =   3015
      Begin VB.TextBox txtN_NuevaSerie 
         Height          =   375
         Left            =   60
         TabIndex        =   10
         Top             =   3420
         Width           =   2895
      End
      Begin VB.CommandButton cmdSerieEnStock 
         Caption         =   "?"
         Height          =   315
         Left            =   2520
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   3120
         Width           =   435
      End
      Begin VB.TextBox txtN_TipoDoc 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   840
         Width           =   1395
      End
      Begin VB.TextBox txtN_NroDoc 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1200
         Width           =   1395
      End
      Begin VB.TextBox txtN_Prod 
         Enabled         =   0   'False
         Height          =   315
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox txtN_Codigo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   0
         Width           =   1395
      End
      Begin VB.TextBox txtN_ProdDesc 
         Enabled         =   0   'False
         Height          =   315
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox txtn_ntdoc 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Doc"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Index           =   2
         Left            =   60
         TabIndex        =   16
         Top             =   1740
         Width           =   1155
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "En stock"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   2160
         TabIndex        =   15
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Numero "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   " Interno"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "nTipoDoc"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   1155
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Nueva serie"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   60
         TabIndex        =   11
         Top             =   3120
         Width           =   1575
      End
   End
   Begin VB.TextBox txtLimiteGrilla 
      Height          =   330
      Left            =   10140
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "100000000"
      Top             =   90
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.OptionButton optSoloStock 
      Caption         =   "Solo Stock"
      Height          =   360
      Left            =   3180
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1440
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar Desde"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   720
      TabIndex        =   34
      Top             =   0
      Width           =   1395
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Documento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   33
      Top             =   360
      Width           =   1155
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   32
      Top             =   780
      Width           =   1155
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Numero Documento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3780
      TabIndex        =   31
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Filtra solo datos cargados  El filtro es acumulativo"
      Height          =   435
      Left            =   120
      TabIndex        =   30
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Serie"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   720
      TabIndex        =   29
      Top             =   1080
      Width           =   1155
   End
   Begin VB.Label lblNoViNada 
      Caption         =   "No encontre datos para:"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   4185
      TabIndex        =   28
      Top             =   1545
      Visible         =   0   'False
      Width           =   7185
   End
   Begin VB.Label lblGrilllaRestringida 
      Caption         =   "Datos limitados a X lineas"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   4170
      TabIndex        =   27
      Top             =   1485
      Visible         =   0   'False
      Width           =   3915
   End
   Begin VB.Label Label7 
      Caption         =   "Limite visualizacion"
      Height          =   300
      Left            =   8685
      TabIndex        =   26
      Top             =   120
      Visible         =   0   'False
      Width           =   1425
   End
End
Attribute VB_Name = "frmModSerie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'Private g As LiGrilla
'Private Const LIMITE_GRILLA = 90000

Private Sub cmdBuscar_Click()
    On Error GoTo fin
'    grilla.Clear
    Dim s As String
    Dim rs As New ADODB.Recordset
    
    lblNoViNada.Visible = False
'    s = "SELECT s.producto  as [ Cod Producto                       ] " _
'        & " , p.descripcion as [ Descripcion Producto                                          ] " _
'        & " , s.serie  as [ Serie                  ] " _
'        & " , c.descripcion AS [ TipoDoc ] " _
'        & " , s.nrocomprobante AS [ Numero ] " _
'        & " FROM Series AS s INNER JOIN TipoComprobantesGrales AS c ON s.comprobante = c.codigo INNER JOIN producto as p on s.producto = p.codigo " _
'        & " WHERE s.activo = 1 and s.fecha_alta >= " & uFecha.ConvertFecha

'    s = "select s.producto  as [ Cod Producto                       ] , p.descripcion as [ Descripcion Producto                                          ]  " _
        & ", s.serie  as [ Serie                  ] , c.descripcion AS [ TipoDoc ] , s.nrocomprobante AS [ Numero ]   " _
        & " from series s INNER JOIN TipoComprobantesGrales AS c ON s.comprobante = c.codigo INNER JOIN producto as p on s.producto = p.codigo where s.activo = 1 " _
        & " and (select count(ss.serie) from series ss where ss.serie=s.serie and ss.producto=s.producto and ss.activo=1) not in (2,4,6,8,10,12,14,16) " _
        & " and not (((select count(ss.serie) from series ss where ss.serie=s.serie and ss.producto=s.producto and ss.activo=1) in (1)) and essalida=1) " _
        & " and not (((select count(ss.serie) from series ss where ss.serie=s.serie and ss.producto=s.producto and ss.activo=1) in (3,5,7)) and " _
        & " ((select count(ss.serie) from series ss where ss.serie=s.serie and ss.producto=s.producto and ss.activo=1 and essalida=1)>(select count(ss.serie) " _
        & " from series ss where ss.serie=s.serie and ss.producto=s.producto and ss.activo=1 and essalida=0))) and  s.fecha_alta >= " & uFecha.ConvertFecha
        
        s = "select s.producto  as [ Cod Producto                       ] , p.descripcion as [ Descripcion Producto                                          ]  " _
        & ", s.serie  as [ Serie                  ] , c.descripcion AS [ TipoDoc ] , s.nrocomprobante AS [ Numero ]   " _
        & " from series s INNER JOIN TipoComprobantesGrales AS c ON s.comprobante = c.codigo INNER JOIN producto as p on s.producto = p.codigo where s.activo = 1 " _
        & " and not (((select count(ss.serie) from series ss where ss.serie=s.serie and ss.producto=s.producto and ss.activo=1) in (1)) and essalida=1) " _
        & " and not (((select count(ss.serie) from series ss where ss.serie=s.serie and ss.producto=s.producto and ss.activo=1) in (3,5,7)) and " _
        & " ((select count(ss.serie) from series ss where ss.serie=s.serie and ss.producto=s.producto and ss.activo=1 and essalida=1)>(select count(ss.serie) " _
        & " from series ss where ss.serie=s.serie and ss.producto=s.producto and ss.activo=1 and essalida=0))) and  s.fecha_alta >= " & uFecha.ConvertFecha


'        & " , s.comprobante as [tdint] " _
'        & " , s.codigo as [co] " _
'        & " , s.EsSalida as [Salida] "
    If txtDocumento > "" Then s = s & " and c.descripcion = '" & Trim$(txtDocumento) & "' "
    If uProducto.codigo > "" Then s = s & " and s.producto = '" & uProducto.codigo & "' "
    If txtNumero > "" Then s = s & " and  s.nrocomprobante  = " & Trim$(txtNumero) & " "
    If txtserie > "" Then s = s & " and s.serie = '" & Trim$(txtserie) & "' "
    's = s & " ORDER BY s.producto, s.nrocomprobante DESC"
    s = s & " ORDER BY s.producto, s.serie DESC"
    
    rs.Open s, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    MsgBox "" & rs.RecordCount
    
    relojito
    lstDetalle.clear
    grilla.rows = 1
'    lblNoViNada.caption = "No encontre datos para: " & txtDocumento & " " & txtNumero & "    " & uProducto.codigo & " " & txtSerie
'    lblGrilllaRestringida.caption = "Grilla limitada a " & LimiteVisualizacion & " elementos"
    
    lblNoViNada.Visible = Not LlenarGrilla(grilla, s, False)
'    lblGrilllaRestringida.Visible = (grilla.rows > LimiteVisualizacion)
    
    Dim i As Long
    For i = 4 To grilla.cols - 1
        grilla.ColWidth(i) = 0
    Next i
    optSoloStock.Value = False
    
    uMenu.BuscarOK
    
fin:
'    optSoloStock.Value = True
    relojito False
End Sub

Private Sub cmdBuscarDoc_Click()
    Dim resu
    lblNoViNada.Visible = False
'    resu = frmBuscar.MostrarSql("select distinct comprobante from series where activo = 1 and fecha_alta > " & uFecha.ConvertFecha)
    resu = frmBuscar.MostrarSql("select descripcion as [ Tipo Doc ] from TipoComprobantesGrales where activo = 1 ")
    If resu > "" Then txtDocumento = resu
End Sub

Private Sub cmdsalir_Click()
    If MsgBox("¿Cerrar formulario?", vbYesNo + vbQuestion, "Confirmacion") = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmdSerieEnStock_Click()
    Dim resu As String
    
    '''resu = Buscar_SeriesEnStock(txtN_Prod)
    resu = SerieStockRepetida(txtN_Prod)
    
    
    If resu = "" Then Exit Sub
    txtN_NuevaSerie = resu
End Sub

Private Sub Command1_Click()
Dim rs As New ADODB.Recordset
Dim Consulta As String

    If grilla.rows > 1 Then
'        CrearConsulta False

        Consulta = "select s.producto  as [ Cod Producto                       ] , p.descripcion as [ Descripcion Producto                                          ]  " _
        & ", s.serie  as [ Serie                  ] , c.descripcion AS [ TipoDoc ] , s.nrocomprobante AS [ Numero ]   " _
        & " from series s INNER JOIN TipoComprobantesGrales AS c ON s.comprobante = c.codigo INNER JOIN producto as p on s.producto = p.codigo where s.activo = 1 " _
        & " and (select count(ss.serie) from series ss where ss.serie=s.serie and ss.producto=s.producto and ss.activo=1) not in (2,4,6,8,10,12,14,16) " _
        & " and not (((select count(ss.serie) from series ss where ss.serie=s.serie and ss.producto=s.producto and ss.activo=1) in (1)) and essalida=1) " _
        & " and not (((select count(ss.serie) from series ss where ss.serie=s.serie and ss.producto=s.producto and ss.activo=1) in (3,5,7)) and " _
        & " ((select count(ss.serie) from series ss where ss.serie=s.serie and ss.producto=s.producto and ss.activo=1 and essalida=1)>(select count(ss.serie) " _
        & " from series ss where ss.serie=s.serie and ss.producto=s.producto and ss.activo=1 and essalida=0))) and  s.fecha_alta >= " & uFecha.ConvertFecha

        
        rs.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        VinculoXl "C:\ListProductos.xls", "Listado de productos", , , rs
        rs.Close
        Set rs = Nothing
    Else
        MsgBox "Debe tirar el listado primero.", vbOKOnly, "Atencion"
    End If
End Sub

'Private Sub cmdSoloStock_Click()
'    Dim i As Long, pr As String, sE As String
'    Dim cc As Long
'    For i = 1 To grilla.rows - 1
'        pr = grilla.TextMatrix(i, 0)
'        sE = grilla.TextMatrix(i, 2)
'        If Not SerieEnStock(sE, pr) Then
'            grilla.RowHidden(i) = True
'            cc = cc + 1
'
'        End If
'    Next i
'     Debug.Print cc
'End Sub

Private Sub Form_Load()
    Dim sqldesc As String, sqlbuscar As String
    sqldesc = "select descripcion from producto where codigo = '###' "
    sqlbuscar = "select codigo as [ Codigo                 ], descripcion as [ Descripcion                                                 ] from producto where activo = 1 order by codigo "
    uProducto.ini sqldesc, sqlbuscar, True
    uFecha.dtfecha #1/1/2000#
'    uMenu.init False, False, False, False, False
    ucXls1.ini grilla, "C:\ListProductos", "Listado de productos"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub

Private Sub Form_Resize()
    Anclar cmdsalir, Me, anclarAbajo + anclarDerecha
    Anclar lstDetalle, Me, anclarArriba + anclarDerecha ' + anclarAbajo
    Anclar grilla, Me, anclarIzquierda + anclarDerecha + anclarArriba + anclarAbajo
    Anclar fraCambioSerie, Me, anclarArriba + anclarDerecha
End Sub


Private Sub optSoloStock_Click()
    On Error GoTo fin
    Dim i As Long, pr As String, sE As String
    Dim cc As Long
cc = 0
    If optSoloStock.Value Then
        'relojito True
        For i = 1 To grilla.rows - 1
            pr = grilla.TextMatrix(i, 0)
            sE = grilla.TextMatrix(i, 2)
            If Not SerieEnStock(sE, pr) Then
                grilla.RowHidden(i) = True
            Else
                cc = cc + 1
            End If
        Next i
'        Debug.Print cc
        MsgBox "" & cc
    End If
fin:
    relojito False
End Sub

Private Sub txtDocumento_LostFocus()
    lblNoViNada.Visible = False
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9" Then KeyAscii = 0
End Sub

Private Sub VerDetalle(ntdoc, nDoc, tdoc, Optional essalida As String)
    With lstDetalle
    Dim tempo
    .clear
    .AddItem IIf(Left(LCase(essalida), 1) = "f", "<---- INGRESO ---- ", "----- EGRESO ----> ")
    .AddItem " "
    Select Case ntdoc
    Case 1 To 4 'factura venta
        tempo = obtenerDeSQL("select f.fecha, f.cliente, c.descripcion from FacturaVenta as f inner join clientes  as c on c.codigo = f.cliente where f.tipodoc = '" & tdoc & "' and NroFactura = " & nDoc)
        If IsEmpty(tempo) Then Exit Sub
        .AddItem "Doc =     " & tdoc
        .AddItem "Nro =     " & nDoc
        .AddItem "Fecha =   " & tempo(0)
        .AddItem "Cliente = " & tempo(1) & "-" & tempo(2)
    Case 5      'rv
        tempo = obtenerDeSQL("select r.fecha, r.cliente, c.descripcion from RemitoVenta as r inner join clientes as c on r.cliente = c.codigo where numero = " & nDoc)
        If IsEmpty(tempo) Then Exit Sub
        .AddItem "Doc =     " & "Remito Venta"
        .AddItem "Nro =     " & nDoc
        .AddItem "Fecha =   " & tempo(0)
        .AddItem "Cliente = " & tempo(1) & "-" & tempo(2)
    Case 6      'rc
        tempo = obtenerDeSQL("select r.fecha, r.NroRemito, r.proveedor, p.descripcion from RemitoCompra as r inner join prov as p on r.proveedor = p.codigo where r.codigo = " & nDoc)
        If IsEmpty(tempo) Then Exit Sub
        .AddItem "Doc =        " & "Remito Prov"
        .AddItem "Nro (int) =  " & nDoc
        .AddItem "Nro (Prov) = " & tempo(1)
        .AddItem "Fecha =      " & tempo(0)
        .AddItem "Cliente =    " & tempo(2) & "-" & tempo(3)
    Case 7, 8   'dif
        tempo = obtenerDeSQL("select r.fecha, r.NroPedido, r.Concepto from RemitoDiferenciaStock as r where r.MovimientoInterno = " & nDoc) 'where r.comprobante= " & nDoc)
        If IsEmpty(tempo) Then Exit Sub
        .AddItem "Doc =        " & "Dif Stock"
        .AddItem "Nro (int) =  " & nDoc
        .AddItem "Fecha =      " & tempo(0)
        .AddItem "Pedido =     " & tempo(1)
        .AddItem "Concepto =   " & sSinNull(obtenerDeSQL("select descripcion from conceptos where codigo = '" & tempo(2) & "'"))
        .AddItem " "
'        .AddItem essalida
    
    End Select
    End With
End Sub

Private Sub txtNumero_LostFocus()
    lblNoViNada.Visible = False
End Sub
Private Sub txtSerie_LostFocus()
    lblNoViNada.Visible = False
End Sub

Private Sub uMenu_AceptarModi()
    If MODO_ON_ERROR_ABM_ON Then On Error GoTo ufamodi
    Dim ntdoc, codi
    codi = s2n(txtN_Codigo)
    
    If codi = 0 Then Exit Sub
    If s2n(obtenerDeSQL("select codigo from series where codigo = '" & codi & "' ")) = 0 Then
        Exit Sub
    End If
    If Trim(txtN_NuevaSerie) = "" Then
        che "falta ingresar nueva serie"
        Exit Sub
    End If
    ntdoc = s2n(txtn_ntdoc)
    If ntdoc <> 6 Then
        If Not SerieEnStock(Trim(txtN_NuevaSerie), Trim(txtN_Prod)) Then
            If Not confirma("Serie NO figura en stock, corfirma su salida?") Then
                Exit Sub
            End If
        End If
    End If
    
    Dim salta, sbaja
    salta = "INSERT INTO Series (producto, serie, comprobante, nrocomprobante, sucursal, concepto, observaciones, consignacion, Fecha, EsSalida, fecha_alta, usuario_alta, activo ) " _
        & " SELECT s.producto, '" & Trim(txtN_NuevaSerie) & "', s.comprobante, s.nrocomprobante, s.sucursal, s.concepto, s.observaciones, s.consignacion, s.Fecha, s.EsSalida, " & ssFecha(Date) & ", " & x2s(UsuarioActual()) & ",  1 " _
        & " FROM Series AS s " _
        & " WHERE s.codigo = '" & codi & "' "
    sbaja = " update series set activo = 0, fecha_baja = " & ssFecha(Date) & ", usuario_baja = " & UsuarioActual() & "  where codigo = " & codi
    DE_BeginTrans
    'DataEnvironment1.dbo_SERIE "B", codi, "", "", 0, 0, 0, 0, "", 0, 0, 0, UsuarioActual(), Date
    'DataEnvironment1.dbo_SERIE "A", 0, txtN_Prod, txtN_NuevaSerie, ntdoc, txtN_NroDoc, 0, 0, "Mod Serie", 0, Date, UsuarioActual(), 0, 0
    DataEnvironment1.Sistema.Execute salta
    DataEnvironment1.Sistema.Execute sbaja
    DE_CommitTrans

    che "Serie cambiada"
    uMenu.AceptarOk
    
    txtNumero = txtN_NroDoc
    txtDocumento = txtN_TipoDoc
    uProducto.codigo = txtN_Prod
    cmdBuscar_Click
fin:
    Exit Sub
ufamodi:
    DE_RollbackTrans
    ufa "Fallo grabacion de modificacion", "modiserie"
    Resume fin
End Sub
Private Sub uMenu_BorrarControles()
'
End Sub
Private Sub uMenu_HabilitarEdicion(sino As Boolean)
    fraCambioSerie.Visible = sino
End Sub
Private Sub uMenu_SALIR()
    Unload Me
End Sub

'Private Function LimiteVisualizacion() As Long
'    LimiteVisualizacion = Abs(s2n(txtLimiteGrilla, 0))
'End Function

Private Sub uProducto_cambio(codigo As Variant)
    lblNoViNada.Visible = False
End Sub



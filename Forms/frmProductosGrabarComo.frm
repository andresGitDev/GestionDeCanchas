VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmProductosGrabarComo 
   Caption         =   "Crear Producto y Formula"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9450
   Icon            =   "frmProductosGrabarComo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   9450
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex7LCtl.VSFlexGrid gri 
      Height          =   3015
      Left            =   135
      TabIndex        =   11
      Top             =   3240
      Width           =   9165
      _cx             =   16166
      _cy             =   5318
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
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   825
      Left            =   8100
      Picture         =   "frmProductosGrabarComo.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton cmdCopiar 
      Caption         =   "&Copiar"
      Height          =   825
      Left            =   7485
      Picture         =   "frmProductosGrabarComo.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox txtCodigoNuevo 
      Height          =   330
      Left            =   2220
      TabIndex        =   8
      Top             =   2850
      Width           =   1950
   End
   Begin VB.TextBox txtCodigo 
      Height          =   330
      Left            =   2190
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1455
      Width           =   1950
   End
   Begin Gestion.ucCoDe uProd 
      Height          =   330
      Left            =   2175
      TabIndex        =   0
      Top             =   135
      Width           =   5040
      _ExtentX        =   12541
      _ExtentY        =   582
      CodigoWidth     =   1000
   End
   Begin Gestion.ucCoDe ugrupo 
      Height          =   330
      Left            =   2205
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   585
      Width           =   4980
      _ExtentX        =   8784
      _ExtentY        =   582
      CodigoWidth     =   1000
   End
   Begin Gestion.ucCoDe uSubgrupo 
      Height          =   330
      Left            =   2205
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1005
      Width           =   4980
      _ExtentX        =   8784
      _ExtentY        =   582
      CodigoWidth     =   1000
   End
   Begin Gestion.ucCoDe uGrupoNuevo 
      Height          =   330
      Left            =   2235
      TabIndex        =   6
      Top             =   1935
      Width           =   4980
      _ExtentX        =   8784
      _ExtentY        =   582
      CodigoWidth     =   1000
   End
   Begin Gestion.ucCoDe uSubgrupoNuevo 
      Height          =   330
      Left            =   2235
      TabIndex        =   7
      Top             =   2415
      Width           =   4980
      _ExtentX        =   8784
      _ExtentY        =   582
      CodigoWidth     =   1000
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo Nuevo"
      Height          =   330
      Index           =   6
      Left            =   1140
      TabIndex        =   16
      Top             =   2865
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "SubGrupo Nuevo"
      Height          =   330
      Index           =   5
      Left            =   915
      TabIndex        =   15
      Top             =   2445
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo Compuesto Original"
      Height          =   255
      Index           =   4
      Left            =   195
      TabIndex        =   14
      Top             =   1500
      Width           =   2025
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo SubGrupo Original"
      Height          =   255
      Index           =   3
      Left            =   270
      TabIndex        =   13
      Top             =   1080
      Width           =   2085
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo Grupo Original"
      Height          =   255
      Index           =   2
      Left            =   540
      TabIndex        =   12
      Top             =   630
      Width           =   2040
   End
   Begin VB.Label Label1 
      Caption         =   "Grupo Nuevo"
      Height          =   330
      Index           =   1
      Left            =   1185
      TabIndex        =   2
      Top             =   1965
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo Producto Original"
      Height          =   255
      Index           =   0
      Left            =   345
      TabIndex        =   1
      Top             =   195
      Width           =   2235
   End
End
Attribute VB_Name = "frmProductosGrabarComo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' *** OJO ***  Este form asume CodProdCompuesto,
'               No vale la pena adaprtarlo al parametro,
'               si no es compuesto es MUCHO mas facil hacer un form distinto (o dentro de producto)

Private Sub cmdCopiar_Click()

    Dim tempo
    
    'verifico
    ' si hay algo distinto
    If uProd.codigo = "" Then Exit Sub
    
    If uProd.codigo = conu() Then Exit Sub
    

    tempo = obtenerDeSQL("select idproducto, activo from producto where codigo = '" & conu & "'  ")
    
    
    If IsEmpty(tempo) Then
        'no existe
        copiar
        
    ElseIf tempo(1) = 0 Then
        'existe, pero borrado
        If confirma("Producto " & conu & " figura borrado, desea activarlo") Then
            DataEnvironment1.Sistema.Execute "update producto set activo = 1 where idproducto = " & tempo(0)
        End If
        
    Else
        'existe, gil.
        che "Producto ya existe"
        
    End If
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub

Private Sub Form_Load()
    uProd.ini "Select descripcion from producto where codigo = '###' and activo = 1 ", "select codigo as [Codigo       ], descripcion as [ Descripcion                         ], formula  from producto where activo = 1 ", True
    
    
    ugrupo.ini "Select descripcion from gruposproducto where codigo = '###' ", "select codigo as grupo, descripcion from gruposproducto", True
    uSubgrupo.ini "Select descripcion from subgruposproducto where codigo = '###' ", "select codigo as subgrupo, descripcion from subgruposproducto", True
    ugrupo.enabled = False
    uSubgrupo.enabled = False
    txtCodigo.enabled = False
    
    uGrupoNuevo.ini "Select descripcion from gruposproducto where codigo = '###' ", "select codigo as grupo, descripcion from gruposproducto", True
    uSubgrupoNuevo.ini "Select descripcion from subgruposproducto where codigo = '###' ", "select codigo as subgrupo, descripcion from subgruposproducto", True
    
    gri.cols = 2
    gri.FixedCols = 0
End Sub


Private Function extraeGrupo(deque)
    extraeGrupo = Left(deque, 3)
End Function
Private Function extraeSubGrupo(deque)
    extraeSubGrupo = Mid(deque, 4, 3)
End Function

Private Sub uProd_cambio(codigo As Variant)
    ugrupo.codigo = extraeGrupo(uProd.codigo)
    uGrupoNuevo.codigo = extraeGrupo(uProd.codigo)
    txtCodigo = Mid(uProd.codigo, 7)
    
    uSubgrupo.codigo = extraeSubGrupo(uProd.codigo)
    uSubgrupoNuevo.codigo = extraeSubGrupo(uProd.codigo)
    txtCodigoNuevo = Mid(uProd.codigo, 7)
    
    LlenarGrilla gri, "select componente, cantidad from formulas where activo = 1 and codigo = '" & uProd.codigo & "' ", False
    grillaWidth gri, Array(1500, 800)
End Sub
Public Function s3(que)
    s3 = Left(que & "     ", 3)
End Function
Private Function conu()
    conu = s3(uGrupoNuevo.codigo) & s3(uSubgrupoNuevo.codigo) & Trim(txtCodigoNuevo)
End Function
Private Sub copiar()
    Dim s
    Dim rs As New ADODB.Recordset
    
    DE_BeginTrans
     With rs
        'producto
        s = "select * from producto where codigo = '" & uProd.codigo & "'  "
        .Open s, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        DataEnvironment1.dbo_PRODUCTOS "A", conu, uGrupoNuevo.codigo, uSubgrupoNuevo.codigo, !DESCRIPCION, _
                !UMedida, !costobase, !CALCSINCOSTO, !COSTOPROM, !precio, !precio2, !precio3, !precio4, _
                !TIPOPROD, nSinNull(!COSTOPROV), !STOCKMIN, !pedmin, !CODIGOBARRA, !Serie, !Iva, 0, !TIEMPOELABORACION, _
                !grafico, !letra, !CANTCONTROL, !observaciones, !PUEDOFAC, nSinNull(!moneda), !formula, sSinNull(!Alias), Date, UsuarioActual(), !CUENTA, !ManejaStock, !ESTADO, 0
        .Close
        
        'formula
        s = " insert  into formulas (codigo, Componente, cantidad, fecha_alta, usuario_alta, activo ) " & _
            " select '" & conu() & "' as conu, Componente, cantidad, " & ssFecha(Date) & ", " & UsuarioActual() & ", 1 " & _
            " from formulas where codigo = '" & uProd.codigo & "' "
    
        DataEnvironment1.Sistema.Execute s
     End With
    DE_CommitTrans
    MsgBox "Copiado.", vbInformation
    Set rs = Nothing
End Sub

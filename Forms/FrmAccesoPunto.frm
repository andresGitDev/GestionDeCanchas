VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form FrmAccesoPunto 
   Caption         =   "Puntos de venta por defecto"
   ClientHeight    =   6435
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin VSFlex7Ctl.VSFlexGrid GRILLA 
      Height          =   5535
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   6615
      _cx             =   11668
      _cy             =   9763
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
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar cambios"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdRefrescar 
      Caption         =   "Refrescar"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "FrmAccesoPunto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGuardar_Click()
    Dim i As Long
    Dim str As String
    Dim punto As String
    Dim sTipo As String
    
    If GRILLA.rows < 2 Then Exit Sub
    i = 1
    While i < GRILLA.rows
        sTipo = Tipo(i) '(cmbPunto.ListIndex)
        If sTipo <> "" Then
            punto = nSinNull(obtenerDeSQL("select puntoventa from documentoscae where tipopunto=" & ssTexto(sTipo) & " and tipo in('FAA','FAB','NCA','NCB')"))
        
            If s2n(obtenerDeSQL("select id from puntodefecto where usuario=" & GRILLA.TextMatrix(i, 0))) > 0 Then
                'modifico
                str = "update puntodefecto set puntodefecto=" & ssTexto(punto) & " where usuario=" & GRILLA.TextMatrix(i, 0)
                DataEnvironment1.Sistema.Execute str
            Else
                'agrego
                str = "insert into puntodefecto(usuario,puntodefecto) values(" & GRILLA.TextMatrix(i, 0) & "," & ssTexto(punto) & ")"
                DataEnvironment1.Sistema.Execute str
            End If
        End If
        i = i + 1
    Wend
    MsgBox "Se ha guardado con exito.", , "ATENCION"
    
End Sub

Private Function Tipo(item As Long)
    If Trim(GRILLA.TextMatrix(item, 2)) = "PRE-IMPRESA" Then
        Tipo = "PI"
    ElseIf Trim(GRILLA.TextMatrix(item, 2)) = "ONLINE" Then
        Tipo = "OL"
    ElseIf Trim(GRILLA.TextMatrix(item, 2)) = "WEBSERVICE" Then
        Tipo = "WS"
    ElseIf Trim(GRILLA.TextMatrix(item, 2)) = "WEBSERVICE 2" Then
        Tipo = "WS2"
    ElseIf Trim(GRILLA.TextMatrix(item, 2)) = "" Then
        Tipo = ""
    End If
End Function

Private Sub Form_Load()
    ini
    llenar
End Sub

Private Sub ini()
    Dim Lista As String
    
    'GRILLA.rows = 0
    GRILLA.rows = 1
    GRILLA.cols = 0
    GRILLA.cols = 3
    
    GRILLA.TextMatrix(0, 0) = "Codigo"
    GRILLA.TextMatrix(0, 1) = "Usuario"
    GRILLA.TextMatrix(0, 2) = "P.Venta"
    
    Lista = "PRE-IMPRESA|ONLINE|WEBSERVICE|WEBSERVICE 2"
    GRILLA.ColComboList(2) = Lista
    
    GRILLA.ColHidden(0) = True
    
    GRILLA.ColWidth(1) = 3000
    GRILLA.ColWidth(2) = 2000
    
    GRILLA.Editable = True
End Sub

Private Sub llenar()
    Dim str As String
    Dim rs As New ADODB.Recordset
    Dim punto As String
    Dim sTipo As String
    Dim descri As String
    
    str = "select * from usuarios where activo=1 and codigo>1"
    rs.Open str, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    While Not rs.EOF
        
        punto = sSinNull(obtenerDeSQL("select puntodefecto from puntodefecto where usuario=" & rs!codigo))
        If punto = "" Then
            GRILLA.AddItem rs!codigo & Chr(9) & rs!usuario
        Else
            sTipo = obtenerDeSQL("select tipopunto from documentoscae where puntoventa=" & ssTexto(punto) & " and tipo  in('FAA','FAB','NCA','NCB')")
            If sTipo = "PI" Then
                descri = "PRE-IMPRESA"
            ElseIf sTipo = "OL" Then
                descri = "ONLINE"
            ElseIf sTipo = "WS" Then
                descri = "WEBSERVICE"
            ElseIf sTipo = "WS2" Then
                descri = "WEBSERVICE 2"
            Else
                descri = ""
            End If
            GRILLA.AddItem rs!codigo & Chr(9) & rs!usuario & Chr(9) & descri
        End If
        
        
        rs.MoveNext
    Wend
End Sub


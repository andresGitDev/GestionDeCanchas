VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.UserControl ucSeries 
   ClientHeight    =   3075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8625
   ScaleHeight     =   3075
   ScaleWidth      =   8625
   Begin VB.CommandButton cmdLlenaSerie 
      Caption         =   "Llenar Serie"
      Height          =   315
      Left            =   3060
      TabIndex        =   2
      ToolTipText     =   "Seleccione filas a llenar"
      Top             =   120
      Width           =   1575
   End
   Begin VB.CheckBox chkSinSeries 
      Alignment       =   1  'Right Justify
      Caption         =   "Sin Series"
      Height          =   315
      Left            =   7500
      TabIndex        =   1
      Top             =   120
      Width           =   1035
   End
   Begin VSFlex7LCtl.VSFlexGrid grillaSeries 
      Align           =   2  'Align Bottom
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   540
      Width           =   8625
      _cx             =   15214
      _cy             =   4471
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
   Begin VB.Label Label10 
      Caption         =   "Puede hacer 'Doble Clic' en el campo  Nro.Serie"
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   2235
   End
   Begin VB.Label lblErrorSeries 
      Caption         =   "--------------"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   4800
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   2475
   End
End
Attribute VB_Name = "ucSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private g3 As LiGrilla
Private mGrillaOrigen As LiGrilla
Private mColProd  As Long
Private mColCant  As Long
Private mColCons  As Long

Private g3ITEM As Long
Private g3PROD As Long
Private g3NSER As Long
Private g3HIDD As Long
Private g3CONS As Long
Private g3ALTA As Long
'
Private Const CTE_SERIE_AGREGAR = "Registrar"

Private mModoSalida      As Boolean
Private mEnabled         As Boolean
Private mPropio          As Boolean
Private mConConsignacion As Boolean

Public Enum ucSerieColNombre
    gsItem
    gsProdu
    gsNser
    gsHidd
    gsCons
    gsAlta
End Enum
'

'***********************************************
Public Property Get modoSalida() As Boolean
    modoSalida = mModoSalida
End Property
Public Property Let modoSalida(cual As Boolean)
    mModoSalida = cual
End Property
Public Property Get Propio() As Boolean
    Propio = mPropio
End Property
Public Property Let Propio(sino As Boolean)
    mPropio = sino
End Property

Public Property Get enabled() As Boolean
    enabled = mEnabled
End Property
Public Property Let enabled(sino As Boolean)
    mEnabled = sino
    cmdLlenaSerie.enabled = sino
    chkSinSeries.enabled = sino
    grillaSeries.Editable = IIf(sino, flexEDKbdMouse, flexEDNone)
End Property
Public Property Get rows() As Long
    rows = g3.rows ' PrimerVacio(g3PROD)
End Property

Public Function cell(Row As Long, colNombre As ucSerieColNombre) As String
    Dim i As Long
    i = Row
    If i = 0 Or i > grillaSeries.rows - 1 Then Exit Function
    cell = g3.tx(Row, colNombre)
End Function

'---------------------------------------------------------------------
Public Sub ini(g As LiGrilla, colProd As Long, colCant As Long, Propio As Boolean, Optional colCons)
    mConConsignacion = IIf(IsMissing(colCons), False, True)
    Set mGrillaOrigen = g
    mColProd = colProd
    mColCant = colCant
    mColCons = s2n(colCons)
    mPropio = Propio
End Sub
Public Function FaltaSeries() As Boolean
    LlenoGrillaSeries
    
    FaltaSeries = FaltaSeriesFuncion(mModoSalida)
End Function

Private Function FaltaSeriesFuncion(checkExistencia As Boolean) As Boolean
    Dim r As Long, i As Long, ns As String
    Dim seri As String, prod As String
    
    FaltaSeriesFuncion = False
    lblErrorSeries.Visible = False
    
    If chkSinSeries.Value = vbChecked Then Exit Function
    
'    LlenoGrillaSeries
    r = g3.rows
    
    'vacio
    If r > 1 And g3.Buscar(g3NSER, "") > 0 Then
        
'        tabRemito.Tab = 2
        grillaSeries.SetFocus
        grillaSeries.Select g3.PrimerVacio(g3NSER), g3NSER
        
        FaltaSeriesFuncion = True
        Exit Function
    End If
    
    If checkExistencia Then
        'existe serie ?
        For i = 1 To r - 1
            seri = g3.tx(i, g3NSER)
            prod = g3.tx(i, g3PROD)
            If Not SerieEnStock(seri, prod) Then
                If g3.tx(i, g3ALTA) <> CTE_SERIE_AGREGAR Then
                    che "No figura en stock " & vbCrLf & "Producto: " & prod & vbCrLf & "Serie: " & seri
                    
    '                tabRemito.Tab = 2
                    grillaSeries.SetFocus
                    grillaSeries.Select i, g3NSER
                    If confirma("Desea registrarlo ahora") Then
                        g3.tx i, g3ALTA, CTE_SERIE_AGREGAR
                    Else
                        FaltaSeriesFuncion = True
                        Exit Function
                    End If
                End If
            End If
        Next i
    End If
     
    If r > 1 Then
        For i = 1 To r - 2
            ns = g3.tx(i, g3NSER)
            If ns <> "" And g3.Buscar(g3NSER, ns, i + 1) > 0 Then
'                tabRemito.Tab = 2
                grillaSeries.SetFocus
                grillaSeries.Select i, g3NSER, g3.Buscar(g3NSER, ns, i + 1), g3NSER
                'grillaSeries.Select g3.Buscar(g3NSER, ns, i + 1), g3NSER
                FaltaSeriesFuncion = True
                Exit Function
            End If
        Next i
    End If
End Function

Private Sub LlenoGrillaSeries()
    Dim i As Long, j As Long, prod As String, cant As Long, cons

        'borrosinserie
        i = 1
        While i < g3.rows
            If g3.tx(i, g3NSER) = "" Then
                g3.delRow i
                i = i - 1
            End If
            i = i + 1
        Wend

        'borro marcas
        For i = 1 To g3.rows - 1
            grillaSeries.TextMatrix(i, g3HIDD) = ""
        Next i

        'marco o agrego en grilla series
        With mGrillaOrigen
            For i = 1 To .rows - 1
                prod = Trim(.tx(i, mColProd))
                cant = s2n(.tx(i, mColCant))
                cons = IIf(mConConsignacion, s2n(.tx(i, mColCons)), 0)
                
                If ProductoConSerie(prod, mPropio) Then
                    For j = 1 To cant
                        If marcoG3(prod, cons) Then
                            cons = cons - 1
                        End If
                    Next j
                End If
            Next i
        End With
        'borro no marcadas
        i = 1
        While i < g3.rows
            If g3.tx(i, g3HIDD) = "" Then
                g3.delRow i
                i = i - 1
            End If
            i = i + 1
        Wend
End Sub

'***********************************************

Private Sub cmdLlenaSerie_Click()
    Dim i As Long, Row As Long
    Dim s As String, n As Long       ' current cell
    Dim ss As String, ns As Long     ' the longer one
    ss = ""
    n = 0
    With grillaSeries ' uso propiedades que no exporte
        For i = 0 To .SelectedRows - 1
            Row = .SelectedRow(i)
            s = .TextMatrix(Row, g3NSER)
            
            If Len(s) > Len(ss) Then
                ss = s
                ns = Len(ss)
            End If
        Next i
        For i = 0 To .SelectedRows - 1
            Row = .SelectedRow(i)
            s = .TextMatrix(Row, g3NSER)
            n = Len(s)
            If n < ns And n > 0 Then .TextMatrix(Row, g3NSER) = Left(ss, ns - n) & s
        Next i
    End With
End Sub

Private Sub UserControl_GotFocus()
    If mEnabled Then LlenoGrillaSeries
End Sub

Private Sub UserControl_Initialize()
    On Error Resume Next
    Set g3 = New LiGrilla
    g3.init grillaSeries
    
    'nombres g3 no se usan, ahora esta el enum para eso
    'addcol DEBE estar en el orden del enum
    g3ITEM = g3.AddCol("  -  ", "A")
    g3PROD = g3.AddCol(" Producto                      ")
    g3NSER = g3.AddCol(" Numero de Serie            ", "S") ' editable
    g3CONS = g3.AddCol("Consignacion", "K")
    g3HIDD = g3.AddCol("", "H")
    g3ALTA = g3.AddCol("                ")
    'addcol DEBE estar en el orden del enum
    '--------------------------------------------------

    g3.Borrar
    grillaSeries.SelectionMode = flexSelectionListBox
End Sub

Private Sub UserControl_Resize()
    grillaSeries.Height = UserControl.Height - 500
End Sub

Private Sub UserControl_Terminate()
    Set g3 = Nothing
End Sub
'
Public Sub MetoGrillaSeries(prod As String, ByVal cant As Long, ByVal consig As Long)
    On Error GoTo ERR_FIN
    Dim i As Long, r As Long

    If ProductoConSerie(prod, mPropio) Then
        'asserts saludables
'        If g3.rows > 100 Then
'            ufa " Demasiados items para num de serie ", "Remito Venta" ', Err
'            Exit Sub
'        End If
        If cant < 1 Then
            ufa "", "Cantidad para num serie < 1 - ucseries MetoGrillaSeries()" ', Err
            Exit Sub
        End If

        For i = 1 To cant
            r = g3.addRow()
            grillaSeries.TextMatrix(r, g3PROD) = prod
            If consig > 0 Then
                grillaSeries.TextMatrix(r, g3CONS) = flexChecked
                consig = consig - 1
            End If
        Next i
    End If

    GoTo fin
ERR_FIN:
    ufa "err en series", " ucseries MetoGrillaSeries()" ', Err
fin:
End Sub

Private Sub grillaSeries_dblClick()
    Dim r As Long, prod As String, resu As String

    If Not mModoSalida Then Exit Sub
    
    r = g3.Row
    If r < 1 Then Exit Sub
    prod = VerProductoMio(g3.tx(r, g3PROD), mPropio)

    resu = Buscar_SeriesEnStock(prod)
    If resu > "" Then grillaSeries.TextMatrix(r, g3NSER) = resu
End Sub

Private Function marcoG3(codi, ByVal cons) '
    Dim i As Long
    
    For i = 1 To g3.rows - 1
        If g3.tx(i, g3PROD) = codi And g3.tx(i, g3HIDD) = "" Then
            grillaSeries.TextMatrix(i, g3HIDD) = "X"
            Exit Function
        End If
    Next i
    i = g3.addRow()
    grillaSeries.TextMatrix(i, g3PROD) = codi
    grillaSeries.TextMatrix(i, g3HIDD) = "X"
    grillaSeries.TextMatrix(i, g3CONS) = IIf(cons > 0, "-1", "0")
    
    marcoG3 = (g3.tx(i, g3CONS) = "-1")
End Function





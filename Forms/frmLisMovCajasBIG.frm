VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmLisMovCajasBIG 
   Caption         =   "Movimientos de caja"
   ClientHeight    =   8475
   ClientLeft      =   90
   ClientTop       =   345
   ClientWidth     =   10140
   Icon            =   "frmLisMovCajasBIG.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   10140
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   540
      Left            =   1320
      TabIndex        =   22
      Top             =   960
      Width           =   6195
      Begin VB.OptionButton OptFecha 
         Caption         =   "Ordenar por fecha"
         Height          =   195
         Left            =   60
         TabIndex        =   25
         Top             =   180
         Value           =   -1  'True
         Width           =   1665
      End
      Begin VB.OptionButton OptImporte 
         Caption         =   "Ordenar por importe"
         Height          =   195
         Left            =   1890
         TabIndex        =   24
         Top             =   195
         Width           =   1785
      End
      Begin VB.OptionButton OptMovimiento 
         Caption         =   "Ordenar por movimiento"
         Height          =   195
         Left            =   3945
         TabIndex        =   23
         Top             =   195
         Width           =   2025
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   885
      Left            =   7485
      Picture         =   "frmLisMovCajasBIG.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7545
      Width           =   870
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   885
      Left            =   9195
      Picture         =   "frmLisMovCajasBIG.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7545
      Width           =   840
   End
   Begin VB.CommandButton cmdMostrar 
      Caption         =   "&Mostrar"
      Height          =   885
      Left            =   6585
      Picture         =   "frmLisMovCajasBIG.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7545
      Width           =   900
   End
   Begin Gestion.ucXls ucXls1 
      Height          =   900
      Left            =   8355
      TabIndex        =   18
      Top             =   7545
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   1588
   End
   Begin Gestion.ucCoDe uCaja 
      Height          =   285
      Left            =   1380
      TabIndex        =   17
      Top             =   465
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   503
      CodigoWidth     =   1455
      CodigoInvalido  =   0
   End
   Begin Gestion.ucFecha uFecha2 
      Height          =   330
      Left            =   2595
      TabIndex        =   16
      Top             =   60
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   582
      FechaInit       =   0
   End
   Begin Gestion.ucFecha uFecha1 
      Height          =   330
      Left            =   1395
      TabIndex        =   15
      Top             =   60
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   582
      FechaInit       =   0
   End
   Begin VB.CheckBox chkOrdernarFecha 
      Caption         =   "Ordenar por fecha"
      Height          =   285
      Left            =   6030
      TabIndex        =   14
      Top             =   480
      Width           =   2670
   End
   Begin VB.Frame fraOpcion 
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   1365
      TabIndex        =   11
      Top             =   1710
      Width           =   8550
      Begin VB.OptionButton optFiltro 
         Caption         =   "Todas"
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   15
         Width           =   1400
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Todas"
         Height          =   375
         Index           =   1
         Left            =   1425
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   15
         Width           =   1400
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Todas"
         Height          =   375
         Index           =   2
         Left            =   2850
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   15
         Width           =   1400
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Todas"
         Height          =   375
         Index           =   3
         Left            =   4275
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   15
         Width           =   1400
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Todas"
         Height          =   375
         Index           =   4
         Left            =   5700
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   15
         Width           =   1400
      End
      Begin VB.OptionButton optFiltro 
         Caption         =   "Todas"
         Height          =   375
         Index           =   5
         Left            =   7140
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   15
         Width           =   1400
      End
   End
   Begin VB.Frame fraTot 
      Height          =   705
      Left            =   30
      TabIndex        =   8
      Top             =   6750
      Width           =   10035
      Begin VB.TextBox txtSaldoAnt 
         Height          =   330
         Left            =   1140
         TabIndex        =   9
         Top             =   210
         Width           =   1905
      End
      Begin VB.Label Label3 
         Caption         =   "Saldo Anterior"
         Height          =   315
         Left            =   75
         TabIndex        =   10
         Top             =   255
         Width           =   1065
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Height          =   4515
      Left            =   30
      TabIndex        =   7
      Top             =   2205
      Width           =   10020
      _cx             =   17674
      _cy             =   7964
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
   Begin VB.Label Label1 
      Caption         =   "Caja (0 todas)"
      Height          =   390
      Index           =   1
      Left            =   255
      TabIndex        =   13
      Top             =   480
      Width           =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "Filtro "
      Height          =   390
      Index           =   2
      Left            =   825
      TabIndex        =   12
      Top             =   1800
      Width           =   465
   End
   Begin VB.Label Label1 
      Caption         =   "Rango de fechas"
      Height          =   390
      Index           =   0
      Left            =   75
      TabIndex        =   6
      Top             =   90
      Width           =   1695
   End
End
Attribute VB_Name = "frmLisMovCajasBIG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mFiltros As Variant
Private mTitulos As Variant
Private mFiltroActual As String
Private mTituloActual As String

Private msCajaTemp As String

Private Enum griCampos
    griCAJA
    griMOVI
    griFECH
    griI_E_
    griTIPO
    griIMPO
    griINTE
    griCONC
End Enum
Private Const griCamposSel = " caja, movimiento, fecha as [Fecha   ],ing_egr as [ I/E ], tipo as [ Valor   ], importe,  interno as [Nro interno], concepto as [Concepto            ] "

Private Sub cmdImprimir_Click()
    Dim STR As String, tipo As String
    Dim total As Double, saldoanterior As Double, diant As Variant
    Dim punto1 As Double, punto2 As Double, punto3 As Double, punto4 As Double, punto5 As Double, punto6 As Double, punto7 As Double, punto8 As Double, punto9 As Double
    Dim rs As New ADODB.Recordset

'generacion tablatemp
    'If msCajaTemp = "" Then
    msCajaTemp = TablaTempCrear(tt_CajasTemp)

        DataEnvironment1.Sistema.Execute "CREATE  INDEX ixi ON " & msCajaTemp & " ([fecha]) ON [PRIMARY] "

    'End If

    'DataEnvironment1.SISTEMA.Execute "delete from " & msCajaTemp 'CajasTemp"
    
    
'PARA UNA CAJA SOLA
        rs.Open "select * from movicaja where fecha " & ssBetween(uFecha1.strFecha, uFecha2.strFecha) & " and caja = " & Val(uCaja.codigo) & " and activo=1 order by fecha", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
'        If Not rs.EOF Then
            While Not rs.EOF
                diant = rs!fecha
                total = 0
                Do While rs!fecha = diant
                    Select Case rs!tipo
                        Case "E":   tipo = "Efectivo"
                        Case "C":   tipo = "Ch. 3º"
                        Case "P":   tipo = "Ch. Propio"
                        Case "T":   tipo = "Transferencia"
                        Case "G":   tipo = "Gasto B."
                        Case "D":   tipo = "Créd. B."
                    End Select
                    If rs!Ing_egr = "I" Then
                            DataEnvironment1.Sistema.Execute "insert into " & msCajaTemp & " (fecha, caja, movimiento, ingegr, importe, desctipo, interno, concepto) " _
                            & "values(" & ssFecha(rs!fecha) & ", " & rs!caja & ", " & rs!movimiento & ", 'E', " & Replace(rs!importe, ",", ".") & ", '" & tipo & "', " & IIf(rs!interno <> Null, rs!interno, 0) & ", '" & rs!concepto & "')"
                    Else

                            DataEnvironment1.Sistema.Execute "insert into " & msCajaTemp & " (fecha, caja, movimiento, ingegr, importe, desctipo, interno, concepto) " _
                            & "values(" & ssFecha(rs!fecha) & ", " & rs!caja & ", " & rs!movimiento & ", 'S', " & Replace(rs!importe * -1, ",", ".") & ", '" & tipo & "', " & IIf(rs!interno <> Null, rs!interno, 0) & ", '" & rs!concepto & "')"
                    End If
                    
                    If rs!TIPODOC = "REA" Or rs!TIPODOC = "REB" Or rs!TIPODOC = "RAA" Or (Left(rs!concepto, 6) = "FacCdo") Then
                        punto1 = punto1 + rs!importe
                        If Left(rs!concepto, 6) = "FacCdo" Then
                            punto2 = punto2 + rs!importe
                        Else
                            punto3 = punto3 + rs!importe
                        End If
                    Else
                        If rs!TIPODOC = "FAC" Then
                            punto4 = punto4 + rs!importe
                            If buscotipopago(rs!codfp) = 1 Then
                                punto5 = punto5 + rs!importe 'TARJETA
                            Else
                                If buscotipopago(rs!codfp) = 2 Then
                                    punto6 = punto6 + rs!importe 'EFECTIVO
                                End If
                            End If
                        Else
                            If rs!TIPODOC = "O/P" Or rs!TIPODOC = "RAC" Then
                                punto7 = punto7 + rs!importe
                            Else
                                punto8 = punto8 + rs!importe
                            End If
                        End If
                    End If
                    rs.MoveNext
                    
                    If rs.EOF Then Exit Do
                Loop
            Wend
'        End If
        punto9 = punto9 + punto1 + punto8 - punto4 - punto7
        punto4 = punto4 * -1
        punto7 = punto7 * -1
        rs.Close
        Set rs = Nothing
    
    STR = "select * from " & msCajaTemp & " order by fecha"
    RptLisMovCajas.lblfecha = Date
    RptLisMovCajas.lblTitulo = "Listado de Mov. de Cajas del " & CStr(uFecha1.strFecha) & " al " & CStr(uFecha2.strFecha)
    RptLisMovCajas.lblsaldo = TraigosaldoAnterior()
    RptLisMovCajas.lblfechasaldo = uFecha1.strFecha
    
    'RptLisMovCajas.Field12 = obtenerDeSQL("select sum (importe) from " & msCajaTemp)'PARA  SABER  TODOS LOS MOVIMIENTOS
    'RptLisMovCajas.Field12 = grilla.TextMatrix(grilla.rows - 1, griIMPO)'SALDO ANTERIROR DE LA GRILLA
    
    RptLisMovCajas.lbl11 = punto1
    RptLisMovCajas.lbl22 = punto2
    RptLisMovCajas.lbl33 = punto3
    RptLisMovCajas.lbl44 = punto4
    RptLisMovCajas.lbl55 = punto5
    RptLisMovCajas.lbl66 = punto6
    RptLisMovCajas.lbl77 = punto7
    RptLisMovCajas.lbl88 = punto8
    RptLisMovCajas.lbl99 = punto9

        RptLisMovCajas.Data.Connection = DataEnvironment1.Sistema

    RptLisMovCajas.PageSettings.Orientation = ddOLandscape
    RptLisMovCajas.Data.Source = STR
    
    RptLisMovCajas.Show
End Sub

Function buscotipopago(tipo As Long) As Long
    Dim rs1 As New ADODB.Recordset

        rs1.Open "select * from formaspago where codigo = " & tipo & " and  activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic

    If Not rs1.EOF Then
        If rs1!tarjeta = True Then
            buscotipopago = 1
        Else
            If rs1!Efectivo = True Then
                buscotipopago = 2
            Else
                buscotipopago = 3
            End If
        End If
    End If
    rs1.Close
    Set rs1 = Nothing
End Function

Function TraigosaldoAnterior() As Double
    Dim rs1 As New ADODB.Recordset
    Dim tot As Double
    
    Dim tI As Double, tE As Double
    
    TraigosaldoAnterior = s2n(sumaimportes("I", s2n(uCaja.codigo)) - sumaimportes("E", s2n(uCaja.codigo)))
End Function
Private Function sumaimportes(Ing_egr As String, caja As Long)
        sumaimportes = s2n(obtenerDeSQL("select sum(importe) from movicaja where activo = 1 and ing_egr = '" & Ing_egr & "' and  caja = '" & caja & "' and fecha < " & ssFecha(uFecha1.strFecha)))
End Function


Private Sub cmdMostrar_Click()
    Dim ss As String, i As Long, ordernar_fecha As String
    Dim d, prov
'    If chkOrdernarFecha Then
'        ordernar_fecha = " order by fecha "
'    Else
'        ordernar_fecha = " "
'    End If

    If optFecha Then
        ordernar_fecha = " order by fecha "
    ElseIf OptImporte Then
        ordernar_fecha = " order by importe "
    ElseIf OptMovimiento Then
        ordernar_fecha = " order by movimiento "
    End If
    
    ss = "select " & griCamposSel & _
        " from movicaja " & _
        " where activo = 1 and fecha between " & uFecha1.ssFecha & " and " & uFecha2.ssFecha & " " & andFiltro & ordernar_fecha

        LlenarGrilla grilla, ss, False, griFECH, griIMPO
    CorrijoColumnas
    'grillaSumarizo grilla, Array(griIMPO) ' anda pa el ojete
    
    txtSaldoAnt = SaldoAnt()
    grillaSumariso grilla, griIMPO, CDbl(txtSaldoAnt)
    
    With grilla
        For i = 0 To .rows - 1
            If Left(.TextMatrix(i, griCONC), 3) = "O/P" Then
                d = Split(.TextMatrix(i, griCONC), ".")
                prov = d(UBound(d))
                .TextMatrix(i, griCONC) = .TextMatrix(i, griCONC) & " " & obtenerDeSQL("select descripcion from prov where codigo=" & prov)
            End If
        Next
    End With
End Sub

Private Sub grillaSumariso(g As Control, C As Integer, s As Double)
Dim aux_total As Double, RENGLON As Long
    
    aux_total = s
    With g
        For RENGLON = 1 To .rows - 1
            If .TextMatrix(RENGLON, C) <> "" Then
                aux_total = aux_total + .TextMatrix(RENGLON, C)
            End If
        Next
        .rows = .rows + 1
        .Row = .rows - 1
        .TextMatrix(.Row, 0) = "  Total  "
        .TextMatrix(.Row, C) = s2n(aux_total)
        .cell(flexcpFontBold, .Row, 0) = True
        .cell(flexcpFontBold, .Row, C) = True
    End With
End Sub

Private Sub CorrijoColumnas()
    Dim i As Long
    With grilla
        For i = 1 To .rows - 1
            If .TextMatrix(i, griI_E_) = "E" Then .TextMatrix(i, griIMPO) = -s2n(.TextMatrix(i, griIMPO))
            cambiogrilla i, griTIPO, "ECP", Array("Efectivo", "Ch 3ros", "Ch Propio")
            cambiogrilla i, griI_E_, "IE", Array("Ingreso", "Egreso")
        Next i
    End With
End Sub
Private Sub cambiogrilla(Y, X, viejo, Nuevo)
    Dim i As Long
    If grilla.TextMatrix(Y, X) > "" Then
        i = InStr(viejo, grilla.TextMatrix(Y, X))
        If i > 1 Then grilla.TextMatrix(Y, X) = Nuevo(i - 1)
    End If
    
End Sub



Private Function SaldoAnt() As Double
    ' saldo considerando:   ingr_egr, activo, fecha, y filtroactual
    Dim saling As Double, salegr As Double

        saling = s2n(obtenerDeSQL("select sum(importe) from movicaja where activo = 1  and ing_egr = 'I' " & andFiltro() & " and fecha < " & uFecha1.ssFecha))
        salegr = s2n(obtenerDeSQL("select sum(importe) from movicaja where activo = 1  and ing_egr = 'E' " & andFiltro() & " and fecha < " & uFecha1.ssFecha))
    SaldoAnt = Round(saling - salegr, 2)
End Function
Private Function andFiltro() As String
     If mFiltroActual > "" Then andFiltro = " and " & mFiltroActual
End Function

' -----------------------------------

Private Sub cmdsalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
        uCaja.ini "select responsable from cajas where codigo = '###' ", "Select codigo, responsable as [Descripcion              ] from cajas where activo = 1", False
    ucXls1.ini grilla, "C:\LisMovCaja"
    botones
    uCaja.codigo = 1
    Form_Resize
End Sub
Private Sub Form_Resize()
    Anclar grilla, Me, anclarLadosTodos
    Anclar fraTot, Me, anclarAbajo + anclarIzquierda
    Anclar cmdMostrar, Me, anclarAbajo + anclarDerecha
    Anclar CmdImprimir, Me, anclarAbajo + anclarDerecha
    Anclar ucXls1, Me, anclarAbajo
    Anclar cmdsalir, Me, anclarAbajo + anclarDerecha
End Sub

Private Sub optFiltro_Click(Index As Integer)
    mFiltroActual = mFiltros(Index)
    mTituloActual = mTitulos(Index)
End Sub
Private Sub optFiltro_Validate(Index As Integer, Cancel As Boolean)
    'cmdMostrar.Value = True
    mFiltroActual = mFiltros(Index)
    mTituloActual = mTitulos(Index)
End Sub
Private Sub uCaja_cambio(codigo As Variant)
    If codigo = 0 Then
        fraOpcion.Visible = True
        optFiltro(1).Value = True
    Else
        fraOpcion.Visible = False
        mFiltroActual = " caja = " & codigo & " "
    End If
End Sub
Private Sub botones()
    Dim i As Long
    mTitulos = Array("Caja #", "Todo", "Solo Banco", "Efectivo", "Ch 3ros", "Efect y Ch 3ros")
    mFiltros = Array("caja = 1", "", "caja = 0 ", "caja >0", "(caja = 0 and tipo = 'C') ", " ((caja = 0 and tipo = 'C') or (caja > 0))  ")
    For i = 0 To 5
        With optFiltro(i)
            .Width = 1350
            .Height = 360
            .Left = i * 1350
            .Top = 0
            .caption = mTitulos(i)
        End With
    Next i
End Sub
Private Sub ucXls1_Clic(Cancel As Boolean)
    ucXls1.aTitulo = "Movimiento para " & mTituloActual & " " & strRango()
End Sub
Private Function strRango()
    strRango = "Entre  " & uFecha1.dtfecha & " y  " & uFecha2.dtfecha
End Function

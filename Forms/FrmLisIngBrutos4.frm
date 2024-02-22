VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmLisIngBrutos4 
   Caption         =   "Listado de Ingreso Brutos por jurisdiccion (Compras) <No esta terminado>"
   ClientHeight    =   9390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12435
   Icon            =   "FrmLisIngBrutos4.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9390
   ScaleWidth      =   12435
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAsientos 
      Caption         =   "Ver Asientos"
      Height          =   990
      Left            =   10410
      Picture         =   "FrmLisIngBrutos4.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   345
      Width           =   915
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   990
      Left            =   9525
      Picture         =   "FrmLisIngBrutos4.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   360
      Width           =   855
   End
   Begin Gestion.ucXls ucXls1 
      Height          =   1005
      Left            =   8670
      TabIndex        =   8
      Top             =   360
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   1773
   End
   Begin VB.Frame frameoption 
      Height          =   945
      Left            =   150
      TabIndex        =   5
      Top             =   405
      Width           =   7575
      Begin Gestion.ucCoDe uJurisdiccion 
         Height          =   315
         Left            =   2145
         TabIndex        =   9
         Top             =   540
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   556
         CodigoWidth     =   1455
         CodigoInvalido  =   0
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Una jurisdiccion"
         Height          =   255
         Left            =   165
         TabIndex        =   7
         Top             =   555
         Width           =   1845
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Todas las jurisdicciones"
         Height          =   225
         Left            =   165
         TabIndex        =   6
         Top             =   225
         Value           =   -1  'True
         Width           =   2730
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   990
      Left            =   11370
      Picture         =   "FrmLisIngBrutos4.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   345
      Width           =   870
   End
   Begin VB.CommandButton cmdVer 
      Caption         =   "Mostrar"
      Height          =   1005
      Left            =   7785
      Picture         =   "FrmLisIngBrutos4.frx":2328
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   330
      Width           =   825
   End
   Begin MSComCtl2.DTPicker dtHasta 
      Height          =   315
      Left            =   1785
      TabIndex        =   2
      Top             =   90
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   556
      _Version        =   393216
      Format          =   63569921
      CurrentDate     =   40241
   End
   Begin MSComCtl2.DTPicker dtDesde 
      Height          =   285
      Left            =   150
      TabIndex        =   3
      Top             =   90
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   503
      _Version        =   393216
      Format          =   63569921
      CurrentDate     =   40241
   End
   Begin VSFlex7LCtl.VSFlexGrid gIIBB 
      Height          =   7935
      Left            =   150
      TabIndex        =   4
      Top             =   1410
      Width           =   12150
      _cx             =   21431
      _cy             =   13996
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
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuCambiarJ 
         Caption         =   "Cambiar Jurisdiccion"
      End
   End
End
Attribute VB_Name = "FrmLisIngBrutos4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAsientos_Click()
Dim idAs() As Long, i As Long, q As Long
Dim sConsul As String
With gIIBB
    If .rows > 1 Then
        q = -1
        For i = 0 To .rows - 1
            If s2n(.TextMatrix(i, 1)) > 0 Then
                q = q + 1
                ReDim Preserve idAs(q)
                idAs(q) = s2n(.TextMatrix(i, 1))
            End If
        Next
    End If
End With

For i = 0 To UBound(idAs)
    If i = 0 Then
        sConsul = " IDDOC IN (" & idAs(i)
    ElseIf i = UBound(idAs) Then
        sConsul = sConsul & "," & idAs(i) & ") ORDER BY FECHA"
    Else
        sConsul = sConsul & "," & idAs(i)
    End If
Next
frmAsientosIDDOC.MostrarDif sConsul
End Sub

Private Sub cmdImprimir_Click()

PrintG gIIBB, pVertical, "IIBB", Date, "IIBB POR JURISDICCION Y POR CUENTA", Printer.PaperSize

End Sub

Private Sub cmdVer_Click()
Dim sCon As String, i As Long, sWhe, sWhe2
Dim tot As Double
If uJurisdiccion.codigo = 0 Or uJurisdiccion.codigo = "" Then
    'el primero limpio traigo todo
    sWhe = ""
    sCon = "SELECT i.TIPODOC AS [DOCUMENTO],i.NRODOC AS [NRO  ],c.razonsocialprov as [PROVEEDOR],cuitprov as [CUIT],i.FECHA AS [FECHA     ],i.IMPORTE AS [IMPORTE       ],i.JURISDICCION AS [JURISDICCION                                    ] FROM IIBBJURISDICCION i inner join compras c on c.iddoc=i.iddoc WHERE i.ACTIVO=1 AND (i.FECHA >=" & ssFecha(dtDesde) & " AND i.FECHA<=" & ssFecha(dtHasta) & ") " _
            & " union " _
            & "SELECT i.TIPODOC AS [DOCUMENTO],i.NRODOC AS [NRO  ],c.razonsocialprov as [PROVEEDOR],cuitprov as [CUIT],i.FECHA AS [FECHA     ],i.IMPORTE AS [IMPORTE       ],i.JURISDICCION AS [JURISDICCION                                    ] FROM IIBBJURISDICCION i inner join transcom c on c.iddoc=i.iddoc WHERE i.ACTIVO=1 AND (i.FECHA >=" & ssFecha(dtDesde) & " AND i.FECHA<=" & ssFecha(dtHasta) & ") " _
            & " union " _
            & "SELECT TIPODOC AS [DOCUMENTO],NRODOC AS [NRO  ],razonsocialprov as [PROVEEDOR],cuitprov as [CUIT],FECHA AS [FECHA     ],IBCapital AS [IMPORTE       ],'CAPITAL FEDERAL'AS [JURISDICCION                                    ] FROM compras WHERE ACTIVO=1 AND (FECHA >=" & ssFecha(dtDesde) & " AND FECHA<=" & ssFecha(dtHasta) & ") and ibcapital>0 " _
            & " union " _
            & "SELECT TIPODOC AS [DOCUMENTO],NRODOC AS [NRO  ],razonsocialprov as [PROVEEDOR],cuitprov as [CUIT],FECHA AS [FECHA     ],IBCapital AS [IMPORTE       ],'CAPITAL FEDERAL'AS [JURISDICCION                                    ] FROM transcom WHERE ACTIVO=1 AND (FECHA >=" & ssFecha(dtDesde) & " AND FECHA<=" & ssFecha(dtHasta) & ") and ibcapital>0 "
ElseIf uJurisdiccion.codigo = "*" And (uJurisdiccion.codigo = 0 Or uJurisdiccion.codigo = "" Or uJurisdiccion.codigo = "*") Then
    'el primero * pero el segundo vacio traigo solo ese
    sCon = "SELECT TIPODOC AS [DOCUMENTO],NRODOC AS [NRO  ],razonsocialprov as [PROVEEDOR],cuitprov as [CUIT],FECHA AS [FECHA     ],IBCapital AS [IMPORTE       ],'" & uJurisdiccion.DESCRIPCION & "'AS [JURISDICCION                                    ] FROM compras WHERE ACTIVO=1 AND (FECHA >=" & ssFecha(dtDesde) & " AND FECHA<=" & ssFecha(dtHasta) & ") and ibcapital>0 " _
            & " union " _
            & "SELECT TIPODOC AS [DOCUMENTO],NRODOC AS [NRO  ],razonsocialprov as [PROVEEDOR],cuitprov as [CUIT],FECHA AS [FECHA     ],IBCapital AS [IMPORTE       ],'" & uJurisdiccion.DESCRIPCION & "'AS [JURISDICCION                                    ] FROM transcom WHERE ACTIVO=1 AND (FECHA >=" & ssFecha(dtDesde) & " AND FECHA<=" & ssFecha(dtHasta) & ") and ibcapital>0 "
ElseIf uJurisdiccion.codigo <> "" And uJurisdiccion.codigo <> "*" And (uJurisdiccion.codigo = 0 Or uJurisdiccion.codigo = "") Then
    'el primero con algo distinto a * pero el segundo vacio, traigo uno solo
    sWhe = " AND i.CODJUR=" & ssTexto(uJurisdiccion.codigo)
    sCon = "SELECT i.TIPODOC AS [DOCUMENTO],i.NRODOC AS [NRO  ],razonsocialprov as [PROVEEDOR],cuitprov as [CUIT],i.FECHA AS [FECHA     ],i.IMPORTE AS [IMPORTE       ],i.JURISDICCION AS [JURISDICCION                                    ] FROM IIBBJURISDICCION i inner join compras c on c.iddoc=i.iddoc WHERE i.ACTIVO=1 AND (i.FECHA >=" & ssFecha(dtDesde) & " AND i.FECHA<=" & ssFecha(dtHasta) & ") " & sWhe & _
            " union " & _
            "SELECT i.TIPODOC AS [DOCUMENTO],i.NRODOC AS [NRO  ],razonsocialprov as [PROVEEDOR],cuitprov as [CUIT],i.FECHA AS [FECHA     ],i.IMPORTE AS [IMPORTE       ],i.JURISDICCION AS [JURISDICCION                                    ] FROM IIBBJURISDICCION i inner join transcom c on c.iddoc=i.iddoc WHERE i.ACTIVO=1 AND (i.FECHA >=" & ssFecha(dtDesde) & " AND i.FECHA<=" & ssFecha(dtHasta) & ") " & sWhe
ElseIf uJurisdiccion.codigo = "*" And uJurisdiccion.codigo <> "" Then
    'traigo el rango seleccionado con la inclusion del *
    'sWhe = ""
    sWhe = " AND i.CODJUR>=" & ssTexto(uJurisdiccion.codigo)
    sWhe2 = " AND i.CODJUR<=" & ssTexto(uJurisdiccion.codigo)
    sCon = "SELECT i.TIPODOC AS [DOCUMENTO],i.NRODOC AS [NRO  ],razonsocialprov as [PROVEEDOR],cuitprov as [CUIT],i.FECHA AS [FECHA     ],i.IMPORTE AS [IMPORTE       ],i.JURISDICCION AS [JURISDICCION                                    ] FROM IIBBJURISDICCION i inner join compras c on c.iddoc=i.iddoc WHERE i.ACTIVO=1 AND (i.FECHA >=" & ssFecha(dtDesde) & " AND i.FECHA<=" & ssFecha(dtHasta) & ") " & sWhe & sWhe2 _
            & " union " _
            & "SELECT i.TIPODOC AS [DOCUMENTO],i.NRODOC AS [NRO  ],razonsocialprov as [PROVEEDOR],cuitprov as [CUIT],i.FECHA AS [FECHA     ],i.IMPORTE AS [IMPORTE       ],i.JURISDICCION AS [JURISDICCION                                    ] FROM IIBBJURISDICCION i inner join transcom c on c.iddoc=i.iddoc WHERE i.ACTIVO=1 AND (i.FECHA >=" & ssFecha(dtDesde) & " AND i.FECHA<=" & ssFecha(dtHasta) & ") " & sWhe & sWhe2 _
            & " union " _
            & "SELECT TIPODOC AS [DOCUMENTO],NRODOC AS [NRO  ],razonsocialprov as [PROVEEDOR],cuitprov as [CUIT],FECHA AS [FECHA     ],IBCapital AS [IMPORTE       ],'CAPITAL FEDERAL'AS [JURISDICCION                                    ] FROM compras WHERE ACTIVO=1 AND (FECHA >=" & ssFecha(dtDesde) & " AND FECHA<=" & ssFecha(dtHasta) & ") and ibcapital>0 " _
            & " union " _
            & "SELECT TIPODOC AS [DOCUMENTO],NRODOC AS [NRO  ],razonsocialprov as [PROVEEDOR],cuitprov as [CUIT],FECHA AS [FECHA     ],IBCapital AS [IMPORTE       ],'CAPITAL FEDERAL'AS [JURISDICCION                                    ] FROM transcom WHERE ACTIVO=1 AND (FECHA >=" & ssFecha(dtDesde) & " AND FECHA<=" & ssFecha(dtHasta) & ") and ibcapital>0 "
ElseIf uJurisdiccion.codigo <> "" And uJurisdiccion.codigo <> "*" And uJurisdiccion.codigo <> "" Then
    'el primero con algo distinto a * pero el segundo vacio, traigo uno solo
    sWhe = " AND i.CODJUR>=" & ssTexto(uJurisdiccion.codigo)
    sWhe2 = " AND i.CODJUR<=" & ssTexto(uJurisdiccion.codigo)
    sCon = "SELECT i.TIPODOC AS [DOCUMENTO],i.NRODOC AS [NRO  ],razonsocialprov as [PROVEEDOR],cuitprov as [CUIT],i.FECHA AS [FECHA     ],i.IMPORTE AS [IMPORTE       ],i.JURISDICCION AS [JURISDICCION                                    ] FROM IIBBJURISDICCION i inner join compras c on c.iddoc=i.iddoc WHERE i.ACTIVO=1 AND (i.FECHA >=" & ssFecha(dtDesde) & " AND i.FECHA<=" & ssFecha(dtHasta) & ") " & sWhe & sWhe2 _
            & " union " _
            & "SELECT i.TIPODOC AS [DOCUMENTO],i.NRODOC AS [NRO  ],razonsocialprov as [PROVEEDOR],cuitprov as [CUIT],i.FECHA AS [FECHA     ],i.IMPORTE AS [IMPORTE       ],i.JURISDICCION AS [JURISDICCION                                    ] FROM IIBBJURISDICCION i inner join transcom c on c.iddoc=i.iddoc WHERE i.ACTIVO=1 AND (i.FECHA >=" & ssFecha(dtDesde) & " AND i.FECHA<=" & ssFecha(dtHasta) & ") " & sWhe & sWhe2
End If
'sCon = "SELECT TIPODOC AS [DOCUMENTO  ],NRODOC AS [NRO  ],FECHA AS [FECHA     ],IMPORTE AS [IMPORTE       ],JURISDICCION AS [JURISDICCION                                    ] FROM IIBBJURISDICCION WHERE ACTIVO=1 AND (FECHA >=" & ssFecha(dtDesde) & " AND FECHA<=" & ssFecha(dtHasta) & ") " & sWhe


tot = 0
LlenarGrilla gIIBB, sCon, False
With gIIBB
    If .rows > 1 Then
        .ColWidth(0) = 1200
        .ColWidth(1) = 900
        .ColWidth(2) = 3000
        .ColWidth(3) = 1300
        .ColWidth(4) = 1000
        .ColWidth(5) = 1200
        .ColWidth(6) = 1700
        For i = 1 To .rows - 1
            .TextMatrix(i, 5) = s2n(.TextMatrix(i, 5), 2, True)
            tot = tot + s2n(.TextMatrix(i, 5))
        Next
        If tot <> 0 Then
            .AddItem "" & Chr(9) & "" & Chr(9) & "TOTAL" & Chr(9) & "" & Chr(9) & "" & Chr(9) & tot
        End If
    End If
End With

End Sub

Private Function a_que_columna(fIdDoc As Long)
Dim ctasMayorista As String, ctasMinorista As String, ctasReparacion As String, ctasAlquiler As String, i As Long
Dim ctaFactura As String, rsFacturas As New ADODB.Recordset
Dim rMayorista As Double, rMinorista As Double, rReparacion As Double, rAlquiler As Double, rOtros As Double, rImporte As Double
Dim tmp
tmp = obtenerDeSQL("select ctasmayorista,ctasminorista,ctasreparacion,ctasalquiler from datosempresa where idempresa=" & gEMPR_idEmpresa)
ctasMayorista = Trim(tmp(0))
ctasMinorista = Trim(tmp(1))
ctasReparacion = Trim(tmp(2))
ctasAlquiler = Trim(tmp(3))
'ctaFactura = Trim(sSinNull(obtenerDeSQL("select m.cuenta from asientos a inner join mayor m on a.idasiento=m.idasiento where a.activo=1 and m.cuenta like '4%' and m.haber >0 and a.iddoc=" & fIddoc)))
ctaFactura = "select m.cuenta,m.debe,m.haber from asientos a inner join mayor m on a.idasiento=m.idasiento where a.activo=1 and m.cuenta like '4%' and a.iddoc=" & fIdDoc
rsFacturas.Open ctaFactura, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
rMayorista = 0
rMinorista = 0
rReparacion = 0
rAlquiler = 0
rOtros = 0

'If fIdDoc = 5891 Then Stop

With rsFacturas
    If .EOF And .BOF Then
    Else
        .MoveFirst
        For i = 0 To .RecordCount - 1
            rImporte = 0
            If s2n(!haber) > 0 Then
                rImporte = s2n(!haber)
            ElseIf s2n(!Debe) > 0 Then
                rImporte = -(s2n(!Debe))
            End If
            
            If InStr(ctasMayorista, Trim(!Cuenta)) Then
                rMayorista = rMayorista + rImporte
            ElseIf InStr(ctasMinorista, Trim(!Cuenta)) Then
                rMinorista = rMinorista + rImporte
            ElseIf InStr(ctasReparacion, Trim(!Cuenta)) Then
                rReparacion = rReparacion + rImporte
            ElseIf InStr(ctasAlquiler, Trim(!Cuenta)) Then
                rAlquiler = rAlquiler + rImporte
            Else
                rOtros = rOtros + rImporte
            End If
            .MoveNext
            
        Next
    End If
End With

a_que_columna = Array(rMayorista, rMinorista, rReparacion, rAlquiler, rOtros)
End Function

Private Function addGrilla(gCodFac As String, gIDDOC As String, gFactura As String, gMayorista As String, gMinorista As String, gReparacion As String, gAlquiler As String, gOtros As String)
Dim rr As Long
With gIIBB
    .AddItem ""
    rr = .rows - 1
    .TextMatrix(rr, 0) = gCodFac
    .TextMatrix(rr, 1) = gIDDOC
    .TextMatrix(rr, 2) = gFactura
    .TextMatrix(rr, 3) = gMayorista
    .TextMatrix(rr, 4) = gMinorista
    .TextMatrix(rr, 5) = gReparacion
    .TextMatrix(rr, 6) = gAlquiler
    .TextMatrix(rr, 7) = gOtros
End With
End Function

Private Function iGrilla()
With gIIBB
    .clear
    .rows = 0
    .rows = 1
    .cols = 0
    .cols = 8
    '.TextMatrix(0, 0) = "CODFAC"
    '.TextMatrix(0, 1) = "IDDOC"
    '.TextMatrix(0, 2) = "FACTURA"
    '.TextMatrix(0, 3) = "MAYORISTA"
    '.TextMatrix(0, 4) = "MINORISTA"
    '.TextMatrix(0, 5) = "REPARACION"
    '.TextMatrix(0, 6) = "ALQUIER"
    '.TextMatrix(0, 7) = "OTROS"
    .ColWidth(0) = 0
    .ColWidth(1) = 0
    .ColWidth(2) = 3000
    .ColWidth(3) = 1500
    .ColWidth(4) = 1500
    .ColWidth(5) = 1500
    .ColWidth(6) = 1500
    .ColWidth(7) = 1500
End With
End Function



Private Sub Form_Load()
dtDesde = CDate("01/01/" & Year(Date))
dtHasta = Date
uJurisdiccion.ini "Select descripcion from provincias where codigo='###'", "Select [CODIGO     ],[DESCRIPCION           ] FROM PROVINCIAS WHERE ACTIVO=1", True
ucXls1.ini gIIBB, "C:\IIBB_JURISDICCION_CUENTAS.XLS"
iGrilla
End Sub


Private Sub gIIBB_DblClick()
If s2n(gIIBB.TextMatrix(gIIBB.Row, 1)) > 0 Then
    frmAsientosIDDOC.MostrarDif "  IDDOC=" & s2n(gIIBB.TextMatrix(gIIBB.Row, 1))
End If

End Sub



Private Sub gIIBB_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 Then
    Me.PopupMenu mnuMenu
End If
End Sub

Private Sub mnuCambiarJ_Click()
Dim Res
Dim rr As Long, cf As Long, uf As String
'MsgBox "A continuacion indique la nueva jurisdiccion.", vbInformation
Res = frmBuscar.MostrarSql("Select CODIGO, DESCRIPCION from provincias where activo=1")
If Res > "" Then
    rr = gIIBB.Row
    If rr < 0 Then
        MsgBox "Seleccione un comprobante......", vbInformation
    Else
        cf = s2n(gIIBB.TextMatrix(rr, 0))
        If cf > 0 Then
            uf = "update facturaventa set provincia=" & ssTexto(Res) & " where codigo=" & cf
            DataEnvironment1.Sistema.Execute uf
            If MsgBox("Guardado...Actualizar listado?", vbInformation + vbYesNo) = vbYes Then
                cmdVer_Click
            End If
        End If
    End If
End If
End Sub

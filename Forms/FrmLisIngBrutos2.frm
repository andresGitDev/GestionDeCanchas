VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmLisIngBrutos2 
   Caption         =   "LISTADO DE IIBB POR JURISDICCION (Compras)"
   ClientHeight    =   9570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   ScaleHeight     =   9570
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   840
      Left            =   8280
      Picture         =   "FrmLisIngBrutos2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   465
      Width           =   840
   End
   Begin VB.CommandButton cmdVer 
      Caption         =   "Ejecutar"
      Height          =   825
      Left            =   5400
      Picture         =   "FrmLisIngBrutos2.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   465
      Width           =   825
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   825
      Left            =   7320
      Picture         =   "FrmLisIngBrutos2.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   480
      Width           =   825
   End
   Begin Gestion.ucXls ucXls1 
      Height          =   825
      Left            =   6360
      TabIndex        =   4
      Top             =   480
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   1455
   End
   Begin Gestion.ucCoDe uJurisdiccion 
      Height          =   330
      Left            =   135
      TabIndex        =   3
      Tag             =   "Coloque 0 para ver todos"
      Top             =   555
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   582
      CodigoWidth     =   1455
      CodigoInvalido  =   0
   End
   Begin MSComCtl2.DTPicker dtHasta 
      Height          =   315
      Left            =   1755
      TabIndex        =   2
      Top             =   150
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   556
      _Version        =   393216
      Format          =   62849025
      CurrentDate     =   40241
   End
   Begin MSComCtl2.DTPicker dtDesde 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   150
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   503
      _Version        =   393216
      Format          =   62849025
      CurrentDate     =   40241
   End
   Begin VSFlex7LCtl.VSFlexGrid gIIBB 
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   1470
      Width           =   9285
      _cx             =   16378
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
   Begin Gestion.ucCoDe uJurisdiccion2 
      Height          =   330
      Left            =   120
      TabIndex        =   5
      Tag             =   "Coloque 0 para ver todos"
      Top             =   960
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   582
      CodigoWidth     =   1455
      CodigoInvalido  =   0
   End
End
Attribute VB_Name = "FrmLisIngBrutos2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdImprimir_Click()
    PrintG gIIBB, pVertical, "IIBB compras", Date, "IIBB compras"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdVer_Click()
Dim sDiv As String, sCon As String, i As Long, sEmpiesa As Long, sWhe, sWhe2
Dim tot As Double, stot As Double
If uJurisdiccion.codigo = 0 Or uJurisdiccion.codigo = "" Then
    'el primero limpio traigo todo
    sWhe = ""
    sCon = "SELECT i.TIPODOC AS [DOCUMENTO],i.NRODOC AS [NRO  ],c.razonsocialprov as [PROVEEDOR],cuitprov as [CUIT],i.FECHA AS [FECHA     ],i.IMPORTE AS [IMPORTE       ],i.JURISDICCION AS [JURISDICCION                                    ] FROM IIBBJURISDICCION i inner join compras c on c.iddoc=i.iddoc WHERE i.ACTIVO=1 AND (i.FECHA >=" & ssFecha(dtDesde) & " AND i.FECHA<=" & ssFecha(dtHasta) & ") " _
            & " union " _
            & "SELECT i.TIPODOC AS [DOCUMENTO],i.NRODOC AS [NRO  ],c.razonsocialprov as [PROVEEDOR],cuitprov as [CUIT],i.FECHA AS [FECHA     ],i.IMPORTE AS [IMPORTE       ],i.JURISDICCION AS [JURISDICCION                                    ] FROM IIBBJURISDICCION i inner join transcom c on c.iddoc=i.iddoc WHERE i.ACTIVO=1 AND (i.FECHA >=" & ssFecha(dtDesde) & " AND i.FECHA<=" & ssFecha(dtHasta) & ") " _
            & " union " _
            & "SELECT TIPODOC AS [DOCUMENTO],NRODOC AS [NRO  ],razonsocialprov as [PROVEEDOR],cuitprov as [CUIT],FECHA AS [FECHA     ],IBCapital AS [IMPORTE       ],'CAPITAL FEDERAL'AS [JURISDICCION                                    ] FROM compras WHERE ACTIVO=1 AND (FECHA >=" & ssFecha(dtDesde) & " AND FECHA<=" & ssFecha(dtHasta) & ") and ibcapital>0 " _
            & " union " _
            & "SELECT TIPODOC AS [DOCUMENTO],NRODOC AS [NRO  ],razonsocialprov as [PROVEEDOR],cuitprov as [CUIT],FECHA AS [FECHA     ],IBCapital AS [IMPORTE       ],'CAPITAL FEDERAL'AS [JURISDICCION                                    ] FROM transcom WHERE ACTIVO=1 AND (FECHA >=" & ssFecha(dtDesde) & " AND FECHA<=" & ssFecha(dtHasta) & ") and ibcapital>0 " _
            & "ORDER BY JURISDICCION"
ElseIf uJurisdiccion.codigo = "*" And (uJurisdiccion2.codigo = 0 Or uJurisdiccion2.codigo = "" Or uJurisdiccion2.codigo = "*") Then
    'el primero * pero el segundo vacio traigo solo ese
    sCon = "SELECT TIPODOC AS [DOCUMENTO],NRODOC AS [NRO  ],razonsocialprov as [PROVEEDOR],cuitprov as [CUIT],FECHA AS [FECHA     ],IBCapital AS [IMPORTE       ],'" & uJurisdiccion.DESCRIPCION & "'AS [JURISDICCION                                    ] FROM compras WHERE ACTIVO=1 AND (FECHA >=" & ssFecha(dtDesde) & " AND FECHA<=" & ssFecha(dtHasta) & ") and ibcapital>0 " _
            & " union " _
            & "SELECT TIPODOC AS [DOCUMENTO],NRODOC AS [NRO  ],razonsocialprov as [PROVEEDOR],cuitprov as [CUIT],FECHA AS [FECHA     ],IBCapital AS [IMPORTE       ],'" & uJurisdiccion.DESCRIPCION & "'AS [JURISDICCION                                    ] FROM transcom WHERE ACTIVO=1 AND (FECHA >=" & ssFecha(dtDesde) & " AND FECHA<=" & ssFecha(dtHasta) & ") and ibcapital>0 "
ElseIf uJurisdiccion.codigo <> "" And uJurisdiccion.codigo <> "*" And (uJurisdiccion2.codigo = 0 Or uJurisdiccion2.codigo = "") Then
    'el primero con algo distinto a * pero el segundo vacio, traigo uno solo
    sWhe = " AND i.CODJUR=" & ssTexto(uJurisdiccion.codigo)
    sCon = "SELECT i.TIPODOC AS [DOCUMENTO],i.NRODOC AS [NRO  ],razonsocialprov as [PROVEEDOR],cuitprov as [CUIT],i.FECHA AS [FECHA     ],i.IMPORTE AS [IMPORTE       ],i.JURISDICCION AS [JURISDICCION                                    ] FROM IIBBJURISDICCION i inner join compras c on c.iddoc=i.iddoc WHERE i.ACTIVO=1 AND (i.FECHA >=" & ssFecha(dtDesde) & " AND i.FECHA<=" & ssFecha(dtHasta) & ") " & sWhe & _
            " union " & _
            "SELECT i.TIPODOC AS [DOCUMENTO],i.NRODOC AS [NRO  ],razonsocialprov as [PROVEEDOR],cuitprov as [CUIT],i.FECHA AS [FECHA     ],i.IMPORTE AS [IMPORTE       ],i.JURISDICCION AS [JURISDICCION                                    ] FROM IIBBJURISDICCION i inner join transcom c on c.iddoc=i.iddoc WHERE i.ACTIVO=1 AND (i.FECHA >=" & ssFecha(dtDesde) & " AND i.FECHA<=" & ssFecha(dtHasta) & ") " & sWhe
ElseIf uJurisdiccion.codigo = "*" And uJurisdiccion2.codigo <> "" Then
    'traigo el rango seleccionado con la inclusion del *
    'sWhe = ""
    sWhe = " AND i.CODJUR>=" & ssTexto(uJurisdiccion.codigo)
    sWhe2 = " AND i.CODJUR<=" & ssTexto(uJurisdiccion2.codigo)
    sCon = "SELECT i.TIPODOC AS [DOCUMENTO],i.NRODOC AS [NRO  ],razonsocialprov as [PROVEEDOR],cuitprov as [CUIT],i.FECHA AS [FECHA     ],i.IMPORTE AS [IMPORTE       ],i.JURISDICCION AS [JURISDICCION                                    ] FROM IIBBJURISDICCION i inner join compras c on c.iddoc=i.iddoc WHERE i.ACTIVO=1 AND (i.FECHA >=" & ssFecha(dtDesde) & " AND i.FECHA<=" & ssFecha(dtHasta) & ") " & sWhe & sWhe2 _
            & " union " _
            & "SELECT i.TIPODOC AS [DOCUMENTO],i.NRODOC AS [NRO  ],razonsocialprov as [PROVEEDOR],cuitprov as [CUIT],i.FECHA AS [FECHA     ],i.IMPORTE AS [IMPORTE       ],i.JURISDICCION AS [JURISDICCION                                    ] FROM IIBBJURISDICCION i inner join transcom c on c.iddoc=i.iddoc WHERE i.ACTIVO=1 AND (i.FECHA >=" & ssFecha(dtDesde) & " AND i.FECHA<=" & ssFecha(dtHasta) & ") " & sWhe & sWhe2 _
            & " union " _
            & "SELECT TIPODOC AS [DOCUMENTO],NRODOC AS [NRO  ],razonsocialprov as [PROVEEDOR],cuitprov as [CUIT],FECHA AS [FECHA     ],IBCapital AS [IMPORTE       ],'CAPITAL FEDERAL'AS [JURISDICCION                                    ] FROM compras WHERE ACTIVO=1 AND (FECHA >=" & ssFecha(dtDesde) & " AND FECHA<=" & ssFecha(dtHasta) & ") and ibcapital>0 " _
            & " union " _
            & "SELECT TIPODOC AS [DOCUMENTO],NRODOC AS [NRO  ],razonsocialprov as [PROVEEDOR],cuitprov as [CUIT],FECHA AS [FECHA     ],IBCapital AS [IMPORTE       ],'CAPITAL FEDERAL'AS [JURISDICCION                                    ] FROM transcom WHERE ACTIVO=1 AND (FECHA >=" & ssFecha(dtDesde) & " AND FECHA<=" & ssFecha(dtHasta) & ") and ibcapital>0 "
ElseIf uJurisdiccion.codigo <> "" And uJurisdiccion.codigo <> "*" And uJurisdiccion2.codigo <> "" Then
    'el primero con algo distinto a * pero el segundo vacio, traigo uno solo
    sWhe = " AND i.CODJUR>=" & ssTexto(uJurisdiccion.codigo)
    sWhe2 = " AND i.CODJUR<=" & ssTexto(uJurisdiccion2.codigo)
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
        sDiv = .TextMatrix(1, 6)
        sEmpiesa = 1
cambio:
stot = 0
        For i = sEmpiesa To .rows - 1
            
            If .TextMatrix(i, 6) <> sDiv Then
                sDiv = .TextMatrix(i, 6)
                .AddItem "", i
                .TextMatrix(i, 5) = s2n(stot)
                sEmpiesa = i + 1
                GoTo cambio
            Else
                stot = stot + s2n(.TextMatrix(i, 5))
            End If
        Next
        .AddItem "", i
        .TextMatrix(i, 5) = s2n(stot)
        
        
        If tot <> 0 Then
            .AddItem "" & Chr(9) & "" & Chr(9) & "TOTAL" & Chr(9) & "" & Chr(9) & "" & Chr(9) & tot
        End If
    End If
End With

End Sub

Private Sub Form_Load()
dtDesde = CDate("01/01/" & Year(Date))
dtHasta = Date
uJurisdiccion.ini "Select descripcion from provincias where codigo='###'", "Select [CODIGO     ],[DESCRIPCION           ] FROM PROVINCIAS WHERE ACTIVO=1", True
uJurisdiccion2.ini "Select descripcion from provincias where codigo='###'", "Select [CODIGO     ],[DESCRIPCION           ] FROM PROVINCIAS WHERE ACTIVO=1", True
gIIBB.rows = 1
ucXls1.ini gIIBB, "C:\IIBB_JURISDICCION.XLS"
End Sub



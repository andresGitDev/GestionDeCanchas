VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmAsientosGeneracion 
   Caption         =   "Generación de asientos"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8295
   Icon            =   "frmAsientosGeneracion.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   8295
   StartUpPosition =   2  'CenterScreen
   Begin Gestion.ucXls ucXls1 
      Height          =   810
      Left            =   6960
      TabIndex        =   7
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1429
   End
   Begin Gestion.ucFecha uFeDesde 
      Height          =   315
      Left            =   1620
      TabIndex        =   3
      Top             =   540
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      FechaInit       =   0
   End
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Height          =   4515
      Left            =   120
      TabIndex        =   6
      Top             =   1020
      Width           =   8055
      _cx             =   14208
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
   Begin VB.CommandButton cmdGenerarAsiento 
      Caption         =   "Generar Periodo"
      Height          =   375
      Left            =   3780
      TabIndex        =   5
      Top             =   540
      Width           =   1395
   End
   Begin Gestion.ucFecha uFeHasta 
      Height          =   315
      Left            =   2640
      TabIndex        =   4
      Top             =   540
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      FechaInit       =   2
   End
   Begin VB.Label lblVoyPor 
      Height          =   375
      Left            =   5280
      TabIndex        =   8
      Top             =   540
      Width           =   1455
   End
   Begin VB.Label lblEjercicio 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2004"
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   1035
   End
   Begin VB.Label Label2 
      Caption         =   "Periodo"
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   540
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Ejercicio:"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmAsientosGeneracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


' NC
'   en dos, graba NC devolucion como NCT, pero pierde discriminacion A y B
'   aca busca items con formula, para saber si es devolucion, creo q es confiable
'   laura sugiere campo nuevo, me ahorra la barrida de sql, y es mas seguro
'   NC_xDevolucion = 1

'ND
'   en DOS es NDR, quizas sea confiable suponiendo que no hay NDB
' aca es ND_xChequeRechazado = 1



' Cobranza Clientes
'   ??????????????????????????????????????????????????????????????
'   no entiendo los loops, el replace en la tmp HABER, ni porqué separa por movimiento.
'   PARECE q todo movicaja va sumado discriminado por cuenta pal haber, ok.
'   PARECE q en detmovcaja, 2 cuentas se separan, las demas van a otra cuenta PERO
'   solo hay 3 cuentas validas, asi q es lo mismo
'   ACA,
'        no separe x movimiento, sumo discriminado por cuenta, nada mas...
'  ????


'Private Const DATO_FIJO = 1 ' Tonka
' ver como hacemos parametrizacion
Private Const CTA_VALORES_RECHAZADOS = "1113002"

Private Const CTA_DEUDORES_VENTAS_LOCALES = "1131001"
Private Const CTA_DEUDORES_VENTAS_EXTERIOR = "1131002"
Private Const CTA_RET_IMP_GANANCIAS = "1132001"
Private Const CTA_ANTICIPO_A_PROVEEDORES = "1133001"

Private Const CTA_PROVEEDORES = "2110001"
Private Const CTA_PROVEEDORES_DEL_EXTERIOR = "2110002"
Private Const CTA_ANTICIPO_DE_CLIENTES = "2110003"
Private Const CTA_RET_IMP_GCIAS_A_TERCEROS = "2140002"
Private Const CTA_PERC_IIBB_BSAS = "2140004"
Private Const CTA_IVA_VENTAS = "2141001"
Private Const CTA_IVA_VENTAS_NO_INSCRIPTOS = "2141002"
Private Const CTA_IVA_COMPRAS = "2141003"
Private Const CTA_RET_IVA = "2141004"
Private Const CTA_PERCEP_IVA_RG3337 = "2141005"
Private Const CTA_BONOS_DE_CREDITO_FISCAL = "2141006"

Private Const CTA_nRgIb = "2144004" ' ?????????

Private Const CTA_VENTAS_LOCALES = "4110001"
Private Const CTA_VENTAS_DE_EXPORTACION = "4110002"
Private Const CTA_VENTAS_IIBB = "2141010"

Private Const CTA_GASTOS_BANCARIOS = "5340001"

'private const CTA_

Private Debe As Cuenta
Private haber As Cuenta
Private Asiento As Asiento  'clase
Private g As LiGrilla       'clase
Private G_ASIE As Long
Private G_DEBE As Long
Private G_HABE As Long
Private G_CONC As Long
Private G_CUEN As Long
Private G_DIFE As Long
'

Private ssql As String, temp As Variant
'

Private Function sWhereFV() As String
    sWhereFV = " where FacturaVenta.activo = 1 and (FacturaVenta.fecha  between " & uFeDesde.ConvertFecha & " and " & uFeHasta.ConvertFecha & ") "
End Function
Private Function sWhere() As String
    sWhere = " where activo = 1 and (fecha  between " & uFeDesde.ConvertFecha & " and " & uFeHasta.ConvertFecha & ") "
End Function
Private Function sWhereFecha() As String
    sWhereFecha = " where (fecha  between " & uFeDesde.ConvertFecha & " and " & uFeHasta.ConvertFecha & ") "
End Function

Private Function a_FactClientes() As Boolean
    'VENTAS fac Clientes
    Dim nTotCC As Double, nTotCExt As Double, nIvaFac As Double, nNetCont As Double, nIIBB As Double

    'FAA FAB
    ssql = "SELECT Sum(Neto) AS SumNeto, Sum(Total) AS SumTotal, Sum(Iva) AS SumIva, sum(IIBB) as SumIIBB " _
        & " FROM FacturaVenta " _
        & sWhereFV & " and (TipoDoc='FAA' Or TipoDoc='FAB') "
    temp = obtenerDeSQL(ssql)
    nNetCont = s2n(temp(0))
    nTotCC = s2n(temp(1))
    nIvaFac = s2n(temp(2))
    nIIBB = s2n(temp(3))
    'FAE
    ssql = "SELECT Sum(Neto) AS SumNeto, Sum(Total) AS SumTotal, Sum(Iva) AS SumIva " _
        & " FROM FacturaVenta " _
        & sWhereFV & " and (TipoDoc='FAE') "
    temp = obtenerDeSQL(ssql)
    nTotCExt = s2n(temp(1))
    
    With Asiento
        'Facturacion Clientes
        AsNuevo "Facturacion Clientes ", "F"
            .AgregarItem CTA_DEUDORES_VENTAS_LOCALES, nTotCC, 0
            .AgregarItem CTA_DEUDORES_VENTAS_EXTERIOR, nTotCExt, 0
            .AgregarItem CTA_IVA_VENTAS, 0, nIvaFac
            .AgregarItem CTA_VENTAS_IIBB, 0, nIIBB
            .AgregarItem CTA_VENTAS_LOCALES, 0, nNetCont
            .AgregarItem CTA_VENTAS_DE_EXPORTACION, 0, nTotCExt
    '       .AgregarItem "2141002", 0, nNiFac
'            MA_CODCTA:='1111001'''          MA_MONTO:= nTotFac   en DOS esta SIEMPRE EN 0, (????)
'            MA_CODCTA:='1131001'''          MA_MONTO:= nTotCC
'            MA_CODCTA:='1131002'''          MA_MONTO:= nTotCExt
'            MA_CODCTA:='2141001'''          MA_MONTO:=   -nIvaFac
'             MA_CODCTA:='2141002''           MA_MONTO:=   -nNiFac     iva no inscripto
'            MA_CODCTA:='4110001'''          MA_MONTO:=   -nNetCont
'            MA_CODCTA:='4110002'''          MA_MONTO:=   -nTotCExt
    End With
    a_FactClientes = aGrillayGraba()
End Function

Private Function a_NC_Devol() As Boolean
    'NC devolucion
    Dim nNetNC As Double, nTotNCex As Double, nTotIvaCf As Double, nTotNC As Double
    'NCA NCB devolucion
    ssql = "SELECT Sum(Neto) AS SumNeto, Sum(Total) AS SumTotal, Sum(Iva) AS SumIva " _
        & " FROM FacturaVenta " _
        & sWhereFV _
        & " and  (TIPODOC = 'NCA' Or TIPODOC = 'NCB')" _
        & " and  NC_xDevolucion = 1"
    temp = obtenerDeSQL(ssql)
    nTotNC = s2n(temp(1))
    nNetNC = s2n(temp(0))
    nTotIvaCf = s2n(temp(2))
    
    'NCE devolucion
    ssql = "SELECT Sum(Neto) AS SumNeto, Sum(Total) AS SumTotal, Sum(Iva) AS SumIva " _
        & " FROM FacturaVenta " _
        & sWhereFV _
        & " and  (TIPODOC = 'NCE')" _
        & " and  NC_xDevolucion = 1"
    temp = obtenerDeSQL(ssql)
    nTotNCex = s2n(temp(1))
    
    With Asiento
        'nc dev
         AsNuevo "N.Cred Dev. Clientes", "F"
                .AgregarItem CTA_VENTAS_LOCALES, nNetNC, 0
                .AgregarItem CTA_VENTAS_DE_EXPORTACION, nTotNCex, 0
                .AgregarItem CTA_IVA_VENTAS, nTotIvaCf, 0
                
                .AgregarItem CTA_DEUDORES_VENTAS_LOCALES, 0, nTotNC
                .AgregarItem CTA_DEUDORES_VENTAS_EXTERIOR, 0, nTotNCex
    ''          MA_CODCTA:='4110001'''          MA_MONTO:=nNetNc
    '           MA_CODCTA:='4110002''           MA_MONTO:=nTotNcEx
    '           MA_CODCTA:='2141001''           MA_MONTO:=nTotIvaCf
    '           MA_CODCTA:='2141002''          MA_MONTO:=nNiNc
    '           MA_CODCTA:='1131001''          MA_MONTO:=-nTotNc
    '           MA_CODCTA:='1131002''          MA_MONTO:=-nToNcEx
    End With
    a_NC_Devol = aGrillayGraba()
End Function

Private Function a_NC_Otros() As Boolean
    'NC otros
    Dim nTotONc As Double, nNetONc As Double, nTotOIvaCf As Double, nTotONCex As Double
    ' otras NCA NCB O0
    ssql = "SELECT Sum(Neto) AS SumNeto, Sum(Total) AS SumTotal, Sum(Iva) AS SumIva " _
        & " FROM FacturaVenta " _
        & sWhereFV _
        & " and  (TIPODOC = 'NCA' Or TIPODOC = 'NCB')" _
        & " and  NC_xDevolucion = 0"
    temp = obtenerDeSQL(ssql)
    
    nNetONc = s2n(temp(0))
    nTotONc = s2n(temp(1))
    nTotOIvaCf = s2n(temp(2))
    
    'NCE _Otros
    ssql = "SELECT Sum(Total) AS SumTotal " _
        & " FROM FacturaVenta " _
        & sWhereFV _
        & " and  (TIPODOC = 'NCE' ) " _
        & " and  NC_xDevolucion = 0"
    temp = obtenerDeSQL(ssql)
    nTotONCex = s2n(temp)
    
    With Asiento
        'nc otros
        AsNuevo "N.Cred Otr conc. Clientes ", "F"
            .AgregarItem CTA_VENTAS_LOCALES, nNetONc, 0
            .AgregarItem CTA_IVA_VENTAS, nTotOIvaCf, 0
            .AgregarItem CTA_DEUDORES_VENTAS_LOCALES, 0, nTotONc
        '          MA_CODCTA:='4110003''          MA_MONTO:=nNetONc
        '          MA_CODCTA:='2141001''          MA_MONTO:=nTotOIvaCf
        '          MA_CODCTA:='2141002''          MA_MONTO:=nNiONc
        '          MA_CODCTA:='1131001''          MA_MONTO:=-nTotONc
    End With
    a_NC_Otros = aGrillayGraba()
End Function

Public Function a_nd_clientes() As Boolean
    'ND
    Dim nTotNDex As Double, nNetND As Double, nTotND As Double, nIvaND As Double  ', nTotIvai As Double
    'ND otras
    ssql = "SELECT Sum(Neto) AS SumNeto, Sum(Total) AS SumTotal, Sum(Iva) AS SumIva " _
        & " FROM FacturaVenta " _
        & sWhereFV _
        & " and (TipoDoc='NDA' or TipoDoc = 'NDB')  and ND_xChequeRechazado = 0"
    temp = obtenerDeSQL(ssql)
    
    nNetND = s2n(temp(0))
    nTotND = s2n(temp(1))
    nIvaND = s2n(temp(2))
    
    'ND ext
    ssql = "SELECT Sum(Total) AS SumTotal " _
        & " FROM FacturaVenta " _
        & sWhereFV _
        & " and (TipoDoc='NDE') "
    temp = obtenerDeSQL(ssql)
    nTotNDex = s2n(temp)
    
    With Asiento
        'nd  comunes
        AsNuevo "N.Deb Clientes ", "F"
            .AgregarItem CTA_DEUDORES_VENTAS_LOCALES, nTotND, 0
            .AgregarItem CTA_DEUDORES_VENTAS_EXTERIOR, nTotNDex, 0
            .AgregarItem CTA_VENTAS_LOCALES, 0, nNetND
            .AgregarItem CTA_VENTAS_DE_EXPORTACION, 0, nTotNDex
            .AgregarItem CTA_IVA_VENTAS, 0, nIvaND '???nTotIvai
'            .AgregarItem CTA_IVA_VENTAS_NO_INSCRIPTOS, 0, nNiNd
'          MA_CODCTA:='1131001''          MA_MONTO:=nTotNd
'          MA_CODCTA:='1131002''          MA_MONTO:=nTotNdEx
'          MA_CODCTA:='4110001''          MA_MONTO:=-nNetNd
'          MA_CODCTA:='4110002''          MA_MONTO:=-nTotNdEx
'          MA_CODCTA:='2141001''          MA_MONTO:=-nTotIvai
'          MA_CODCTA:='2141002''          MA_MONTO:=-nNiNd
    End With
    a_nd_clientes = aGrillayGraba()
End Function

Public Function a_nd_ch_rechazo() As Boolean
    'ND ch rech
    Dim nTotNdCR As Double, nNetNdCR As Double, nIvaNDcr As Double, nTotRechazo As Double
    'ND cheque rechazado
        ssql = "SELECT Sum(Neto) AS SumNeto, Sum(Total) AS SumTotal, Sum(Iva) AS SumIva, Sum(NoGrav) as SumNoGrav " _
            & " FROM FacturaVenta " _
            & sWhereFV _
            & " and ND_xChequeRechazado = 1" '  and (TipoDoc='NDA' or TipoDoc = 'NDB')
        temp = obtenerDeSQL(ssql)
        nNetNdCR = s2n(temp(0)) 'CTA_GASTOS_BANCARIOS
        nTotNdCR = s2n(temp(1)) 'CTA_DEUDORES_VENTAS_LOCALES
        nIvaNDcr = s2n(temp(2))
        nTotRechazo = s2n(temp(3)) 'CTA_VALORES_RECHAZADOS
''       sSql = " SELECT DETALLEMOVCAJAS.CUENTA, Sum(DETALLEMOVCAJAS.IMPORTE) AS SumaDeIMPORTE, Sum(FacturaVenta.NoGrav) AS SumaDeNoGrav " _
''            & " FROM DETALLEMOVCAJAS INNER JOIN FacturaVenta ON (DETALLEMOVCAJAS.NRODOC = FacturaVenta.NroFactura) AND (DETALLEMOVCAJAS.TIPDOC = FacturaVenta.TipoDoc) " _
''            & sWhereFV _
''            & " and FacturaVenta.ND_xChequeRechazado = 1 " _
''            & " GROUP BY DETALLEMOVCAJAS.CUENTA "
''        temp = obtenerDeSQL(sSql)
 
'ND CHEQUE RECHAZADO
    With Asiento
        AsNuevo "N.Deb. x Ch/Rechazados ", "F"
            .AgregarItem CTA_DEUDORES_VENTAS_LOCALES, nTotNdCR, 0
            .AgregarItem CTA_GASTOS_BANCARIOS, 0, nNetNdCR
            .AgregarItem CTA_IVA_VENTAS, 0, nIvaNDcr
            .AgregarItem CTA_VALORES_RECHAZADOS, 0, nTotRechazo

'            .AgregarItem CTA_IVA_VENTAS_NO_INSCRIPTOS, 0, nNiNdCR
'          MA_CODCTA:='1131001''          MA_MONTO:=nTotNdCR
'          MA_CODCTA:='5340001''          MA_MONTO:=-nNetNdCR
'          MA_CODCTA:='2141001''          MA_MONTO:=-nTotIvaiCR
'          MA_CODCTA:='2141002''          MA_MONTO:=-nNiNdCR
    End With
    a_nd_ch_rechazo = aGrillayGraba()
End Function

Public Function a_cobro_clientes() As Boolean
    'cobro clientes
    Dim nTotCobEx As Double, nTotAdCob As Double, nTotCobCli As Double
    'Dim rs As New ADODB.Recordset
    'ret
''    Dim nRetIva As Double, nRetIB As Double, nRetGan As Double, nRetBon As Double
    
''    'ret
''    sSql = "SELECT Sum(Total)SumTot " _
''        & " FROM FacturaVenta " _
''        & sWhereFV & " and TipoDoc='RET' "
''    temp = obtenerDeSQL(sSql)
''    nRetIva = s2n(temp)
''    'rib
''    sSql = "SELECT Sum(Total)SumTot " _
''        & " FROM FacturaVenta " _
''        & sWhereFV & " and TipoDoc='RIB' "
''    temp = obtenerDeSQL(sSql)
''    nRetIB = s2n(temp)
''    'rga
''    sSql = "SELECT Sum(Total)SumTot " _
''        & " FROM FacturaVenta " _
''        & sWhereFV & " and TipoDoc='RGA' "
''    temp = obtenerDeSQL(sSql)
''    nRetGan = s2n(temp)
''    'rbo
''    sSql = "SELECT Sum(Total)SumTot " _
''        & " FROM FacturaVenta " _
''        & sWhereFV & " and TipoDoc='RBO' "
''    temp = obtenerDeSQL(sSql)
''    nRetBon = s2n(temp)
    
    
'----------------------------------------------
    Dim rs As New ADODB.Recordset, ssM, ssD, ssR, cue As String, tmpRet As Double
    
    ssM = "SELECT m.CUENTA,  Sum(d.IMPORTE) AS SumaDeIMPORTE " _
        & " FROM MoviCaja AS m INNER JOIN DETALLEMOVCAJAS AS d ON m.MOVIMIENTO = d.MOVIMIENTO " _
        & " where activo = 1 and (m.fecha  between " & uFeDesde.ConvertFecha & " and " & uFeHasta.ConvertFecha & ") " _
        & " And (d.Origen = 'RA' Or d.Origen = 'RC')" _
        & " GROUP BY m.CUENTA  order by m.cuenta"
    
    ssD = "SELECT d.CUENTA, Sum(d.IMPORTE) AS SumaDeIMPORTE " _
        & " FROM MoviCaja AS m INNER JOIN DETALLEMOVCAJAS AS d ON m.MOVIMIENTO = d.MOVIMIENTO " _
        & " where activo = 1 and (m.fecha  between " & uFeDesde.ConvertFecha & " and " & uFeHasta.ConvertFecha & ") " _
        & " And (d.Origen = 'RA' Or d.Origen = 'RC') " _
        & " GROUP BY d.CUENTA  order by d.cuenta"
    
    ssR = "select codigo, cuenta from TipoRetenciones where cuenta > ''"
'---------------------------------------------------------------
    With Asiento
        AsNuevo "Cobranzas a Clientes", "R"
            temp = 0
            
            rs.Open ssR, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
            While Not rs.EOF
                ssql = "SELECT Sum(Total)SumTot FROM FacturaVenta " & sWhereFV & " and TipoDoc='" & rs!codigo & "' "
                tmpRet = s2n(obtenerDeSQL(ssql))
                temp = temp + tmpRet
                '.AgregarItem s2n(rs!Cuenta, 0), temp, 0
                .AcumularItem s2n(rs!Cuenta, 0), tmpRet, 0
                rs.MoveNext
            Wend
            rs.Close
            
            '.AgregarItem CTA_RET_IMP_GANANCIAS, nRetGan, 0
            '.AgregarItem CTA_RET_IVA, nRetIva, 0
            '.AgregarItem CTA_BONOS_DE_CREDITO_FISCAL, nRetBon, 0
            
            rs.Open ssM, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
            While Not rs.EOF
                .AcumularItem rs!Cuenta, s2n(rs!SumaDeIMPORTE), 0
                rs.MoveNext
            Wend
            rs.Close
            
    '       While nNroMov==MCD_MOVIMIENTO .And. !MCD->(Eof())
    '              If MCD_CUENTA='2110003' '     nTotAdCob  += MCD_IMPORTE
    '              If MCD_CUENTA='1131002''      nTotCobEx  += MCD_IMPORTE
    '              Else '                        nTotCobCli += MCD_IMPORTE
            rs.Open ssD, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
            While Not rs.EOF
                cue = rs!Cuenta
                If cue = CTA_DEUDORES_VENTAS_EXTERIOR Or cue = CTA_ANTICIPO_DE_CLIENTES Then
                    .AgregarItem cue, 0, rs!SumaDeIMPORTE
                Else
                    temp = temp + rs!SumaDeIMPORTE
                End If
                rs.MoveNext
            Wend
            rs.Close
            .AgregarItem CTA_DEUDORES_VENTAS_LOCALES, 0, (temp) '+ nRetGan + nRetIva + nRetBon)
            
'          MA_CODCTA:='2141004''          MA_MONTO:=nRetIva
'          MA_CODCTA:='1132001''          MA_MONTO:=nRetGan
'          MA_CODCTA:='2141006''          MA_MONTO:=nRetBon
'          MA_CODCTA:='1131001''          MA_MONTO:=-( nTotCobCli +nRetGan+nRetIva+nRetBon)
'          MA_CODCTA:='1131002''          MA_MONTO:=-(nTotCobEx)
'          MA_CODCTA:='2110003''          MA_MONTO:=-nTotAdCob
    End With
    a_cobro_clientes = aGrillayGraba()
    
End Function

Private Function a_Ajus_Cred_Clientes() As Boolean
    'ajustes
    Dim nTotCAju As Double, nTotDExAj As Double, nTotCExAj As Double, nTotDAju As Double
    Dim nTotCEx As Double, nTotDEx As Double
    'ACC  cliente local
    ssql = " SELECT Sum(FacturaVenta.Total) AS SumaDeTotal " _
        & " FROM FacturaVenta INNER JOIN Clientes " _
        & " ON FacturaVenta.Cliente = Clientes.codigo " _
        & sWhereFV _
        & " and Clientes.zona = 1 and FacturaVenta.TipoDoc='ACC'"
    temp = obtenerDeSQL(ssql)
    nTotCAju = s2n(temp)
    'ACC  cliente ext
    ssql = " SELECT Sum(FacturaVenta.Total) AS SumaDeTotal " _
        & " FROM FacturaVenta INNER JOIN Clientes " _
        & " ON FacturaVenta.Cliente = Clientes.codigo " _
        & sWhereFV _
        & " and Clientes.zona <> 1 and FacturaVenta.TipoDoc='ACC'"

    temp = obtenerDeSQL(ssql)
    nTotCExAj = s2n(temp) 'nTotCEx + s2n(temp)

    With Asiento
        AsNuevo "Aj.Cred. Clientes ", "F"
    '          MA_CODCTA:='4110001''          MA_MONTO:=nTotCAju
    '          MA_CODCTA:='4110002''          MA_MONTO:=nTotCEx
    '          MA_CODCTA:='1131001''          MA_MONTO:=-nTotCAju
    '          MA_CODCTA:='1131002''          MA_MONTO:=-nTotCEx
            .AgregarItem CTA_VENTAS_LOCALES, nTotCAju, 0
            .AgregarItem CTA_VENTAS_DE_EXPORTACION, nTotCEx, 0
            .AgregarItem CTA_DEUDORES_VENTAS_LOCALES, 0, nTotCAju
            .AgregarItem CTA_DEUDORES_VENTAS_EXTERIOR, 0, nTotCEx
    End With
    a_Ajus_Cred_Clientes = aGrillayGraba()
End Function

Private Function a_Ajus_Debi_Clientes() As Boolean
    'ajustes
    Dim nTotCAju As Double, nTotDExAj As Double, nTotCExAj As Double, nTotDAju As Double
    Dim nTotCEx As Double, nTotDEx As Double

    'ACD  cliente local
    ssql = " SELECT Sum(FacturaVenta.Total) AS SumaDeTotal " _
        & " FROM FacturaVenta INNER JOIN Clientes " _
        & " ON FacturaVenta.Cliente = Clientes.codigo " _
        & sWhereFV _
        & " and Clientes.zona = 1 and FacturaVenta.TipoDoc='ACD'"

    temp = obtenerDeSQL(ssql)
    nTotDAju = s2n(temp) 'nTotDAju + s2n(temp)
    'ACD  cliente ext
    ssql = " SELECT Sum(FacturaVenta.Total) AS SumaDeTotal " _
        & " FROM FacturaVenta INNER JOIN Clientes " _
        & " ON FacturaVenta.Cliente = Clientes.codigo " _
        & sWhereFV _
        & " and Clientes.zona <> 1 and FacturaVenta.TipoDoc='ACD'"
    
    temp = obtenerDeSQL(ssql)
    nTotDExAj = s2n(temp) 'nTotDEx + s2n(temp)
 
    With Asiento
        AsNuevo "Aj.Deb. Clientes ", "F"
'          MA_CODCTA:='1131001''          MA_MONTO:=nTotDAju
'          MA_CODCTA:='1131002''          MA_MONTO:=nTotDEx
'          MA_CODCTA:='4110001''          MA_MONTO:=-nTotDAju
'          MA_CODCTA:='4110002''          MA_MONTO:=-nTotDEx
            .AgregarItem CTA_DEUDORES_VENTAS_LOCALES, nTotDAju, 0
            .AgregarItem CTA_DEUDORES_VENTAS_EXTERIOR, nTotDEx, 0
            .AgregarItem CTA_VENTAS_LOCALES, 0, nTotDAju
            .AgregarItem CTA_VENTAS_DE_EXPORTACION, 0, nTotDEx
    End With
    a_Ajus_Debi_Clientes = aGrillayGraba()
End Function


Private Function a_imput_anticip_VENtas() As Boolean
    'anticipo
    Dim nTotAntCob As Double
    'AnticiposClientes
    ssql = "SELECT Sum(DETALLEMOVCAJAS.IMPORTE) AS SumaDeIMPORTE " _
        & " From DETALLEMOVCAJAS " _
        & sWhereFecha _
        & " and DETALLEMOVCAJAS.ORIGEN ='IA' "
    temp = obtenerDeSQL(ssql)
    
    nTotAntCob = s2n(temp)

    With Asiento
        'IA
        AsNuevo "Imputaciones de Anticipos ", "R"
            .AgregarItem CTA_ANTICIPO_DE_CLIENTES, nTotAntCob, 0
            .AgregarItem CTA_DEUDORES_VENTAS_LOCALES, 0, nTotAntCob
            'mcd
    '       MA_CODCTA:='2110003''       MA_MONTO:=nTotAntCob
    '       MA_CODCTA:='1131001''       MA_MONTO:=-nTotAntCob
    End With
    a_imput_anticip_VENtas = aGrillayGraba()
End Function

Private Function SignoTipoDocCompra(quetipo) As Double
    Select Case quetipo
    Case "FAC", "N/D", "APD"
        SignoTipoDocCompra = 1
    Case "N/C", "APC"
        SignoTipoDocCompra = -1
    Case "RAC"
        SignoTipoDocCompra = 0
    End Select
End Function
Private Function SignoMCD(queOrigen) As Double
    Select Case queOrigen
    Case "FT", "FC", "ND", "AD": SignoMCD = 1
    Case "NC", "AC": SignoMCD = -1
    Case Else: SignoMCD = 0
    End Select
End Function
Private Function SumImpCompra(prov As Long, TIPODOC As String, NroDoc As Long) As Double
    SumImpCompra = s2n(obtenerDeSQL("select (imp_Int + Der_est + ibCapital  + ibProvincia ) from Transcom where codpr = " & prov & " and TipoDoc = '" & TIPODOC & "' and nroDoc = " & NroDoc))
    If SumImpCompra = 0 Then
      SumImpCompra = s2n(obtenerDeSQL("select (imp_Int + Der_est + ibCapital  + ibProvincia ) from compras where codpr = " & prov & " and TipoDoc = '" & TIPODOC & "' and nroDoc = " & NroDoc))
    End If
End Function
Private Sub AcumulaMoviCaja(prov As Long, TIPODOC As String, NroDoc As Long)
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from  movicaja " _
        & sWhere() _
        & " and codprov = " & prov & " and TipoDoc = '" & TIPODOC & "' and NroDoc = " & NroDoc _
        & " order by cuenta ", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    While Not rs.EOF
        Asiento.AcumularItem rs!Cuenta, 0, rs!Importe
        rs.MoveNext
    Wend
    Set rs = Nothing
End Sub
Private Function a_Compras() As Boolean
    'compras
    Dim nTotIva As Double ', nTotIvaT As Double  'Private nTotIva21 As Double, nTotIva27 As Double, nTotIva10 As Double,', nTotIvaT21 As Double, nTotIvaT27 As Double, nTotIvaT10 As Double,
    Dim nPercep As Double, nPercepIb As Double
    Dim nExeComp As Double, nDerEst As Double, nDerTEst As Double
    Dim nRetIComp As Double, nImInComp As Double, nRgComp As Double, nRgIb As Double, nPercepT As Double, nPercepTIb As Double, nRetITComp As Double, nImInTComp As Double, nRgTComp As Double, nExeTComp As Double
    Dim nAjuCred As Double, nAjuDeb As Double
    Dim nTotProveed As Double, nTotProvExt As Double  ', nRetenido As Double
    
    Dim rs As New ADODB.Recordset, signo As Double

    
    'COMPRAS
    ssql = "SELECT Sum(c.IVA_21 +  c.IVA_27 + c.IVA_10) AS SumaIVA " _
        & " , Sum(c.PERCEPC) AS SumaDePERCEPC, Sum(c.IBprovincia) AS SumaDeIBprovincia, Sum(c.IBcapital) AS SumaDeIBcapital " _
        & " , Sum(c.IVA_9) AS SumaDeIVA_9, Sum(c.IMP_INT) AS SumaDeIMP_INT, Sum(c.RET_GAN) AS SumaDeRET_GAN " _
        & " , Sum(c.DER_EST) AS SumaDeDER_EST, Sum(c.TOTAL) AS SumaDeTOTAL, c.exterior, c.CONTADO, c.TIPODOC " _
        & " , exter " _
        & " FROM COMPRAS AS c  left join prov on prov.codigo = c.codpr " _
        & sWhereFecha() & " and c.activo = 1 " _
        & " GROUP BY c.exterior, c.CONTADO, c.TIPODOC, exter  "
    With rs
        .Open ssql, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        While Not .EOF
            signo = SignoTipoDocCompra(!TIPODOC)
            nTotIva = nTotIva + (signo * !sumaiva)
            nPercep = nPercep + (signo * !SumaDePERCEPC)
            nRgComp = nRgComp + (signo * !SumaDeRET_GAN)
            nRetIComp = nRetIComp + (signo * !SumaDeIVA_9)
'            nPercepIb = nPercepIb +  (signo * ! ) 'PERCEIB_CO
'''            nImInComp = nImInComp + (signo * !SumaDeIMP_INT)
'''            nDerEst = nDerEst + (signo * !SumaDeDER_EST)
'            nRgIb = nRgIb + (signo * ! )  'RETIIBB_CO
            If !exter = "S" Then nTotProvExt = nTotProvExt + (signo * !sumadetotal)
            If !exter <> "S" Then nTotProveed = nTotProveed + (signo * !sumadetotal)
            .MoveNext
        Wend
        .Close
    End With
    'TRANSCOM
    ssql = "SELECT Sum(c.IVA_21 +  c.IVA_27 + c.IVA_10) AS SumaIVA " _
        & " , Sum(c.PERCEPC) AS SumaDePERCEPC, Sum(c.IBprovincia) AS SumaDeIBprovincia, Sum(c.IBcapital) AS SumaDeIBcapital " _
        & " , Sum(c.IVA_9) AS SumaDeIVA_9, Sum(c.IMP_INT) AS SumaDeIMP_INT, Sum(c.RET_GAN) AS SumaDeRET_GAN " _
        & " , Sum(c.DER_EST) AS SumaDeDER_EST, Sum(c.TOTAL) AS SumaDeTOTAL, c.exterior,  c.TIPODOC " _
        & " , exter " _
        & " FROM transcom AS c    left join prov on prov.codigo = c.codpr " _
        & sWhereFecha() & " and c.activo = 1 " _
        & " GROUP BY c.exterior,  c.TIPODOC, exter "
    With rs
        .Open ssql, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        While Not .EOF 'And Not IsNull(!sumadetotal)
            signo = SignoTipoDocCompra(!TIPODOC)
            nTotIva = nTotIva + (signo * s2n(!sumaiva))
            nPercep = nPercep + (signo * s2n(!SumaDePERCEPC))
            nRgComp = nRgComp + (signo * s2n(!SumaDeRET_GAN))
            nRetIComp = nRetIComp + (signo * s2n(!SumaDeIVA_9))
'            nPercepIb = nPercepIb +  (signo * ! ) 'PERCEIB_CO
'''            nImInComp = nImInComp + (signo * s2n(!SumaDeIMP_INT))
'''            nDerEst = nDerEst + (signo * s2n(!SumaDeDER_EST))
'            nRgIb = nRgIb + (signo * ! )  'RETIIBB_CO
            If !exter = "S" Then nTotProvExt = nTotProvExt + (signo * !sumadetotal)
            If !exter <> "S" Then nTotProveed = nTotProveed + (signo * !sumadetotal)
            .MoveNext
        Wend
        .Close
    End With
    
    
    'MOVCJDET
    'usas el while en MCD para totalizar por cuenta, y no uso el sum de sql asi
    'al salir del while, aseguras que buscas UNA SOLA VEZ  la factura relacionada
    Dim ssMCD, ssMC
'    ssMCD = "SELECT Sum(IMPORTE) AS SumaIMPORTE, CUENTA, ORIGEN " _
        & " FROM DETALLEMOVCAJAS " _
        & sWhereFecha() _
        & " GROUP BY CUENTA, ORIGEN " _
        & " ORDER BY CUENTA " '_
        '& " Having (((Origen) = 'FC' Or (Origen) = 'ND' Or (Origen) = 'AD' Or (Origen) = 'FT' Or (Origen) = 'NC' Or (Origen) = 'AC'))"
'    ssMCD_FC = "select distinct codprov, serie, tipdoc, nrodoc from detallemovcajas    " _
'        & sWhereFecha() _
'        & " and origen = 'FC' "
'    ssMC = "SELECT Sum(IMPORTE) AS SumaIMPORTE, CUENTA, "

    ssMCD = "SELECT CODPROV, TIPDOC, NRODOC, CUENTA, ORIGEN, Sum(IMPORTE) AS SumaDeIMPORTE " _
        & " FROM DETALLEMOVCAJAS " _
        & sWhereFecha() _
        & " GROUP BY CODPROV,  TIPDOC, NRODOC, CUENTA, ORIGEN " _
        & " Having (((Origen) = 'FC' Or (Origen) = 'ND' Or (Origen) = 'AD' Or (Origen) = 'FT' Or (Origen) = 'NC' Or (Origen) = 'AC' ))" _
        & " ORDER BY CODPROV,  TIPDOC, NRODOC, CUENTA "
        
    With Asiento
        AsNuevo "Compras ", "C"
            .AgregarItem CTA_IVA_COMPRAS, nTotIva, 0
            .AgregarItem CTA_PERCEP_IVA_RG3337, nPercep, 0
            .AgregarItem CTA_RET_IVA, nRetIComp, 0
            .AgregarItem CTA_RET_IMP_GCIAS_A_TERCEROS, nRgComp, 0
            
'            rs.Open ssMCD, daTaenvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'            While Not rs.EOF
'                If SignoMCD(rs!Origen) <> 0 Then ' pregunta al pedo en version final
'                .AcumularItem rs!Cuenta, (SignoMCD(rs!Origen) * s2n(rs!SumaIMPORTE)), 0
'                End If ' pregunta al pedo en version final
'                rs.MoveNext
'            Wend
'            rs.Close
            
            Dim prevCodProv As Long, prevTipDoc As String, prevNroDoc As Long, prevCuenta As Long, prevOrigen As String
            rs.Open ssMCD, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
            If Not rs.EOF Then
                prevCodProv = rs!CodProv
                prevTipDoc = rs!TipDoc
                prevNroDoc = rs!NroDoc
                prevCuenta = rs!Cuenta
                prevOrigen = rs!Origen
                While Not rs.EOF
                    Do While rs!CodProv = prevCodProv And rs!TipDoc = prevTipDoc And rs!NroDoc = prevNroDoc 'And rs!serie = prevSerie
                        .AcumularItem rs!Cuenta, (SignoMCD(rs!Origen) * s2n(rs!SumaDeIMPORTE)), 0
                        rs.MoveNext
                        If rs.EOF Then Exit Do
                    Loop
                    .AcumularItem prevCuenta, (SignoMCD(prevOrigen) * SumImpCompra(prevCodProv, prevTipDoc, prevNroDoc)), 0
                    AcumulaMoviCaja prevCodProv, prevTipDoc, prevNroDoc
                    If Not rs.EOF Then
                        prevCodProv = rs!CodProv
                        prevTipDoc = rs!TipDoc
                        prevNroDoc = rs!NroDoc
                        prevCuenta = rs!Cuenta
                        prevOrigen = rs!Origen
                    End If
                Wend
            End If
            rs.Close
            .AgregarItem CTA_PROVEEDORES, 0, nTotProveed
            .AgregarItem CTA_PROVEEDORES_DEL_EXTERIOR, 0, nTotProvExt

'    '       MA_CODCTA :='2141003''       MA_MONTO  :=nTotIva21+nTotIva27+nTotIva10
'    '       MA_CODCTA :='2141005''       MA_MONTO  :=nPercep
'    '       MA_CODCTA :='2141004''       MA_MONTO  :=nRetIComp
'    '       MA_CODCTA :='2140002''       MA_MONTO  :=nRGComp
'    '       MA_CODCTA :='2144004''       MA_MONTO  :=nRGIb
'
'    '       Select DEBE'    While !EOF()
'    '                   MA_CODCTA:=DEBE->CODCTA'       MA_MONTO:=DEBE->IMPORTE
'
'    '       Select HABER'    While !EOF()
'    '                   MA_CODCTA:=HABER->CODCTA'       MA_MONTO:=-HABER->IMPORTE
'    '       MA_CODCTA:='2110001''       MA_MONTO:=-nTotProveed
'    '       MA_CODCTA:='2110002''       MA_MONTO:=-nTotProvExt
    End With
    a_Compras = aGrillayGraba()
End Function

Private Function a_pago_proveedores() As Boolean
    Dim rs As New ADODB.Recordset, umov As Long, tempo
    Dim nTotPagProv As Double, nTotPagExt As Double, nTotAdProv As Double, nRetenido As Double

    ssql = "select origen, sum(importe) as SumaImporte, cuenta, movimiento from DetalleMovCajas as d left join prov on prov.codigo = d.codprov " _
        & sWhereFecha _
        & " and  (origen = 'OP' or origen = 'PC') " _
        & " group by origen, cuenta , movimiento "
        
    nRetenido = _
        s2n(obtenerDeSQL("select sum(RetGan) from compras  " & sWhere() & " and tipodoc = 'RAC' ")) + _
        s2n(obtenerDeSQL("select sum(RetGan) from transcom " & sWhere() & " and tipodoc = 'RAC' "))
    
    
'   //GRABA ASIENTO DE PAGO A PROVEEDORES
    With Asiento
        AsNuevo "Pago a Proveedores ", "P"
        
            rs.Open ssql, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
            While Not rs.EOF
                If rs!Origen = "OP" Then
                    If rs!Cuenta = CTA_PROVEEDORES_DEL_EXTERIOR Then
                        .AcumularItem CTA_PROVEEDORES_DEL_EXTERIOR, rs!SumaImporte, 0  ' nTotPagExt = nTotPagExt + rs!SumaImporte
                    Else
                        .AcumularItem CTA_PROVEEDORES, rs!SumaImporte, 0  'nTotPagProv = nTotPagProv + rs!SumaImporte
                    End If
                Else
                    .AcumularItem CTA_ANTICIPO_A_PROVEEDORES, rs!SumaImporte, 0  'nTotAdProv = nTotAdProv + !importe
                End If
                
                If umov <> rs!movimiento Then
                    umov = rs!movimiento
                    tempo = obtenerDeSQL("select cuenta, importe from MoviCaja where movimiento = " & umov)
                    If Not IsEmpty(tempo) Then
                        .AcumularItem tempo(0), 0, s2n(tempo(1))
                    End If
                End If
                rs.MoveNext
            Wend
            .AcumularItem CTA_ANTICIPO_A_PROVEEDORES, nRetenido, 0
            .AgregarItem CTA_RET_IMP_GCIAS_A_TERCEROS, 0, nRetenido

''''       MA_CODCTA:='2110001''       MA_MONTO:=nTotPagProv
''''       MA_CODCTA:='2110002''       MA_MONTO:=nTotPagExt
''''       MA_CODCTA:='1133001''       MA_MONTO:=nTotAdProv+nRetenido
'''            .AgregarItem CTA_PROVEEDORES, nTotPagProv, 0
'''            .AgregarItem CTA_PROVEEDORES_DEL_EXTERIOR, nTotPagExt, 0
'''            .AgregarItem CTA_ANTICIPO_A_PROVEEDORES, nTotAdProv + nRetenido, 0
''''       Select HABER'    While !EOF()
''''                   MA_CODCTA:=HABER->CODCTA'       MA_MONTO:=-HABER->IMPORTE
''''       MA_CODCTA:='2140002''       MA_MONTO:=-nRetenido
'''            .AgregarItem CTA_RET_IMP_GCIAS_A_TERCEROS, 0, nRetenido
    End With
    a_pago_proveedores = aGrillayGraba()
End Function

Private Function a_Imput_anticip_Compras() As Boolean
    'anticipo prov
    Dim nTotAntExt As Double, nTotAntPag As Double

'    // GENERA ASIENTO DE IMPUTACIONES DE ANTICIPOS A  PROVEEDORES
'    nTotAntPag := 0
'    nTotAntExt := 0
    ssql = " select sum(importe) from DetalleMovCajas " _
        & " inner join prov on codprov = prov.codigo " _
        & sWhere _
        & " and origen = 'IP' and prov.exter = 'S' "
    nTotAntExt = s2n(obtenerDeSQL(ssql))
    
    ssql = " select sum(importe) from DetalleMovCajas " _
        & " right join prov on codprov = prov.codigo " _
        & sWhere _
        & " and origen = 'IP' and prov.exter <> 'S' "
    nTotAntPag = s2n(obtenerDeSQL(ssql))


    With Asiento
'    // GRABA ASIENTO DE IMPUTACIONES DE ANTICIPOS A  PROVEEDORES
        AsNuevo "Imput. Anticipos a Prov. ", "R"
            .AgregarItem CTA_PROVEEDORES, nTotAntPag, 0
            .AgregarItem CTA_PROVEEDORES_DEL_EXTERIOR, nTotAntExt, 0
            .AgregarItem CTA_ANTICIPO_A_PROVEEDORES, 0, (nTotAntPag + nTotAntExt)
'           MA_CODCTA:='2110001''          MA_MONTO:=nTotAntPag
'           MA_CODCTA:='2110002''          MA_MONTO:=nTotAntExt
'           MA_CODCTA:='1133001''       MA_MONTO:=-(nTotAntPag+nTotAntExt)
    End With
    a_Imput_anticip_Compras = aGrillayGraba()
End Function

Private Function a_Mov_Caja_y_Bancos() As Boolean
    Dim rs As New ADODB.Recordset
    Dim NroMov As Long, tempo

    ssql = "select sum(importe) as SumaImporte, origen, cuenta, movimiento  from DetalleMovCajas " _
        & sWhereFecha _
        & " and " _
        & "   (Origen = 'IE' or Origen = 'EE' or Origen = 'GB' or Origen = 'CB' " _
        & " or Origen = 'IC' or Origen = 'LC' or Origen = 'LV' " _
        & " or Origen = 'LA' or Origen = 'DE' or Origen = 'DC' or Origen = 'TR') " _
        & " group by movimiento, cuenta, origen"
        '_
        '& " and cuenta <> '" & CTA_VENTAS_LOCALES & "' and cuenta <> '" & CTA_VENTAS_DE_EXPORTACION & "' "

    With Asiento
        AsNuevo "Mov Caja y Bancos", "B"
        rs.Open ssql, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        'If Not rs!EOF Then NroMov = rs!movimiento
        While Not rs.EOF
            If rs!Cuenta <> CTA_VENTAS_LOCALES And rs!Cuenta <> CTA_VENTAS_DE_EXPORTACION Then
                Select Case rs!Origen
                Case "IE", "CB", "IC", "LA":                    .AcumularItem rs!Cuenta, 0, rs!SumaImporte
                Case "EE", "GB", "LC", "LV", "DE", "DC", "TR":  .AcumularItem rs!Cuenta, rs!SumaImporte, 0
                End Select
            End If
            
            If NroMov <> rs!movimiento Then
                NroMov = rs!movimiento
                tempo = obtenerDeSQL("select cuenta,importe from movicaja where movimiento = " & NroMov)
                If Not IsEmpty(tempo) Then
                    Select Case rs!Origen
                    Case "IE", "CB", "IC", "LA":                    .AcumularItem tempo(0), tempo(1), 0
                    Case "EE", "GB", "LC", "LV", "DE", "DC", "TR":  .AcumularItem tempo(0), 0, tempo(1)
                    End Select
                End If
            End If
            rs.MoveNext
        Wend
    
    End With
    a_Mov_Caja_y_Bancos = aGrillayGraba()
End Function

Private Sub GenAsiento() 'nOper)
    If ON_ERROR_HABILITADO Then On Error GoTo UfaGENE
    
    ''DE_BeginTrans
    grilla.Redraw = flexRDNone
    Screen.ActiveForm.MousePointer = vbHourglass
    
    a_FactClientes
    a_NC_Devol
    a_NC_Otros
    a_nd_clientes
    a_nd_ch_rechazo
    a_cobro_clientes
    a_Ajus_Cred_Clientes
    a_Ajus_Debi_Clientes
    a_imput_anticip_VENtas
    a_Compras
    a_pago_proveedores
    a_Imput_anticip_Compras
    a_Mov_Caja_y_Bancos
    

   

'    DE_CommitTrans
fin:
'    Set rs = Nothing
    grilla.Redraw = True
    Screen.ActiveForm.MousePointer = vbDefault
    Exit Sub
UfaGENE:
    On Error Resume Next
'    DE_RollbackTrans
    ufa "err generando asiento", ""
    Resume fin
End Sub


Private Sub MalAsiento()
    On Error Resume Next
'    DE_RollbackTrans
    che "Err en la generacion de asientos" & vbCrLf & " Concepto " & Asiento.concepto & vbCrLf & " Fehca " & CStr(Asiento.Fecha) & vbCrLf & " dif: " & Asiento.Diferencia & vbCrLf & " Nro Items: " & Asiento.CantItems
End Sub
Private Sub AsientoAGrilla()
    Dim r As Long, i As Long, u As Long
    With Asiento
        If .CantItems = 0 Then Exit Sub ' no hay items
       
        '''''''''''' DEBUG '''''''' HABILITAR AL FINAL
        If s2n(.TotalDebe + .TotalHaber) = 0 Then Exit Sub ' no hay valores
        '''''''''''' DEBUG '''''''' HABILITAR AL FINAL
        
        r = g.addRow()
        g.tx r, G_CONC, .concepto
        For i = 0 To .CantItems - 1
            r = g.addRow()
            g.tx r, G_CONC, .ItemConcepto(i)
            g.tx r, G_CUEN, .ItemCuenta(i)
            g.tx r, G_DEBE, .ItemDebe(i)
            g.tx r, G_HABE, .ItemHaber(i)
        Next i
        r = g.addRow
        g.tx r, G_DEBE, .TotalDebe
        g.tx r, G_HABE, .TotalHaber
        g.tx r, G_DIFE, .Diferencia
        g.addRow
    End With
End Sub
Private Sub inigrilla()
    Set g = New LiGrilla
    g.init grilla
'    G_ASIE = g.AddCol(" Nro ")
    G_CONC = g.AddCol(" Concepto                                             ")
    G_CUEN = g.AddCol(" Cuenta         ")
    G_DEBE = g.AddCol(" Debe              ", "9")
    G_HABE = g.AddCol(" Haber             ", "9")
    G_DIFE = g.AddCol(" -                 ", "9")
End Sub
Private Function aGrillayGraba() As Boolean
    aGrillayGraba = True
    Asiento.ordenar
    AsientoAGrilla
    If Asiento.Grabar(0) = 0 Then
        MalAsiento
        aGrillayGraba = False
    End If
End Function

Private Sub Form_Resize()
    encajar grilla, Me, 1200, 20, 120, 20
End Sub
Private Sub cmdGenerarAsiento_Click()
    g.Borrar
    If PeriodoYaGenerado() Then
        che "Periodo ya generado"
        Exit Sub
    End If
    GenAsiento
End Sub
Private Sub Form_Load()
    lblEjercicio = leerEjercicioDenominacion()
    Set Asiento = New Asiento
    Set Debe = New Cuenta
    Set haber = New Cuenta
    inigrilla
    ucXls1.ini g, ".\Asientos", "asientos "
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Asiento = Nothing
    Set g = Nothing
End Sub

Private Function AsNuevo(concepto As String, letra As String)
    Dim Fecha As Date
    Fecha = uFeHasta.dtFecha
    AsNuevo = Asiento.nuevo(concepto & CStr(Month(Fecha)) & "/" & CStr(Year(Fecha)), Fecha, letra)
    lblVoyPor.caption = concepto
End Function

Private Sub uFeDesde_LostFocus()
    uFeHasta.setUltDiaMes uFeDesde.Mes, uFeDesde.Anio
End Sub

Private Function PeriodoYaGenerado() As Boolean

End Function

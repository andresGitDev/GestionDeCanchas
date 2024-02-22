Attribute VB_Name = "ModAccesorios"
Option Explicit
Public pTotalEnCheques As Double
'mod german
Public revoke_producto()
'fin mod german
'mod li
Public Const USUARIO_SABE_LO_QUE_HACE = True
'Public ON_ERROR_HABILITADO As Boolean
'
'Public Type typeFormula
'    componente As String
'    cantidad As Long
'End Type
'Public ColFormula As Collection

Public Const REMITO_CON_PRECIO = True
'Public Const PRODUCTO_CON_FORMULA_ES_VIRTUAL = True
Public Const CHAR_PROD_VIRTUAL = "V"

' *** ' De Tabla TipoComprobantesGrales :  REMVTA = 5 - REMCPRA = 6
Public Const TipoComprobante_CANCELACIONPEDIDO = 8 ' Cancelpedido-remitoDifStock
Public Const TipoComprobante_REMITOVENTA = 5 ' REMVTA
Public Const TipoComprobante_REMITOCOMPRA = 6 ' REMCPRA
Public Const TipoComprobante_DIFSTOCK = 7 ' Dif Stock
Public Const TipoComprobante_REMITOAJUSTE = 9 ' remito venta ajuste
'Public Const TipoComprobante_FACTURAVENTA_A = 1 ' FAA
'Public Const TipoComprobante_FACTURAVENTA_B = 2
' *** ' De Tabla TipoComprobantesGrales :  REMVTA = 5

' *** Tabla IVAS ***
'Public Const IVA_ConsumidorFinal = 1
'

'Tabla BS -  *******************************************

Public Const TABLA_PARAMETROS = "BS"
Public Const CAMPO_BS_NroREMITO = "NUM_RemitoVenta"
Public Const CAMPO_BS_NroFACTURA_A = "NUM_Factura_A"
Public Const CAMPO_BS_FecFACTURA_A = "FEC_Factura_A"
Public Const CAMPO_BS_NroFACTURA_B = "NUM_Factura_B"
Public Const CAMPO_BS_FecFACTURA_B = "FEC_Factura_B"
'Public Const CAMPO_BS_CodFactura_VENTA = "COD_FacturaVenta"
Public Const CAMPO_BS_OrdenPago = "NUM_opago"
Public Const CAMPO_BS_EJERCICIO = "Ejercicio"
Public Const CAMPO_BS_APC = "Num_APC"
Public Const CAMPO_BS_APD = "Num_APD"
Public Const CAMPO_BS_BaseIIBB = "BasePerIIBB"
'fin mod li

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OpenFilename) As Long

Public Const Operacion_ALTA = "A"
Public Const Operacion_BAJA = "B"
Public Const Operacion_MODIFICACION = "M"

Public Const Estado_PENDIENTE = "PENDIENTE"
Public Const Estado_FACTURADO = "FACTURADO"

Public Const FACTURA_A = 1

Public Enum LlenarGrillaComo
    llenagResetear
    llenagAgregar
End Enum
 
Public Type datCD
    dCodigo As Long
    dDescripcion As String
End Type

Private Type OpenFilename
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    iFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Const INVALID_HANDLE_VALUE = -1
Public Const MAX_PATH = 260

Public Declare Function FindFirstFile Lib "kernel32" _
   Alias "FindFirstFileA" _
  (ByVal lpFileName As String, _
   lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" _
  (ByVal hFindFile As Long) As Long


Public Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

Public Enum PrintF
    pHorizontal
    pVertical
End Enum

Public Function ExisteArchivo(sSource As String) As Boolean
   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   
   hFile = FindFirstFile(sSource, WFD)
   ExisteArchivo = hFile <> INVALID_HANDLE_VALUE
   
   Call FindClose(hFile)

End Function

Public Function PrintG(pGrilla As Control, Optional pForma As PrintF = pVertical, Optional pTitulo As String = "", Optional pFecha As Date = CDate("01/01/1900"), Optional pCabezera As String = "Reporte", Optional hoja As Long = pprA4)
Dim frmImp As New FrmImpresiones
With frmImp
    If pTitulo = "" Then
        .caption = "Impresion de Reporte"
    Else
        .caption = Trim(pTitulo)
    End If
    
    If pGrilla.rows < 2 Then Exit Function
    
    pGrilla.GridLines = flexGridNone
    pGrilla.GridLinesFixed = flexGridNone
    
    If pForma = pVertical Then
        .VSPrinter.Orientation = orPortrait
    Else
        .VSPrinter.Orientation = orLandscape
    End If
    
    If hoja = pprA4 Then
        .VSPrinter.PaperSize = pprA4
    Else
        .VSPrinter.PaperSize = hoja
    End If
    .VSPrinter.Preview = True
    .VSPrinter.Font.Name = pGrilla.Font.Name
    .VSPrinter.FontSize = 8
    If pCabezera = "" Then
        .VSPrinter.Header = ""
    Else
        .VSPrinter.Header = pCabezera
    End If
    .VSPrinter.FontSize = 8
    
    .VSPrinter.StartDoc
    If pFecha = CDate("01/01/1900") Then
    Else
        .VSPrinter.Paragraph = "Fecha : " & pFecha
    End If
    '.VSPrinter.Paragraph = "Control de Pre-liquidacion" 'mTitulo
    '.VSPrinter.Paragraph = "Entre fechas : " & dtFDesde & " - " & dtFHasta
    '.VSPrinter.Paragraph = "Periodo : " & frmLiquidacion.armoPeriodo(dtFHasta)
    .VSPrinter.Paragraph = " "
    
    .VSPrinter.TextAlign = taRightBottom
    .VSPrinter.RenderControl = pGrilla.hWnd
    
    
    .VSPrinter.Footer = "||Pagina %d de " & .VSPrinter.PageCount
    .VSPrinter.Zoom = 100
    .VSPrinter.EndDoc
    
    .Show
    pGrilla.GridLines = flexGridFlat
End With
End Function

Public Function VentanaArchivo(f As Object, Optional ext As String = "", Optional tit As String = "") As String
    Dim ofn As OpenFilename
    ofn.lpstrFile = ""
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = f.hWnd
    If ext > "" Then
        ofn.lpstrFilter = "Tipo (" & Trim(ext) & ")" & Chr$(0) & Trim(ext) & Chr$(0) & Chr(0) & Chr(0)
    Else
        ofn.lpstrFilter = "Todos los archivos (*.rei)" & Chr$(0) & "*.*" & Chr$(0) & Chr(0) & Chr(0)
    End If
    ofn.lpstrFile = String(56, 0)
    ofn.nMaxFile = 255
    If tit > "" Then
        ofn.lpstrTitle = Trim(tit)
    Else
        ofn.lpstrTitle = "Buscar Archivo"
    End If
    ofn.Flags = &H800000 + &H1000 + &H8 + &H4
    ofn.lpstrDefExt = "*" + Chr(0)
    GetOpenFileName ofn
    If Mid(ofn.lpstrFile, 1, 1) <> Chr(0) Then VentanaArchivo = Trim(ofn.lpstrFile)
End Function


Public Function mStripRTF(ByVal strNote As String) As String
    Dim strNewNote As String
    Dim blnPrint As Boolean
    Dim blnParen As Boolean
    blnPrint = True
    blnParen = False
    Dim strPre As String
    
    Dim i As Long
    For i = 1 To Len(strNote)
        If Len(Trim(Mid(strNote, i, 1))) = 0 Then
            blnPrint = True
        End If
        If i > 1 Then
            strPre = Mid(strNote, i - 1, 1)
        End If
        Select Case Mid(strNote, i, 1)
        Case "{"
            If strPre <> "\" Then
                blnParen = True
                blnPrint = False
            Else
                blnParen = False
                blnPrint = True
                strNewNote = strNewNote & Mid(strNote, i, 1)
            End If
        Case "}"
            If strPre <> "\" Then
                blnPrint = False
            Else
                blnPrint = True
                strNewNote = strNewNote & Mid(strNote, i, 1)
            End If
            blnParen = False
        Case "\"
            If strPre <> "\" Then
                blnPrint = False
            Else
                blnPrint = True
                strNewNote = strNewNote & Mid(strNote, i, 1)
            End If
        Case Else
            If blnPrint And Not blnParen Then
                strNewNote = strNewNote & Mid(strNote, i, 1)
            End If
        End Select
    Next i
    
    Dim blnTrimLines As String
    blnTrimLines = False
    Dim strTestTrim As String
    strTestTrim = strNewNote
    
    'Trim Extra spaces within the text.
    Do Until blnTrimLines
        strNewNote = Replace(strNewNote, Chr(32) & Chr(32), Chr(32))
        strNewNote = Replace(strNewNote, Chr(13) & Chr(10) & Chr(32), Chr(13) & Chr(10))
        strNewNote = Replace(strNewNote, Chr(32) & Chr(13) & Chr(10), Chr(13) & Chr(10))
        If strTestTrim <> strNewNote Then
            strTestTrim = strNewNote
        Else
            blnTrimLines = True
        End If
    Loop
    
    'Two spaces after a Period or Colon
    strNewNote = Replace(strNewNote, Chr(46) & Chr(32), Chr(46) & Chr(32) & Chr(32))
    strNewNote = Replace(strNewNote, Chr(58) & Chr(32), Chr(58) & Chr(32) & Chr(32))
    
    mStripRTF = Trim$(strNewNote)
    
End Function

Public Function CORTO(cadena As String, IZQ As Integer, DER As Integer) As String
Dim Car As Long, Iz As Long, de As Long
Car = Len(cadena)
If Car = 0 Then CORTO = "": Exit Function
Iz = Car - DER
de = Iz - IZQ
CORTO = Right(Left(cadena, Iz), de)
End Function

Public Function AverEjercicio()
Dim fff, fi As Date, ff As Date, ejer As Long
Dim newfi As Date, newff As Date, exejer
fff = obtenerDeSQL("select fechainicio, fechafin, ejercicio from ejercicio where activo=1")
If IsNull(fff) Or IsEmpty(fff) Then
    fff = Array(Date - 366, Date - 1, 0)
End If
fi = CDate(fff(0))
ff = CDate(fff(1))
ejer = fff(2) + 1
If Date >= fi And Date <= ff Then
Else
    If Date < fi Then
        MsgBox "La fecha actual esta fuera del rango del Ejercicio." & Chr(13) & "Verifique el ejercicio Actual.", vbExclamation, "Fecha Menor a la fecha inicial del ejercicio."
    ElseIf Date > ff Then
        newfi = CDate(Day(fi) & "/" & Month(fi) & "/" & Year(fi) + 1)
        newff = CDate(Day(ff) & "/" & Month(ff) & "/" & Year(ff) + 1)
        exejer = obtenerDeSQL("select idejercicio from ejercicio where fechainicio=" & ssFecha(newfi) & " and fechafin=" & ssFecha(newff))
        If IsNull(exejer) Or IsEmpty(exejer) Then
            DataEnvironment1.Sistema.Execute "update ejercicio set activo=0"
            DataEnvironment1.Sistema.Execute "insert into ejercicio (Ejercicio,Denominacion,FechaInicio,FechaFin,Activo,Cerrado) " _
                                        & " Values (" & ejer & "," & Year(newfi) & "," & ssFecha(newfi) & "," & ssFecha(newff) & ",1,0) "
        Else
            DataEnvironment1.Sistema.Execute "update ejercicio set activo=0"
            DataEnvironment1.Sistema.Execute "update ejercicio set activo=1 where idejercicio=" & exejer
        End If
    End If
End If
   
End Function


Public Function ExistenciaCalculada(producto As String) As Long
Dim t As Variant, rs As New ADODB.Recordset, depo As String, s As String
   
    t = obtenerDeSQL("select existencia, formula from producto where codigo = '" & producto & "' ")
    If IsNull(t) Or IsEmpty(t) Then
        ExistenciaCalculada = 0
        Exit Function
    ElseIf Not t(1) Then
        ExistenciaCalculada = s2n(t(0))
    Else
        s = "SELECT Min(existencia/cantidad) AS MaxArmados " _
            & " FROM  producto as p INNER JOIN Formulas as f ON p.codigo = f.Componente " _
            & " Where f.codigo = '" & producto & "' "
            
        ExistenciaCalculada = Fix(s2n(obtenerDeSQL(s)))
        If ExistenciaCalculada < 0 Then ExistenciaCalculada = 0
    End If
End Function

Public Function ssTexto(dtexto) As String
    ssTexto = " '" & Trim(dtexto) & "' "
End Function

Public Function siFCAE(sTipo As String, sPuntoTipo As String) As Boolean
Dim sGet As Boolean, sPermitido As Boolean
sGet = obtenerDeSQL("select activocae from datosempresa where idempresa=" & gEMPR_idEmpresa)
siFCAE = sGet
If sGet Then
    sPermitido = obtenerDeSQL("select permito_cae from documentoscae where tipopunto =" & ssTexto(sPuntoTipo) & " and tipo=" & ssTexto(sTipo))
    If sPermitido Then
    Else
        MsgBox "El Documento que intenta asignarle CAE no esta permitido.", vbCritical
        siFCAE = False
    End If
Else
    MsgBox "Asignacion de CAE no habilitada.", vbInformation
End If
End Function

Public Function PuntoVentaTipo(indexx As Long) As String
'0 pre-impresa PI
'1 online OL
'2 webservice WS

If indexx = 0 Then PuntoVentaTipo = "PI"
If indexx = 1 Then PuntoVentaTipo = "OL"
If indexx = 2 Then PuntoVentaTipo = "WS"
If indexx = 3 Then PuntoVentaTipo = "WS2"

End Function

'''Public Function FacturaElectronica(IDFactura As Long, CPERMISO As String) As Boolean
'''Dim CAE As New WSAFIPFE.Factura
'''Dim bResultado As Boolean, sModoFE As Long, sEMPRESA As String
'''Dim CuitEmpresa As String, Certificado As String, Licencia As String, rsFactura As New ADODB.Recordset, Identificador As String, IdentificadorF As String, COMPROBANTE As TipoComprobante
'''Dim FFECHA As Date, fneto As Double, ftotal As Double, FCUIT As String, fTipo As String, fCAE As String, fPrDia As Date, fPuntoVenta As Long, fDetalle As New ADODB.Recordset
'''Dim bCuit As String, bCodFactura As String, bPUNTOVENTA As String, bCAE As String, bFechaCAE As String, bBARRA As String
'''Dim i As Long, fexNroTMP
'''Dim tdoc As String
'''Dim zzz As Double, ttt As Double
'''Dim sumItem As Double, iii As Double, ptt As Double, pIva As Double
'''Dim Tipo3 As Double, Tipo4 As Double, Tipo5 As Double
'''Dim Tipo3base As Double, Tipo4base As Double, Tipo5base As Double
'''FacturaElectronica = False
'''sEMPRESA = "BACIGALUPPI"
'''
'''If IDFactura = 0 Then
'''    MsgBox "No se puede seguir con el proceso, ID de factura invalido.", vbCritical
'''    Exit Function
'''End If
'''
'''CuitEmpresa = Trim(Replace(obtenerDeSQL("select cuitempresa from datosempresa where idempresa=" & gEMPR_idEmpresa), "-", ""))
'''If CuitEmpresa = "" Then
'''    MsgBox "No se puede seguir con el proceso, el CUIT de la empresa no esta establecido.", vbCritical
'''    Exit Function
'''End If
'''If Len(CuitEmpresa) < 11 Then
'''    MsgBox "No se puede seguir con el proceso, el CUIT de la empresa no es valido.", vbCritical
'''    Exit Function
'''End If
'''
'''Certificado = App.Path & "\" & sEMPRESA & ".pfx"
'''If ExisteArchivo(Certificado) Then
'''Else
'''    MsgBox "No se puede seguir con el proceso, el certificado no existe.", vbCritical
'''    Exit Function
'''End If
'''
'''Licencia = App.Path & "\" & sEMPRESA & ".lic"
''''Licencia = ""
'''If ExisteArchivo(Licencia) Then
'''Else
'''    MsgBox "No se puede seguir con el proceso, la licencia no existe.", vbCritical
'''    Exit Function
'''End If
'''
'''fCAE = Trim(sSinNull(obtenerDeSQL("Select cae from facturaventa where codigo=" & IDFactura)))
'''If fCAE = "" Then
'''Else
'''    MsgBox "No se puede seguir con el proceso, la factura ya tiene CAE.", vbCritical
'''    Exit Function
'''End If
'''
''''With CAE
'''bPUNTOVENTA = Trim(sSinNull(obtenerDeSQL("Select puntoventa from facturaventa where codigo=" & IDFactura)))
'''
'''
'''
'''sModoFE = obtenerDeSQL("select licenciacae from datosempresa where idempresa=" & gEMPR_idEmpresa)
'''If sModoFE = 0 Then
'''    bResultado = CAE.iniciar(modoFiscal_Test, CuitEmpresa, Certificado, Licencia)
'''Else
'''    bResultado = CAE.iniciar(modoFiscal_Fiscal, CuitEmpresa, Certificado, Licencia)
'''End If
'''
'''
'''
'''If bResultado Then
'''    If CORTO(Trim(obtenerDeSQL("select tipodoc from facturaventa where codigo=" & IDFactura)), 2, 0) = "E" Then
'''        bResultado = CAE.xObtenerTicketAcceso()
'''    Else
'''        bResultado = CAE.f1ObtenerTicketAcceso()
'''    End If
'''End If
'''
'''
'''
'''
'''
'''    If bResultado Then
'''        rsFactura.Open "Select * from facturaventa where codigo=" & IDFactura, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
'''        tdoc = Trim(rsFactura!TIPODOC)
'''
'''        If CORTO(tdoc, 0, 2) = "N" Then
'''            If CORTO(tdoc, 1, 1) = "C" Then
'''                If CORTO(tdoc, 2, 0) = "A" Then
'''                    COMPROBANTE = TipoComprobante_NotaCreditoA
'''                ElseIf CORTO(tdoc, 2, 0) = "B" Then
'''                    COMPROBANTE = TipoComprobante_NotaCreditoB
'''                ElseIf CORTO(tdoc, 2, 0) = "C" Then
'''                    COMPROBANTE = TipoComprobante_NotaCreditoC
'''                ElseIf CORTO(tdoc, 2, 0) = "E" Then
'''                    COMPROBANTE = TipoComprobante_NotaCreditoExterior
'''                End If
'''            ElseIf CORTO(tdoc, 1, 1) = "D" Then
'''                If CORTO(tdoc, 2, 0) = "A" Then
'''                    COMPROBANTE = TipoComprobante_NotaDebitoA
'''                ElseIf CORTO(tdoc, 2, 0) = "B" Then
'''                    COMPROBANTE = TipoComprobante_NotaDebitoB
'''                ElseIf CORTO(tdoc, 2, 0) = "C" Then
'''                    COMPROBANTE = TipoComprobante_NotaDebitoC
'''                ElseIf CORTO(tdoc, 2, 0) = "E" Then
'''                    COMPROBANTE = TipoComprobante_NotaDebitoExterior
'''                End If
'''            End If
'''        ElseIf CORTO(tdoc, 0, 2) = "F" Then
'''            If CORTO(tdoc, 2, 0) = "A" Then
'''                COMPROBANTE = TipoComprobante_FacturaA
'''            ElseIf CORTO(tdoc, 2, 0) = "B" Then
'''                COMPROBANTE = TipoComprobante_FacturaB
'''            ElseIf CORTO(tdoc, 2, 0) = "C" Then
'''                COMPROBANTE = TipoComprobante_FacturaC
'''            ElseIf CORTO(tdoc, 2, 0) = "E" Then
'''                COMPROBANTE = TipoComprobante_FacturaEmportacion
'''            End If
'''        End If
'''        'fPuntoVenta = s2n(obtenerDeSQL("select puntoventafe from datosempresa where idempresa=" & gEMPR_idEmpresa)) 'obtenerDeSQL("select puntoventafe from datosempresa"))
'''        fPuntoVenta = s2n(bPUNTOVENTA)
'''
'''        If CORTO(Trim(obtenerDeSQL("select tipodoc from facturaventa where codigo=" & IDFactura)), 2, 0) = "E" Then
'''            fexNroTMP = CAE.xFEGetLastCMP(fPuntoVenta, COMPROBANTE)
'''        Else
'''            fexNroTMP = CAE.F1CompUltimoAutorizado(fPuntoVenta, COMPROBANTE)
'''        End If
'''
'''        If s2n(s2n(fexNroTMP) + 1) <> s2n(rsFactura!NroFactura) Then
'''            MsgBox "El ultimo comprobante autorizado es el " & s2n(fexNroTMP) & "." & Chr(13) & " El siguiente por aprobar deberia ser el " & s2n(s2n(fexNroTMP) + 1), vbInformation
'''            Exit Function
'''        End If
'''
'''
'''
'''        If COMPROBANTE = TipoComprobante_FacturaEmportacion Or COMPROBANTE = TipoComprobante_NotaDebitoExterior Or COMPROBANTE = TipoComprobante_NotaCreditoExterior Then
'''            CAE.F1CabeceraCantReg = 1
'''            CAE.indice = 0
'''            CAE.xFecha_cbte = afipFecha(CDate(rsFactura!Fecha))
'''            CAE.xtipo_expo = 1
'''            If Trim(sSinNull(rsFactura!permisoembarque)) = "" Then
'''                If COMPROBANTE = TipoComprobante_NotaCreditoExterior Or COMPROBANTE = TipoComprobante_NotaDebitoExterior Then
'''                    CAE.xPermiso_existenteS = ""
'''                    CAE.xPermisoCantidad = 0
'''                    CAE.xPermisoNoInformar = 1
'''                Else
'''                    CAE.xPermiso_existenteS = "N"
'''                    CAE.xPermisoCantidad = 0
'''                    CAE.xPermisoNoInformar = 1
'''                End If
'''            Else
'''                CAE.xPermiso_existenteS = "S"
'''                CAE.xPermisoCantidad = 1
'''                CAE.xPermisoNoInformar = 0
'''            End If
'''            If Trim(sSinNull(rsFactura!paisfae)) = "" Then
'''                CAE.xDst_cmp = 208
'''            Else
'''                CAE.xDst_cmp = s2n(sSinNull(rsFactura!paisfae))
'''            End If
'''            CAE.xCliente = Trim(rsFactura!RAZONSOCIAL)
'''            CAE.xCuit_pais_clienteS = Replace(Trim(rsFactura!CUIT), "-", "") '"50000000016"
'''            CAE.xDomicilio_cliente = Trim(obtenerDeSQL("select direccion from clientes where codigo=" & s2n(rsFactura!cliente)))
'''            '.xId_impositivo = "PJ54482221-l" ' no es obligatorio si tiene cuit del pais
'''            CAE.xMoneda_idS = Trim(nSinNull(obtenerDeSQL("select codigowsfex from monedas where codigo=" & nSinNull(rsFactura!moneda))))
'''            If CAE.xMoneda_idS = "" Then
'''                CAE.xMoneda_idS = "DOL"
'''            End If
'''            CAE.xMoneda_ctzS = s2n(rsFactura!cotizacion, 3)
'''            CAE.xObs_comerciales = "Sin observaciones"
'''
'''            ttt = CDbl(rsFactura!Total)
'''            zzz = CDbl(rsFactura!cotizacion)
'''
'''            CAE.xImp_total = ttt / zzz
'''            CAE.xImp_total = s2n(CAE.xImp_total, 3, True)
'''            CAE.xForma_pago = Trim(obtenerDeSQL("select descripcion from formaspago where codigo=" & rsFactura!formaPago))
'''            If sSinNull(rsFactura!incoterms) = "" Then
'''                MsgBox "Falta INCOTERMS...", vbCritical
'''                Exit Function
'''            Else
'''                CAE.xIncoTerms = CORTO(Trim(sSinNull(rsFactura!incoterms)), 0, Len(Trim(sSinNull(rsFactura!incoterms))) - 3)
'''            End If
'''            CAE.xIncoTerms_ds = Trim(CORTO(Trim(sSinNull(rsFactura!incoterms)), 3, 0))
'''            CAE.xIdioma_cbte = 1
'''
'''            '1 español
'''            '2 ingles
'''            '3 portugues
'''
'''            'CODIGO Y DESCRIPCION DE PAISES
'''            '.xFEGetPARAM_DST_PAIS
'''            'Dim nContador As Integer
'''            'For nContador = 0 To .xPaisItemCantidad - 1
'''            '     .xIndiceItem = nContador
'''            '     DataEnvironment1.AMR.Execute "INSERT INTO FEPAIS (CODIGO,DESCRIPCION) VALUES (" & .xPais_dst_codigo & "," & ssTexto(.xPais_dst_ds) & ")"
'''            '     Debug.Print Chr(9) & Chr(9) & .xPais_dst_codigo & Chr(9) & Chr(9) & .xPais_dst_ds
'''            'Next
'''
'''
'''            'CUIT DE LOS PAISES
'''            '.xFEGetPARAM_DST_CUIT
'''            'For nContador = 0 To .xCuitItemCantidad - 1
'''            '    .xIndiceItem = nContador
'''            '    DataEnvironment1.AMR.Execute "INSERT INTO FECUIT (CUITPAIS,DESCRIPCION) VALUES (" & ssTexto(.xCuit_dst_cuit) & "," & ssTexto(.xCuit_dst_ds) & ")"
'''                'hoja.Cells(nContador + 2, 1).Value = gFe.xCuit_dst_cuit
'''                'hoja.Cells(nContador + 2, 2).Value = gFe.xCuit_dst_ds
'''            'Next
'''
'''            'TIPOS DE UNIDADES DE MEDIDA
'''            '.xFEGetPARAM_uMed
'''
'''            '  For nContador = 0 To .xUMedItemCantidad - 1
'''            '      .xIndiceItem = nContador
'''            '      Debug.Print .xUMed_Id & " - " & .xUMed_DS & " - " & .xUMed_Vig_desde & " - " & .xUMed_Vig_hasta
'''            '  Next
'''
'''            'MONEDAS
'''            '.xFEGetPARAM_MON
'''            '  For nContador = 0 To .xMonedaItemCantidad - 1
'''            '      .xIndiceItem = nContador
'''            '      Debug.Print .xMonedaId & " - " & .xMonedaDS & " - " & .xMonedaVig_desde & " - " & .xMonedaVig_HASTA
'''            '  Next
'''
'''            '.ArchivoXMLRecibido = "c:\recibido.xml"
'''
'''            fDetalle.Open "select * from facturaventadetalle where codigofactura=" & s2n(rsFactura!codigo), DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
'''            sumItem = 0
'''            If fDetalle.EOF And fDetalle.BOF Then
'''            Else
'''                CAE.xItemCantidad = fDetalle.RecordCount
'''                For i = 0 To fDetalle.RecordCount - 1
'''                    CAE.xIndiceItem = i
'''                    CAE.xITEMPro_codigo = Trim(fDetalle!producto)
'''                    CAE.xITEMPro_ds = Trim(fDetalle!DESCRIPCION)
'''
'''                    CAE.xITEMPro_qty = s2n(fDetalle!cantidad)
'''                    If CAE.xITEMPro_qty = 0 Then
'''                        CAE.xITEMPro_umed = 0 'SIN UNIDAD
'''                    Else
'''                        CAE.xITEMPro_umed = 7 'UNIDADES
'''                    End If
'''                    iii = CDbl(fDetalle!PrecioUnitario)
'''                    CAE.xITEMPro_precio_uni = s2n(iii / zzz, 3, True)
'''                    ptt = CDbl(fDetalle!PrecioTotal)
'''                    CAE.xITEMPro_precio_item = s2n(ptt / zzz, 3, True)
'''                    '.xITEMPro_precio_uni = s2n(fDetalle!PrecioUnitario, 8)
'''                    '.xITEMPro_precio_item = s2n(fDetalle!Preciototal, 8)
'''
'''                    sumItem = sumItem + ptt
'''                    fDetalle.MoveNext
'''                Next
'''                sumItem = sumItem / zzz
'''                CAE.xImp_total = s2n(sumItem, 3, True)
'''            End If
'''
'''            '.xIndiceItem = 0
'''            '.xITEMPro_codigo = "PRO1"
'''            '.xITEMPro_ds = "Producto Tipo 1 Exportacion MERCOSUR ISO 9001"
'''            '.xITEMPro_qty = 1
'''            '.xITEMPro_umed = 7
'''            '.xITEMPro_precio_uni = 250
'''            '.xITEMPro_precio_item = 250
'''
'''            '.xIndiceItem = 1
'''            '.xITEMPro_codigo = "PRO1"
'''            '.xITEMPro_ds = "Producto Tipo 1 Exportacion MERCOSUR ISO 9001"
'''            '.xITEMPro_qty = 1
'''            '.xITEMPro_umed = 7
'''            '.xITEMPro_precio_uni = 250
'''            '.xITEMPro_precio_item = 250
'''
'''
'''
'''            '.xIndiceItem = 0
'''            '.xPERMISO_id_permiso = "09052EC01006154G"
'''            '.xPERMISO_dst_merc = 203
'''
'''            '.xIndiceItem = 1
'''            '.xPERMISO_id_permiso = "09052EC01006154G"
'''            '.xPERMISO_dst_merc = 202
'''
'''            '.xCmps_asocCantidad = 0
'''        Else
'''            fPrDia = CDate("01/" & Month(CDate(rsFactura!Fecha)) & "/" & Year(CDate(rsFactura!Fecha)))
'''
'''            CAE.F1CabeceraCantReg = 1
'''
'''            zzz = CDbl(rsFactura!cotizacion)
'''            If zzz = 0 Then zzz = 1
'''            If zzz = 1 Then
'''                CAE.F1DetalleMonId = "PES"
'''                CAE.F1DetalleMonCotiz = 1
'''            Else
'''                CAE.F1DetalleMonId = "DOL"
'''                CAE.F1DetalleMonCotiz = zzz
'''            End If
'''            CAE.F1CabeceraPtoVta = fPuntoVenta
'''            CAE.F1CabeceraCbteTipo = COMPROBANTE
'''            CAE.f1Indice = 0
'''            CAE.F1DetalleConcepto = 3
'''            'CAE.F1DetalleDocTipo = 80
'''            CAE.F1DetalleDocNro = s2n(Replace(rsFactura!CUIT, "-", ""))
'''
'''
'''
'''            If Len(Trim(CAE.F1DetalleDocNro)) <= 8 Then
'''                CAE.F1DetalleDocTipo = TipoDocumento_DNI
'''            ElseIf Len(Trim(CAE.F1DetalleDocNro)) = 11 Then
'''                CAE.F1DetalleDocTipo = TipoDocumento_CUIT
'''            Else
'''                CAE.F1DetalleDocTipo = TipoDocumento_Pasaporte
'''            End If
'''
'''            If FrmPrincipal.chkDesbichando Then
'''                MsgBox "DocNro: " & CAE.F1DetalleDocNro & Chr(13) & "DocTipo: " & CAE.F1DetalleDocTipo
'''            End If
'''
'''
'''            'CAE.FEDetalleNro_doc = s2n(Replace(rsFactura!Cuit, "-", ""))
'''            CAE.F1DetalleDocNro = s2n(Replace(rsFactura!CUIT, "-", ""))
'''            CAE.F1DetalleCbteDesde = s2n(rsFactura!NroFactura)
'''            CAE.F1DetalleCbteHasta = s2n(rsFactura!NroFactura)
'''            'CAE.FEDetalleFecha_cbte = afipFecha(CDate(rsFactura!fecha))
'''            CAE.F1DetalleCbteFch = afipFecha(CDate(rsFactura!Fecha))
'''            'CAE.FEDetalleImp_total = s2n(rsFactura!Total)
'''            CAE.F1DetalleImpTotal = s2n(rsFactura!Total / zzz)
'''            CAE.F1DetalleImpTotalConc = 0
'''            'CAE.FEDetalleImp_neto = s2n(rsFactura!Neto)
'''            CAE.F1DetalleImpNeto = s2n(rsFactura!Neto / zzz)
'''            CAE.F1DetalleImpOpEx = 0
'''            CAE.F1DetalleImpTrib = 0
'''            CAE.F1DetalleImpIva = s2n(rsFactura!Iva / zzz)
'''
'''            CAE.F1DetalleFchServDesde = afipFecha(fPrDia)
'''            CAE.F1DetalleFchServHasta = afipFecha(ultimoDiaDelMes(fPrDia))
'''            'CAE.F1DetalleFchServDesde = afipFecha(CDate("01/01/" & Year(Date)))
'''            'CAE.F1DetalleFchServHasta = afipFecha(ultimoDiaDelMes(CDate("01/12/" & Year(Date))))
'''
'''            'CAE.F1DetalleFchVtoPago = afipFecha(Date)
'''            CAE.F1DetalleFchVtoPago = afipFecha(CDate(rsFactura!Vencimiento))
'''            ttt = CDbl(rsFactura!Total)
'''
'''
'''
'''
'''
'''            CAE.F1DetalleTributoItemCantidad = 0
'''            CAE.f1IndiceItem = 0
'''            CAE.F1DetalleTributoId = 0
'''            CAE.F1DetalleTributoDesc = ""
'''            CAE.F1DetalleTributoBaseImp = 0
'''            CAE.F1DetalleTributoAlic = 0
'''            CAE.F1DetalleTributoImporte = 0
'''
'''
'''
'''            Set fDetalle = Nothing
'''            fDetalle.Open "select *,_iva as piva from facturaventadetalle where codigofactura=" & s2n(rsFactura!codigo), DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
'''            sumItem = 0
'''            pIva = s2n(rsFactura!PorcentajeIva, 4)
'''            If fDetalle.EOF And fDetalle.BOF Then
'''            Else
'''
'''                'CAE.F1DetalleIvaId = 5
'''                'CAE.F1DetalleIvaBaseImp = 100
'''                'CAE.F1DetalleIvaImporte = 21
'''
'''                'CAE.F1DetalleIvaItemCantidad = fDetalle.RecordCount
'''                Tipo3 = 0
'''                Tipo4 = 0
'''                Tipo5 = 0
'''                Tipo3base = 0
'''                Tipo4base = 0
'''                Tipo5base = 0
'''                For i = 0 To fDetalle.RecordCount - 1
''''                    CAE.f1IndiceItem = i
''''                    'CAE.F1DetalleIvaItemCantidad = Trim(fDetalle!descripcion)
'''
''''                    'CAE.xITEMPro_qty = s2n(fDetalle!cantidad)
''''                    'CAE.xITEMPro_umed = 7 'UNIDADES
'''
'''                    ptt = CDbl(fDetalle!PrecioTotal)
'''                    pIva = s2n(fDetalle!pIva / 100, 4)
'''                    If CORTO(Trim(rsFactura!TIPODOC), 2, 0) = "B" Then
'''                        pIva = 0
'''                    End If
'''                    iii = 0
'''                    If CORTO(Trim(rsFactura!TIPODOC), 2, 0) = "A" Then
'''                        iii = s2n(s2n((ptt * pIva), 3) / zzz, 3)
'''                    ElseIf CORTO(Trim(rsFactura!TIPODOC), 2, 0) = "B" Then
'''                        iii = s2n(s2n(ptt - (ptt / (1 + pIva)), 3) / zzz, 3)
'''                    End If
'''
'''                    If pIva * 100 = 21 Then
'''                        'CAE.F1DetalleIvaId = 5 '5 es 21 , 4 es 10.5
'''                        Tipo5base = Tipo5base + s2n(ptt / zzz, 3, True)
'''                        Tipo5 = s2n(Tipo5 + iii, 3)
'''                    ElseIf pIva * 100 = 10.5 Then
'''                        'CAE.F1DetalleIvaId = 4 '5 es 21 , 4 es 10.5
'''                        Tipo4base = Tipo4base + s2n(ptt / zzz, 3, True)
'''                        Tipo4 = s2n(Tipo4 + iii, 3)
'''                    Else
'''                        'CAE.F1DetalleIvaId = 3
'''                        Tipo3base = Tipo3base + s2n(ptt / zzz, 3, True)
'''                        Tipo3 = s2n(Tipo3 + iii, 3)
'''                    End If
'''
'''                    'CAE.F1DetalleIvaBaseImp = s2n(ptt / zzz, 3, True)
'''
'''
'''                    'iii = CDbl(fDetalle!PrecioUnitario)
'''
'''                    'CAE.xITEMPro_precio_uni = s2n(iii / zzz, 3, True)
''''                    'ptt = CDbl(fDetalle!Preciototal)
''''                    'CAE.xITEMPro_precio_item = s2n(ptt / zzz, 3, True)
''''                    '.xITEMPro_precio_uni = s2n(fDetalle!PrecioUnitario, 8)
''''                    '.xITEMPro_precio_item = s2n(fDetalle!Preciototal, 8)
'''                    sumItem = sumItem + ptt
'''                    fDetalle.MoveNext
'''                Next
'''                CAE.F1DetalleIvaItemCantidad = 0
'''                If s2n(Tipo3base) > 0 Then
'''                    CAE.F1DetalleIvaItemCantidad = CAE.F1DetalleIvaItemCantidad + 1
'''                End If
'''                If s2n(Tipo4base) > 0 Then
'''                    CAE.F1DetalleIvaItemCantidad = CAE.F1DetalleIvaItemCantidad + 1
'''                End If
'''                If s2n(Tipo5base) > 0 Then
'''                    CAE.F1DetalleIvaItemCantidad = CAE.F1DetalleIvaItemCantidad + 1
'''                End If
'''
'''                For i = 0 To CAE.F1DetalleIvaItemCantidad - 1
'''                    CAE.f1IndiceItem = i
'''                    If s2n(Tipo3base) > 0 Then
'''                        CAE.F1DetalleIvaId = 3
'''                        CAE.F1DetalleIvaBaseImp = s2n(Tipo3base, 2)
'''                        CAE.F1DetalleIvaImporte = s2n(Tipo3, 2)
'''                        Tipo3base = 0
'''                    'End If
'''                    ElseIf s2n(Tipo4base) > 0 Then
'''                        CAE.F1DetalleIvaId = 4
'''                        CAE.F1DetalleIvaBaseImp = s2n(Tipo4base, 2)
'''                        CAE.F1DetalleIvaImporte = s2n(Tipo4, 2)
'''                        Tipo4base = 0
'''                    'End If
'''                    ElseIf s2n(Tipo5base) > 0 Then
'''                        CAE.F1DetalleIvaId = 5
'''                        CAE.F1DetalleIvaBaseImp = s2n(Tipo5base, 2)
'''                        CAE.F1DetalleIvaImporte = s2n(Tipo5, 2)
'''                        Tipo5base = 0
'''                    End If
'''
'''                Next
'''
'''                sumItem = sumItem / zzz
''''                'CAE.xImp_total = s2n(sumItem, 3, True)
'''            End If
'''        End If
'''
'''        IdentificadorF = sSinNull(obtenerDeSQL("select identificador from facturaventa where codigo=" & IDFactura))
'''        If IdentificadorF = "" Then
'''            IdentificadorF = CAE.Identificador
'''            DataEnvironment1.Sistema.Execute "update facturaventa set identificador=" & ssTexto(IdentificadorF) & " where codigo=" & IDFactura
'''        Else
'''            CAE.Identificador = IdentificadorF
'''        End If
'''
'''
'''
'''        If Right(Trim(rsFactura!TIPODOC), 1) = "E" Then
'''            'CAE.FEDetalleTipo_doc = TipoDocumento_CIExtranjera
'''            bResultado = CAE.xRegistrar(fPuntoVenta, COMPROBANTE, IdentificadorF)
'''            'bResultado = .xRegistrarConNumero(fPuntoVenta, comprobante, IdentificadorF, s2n(rsFactura!NroFactura))
'''        Else
'''            'CAE.FEDetalleTipo_doc = TipoDocumento_CUIT
'''            'bResultado = CAE.Registrar(fPuntoVenta, comprobante, CAE.Identificador)
'''            'CAE.FEDetalleTipo_doc = TipoDocumento_CUIT
'''            'CAE.F1DetalleDocTipo = TipoDocumento_CUIT
'''            'bResultado = CAE.Registrar(fPuntoVenta, comprobante, CAE.Identificador)
'''            CAE.ArchivoXMLEnviado = "c:\WSFEv1_enviado.xml"
'''            CAE.ArchivoXMLRecibido = "c:\WSFEv1_recibido.xml"
'''            bResultado = CAE.F1CAESolicitar()
'''        End If
'''
'''
'''
'''        'bResultado = .FE.Registrar(1, FacturaA, "")
'''        If COMPROBANTE = TipoComprobante_FacturaEmportacion Or COMPROBANTE = TipoComprobante_NotaDebitoExterior Or COMPROBANTE = TipoComprobante_NotaCreditoExterior Then
'''            If bResultado Then
'''                bCuit = Format(CuitEmpresa, "00000000000")
'''                'bCodFactura = Format(Trim(obtenerDeSQL("select codfacturae from datosempresa where idempresa=" & gEMPR_idEmpresa)), "00")
'''                bPUNTOVENTA = bPUNTOVENTA
'''                bCodFactura = Format(Trim(obtenerDeSQL("select codfactura from documentoscae where tipopunto='WS' and tipo=" & ssTexto(rsFactura!TIPODOC) & " and puntoventa=" & ssTexto(bPUNTOVENTA))), "00")
'''                'bPUNTOVENTA = bPUNTOVENTA 'Format(Trim(obtenerDeSQL("select puntoventa from datosempresa where idempresa=" & gEMPR_idEmpresa)), "0000")
'''
'''                bCAE = CAE.xRespuestaCAE
'''
'''                bFechaCAE = CAE.xRespuestaFch_vence_cae
'''
'''                'bBARRA = bCuit & bCodFactura & bPuntoVenta & bCAE & bFechaCAE & "8" 'MOMENTAMEAMENTE VA 8 FIJO HASTA QUE HABERIGUE DE DONDE SALE
'''                bBARRA = bCuit & bCodFactura & bPUNTOVENTA & bCAE & bFechaCAE
'''                bBARRA = bBARRA & CodVerificador(bBARRA)
'''
'''    '            .FE.FERespuestaDetalleFecha_vto
'''                DataEnvironment1.Sistema.Execute "update facturaventa set barra=" & ssTexto(bBARRA) & ",identificador=" & ssTexto(IdentificadorF) & ",caev=" & ssFecha(aFecha(CAE.xRespuestaFch_vence_cae)) & ", cae=" & ssTexto(CAE.xRespuestaCAE) & ",PermisoEmbarque=" & ssTexto(CPERMISO) & " where codigo=" & IDFactura
'''                'MsgBox ("CAE: " + .FE.FERespuestaDetalleCae + Chr(10) + "MOTIVO: " + .FE.FERespuestaDetalleMotivo + Chr(10) + "PROCESO: " + .FE.FERespuestaReproceso + Chr(10) + "Numero: " + STR(.FE.FERespuestaDetalleCbt_desde))
'''                MsgBox "CAE: " & CAE.xRespuestaCAE, vbInformation
'''                MsgBox "Reproceso: " & CAE.xRespuestaReproceso & Chr(10) & "Comprobante: " & CAE.xRespuestacbte_numeroS & Chr(10) & "Resultado:" & CAE.xRespuestaResultado
'''
'''            Else
'''                'MsgBox ("Motivo: " + Me.FE.FERespuestaDetalleMotivo + Chr(10) + " Error " + Me.FE.Permsg + "Detalle: " + Me.FE.UltimoMensajeError)
'''                MsgBox ("Motivo: " & Chr(10) & " Error " & CAE.Permsg & "Detalle: " & CAE.xerrmsg & " - " & CAE.UltimoNumeroError)
'''                MsgBox ("Reproceso: " & CAE.xRespuestaReproceso & Chr(10) & "Comprobante: " & CAE.xRespuestacbte_numeroS & Chr(10) & "Resultado:" & CAE.xRespuestaResultado)
'''            End If
'''        Else
'''            If bResultado Then
'''                bCuit = Format(CuitEmpresa, "00000000000")
'''                'bCodFactura = Format(Trim(obtenerDeSQL("select codfacturae from datosempresa where idempresa=" & gEMPR_idEmpresa)), "00")
'''                bCodFactura = Format(Trim(obtenerDeSQL("select codfactura from documentoscae where tipopunto='WS' and tipo=" & ssTexto(rsFactura!TIPODOC) & " and puntoventa=" & ssTexto(bPUNTOVENTA))), "00")
'''                bPUNTOVENTA = bPUNTOVENTA 'Format(Trim(obtenerDeSQL("select puntoventa from datosempresa where idempresa=" & gEMPR_idEmpresa)), "0000")
'''                If Trim(CAE.F1RespuestaDetalleCae) = "" Then
'''                    MsgBox "CAE no valido", vbInformation
'''                    If FrmPrincipal.chkDesbichando Then
'''                        With CAE
'''                            MsgBox .F1RespuestaResultado
'''                            MsgBox .F1RespuestaReProceso
'''                            MsgBox .f1ErrorCode1
'''                            MsgBox .f1ErrorMsg1
'''                            MsgBox .UltimoNumeroError
'''                            MsgBox .UltimoMensajeError
'''                            MsgBox .F1RespuestaCantidadReg
'''                            MsgBox .F1RespuestaDetalleResultado
'''                            MsgBox .F1RespuestaDetalleCae
'''                            MsgBox .F1RespuestaDetalleCAEFchVto
'''                            MsgBox .F1RespuestaDetalleObservacionCode1
'''                            MsgBox .F1RespuestaDetalleObservacionMsg1
'''                            MsgBox .F1RespuestaDetalleObservacionItemCantidad
'''                        End With
'''                    End If
'''                    Exit Function
'''                End If
'''                'bCAE = CAE.FERespuestaDetalleCae
'''                bCAE = CAE.F1RespuestaDetalleCae
'''                'bFechaCAE = CAE.FERespuestaDetalleFecha_vto
'''                bFechaCAE = CAE.F1RespuestaDetalleCAEFchVto
'''                'bBARRA = bCuit & bCodFactura & bPuntoVenta & bCAE & bFechaCAE & "8" 'MOMENTAMEAMENTE VA 8 FIJO HASTA QUE HABERIGUE DE DONDE SALE
'''                bBARRA = bCuit & bCodFactura & bPUNTOVENTA & bCAE & bFechaCAE
'''                bBARRA = bBARRA & CodVerificador(bBARRA)
'''                'DataEnvironment1.AMR.Execute "update facturaventa set barra=" & ssTexto(bBARRA) & ",identificador=" & ssTexto(Identificador) & ",caev=" & ssFecha(aFecha(CAE.FERespuestaDetalleFecha_vto)) & ", cae=" & ssTexto(CAE.FERespuestaDetalleCae) & ",PermisoEmbarque=" & ssTexto(CPERMISO) & " where codigo=" & IDFactura
'''                DataEnvironment1.Sistema.Execute "update facturaventa set barra=" & ssTexto(bBARRA) & ",identificador=" & ssTexto(Identificador) & ",caev=" & ssFecha(aFecha(CAE.F1RespuestaDetalleCAEFchVto)) & ", cae=" & ssTexto(CAE.F1RespuestaDetalleCae) & ",PermisoEmbarque=" & ssTexto(CPERMISO) & " where codigo=" & IDFactura
'''                'MsgBox ("CAE: " + .fe.FERespuestaDetalleCae + Chr(10) + "MOTIVO: " + .fe.FERespuestaDetalleMotivo + Chr(10) + "PROCESO: " + .fe.FERespuestaReproceso + Chr(10) + "Numero: " + str(.fe.FERespuestaDetalleCbt_desde))
'''                MsgBox "CAE: " & CAE.F1RespuestaDetalleCae, vbInformation
'''                'MsgBox "Reproceso: " & CAE.FERespuestaReproceso & Chr(10) & "Cantidad de Reg: " & CAE.FERespuestaCantidadReg & Chr(10) & "Resultado:" & CAE.FERespuestaResultado
'''                MsgBox "Resultado global AFIP: " + CAE.F1RespuestaResultado & Chr(13) & "Es reproceso? " + CAE.F1RespuestaReProceso & Chr(13) & "Registros procesados por AFIP: " + str(CAE.F1RespuestaCantidadReg) & Chr(13) & "ERROR genérico global:" + CAE.f1ErrorMsg1, vbInformation
'''
'''            Else
'''                MsgBox "Resultado global AFIP: " + CAE.F1RespuestaResultado & Chr(13) & "Es reproceso? " + CAE.F1RespuestaReProceso & Chr(13) & "Registros procesados por AFIP: " + str(CAE.F1RespuestaCantidadReg) & Chr(13) & "ERROR genérico global:" + CAE.f1ErrorMsg1, vbInformation
'''            End If
'''        End If
'''        FacturaElectronica = bResultado
'''    Else
'''        MsgBox "Resultado global AFIP: " + CAE.F1RespuestaResultado & Chr(13) & "Es reproceso? " + CAE.F1RespuestaReProceso & Chr(13) & "Registros procesados por AFIP: " + str(CAE.F1RespuestaCantidadReg) & Chr(13) & "ERROR genérico global:" + CAE.f1ErrorMsg1, vbInformation
'''    End If
'''
''''end with
'''
'''End Function

Public Function CodVerificador(vBarra As String) As String
Dim i As Long, L As Long
Dim aImpar As Double, aPar As Double, aValor As String, aResultado As String, aVerificador As Long
L = Len(vBarra)
aImpar = 0
For i = 1 To L Step 2
    aValor = Mid(vBarra, i, 1)
    aImpar = aImpar + s2n(aValor)
Next
aImpar = aImpar * 3
aPar = 0
For i = 2 To L Step 2
    aValor = Mid(vBarra, i, 1)
    aPar = aPar + s2n(aValor)
Next
aResultado = s2n(aImpar + aPar)
aVerificador = 10 - s2n(CORTO(aResultado, Len(aResultado) - 1, 0))
CodVerificador = aVerificador
End Function



Public Function NuevoCodigoMixto(nProv As Long, nFecha As Date)
Dim i As Long, nEx
Dim nYear As String, nMonth As String, nNumero As String, nNuevo As String
nYear = Right(Year(nFecha), 2)
nMonth = Format(Month(nFecha), "00")
For i = 1 To 9999
    nNumero = Format(i, "0000")
    nNuevo = nYear & nMonth & nNumero
    'nEx = obtenerDeSQL("select codmixto from transcom where codpr=" & nProv & " and codmixto=" & ssTexto(nNuevo))
    nEx = obtenerDeSQL("select codmixto from transcom where codmixto=" & ssTexto(nNuevo))
    If IsNull(nEx) Or IsEmpty(nEx) Then
        'nEx = obtenerDeSQL("select codmixto from compras where codpr=" & nProv & " and codmixto=" & ssTexto(nNuevo))
        nEx = obtenerDeSQL("select codmixto from compras where codmixto=" & ssTexto(nNuevo))
        If IsNull(nEx) Or IsEmpty(nEx) Then
            Exit For
        End If
    End If
Next
NuevoCodigoMixto = nNuevo
End Function


Public Function SoloNum(KeyAscy As Integer, Optional ConComa As Boolean = True, Optional Neg As Boolean = True) As Integer
'    If (KeyAscy < 47 Or KeyAscy > 57) And (KeyAscy <> 46 And KeyAscy <> 44) And (KeyAscy <> 7 And KeyAscy <> 8) Then
'        soloNum = 0
'    Else
'        soloNum = KeyAscy
'    End If
    If (KeyAscy < 47 Or KeyAscy > 57) And (KeyAscy <> 46 And KeyAscy <> 44) And (KeyAscy <> 7 And KeyAscy <> 8) Then
        SoloNum = 0
    Else
        SoloNum = KeyAscy
    End If
    If ConComa = False Then
        If KeyAscy = 44 Or KeyAscy = 46 Then SoloNum = 0
    End If
    If Neg Then
        If KeyAscy = 45 Then SoloNum = KeyAscy
    End If

End Function



'*********20/4/07*****CONVERTIDOR DE NUMEROS***RAUL
Function aPositivo(ByVal valorN As Double) As Double 'esto convierte todo numero negativo a positivo
        If (valorN < 0) Then
            aPositivo = s2n(valorN - (valorN * 2))
        End If
End Function

'********23/4/07*****CONVERTIDOR DE NUMERO 2 **RAUL
Function aNegativo(ByRef valorP As Double) As Double 'esto convierte todo numero positivo a negativo
        If (valorP > 0) Then
            aNegativo = s2n(valorP - (valorP * 2))
        End If
End Function

Function CambiarSigno(ByVal valorX As Double) As Double 'esto cambia el signo del numero cualquiera sea
        If (valorX > 0) Then
            CambiarSigno = -(valorX)
        End If
        If (valorX < 0) Then
            CambiarSigno = -(valorX)
        End If
End Function

Public Function ultimoDiaDelMes(Fecha As Date) As Date
ultimoDiaDelMes = DateAdd("m", 1, Fecha)
ultimoDiaDelMes = DateSerial(Year(ultimoDiaDelMes), Month(ultimoDiaDelMes), 1)
ultimoDiaDelMes = DateAdd("d", -1, ultimoDiaDelMes)
End Function




'*********19/4/07*****VERIFICADOR DE DATOS******RAUL
Public Function Verificar_Dato(Dat As Variant, Mode As Integer) As Variant 'esto sirve para cuando un dato esta vacio
Dim datoV As Variant
datoV = ""
    Select Case Mode: 'a ninguno le pongo un valor por defecto por que hay que mostrarle al usuario que lo ingreso mal
        Case 1: 'verifica un entero vacio
            If Dat = "" Or VarType(Dat) = vbNull Then
                datoV = "Sin Numero"
            Else
                datoV = Dat
            End If
        Case 2: 'verifica una fecha vacio
            If Dat = "" Or VarType(Dat) = vbNull Then
                datoV = "Sin Fecha"
            Else
                datoV = Dat
            End If
        Case 3: 'verifica un string vacio
            If Dat = "" Or VarType(Dat) = vbNull Then
                datoV = "Sin Datos"
            Else
                datoV = Dat
            End If
        Case 4: 'verifico un precio vacio
            If Dat = "" Or VarType(Dat) = vbNull Then
                datoV = "Sin Valor"
            Else
                datoV = Dat
            End If
    End Select
Verificar_Dato = datoV
End Function

'*****************EJERCICIOS SIN CERRAR***********************12/4/07****raul
Function TraerEjerSinCerrar() As String
Dim i As Integer
Dim rsEjerSinCerrar As New ADODB.Recordset
Dim cerrar, sql As String

        sql = "SELECT * from ejercicio where Cerrado =0"
        rsEjerSinCerrar.Open sql, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
        cerrar = " "
        For i = 1 To rsEjerSinCerrar.RecordCount Step 1
            If cerrar = " " Then
                cerrar = cerrar & rsEjerSinCerrar!idejercicio
            Else
                cerrar = cerrar & " or (Ejercicio.idEjercicio) = " & rsEjerSinCerrar!idejercicio
            End If
            rsEjerSinCerrar.MoveNext
        Next i
Set rsEjerSinCerrar = Nothing
TraerEjerSinCerrar = cerrar
End Function

'****************MODO DE ACCESO AL PROGRAMA********************7/5/07******RAUL

Function modoDacceso(acceso As String) As Boolean 'esto es para saber como y quien accede, por ahora solo se usa cuando ingresa el contador
    If acceso > "" Then
        FrmPrincipal.lblModo.Visible = True
        FrmPrincipal.lblModo.caption = acceso
        FrmPrincipal.lblModo.Width = Len(FrmPrincipal.lblModo) * 130
        FrmPrincipal.lblModo.Top = 440
        FrmPrincipal.lblModo.Left = 3200
    Else
        FrmPrincipal.lblModo.Visible = False
    End If
End Function
'fin mod raul

'mod sebastian


Public Function ObtenerDatoDB(tabla As String, ColumnaABuscar As String, DatoABuscar, ColumnaADevolver As String) As Variant
'Busca DATOABUSCAR en la COLUMNAABUSCAR en TABLA me devuelve el contenido de la COLUMNAADEVOLVER

Dim Consulta As String
Dim rsaux As New ADODB.Recordset

    If tabla <> "" And ColumnaABuscar <> "" And ColumnaADevolver <> "" And DatoABuscar <> "" Then
        If IsNumeric(DatoABuscar) Then
            Consulta = "Select " & ColumnaADevolver & " From " & tabla & " Where " & ColumnaABuscar & " = " & DatoABuscar
        Else
            Consulta = "Select " & ColumnaADevolver & " From " & tabla & " Where " & ColumnaABuscar & " = '" & DatoABuscar & "'"
        End If
        rsaux.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rsaux.EOF Then ObtenerDatoDB = rsaux.Fields(0)
        rsaux.Close
        Set rsaux = Nothing
        
    End If
End Function

Public Function ObtenerDatoDB2(tabla As String, ColumnaABuscar As String, DatoABuscar, ColumnaADevolver As String) As Variant
'Busca DATOABUSCAR en la COLUMNAABUSCAR en TABLA me devuelve el contenido de la COLUMNAADEVOLVER

Dim Consulta As String
Dim rsaux As New ADODB.Recordset

    If tabla <> "" And ColumnaABuscar <> "" And ColumnaADevolver <> "" And DatoABuscar <> "" Then
        If IsNumeric(DatoABuscar) Then
            Consulta = "Select " & ColumnaADevolver & " From " & tabla & " Where Activo=1 and " & ColumnaABuscar & " = " & DatoABuscar
        Else
            Consulta = "Select " & ColumnaADevolver & " From " & tabla & " Where Activo=1 and " & ColumnaABuscar & " = '" & DatoABuscar & "'"
        End If
        rsaux.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rsaux.EOF Then ObtenerDatoDB2 = rsaux.Fields(0)
        rsaux.Close
        Set rsaux = Nothing
        
    End If
End Function

Public Sub LimpiarGrilla(grilla As Control, Optional Filas As Long = 2, Optional Columnas As Long = 2)
    grilla.clear
    grilla.rows = Filas
    grilla.cols = Columnas
End Sub


Public Function LlenarGrilla(grilla As Control, ConsultaSQL As String, AjustarAnchos As Boolean, Optional nColCorte, Optional nColSum, Optional llenacomo As LlenarGrillaComo = llenagResetear) As Boolean
    ' agregado corte, no implementada la suma aun
    'agregado col invisible, alias empieza con "_H_"     ejnombre = "_H_idRegistro"
    Dim rsaux As New ADODB.Recordset
    Dim C As Long
    Dim Encabezado As String
    Dim ConCorte As Boolean, ColCorte As Long, ColTMP()  ' todo para corte
    

    ColCorte = s2n(nColCorte)
    ConCorte = Not IsMissing(nColCorte) And ColCorte >= 0
    
    
    If ConsultaSQL <> "" Then
        If llenacomo = llenagResetear Then grilla.clear: grilla.rows = 1
        
        rsaux.Open ConsultaSQL, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        If Not rsaux.EOF Then
            With rsaux
                If ColCorte > .Fields.Count Then ConCorte = False
                
                
              If llenacomo = llenagResetear Then
                'hago el encabezado de la grilla
                grilla.FixedCols = 0
                grilla.cols = .Fields.Count
                grilla.Row = 0
                For C = 0 To .Fields.Count - 1
                    If C > 0 Then Encabezado = Encabezado & "|"
                    Encabezado = Encabezado & .Fields(C).Name
                    
                    grilla.TextMatrix(0, C) = .Fields(C).Name
                Next C
'                grilla.FormatString = Encabezado
                
                 'modifico los anchos de las columnas
                If AjustarAnchos Then
                    grilla.Row = 0
                    For C = 0 To .Fields.Count - 1
                            Select Case .Fields(C).Type
                                Case adVarChar, adChar  '200
                                    grilla.ColWidth(C) = 3000
                                Case adInteger
                                    grilla.ColWidth(C) = 1000
                                Case adDouble
                                    grilla.ColWidth(C) = 2000
                                Case adDate
                                    grilla.ColWidth(C) = 1200
                                Case adBoolean
                                    grilla.ColWidth(C) = 200
                                Case Else
                                    grilla.ColWidth(C) = 1000
                            End Select
                    Next C
                End If
                ' oculto columnas
                For C = 0 To .Fields.Count - 1
                        If Left(.Fields(C).Name, 3) = "_H_" Then grilla.ColHidden(C) = True
                        'grilla.ColWidth(C) =
                Next C
                
                grilla.rows = 1
              
              End If
              
                'lleno la grilla con los datos de la consulta
                grilla.cols = .Fields.Count

                While Not .EOF
                    grilla.rows = grilla.rows + 1
                    If ConCorte And grilla.rows > 2 And grilla.TextMatrix(grilla.rows - 2, ColCorte) <> CStr(.Fields(ColCorte)) Then
                        grilla.rows = grilla.rows + 1
                    End If
                    For C = 0 To .Fields.Count - 1
                        grilla.TextMatrix(grilla.rows - 1, C) = IIf(IsNull(.Fields(C)), "", .Fields(C))
                    Next C
                    .MoveNext
                Wend
            End With
            LlenarGrilla = True
        Else
            LlenarGrilla = False
        End If
    End If
    
    
    Set rsaux = Nothing
End Function

Public Function LlenarGrilla2(grilla As Control, ConsultaSQL As String, consultasql2 As String, consultasql3 As String, AjustarAnchos As Boolean, Optional nColCorte, Optional nColSum, Optional llenacomo As LlenarGrillaComo = llenagResetear) As Boolean
    ' agregado corte, no implementada la suma aun
    'agregado col invisible, alias empieza con "_H_"     ejnombre = "_H_idRegistro"
    Dim rsaux As New ADODB.Recordset
    Dim C As Long
    Dim Encabezado As String
    Dim ConCorte As Boolean, ColCorte As Long, ColTMP()  ' todo para corte
    

    ColCorte = s2n(nColCorte)
    ConCorte = Not IsMissing(nColCorte) And ColCorte >= 0
    
    
    If ConsultaSQL <> "" Then
        If llenacomo = llenagResetear Then grilla.clear: grilla.rows = 1
        
        rsaux.Open ConsultaSQL, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        If Not rsaux.EOF Then
            With rsaux
                If ColCorte > .Fields.Count Then ConCorte = False
                
                
              If llenacomo = llenagResetear Then
                'hago el encabezado de la grilla
                grilla.FixedCols = 0
                grilla.cols = .Fields.Count
                grilla.Row = 0
                For C = 0 To .Fields.Count - 1
                    If C > 0 Then Encabezado = Encabezado & "|"
                    Encabezado = Encabezado & .Fields(C).Name
                    
                    grilla.TextMatrix(0, C) = .Fields(C).Name
                Next C
'                grilla.FormatString = Encabezado
                
                 'modifico los anchos de las columnas
                If AjustarAnchos Then
                    grilla.Row = 0
                    For C = 0 To .Fields.Count - 1
                            Select Case .Fields(C).Type
                                Case adVarChar, adChar  '200
                                    grilla.ColWidth(C) = 3000
                                Case adInteger
                                    grilla.ColWidth(C) = 1000
                                Case adDouble
                                    grilla.ColWidth(C) = 2000
                                Case adDate
                                    grilla.ColWidth(C) = 1200
                                Case adBoolean
                                    grilla.ColWidth(C) = 200
                                Case Else
                                    grilla.ColWidth(C) = 1000
                            End Select
                    Next C
                End If
                ' oculto columnas
                For C = 0 To .Fields.Count - 1
                        If Left(.Fields(C).Name, 3) = "_H_" Then grilla.ColHidden(C) = True
                        'grilla.ColWidth(C) =
                Next C
                
                grilla.rows = 1
              
              End If
              
                'lleno la grilla con los datos de la consulta
                grilla.cols = .Fields.Count

                While Not .EOF
                    grilla.rows = grilla.rows + 1
                    If ConCorte And grilla.rows > 2 And grilla.TextMatrix(grilla.rows - 2, ColCorte) <> CStr(.Fields(ColCorte)) Then
                        grilla.rows = grilla.rows + 1
                    End If
                    For C = 0 To .Fields.Count - 1
                        grilla.TextMatrix(grilla.rows - 1, C) = IIf(IsNull(.Fields(C)), "", .Fields(C))
                    Next C
                    .MoveNext
                Wend
                
                Set rsaux = Nothing
                If consultasql2 <> "" Then
                    rsaux.Open consultasql2, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                    If rsaux.EOF = True And rsaux.BOF = True Then
                    Else
                        rsaux.MoveFirst
                        While Not rsaux.EOF
                            grilla.rows = grilla.rows + 1
                            If ConCorte And grilla.rows > 2 And grilla.TextMatrix(grilla.rows - 2, ColCorte) <> CStr(rsaux.Fields(ColCorte)) Then
                                grilla.rows = grilla.rows + 1
                            End If
                            For C = 0 To .Fields.Count - 1
                                grilla.TextMatrix(grilla.rows - 1, C) = IIf(IsNull(rsaux.Fields(C)), "", rsaux.Fields(C))
                            Next C
                            rsaux.MoveNext
                        Wend
                    End If
                End If
                Set rsaux = Nothing
                If consultasql3 <> "" Then
                    rsaux.Open consultasql3, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                    If rsaux.EOF = True And rsaux.BOF = True Then
                    Else
                        rsaux.MoveFirst
                        While Not rsaux.EOF
                            grilla.rows = grilla.rows + 1
                            If ConCorte And grilla.rows > 2 And grilla.TextMatrix(grilla.rows - 2, ColCorte) <> CStr(rsaux.Fields(ColCorte)) Then
                                grilla.rows = grilla.rows + 1
                            End If
                            For C = 0 To .Fields.Count - 1
                                grilla.TextMatrix(grilla.rows - 1, C) = IIf(IsNull(rsaux.Fields(C)), "", rsaux.Fields(C))
                            Next C
                            rsaux.MoveNext
                        Wend
                    End If
                End If
                
            End With
            LlenarGrilla2 = True
        Else
            LlenarGrilla2 = False
        End If
    End If
    
    
    Set rsaux = Nothing
End Function

Public Function ExisteDato(tabla As String, Columna As String, DatoABuscar As Variant) As Boolean
Dim rsaux As New ADODB.Recordset
Dim Consulta As String

    If IsNumeric(DatoABuscar) Then
        Consulta = "Select * From " & tabla & " Where " & Columna & " = " & DatoABuscar
    Else
        Consulta = "Select * From " & tabla & " Where " & Columna & " = '" & DatoABuscar & "'"
    End If
    rsaux.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    ExisteDato = Not rsaux.EOF
    rsaux.Close
    Set rsaux = Nothing
    
End Function
' fin mod sebastian
'mod german


Public Function cargarPos(campo As String, impresionDe As String, propiedad As String) As Long 'pido el registro campo, el tipo de impresion y la propiedad que es el campo
    Dim str As String
    Dim RSimprime As New ADODB.Recordset
    If Not propiedad = "" Then
        str = "select " & propiedad & " from posicionar where nombre='" & campo & "' and imprecionde='" & impresionDe & "'"
        
        RSimprime.Open str, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
        If Not RSimprime.Fields(0).Value = "" Or Not IsNull(RSimprime.Fields(0).Value) Then
            cargarPos = RSimprime.Fields(0)
        Else
        '    MsgBox "No se ha encontrado dato para " & propiedad
            cargarPos = 0
        End If
    End If
End Function
Public Function cargarColor(campo As String, impresionDe As String, propiedad As String) As String 'pido el registro campo, el tipo de impresion y la propiedad que es el campo
    Dim str As String
    Dim RSimprime As New ADODB.Recordset
    If Not propiedad = "" Then
        str = "select " & propiedad & " from posicionar where nombre='" & campo & "' and imprecionde='" & impresionDe & "'"
        
        RSimprime.Open str, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
        If Not RSimprime.Fields(0).Value = "" Or Not IsNull(RSimprime.Fields(0).Value) Then
            cargarColor = RSimprime.Fields(0)
        Else
        '    MsgBox "No se ha encontrado dato para " & propiedad
            cargarColor = 0
        End If
    End If
End Function
Public Function PasarTamano(Valor As Long) As Long 'paso de milimetro a twips
    'If Not valor = "" And
    If IsNumeric(Valor) Then
        'PasarTamano = sSinNull((1440 * Valor) / 254)  'paso una medida en milimetros ing en la base a twips
        PasarTamano = sSinNull(56.6929133858 * Valor)
    'Else
    '    MsgBox "No se ha encontrado ningun valor para convertir."
    End If
End Function
'Public Function Actualizar(ByVal cadena As String) As Boolean
'    Set comm = New ADODB.Command
'    comm.ActiveConnection DataEnvironment1.Sistema
'    comm.CommandType = adCmdStoredProc
'    comm.CommandText = cadena
'    comm.Parameters.Refresh
'    For i = 0 To comm.Parameters.Count - 1
'        comm.Parameters(i).Value = Array(i + 1)
'    Next
'    comm.Execute , , adExecuteNoRecords
'End Function
Public Function PasoColor(ByVal color As String) As String
    If color = "Transparente" Then
        PasoColor = "0" '"ddBKTransparent"
        'rptFactura.Senor.BackStyle = ddBKTransparent
    ElseIf color = "Normal" Then
        PasoColor = "1" '"ddBKNormal"
        'rptFactura.Senor.BackStyle = ddBKNormal
    End If
End Function

Public Function PasoMes(ByVal Mes As Long) As String
    Select Case Mes
        Case 1
            PasoMes = "Enero"
        Case 2
            PasoMes = "Febrero"
        Case 3
            PasoMes = "Marzo"
        Case 4
            PasoMes = "Abril"
        Case 5
            PasoMes = "Mayo"
        Case 6
            PasoMes = "Junio"
        Case 7
            PasoMes = "Julio"
        Case 8
            PasoMes = "Agosto"
        Case 9
            PasoMes = "Septiembre"
        Case 10
            PasoMes = "Octubre"
        Case 11
            PasoMes = "Noviembre"
        Case 12
            PasoMes = "Diciembre"
    End Select
End Function

Public Function Posicionar(Tipo As Boolean) As Boolean
    If Tipo = "0" Then 'remito
        seniorRemito
        direccionRemito
        CUITremito
        DiaRemito
        MesRemito
        AnoRemito
        MesesRemito
        Transporte
        presupuestoRemito
        ordenCompraRemito
        AtencionRemito
        LocalidadRemito
        FacturaRemito
        FechaRemito
        ComprobanteRemito
        IvaRemito
        ReferenciaRemito
        NroReferenciaRemito
        NroProvinciaRemito
        TacharRemito
        cantidadRemito
        ArticuloRemito
        DescripcionRemito
    
    Else 'Factura
        senior
        direccion
        CUIT
        Dia
        Mes
        Ano
        Meses
        Iva
        condicion
        bruto
        cliente
        presupuesto
        ordenCompra
        Debe
        producto
        NumeroRemito
        ResponsableInsc
        ResponsableNoInsc
        Localidad
        Factura
        Provincia
        Pais
        Fecha
        Postal
        cantidad
        Articulo
        DESCRIPCION
        Unitario
        PrecioTotal
        subtotal
        Impuesto
        Subtotal2
        ivainscripto
        ivaNoinscripto
        Total
        Descuento
        DescuentoP
        IvaIn
        IvaInP
        IIBB
        IIBBP
    End If
End Function

'fin mod german

'mod li
'' ***********  Comprobantes ***********
'
Public Function TipoFormVenta(codigoIva) As String
    If ON_ERROR_HABILITADO Then On Error GoTo fin
    TipoFormVenta = obtenerDeSQL("select letra from ivas where codigo = " & codigoIva)
fin:
    If TipoFormVenta = "" Then ufa "err: Formulario de tipo iva no definido ", "TipoFormVenta: ivas =" & codigoIva ', Err
End Function

' **************************************

' ************* Producto ************
Public Function rsFormulaComponentes(productoBase As String) As ADODB.Recordset
    'OJO cerrar rs donde lo llama
    If Not DE_EstaAbierto Then DataEnvironment1.Sistema.Open
    Set rsFormulaComponentes = New ADODB.Recordset
    rsFormulaComponentes.Open "select Componente, cantidad from Formulas where activo = 1 and codigo = '" & productoBase & "'", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
End Function
' Codigo de producto PROPIO
Public Function VerProductoCliente(codigo As String, codPropio As Boolean, codClie As Long)
    On Error Resume Next
    VerProductoCliente = ""
    
    If codigo = "" Then
        VerProductoCliente = ""
    ElseIf codPropio Then
        VerProductoCliente = codigo
    Else
        VerProductoCliente = obtenerDeSQL("select productoCliente from Relacion_Producto_Cliente where producto = '" & codigo & "' and cliente = '" & codClie & "' ")
    End If
'    VerProductoCliente = IIf(codPropio, obtenerDeSQL("select productoCliente from Relacion_Producto_Cliente where producto = '" & codigo & "'"), codigo)
End Function
Public Function VerProductoMio(codigo, Optional codPropio As Boolean)
    On Error Resume Next
    VerProductoMio = ""
    If codPropio Then
        VerProductoMio = codigo
    Else
        VerProductoMio = obtenerDeSQL("select producto from Relacion_Producto_Cliente where productoCliente = '" & codigo & "'")
    End If
End Function
Public Function ProductoConSerie(COD As String, Optional bPropio As Boolean = True) As Boolean
    On Error Resume Next
    Dim conSerie 'variant
    conSerie = obtenerDato("Producto", "'" & VerProductoMio(COD, bPropio) & "'", "serie")
    ProductoConSerie = conSerie
End Function

Public Function EsProductoVirtual(ProdMio As String) As Boolean
    Dim tmp
    EsProductoVirtual = False
    'If PRODUCTO_CON_FORMULA_ES_VIRTUAL Then
    If gEMPR_FormulaEsVirtual Then
        tmp = obtenerDeSQL("select codigo from formulas where activo = 1 and codigo = '" & ProdMio & "'")
        EsProductoVirtual = Not (IsEmpty(tmp))
    End If
End Function

' ***************************************

'************** Stock  - Deposito ********************
Public Function HayProducto(codigo, codDeposito)
    HayProducto = obtenerDeSQL("select " & DepositoCod2Campo(codDeposito) & " from producto where codigo = '" & codigo & "'")
End Function
'
'
Public Function DepositoCod2Campo(COD)
    Dim t As Variant
    t = Array("existencia", "dep1", "dep2", "dep3", "dep4")
    DepositoCod2Campo = t(COD)
End Function
'
'************************************

Public Function AyudaProducto(codCliente As Long, codPropio As Boolean)
    If codPropio Then
        frmBuscar.MostrarSql "select codigo as [ Producto             ], alias as [ Alias               ], descripcion  as [ Descripcion                                              ] from producto where activo = 1"
    Else
        frmBuscar.MostrarSql "" _
            & " select relacion_producto_cliente.productoCliente as [ Producto             ], producto.codigo as [ Codigo Interno       ] , producto.descripcion as [ Descripcion                                 ] ,relacion_producto_cliente.Precio " _
            & " from producto  " _
            & " inner join relacion_Producto_Cliente " _
            & " on producto.codigo = relacion_Producto_cliente.producto " _
            & " where cliente = " & codCliente _
            & " and producto.activo = 1 and relacion_producto_cliente.activo = 1 " _
            & " order by producto"
    End If
    AyudaProducto = frmBuscar.resultado()
End Function

Public Function DescripcionProducto(cual As String) As String
    DescripcionProducto = sSinNull(obtenerDeSQL("select descripcion from producto where codigo = '" & cual & "' and activo = 1 "))
End Function


' **************** SQL ***************************
Public Function obtenerParametro(cual) As Variant 'As long
    Dim rs As New ADODB.Recordset, ssql  As String
    
    ssql = "select " & cual & " from " & TABLA_PARAMETROS
    rs.Open ssql, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    obtenerParametro = rs.Fields(0)
    
    Set rs = Nothing
End Function
Public Function obtenerParametroDE(cual) As Variant 'As long
    Dim rs As New ADODB.Recordset, ssql  As String
    
    ssql = "select " & cual & " from " & TABLA_PARAMETROS
    rs.Open ssql, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    obtenerParametroDE = rs.Fields(0)
    
    Set rs = Nothing
End Function
Public Function obtenerParametroConDefault(cual, queDefault) As Variant 'As long
    On Error GoTo ufaChe
    Dim rs As New ADODB.Recordset, ssql  As String
    
    ssql = "select " & cual & " from " & TABLA_PARAMETROS
    rs.Open ssql, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    obtenerParametroConDefault = rs.Fields(0)
    
    Set rs = Nothing
    Exit Function
ufaChe:
    ' si fallo por algo, mando default
    obtenerParametroConDefault = queDefault
End Function

Public Function AumentarParametroN(cual, nuevo) As Boolean
    Dim viejo
    
    viejo = obtenerParametro(cual)
    
    If viejo > nuevo Then
        ufa "PrgErr: Intento grabar numero menor", "Maybe UserFault: AumentarParametro" ', Err
        Exit Function
    End If
    
    DataEnvironment1.Sistema.Execute "update bs set " & cual & " = " & nuevo
'     daTaenvironment1.Sistema.Execute "update bs set " & cual & " = " & Nuevo
    AumentarParametroN = True
End Function
Public Function AumentarParametroD(cual As String, nuevo As Date) As Boolean
    Dim viejo
    
    viejo = obtenerParametro(cual)
    
    If viejo > nuevo Then
        ufa "PrgErr: Intento grabar fecha menor", "aum parametro" ', Err
        Exit Function
    End If
    
    DataEnvironment1.Sistema.Execute "update bs set " & cual & " = " & ssFecha(nuevo)
'    daTaenvironment1.Sistema.Execute "update bs set " & cual & " = " & ssFecha(Nuevo)
    AumentarParametroD = True
End Function
Public Function CambiarParametroS(cual As String, nuevo As String) As Boolean
    Dim viejo As String
    
    viejo = obtenerParametro(cual)
    
    If viejo = "" Then
        ufa "PrgErr: Intento grabar un vacio", "aum parametro" ', Err
        Exit Function
    End If
    
    DataEnvironment1.Sistema.Execute "update " & TABLA_PARAMETROS & " set " & cual & " = '" & nuevo & "'"
    CambiarParametroS = True
End Function
Public Function CambiarParametroN(cual As String, nuevo As String) As Boolean
    Dim viejo
    
    viejo = obtenerParametro(cual)
    
    If IsNull(viejo) Then
        ufa "", "camb parametro n" & cual & nuevo ', Err
        If Not confirma("dato previo vacio - Grabo?") Then Exit Function
    End If
    
    DataEnvironment1.Sistema.Execute "update " & TABLA_PARAMETROS & " set " & cual & " = " & nuevo
    CambiarParametroN = True
End Function

Public Function YaEstaRecibo(numero) As Boolean
    Dim tmp
    numero = s2n(numero) ' por las dudas, si es string
    
    tmp = obtenerDeSQL("select cliente from recibos where activo = 1 and numero = " & numero)
    If Not IsEmpty(tmp) Then
        YaEstaRecibo = True
    End If
    tmp = obtenerDeSQL("select cliente from facturaVenta where activo = 1 and tipodoc = '" & TipoDoc_RECIBO & "' and NroFactura = " & numero)
    If Not IsEmpty(tmp) Then
        YaEstaRecibo = True
    End If
End Function

' *************************************************

Public Function nuevoCodigo(TablaDE As String, Optional cpoCodigo As String, Optional whe As String) As Long
    'Dim rs As New ADODB.Recordset
    Dim ssql As String, neww

    If cpoCodigo = "" Then cpoCodigo = "codigo"

    ssql = "Select max (" & cpoCodigo & ")  as NN From " & TablaDE
    If whe > "" Then ssql = ssql & " where " & whe
    
    neww = obtenerDeSQL(ssql)
    If IsNull(neww) Or IsEmpty(neww) Then
        nuevoCodigo = 1
    Else
        nuevoCodigo = neww + 1
    End If
End Function
'Public Function nuevoCodigoDB(TablaDB As String, Optional cpocodigo As String, Optional whe As String) As long
'    Dim rs As New ADODB.Recordset
'    Dim sSql As String
'
'    If cpocodigo = "" Then cpocodigo = "codigo"
'
'    sSql = "Select max (" & cpocodigo & ")  as NN From " & TablaDB
'    If whe > "" Then sSql = sSql & " where " & whe
'
'    DE_abrir
'    rs.Open sSql, DataEnvironment1.Sistema, adOpenDynamic, adLockReadOnly
'
'    If Not rs.EOF Then
'        If IsNull(rs.Fields("NN")) Then
'            nuevoCodigoDB = 1
'        Else
'            nuevoCodigoDB = rs.Fields("NN") + 1
'        End If
'    Else
'        nuevoCodigoDB = 1
'    End If
'
'    Set rs = Nothing
'End Function



' -****************  TRANSACCIONES daTaenvironment1 *************************
' OJO un solo nivel
Public Function DE_BeginTrans() As Boolean 'adodb.Connection
    If Not DE_EstaAbierto() Then DataEnvironment1.Sistema.Open
    DataEnvironment1.Sistema.BeginTrans
    DE_BeginTrans = True
End Function
Public Function DE_CommitTrans() As Boolean
On Error GoTo ufaCT
    DataEnvironment1.Sistema.CommitTrans
    DE_CommitTrans = True
    Exit Function
ufaCT:
    MsgBox "Error al completar transaccion.", vbCritical, "No se pudo completar"
End Function
Public Function DE_RollbackTrans() As Boolean ' OJO muere silenciosamente
   On Error GoTo ufaCT
   DataEnvironment1.Sistema.RollbackTrans
   DE_RollbackTrans = True
fin:
    Exit Function
ufaCT:
    ufa "", "DE_RollbackTrans Falla rollBack DE "
    Resume fin
End Function
Public Function DE_EstaAbierto() As Boolean
On Error GoTo UfaEA
    DE_EstaAbierto = ((DataEnvironment1.Sistema.State And adStateOpen) > 0)
    Exit Function
UfaEA:
    MsgBox "Error revisando conexion.", vbInformation, "Error en estado"
End Function
Public Function DE_abrir() As Boolean
    On Error GoTo UfaDA
'    daTaenvironment1.Sistema.Open
    If Not DE_EstaAbierto() Then DataEnvironment1.Sistema.Open
    DE_abrir = True
fin:
    Exit Function
UfaDA:
'    ufa "fallo abriendo conexion", "DE_Abrir"
'    DE_abrir = False
'    Resume fin
    End
End Function
' -****************  TRANSACCIONES  daTaenvironment1 *************************
''''' -****************  TRANSACCIONES  DB  *************************
''''' OJO un solo nivel
''''Public Function DB_BeginTrans() As Boolean 'adodb.Connection
''''   On Error GoTo ufaBT
''''    daTaenvironment1.Sistema.BeginTrans
''''    DE_BeginTrans = True
''''Fin:
''''    Exit Function
''''ufaBT:
''''    ufa "fallo intento de Comenzar transaccion", "Db_BeginTrans"
''''    Resume Fin
''''End Function
''''Public Function DB_CommitTrans() As Boolean
''''   On Error GoTo ufaCT
''''   daTaenvironment1.Sistema.CommitTrans
''''   DE_CommitTrans = True
''''Fin:
''''    Exit Function
''''ufaCT:
''''    ufa "fallo intento de completar transaccion", "Db_CommitTrans"
''''    Resume Fin
''''End Function
''''Public Function DB_RollbackTrans() As Boolean ' OJO muere silenciosamente
''''   On Error GoTo ufaCT
''''   daTaenvironment1.Sistema.RollbackTrans
''''   DE_RollbackTrans = True
''''Fin:
''''    Exit Function
''''ufaCT:
''''    ufa "", "DE_RollbackTrans Falla rollBack daTaenvironment1.Sistema "
''''    Resume Fin
''''End Function
''''' -****************  TRANSACCIONES  DB  *************************

Public Function leerEjercicioDenominacion()
    leerEjercicioDenominacion = obtenerDeSQL("select denominacion from Ejercicio where activo = 1")
End Function
Public Function leerEjercicioId(Optional denominacion As String) As Long
    If denominacion = "" Then
        leerEjercicioId = obtenerDeSQL("select ejercicio from Ejercicio where activo = 1")
    Else
        leerEjercicioId = obtenerDeSQL("select ejercicio from Ejercicio where denominacion='" & Trim(denominacion) & "'")
    End If
End Function

Public Function HayProdEnEdicion(strDescrProd As String) As Boolean
    If Trim$(strDescrProd) = "" Then
        HayProdEnEdicion = False
    Else
        HayProdEnEdicion = Not confirma("Hay un Producto en la linea de edicion." & vbCrLf & "¿Desea descartar ese producto?")
    End If
End Function


'Saldo Productos
Public Function ProductosPedidos(CodProducto As String, Limitar5UltimosDias As Boolean)
    Dim s As String, tempo
    
    s = "SELECT Sum(I.Saldo) AS SumaDeSaldo FROM ItemPedidoCliente AS I INNER JOIN Pedidos_Clientes AS P ON I.PEDIDO = P.numero " _
        & " Where P.activo = 1 And P.cancelado = 0 And i.Producto = '" & CodProducto & "' "
    
    If Limitar5UltimosDias Then s = s & " and p.fecha > " & ssFecha(Date - 5)
    
    ProductosPedidos = s2n(obtenerDeSQL(s))
End Function


'series  MODIFICAR PARA Q FUNCIONE
Public Function SerieEnStock(cualSerie As String, cualProducto As String) As Boolean
    '    ss = "SELECT  serie as [ Serie                 ], producto as  [ Producto              ] , MAX(codigo) as  [Movimiento ] From SERIES Where (activo = 1 and producto = '" & prod & "') GROUP BY  producto, serie"
    Dim tempo
    tempo = obtenerDeSQL("select serie from series where serie = '" & cualSerie & "' and producto = '" & cualProducto & "' ")
    SerieEnStock = (sSinNull(tempo) > "")
End Function
Public Function ProductoDescripcion(codi) As String
    codi = sSinNull(codi)
    If codi = "" Then Exit Function
    ProductoDescripcion = obtenerDeSQL("select descripcion from producto where codigo = '" & Trim(codi) & "' and activo = 1 ")
End Function
Public Function Buscar_SeriesEnStock(producto As String) As String
    On Error GoTo UfaBuscaSer
    Dim ss As String, tmpTablaSeries As String
    
    tmpTablaSeries = TablaTempCrear(tt_SeriesEnStockTemp)
    ss = "INSERT INTO " & tmpTablaSeries & " ( Codigo, Producto, Serie, Descripcion ) " _
        & " SELECT max(series.Codigo) as UltimoCodigo, producto, Series.serie, Descripcion From Series " _
        & " inner join producto on producto = producto.codigo " _
        & " Where Series.activo = 1 and producto = '" & producto & "' " _
        & " GROUP BY producto, Descripcion, Series.serie order by Series.serie "
        
' debug
    If producto = "" Then
        ss = "INSERT INTO " & tmpTablaSeries & " ( Codigo, Producto, Serie ) " _
        & " SELECT max(Codigo) as UltimoCodigo, producto, serie From Series " _
        & " Where activo = 1  " _
        & " GROUP BY producto, serie order by serie "
    End If
' '''''
        
    DataEnvironment1.Sistema.Execute ss
    
    ss = "SELECT t.Serie as [ Serie               ], s.Producto as [ Producto                  ], t.Descripcion as [ Descripcion                                              ], s.comprobante as [c], t.codigo  as [i]" _
        & " FROM " & tmpTablaSeries & " AS t INNER JOIN Series AS s ON t.codigo = s.codigo left join conceptos as c on c.codigo = s.concepto " _
        & " WHERE s.comprobante = 6 or s.comprobante = 3 or s.comprobante = 4 or (s.comprobante = 7 and c.movimiento <> 'R' )  "
    Buscar_SeriesEnStock = frmBuscar.MostrarSql(ss)
    
fin:
    Exit Function
UfaBuscaSer:
    ufa "err: buscando series", "Prod: " & producto
    Resume fin
End Function

Public Function GeneraExistenciaCalculada()
'    On Error GoTo UfaCalcExistencia
    If Not gEMPR_FormulaEsVirtual Then Exit Function
    
    
    Dim ss0 As String, ss As String, ss1 As String, ss2 As String
    Dim rs As New ADODB.Recordset, rs2 As New ADODB.Recordset
    Dim CodiZZZ As String, CantMin As Double, reserv As Double, tempo
    
    ss0 = "update producto set ExistenciaCalculada = Existencia, ReservaCalculada = 0"
    DataEnvironment1.Sistema.Execute ss0
    
    'RESERVADOS --- ACA SE TIENE Q HACER LA VERIFICACION DE VENCIMIENTO (fecha venc pedido) ----
    ss0 = "SELECT i.Producto, i.Saldo FROM Pedidos_Clientes AS p INNER JOIN ItemPedidoCliente AS i " _
        & " ON p.numero = i.PEDIDO " _
        & " WHERE (((i.Saldo)>0) AND ((p.activo)=1) AND ((p.cancelado)=0)) " 'and fechavencimiento > xxx
    rs.Open ss0, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    While Not rs.EOF
        DataEnvironment1.Sistema.Execute "update producto set ReservaCalculada =  ReservaCalculada + " & x2s(rs!saldo) & " where codigo = '" & rs!producto & "' "
        rs.MoveNext
    Wend
    rs.Close
     
    If gEMPR_FormulaEsVirtual Then
        ss = "select distinct codigo from formulas where activo = 1"
        rs.Open ss, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        While Not rs.EOF
            CodiZZZ = rs!codigo
            
'            If Trim(CodiZZZ) = "ZZZZZZ553TFH0904" Then Stop
            
            ' averiguo max cant q se pueden armar (min de componente)
            ss1 = "SELECT Min(existenciaCalculada/cantidad) AS MaxArmados, min(ReservaCalculada) as reservado " _
                & " FROM  producto as p INNER JOIN Formulas as f ON p.codigo = f.Componente " _
                & " Where f.codigo = '" & CodiZZZ & "' "
            
            tempo = obtenerDeSQL(ss1)
            CantMin = Fix(s2n(tempo(0)))
            reserv = s2n(tempo(1))
            
            'update virtual cant q se pueden armar
            DataEnvironment1.Sistema.Execute "update producto set ExistenciaCalculada = " & x2s(CantMin) & " where codigo = '" & CodiZZZ & "' and activo = 1 "
            
            'update componentes, resta de virutuales
            If CantMin > 0 Then
                ss2 = "select componente from formulas where activo = 1 and codigo = '" & CodiZZZ & "'"
                rs2.Open ss2, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                While Not rs2.EOF
                    DataEnvironment1.Sistema.Execute "update Producto set " _
                        & " ExistenciaCalculada = ExistenciaCalculada - " & x2s(CantMin) _
                        & ", ReservaCalculada = ReservaCalculada - " & x2s(reserv) _
                        & " where codigo = '" & rs2!Componente & "' "
                    rs2.MoveNext
                Wend
                rs2.Close
            End If
            rs.MoveNext
        Wend
    End If
    
fin:
    Set rs = Nothing
    Set rs2 = Nothing
    Exit Function
   
UfaCalcExistencia:
    ufa "err: Buscando stock", "ExistenciaCalculada"
    Resume fin

End Function

Public Function RevisaNroYFechaOk(sTabla As String, sNum As String, sFec As String, numero As Long, Fecha As Date, masWhere As String, Optional REVFECHA As Boolean = True, Optional PuntoVenta As String = "0001") As Boolean
    Dim ss As String, tempo As Variant
    Dim sWhere2 As String

    RevisaNroYFechaOk = False
    
    If masWhere > "" Then masWhere = " AND " & masWhere
    If PuntoVenta > "" Then sWhere2 = " AND PUNTOVENTA=" & ssTexto(PuntoVenta)
    
    ' reviso numero
    ss = "select " & sFec & " from " & sTabla & " where " & sNum & " = " & numero & masWhere & sWhere2
    tempo = obtenerDeSQL(ss)
    If Not IsEmpty(tempo) Then
        che "Numero " & numero & " existe con fecha  " & tempo
        Exit Function
    End If
    
    If REVFECHA Then
        'Reviso Fecha anterior
        ss = "select max(" & sNum & ") from " & sTabla & " where " & sNum & " < " & numero & masWhere & sWhere2
        tempo = s2n(obtenerDeSQL(ss), 0)
        If tempo > 0 Then
            ss = "select " & sFec & " from " & sTabla & " where " & sNum & " = " & tempo & masWhere & sWhere2
            tempo = obtenerDeSQL(ss)
            If tempo > Fecha Then
                che "Documento anterior tiene fecha " & tempo
                Exit Function
            End If
        End If
        
        'Reviso Fecha Posterior
        ss = "select min(" & sNum & ") from " & sTabla & " where " & sNum & " > " & numero & masWhere & sWhere2
        tempo = s2n(obtenerDeSQL(ss), 0)
        If tempo > 0 Then
            ss = "select " & sFec & " from " & sTabla & " where " & sNum & " = " & tempo & masWhere & sWhere2
            tempo = obtenerDeSQL(ss)
            If tempo < Fecha Then
                che "Documento posterior tiene fecha " & tempo
                Exit Function
            End If
        End If
    End If
    RevisaNroYFechaOk = True
End Function
Public Function RevisaNro(sTabla As String, sNum As String, sFec As String, numero As Long, masWhere As String) As Boolean
    Dim ss As String, tempo As Variant

    RevisaNro = False
    
    If masWhere > "" Then masWhere = " AND " & masWhere
    
    ' reviso numero
    ss = "select " & sFec & " from " & sTabla & " where " & sNum & " = " & numero & masWhere
    tempo = obtenerDeSQL(ss)
    If Not IsEmpty(tempo) Then
        che "Numero " & numero & " existe con fecha  " & tempo
        Exit Function
    End If
  
    RevisaNro = True
End Function

Public Function nuevoCodigoOP() As Long
    Dim tmp As Long
    
    tmp = nuevoCodigo("Rec_Comp", "Nro")
    nuevoCodigoOP = tmp
    
    tmp = nuevoCodigo("transcom", "NroDoc", "TipoDoc = 'RAC'")
    If tmp > nuevoCodigoOP Then nuevoCodigoOP = tmp
    
    tmp = nuevoCodigo("Compras", "NroDoc", "TipoDoc = 'RAC'")
    If tmp > nuevoCodigoOP Then nuevoCodigoOP = tmp
    
' mod 2006 TONKA ' provisorio
    tmp = NuevoNroPago()
    If tmp > nuevoCodigoOP Then nuevoCodigoOP = tmp
' mod 2006 TONKA ' provisorio, no jode si queda asi .
    
    
    ' DESPUES DE LA IMPLEMENTACION, ESTA ES LA UNICA LINEA QUE VA
    ' nuevoCodigoOP = NuevoNroPago()
    ' DESPUES DE LA IMPLEMENTACION, ESTA ES LA UNICA LINEA QUE VA
    
End Function
Public Function existeOP(cual) As Boolean
    Dim tempo
    tempo = obtenerDeSQL("Select TipoDoc, NroDoc from transcom where tipodoc = 'RAC' and NroDoc = " & cual)
    existeOP = Not IsEmpty(tempo)
    If existeOP Then Exit Function
    
    tempo = obtenerDeSQL("Select TipoDoc, NroDoc from compras where tipodoc = 'RAC' and NroDoc = " & cual)
    existeOP = Not IsEmpty(tempo)
    If existeOP Then Exit Function
    
    tempo = obtenerDeSQL("Select Nro, id from rec_comp where activo = 1 and Nro = " & cual)
    existeOP = Not IsEmpty(tempo)
    If existeOP Then Exit Function
End Function
'Public Function CodigoDeAlias(sAlias)
'    CodigoDeAlias = sSinNull(obtenerDeSQL("select codigo from producto where activo = 1 and alias = '" & sAlias & "' "))
'End Function
'Public Function AliasDeCodigo(sCod)
'    AliasDeCodigo = sSinNull(obtenerDeSQL("select alias from producto where activo = 1 and codigo = '" & sCod & "' "))
'End Function
Public Sub ReordenarAsientos()
   If ON_ERROR_HABILITADO Then On Error GoTo ufaChe
    Dim i, n, rs As New ADODB.Recordset
    n = 1
    i = leerEjercicioId()
    With rs
        .Open "select * from asientos where activo = 1 and Ejercicio = " & i & " order by fecha, NroAsiento ", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        While Not .EOF
            !NroAsiento = n
            .Update
            n = n + 1
            .MoveNext
        Wend
    End With
    che "Ejercicio ordenado"
fin:
    Set rs = Nothing
    Exit Sub
ufaChe:
    ufa "err al ordenar asientos", "i = " & i & " n= " & n
    Resume fin
End Sub
Public Function CuentaDescripcion(CUENTA, Optional conNoImputables As Boolean = True, Optional conNoActivas As Boolean = True) As String
    Dim ss As String
    ss = "select descripcion from cuentas where cuenta = '" & CUENTA & "' "
    If conNoImputables = False Then ss = ss & " and imputable = 1 "
    If conNoActivas = False Then ss = ss & " and activo = 1 "
    CuentaDescripcion = sSinNull(obtenerDeSQL(ss))
End Function
Public Function BuscarCuenta(Optional conNoImputables As Boolean = True, Optional conNoActivas As Boolean = True, Optional prov As Long, Optional clie As Long) As String
    Dim ss As String, sWhe As String, sCtas
    
    ss = "select Cuenta as [ Cuenta        ], descripcion as [ Descripcion                                ], Imputable, Activo  from cuentas " & sWhe
    
    If prov > 0 Then
        sCtas = sSinNull(obtenerDeSQL("select cuentascompras from prov where codigo=" & prov))
        If sCtas = "" Then
        Else
        sWhe = " where cuenta in (" & Replace(sCtas, "#", "'") & ")"
        End If
        ss = ss & sWhe
    Else
        If conNoImputables = False Then
             ss = ss & " where imputable = 1 "
        End If
        If conNoActivas = False Then
            If conNoImputables = False Then
                ss = ss & " and activo = 1 "
            Else
                ss = ss & " where activo = 1"
            End If
        End If
    End If
    
    BuscarCuenta = frmBuscar.MostrarSql(ss, , "CUENTAS", "-", "", "No")
End Function
Public Function ManejaStock(prod As String) As Long
    Dim rs As New ADODB.Recordset
    
    rs.Open "select ManejaStock from Producto where codigo = '" & prod & "' ", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    If rs.Fields(0) = True Then
        ManejaStock = 1
    Else
        ManejaStock = 0
    End If
    Set rs = Nothing
    'ManejaStock = s2n(obtenerDeSQL("select ManejaStock from Producto where codigo = '" & prod & "' "))
    
End Function

Public Function CalcPercIIBB(monto As Double, cliente As Long) As Double
    'calcula iibb  en base al monto y cliente, contempla base imponible, categoria.
    
    Dim tempo
    
    CalcPercIIBB = 0 ' para empezar...
    If cliente = 0 Then Exit Function
    
    ' monto no alcanza
    If monto < s2n(obtenerDeSQL("select baseperiibb from bs")) Then Exit Function
      
    tempo = obtenerDeSQL("select ConPercIIBB, iva, coefPercIIBB  from clientes left join ivas on clientes.iva = ivas.codigo where clientes.codigo = '" & x2s(cliente) & " ' and clientes.activo = 1 and ivas.activo = 1")
    
    'cliente no alcanzado
    If tempo(0) = False Then Exit Function
    
    CalcPercIIBB = monto * s2n(tempo(2))
End Function
Public Function ON_ERROR_HABILITADO() As Boolean
    ON_ERROR_HABILITADO = (Not (FrmPrincipal.chkDesbichando.Value = vbChecked))
End Function
Public Function PREVIEW_IMPRESIONES() As Boolean
    PREVIEW_IMPRESIONES = (FrmPrincipal.chkPreviewImpresion.Value = vbChecked) Or VerParametro(BS_PREVIEW_IMPRESIONES)
End Function
Public Function TABLA_TEMP_FIJA() As Boolean
    TABLA_TEMP_FIJA = (FrmPrincipal.chkVerTablaTemp.Value = vbChecked)
End Function


Public Function precioProducto(CodigoCli As String, Propio As Boolean, cliente As Long) As Double
    On Error Resume Next ' por las dudas
    
    Dim preci As Double
'    If Not Propio Then
        preci = s2n(obtenerDeSQL("select precio from relacion_producto_cliente where  productocliente = '" & CodigoCli & "' and cliente = " & cliente), 4)
'    End If
    If preci = 0 Then
        preci = s2n(obtenerDeSQL("select precio from relacion_producto_cliente where  producto = '" & CodigoCli & "' and cliente = " & cliente), 4)
    End If
    If preci = 0 Then
        preci = s2n(obtenerDeSQL("select precio from producto where codigo = '" & CodigoCli & "' "), 4)
    End If
    precioProducto = preci
End Function

Private Function ExisteFacCompra(prov, suc, Nro) As String
    ' busca si hay FAC ND NC
    ' devuelve fecha como string o vacio
    Dim resu
    Dim whe  As String
    
    If prov = 0 Then Exit Function
  
    whe = " where  (TIPODOC = 'N/C' OR TIPODOC = 'N/D' OR TIPODOC = 'FAC') and codpr = " & prov & " and suc = " & suc & " and NroDoc = " & Nro
    
    resu = obtenerDeSQL("select tipodoc, fecha from compras " & whe)
    If Not IsEmpty(resu) Then
        ExisteFacCompra = resu(0) & " " & CStr(resu(1))
        Exit Function
    End If
        
    resu = obtenerDeSQL("select tipodoc, fecha from transcom " & whe)
    If Not IsEmpty(resu) Then
        ExisteFacCompra = resu(0) & " " & CStr(resu(1))
        Exit Function
    End If
End Function

Private Function ExisteDocBanco(prov, suc, Nro) As String
    ' busca si hay FAC ND NC
    ' devuelve fecha como string o vacio
    Dim resu
    Dim whe  As String
    
    If prov = 0 Then Exit Function
  
    whe = " where codbanco = " & prov & " and suc = " & suc & " and NroDoc = " & Nro
    
    resu = obtenerDeSQL("select tipodoc, fecha from gastosbancarios " & whe)
    If IsEmpty(resu) Or IsNull(resu) Then
    Else
        ExisteDocBanco = resu(0) & " " & CStr(resu(1))
        Exit Function
    End If
End Function

Public Function ExisteFacCompraMSG(prov, suc, Nro) As Boolean
    ' si existe, tira mensaje al usuario y devueve TRUE
    Dim resu As String
    
     resu = ExisteFacCompra(prov, suc, Nro)
     If resu > "" Then
        MsgBox "Documento existente: " & resu
        ExisteFacCompraMSG = True
    End If
End Function


Public Function ExisteDocBancoMSG(prov, suc, Nro) As Boolean
    ' si existe, tira mensaje al usuario y devueve TRUE
    Dim resu As String
    
     resu = ExisteDocBanco(prov, suc, Nro)
     If resu > "" Then
        MsgBox "Documento existente: " & resu
        ExisteDocBancoMSG = True
    End If
End Function

Public Function ExisteBoletaMSG(prov, suc, Nro) As Boolean
    ' si existe, tira mensaje al usuario y devueve TRUE
    Dim resu As String
    
     resu = ExisteBoleta(prov, suc, Nro)
     If resu > "" Then
        MsgBox "Documento existente: " & resu
        ExisteBoletaMSG = True
    End If
End Function

Private Function ExisteBoleta(prov, suc, Nro) As String
    ' busca si hay FAC ND NC
    ' devuelve fecha como string o vacio
    Dim resu
    Dim whe  As String
    
    If prov = 0 Then Exit Function
  
    whe = " where codpr = " & prov & " and suc = " & suc & " and NroDoc = " & Nro
    
    resu = obtenerDeSQL("select tipodoc, fecha from gastosboletas " & whe)
    If IsEmpty(resu) Or IsNull(resu) Then
    Else
        ExisteBoleta = resu(0) & " " & CStr(resu(1))
        Exit Function
    End If
End Function

Public Function ProvCoefIVA(CodProv) As Double
    ProvCoefIVA = s2n(obtenerDeSQL("select porcentaje from prov inner join porcentajesiva as ivas  on prov.tipoiva = ivas.iva where prov.codigo = '" & CodProv & "'"))
End Function


Public Function NuevoMoviCaja() As Long
    NuevoMoviCaja = nuevoCodigo("Movicaja", "movimiento")
End Function
Public Function NuevoMovibanc() As Long
    NuevoMovibanc = nuevoCodigo("movibanc", "movbanco")
End Function


Public Function BuscarProducto(Optional Propio As Boolean = True) As String
    Dim s As String
    If Propio Then
        s = "select p.codigo as [Codigo          ], p.letra as [Letra], Descripcion as [ Descripcion                               ], e.estado as [ Estado         ] from producto as p inner join ProductoEstado as e on e.codigo = p.estado where p.activo = 1"
    Else
    End If
    BuscarProducto = frmBuscar.MostrarSql(s)
End Function

Public Function VerProdProv(prod As String, prov As Long, Optional DefaultProd As Boolean = False)
    'Dim s
    VerProdProv = ssStr(obtenerDeSQL("select CodigoProveedor From Relacion_Producto_Proveedor where producto = '" & ssStr(prod, True) & "' and proveedor = " & prov))
'    VerProdClie = IIf(Trim(s) = "", prod, s)
End Function

Public Function BuscarCliente(Optional arrayresultado As Variant) As String
    ' devuelve string codigo, o modifica arrayresultado 1 = codigo 2 = descripcion
    BuscarCliente = frmBuscar.MostrarSql("select codigo, descripcion as [Nombre Cliente                    ] from clientes where activo = 1", , , "", , , True)
    
    arrayresultado = Array(frmBuscar.resultado(1), frmBuscar.resultado(2))
End Function

Public Function reformateoProducto(cual As String) As String
    ' devuelve formato  XXX-XXX-XXXXX
    
    If Trim(cual) = "" Then Exit Function
    
    reformateoProducto = Left(Left(cual, 3) & "-" & Mid(cual, 4, 3) & "-" & Mid(cual, 7) & "              ", 13)
End Function


Public Function Prov_NumIIBB(cual As Long) As String
    Prov_NumIIBB = sSinNull(obtenerDeSQL("select numiibb from prov where codigo = " & cual))
End Function

Public Function verDescBanco(codi)
    On Error Resume Next
    verDescBanco = sSinNull(obtenerDeSQL("select descripcion from bancosGrales where codigo = " & codi))
End Function

Public Sub grillaMarcoSaldosFinales(quegrilla, ColCorte As Long, colmarca As Long, colsaldo As Long)
    Dim i As Long, s As String, v As String
    
    With quegrilla
        For i = .rows - 1 To 1 Step -1
            s = .TextMatrix(i, ColCorte)
            If v <> s Then .TextMatrix(i, colmarca) = .TextMatrix(i, colsaldo)
            v = s
        Next i
    End With
End Sub

Public Function TraerSumariza(cta As String) As String
Dim rs As New ADODB.Recordset


rs.Open "select sumariza from cuentas where cuenta='" & Trim(cta) & "'", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
If Not rs.EOF Then
    TraerSumariza = rs!SUMARIZA
Else
    TraerSumariza = "0"
End If
rs.Close
Set rs = Nothing

End Function

Public Function VentanaCarpeta(Optional Titulo As String, Optional Path_Inicial As Variant) As String
  
On Local Error GoTo errFunction
       
    Dim objShell As Object
    Dim objFolder As Object
    Dim o_Carpeta As Object
       
    ' Nuevo objeto Shell.Application
    Set objShell = CreateObject("Shell.Application")
       
    On Error Resume Next
    'Abre el cuadro de diálogo para seleccionar
    Set objFolder = objShell.BrowseForFolder( _
                            0, _
                            Titulo, _
                            0, _
                            Path_Inicial)
       
    ' Devuelve solo el nombre de carpeta
    Set o_Carpeta = objFolder.Self
       
    ' Devuelve la ruta completa seleccionada en el diálogo
    VentanaCarpeta = o_Carpeta.Path
  
Exit Function
'Error
errFunction:
    MsgBox Err.Description, vbCritical
    VentanaCarpeta = vbNullString
  
End Function


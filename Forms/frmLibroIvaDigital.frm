VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLibroIvaDigital 
   Caption         =   "LID-CITI"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3465
   LinkTopic       =   "Form1"
   ScaleHeight     =   2580
   ScaleWidth      =   3465
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCiti 
      Caption         =   "CITI"
      Height          =   810
      Left            =   2070
      TabIndex        =   10
      Top             =   600
      Width           =   750
   End
   Begin VB.OptionButton optFecha 
      Caption         =   "Por fecha de comprobante"
      Height          =   255
      Left            =   570
      TabIndex        =   9
      Top             =   2100
      Value           =   -1  'True
      Width           =   2415
   End
   Begin VB.OptionButton optImputacion 
      Caption         =   "Por periodo de imputacion"
      Height          =   255
      Left            =   555
      TabIndex        =   8
      Top             =   1740
      Width           =   2295
   End
   Begin VB.CommandButton cmdLibroDigitalVentasCompras 
      Caption         =   "LID"
      Height          =   810
      Left            =   1245
      TabIndex        =   7
      Top             =   615
      Width           =   750
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Otras Percepciones"
      Height          =   375
      Left            =   6435
      TabIndex        =   5
      Top             =   345
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Detalle de Facturas"
      Height          =   375
      Left            =   6360
      TabIndex        =   4
      Top             =   1515
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cabecera de Facturas"
      Height          =   375
      Left            =   6435
      TabIndex        =   3
      Top             =   915
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Libro COMPRAS"
      Height          =   375
      Left            =   6345
      TabIndex        =   2
      Top             =   2145
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker dtPeriodo 
      Height          =   375
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      OLEDropMode     =   1
      CustomFormat    =   "MM/yyyy"
      Format          =   241696771
      CurrentDate     =   39361
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Libro VENTAS"
      Height          =   375
      Left            =   6330
      TabIndex        =   0
      Top             =   2790
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Periodo"
      Height          =   255
      Left            =   615
      TabIndex        =   6
      Top             =   195
      Width           =   1920
   End
End
Attribute VB_Name = "frmLibroIvaDigital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function LibroDigitalCompras()
On Error GoTo compras_err
Dim cFecha As String, cTipo As String, cPunto As String, cNro As String, cNroDespacho As String, cCodigo As String, cIdentificador As String, cApellidoProveedor As String, cTotal As String, cNoGravado As String, cExentas As String, cPercepciones1 As String, cPercepciones2 As String, cIngresosBrutos As String, cMunicipales As String, cInternos As String, cCodigoMoneda As String, cCambio As String, cCantidadAlicuotas As String, cCodOperacion As String, cCreditoFiscal As String, cTributos As String, cCuitEmisor As String, cDenominacionEmisor As String, cComision As String
Dim cPrimerDia As Date, cUltimoDia As Date, str As String, i As Long, sNombreFile As String, sCarpeta As String, sArchivoCompleto As String, sArchivoCompleto2 As String, sArchivoCompleto3 As String, pLetra As String, pTipoDoc As String, a As Integer
Dim cNeto As String, cAlicuota As String, cImpuestoLiquidado As String
Dim dDESPACHO As String, dNETO As String, dALICUOTA As String, dIMPUESTO As String, sss As String
Dim rs1 As New ADODB.Recordset
Dim sCampofecha As String

    cPrimerDia = "01/" & Month(dtPeriodo) & "/" & Year(dtPeriodo)
    cUltimoDia = ultimoDiaDelMes(dtPeriodo)
    
    If optFecha.Value = True Then
        sCampofecha = " c.fecha>=" & ssFecha(cPrimerDia) & " and c.fecha<=" & ssFecha(cUltimoDia) & ""
    Else
         sCampofecha = " c.mesimp=" & Month(dtPeriodo) & " and c.anoimp=" & Year(dtPeriodo) & " "
    End If
    
        sss = "select c.fecha,codpr,nrodoc,iva_21,iva_27,0 as iva_9,iva_10,imp_int,razonsocialprov,cuitprov,suc,anoimp,TIPODOC,TOTAL,NETO,EXENTO,(IBCAPITAL+IBPROVINCIA) AS IIBB,LETRA,(PERCEPC+DER_EST+iva_9) AS RETIVA,0 as retiva2,'' as nrodespacho,0 as neto27,0 as neto10,nogravado " & _
                    " from TRANSCOM as c " & _
                    " where " & sCampofecha & " and tipodoc<>'RAC' and tipodoc<>'APC' and tipodoc<>'APD' " & _
            "union " & _
                " select c.fecha,codpr,nrodoc,iva_21,iva_27,0 as iva_9,iva_10,imp_int,razonsocialprov,cuitprov,suc,anoimp,TIPODOC,TOTAL,NETO,EXENTO,(IBCAPITAL+IBPROVINCIA) AS IIBB,LETRA,(PERCEPC+DER_EST+iva_9) AS RETIVA,0 as retiva2,'' as  nrodespacho,0 as neto27,0 as neto10,nogravado " & _
                    " from compras as c " & _
                    " where " & sCampofecha & " and tipodoc<>'RAC' and tipodoc<>'APC' and tipodoc<>'APD' " & _
             "union " & _
                " select d.fecha,d.proveedor as codpr,0 as nrodoc, d.iva as iva21,0 as iva_27,0 as iva_9,0 as iva_10,0 as imp_int,P.descripcion as razonsocialprov, P.cuit as cuitprov, 0 as suc,year(fecha) as fecha,'DDD' as tipodoc,(D.base + D.total) as total,D.base as neto,D.exento,0 as iibb,'D' as letra,D.percrg3431 as  retiva,D.percganancia as  retiva2,D.nrodespacho,0 as neto27, 0 as neto10,0 as nogravado " & _
                " from despachodeimportacion as d inner join prov as p  on d.proveedor=p.codigo " & _
                " where d.ACTIVO=1 AND d.fecha>=" & ssFecha(cPrimerDia) & " and d.fecha<=" & ssFecha(cUltimoDia)
'            "union " & _
'                " select gastosbancarios.fecha,codbanco as codpr, nrodoc, iva_21 , iva_27,0 as iva_9,iva_10,0 as imp_int,razonsocialbanco as razonsocialprov, cuitbanco as cuitprov, 0 as suc,year(gastosbancarios.fecha) as fecha,tipodoc,(mantcta+gastoschqra+gastosvarios+sellado+intxgiro+valnoconfor+iva_21+iva_27+iva_10) as total,(mantcta+gastoschqra+gastosvarios+sellado+intxgiro+valnoconfor) as neto,0 as exento,0 as iibb, letra,0 as  retiva,0 as  retiva2,'' as nrodespacho " & _
'                " from gastosbancarios inner join prov  on gastosbancarios.codbanco=prov.codigo " & _
'                " where  gastosbancarios.fecha>=" & ssFecha(cPrimerDia) & " and gastosbancarios.fecha<=" & ssFecha(cUltimoDia) & " " & _
'            "union " & _
'                " select d.fecha,c.codigo as codpr,d.numero as nrodoc, d.iva as iva21,0 as iva_27,0 as iva_9,0 as iva_10,0 as imp_int,b.descripcion as razonsocialprov, c.cuit as cuitprov, 0 as suc,year(d.fecha) as fecha,'LP' as tipodoc,(d.capital + d.interes + d.iva) as total,(d.capital+d.interes) as neto,0 as exento,0 as iibb,'A' as letra,0 as  retiva,0 as  retiva2,'' as nrodespacho " & _
'                " from ((prestamosldetalle as d inner join prestamosl as p on d.idprestamo=p.idprestamo) inner join CTASBANK as c on p.cuenta=c.codigo) inner join bancosgrales as b on c.banco=b.codigo " & _
'                " where d.fecha>=" & ssFecha(cPrimerDia) & " and d.fecha<=" & ssFecha(cUltimoDia) & " " & _
'            "union " & _
'                " select d.fecha,c.codigo as codpr,d.numero as nrodoc, d.iva as iva21,0 as iva_27,0 as iva_9,0 as iva_10,0 as imp_int,b.descripcion as razonsocialprov, c.cuit as cuitprov, 0 as suc,year(d.fecha) as fecha,'LP' as tipodoc,(d.honorarios + d.iva) as total,(d.honorarios) as neto,0 as exento,0 as iibb,'A' as letra,0 as  retiva,0 as  retiva2,'' as nrodespacho " & _
'                " from ((prestamosldetalleg as d inner join prestamosl as p on d.idprestamo=p.idprestamo) inner join CTASBANK as c on p.cuenta=c.codigo) inner join bancosgrales as b on c.banco=b.codigo " & _
'                " where d.fecha>=" & ssFecha(cPrimerDia) & " and d.fecha<=" & ssFecha(cUltimoDia) & "  " & _
'             "union " & _
'                " select d.fecha,d.proveedor as codpr,0 as nrodoc, d.iva as iva21,0 as iva_27,0 as iva_9,0 as iva_10,0 as imp_int,P.descripcion as razonsocialprov, P.cuit as cuitprov, 0 as suc,year(fecha) as fecha,'DDD' as tipodoc,(D.base + D.total) as total,D.base as neto,D.exento,0 as iibb,'D' as letra,D.percrg3431 as  retiva,D.percganancia as  retiva2,D.nrodespacho " & _
'                " from despachodeimportacion as d inner join prov as p  on d.proveedor=p.codigo " & _
'                " where d.ACTIVO=1 AND d.fecha>=" & ssFecha(cPrimerDia) & " and d.fecha<=" & ssFecha(cUltimoDia)
                
        rs1.Open sss, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly

    
    With rs1
        
        If rs1.EOF And rs1.BOF Then
            MsgBox "No se ha encontrado datos para este periodo.", vbInformation, "ATENCION"
            Exit Function
        Else
            
            
            sCarpeta = "C:\"
            sCarpeta = Trim(VentanaCarpeta("Carpeta Destino txt COMPRAS", sCarpeta))
            If Trim(sCarpeta) = "" Then Exit Function
            If CORTO(sCarpeta, Len(sCarpeta) - 1, 0) <> "\" Then sCarpeta = sCarpeta & "\"
            sNombreFile = "LIBRO-IVA-DIGITAL-COMPRAS_" & Year(dtPeriodo) & Format(Month(dtPeriodo), "00") & "_" & Cuit_Empresa_Carga & ".txt"
            sNombreFile2 = "LIBRO-IVA-DIGITAL-COMPRAS_ALICUOTAS_" & Year(dtPeriodo) & Format(Month(dtPeriodo), "00") & "_" & Cuit_Empresa_Carga & ".txt"
            sNombreFile3 = "LIBRO-IVA-DIGITAL-DESPACHOS_ALICUOTAS_" & Year(dtPeriodo) & Format(Month(dtPeriodo), "00") & "_" & Cuit_Empresa_Carga & ".txt"
            sArchivoCompleto = sCarpeta & sNombreFile
            sArchivoCompleto2 = sCarpeta & sNombreFile2
            sArchivoCompleto3 = sCarpeta & sNombreFile3
            If ExisteArchivo(sArchivoCompleto) Then
                Kill sArchivoCompleto
            End If
            If ExisteArchivo(sArchivoCompleto2) Then
                Kill sArchivoCompleto2
            End If
            Open sArchivoCompleto For Output As #1
            Open sArchivoCompleto2 For Output As #2
            
            Open sArchivoCompleto3 For Output As #3
            
            .MoveFirst
            For i = 0 To .RecordCount - 1
            
                cFecha = "19000101"
                cTipo = "000"
                cPunto = "00000"
                cNro = "00000000000000000000"
                cNroDespacho = "                "
                cCodigo = "80"
                cIdentificador = "00000000000000000000"
                cApellidoProveedor = "----------SIN DATOS ----------"
                cTotal = "000000000000000"
                cNoGravado = "000000000000000"
                cExentas = "000000000000000"
                cPercepciones1 = "000000000000000"
                cPercepciones2 = "000000000000000"
                cIngresosBrutos = "000000000000000"
                cMunicipales = "000000000000000"
                cInternos = "000000000000000"
                cCodigoMoneda = "PES"
                cCambio = "0001000000"
                cCantidadAlicuotas = "1"
                cCodOperacion = " "
                cCreditoFiscal = "000000000000000"
                cTributos = "000000000000000"
                cCuitEmisor = "00000000000"
                cDenominacionEmisor = "                              "
                cComision = "000000000000000"
            
                cNeto = "000000000000000"
                cAlicuota = "0000"
                cImpuestoLiquidado = "000000000000000"
                
                cFecha = Format(!Fecha, "YYYYMMDD")
                
                pLetra = Trim(sSinNull(!letra))
                If pLetra = "" Then
                    pLetra = Trim(obtenerDeSQL("select i.LETRA from prov p inner join ivas i on i.codigo=p.tipoiva where p.codigo=" & s2n(!CODPR)))
                End If
                
                pTipoDoc = Trim(!TIPODOC) & pLetra
                Select Case Trim(pTipoDoc)
                    Case "FACA": cTipo = "001"
                    Case "BGA": cTipo = "001"
                    Case "PLA": cTipo = "001"
                    Case "N/DA": cTipo = "002"
                    Case "N/CA": cTipo = "003"
                    Case "FACB": cTipo = "006"
                    Case "N/DB": cTipo = "007"
                    Case "N/CB": cTipo = "008"
                    Case "FACC": cTipo = "011"
                    Case "FACE": cTipo = "019"
                    Case "N/CE": cTipo = "021"
                    Case "DDDD": cTipo = "066"
                End Select
                'cTipo = Format(Trim(!TIPODOC2), "000")
                
                cPunto = Format(!suc, "00000")
                If cPunto = "00000" Then cPunto = "00001"
                cNro = Format(!NroDoc, "00000000000000000000")
                cIdentificador = Format(Replace(!cuitprov, "-", ""), "00000000000000000000")
                
                
                cApellidoProveedor = Trim(!razonsocialprov)
                If Len(cApellidoProveedor) > 30 Then
                    cApellidoProveedor = CORTO(cApellidoProveedor, 0, Len(cApellidoProveedor) - 30)
                Else
                    While Len(cApellidoProveedor) < 30
                        cApellidoProveedor = Chr(32) & cApellidoProveedor
                    Wend
                End If
                
                cTotal = Format(Replace(s2n(!Total, 2, True), ",", ""), "000000000000000")
                cExentas = Format(Replace(s2n(!EXENTO, 2, True), ",", ""), "000000000000000")
                cNoGravado = Format(Replace(s2n(!nogravado, 2, True), ",", ""), "000000000000000")
                cNeto = Format(Replace(s2n(!Neto, 2, True), ",", ""), "000000000000000")
                If pLetra = "C" Or pLetra = "B" Then
                    cNeto = "000000000000000"
                    cNoGravado = "000000000000000"
                    cExentas = "000000000000000"
                End If
                cPercepciones1 = Format(Replace(s2n(!retIva, 2, True), ",", ""), "000000000000000")
                cPercepciones2 = Format(Replace(s2n(!retIva2, 2, True), ",", ""), "000000000000000")
                cIngresosBrutos = Format(Replace(s2n(!IIBB, 2, True), ",", ""), "000000000000000")
                cInternos = Format(Replace(s2n(!imp_int, 2, True), ",", ""), "000000000000000")
                
                a = 0
                If pLetra = "B" Or pLetra = "C" Then
                    cAlicuota = "0003"
                    Print #2, cTipo & cPunto & cNro & cCodigo & cIdentificador & cNeto & cAlicuota & cImpuestoLiquidado
                Else
                    If pLetra <> "D" Then
                        If s2n(!IVA_21) > 0 Then
                            a = a + 1
                            cAlicuota = "0005"
                            cImpuestoLiquidado = Format(Replace(s2n(!IVA_21, 2, True), ",", ""), "000000000000000")
                            cNeto = s2n(!IVA_21 / 0.21, 2, True) 's2n(!Neto, 2, True)  's2n(!IVA_21 / 0.21, 2, True)
                            cNeto = Format(Replace(s2n(cNeto, 2, True), ",", ""), "000000000000000")
                            Print #2, cTipo & cPunto & cNro & cCodigo & cIdentificador & cNeto & cAlicuota & cImpuestoLiquidado
                        End If
                        If s2n(!IVA_27) > 0 Then
                            a = a + 1
                            cAlicuota = "0006"
                            cImpuestoLiquidado = Format(Replace(s2n(!IVA_27, 2, True), ",", ""), "000000000000000")
                            cNeto = s2n(!IVA_27 / 0.27, 2, True) 's2n(!neto27, 2, True) 's2n(!IVA_27 / 0.27, 2, True)
                            cNeto = Format(Replace(s2n(cNeto, 2, True), ",", ""), "000000000000000")
                            Print #2, cTipo & cPunto & cNro & cCodigo & cIdentificador & cNeto & cAlicuota & cImpuestoLiquidado
                        End If
                        If s2n(!iva_10) > 0 Then
                            a = a + 1
                            cAlicuota = "0004"
                            cImpuestoLiquidado = Format(Replace(s2n(!iva_10, 2, True), ",", ""), "000000000000000")
                            cNeto = s2n(!iva_10 / 0.105, 2, True) 's2n(!Neto10, 2, True) ' s2n(!iva_10 / 0.105, 2, True)
                            cNeto = Format(Replace(s2n(cNeto, 2, True), ",", ""), "000000000000000")
                            Print #2, cTipo & cPunto & cNro & cCodigo & cIdentificador & cNeto & cAlicuota & cImpuestoLiquidado
                        End If
                    End If
                End If
                cCantidadAlicuotas = a
                cCreditoFiscal = Format(Replace(s2n(s2n(!iva_10) + s2n(!IVA_21) + s2n(!IVA_27), 2, True), ",", ""), "000000000000000")
                
                If pLetra = "A" And s2n(a) = 0 Then
                    cCodOperacion = "E"
                    cAlicuota = "0003"
                    cCantidadAlicuotas = "1"
                    Print #2, cTipo & cPunto & cNro & cCodigo & cIdentificador & cNeto & cAlicuota & cImpuestoLiquidado
                End If
                
                If pLetra = "D" Then
                    cCantidadAlicuotas = 1
                    cPunto = "00000"
                    dDESPACHO = "0000000000000000"
                    dNETO = "0000"
                    dALICUOTA = "000000000000000"
                    dIMPUESTO = "000000000000000"
                    
                    dDESPACHO = sSinNull(!nrodespacho)
                    cNroDespacho = dDESPACHO
                    dNETO = Format(Replace(s2n(!Neto, 2, True), ",", ""), "000000000000000")
                    dALICUOTA = "0005"
                    dIMPUESTO = Format(Replace(s2n(!IVA_21, 2, True), ",", ""), "000000000000000")
                    
                    Print #3, dDESPACHO & dNETO & dALICUOTA & dIMPUESTO
                End If
                
                
                Select Case pTipoDoc
                    Case "N/DA": If a = 0 Then cCodOperacion = "A"
                    Case "N/DB": If a = 0 Then cCodOperacion = "A"
                    Case "FACE": cCodOperacion = "X"
                    Case "N/CE": cCodOperacion = "X"
                End Select
                
                cCuitEmisor = Format(cIdentificador, "00000000000")
                cDenominacionEmisor = cApellidoProveedor
                
                Print #1, cFecha & cTipo & cPunto & cNro & cNroDespacho & cCodigo & cIdentificador & cApellidoProveedor & cTotal & cNoGravado & cExentas & cPercepciones1 & cPercepciones2 & cIngresosBrutos & cMunicipales & cInternos & cCodigoMoneda & cCambio & cCantidadAlicuotas & cCodOperacion & cCreditoFiscal & cTributos & cCuitEmisor & cDenominacionEmisor & cComision
                
                .MoveNext
            Next
            
            
            

        End If
        
    End With
    
    
    Close #1
    Close #2
    Close #3
    
    Set rs1 = Nothing

    MsgBox "Archivo generado correctamente.", vbInformation, "Compras"

Exit Function
compras_err:
MsgBox "Error... " & Err.Number & " " & Err.Description


End Function

Private Function LibroDigitalVentas()
On Error GoTo ventas_err
Dim vFecha As String, vTipo As String, vPunto As String, vNro As String, vNroHasta As String, vCodigo As String, vIdentificador As String, vApellidoCliente As String, vTotal As String, vNoGravado As String, vPercepcionNN As String, vExentas As String, vPercepciones As String, vIngresosBrutos As String, vMunicipales As String, vInternos As String, vCodigoMoneda As String, vCambio As String, vCantidadAlicuotas As String, vCodOperacion As String, vTributos As String, vVencimiento As String
Dim vPrimerDia As Date, vUltimoDia As Date, str As String, i As Long, sNombreFile2 As String, sNombreFile As String, sCarpeta As String, sArchivoCompleto As String, sArchivoCompleto2 As String
Dim vNeto As String, vAlicuota As String, vImpuestoLiquidado As String, nnNeto As Double, a As Integer
Dim rs1 As New ADODB.Recordset

    vPrimerDia = "01/" & Month(dtPeriodo) & "/" & Year(dtPeriodo)
    vUltimoDia = ultimoDiaDelMes(dtPeriodo)
    str = "select * from facturaventa where fecha>=" & ssFecha(vPrimerDia) & " and fecha<=" & ssFecha(vUltimoDia) & " and tipodoc in ('NDB','NDA','FAE','FAA','FAB','NCA','NCB','FEA','FEB','FEC','DEA','DEB','DEC','CEA','CEB','CEC') order by fecha,tipodoc,nrofactura"
    rs1.Open str, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly

    With rs1
        If .EOF And .BOF Then
            MsgBox "No se ha encontrado datos para este periodo.", vbInformation, "ATENCION"
            Exit Function
        Else
            sCarpeta = "C:\"
            sCarpeta = Trim(VentanaCarpeta("Carpeta Destino txt VENTAS", sCarpeta))
            If Trim(sCarpeta) = "" Then Exit Function
            If CORTO(sCarpeta, Len(sCarpeta) - 1, 0) <> "\" Then sCarpeta = sCarpeta & "\"
            sNombreFile = "LIBRO-IVA-DIGITAL-VENTAS_" & Year(dtPeriodo) & Format(Month(dtPeriodo), "00") & ".txt"
            sNombreFile2 = "LIBRO-IVA-DIGITAL-VENTAS_ALICUOTAS_" & Year(dtPeriodo) & Format(Month(dtPeriodo), "00") & ".txt"
            sArchivoCompleto = sCarpeta & sNombreFile
            sArchivoCompleto2 = sCarpeta & sNombreFile2
            If ExisteArchivo(sArchivoCompleto) Then
                Kill sArchivoCompleto
            End If
            If ExisteArchivo(sArchivoCompleto2) Then
                Kill sArchivoCompleto2
            End If
            Open sArchivoCompleto For Output As #1
            Open sArchivoCompleto2 For Output As #2
            
            .MoveFirst
            For i = 0 To .RecordCount - 1
                vFecha = "19000101"
                vTipo = "000"
                vPunto = "00000"
                vNro = "00000000000000000000"
                vNroHasta = "00000000000000000000"
                vCodigo = "00"
                vIdentificador = "00000000000000000000"
                vApellidoCliente = "----------SIN DATOS ----------"
                vTotal = "000000000000000"
                vNoGravado = "000000000000000"
                vPercepcionNN = "000000000000000"
                vExentas = "000000000000000"
                vPercepciones = "000000000000000"
                vIngresosBrutos = "000000000000000"
                vMunicipales = "000000000000000"
                vInternos = "000000000000000"
                vCodigoMoneda = "PES"
                vCambio = "0001000000"
                vCantidadAlicuotas = "1"
                vCodOperacion = " "
                vTributos = "000000000000000"
                vVencimiento = "00000000"
            
                vNeto = "000000000000000"
                vAlicuota = "0000"
                vImpuestoLiquidado = "000000000000000"
                
                
                vFecha = Format(!Fecha, "YYYYMMDD")
                
                
                Select Case Trim(!TIPODOC)
                    Case "FAA": vTipo = "001"
                    Case "FAB": vTipo = "006"
                    Case "NCA": vTipo = "003"
                    Case "NCB": vTipo = "008"
                    Case "NDA": vTipo = "002"
                    Case "NDB": vTipo = "007"
                    Case "FAE": vTipo = "019"
                    Case "NCE": vTipo = "021"
                    Case "FEA": vTipo = "201"
                    Case "FEB": vTipo = "206"
                    Case "FEC": vTipo = "211"
                    Case "DEA": vTipo = "202"
                    Case "DEB": vTipo = "207"
                    Case "DEC": vTipo = "212"
                    Case "CEA": vTipo = "203"
                    Case "CEB": vTipo = "208"
                    Case "CEC": vTipo = "213"
                End Select
                
                If vTipo <> "019" And vTipo <> "021" Then
                    vVencimiento = Format(!Vencimiento, "YYYYMMDD")
                End If
                
                'If vTipo = "019" Then Stop
                
                vPunto = Format(!PuntoVenta, "00000")
                vNro = Format(!NroFactura, "00000000000000000000")
                vNroHasta = vNro
                
                vIdentificador = Replace(!CUIT, "-", "")
                vIdentificador = s2n(vIdentificador)
                                    
                Select Case !tipoiva
                    Case 2, 4, 7, 9, 10: vCodigo = 80 'cuit
                    Case 1: vCodigo = IIf(Len(vIdentificador) > 8, 86, 96) '96=dni,86=cuil
                    Case Else: vCodigo = 86
                End Select
                
                vIdentificador = Format(vIdentificador, "00000000000000000000")
                
                vApellidoCliente = Trim(!RAZONSOCIAL)
                If Len(vApellidoCliente) > 30 Then
                    vApellidoCliente = CORTO(vApellidoCliente, 0, Len(vApellidoCliente) - 30)
                Else
                    While Len(vApellidoCliente) < 30
                        vApellidoCliente = Chr(32) & vApellidoCliente
                    Wend
                End If
                
                vCantidadAlicuotas = obtenerDeSQL("select count(distinct(_iva)) from facturaventadetalle where codigofactura=" & !codigo)
                
                If !ND_xChequeRechazado Then
                    vCodOperacion = "A"
                ElseIf CORTO(!TIPODOC, 2, 0) = "E" Then
                    vCodOperacion = "X"
                End If
                
                
                vTotal = Format(Replace(s2n(!Total, 2, True), ",", ""), "000000000000000")
                vNoGravado = Format(Replace(s2n(!NoGrav, 2, True), ",", ""), "000000000000000")
                If vTipo = "019" Or vTipo = "021" Then
                    vExentas = vTotal 'Format(!NoGrav, "000000000000000")
                End If
                vIngresosBrutos = Format(Replace(s2n(!IIBB, 2, True), ",", ""), "000000000000000")
                
'                nnNeto = s2n(!Neto)
'                If nnNeto = 0 Then
'                    nnNeto = s2n(!total)
'                End If
'
'                If vTipo = "019" Or vTipo = "021" Then
'                    vAlicuota = "0003"
'                ElseIf !ND_xChequeRechazado Then
'                    vAlicuota = "0003"
'                    vNoGravado = Format(Replace(s2n(!Neto, 2, True), ",", ""), "000000000000000")
'                    vNeto = "000000000000000"
'                    'vImpuestoLiquidado = Format(Replace(s2n(!Iva, 2, True), ",", ""), "000000000000000")
'                    vImpuestoLiquidado = "000000000000000"
'                ElseIf s2n(s2n(!total) / s2n(nnNeto)) > 1.2 Then
'                    vAlicuota = "0005"
'                    vNeto = Format(Replace(s2n(!Neto, 2, True), ",", ""), "000000000000000")
'                    vImpuestoLiquidado = Format(Replace(s2n(!Iva, 2, True), ",", ""), "000000000000000")
'                ElseIf s2n(s2n(!total) / s2n(nnNeto)) < 1.2 Then
'                    vAlicuota = "0004"
'                    vNeto = Format(Replace(s2n(!Neto, 2, True), ",", ""), "000000000000000")
'                    vImpuestoLiquidado = Format(Replace(s2n(!Iva, 2, True), ",", ""), "000000000000000")
'                End If


                a = 0
                If vTipo = "019" Or vTipo = "021" Then
                    vAlicuota = "0003"
                    Print #2, vTipo & vPunto & vNro & vNeto & vAlicuota & vImpuestoLiquidado
'                ElseIf vTipo = "006" Or vTipo = "007" Or vTipo = "008" Then
'                    nIva21 = nSinNull(obtenerDeSQL("select (sum(preciototal)/1.21) as preciototal from facturaventadetalle where _iva=21 and codigofactura=" & !codigo))
'                    If s2n(nIva21) > 0 Then
'                        a = a + 1
'                        vAlicuota = "0005"
'                        If !ND_xChequeRechazado Then vAlicuota = "0003"
'                        vNeto = Format(Replace(s2n(nIva21, 2, True), ",", ""), "000000000000000")
'                        vImpuestoLiquidado = Format(Replace(s2n(nIva21 * 0.21, 2, True), ",", ""), "000000000000000")
'                        Print #2, vTipo & vPunto & vNro & vNeto & vAlicuota & vImpuestoLiquidado
'                    End If
'                    nIva10 = nSinNull(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where _iva=10.5 and codigofactura=" & !codigo))
'                    If s2n(nIva10) > 0 Then
'                        a = a + 1
'                        vAlicuota = "0004"
'                        If !ND_xChequeRechazado Then vAlicuota = "0003"
'                        vNeto = Format(Replace(s2n(nIva10, 2, True), ",", ""), "000000000000000")
'                        vImpuestoLiquidado = Format(Replace(s2n(nIva10 * 0.105, 2, True), ",", ""), "000000000000000")
'                        Print #2, vTipo & vPunto & vNro & vNeto & vAlicuota & vImpuestoLiquidado
'                    End If
                Else
                    If vTipo = "006" Or vTipo = "007" Or vTipo = "008" Or vTipo = "206" Or vTipo = "207" Or vTipo = "208" Then
                        nIva21 = nSinNull(obtenerDeSQL("select (sum(preciototal)/1.21) as preciototal from facturaventadetalle where _iva=21 and codigofactura=" & !codigo))
                    Else
                        nIva21 = nSinNull(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where _iva=21 and codigofactura=" & !codigo))
                    End If
                    If s2n(nIva21) > 0 Then
                        a = a + 1
                        vAlicuota = "0005"
                        If !ND_xChequeRechazado Then vAlicuota = "0003"
                        vNeto = Format(Replace(s2n(nIva21, 2, True), ",", ""), "000000000000000")
                        vImpuestoLiquidado = Format(Replace(s2n(nIva21 * 0.21, 2, True), ",", ""), "000000000000000")
                        Print #2, vTipo & vPunto & vNro & vNeto & vAlicuota & vImpuestoLiquidado
                    End If
                    
                    If vTipo = "006" Or vTipo = "007" Or vTipo = "008" Then
                        nIva10 = nSinNull(obtenerDeSQL("select (sum(preciototal)/1.105) as preciototal  from facturaventadetalle where _iva=10.5 and codigofactura=" & !codigo))
                    Else
                        nIva10 = nSinNull(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where _iva=10.5 and codigofactura=" & !codigo))
                    End If
                    If s2n(nIva10) > 0 Then
                        a = a + 1
                        vAlicuota = "0004"
                        If !ND_xChequeRechazado Then vAlicuota = "0003"
                        vNeto = Format(Replace(s2n(nIva10, 2, True), ",", ""), "000000000000000")
                        vImpuestoLiquidado = Format(Replace(s2n(nIva10 * 0.105, 2, True), ",", ""), "000000000000000")
                        Print #2, vTipo & vPunto & vNro & vNeto & vAlicuota & vImpuestoLiquidado
                    End If
                End If
                vCantidadAlicuotas = a
                
                
                
                Print #1, vFecha & vTipo & vPunto & vNro & vNroHasta & vCodigo & vIdentificador & vApellidoCliente & vTotal & vNoGravado & vPercepcionNN & vExentas & vPercepciones & vIngresosBrutos & vMunicipales & vInternos & vCodigoMoneda & vCambio & vCantidadAlicuotas & vCodOperacion & vTributos & vVencimiento
                'Print #2, vTipo & vPunto & vNro & vNeto & vAlicuota & vImpuestoLiquidado
            
                .MoveNext
            Next
        End If
    End With

    
    Close #1
    Close #2
    Set rs1 = Nothing

    MsgBox "Archivo generado correctamente.", vbInformation, "Ventas"
                
            
Exit Function
ventas_err:
MsgBox "Error... " & Err.Number & " " & Err.Description

End Function

Private Function CitiCompras2()
On Error GoTo compras_err
Dim cFecha As String, cTipo As String, cPunto As String, cNro As String, cNroDespacho As String, cCodigo As String, cIdentificador As String, cApellidoProveedor As String, cTotal As String, cNoGravado As String, cExentas As String, cPercepciones1 As String, cPercepciones2 As String, cIngresosBrutos As String, cMunicipales As String, cInternos As String, cCodigoMoneda As String, cCambio As String, cCantidadAlicuotas As String, cCodOperacion As String, cCreditoFiscal As String, cTributos As String, cCuitEmisor As String, cDenominacionEmisor As String, cComision As String
Dim cPrimerDia As Date, cUltimoDia As Date, str As String, i As Long, sNombreFile As String, sCarpeta As String, sArchivoCompleto As String, sArchivoCompleto2 As String, sArchivoCompleto3 As String, pLetra As String, pTipoDoc As String, a As Integer
Dim cNeto As String, cAlicuota As String, cImpuestoLiquidado As String
Dim dDESPACHO As String, dNETO As String, dALICUOTA As String, dIMPUESTO As String, sss As String
Dim rs1 As New ADODB.Recordset

    cPrimerDia = "01/" & Month(dtPeriodo) & "/" & Year(dtPeriodo)
    cUltimoDia = ultimoDiaDelMes(dtPeriodo)
    
    If optFecha.Value = True Then
        sss = "select TRANSCOM.fecha,codpr,nrodoc,iva_21,iva_27,0 as iva_9,iva_10,imp_int,razonsocialprov,cuitprov,suc,anoimp,TIPODOC,TOTAL,NETO,EXENTO,(IBCAPITAL+IBPROVINCIA) AS IIBB,LETRA,(PERCEPC+DER_EST+iva_9) AS RETIVA,otrasretimp as retiva2,'' as nrodespacho " & _
                    " from TRANSCOM " & _
                    " where TRANSCOM.fecha>=" & ssFecha(cPrimerDia) & " and TRANSCOM.fecha<=" & ssFecha(cUltimoDia) & " and tipodoc<>'RAC' and tipodoc<>'APC' and tipodoc<>'APD' " & _
            "union " & _
                " select compras.fecha,codpr,nrodoc,iva_21,iva_27,0 as iva_9,iva_10,imp_int,razonsocialprov,cuitprov,suc,anoimp,TIPODOC,TOTAL,NETO,EXENTO,(IBCAPITAL+IBPROVINCIA) AS IIBB,LETRA,(PERCEPC+DER_EST+iva_9) AS RETIVA,otrasretimp as retiva2,'' as  nrodespacho " & _
                    " from compras " & _
                    " where compras.fecha>=" & ssFecha(cPrimerDia) & " and compras.fecha<=" & ssFecha(cUltimoDia) & " and tipodoc<>'RAC' and tipodoc<>'APC' and tipodoc<>'APD' " & _
            "union " & _
                " select gastosbancarios.fecha,codbanco as codpr, nrodoc, iva_21 , iva_27,0 as iva_9,iva_10,0 as imp_int,razonsocialbanco as razonsocialprov, cuitbanco as cuitprov, 0 as suc,year(gastosbancarios.fecha) as fecha,tipodoc,(mantcta+gastoschqra+gastosvarios+sellado+intxgiro+valnoconfor+iva_21+iva_27+iva_10) as total,(mantcta+gastoschqra+gastosvarios+sellado+intxgiro+valnoconfor) as neto,0 as exento,0 as iibb, letra,0 as  retiva,0 as  retiva2,'' as nrodespacho " & _
                " from gastosbancarios inner join prov  on gastosbancarios.codbanco=prov.codigo " & _
                " where  gastosbancarios.fecha>=" & ssFecha(cPrimerDia) & " and gastosbancarios.fecha<=" & ssFecha(cUltimoDia) & " " & _
            "union " & _
                " select d.fecha,c.codigo as codpr,d.numero as nrodoc, d.iva as iva21,0 as iva_27,0 as iva_9,0 as iva_10,0 as imp_int,b.descripcion as razonsocialprov, c.cuit as cuitprov, 0 as suc,year(d.fecha) as fecha,'LP' as tipodoc,(d.capital + d.interes + d.iva) as total,(d.capital+d.interes) as neto,0 as exento,0 as iibb,'A' as letra,0 as  retiva,0 as  retiva2,'' as nrodespacho " & _
                " from ((prestamosldetalle as d inner join prestamosl as p on d.idprestamo=p.idprestamo) inner join CTASBANK as c on p.cuenta=c.codigo) inner join bancosgrales as b on c.banco=b.codigo " & _
                " where d.fecha>=" & ssFecha(cPrimerDia) & " and d.fecha<=" & ssFecha(cUltimoDia) & " " & _
            "union " & _
                " select d.fecha,c.codigo as codpr,d.numero as nrodoc, d.iva as iva21,0 as iva_27,0 as iva_9,0 as iva_10,0 as imp_int,b.descripcion as razonsocialprov, c.cuit as cuitprov, 0 as suc,year(d.fecha) as fecha,'LP' as tipodoc,(d.honorarios + d.iva) as total,(d.honorarios) as neto,0 as exento,0 as iibb,'A' as letra,0 as  retiva,0 as  retiva2,'' as nrodespacho " & _
                " from ((prestamosldetalleg as d inner join prestamosl as p on d.idprestamo=p.idprestamo) inner join CTASBANK as c on p.cuenta=c.codigo) inner join bancosgrales as b on c.banco=b.codigo " & _
                " where d.fecha>=" & ssFecha(cPrimerDia) & " and d.fecha<=" & ssFecha(cUltimoDia) & "  "
        rs1.Open sss, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    Else
        rs1.Open "select fecha,codpr,nrodoc,iva_21,iva_27,0 as iva_9,iva_10,imp_int,razonsocialprov,cuitprov,suc,anoimp,TIPODOC,TOTAL,NETO,EXENTO,(IBCAPITAL+IBPROVINCIA) AS IIBB,LETRA,(PERCEPC+DER_EST+iva_9) AS RETIVA,otrasretimp as retiva2,'' as nrodespacho " & _
                    " from TRANSCOM " & _
                    " where mesimp=" & Month(dtPeriodo) & " and anoimp=" & Year(dtPeriodo) & " and tipodoc<>'RAC' and tipodoc<>'APC' and tipodoc<>'APD' " & _
            "union " & _
                " select fecha,codpr,nrodoc,iva_21,iva_27,0 as iva_9,iva_10,imp_int,razonsocialprov,cuitprov,suc,anoimp,TIPODOC,TOTAL,NETO,EXENTO,(IBCAPITAL+IBPROVINCIA) AS IIBB,LETRA,(PERCEPC+DER_EST+iva_9) AS RETIVA,otrasretimp as retiva2,'' as  nrodespacho " & _
                    " from compras " & _
                    " where mesimp=" & Month(dtPeriodo) & " and anoimp=" & Year(dtPeriodo) & " and tipodoc<>'RAC' and tipodoc<>'APC' and tipodoc<>'APD' ", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    End If
    
    With rs1
        
        If rs1.EOF And rs1.BOF Then
            MsgBox "No se ha encontrado datos para este periodo.", vbInformation, "ATENCION"
            Exit Function
        Else
            
            sCarpeta = "C:\"
            sCarpeta = Trim(VentanaCarpeta("Carpeta Destino txt COMPRAS", sCarpeta))
            If Trim(sCarpeta) = "" Then Exit Function
            If CORTO(sCarpeta, Len(sCarpeta) - 1, 0) <> "\" Then sCarpeta = sCarpeta & "\"
            sNombreFile = "COMPRAS_" & Year(dtPeriodo) & Format(Month(dtPeriodo), "00") & "_" & Cuit_Empresa_Carga & ".txt"
            sNombreFile2 = "COMPRAS_ALICUOTAS_" & Year(dtPeriodo) & Format(Month(dtPeriodo), "00") & "_" & Cuit_Empresa_Carga & ".txt"
            sNombreFile3 = "DESPACHOS_ALICUOTAS_" & Year(dtPeriodo) & Format(Month(dtPeriodo), "00") & "_" & Cuit_Empresa_Carga & ".txt"
            sArchivoCompleto = sCarpeta & sNombreFile
            sArchivoCompleto2 = sCarpeta & sNombreFile2
            sArchivoCompleto3 = sCarpeta & sNombreFile3
            If ExisteArchivo(sArchivoCompleto) Then
                Kill sArchivoCompleto
            End If
            If ExisteArchivo(sArchivoCompleto2) Then
                Kill sArchivoCompleto2
            End If
            Open sArchivoCompleto For Output As #1
            Open sArchivoCompleto2 For Output As #2
            
            Open sArchivoCompleto3 For Output As #3
            
            .MoveFirst
            For i = 0 To .RecordCount - 1
            
                cFecha = "19000101"
                cTipo = "000"
                cPunto = "00000"
                cNro = "00000000000000000000"
                cNroDespacho = "                "
                cCodigo = "80"
                cIdentificador = "00000000000000000000"
                cApellidoProveedor = "----------SIN DATOS ----------"
                cTotal = "000000000000000"
                cNoGravado = "000000000000000"
                cExentas = "000000000000000"
                cPercepciones1 = "000000000000000"
                cPercepciones2 = "000000000000000"
                cIngresosBrutos = "000000000000000"
                cMunicipales = "000000000000000"
                cInternos = "000000000000000"
                cCodigoMoneda = "PES"
                cCambio = "0001000000"
                cCantidadAlicuotas = "1"
                cCodOperacion = "0"
                cCreditoFiscal = "000000000000000"
                cTributos = "000000000000000"
                cCuitEmisor = "00000000000"
                cDenominacionEmisor = "                              "
                cComision = "000000000000000"
            
                cNeto = "000000000000000"
                cAlicuota = "0000"
                cImpuestoLiquidado = "000000000000000"
                
                cFecha = Format(!Fecha, "YYYYMMDD")
                
                pLetra = Trim(sSinNull(!letra))
                If pLetra = "" Then
                    pLetra = Trim(obtenerDeSQL("select i.LETRA from prov p inner join ivas i on i.codigo=p.tipoiva where p.codigo=" & s2n(!CODPR)))
                End If
                pTipoDoc = Trim(!TIPODOC) & pLetra
                Select Case Trim(pTipoDoc)
                    Case "FACA": cTipo = "001"
                    Case "BGA": cTipo = "001"
                    Case "PLA": cTipo = "001"
                    Case "N/DA": cTipo = "002"
                    Case "N/CA": cTipo = "003"
                    Case "FACB": cTipo = "006"
                    Case "N/DB": cTipo = "007"
                    Case "N/CB": cTipo = "008"
                    Case "FACC": cTipo = "011"
                    Case "FACE": cTipo = "019"
                    Case "N/CE": cTipo = "021"
                End Select
                'cTipo = Format(Trim(!TIPODOC2), "000")
                
                cPunto = Format(!suc, "00000")
                If cPunto = "00000" Then cPunto = "00001"
                cNro = Format(!NroDoc, "00000000000000000000")
                cIdentificador = Format(Replace(!cuitprov, "-", ""), "00000000000000000000")
                
                
                cApellidoProveedor = Trim(!razonsocialprov)
                If Len(cApellidoProveedor) > 30 Then
                    cApellidoProveedor = CORTO(cApellidoProveedor, 0, Len(cApellidoProveedor) - 30)
                Else
                    While Len(cApellidoProveedor) < 30
                        cApellidoProveedor = Chr(32) & cApellidoProveedor
                    Wend
                End If
                
                cTotal = Format(Replace(s2n(!Total, 2, True), ",", ""), "000000000000000")
                cExentas = Format(Replace(s2n(!EXENTO, 2, True), ",", ""), "000000000000000")
                cNeto = Format(Replace(s2n(!Neto, 2, True), ",", ""), "000000000000000")
                If pLetra = "C" Or pLetra = "B" Then
                    cNeto = "000000000000000"
                    cNoGravado = "000000000000000"
                    cExentas = "000000000000000"
                End If
                cPercepciones1 = Format(Replace(s2n(!retIva, 2, True), ",", ""), "000000000000000")
                cPercepciones2 = Format(Replace(s2n(!retIva2, 2, True), ",", ""), "000000000000000")
                cIngresosBrutos = Format(Replace(s2n(!IIBB, 2, True), ",", ""), "000000000000000")
                cInternos = Format(Replace(s2n(!imp_int, 2, True), ",", ""), "000000000000000")
                
                a = 0
                If pLetra = "B" Or pLetra = "C" Then
                    cAlicuota = "0003"
                    Print #2, cTipo & cPunto & cNro & cCodigo & cIdentificador & cNeto & cAlicuota & cImpuestoLiquidado
                Else
                    If pLetra <> "D" Then
                        If s2n(!IVA_21) > 0 Then
                            a = a + 1
                            cAlicuota = "0005"
                            cImpuestoLiquidado = Format(Replace(s2n(!IVA_21, 2, True), ",", ""), "000000000000000")
                            cNeto = s2n(!IVA_21 / 0.21, 2, True)
                            cNeto = Format(Replace(s2n(cNeto, 2, True), ",", ""), "000000000000000")
                            Print #2, cTipo & cPunto & cNro & cCodigo & cIdentificador & cNeto & cAlicuota & cImpuestoLiquidado
                        End If
                        If s2n(!IVA_27) > 0 Then
                            a = a + 1
                            cAlicuota = "0006"
                            cImpuestoLiquidado = Format(Replace(s2n(!IVA_27, 2, True), ",", ""), "000000000000000")
                            cNeto = s2n(!IVA_27 / 0.27, 2, True)
                            cNeto = Format(Replace(s2n(cNeto, 2, True), ",", ""), "000000000000000")
                            Print #2, cTipo & cPunto & cNro & cCodigo & cIdentificador & cNeto & cAlicuota & cImpuestoLiquidado
                        End If
                        If s2n(!iva_10) > 0 Then
                            a = a + 1
                            cAlicuota = "0004"
                            cImpuestoLiquidado = Format(Replace(s2n(!iva_10, 2, True), ",", ""), "000000000000000")
                            cNeto = s2n(!iva_10 / 0.105, 2, True)
                            cNeto = Format(Replace(s2n(cNeto, 2, True), ",", ""), "000000000000000")
                            Print #2, cTipo & cPunto & cNro & cCodigo & cIdentificador & cNeto & cAlicuota & cImpuestoLiquidado
                        End If
                    End If
                End If
                cCantidadAlicuotas = a
                
                
                If pLetra = "A" And s2n(a) = 0 Then
                    cCodOperacion = "E"
                    cAlicuota = "0003"
                    cCantidadAlicuotas = "1"
                    Print #2, cTipo & cPunto & cNro & cCodigo & cIdentificador & cNeto & cAlicuota & cImpuestoLiquidado
                End If
                
                If pLetra = "D" Then
                    cCantidadAlicuotas = 1
                    cPunto = "00000"
                    dDESPACHO = "0000000000000000"
                    dNETO = "0000"
                    dALICUOTA = "000000000000000"
                    dIMPUESTO = "000000000000000"
                    
                    dDESPACHO = sSinNull(!nrodespacho)
                    cNroDespacho = dDESPACHO
                    dNETO = Format(Replace(s2n(!Neto, 2, True), ",", ""), "000000000000000")
                    dALICUOTA = "0005"
                    dIMPUESTO = Format(Replace(s2n(!IVA_21, 2, True), ",", ""), "000000000000000")
                    
                    Print #3, dDESPACHO & dNETO & dALICUOTA & dIMPUESTO
                End If
                
                
                Select Case pTipoDoc
                    Case "N/DA": If a = 0 Then cCodOperacion = "A"
                    Case "N/DB": If a = 0 Then cCodOperacion = "A"
                    Case "FACE": cCodOperacion = "X"
                    Case "N/CE": cCodOperacion = "X"
                End Select
                
                'cCuitEmisor = Format(cIdentificador, "00000000000")
                'cDenominacionEmisor = cApellidoProveedor
                
                Print #1, cFecha & cTipo & cPunto & cNro & cNroDespacho & cCodigo & cIdentificador & cApellidoProveedor & cTotal & cNoGravado & cExentas & cPercepciones1 & cPercepciones2 & cIngresosBrutos & cMunicipales & cInternos & cCodigoMoneda & cCambio & cCantidadAlicuotas & cCodOperacion & cCreditoFiscal & cTributos & cCuitEmisor & cDenominacionEmisor & cComision
                
                .MoveNext
            Next
            
            
            

        End If
        
    End With
    
    
    Close #1
    Close #2
    Close #3
    
    Set rs1 = Nothing

    MsgBox "Archivo generado correctamente.", vbInformation, "Compras"

Exit Function
compras_err:
MsgBox "Error... " & Err.Number & " " & Err.Description

End Function

Private Function CitiVentas2()
On Error GoTo ventas_err
Dim vFecha As String, vTipo As String, vPunto As String, vNro As String, vNroHasta As String, vCodigo As String, vIdentificador As String, vApellidoCliente As String, vTotal As String, vNoGravado As String, vPercepcionNN As String, vExentas As String, vPercepciones As String, vIngresosBrutos As String, vMunicipales As String, vInternos As String, vCodigoMoneda As String, vCambio As String, vCantidadAlicuotas As String, vCodOperacion As String, vTributos As String, vVencimiento As String
Dim vPrimerDia As Date, vUltimoDia As Date, str As String, i As Long, sNombreFile2 As String, sNombreFile As String, sCarpeta As String, sArchivoCompleto As String, sArchivoCompleto2 As String
Dim vNeto As String, vAlicuota As String, vImpuestoLiquidado As String, nnNeto As Double, a As Integer
Dim rs1 As New ADODB.Recordset

    vPrimerDia = "01/" & Month(dtPeriodo) & "/" & Year(dtPeriodo)
    vUltimoDia = ultimoDiaDelMes(dtPeriodo)
    str = "select * from facturaventa where fecha>=" & ssFecha(vPrimerDia) & " and fecha<=" & ssFecha(vUltimoDia) & " and tipodoc in ('NDB','NDA','FAE','FAA','FAB','NCA','NCB','FEA','FEB','FEC','DEA','DEB','DEC','CEA','CEB','CEC') order by fecha,tipodoc,nrofactura"
    rs1.Open str, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly

    With rs1
        If .EOF And .BOF Then
            MsgBox "No se ha encontrado datos para este periodo.", vbInformation, "ATENCION"
            Exit Function
        Else
            sCarpeta = "C:\"
            sCarpeta = Trim(VentanaCarpeta("Carpeta Destino txt VENTAS", sCarpeta))
            If Trim(sCarpeta) = "" Then Exit Function
            If CORTO(sCarpeta, Len(sCarpeta) - 1, 0) <> "\" Then sCarpeta = sCarpeta & "\"
            sNombreFile = "VENTAS_" & Year(dtPeriodo) & Format(Month(dtPeriodo), "00") & ".txt"
            sNombreFile2 = "VENTAS_ALICUOTAS_" & Year(dtPeriodo) & Format(Month(dtPeriodo), "00") & ".txt"
            sArchivoCompleto = sCarpeta & sNombreFile
            sArchivoCompleto2 = sCarpeta & sNombreFile2
            If ExisteArchivo(sArchivoCompleto) Then
                Kill sArchivoCompleto
            End If
            If ExisteArchivo(sArchivoCompleto2) Then
                Kill sArchivoCompleto2
            End If
            Open sArchivoCompleto For Output As #1
            Open sArchivoCompleto2 For Output As #2
            
            .MoveFirst
            For i = 0 To .RecordCount - 1
                vFecha = "19000101"
                vTipo = "000"
                vPunto = "00000"
                vNro = "00000000000000000000"
                vNroHasta = "00000000000000000000"
                vCodigo = "00"
                vIdentificador = "00000000000000000000"
                vApellidoCliente = "----------SIN DATOS ----------"
                vTotal = "000000000000000"
                vNoGravado = "000000000000000"
                vPercepcionNN = "000000000000000"
                vExentas = "000000000000000"
                vPercepciones = "000000000000000"
                vIngresosBrutos = "000000000000000"
                vMunicipales = "000000000000000"
                vInternos = "000000000000000"
                vCodigoMoneda = "PES"
                vCambio = "0001000000"
                vCantidadAlicuotas = "1"
                vCodOperacion = "0"
                vTributos = "000000000000000"
                vVencimiento = "00000000"
            
                vNeto = "000000000000000"
                vAlicuota = "0000"
                vImpuestoLiquidado = "000000000000000"
                
                
                vFecha = Format(!Fecha, "YYYYMMDD")
                
                
                Select Case Trim(!TIPODOC)
                    Case "FAA": vTipo = "001"
                    Case "FAB": vTipo = "006"
                    Case "NCA": vTipo = "003"
                    Case "NCB": vTipo = "008"
                    Case "NDA": vTipo = "002"
                    Case "NDB": vTipo = "007"
                    Case "FAE": vTipo = "019"
                    Case "NCE": vTipo = "021"
                    Case "FEA": vTipo = "201"
                    Case "FEB": vTipo = "206"
                    Case "FEC": vTipo = "211"
                    Case "DEA": vTipo = "202"
                    Case "DEB": vTipo = "207"
                    Case "DEC": vTipo = "212"
                    Case "CEA": vTipo = "203"
                    Case "CEB": vTipo = "208"
                    Case "CEC": vTipo = "213"
                End Select
                
                If vTipo <> "019" And vTipo <> "021" Then
                    vVencimiento = Format(!Vencimiento, "YYYYMMDD")
                End If
                
                'If vTipo = "019" Then Stop
                
                vPunto = Format(!PuntoVenta, "00000")
                vNro = Format(!NroFactura, "00000000000000000000")
                vNroHasta = vNro
                
                vIdentificador = Replace(!CUIT, "-", "")
                vIdentificador = s2n(vIdentificador)
                                    
                Select Case !tipoiva
                    Case 2, 4, 7, 9, 10: vCodigo = 80 'cuit
                    Case 1: vCodigo = IIf(Len(vIdentificador) > 8, 86, 96) '96=dni,86=cuil
                    Case Else: vCodigo = 86
                End Select
                
                vIdentificador = Format(vIdentificador, "00000000000000000000")
                
                vApellidoCliente = Trim(!RAZONSOCIAL)
                If Len(vApellidoCliente) > 30 Then
                    vApellidoCliente = CORTO(vApellidoCliente, 0, Len(vApellidoCliente) - 30)
                Else
                    While Len(vApellidoCliente) < 30
                        vApellidoCliente = Chr(32) & vApellidoCliente
                    Wend
                End If
                
                vCantidadAlicuotas = obtenerDeSQL("select count(distinct(_iva)) from facturaventadetalle where codigofactura=" & !codigo)
                
                If !ND_xChequeRechazado Then
                    vCodOperacion = "A"
                ElseIf CORTO(!TIPODOC, 2, 0) = "E" Then
                    vCodOperacion = "X"
                End If
                
                
                vTotal = Format(Replace(s2n(!Total, 2, True), ",", ""), "000000000000000")
                vNoGravado = Format(Replace(s2n(!NoGrav, 2, True), ",", ""), "000000000000000")
                If vTipo = "019" Or vTipo = "021" Then
                    vExentas = vTotal 'Format(!NoGrav, "000000000000000")
                End If
                vIngresosBrutos = Format(Replace(s2n(!IIBB, 2, True), ",", ""), "000000000000000")
                
'                nnNeto = s2n(!Neto)
'                If nnNeto = 0 Then
'                    nnNeto = s2n(!total)
'                End If
'
'                If vTipo = "019" Or vTipo = "021" Then
'                    vAlicuota = "0003"
'                ElseIf !ND_xChequeRechazado Then
'                    vAlicuota = "0003"
'                    vNoGravado = Format(Replace(s2n(!Neto, 2, True), ",", ""), "000000000000000")
'                    vNeto = "000000000000000"
'                    'vImpuestoLiquidado = Format(Replace(s2n(!Iva, 2, True), ",", ""), "000000000000000")
'                    vImpuestoLiquidado = "000000000000000"
'                ElseIf s2n(s2n(!total) / s2n(nnNeto)) > 1.2 Then
'                    vAlicuota = "0005"
'                    vNeto = Format(Replace(s2n(!Neto, 2, True), ",", ""), "000000000000000")
'                    vImpuestoLiquidado = Format(Replace(s2n(!Iva, 2, True), ",", ""), "000000000000000")
'                ElseIf s2n(s2n(!total) / s2n(nnNeto)) < 1.2 Then
'                    vAlicuota = "0004"
'                    vNeto = Format(Replace(s2n(!Neto, 2, True), ",", ""), "000000000000000")
'                    vImpuestoLiquidado = Format(Replace(s2n(!Iva, 2, True), ",", ""), "000000000000000")
'                End If


                a = 0
                If vTipo = "019" Or vTipo = "021" Then
                    vAlicuota = "0003"
                    Print #2, vTipo & vPunto & vNro & vNeto & vAlicuota & vImpuestoLiquidado
'                ElseIf vTipo = "006" Or vTipo = "007" Or vTipo = "008" Then
'                    nIva21 = nSinNull(obtenerDeSQL("select (sum(preciototal)/1.21) as preciototal from facturaventadetalle where _iva=21 and codigofactura=" & !codigo))
'                    If s2n(nIva21) > 0 Then
'                        a = a + 1
'                        vAlicuota = "0005"
'                        If !ND_xChequeRechazado Then vAlicuota = "0003"
'                        vNeto = Format(Replace(s2n(nIva21, 2, True), ",", ""), "000000000000000")
'                        vImpuestoLiquidado = Format(Replace(s2n(nIva21 * 0.21, 2, True), ",", ""), "000000000000000")
'                        Print #2, vTipo & vPunto & vNro & vNeto & vAlicuota & vImpuestoLiquidado
'                    End If
'                    nIva10 = nSinNull(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where _iva=10.5 and codigofactura=" & !codigo))
'                    If s2n(nIva10) > 0 Then
'                        a = a + 1
'                        vAlicuota = "0004"
'                        If !ND_xChequeRechazado Then vAlicuota = "0003"
'                        vNeto = Format(Replace(s2n(nIva10, 2, True), ",", ""), "000000000000000")
'                        vImpuestoLiquidado = Format(Replace(s2n(nIva10 * 0.105, 2, True), ",", ""), "000000000000000")
'                        Print #2, vTipo & vPunto & vNro & vNeto & vAlicuota & vImpuestoLiquidado
'                    End If
                Else
                    If vTipo = "006" Or vTipo = "007" Or vTipo = "008" Or vTipo = "206" Or vTipo = "207" Or vTipo = "208" Then
                        nIva21 = nSinNull(obtenerDeSQL("select (sum(preciototal)/1.21) as preciototal from facturaventadetalle where _iva=21 and codigofactura=" & !codigo))
                    Else
                        nIva21 = nSinNull(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where _iva=21 and codigofactura=" & !codigo))
                    End If
                    If s2n(nIva21) > 0 Then
                        a = a + 1
                        vAlicuota = "0005"
                        If !ND_xChequeRechazado Then vAlicuota = "0003"
                        vNeto = Format(Replace(s2n(nIva21, 2, True), ",", ""), "000000000000000")
                        vImpuestoLiquidado = Format(Replace(s2n(nIva21 * 0.21, 2, True), ",", ""), "000000000000000")
                        Print #2, vTipo & vPunto & vNro & vNeto & vAlicuota & vImpuestoLiquidado
                    End If
                    
                    If vTipo = "006" Or vTipo = "007" Or vTipo = "008" Then
                        nIva10 = nSinNull(obtenerDeSQL("select (sum(preciototal)/1.105) as preciototal  from facturaventadetalle where _iva=10.5 and codigofactura=" & !codigo))
                    Else
                        nIva10 = nSinNull(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where _iva=10.5 and codigofactura=" & !codigo))
                    End If
                    If s2n(nIva10) > 0 Then
                        a = a + 1
                        vAlicuota = "0004"
                        If !ND_xChequeRechazado Then vAlicuota = "0003"
                        vNeto = Format(Replace(s2n(nIva10, 2, True), ",", ""), "000000000000000")
                        vImpuestoLiquidado = Format(Replace(s2n(nIva10 * 0.105, 2, True), ",", ""), "000000000000000")
                        Print #2, vTipo & vPunto & vNro & vNeto & vAlicuota & vImpuestoLiquidado
                    End If
                End If
                vCantidadAlicuotas = a
                
                
                
                Print #1, vFecha & vTipo & vPunto & vNro & vNroHasta & vCodigo & vIdentificador & vApellidoCliente & vTotal & vNoGravado & vPercepcionNN & vExentas & vPercepciones & vIngresosBrutos & vMunicipales & vInternos & vCodigoMoneda & vCambio & vCantidadAlicuotas & vCodOperacion & vTributos & vVencimiento
                'Print #2, vTipo & vPunto & vNro & vNeto & vAlicuota & vImpuestoLiquidado
            
                .MoveNext
            Next
        End If
    End With

    
    Close #1
    Close #2
    Set rs1 = Nothing

    MsgBox "Archivo generado correctamente.", vbInformation, "Ventas"
                
            
Exit Function
ventas_err:
MsgBox "Error... " & Err.Number & " " & Err.Description
End Function

Private Sub cmdCiti_Click()
    If MsgBox("¿Generar Ventas?", vbYesNo + vbInformation) = vbYes Then CitiVentasB
    If MsgBox("¿Generar Compras?", vbYesNo + vbInformation) = vbYes Then CitiComprasB

End Sub

Private Sub cmdLibroDigitalVentasCompras_Click()
    
    If MsgBox("¿Generar Ventas?", vbYesNo + vbInformation) = vbYes Then LibroDigitalVentas
    If MsgBox("¿Generar Compras?", vbYesNo + vbInformation) = vbYes Then LibroDigitalCompras
End Sub

Private Sub Command1_Click()  'LIBRO DE VENTAS
    Dim ARCH As String
    Dim documento As Integer
    Dim nom As String
    Dim tI As Double
    Dim impLIQ As String
    Dim impRNI As String
    Dim impEXE As String
    Dim totimpLIQ As Double
    Dim totimpRNI As Double
    Dim totimpEXE As Double
    Dim IIBB As Double
    Dim fec As String
    Dim pri As Date
    Dim seg As Date
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim rs3 As New ADODB.Recordset
    Dim i As Long
    Dim j As Long
    Dim tot As Double
    Dim netgra As Double
    Dim netNOgra As Double
    Dim rni As Double
    Dim relleno As String
    Dim relleno2 As String
    Dim a
    
    
    pri = "01/" & Month(dtPeriodo.Value) & "/" & Year(dtPeriodo.Value)
    seg = ultimoDiaDelMes(dtPeriodo.Value)
    a = "select * from facturaventa where fecha>=" & ssFecha(pri) & " and fecha<=" & ssFecha(seg) & " and (tipodoc='FAA' or tipodoc='FAB' or tipodoc='NCA' or tipodoc='NCB') order by fecha,tipodoc,nrofactura"
    rs3.Open "select * from facturaventa where fecha>=" & ssFecha(pri) & " and fecha<=" & ssFecha(seg) & " and (tipodoc='FAA' or tipodoc='FAB' or tipodoc='NCA' or tipodoc='NCB') order by fecha,tipodoc,nrofactura", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    
'********************************** ARCHIVO

    ARCH = "VENTAS_" & Year(dtPeriodo.Value) & Format(Month(dtPeriodo.Value), "00")
    rs.Open "select * from datosempresa where idempresa=15", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    
    If (rs3.EOF = True And rs3.BOF = True) Or IsNull(rs3!codigo) Then
        MsgBox "No se ha encontrado datos para este periodo.", , "ATENCION"
        Exit Sub
    End If
    
    rs3.MoveFirst
    If Not IsNull(rs3) Or Not IsEmpty(rs3) Or Not (rs3.EOF = True And rs3.BOF = True) Then
        Open "C:\gestion\SIRED\" & ARCH & ".txt" For Output As #1
    End If
    i = 0
    While Not rs3.EOF
    
        rs2.Open "select i.*,C.CUIT from clientes c inner join ivas i on i.codigo=c.iva where c.codigo=" & rs3!cliente, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        If rs3!tipoiva = 1 Or rs3!tipoiva = 4 Or rs3!tipoiva = 6 Then
        'If ComboCodigo(cmbTipoIva) = 1 Or ComboCodigo(cmbTipoIva) = 4 Or ComboCodigo(cmbTipoIva) = 6 Then
            documento = 80
        Else
            documento = 86
        End If
        nom = Trim(rs3!RAZONSOCIAL)
        While Len(nom) < 30
            nom = Chr(32) & nom
        Wend
        
        fec = Year(rs3!Fecha) & Format(Month(rs3!Fecha), "00") & Format(Day(rs3!Fecha), "00")
        'ver si tengo que multiplicar * 100 los valores para sacarle la coma!!!!!!
        tI = Trim(rs3!tipoiva)
        If tI = 5 Or tI = 4 Or tI = 6 Or tI = 10 Then 'facturas B
            impLIQ = "000000000000000"
            impRNI = "000000000000000"
            impEXE = "000000000000000" 'Format(TxtIVA * 100, "000000000000000")
        ElseIf tI = 1 Or tI = 2 Then  'facturas A
            If tI = 1 Then
                impLIQ = Format(rs3!Iva * 100, "000000000000000")
                impRNI = "000000000000000"
                impEXE = "000000000000000"
            ElseIf tI = 2 Then
                impLIQ = "000000000000000"
                impRNI = Format(rs3!Iva * 100, "000000000000000")
                impEXE = "000000000000000"
            ElseIf tI = 8 Then
                impLIQ = "000000000000000"
                impRNI = "000000000000000"
                impEXE = Format(rs3!Iva * 100, "000000000000000")
            End If
        End If
        
        If Trim(rs3!TIPODOC) = "FAA" Or Trim(rs3!TIPODOC) = "FAB" Or Trim(rs3!TIPODOC) = "NDA" Or Trim(rs3!TIPODOC) = "NDB" Then
            tot = tot + CDbl(rs3!Total)
            totimpLIQ = totimpLIQ + CDbl(impLIQ)
            totimpRNI = totimpRNI + CDbl(impRNI)
            totimpEXE = totimpEXE + CDbl(impEXE)
            IIBB = IIBB + CDbl(rs3!IIBB)
        ElseIf Trim(rs3!TIPODOC) = "NCA" Or Trim(rs3!TIPODOC) = "NCB" Then
            tot = tot - CDbl(rs3!Total)
            totimpLIQ = totimpLIQ - CDbl(impLIQ)
            totimpRNI = totimpRNI - CDbl(impRNI)
            totimpEXE = totimpEXE - CDbl(impEXE)
            IIBB = IIBB - CDbl(rs3!IIBB)
        End If
        
        
        If Trim(rs3!TIPODOC) = "FAA" Then  'en las tablas de la afip el 01 es factura A
            'Open "C:\gestion\SIRED\" & arch & ".txt" For Output As #1
            'Write #1, cmbCliente.Text & cmbCliente.Text & Chr(32) & cmbCliente.Text            ' Text1.Text el chr(9)=tab  vbCrLf=enter   chr(32)=un espacio
            relleno2 = ""
            j = 1
            While j < 61
                relleno2 = relleno2 & Chr(32)
                j = j + 1
            Wend
            Print #1, "1" & fec & "01" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & "2100" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & Chr(32) & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
            "Sin comentarios" & relleno2
            'Close #1
            netgra = netgra + CDbl(rs3!Neto)
            i = i + 1
        ElseIf Trim(rs3!TIPODOC) = "FAB" Then
            'If CDbl(txtTotal.caption) >= CDbl("10000") Then  'lo quieren hacer uno por uno
                'Open "C:\gestion\SIRED\" & arch & ".txt" For Output As #1
                'Write #1, cmbCliente.Text & cmbCliente.Text & Chr(32) & cmbCliente.Text            ' Text1.Text el chr(9)=tab  vbCrLf=enter
                relleno2 = ""
                j = 1
                While j < 61
                    relleno2 = relleno2 & Chr(32)
                    j = j + 1
                Wend
                Print #1, "1" & fec & "06" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & "E" & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                "Sin comentarios" & relleno2 '& vbCrLf
                'txtneto
                netNOgra = netNOgra + CDbl(rs3!Total)
                'Close #1
            'End If
            i = i + 1
        ElseIf Trim(rs3!TIPODOC) = "NCA" Then
            relleno2 = ""
            j = 1
            While j < 61
                relleno2 = relleno2 & Chr(32)
                j = j + 1
            Wend
            Print #1, "1" & fec & "03" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & "2100" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & Chr(32) & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
            "Sin comentarios" & relleno2
            netgra = netgra - CDbl(rs3!Neto)
            i = i + 1
        ElseIf Trim(rs3!TIPODOC) = "NDA" Then
            relleno2 = ""
            j = 1
            While j < 61
                relleno2 = relleno2 & Chr(32)
                j = j + 1
            Wend
            Print #1, "1" & fec & "02" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & "2100" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & Chr(32) & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
            "Sin comentarios" & relleno2
            netgra = netgra + CDbl(rs3!Neto)
            i = i + 1
        ElseIf Trim(rs3!TIPODOC) = "NCB" Then
            relleno2 = ""
            j = 1
            While j < 61
                relleno2 = relleno2 & Chr(32)
                j = j + 1
            Wend
            Print #1, "1" & fec & "08" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & "E" & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                "Sin comentarios" & relleno2 '& vbCrLf
            netNOgra = netNOgra - CDbl(rs3!Total)
            i = i + 1
        ElseIf Trim(rs3!TIPODOC) = "NDB" Then
            relleno2 = ""
            j = 1
            While j < 61
                relleno2 = relleno2 & Chr(32)
                j = j + 1
            Wend
            Print #1, "1" & fec & "07" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & "E" & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                "Sin comentarios" & relleno2 '& vbCrLf
            netNOgra = netNOgra + CDbl(rs3!Total)
            i = i + 1
        End If
        
        
        rs3.MoveNext
        Set rs2 = Nothing
    Wend
    'registro de tipo 2
    j = 1
    While j < 123
        relleno = relleno & Chr(32)
        j = j + 1
    Wend
    relleno2 = ""
    j = 1
    While j < 30
        relleno2 = relleno2 & Chr(32)
        j = j + 1
    Wend
    
    Print #1, "2" & Year(dtPeriodo.Value) & Format(Month(dtPeriodo.Value), "00") & relleno2 & Format(i, "000000000000") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Trim(Mid(rs!CuitEmpresa, 1, 2)) & Trim(Mid(rs!CuitEmpresa, 4, 8)) & Trim(Mid(rs!CuitEmpresa, 13, 1)) & relleno2 & Chr(32) & _
     Format(tot * 100, "000000000000000") & Format(netNOgra * 100, "000000000000000") & Format(netgra * 100, "000000000000000") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Format(totimpLIQ, "000000000000000") & Format(totimpRNI, "000000000000000") & Format(totimpEXE, "000000000000000") & "000000000000000" & Format(IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & relleno
    
    
    If Not IsNull(rs3) Or Not IsEmpty(rs3) Or Not (rs3.EOF = True And rs3.BOF = True) Then
        Close #1
    End If
    Set rs3 = Nothing
    Set rs = Nothing
    MsgBox "Archivo generado correctamente."
                
            
    '******************************************************************************************

End Sub

Private Sub Command2_Click() 'LIBRO DE COMPRAS
    Dim ARCH As String
    Dim documento As Integer
    Dim nom As String
    Dim tI As Double
    Dim impLIQ As String
    Dim impRNI As String
    Dim impEXE As String
    Dim totimpLIQ As Double
    Dim totimpRNI As Double
    Dim totimpEXE As Double
    Dim totPerAgr As Double
    Dim totPerOtro As Double
    Dim totImpInt As Double
    Dim IIBB As Double
    Dim fec As String
    Dim pri As Date
    Dim seg As Date
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim rs3 As New ADODB.Recordset
    Dim i As Long
    Dim j As Long
    Dim x As Long
    Dim p As Long
    Dim tot As Double
    Dim netgra As Double
    Dim netNOgra As Double
    Dim rni As Double
    Dim relleno As String
    Dim relleno2 As String
    Dim val As Double
    Dim val2 As Double
    Dim val3 As Double
    Dim val4 As Double
    Dim val5 As Double
    Dim val6 As Double
    Dim val7 As Double
    Dim val8 As Double
    Dim val9 As Double
    Dim a As Integer
    Dim fecCAI As String
    Dim cui As String
    
    
    pri = "01/" & Month(dtPeriodo.Value) & "/" & Year(dtPeriodo.Value)
    seg = ultimoDiaDelMes(dtPeriodo.Value)
    rs3.Open "select fecha,codpr,tipoiva,tipodoc,nrodoc,total,neto,exento,iva_21,iva_27,iva_9,iva_10,imp_int,percepc,der_est,ibcapital,ibprovincia,retgan,retganpago,formadepago,razonsocialprov,cuitprov,controlador,suc,anoimp,aduana,destinacion,verifidespacho,despacho,cai,vencecai from TRANSCOM where fecha>=" & ssFecha(pri) & " and fecha<=" & ssFecha(seg) & " and tipodoc<>'RAC'" & _
        "union select fecha,codpr,tipoiva,tipodoc,nrodoc,total,neto,exento,iva_21,iva_27,iva_9,iva_10,imp_int,percepc,der_est,ibcapital,ibprovincia,retgan,retganpago,formadepago,razonsocialprov,cuitprov,controlador,suc,anoimp,aduana,destinacion,verifidespacho,despacho,cai,vencecai from compras where fecha>=" & ssFecha(pri) & " and fecha<=" & ssFecha(seg) & " and tipodoc<>'RAC' order by fecha,tipodoc,nrodoc", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    
    'MsgBox "" & rs3.RecordCount
'********************************** ARCHIVO
                ARCH = "COMPRAS_" & Year(dtPeriodo.Value) & Format(Month(dtPeriodo.Value), "00")
                rs.Open "select * from datosempresa where idempresa=15", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                
                If (rs3.EOF = True And rs3.BOF = True) Or IsNull(rs3) Then
                    MsgBox "No se ha encontrado datos para este periodo.", , "ATENCION"
                    Exit Sub
                End If
                
                rs3.MoveFirst
                If Not IsNull(rs3) Or Not IsEmpty(rs3) Or Not (rs3.EOF = True And rs3.BOF = True) Then
                    Open "C:\gestion\SIRED\" & ARCH & ".txt" For Output As #1
                End If
                i = 0
                While Not rs3.EOF
                
                    rs2.Open "select i.*,C.CUIT from prov c inner join ivas i on i.codigo=c.tipoiva where c.codigo=" & rs3!CODPR, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                    a = 0
                    If rs3!tipoiva = 1 Or rs3!tipoiva = 4 Or rs3!tipoiva = 6 Then
                    'If ComboCodigo(cmbTipoIva) = 1 Or ComboCodigo(cmbTipoIva) = 4 Or ComboCodigo(cmbTipoIva) = 6 Then
                        documento = 80
                    Else
                        documento = 86
                    End If
                    
                    If rs3!cuitprov = "" Then
                        cui = Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32)
                        documento = 99
                    Else
                        cui = Trim(Mid(rs3!cuitprov, 1, 2)) & Trim(Mid(rs3!cuitprov, 4, 8)) & Trim(Mid(rs3!cuitprov, 13, 1))
                    End If
                    
                    nom = Trim(rs3!razonsocialprov)
                    If Trim(rs3!razonsocialprov) = "Nextel" Then
                        nom = ""
                    End If
                    While Len(nom) < 30
                        nom = Chr(32) & nom
                    Wend
                    
                    fec = Year(rs3!Fecha) & Format(Month(rs3!Fecha), "00") & Format(Day(rs3!Fecha), "00")
                    'ver si tengo que multiplicar * 100 los valores para sacarle la coma!!!!!!
                    tI = Trim(rs3!tipoiva)
                    If tI = 5 Or tI = 4 Or tI = 6 Or tI = 10 Then 'facturas B
                        If tI = 5 Or tI = 6 Then
                            impLIQ = "000000000000000"
                            impRNI = "000000000000000"
                            impEXE = "000000000000000" 'Format(TxtIVA * 100, "000000000000000")
                        ElseIf tI = 4 Or tI = 10 Then
                            impLIQ = "000000000000000"
                            impRNI = "000000000000000"
                            impEXE = Format(rs3!EXENTO * 100, "000000000000000")
                        End If
                    ElseIf tI = 1 Or tI = 2 Then  'facturas A
                        If tI = 1 Then
                            impLIQ = Format(rs3!IVA_21 * 100, "000000000000000")
                            impRNI = "000000000000000"
                            impEXE = "000000000000000"
                        ElseIf tI = 2 Then
                            impLIQ = "000000000000000"
                            impRNI = Format(rs3!IVA_21 * 100, "000000000000000")
                            impEXE = "000000000000000"
                        
                        End If
                    End If
                    
                    tot = tot + CDbl(rs3!Total)
                    totimpLIQ = totimpLIQ + CDbl((rs3!IVA_21 + rs3!iva_10 + rs3!IVA_27) * 100)
                    totimpRNI = totimpRNI + CDbl(impRNI)
                    totimpEXE = totimpEXE + CDbl(rs3!EXENTO * 100)
                    totPerAgr = totPerAgr + CDbl(rs3!percepc * 100)
                    totPerOtro = totPerOtro + CDbl(rs3!der_est * 100)
                    IIBB = IIBB + CDbl((rs3!ibcapital + rs3!ibprovincia) * 100)
                    totImpInt = totImpInt + CDbl(rs3!imp_int * 100)
                    
                    x = 0
                    If rs3!IVA_21 > 0 Then x = x + 1
                    If rs3!IVA_27 > 0 Then x = x + 1
                    If rs3!iva_10 > 0 Then x = x + 1
                    
                    j = 1
                    relleno = "Sin comentarios"
                    While j < 61
                        relleno = relleno & Chr(32)
                        j = j + 1
                    Wend
                    
                    fecCAI = IIf(Trim(rs3!vencecai) = "01/01/1900", "00000000", Year(rs3!vencecai) & Format(Month(rs3!vencecai), "00") & Format(Day(rs3!vencecai), "00"))
                    
                    'If tI = 1 Or tI = 2 Then 'Trim(rs3!TIPODOC) = "FAA" Then 'en las tablas de la afip el 01 es factura A
                        'Open "C:\gestion\SIRED\" & arch & ".txt" For Output As #1
                        'Write #1, cmbCliente.Text & cmbCliente.Text & Chr(32) & cmbCliente.Text            ' Text1.Text el chr(9)=tab  vbCrLf=enter   chr(32)=un espacio
                        If x = 0 Then
                            Print #1, "1" & fec & "01" & IIf(rs3!controlador = 1, "C", Chr(32)) & Format(rs3!suc, "0000") & Format(rs3!NroDoc, "00000000000000000000") & fec & Format(IIf(Trim(rs3!aduana) = "", 0, Trim(rs3!aduana)), "000") & IIf(IsNull(rs3!destinacion) Or Trim(rs3!destinacion) = "", "    ", Trim(rs3!destinacion)) & Format(Trim(rs3!despacho), "000000") & IIf(Trim(rs3!verifidespacho) = "", Chr(32), Trim(rs3!verifidespacho)) & documento & cui & nom & _
                                Format(rs3!Total * 100, "000000000000000") & Format(rs3!EXENTO * 100, "000000000000000") & Format(rs3!Neto * 100, "000000000000000") & "0000" & impLIQ & Format(rs3!EXENTO * 100, "000000000000000") & Format(rs3!percepc * 100, "000000000000000") & Format(rs3!der_est * 100, "000000000000000") & Format((rs3!ibcapital + rs3!ibprovincia) * 100, "000000000000000") & "000000000000000" & Format(rs3!imp_int * 100, "000000000000000") & _
                                Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & "E" & Format(rs3!cai, "00000000000000") & fecCAI & relleno
                            i = i + 1
                        Else
                            For p = 1 To x
                                If x = p Then
                                    val = rs3!Total * 100
                                    val2 = rs3!EXENTO * 100  '(rs3!EXENTO + rs3!imp_int + rs3!IVA_21 + rs3!IVA_27 + rs3!iva_10) * 100 '
                                    val3 = rs3!Neto * 100
                                    val4 = rs3!imp_int * 100
                                    val5 = rs3!EXENTO * 100
                                    'val6 = (rs3!IVA_21 + rs3!IVA_27 + rs3!iva_10) * 100
                                    val7 = rs3!percepc
                                    val8 = rs3!der_est
                                    val9 = (rs3!ibcapital + rs3!ibprovincia) * 100
                                Else
                                    val = 0 ' SE MUESTRA SOLO EN UN REGISTRO EL TOTAL QUE ES EL ULTIMO EN EL CASO DE TENER VARIAS ALICUOTAS
                                    val2 = 0
                                    val3 = 0
                                    val4 = 0
                                    val5 = 0
                                    'val6 = 0
                                    val7 = 0
                                    val8 = 0
                                    val9 = 0
                                End If
                                If x > 0 Then
                                    If rs3!IVA_21 > 0 And a = 0 Then
                                        val6 = rs3!IVA_21 * 100
                                        'val2 = rs3!IVA_21 * 100
                                    ElseIf rs3!IVA_27 > 0 And (a = 1 Or a = 0) Then
                                        val6 = rs3!IVA_27 * 100
                                        'val2 = rs3!IVA_27 * 100
                                    ElseIf rs3!iva_10 And (a = 2 Or a = 1 Or a = 0) Then
                                        val6 = rs3!iva_10 * 100
                                        'val2 = rs3!iva_10 * 100
                                    End If
                                Else
                                    val6 = 0
                                    'val2 = (CDbl(rs3!EXENTO) + CDbl(rs3!imp_int) + CDbl(rs3!IVA_21) + CDbl(rs3!IVA_27) + CDbl(rs3!iva_10)) * 100
                                End If
                                
                                If rs3!IVA_21 > 0 And a = 0 Then
                                    'If val3 = 0 Then val3 = rs3!IVA_21
                                    Print #1, "1" & fec & "01" & IIf(rs3!controlador = 1, "C", Chr(32)) & Format(rs3!suc, "0000") & Format(rs3!NroDoc, "00000000000000000000") & fec & Format(IIf(Trim(rs3!aduana) = "", 0, Trim(rs3!aduana)), "000") & IIf(IsNull(rs3!destinacion) Or Trim(rs3!destinacion) = "", "    ", Trim(rs3!destinacion)) & Format(Trim(rs3!despacho), "000000") & IIf(Trim(rs3!verifidespacho) = "", Chr(32), Trim(rs3!verifidespacho)) & documento & cui & nom & _
                                        Format(val, "000000000000000") & Format(val2, "000000000000000") & Format(val3, "000000000000000") & "2100" & Format(val6, "000000000000000") & Format(val5, "000000000000000") & Format(val7 * 100, "000000000000000") & Format(val8 * 100, "000000000000000") & Format(val9, "000000000000000") & "000000000000000" & Format(val4, "000000000000000") & _
                                        Format(rs3!tipoiva, "00") & "PES" & "0001000000" & Format(x, "0") & Chr(32) & Format(rs3!cai, "00000000000000") & fecCAI & relleno
                                    a = a + 1
                                    i = i + 1
                                Else
                                    If rs3!IVA_27 > 0 And (a = 1 Or a = 0) Then
                                        'If val3 = 0 Then val3 = rs3!IVA_27
                                        Print #1, "1" & fec & "01" & IIf(rs3!controlador = 1, "C", Chr(32)) & Format(rs3!suc, "0000") & Format(rs3!NroDoc, "00000000000000000000") & fec & Format(IIf(Trim(rs3!aduana) = "", 0, Trim(rs3!aduana)), "000") & IIf(IsNull(rs3!destinacion) Or Trim(rs3!destinacion) = "", "    ", Trim(rs3!destinacion)) & Format(Trim(rs3!despacho), "000000") & IIf(Trim(rs3!verifidespacho) = "", Chr(32), Trim(rs3!verifidespacho)) & documento & cui & nom & _
                                            Format(val, "000000000000000") & Format(val2, "000000000000000") & Format(val3, "000000000000000") & "2700" & Format(val6, "000000000000000") & Format(val5, "000000000000000") & Format(val7 * 100, "000000000000000") & Format(val8 * 100, "000000000000000") & Format(val9, "000000000000000") & "000000000000000" & Format(val4, "000000000000000") & _
                                            Format(rs3!tipoiva, "00") & "PES" & "0001000000" & Format(x, "0") & Chr(32) & Format(rs3!cai, "00000000000000") & fecCAI & relleno
                                        a = a + 1
                                        i = i + 1
                                    Else
                                        If rs3!iva_10 > 0 And (a = 2 Or a = 1 Or a = 0) Then
                                            'If val3 = 0 Then val3 = rs3!iva_10
                                            Print #1, "1" & fec & "01" & IIf(rs3!controlador = 1, "C", Chr(32)) & Format(rs3!suc, "0000") & Format(rs3!NroDoc, "00000000000000000000") & fec & Format(IIf(Trim(rs3!aduana) = "", 0, Trim(rs3!aduana)), "000") & IIf(IsNull(rs3!destinacion) Or Trim(rs3!destinacion) = "", "    ", Trim(rs3!destinacion)) & Format(Trim(rs3!despacho), "000000") & IIf(Trim(rs3!verifidespacho) = "", Chr(32), Trim(rs3!verifidespacho)) & documento & cui & nom & _
                                                Format(val, "000000000000000") & Format(val2, "000000000000000") & Format(val3, "000000000000000") & "1050" & Format(val6, "000000000000000") & Format(val5, "000000000000000") & Format(val7 * 100, "000000000000000") & Format(val8 * 100, "000000000000000") & Format(val9, "000000000000000") & "000000000000000" & Format(val4, "000000000000000") & _
                                                Format(rs3!tipoiva, "00") & "PES" & "0001000000" & Format(x, "0") & Chr(32) & Format(rs3!cai, "00000000000000") & fecCAI & relleno
                                            i = i + 1
                                        End If
                                    End If
                                End If
                            Next p
                        End If
                        'Close #1
                        If tI = 1 Or tI = 2 Then
                            netgra = netgra + CDbl(rs3!Neto)
                            netNOgra = netNOgra + CDbl(rs3!EXENTO) '+ CDbl(rs3!imp_int) + CDbl(rs3!IVA_21) + CDbl(rs3!IVA_27) + CDbl(rs3!iva_10)
                        ElseIf tI = 4 Or tI = 6 Or tI = 10 Then    'Or tI = 5
                            netNOgra = netNOgra + CDbl(rs3!Neto) '+ CDbl(rs3!imp_int) + CDbl(rs3!IVA_21) + CDbl(rs3!IVA_27) + CDbl(rs3!iva_10)
                        ElseIf tI = 5 Then
                            netNOgra = netNOgra + CDbl(rs3!EXENTO) '+ CDbl(rs3!imp_int) + CDbl(rs3!IVA_21) + CDbl(rs3!IVA_27) + CDbl(rs3!iva_10)
                            netgra = netgra + CDbl(rs3!Neto)
                        End If
                        
                    'ElseIf tI = 4 Or tI = 6 Or tI = 10 Then 'Trim(rs3!TIPODOC) = "FAB" Then
                        'If CDbl(txtTotal.caption) >= CDbl("10000") Then  'lo quieren hacer uno por uno
                            'Open "C:\gestion\SIRED\" & arch & ".txt" For Output As #1
                            'Write #1, cmbCliente.Text & cmbCliente.Text & Chr(32) & cmbCliente.Text            ' Text1.Text el chr(9)=tab  vbCrLf=enter
                    '        Print #1, "1" & fec & "06" & IIf(rs3!controlador = 1, "C", Chr(32)) & Format(rs3!suc, "0000") & Format(rs3!NroDoc, "00000000000000000000") & fec & Format(Trim(rs3!aduana), "000") & IIf(IsNull(rs3!destinacion) Or Trim(rs3!destinacion) = "", "    ", Trim(rs3!destinacion)) & Format(Trim(rs3!despacho), "000000") & IIf(Trim(rs3!verifidespacho) = "", Chr(32), Trim(rs3!verifidespacho)) & documento & Trim(Mid(rs3!cuitprov, 1, 2)) & Trim(Mid(rs3!cuitprov, 4, 8)) & Trim(Mid(rs3!cuitprov, 13, 1)) & nom & _
                    '            Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!iibb * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & "E" & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & "00000000" & _
                    '            Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(13)
                            'txtneto
                    '        netNOgra = netNOgra + CDbl(rs3!Total)
                            'Close #1
                        'End If
                    '    i = i + 1
                    'ElseIf tI = 5 Then 'consumidor final
                    '    Print #1, "1" & fec & "06" & IIf(rs3!controlador = 1, "C", Chr(32)) & Format(rs3!suc, "0000") & Format(rs3!NroDoc, "00000000000000000000") & fec & Format(Trim(rs3!aduana), "000") & IIf(IsNull(rs3!destinacion) Or Trim(rs3!destinacion) = "", "    ", Trim(rs3!destinacion)) & Format(Trim(rs3!despacho), "000000") & IIf(Trim(rs3!verifidespacho) = "", Chr(32), Trim(rs3!verifidespacho)) & documento & Trim(Mid(rs3!cuitprov, 1, 2)) & Trim(Mid(rs3!cuitprov, 4, 8)) & Trim(Mid(rs3!cuitprov, 13, 1)) & nom & _
                    '        Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!iibb * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & "E" & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & "00000000" & _
                    '        Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(13)
                    'End If
                    rs3.MoveNext
                    Set rs2 = Nothing
                Wend
                'registro de tipo 2
                relleno = ""
                j = 1
                While j < 115
                    relleno = relleno & Chr(32)
                    j = j + 1
                Wend
                
                relleno2 = ""
                j = 1
                While j < 11
                    relleno2 = relleno2 & Chr(32)
                    j = j + 1
                Wend
                
                Print #1, "2" & Year(dtPeriodo.Value) & Format(Month(dtPeriodo.Value), "00") & relleno2 & Format(i, "000000000000") & relleno2 & relleno2 & relleno2 & Chr(32) & Trim(Mid(rs!CuitEmpresa, 1, 2)) & Trim(Mid(rs!CuitEmpresa, 4, 8)) & Trim(Mid(rs!CuitEmpresa, 13, 1)) & relleno2 & relleno2 & relleno2 & _
                 Format(tot * 100, "000000000000000") & Format(netNOgra * 100, "000000000000000") & Format(netgra * 100, "000000000000000") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Format(totimpLIQ, "000000000000000") & Format(totimpEXE, "000000000000000") & Format(totPerAgr, "000000000000000") & Format(totPerOtro, "000000000000000") & Format(IIBB, "000000000000000") & "000000000000000" & Format(totImpInt, "000000000000000") & relleno

                
                If Not IsNull(rs3) Or Not IsEmpty(rs3) Or Not (rs3.EOF = True And rs3.BOF = True) Then
                    Close #1
                End If
                Set rs3 = Nothing
                Set rs = Nothing
                MsgBox "Archivo generado correctamente."
                
            
    '******************************************************************************************

End Sub

Private Sub Command3_Click()   'CABECERA DE FACTURAS
    Dim ARCH As String
    Dim documento As Integer
    Dim nom As String
    Dim tI As Double
    Dim impLIQ As String
    Dim impRNI As String
    Dim impEXE As String
    Dim totimpLIQ As Double
    Dim totimpRNI As Double
    Dim totimpEXE As Double
    Dim IIBB As Double
    Dim fec As String
    Dim pri As Date
    Dim seg As Date
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim rs3 As New ADODB.Recordset
    Dim i As Long
    Dim j As Long
    Dim tot As Double
    Dim netgra As Double
    Dim netNOgra As Double
    Dim rni As Double
    Dim relleno As String
    
    
    pri = "01/" & Month(dtPeriodo.Value) & "/" & Year(dtPeriodo.Value)
    seg = ultimoDiaDelMes(dtPeriodo.Value)
    rs3.Open "select * from facturaventa where fecha>=" & ssFecha(pri) & " and fecha<=" & ssFecha(seg) & " and (tipodoc='FAA' or tipodoc='FAB' or tipodoc='NCA' or tipodoc='NCB') order by fecha,tipodoc,nrofactura", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'********************************** ARCHIVO
                ARCH = "CABECERA_" & Year(dtPeriodo.Value) & Format(Month(dtPeriodo.Value), "00")
                rs.Open "select * from datosempresa where idempresa=15", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                
                If (rs3.EOF = True And rs3.BOF = True) Or IsNull(rs3) Then
                    MsgBox "No se ha encontrado datos para este periodo.", , "ATENCION"
                    Exit Sub
                End If
                
                rs3.MoveFirst
                If Not IsNull(rs3) Or Not IsEmpty(rs3) Or Not (rs3.EOF = True And rs3.BOF = True) Then
                    Open "C:\gestion\SIRED\" & ARCH & ".txt" For Output As #1
                End If
                i = 0
                While Not rs3.EOF
                
                    rs2.Open "select i.*,C.CUIT from clientes c inner join ivas i on i.codigo=c.iva where c.codigo=" & rs3!cliente, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                    If rs3!tipoiva = 1 Or rs3!tipoiva = 4 Or rs3!tipoiva = 6 Then
                    'If ComboCodigo(cmbTipoIva) = 1 Or ComboCodigo(cmbTipoIva) = 4 Or ComboCodigo(cmbTipoIva) = 6 Then
                        documento = 80
                    Else
                        documento = 86
                    End If
                    nom = Trim(rs3!RAZONSOCIAL)
                    While Len(nom) < 30
                        nom = Chr(32) & nom
                    Wend
                    
                    fec = Year(rs3!Fecha) & Format(Month(rs3!Fecha), "00") & Format(Day(rs3!Fecha), "00")
                    'ver si tengo que multiplicar * 100 los valores para sacarle la coma!!!!!!
                    tI = Trim(rs3!tipoiva)
                    If tI = 5 Or tI = 4 Or tI = 6 Or tI = 10 Then 'facturas B
                        impLIQ = "000000000000000"
                        impRNI = "000000000000000"
                        impEXE = "000000000000000" 'Format(TxtIVA * 100, "000000000000000")
                    ElseIf tI = 1 Or tI = 2 Then  'facturas A
                        If tI = 1 Then
                            impLIQ = Format(rs3!Iva * 100, "000000000000000")
                            impRNI = "000000000000000"
                            impEXE = "000000000000000"
                        ElseIf tI = 2 Then
                            impLIQ = "000000000000000"
                            impRNI = Format(rs3!Iva * 100, "000000000000000")
                            impEXE = "000000000000000"
                        ElseIf tI = 8 Then
                            impLIQ = "000000000000000"
                            impRNI = "000000000000000"
                            impEXE = Format(rs3!Iva * 100, "000000000000000")
                        End If
                    End If
                    
                    If Trim(rs3!TIPODOC) = "FAA" Or Trim(rs3!TIPODOC) = "FAB" Or Trim(rs3!TIPODOC) = "NDA" Or Trim(rs3!TIPODOC) = "NDB" Then
                        tot = tot + CDbl(rs3!Total)
                        totimpLIQ = totimpLIQ + CDbl(impLIQ)
                        totimpRNI = totimpRNI + CDbl(impRNI)
                        totimpEXE = totimpEXE + CDbl(impEXE)
                        IIBB = IIBB + CDbl(rs3!IIBB)
                    ElseIf Trim(rs3!TIPODOC) = "NCA" Or Trim(rs3!TIPODOC) = "NCB" Then
                        tot = tot - CDbl(rs3!Total)
                        totimpLIQ = totimpLIQ - CDbl(impLIQ)
                        totimpRNI = totimpRNI - CDbl(impRNI)
                        totimpEXE = totimpEXE - CDbl(impEXE)
                        IIBB = IIBB - CDbl(rs3!IIBB)
                    End If
                                        
                    
                    If Trim(rs3!TIPODOC) = "FAA" Then 'en las tablas de la afip el 01 es factura A
                        'Open "C:\gestion\SIRED\" & arch & ".txt" For Output As #1
                        'Write #1, cmbCliente.Text & cmbCliente.Text & Chr(32) & cmbCliente.Text            ' Text1.Text el chr(9)=tab  vbCrLf=enter   chr(32)=un espacio
                        Print #1, "1" & fec & "01" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000") & Format(rs3!NroFactura, "00000000") & "001" & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & Chr(32) & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) '& Chr(13)
                        'Close #1
                        netgra = netgra + CDbl(rs3!Neto)
                        i = i + 1
                    ElseIf Trim(rs3!TIPODOC) = "FAB" Then
                        'If CDbl(txtTotal.caption) >= CDbl("10000") Then  'lo quieren hacer uno por uno
                            'Open "C:\gestion\SIRED\" & arch & ".txt" For Output As #1
                            'Write #1, cmbCliente.Text & cmbCliente.Text & Chr(32) & cmbCliente.Text            ' Text1.Text el chr(9)=tab  vbCrLf=enter
                            Print #1, "1" & fec & "06" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000") & Format(rs3!NroFactura, "00000000") & "001" & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & "E" & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) '& Chr(13)
                            'txtneto
                            netNOgra = netNOgra + CDbl(rs3!Total)
                            'Close #1
                        'End If
                        i = i + 1
                    ElseIf Trim(rs3!TIPODOC) = "NCA" Then
                        
                        Print #1, "1" & fec & "03" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000") & Format(rs3!NroFactura, "00000000") & "001" & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & Chr(32) & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) '& Chr(13)
                        netgra = netgra - CDbl(rs3!Neto)
                        i = i + 1
                    ElseIf Trim(rs3!TIPODOC) = "NDA" Then
                        
                        Print #1, "1" & fec & "02" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000") & Format(rs3!NroFactura, "00000000") & "001" & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & Chr(32) & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) '& Chr(13)
                        netgra = netgra + CDbl(rs3!Neto)
                        i = i + 1
                    ElseIf Trim(rs3!TIPODOC) = "NCB" Then
                        
                        Print #1, "1" & fec & "08" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000") & Format(rs3!NroFactura, "00000000") & "001" & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & "E" & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) '& Chr(13)
                        netNOgra = netNOgra - CDbl(rs3!Total)
                        i = i + 1
                    ElseIf Trim(rs3!TIPODOC) = "NDB" Then
                        
                        Print #1, "1" & fec & "07" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000") & Format(rs3!NroFactura, "00000000") & "001" & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & "E" & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) '& Chr(13)
                        netNOgra = netNOgra + CDbl(rs3!Total)
                        i = i + 1
                        
                    End If
                    rs3.MoveNext
                    Set rs2 = Nothing
                Wend
                'registro de tipo 2
                j = 1
                While j < 63
                    relleno = relleno & Chr(32)
                    j = j + 1
                Wend
                
                Print #1, "2" & Year(dtPeriodo.Value) & Format(Month(dtPeriodo.Value), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Format(i, "00000000") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Trim(Mid(rs!CuitEmpresa, 1, 2)) & Trim(Mid(rs!CuitEmpresa, 4, 8)) & Trim(Mid(rs!CuitEmpresa, 13, 1)) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                 Format(tot * 100, "000000000000000") & Format(netNOgra * 100, "000000000000000") & Format(netgra * 100, "000000000000000") & Format(totimpLIQ, "000000000000000") & Format(totimpRNI, "000000000000000") & Format(totimpEXE, "000000000000000") & "000000000000000" & Format(IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & relleno

                
                If Not IsNull(rs3) Or Not IsEmpty(rs3) Or Not (rs3.EOF = True And rs3.BOF = True) Then
                    Close #1
                End If
                Set rs3 = Nothing
                Set rs = Nothing
                MsgBox "Archivo generado correctamente."
                
            
    '******************************************************************************************

End Sub

Private Sub Command4_Click() 'DETALLE DE FACTURAS
    Dim ARCH As String
    Dim documento As Integer
    Dim nom As String
    Dim prod As String
    Dim tI As Double
    Dim impLIQ As String
    Dim impRNI As String
    Dim impEXE As String
    Dim totimpLIQ As Double
    Dim totimpRNI As Double
    Dim totimpEXE As Double
    Dim IIBB As Double
    Dim fec As String
    Dim pri As Date
    Dim seg As Date
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim rs3 As New ADODB.Recordset
    Dim rs4 As New ADODB.Recordset
    Dim i As Long
    Dim j As Long
    Dim z As Long
    Dim tot As Double
    Dim netgra As Double
    Dim netNOgra As Double
    Dim rni As Double
    Dim relleno As String
    Dim cant As Double
    Dim PU As Double
    Dim item As Long
    Dim PT As Double
    
    
    pri = "01/" & Month(dtPeriodo.Value) & "/" & Year(dtPeriodo.Value)
    seg = ultimoDiaDelMes(dtPeriodo.Value)
    rs3.Open "select * from facturaventa where fecha>=" & ssFecha(pri) & " and fecha<=" & ssFecha(seg) & " order by fecha,tipodoc,nrofactura", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'********************************** ARCHIVO
                ARCH = "DETALLE_" & Year(dtPeriodo.Value) & Format(Month(dtPeriodo.Value), "00")
                rs.Open "select * from datosempresa where idempresa=15", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                
                If (rs3.EOF = True And rs3.BOF = True) Or IsNull(rs3) Then
                    MsgBox "No se ha encontrado datos para este periodo.", , "ATENCION"
                    Exit Sub
                End If
                
                rs3.MoveFirst
                If Not IsNull(rs3) Or Not IsEmpty(rs3) Or Not (rs3.EOF = True And rs3.BOF = True) Then
                    Open "C:\gestion\SIRED\" & ARCH & ".txt" For Output As #1
                End If
                i = 0
                
                'j = 1
                'While j < 201
                '    relleno = relleno & Chr(32)
                '    j = j + 1
                'Wend
                
                While Not rs3.EOF
                
                    rs2.Open "select i.*,C.CUIT from clientes c inner join ivas i on i.codigo=c.iva where c.codigo=" & rs3!cliente, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                    If rs3!tipoiva = 1 Or rs3!tipoiva = 4 Or rs3!tipoiva = 6 Then
                    'If ComboCodigo(cmbTipoIva) = 1 Or ComboCodigo(cmbTipoIva) = 4 Or ComboCodigo(cmbTipoIva) = 6 Then
                        documento = 80
                    Else
                        documento = 86
                    End If
                    nom = Trim(rs3!RAZONSOCIAL)
                    While Len(nom) < 30
                        nom = Chr(32) & nom
                    Wend
                    
                    fec = Year(rs3!Fecha) & Format(Month(rs3!Fecha), "00") & Format(Day(rs3!Fecha), "00")
                    'ver si tengo que multiplicar * 100 los valores para sacarle la coma!!!!!!
                    'tI = Trim(rs3!tipoiva)
                    'If tI = 5 Or tI = 4 Or tI = 6 Or tI = 10 Then 'facturas B
                    '    impLIQ = "000000000000000"
                    '    impRNI = "000000000000000"
                    '    impEXE = "000000000000000" 'Format(TxtIVA * 100, "000000000000000")
                    'ElseIf tI = 1 Or tI = 2 Then  'facturas A
                    '    If tI = 1 Then
                    '        impLIQ = Format(rs3!iva * 100, "000000000000000")
                    '        impRNI = "000000000000000"
                    '        impEXE = "000000000000000"
                    '    ElseIf tI = 2 Then
                    '        impLIQ = "000000000000000"
                    '        impRNI = Format(rs3!iva * 100, "000000000000000")
                    '        impEXE = "000000000000000"
                    '    ElseIf tI = 8 Then
                    '        impLIQ = "000000000000000"
                    '        impRNI = "000000000000000"
                    '        impEXE = Format(rs3!iva * 100, "000000000000000")
                    '    End If
                    'End If
                    
                    'tot = tot + CDbl(rs3!Total)
                    'totimpLIQ = totimpLIQ + CDbl(impLIQ)
                    'totimpRNI = totimpRNI + CDbl(impRNI)
                    'totimpEXE = totimpEXE + CDbl(impEXE)
                    'iibb = iibb + CDbl(rs3!iibb)
                    
                    item = obtenerDeSQL("select count(distinct(producto)) from facturaventadetalle where tipodoc='" & rs3!TIPODOC & "' and nrofactura=" & rs3!NroFactura)
                    
                    rs4.Open "select distinct(producto) as prod,* from facturaventadetalle where tipodoc='" & rs3!TIPODOC & "' and nrofactura=" & rs3!NroFactura, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                    
                    If Trim(rs3!TIPODOC) = "FAA" Then 'en las tablas de la afip el 01 es factura A
                        'Open "C:\gestion\SIRED\" & arch & ".txt" For Output As #1
                        'Write #1, cmbCliente.Text & cmbCliente.Text & Chr(32) & cmbCliente.Text            ' Text1.Text el chr(9)=tab  vbCrLf=enter   chr(32)=un espacio
                        z = 0
                        While z < item
                            cant = obtenerDeSQL("select sum(cantidad) from facturaventadetalle where tipodoc='" & Trim(rs3!TIPODOC) & "' and nrofactura=" & rs3!NroFactura & " and producto='" & Trim(rs4!prod) & "'")
                            PU = obtenerDeSQL("select preciounitario from facturaventadetalle where tipodoc='" & Trim(rs3!TIPODOC) & "' and nrofactura=" & rs3!NroFactura & " and producto='" & Trim(rs4!prod) & "'")
                            PT = obtenerDeSQL("select sum(preciototal) from facturaventadetalle where tipodoc='" & Trim(rs3!TIPODOC) & "' and nrofactura=" & rs3!NroFactura & " and producto='" & Trim(rs4!prod) & "'")
                            prod = Trim(rs4!DESCRIPCION)
                            While Len(prod) < 75
                                prod = prod & Chr(32)
                            Wend
                            
                            Print #1, "01" & Chr(32) & fec & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000") & Format(rs3!NroFactura, "00000000") & Format(cant * 100000, "000000000000") & "98" & Format(PU * 1000, "0000000000000000") & "000000000000000" & "0000000000000000" & Format(PT * 1000, "0000000000000000") & "2100" & "G" & Chr(32) & prod '& Chr(13)
                            z = z + 1
                            rs4.MoveNext
                        Wend
                        'Close #1
                        'netgra = netgra + CDbl(rs3!neto)
                        'i = i + 1
                    ElseIf Trim(rs3!TIPODOC) = "FAB" Then
                        'If CDbl(txtTotal.caption) >= CDbl("10000") Then  'lo quieren hacer uno por uno
                            'Open "C:\gestion\SIRED\" & arch & ".txt" For Output As #1
                            'Write #1, cmbCliente.Text & cmbCliente.Text & Chr(32) & cmbCliente.Text            ' Text1.Text el chr(9)=tab  vbCrLf=enter
                        z = 0
                        While z < item
                            cant = obtenerDeSQL("select sum(cantidad) from facturaventadetalle where tipodoc='" & Trim(rs3!TIPODOC) & "' and nrofactura=" & rs3!NroFactura & " and producto='" & Trim(rs4!prod) & "'")
                            PU = obtenerDeSQL("select preciounitario from facturaventadetalle where tipodoc='" & Trim(rs3!TIPODOC) & "' and nrofactura=" & rs3!NroFactura & " and producto='" & Trim(rs4!prod) & "'")
                            PT = obtenerDeSQL("select preciototal from facturaventadetalle where tipodoc='" & Trim(rs3!TIPODOC) & "' and nrofactura=" & rs3!NroFactura & " and producto='" & Trim(rs4!prod) & "'")
                            prod = Trim(rs4!DESCRIPCION)
                            While Len(prod) < 75
                                prod = prod & Chr(32)
                            Wend
                            
                            Print #1, "06" & Chr(32) & fec & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000") & Format(rs3!NroFactura, "00000000") & Format(cant * 100000, "000000000000") & "98" & Format(PU * 1000, "0000000000000000") & "000000000000000" & "0000000000000000" & Format(PT * 1000, "0000000000000000") & "0000" & "E" & Chr(32) & prod '& Chr(13)
                            z = z + 1
                            rs4.MoveNext
                        Wend
                            'txtneto
                            'netNOgra = netNOgra + CDbl(rs3!Total)
                            'Close #1
                        'End If
                        'i = i + 1
                    End If
                    rs3.MoveNext
                    Set rs2 = Nothing
                    Set rs4 = Nothing
                Wend
                'registro de tipo 2
                
                
                'Print #1, "2" & Year(dtperiodo.Value) & Format(Month(dtperiodo.Value), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Format(i, "000000000000") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Trim(Mid(rs!cuitempresa, 1, 2)) & Trim(Mid(rs!cuitempresa, 4, 8)) & Trim(Mid(rs!cuitempresa, 13, 1)) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                ' Format(tot * 100, "000000000000000") & Format(netNOgra * 100, "000000000000000") & Format(netgra * 100, "000000000000000") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Format(totimpLIQ, "000000000000000") & Format(totimpRNI, "000000000000000") & Format(totimpEXE, "000000000000000") & "000000000000000" & Format(iibb * 100, "000000000000000") & "000000000000000" & "000000000000000" & relleno & "j"

                
                If Not IsNull(rs3) Or Not IsEmpty(rs3) Or Not (rs3.EOF = True And rs3.BOF = True) Then
                    Close #1
                End If
                Set rs3 = Nothing
                Set rs = Nothing
                MsgBox "Archivo generado correctamente."
                
            
    '******************************************************************************************

End Sub

Private Sub Command5_Click() 'OTRAS PERCEPCIONES
    Dim ARCH As String
    Dim documento As Integer
    Dim nom As String
    Dim tI As Double
    Dim impLIQ As String
    Dim impRNI As String
    Dim impEXE As String
    Dim totimpLIQ As Double
    Dim totimpRNI As Double
    Dim totimpEXE As Double
    Dim IIBB As Double
    Dim fec As String
    Dim pri As Date
    Dim seg As Date
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim rs3 As New ADODB.Recordset
    Dim i As Long
    Dim j As Long
    Dim tot As Double
    Dim netgra As Double
    Dim netNOgra As Double
    Dim rni As Double
    Dim relleno As String
    
    
    pri = "01/" & Month(dtPeriodo.Value) & "/" & Year(dtPeriodo.Value)
    seg = ultimoDiaDelMes(dtPeriodo.Value)
    rs3.Open "select * from facturaventa where fecha>=" & ssFecha(pri) & " and fecha<=" & ssFecha(seg) & " and iibb<>0 order by fecha,tipodoc,nrofactura", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'********************************** ARCHIVO
                ARCH = "OTRAS_PERCEP_" & Year(dtPeriodo.Value) & Format(Month(dtPeriodo.Value), "00")
                rs.Open "select * from datosempresa where idempresa=15", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                
                If (rs3.EOF = True And rs3.BOF = True) Or IsNull(rs3) Then
                    MsgBox "No se ha encontrado datos para este periodo.", , "ATENCION"
                    Exit Sub
                End If
                
                rs3.MoveFirst
                If Not IsNull(rs3) Or Not IsEmpty(rs3) Or Not (rs3.EOF = True And rs3.BOF = True) Then
                    Open "C:\gestion\SIRED\" & ARCH & ".txt" For Output As #1
                End If
                i = 0
                
                j = 1
                While j < 41
                    relleno = relleno & Chr(32)
                    j = j + 1
                Wend
                
                While Not rs3.EOF
                
                    rs2.Open "select i.*,C.CUIT from clientes c inner join ivas i on i.codigo=c.iva where c.codigo=" & rs3!cliente, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                    If rs3!tipoiva = 1 Or rs3!tipoiva = 4 Or rs3!tipoiva = 6 Then
                    'If ComboCodigo(cmbTipoIva) = 1 Or ComboCodigo(cmbTipoIva) = 4 Or ComboCodigo(cmbTipoIva) = 6 Then
                        documento = 80
                    Else
                        documento = 86
                    End If
                    'nom = Trim(rs3!RAZONSOCIAL)
                    'While Len(nom) < 30
                    '    nom = Chr(32) & nom
                    'Wend
                    
                    fec = Year(rs3!Fecha) & Format(Month(rs3!Fecha), "00") & Format(Day(rs3!Fecha), "00")
                    'ver si tengo que multiplicar * 100 los valores para sacarle la coma!!!!!!
                    'tI = Trim(rs3!tipoiva)
                    'If tI = 5 Or tI = 4 Or tI = 6 Or tI = 10 Then 'facturas B
                    '    impLIQ = "000000000000000"
                    '    impRNI = "000000000000000"
                    '    impEXE = "000000000000000" 'Format(TxtIVA * 100, "000000000000000")
                    'ElseIf tI = 1 Or tI = 2 Then  'facturas A
                    '    If tI = 1 Then
                    '        impLIQ = Format(rs3!iva * 100, "000000000000000")
                    '        impRNI = "000000000000000"
                    '        impEXE = "000000000000000"
                    '    ElseIf tI = 2 Then
                    '        impLIQ = "000000000000000"
                    '        impRNI = Format(rs3!iva * 100, "000000000000000")
                    '        impEXE = "000000000000000"
                    '    ElseIf tI = 8 Then
                    '        impLIQ = "000000000000000"
                    '        impRNI = "000000000000000"
                    '        impEXE = Format(rs3!iva * 100, "000000000000000")
                    '    End If
                    'End If
                    
                    'tot = tot + CDbl(rs3!Total)
                    'totimpLIQ = totimpLIQ + CDbl(impLIQ)
                    'totimpRNI = totimpRNI + CDbl(impRNI)
                    'totimpEXE = totimpEXE + CDbl(impEXE)
                    'iibb = iibb + CDbl(rs3!iibb)
                    
                    
                    If Trim(rs3!TIPODOC) = "FAA" Then 'en las tablas de la afip el 01 es factura A
                        'Open "C:\gestion\SIRED\" & arch & ".txt" For Output As #1
                        'Write #1, cmbCliente.Text & cmbCliente.Text & Chr(32) & cmbCliente.Text            ' Text1.Text el chr(9)=tab  vbCrLf=enter   chr(32)=un espacio
                        Print #1, fec & "01" & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000") & provi(rs3!Provincia) & Format(rs3!IIBB * 100, "000000000000000") & relleno & "000000000000000" '& Chr(13)
                        'Close #1
                        netgra = netgra + CDbl(rs3!Neto)
                        i = i + 1
                    ElseIf Trim(rs3!TIPODOC) = "FAB" Then
                        'If CDbl(txtTotal.caption) >= CDbl("10000") Then  'lo quieren hacer uno por uno
                            'Open "C:\gestion\SIRED\" & arch & ".txt" For Output As #1
                            'Write #1, cmbCliente.Text & cmbCliente.Text & Chr(32) & cmbCliente.Text            ' Text1.Text el chr(9)=tab  vbCrLf=enter
                            Print #1, fec & "06" & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000") & provi(rs3!Provincia) & Format(rs3!IIBB * 100, "000000000000000") & relleno & "000000000000000" '& Chr(13)
                            'txtneto
                            netNOgra = netNOgra + CDbl(rs3!Total)
                            'Close #1
                        'End If
                        i = i + 1
                    End If
                    rs3.MoveNext
                    Set rs2 = Nothing
                Wend
                'registro de tipo 2
                
                
                'Print #1, "2" & Year(dtperiodo.Value) & Format(Month(dtperiodo.Value), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Format(i, "00000000") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Trim(Mid(rs!cuitempresa, 1, 2)) & Trim(Mid(rs!cuitempresa, 4, 8)) & Trim(Mid(rs!cuitempresa, 13, 1)) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                ' Format(tot * 100, "000000000000000") & Format(netNOgra * 100, "000000000000000") & Format(netgra * 100, "000000000000000") & Format(totimpLIQ, "000000000000000") & Format(totimpRNI, "000000000000000") & Format(totimpEXE, "000000000000000") & "000000000000000" & Format(iibb * 100, "000000000000000") & "000000000000000" & "000000000000000" & relleno

                
                If Not IsNull(rs3) Or Not IsEmpty(rs3) Or Not (rs3.EOF = True And rs3.BOF = True) Then
                    Close #1
                End If
                Set rs3 = Nothing
                Set rs = Nothing
                MsgBox "Archivo generado correctamente."
                
            
    '******************************************************************************************

End Sub



Public Function ultimoDiaDelMes(Fecha As Date) As Date
ultimoDiaDelMes = DateAdd("m", 1, Fecha)
ultimoDiaDelMes = DateSerial(Year(ultimoDiaDelMes), Month(ultimoDiaDelMes), 1)
ultimoDiaDelMes = DateAdd("d", -1, ultimoDiaDelMes)
End Function

Private Function provi(p As String) As String
    Select Case p
        Case "*":
            provi = "00"
        Case "B":
            provi = "01"
        Case "S":
            provi = "12"
        Case "Z":
            provi = "23"
        Case "K":
            provi = "02"
        Case "H":
            provi = "16"
        Case "U":
            provi = "17"
        Case "X":
            provi = "03"
        Case "W":
            provi = "04"
        Case "E":
            provi = "05"
        Case "P":
            provi = "18"
        Case "Y":
            provi = "06"
        Case "L":
            provi = "21"
        Case "F":
            provi = "08"
        Case "M":
            provi = "07"
        Case "N":
            provi = "19"
        Case "Q":
            provi = "20"
        Case "R":
            provi = "22"
        Case "A":
            provi = "09"
        Case "J":
            provi = "10"
        Case "D":
            provi = "11"
        Case "G":
            provi = "13"
        Case "V":
            provi = "24"
        Case "T":
            provi = "14"
    End Select
End Function

Private Sub Command7_Click()
    Dim ARCH As String
    Dim documento As Integer
    Dim nom As String
    Dim tI As Double
    Dim impLIQ As String
    Dim impLIQ2 As Double
    Dim impRNI As String
    Dim impEXE As String
    Dim totimpLIQ As Double
    Dim totimpRNI As Double
    Dim totimpEXE As Double
    Dim IIBB As Double
    Dim fec As String
    Dim pri As Date
    Dim seg As Date
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim rs3 As New ADODB.Recordset
    Dim rs4 As New ADODB.Recordset
    Dim i As Long
    Dim j As Long
    Dim tot As Double
    Dim netgra As Double
    Dim netNOgra As Double
    Dim rni As Double
    Dim relleno As String
    Dim relleno2 As String
    Dim str As String
    Dim cantIVA As Long
    Dim Iva As String ' Double
    Dim Neto As Double
    Dim Tipo As String
    
    
    pri = "01/" & Month(dtPeriodo.Value) & "/" & Year(dtPeriodo.Value)
    seg = ultimoDiaDelMes(dtPeriodo.Value)
    str = "select * from facturaventa where fecha>=" & ssFecha(pri) & " and fecha<=" & ssFecha(seg) & " and (tipodoc='FAA' or tipodoc='FAB' or tipodoc='NCA' or tipodoc='NCB') order by fecha,tipodoc,nrofactura"
    rs3.Open str, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    
'********************************** ARCHIVO

    ARCH = "VENTAS_" & Year(dtPeriodo.Value) & Format(Month(dtPeriodo.Value), "00")
    rs.Open "select * from datosempresa where idempresa=4", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    
    If (rs3.EOF = True And rs3.BOF = True) Or IsNull(rs3!codigo) Then
        MsgBox "No se ha encontrado datos para este periodo.", , "ATENCION"
        Exit Sub
    End If
    
    rs3.MoveFirst
    If Not IsNull(rs3) Or Not IsEmpty(rs3) Or Not (rs3.EOF = True And rs3.BOF = True) Then
        'Call Verificar_Existe("C:\CITI\") '& ARCH & ".txt"
        
        Open "C:\CITI\" & ARCH & ".txt" For Output As #1
    End If
    i = 0
    While Not rs3.EOF
        If rs3!activo = False Then 'si esta anulado va todo en cero
            fec = Year(rs3!Fecha) & Format(Month(rs3!Fecha), "00") & Format(Day(rs3!Fecha), "00")
            If Trim(rs3!TIPODOC) = "FAA" Then
                Tipo = "01"
            ElseIf Trim(rs3!TIPODOC) = "FAB" Then
                Tipo = "06"
            ElseIf Trim(rs3!TIPODOC) = "NCA" Then
                Tipo = "03"
            ElseIf Trim(rs3!TIPODOC) = "NCB" Then
                Tipo = "08"
            ElseIf Trim(rs3!TIPODOC) = "NDA" Then
                Tipo = "02"
            ElseIf Trim(rs3!TIPODOC) = "NDB" Then
                Tipo = "07"
            End If
            If rs3!tipoiva = 2 Or rs3!tipoiva = 4 Or rs3!tipoiva = 7 Then 'revisar
                documento = 80 'cuit
            ElseIf rs3!tipoiva = 1 Then
                documento = IIf(Mid(sSinNull(rs3!CUIT), 1, 4) = "00-0", 96, 86) '96=dni,86=cuil
            Else
                documento = 86
            End If
            nom = Mid(Trim(rs3!RAZONSOCIAL), 1, 30)
            While Len(nom) < 30
                nom = Chr(32) & nom
            Wend
            impLIQ = "000000000000000"
            impRNI = "000000000000000"
            impEXE = "000000000000000"
            cantIVA = 1
            relleno2 = ""
            j = 1
            While j <= 75
                relleno2 = relleno2 & Chr(32)
                j = j + 1
            Wend
                        
            Print #1, "1" & fec & Tipo & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim("00000000000") & nom & Format(0, "000000000000000") & "000000000000000" & Format(0, "000000000000000") & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & "000000000000000" & "000000000000000" & "000000000000000" & "00" & "PES" & "0000000000" & cantIVA & Chr(32) & "00000000000000" & "00000000" & "00000000" & _
                            relleno2 & "00000000" & "000000000000000"
        Else
            cantIVA = obtenerDeSQL("select count(distinct(_iva)) from facturaventadetalle where codigofactura=" & rs3!codigo)
            If cantIVA > 1 Then
                
                rs4.Open "select *,_iva as iva from facturaventadetalle where codigofactura=" & rs3!codigo & " order by _iva", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        
                rs2.Open "select i.*,C.CUIT from clientes c inner join ivas i on i.codigo=c.iva where c.codigo=" & rs3!cliente, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                If rs3!tipoiva = 2 Or rs3!tipoiva = 4 Or rs3!tipoiva = 7 Then 'revisar
                    documento = 80 'cuit
                ElseIf rs3!tipoiva = 1 Then
                    documento = IIf(Mid(sSinNull(rs3!CUIT), 1, 4) = "00-0", 96, 86) '96=dni,86=cuil
                Else
                    documento = 86
                End If
                nom = Mid(Trim(rs3!RAZONSOCIAL), 1, 30)
                While Len(nom) < 30
                    nom = Chr(32) & nom
                Wend
                
                fec = Year(rs3!Fecha) & Format(Month(rs3!Fecha), "00") & Format(Day(rs3!Fecha), "00")
                tI = Trim(rs3!tipoiva)
                                        
                If tI = 1 Or tI = 4 Or tI = 7 Or tI = 8 Then 'facturas B
                    impLIQ = "000000000000000"
                    impRNI = "000000000000000"
                    impEXE = "000000000000000"
                ElseIf tI = 2 Or tI = 3 Then  'facturas A
                    If tI = 2 Then
                        impLIQ = Format(rs3!Iva * 100, "000000000000000")
                        impRNI = "000000000000000"
                        impEXE = "000000000000000"
                    ElseIf tI = 3 Then
                        impLIQ = "000000000000000"
                        impRNI = Format(rs3!Iva * 100, "000000000000000")
                        impEXE = "000000000000000"
                    ElseIf tI = 8 Then ' a este no entra nunca, pero asi estaba en tavi...
                        impLIQ = "000000000000000"
                        impRNI = "000000000000000"
                        impEXE = Format(rs3!Iva * 100, "000000000000000")
                    End If
                End If
                
                If Trim(rs3!TIPODOC) = "FAA" Or Trim(rs3!TIPODOC) = "FAB" Or Trim(rs3!TIPODOC) = "NDA" Or Trim(rs3!TIPODOC) = "NDB" Then
                    tot = tot + CDbl(rs3!Total)
                    totimpLIQ = totimpLIQ + CDbl(impLIQ)
                    totimpRNI = totimpRNI + CDbl(impRNI)
                    totimpEXE = totimpEXE + CDbl(impEXE)
                    IIBB = IIBB + CDbl(rs3!IIBB)
                ElseIf Trim(rs3!TIPODOC) = "NCA" Or Trim(rs3!TIPODOC) = "NCB" Then
                    tot = tot - CDbl(rs3!Total)
                    totimpLIQ = totimpLIQ - CDbl(impLIQ)
                    totimpRNI = totimpRNI - CDbl(impRNI)
                    totimpEXE = totimpEXE - CDbl(impEXE)
                    IIBB = IIBB - CDbl(rs3!IIBB)
                End If
                
                '***************************
                If Trim(rs3!TIPODOC) = "FAA" Then
                    Tipo = "01"
                ElseIf Trim(rs3!TIPODOC) = "FAB" Then
                    Tipo = "06"
                ElseIf Trim(rs3!TIPODOC) = "NCA" Then
                    Tipo = "03"
                ElseIf Trim(rs3!TIPODOC) = "NCB" Then
                    Tipo = "08"
                ElseIf Trim(rs3!TIPODOC) = "NDA" Then
                    Tipo = "02"
                ElseIf Trim(rs3!TIPODOC) = "NDB" Then
                    Tipo = "07"
                End If
                Do While Not rs4.EOF
                    If tI = 1 Or tI = 4 Or tI = 7 Or tI = 8 Then 'facturas B
                        'impLIQ = Format(s2n(rs4!PrecioTotal * 1 + (rs4!Iva / 100)) * 100, "000000000000000")
                        impLIQ = Format(s2n(rs4!PrecioTotal - (rs4!PrecioTotal / (1 + (rs4!Iva / 100)))) * 100, "000000000000000")
                        Iva = Format(rs4!Iva * 100, "0000")
                        Neto = rs4!PrecioTotal - s2n(rs4!PrecioTotal - (rs4!PrecioTotal / (1 + (rs4!Iva / 100))))
                    ElseIf tI = 2 Then
                        impLIQ = Format(s2n(rs4!PrecioTotal * (rs4!Iva / 100)) * 100, "000000000000000")
                        Iva = Format(rs4!Iva * 100, "0000")
                        Neto = rs4!PrecioTotal
                    End If
                    If rs4.AbsolutePosition <> rs4.RecordCount Then
                        relleno2 = ""
                        j = 1
                        While j <= 75
                            relleno2 = relleno2 & Chr(32)
                            j = j + 1
                        Wend
                        'Print #1, "1" & fec & tipo & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(0, "000000000000000") & "000000000000000" & Format(rs4!PrecioTotal * 100, "000000000000000") & (rs4!Iva * 100) & impLIQ & impRNI & impEXE & "000000000000000" & "000000000000000" & "000000000000000" & "000000000000000" & "00" & Chr(32) & Chr(32) & Chr(32) & "0000000000" & cantIVA & Chr(32) & "00000000000000" & "00000000" & "00000000" & _
                            relleno2 & "00000000" & "000000000000000"
                        Print #1, "1" & fec & Tipo & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(0, "000000000000000") & "000000000000000" & Format(Neto * 100, "000000000000000") & Format((rs4!Iva * 100), "0000") & impLIQ & impRNI & impEXE & "000000000000000" & "000000000000000" & "000000000000000" & "000000000000000" & "00" & Chr(32) & Chr(32) & Chr(32) & "0000000000" & cantIVA & Chr(32) & "00000000000000" & "00000000" & "00000000" & _
                            relleno2 & "00000000" & "000000000000000"
                        rs4.MoveNext
                    Else
                        Exit Do
                    End If
                Loop
                Set rs4 = Nothing
                '*********************************************
                
                If Trim(rs3!TIPODOC) = "FAA" Then  'en las tablas de la afip el 01 es factura A
                    relleno2 = ""
                    j = 1
                    While j <= 75
                        relleno2 = relleno2 & Chr(32)
                        j = j + 1
                    Wend
        ''            Print #1, "1" & fec & "01" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & "2100" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & Chr(32) & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                    "Sin comentarios" & relleno2
    '''                Print #1, "1" & fec & "01" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & "2100" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0000000000" & "1" & Chr(32) & Format(nSinNull(rs3!CAE), "00000000000000") & Year(rs3!CAEV) & Format(Month(rs3!CAEV), "00") & Format(Day(rs3!CAEV), "00") & "00000000" & _
                    relleno2 & "00000000" & "000000000000000"
                    Print #1, "1" & fec & "01" & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(Neto * 100, "000000000000000") & Format(Iva, "0000") & impLIQ & impRNI & impEXE & "000000000000000" & "000000000000000" & "000000000000000" & "000000000000000" & "00" & Chr(32) & Chr(32) & Chr(32) & "0000000000" & cantIVA & Chr(32) & "00000000000000" & "00000000" & "00000000" & _
                    relleno2 & "00000000" & "000000000000000"
                    'Close #1
                    netgra = netgra + CDbl(rs3!Neto)
                    i = i + 1
                ElseIf Trim(rs3!TIPODOC) = "FAB" Then
                    'If CDbl(txtTotal.caption) >= CDbl("10000") Then  'lo quieren hacer uno por uno
                        'Open "C:\gestion\SIRED\" & arch & ".txt" For Output As #1
                        'Write #1, cmbCliente.Text & cmbCliente.Text & Chr(32) & cmbCliente.Text            ' Text1.Text el chr(9)=tab  vbCrLf=enter
                        relleno2 = ""
                        j = 1
                        While j <= 75
                            relleno2 = relleno2 & Chr(32)
                            j = j + 1
                        Wend
        ''                Print #1, "1" & fec & "06" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & "E" & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                        "Sin comentarios" & relleno2 '& vbCrLf
    '''                    Print #1, "1" & fec & "06" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0000000000" & "1" & Chr(32) & Format(nSinNull(rs3!CAE), "00000000000000") & Year(rs3!CAEV) & Format(Month(rs3!CAEV), "00") & Format(Day(rs3!CAEV), "00") & "00000000" & _
                        relleno2 & "00000000" & "000000000000000"
                        Print #1, "1" & fec & "06" & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(0, "000000000000000") & Format(Neto * 100, "000000000000000") & Format(Iva, "0000") & impLIQ & impRNI & impEXE & "000000000000000" & "000000000000000" & "000000000000000" & "000000000000000" & "00" & Chr(32) & Chr(32) & Chr(32) & "0000000000" & cantIVA & Chr(32) & "00000000000000" & "00000000" & "00000000" & _
                        relleno2 & "00000000" & "000000000000000"
                        'txtneto
                        netNOgra = netNOgra + CDbl(rs3!Total)
                        'Close #1
                    'End If
                    i = i + 1
                    Set rs4 = Nothing
                ElseIf Trim(rs3!TIPODOC) = "NCA" Then
                    relleno2 = ""
                    j = 1
                    While j <= 75
                        relleno2 = relleno2 & Chr(32)
                        j = j + 1
                    Wend
        ''            Print #1, "1" & fec & "03" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & "2100" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & Chr(32) & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                    "Sin comentarios" & relleno2
    '''                Print #1, "1" & fec & "03" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & "2100" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0000000000" & "1" & Chr(32) & Format(nSinNull(rs3!CAE), "00000000000000") & Year(rs3!CAEV) & Format(Month(rs3!CAEV), "00") & Format(Day(rs3!CAEV), "00") & "00000000" & _
                    relleno2 & "00000000" & "000000000000000"
                    Print #1, "1" & fec & "03" & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(Neto * 100, "000000000000000") & Format(Iva, "0000") & impLIQ & impRNI & impEXE & "000000000000000" & "000000000000000" & "000000000000000" & "000000000000000" & "00" & Chr(32) & Chr(32) & Chr(32) & "0000000000" & cantIVA & Chr(32) & "00000000000000" & "00000000" & "00000000" & _
                    relleno2 & "00000000" & "000000000000000"
                    netgra = netgra - CDbl(rs3!Neto)
                    i = i + 1
                ElseIf Trim(rs3!TIPODOC) = "NDA" Then
                    relleno2 = ""
                    j = 1
                    While j <= 75
                        relleno2 = relleno2 & Chr(32)
                        j = j + 1
                    Wend
        ''            Print #1, "1" & fec & "02" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & "2100" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & Chr(32) & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                    "Sin comentarios" & relleno2
    '''                Print #1, "1" & fec & "02" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & "2100" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0000000000" & "1" & Chr(32) & Format(nSinNull(rs3!CAE), "00000000000000") & Year(rs3!CAEV) & Format(Month(rs3!CAEV), "00") & Format(Day(rs3!CAEV), "00") & "00000000" & _
                    relleno2 & "00000000" & "000000000000000"
                    Print #1, "1" & fec & "02" & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(Neto * 100, "000000000000000") & Format(Iva, "0000") & impLIQ & impRNI & impEXE & "000000000000000" & "000000000000000" & "000000000000000" & "000000000000000" & "00" & Chr(32) & Chr(32) & Chr(32) & "0000000000" & cantIVA & Chr(32) & "00000000000000" & "00000000" & "00000000" & _
                    relleno2 & "00000000" & "000000000000000"
                    netgra = netgra + CDbl(rs3!Neto)
                    i = i + 1
                ElseIf Trim(rs3!TIPODOC) = "NCB" Then
                    relleno2 = ""
                    j = 1
                    While j <= 75
                        relleno2 = relleno2 & Chr(32)
                        j = j + 1
                    Wend
        ''            Print #1, "1" & fec & "08" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & "E" & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                        "Sin comentarios" & relleno2 '& vbCrLf
    '''                Print #1, "1" & fec & "08" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0000000000" & "1" & Chr(32) & Format(nSinNull(rs3!CAE), "00000000000000") & Year(rs3!CAEV) & Format(Month(rs3!CAEV), "00") & Format(Day(rs3!CAEV), "00") & "00000000" & _
                        relleno2 & "00000000" & "000000000000000"
                    Print #1, "1" & fec & "08" & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(Neto * 100, "000000000000000") & Format(Iva, "0000") & impLIQ & impRNI & impEXE & "000000000000000" & "000000000000000" & "000000000000000" & "000000000000000" & "00" & Chr(32) & Chr(32) & Chr(32) & "0000000000" & cantIVA & Chr(32) & "00000000000000" & "00000000" & "00000000" & _
                        relleno2 & "00000000" & "000000000000000"
                    netNOgra = netNOgra - CDbl(rs3!Total)
                    i = i + 1
                ElseIf Trim(rs3!TIPODOC) = "NDB" Then
                    relleno2 = ""
                    j = 1
                    While j <= 75
                        relleno2 = relleno2 & Chr(32)
                        j = j + 1
                    Wend
        ''            Print #1, "1" & fec & "07" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & "E" & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                        "Sin comentarios" & relleno2 '& vbCrLf
    '''                Print #1, "1" & fec & "07" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0000000000" & "1" & Chr(32) & Format(nSinNull(rs3!CAE), "00000000000000") & Year(rs3!CAEV) & Format(Month(rs3!CAEV), "00") & Format(Day(rs3!CAEV), "00") & "00000000" & _
                        relleno2 & "00000000" & "000000000000000"
                    Print #1, "1" & fec & "07" & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(Neto * 100, "000000000000000") & Format(Iva, "0000") & impLIQ & impRNI & impEXE & "000000000000000" & "000000000000000" & "000000000000000" & "000000000000000" & "00" & Chr(32) & Chr(32) & Chr(32) & "0000000000" & cantIVA & Chr(32) & "00000000000000" & "00000000" & "00000000" & _
                        relleno2 & "00000000" & "000000000000000"
                    netNOgra = netNOgra + CDbl(rs3!Total)
                    i = i + 1
                End If
                k = k + 1
            Else
                
                rs4.Open "select *,_iva as iva from facturaventadetalle where codigofactura=" & rs3!codigo & " order by _iva", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                
                rs2.Open "select i.*,C.CUIT from clientes c inner join ivas i on i.codigo=c.iva where c.codigo=" & rs3!cliente, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                If rs3!tipoiva = 2 Or rs3!tipoiva = 4 Or rs3!tipoiva = 7 Then 'revisar
                    documento = 80
                ElseIf rs3!tipoiva = 1 Then
                    documento = IIf(Mid(sSinNull(rs3!CUIT), 1, 4) = "00-0", 96, 86) '96=dni,86=cuil
                Else
                    documento = 86
                End If
                nom = Mid(Trim(rs3!RAZONSOCIAL), 1, 30)
                While Len(nom) < 30
                    nom = Chr(32) & nom
                Wend
                
                fec = Year(rs3!Fecha) & Format(Month(rs3!Fecha), "00") & Format(Day(rs3!Fecha), "00")
                tI = Trim(rs3!tipoiva)
                If tI = 1 Or tI = 4 Or tI = 7 Or tI = 8 Then 'facturas B
                    impLIQ = "000000000000000"
                    impRNI = "000000000000000"
                    impEXE = "000000000000000"
                ElseIf tI = 2 Or tI = 3 Then  'facturas A
                    If tI = 2 Then
                        impLIQ = Format(rs3!Iva * 100, "000000000000000")
                        impRNI = "000000000000000"
                        impEXE = "000000000000000"
                    ElseIf tI = 3 Then
                        impLIQ = "000000000000000"
                        impRNI = Format(rs3!Iva * 100, "000000000000000")
                        impEXE = "000000000000000"
                    ElseIf tI = 8 Then ' a este no entra nunca, pero asi estaba en tavi...
                        impLIQ = "000000000000000"
                        impRNI = "000000000000000"
                        impEXE = Format(rs3!Iva * 100, "000000000000000")
                    End If
                End If
                
                If Trim(rs3!TIPODOC) = "FAA" Or Trim(rs3!TIPODOC) = "FAB" Or Trim(rs3!TIPODOC) = "NDA" Or Trim(rs3!TIPODOC) = "NDB" Then
                    tot = tot + CDbl(rs3!Total)
                    totimpLIQ = totimpLIQ + CDbl(impLIQ)
                    totimpRNI = totimpRNI + CDbl(impRNI)
                    totimpEXE = totimpEXE + CDbl(impEXE)
                    IIBB = IIBB + CDbl(rs3!IIBB)
                ElseIf Trim(rs3!TIPODOC) = "NCA" Or Trim(rs3!TIPODOC) = "NCB" Then
                    tot = tot - CDbl(rs3!Total)
                    totimpLIQ = totimpLIQ - CDbl(impLIQ)
                    totimpRNI = totimpRNI - CDbl(impRNI)
                    totimpEXE = totimpEXE - CDbl(impEXE)
                    IIBB = IIBB - CDbl(rs3!IIBB)
                End If
                '***************************
                If Trim(rs3!TIPODOC) = "FAA" Then
                    Tipo = "01"
                ElseIf Trim(rs3!TIPODOC) = "FAB" Then
                    Tipo = "06"
                ElseIf Trim(rs3!TIPODOC) = "NCA" Then
                    Tipo = "03"
                ElseIf Trim(rs3!TIPODOC) = "NCB" Then
                    Tipo = "08"
                ElseIf Trim(rs3!TIPODOC) = "NDA" Then
                    Tipo = "02"
                ElseIf Trim(rs3!TIPODOC) = "NDB" Then
                    Tipo = "07"
                End If
                impLIQ2 = 0
                Neto = 0
                Do While Not rs4.EOF
                    If tI = 1 Or tI = 4 Or tI = 7 Or tI = 8 Then 'facturas B
                        impLIQ2 = impLIQ2 + s2n(rs4!PrecioTotal - (rs4!PrecioTotal / (1 + (rs4!Iva / 100)))) * 100
                        impLIQ = Format(s2n(impLIQ2), "000000000000000")
                        'impLIQ = Format(s2n(rs4!PrecioTotal - (rs4!PrecioTotal / (1 + (rs4!Iva / 100)))) * 100, "000000000000000")
                        Iva = Format(rs4!Iva * 100, "0000")
                        Neto = Neto + rs4!PrecioTotal - s2n(rs4!PrecioTotal - (rs4!PrecioTotal / (1 + (rs4!Iva / 100))))
                    ElseIf tI = 2 Then
                        impLIQ2 = impLIQ2 + s2n(rs4!PrecioTotal * (rs4!Iva / 100)) * 100
                        impLIQ = Format(s2n(impLIQ), "000000000000000")
                        'impLIQ = Format(s2n(rs4!PrecioTotal * (rs4!Iva / 100)) * 100, "000000000000000")
                        Iva = Format(rs4!Iva * 100, "0000")
                        Neto = Neto + rs4!PrecioTotal
                    End If
    '                If rs4.AbsolutePosition <> rs4.RecordCount Then
    '                    relleno2 = ""
    '                    j = 1
    '                    While j <= 75
    '                        relleno2 = relleno2 & Chr(32)
    '                        j = j + 1
    '                    Wend
    '                    'Print #1, "1" & fec & tipo & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(0, "000000000000000") & "000000000000000" & Format(rs4!PrecioTotal * 100, "000000000000000") & (rs4!Iva * 100) & impLIQ & impRNI & impEXE & "000000000000000" & "000000000000000" & "000000000000000" & "000000000000000" & "00" & Chr(32) & Chr(32) & Chr(32) & "0000000000" & cantIVA & Chr(32) & "00000000000000" & "00000000" & "00000000" & _
    '                        relleno2 & "00000000" & "000000000000000"
    '                    Print #1, "1" & fec & tipo & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(0, "000000000000000") & "000000000000000" & Format(Neto * 100, "000000000000000") & (rs4!Iva * 100) & impLIQ & impRNI & impEXE & "000000000000000" & "000000000000000" & "000000000000000" & "000000000000000" & "00" & Chr(32) & Chr(32) & Chr(32) & "0000000000" & cantIVA & Chr(32) & "00000000000000" & "00000000" & "00000000" & _
    '                        relleno2 & "00000000" & "000000000000000"
                        rs4.MoveNext
    '                Else
    '                    Exit Do
    '                End If
                Loop
                Set rs4 = Nothing
                '*********************************************
                
                If Trim(rs3!TIPODOC) = "FAA" Then  'en las tablas de la afip el 01 es factura A
                    relleno2 = ""
                    j = 1
                    While j <= 75
                        relleno2 = relleno2 & Chr(32)
                        j = j + 1
                    Wend
        ''            Print #1, "1" & fec & "01" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & "2100" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & Chr(32) & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                    "Sin comentarios" & relleno2
                    Print #1, "1" & fec & "01" & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & "2100" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0000000000" & "1" & Chr(32) & Format(nSinNull(rs3!CAE), "00000000000000") & Year(nSinNull(rs3!CAEV)) & Format(Month(nSinNull(rs3!CAEV)), "00") & Format(Day(nSinNull(rs3!CAEV)), "00") & "00000000" & _
                    relleno2 & "00000000" & "000000000000000"
                    'Close #1
                    netgra = netgra + CDbl(rs3!Neto)
                    i = i + 1
                ElseIf Trim(rs3!TIPODOC) = "FAB" Then
                    'If CDbl(txtTotal.caption) >= CDbl("10000") Then  'lo quieren hacer uno por uno
                        'Open "C:\gestion\SIRED\" & arch & ".txt" For Output As #1
                        'Write #1, cmbCliente.Text & cmbCliente.Text & Chr(32) & cmbCliente.Text            ' Text1.Text el chr(9)=tab  vbCrLf=enter
                        relleno2 = ""
                        j = 1
                        While j <= 75
                            relleno2 = relleno2 & Chr(32)
                            j = j + 1
                        Wend
        ''                Print #1, "1" & fec & "06" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & "E" & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                        "Sin comentarios" & relleno2 '& vbCrLf
                        'Print #1, "1" & fec & "06" & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0000000000" & "1" & Chr(32) & Format(nSinNull(rs3!CAE), "00000000000000") & Year(nSinNull(rs3!CAEV)) & Format(Month(nSinNull(rs3!CAEV)), "00") & Format(Day(nSinNull(rs3!CAEV)), "00") & "00000000" & _
                            relleno2 & "00000000" & "000000000000000"
                        Print #1, "1" & fec & "06" & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(0, "000000000000000") & Format(Neto * 100, "000000000000000") & Format(Iva, "0000") & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & "00" & Chr(32) & Chr(32) & Chr(32) & "0000000000" & cantIVA & Chr(32) & "00000000000000" & "00000000" & "00000000" & _
                            relleno2 & "00000000" & "000000000000000"
                        'txtneto
                        netNOgra = netNOgra + CDbl(rs3!Total)
                        'Close #1
                    'End If
                    i = i + 1
                ElseIf Trim(rs3!TIPODOC) = "NCA" Then
                    relleno2 = ""
                    j = 1
                    While j <= 75
                        relleno2 = relleno2 & Chr(32)
                        j = j + 1
                    Wend
        ''            Print #1, "1" & fec & "03" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & "2100" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & Chr(32) & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                    "Sin comentarios" & relleno2
                    Print #1, "1" & fec & "03" & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & "2100" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0000000000" & "1" & Chr(32) & Format(nSinNull(rs3!CAE), "00000000000000") & Year(nSinNull(rs3!CAEV)) & Format(Month(nSinNull(rs3!CAEV)), "00") & Format(Day(nSinNull(rs3!CAEV)), "00") & "00000000" & _
                    relleno2 & "00000000" & "000000000000000"
                    netgra = netgra - CDbl(rs3!Neto)
                    i = i + 1
                ElseIf Trim(rs3!TIPODOC) = "NDA" Then
                    relleno2 = ""
                    j = 1
                    While j <= 75
                        relleno2 = relleno2 & Chr(32)
                        j = j + 1
                    Wend
        ''            Print #1, "1" & fec & "02" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & "2100" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & Chr(32) & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                    "Sin comentarios" & relleno2
                    Print #1, "1" & fec & "02" & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & "2100" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0000000000" & "1" & Chr(32) & Format(nSinNull(rs3!CAE), "00000000000000") & Year(nSinNull(rs3!CAEV)) & Format(Month(nSinNull(rs3!CAEV)), "00") & Format(Day(nSinNull(rs3!CAEV)), "00") & "00000000" & _
                    relleno2 & "00000000" & "000000000000000"
                    netgra = netgra + CDbl(rs3!Neto)
                    i = i + 1
                ElseIf Trim(rs3!TIPODOC) = "NCB" Then
                    relleno2 = ""
                    j = 1
                    While j <= 75
                        relleno2 = relleno2 & Chr(32)
                        j = j + 1
                    Wend
        ''            Print #1, "1" & fec & "08" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & "E" & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                        "Sin comentarios" & relleno2 '& vbCrLf
                    Print #1, "1" & fec & "08" & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0000000000" & "1" & Chr(32) & Format(nSinNull(rs3!CAE), "00000000000000") & Year(nSinNull(rs3!CAEV)) & Format(Month(nSinNull(rs3!CAEV)), "00") & Format(Day(nSinNull(rs3!CAEV)), "00") & "00000000" & _
                        relleno2 & "00000000" & "000000000000000"
                    netNOgra = netNOgra - CDbl(rs3!Total)
                    i = i + 1
                ElseIf Trim(rs3!TIPODOC) = "NDB" Then
                    relleno2 = ""
                    j = 1
                    While j <= 75
                        relleno2 = relleno2 & Chr(32)
                        j = j + 1
                    Wend
        ''            Print #1, "1" & fec & "07" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & "E" & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                        "Sin comentarios" & relleno2 '& vbCrLf
                    Print #1, "1" & fec & "07" & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0000000000" & "1" & Chr(32) & Format(nSinNull(rs3!CAE), "00000000000000") & Year(nSinNull(rs3!CAEV)) & Format(Month(nSinNull(rs3!CAEV)), "00") & Format(Day(nSinNull(rs3!CAEV)), "00") & "00000000" & _
                        relleno2 & "00000000" & "000000000000000"
                    netNOgra = netNOgra + CDbl(rs3!Total)
                    i = i + 1
                End If
            End If
        End If
        
        rs3.MoveNext
        Set rs2 = Nothing
    Wend
    'registro de tipo 2
    j = 1
    While j < 123
        relleno = relleno & Chr(32)
        j = j + 1
    Wend
    relleno2 = ""
    j = 1
    While j < 30
        relleno2 = relleno2 & Chr(32)
        j = j + 1
    Wend
    
''    Print #1, "2" & Year(dtperiodo.Value) & Format(Month(dtperiodo.Value), "00") & relleno2 & Format(i, "000000000000") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Trim(Mid(rs!CuitEmpresa, 1, 2)) & Trim(Mid(rs!CuitEmpresa, 4, 8)) & Trim(Mid(rs!CuitEmpresa, 13, 1)) & relleno2 & Chr(32) & _
     Format(tot * 100, "000000000000000") & Format(netNOgra * 100, "000000000000000") & Format(netgra * 100, "000000000000000") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Format(totimpLIQ, "000000000000000") & Format(totimpRNI, "000000000000000") & Format(totimpEXE, "000000000000000") & "000000000000000" & Format(IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & relleno
    
    
    If Not IsNull(rs3) Or Not IsEmpty(rs3) Or Not (rs3.EOF = True And rs3.BOF = True) Then
        Close #1
    End If
    Set rs3 = Nothing
    Set rs = Nothing
    MsgBox "Archivo generado correctamente."
                
            
    '******************************************************************************************


End Sub

Private Sub Command8_Click()
    Dim ARCH As String
    Dim nom As String
    Dim impLIQ As String
    Dim fec As String
    Dim pri As Date
    Dim seg As Date
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim rs3 As New ADODB.Recordset
    Dim i As Long
    Dim j As Long
    Dim relleno As String
    Dim cui As String
    Dim Tipo As String
        
    pri = "01/" & Month(dtPeriodo.Value) & "/" & Year(dtPeriodo.Value)
    seg = ultimoDiaDelMes(dtPeriodo.Value)
    'rs3.Open "select fecha,codpr,tipoiva,tipodoc,nrodoc,total,neto,exento,iva_21,iva_27,iva_9,iva_10,imp_int,percepc,der_est,ibcapital,ibprovincia,retgan,retganpago,formadepago,razonsocialprov,cuitprov,suc,anoimp,tipocompro from TRANSCOM where fecha>=" & ssFecha(pri) & " and fecha<=" & ssFecha(seg) & " and tipodoc<>'RAC'" & _
        "union select fecha,codpr,tipoiva,tipodoc,nrodoc,total,neto,exento,iva_21,iva_27,iva_9,iva_10,imp_int,percepc,der_est,ibcapital,ibprovincia,retgan,retganpago,formadepago,razonsocialprov,cuitprov,suc,anoimp,tipocompro from compras where fecha>=" & ssFecha(pri) & " and fecha<=" & ssFecha(seg) & " and tipodoc<>'RAC' order by fecha,tipodoc,nrodoc", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    rs3.Open "select fecha,codpr,nrodoc,iva_21,iva_27,iva_9,iva_10,imp_int,razonsocialprov,cuitprov,suc,anoimp,tipocompro,(iva_21+iva_27+iva_10) as iva " & _
                " from TRANSCOM " & _
                " where fecha>=" & ssFecha(pri) & " and fecha<=" & ssFecha(seg) & " and tipodoc<>'RAC'" & _
                " group by fecha,codpr,nrodoc,iva_21,iva_27,iva_9,iva_10,imp_int,razonsocialprov,cuitprov,suc,anoimp,tipocompro " & _
                " Having (IVA_21 + IVA_27 + iva_10) >= 500 " & _
        "union " & _
            " select fecha,codpr,nrodoc,iva_21,iva_27,iva_9,iva_10,imp_int,razonsocialprov,cuitprov,suc,anoimp,tipocompro,(iva_21+iva_27+iva_10) as iva " & _
                " from compras " & _
                " where fecha>=" & ssFecha(pri) & " and fecha<=" & ssFecha(seg) & " and tipodoc<>'RAC' " & _
                " group by fecha,codpr,nrodoc,iva_21,iva_27,iva_9,iva_10,imp_int,razonsocialprov,cuitprov,suc,anoimp,tipocompro " & _
                " Having (IVA_21 + IVA_27 + iva_10) >= 500 " & _
                " order by fecha,nrodoc", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'********************************** ARCHIVO
        ARCH = "COMPRAS_" & Year(dtPeriodo.Value) & Format(Month(dtPeriodo.Value), "00")
        rs.Open "select * from datosempresa where idempresa=4", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        
        If (rs3.EOF = True And rs3.BOF = True) Or IsNull(rs3) Then
            MsgBox "No se ha encontrado datos para este periodo.", , "ATENCION"
            Exit Sub
        End If
        
        rs3.MoveFirst
        If Not IsNull(rs3) Or Not IsEmpty(rs3) Or Not (rs3.EOF = True And rs3.BOF = True) Then
            'Call Verificar_Existe("C:\CITI\") '& ARCH & ".txt"

            Open "C:\CITI\" & ARCH & ".txt" For Output As #1
        End If
        i = 0
        While Not rs3.EOF
            
                rs2.Open "select i.*,C.CUIT from prov c inner join ivas i on i.codigo=c.tipoiva where c.codigo=" & rs3!CODPR, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                Tipo = Format(rs3!tipocompro, "00")
                If rs3!cuitprov = "" Then
                    cui = Format(0, "00000000000")
                Else
                    cui = Trim(Mid(rs3!cuitprov, 1, 2)) & Trim(Mid(rs3!cuitprov, 4, 8)) & Trim(Mid(rs3!cuitprov, 13, 1))
                End If
                nom = Trim(rs3!razonsocialprov)
                While Len(nom) < 25
                    nom = Chr(32) & nom
                Wend
                
                fec = Year(rs3!Fecha) & Format(Month(rs3!Fecha), "00") & Format(Day(rs3!Fecha), "00")
                
                impLIQ = Format((rs3!IVA_21 + rs3!IVA_27 + rs3!iva_10) * 100, "000000000000")
                
                j = 1
                relleno = " "
                While j < 25
                    relleno = relleno & Chr(32)
                    j = j + 1
                Wend
                
                
                Print #1, Tipo & Format(rs3!suc, "0000") & Format(rs3!NroDoc, "00000000000000000000") & fec & cui & nom & _
                    impLIQ & "00000000000" & relleno & "000000000000"
    
                rs3.MoveNext
                Set rs2 = Nothing
        Wend
        
        If Not IsNull(rs3) Or Not IsEmpty(rs3) Or Not (rs3.EOF = True And rs3.BOF = True) Then
            Close #1
        End If
        Set rs3 = Nothing
        Set rs = Nothing
        MsgBox "Archivo generado correctamente."
                
            
    '******************************************************************************************

End Sub


Private Sub Form_Load()
    dtPeriodo.Value = Date
    Me.caption = Me.caption & " - Cliente : " & Cuit_Empresa_Carga
    
End Sub



Private Function CitiVentas()
On Error GoTo ventas_err
Dim vFecha As String, vTipo As String, vPunto As String, vNro As String, vNroHasta As String, vCodigo As String, vIdentificador As String, vApellidoCliente As String, vTotal As String, vNoGravado As String, vPercepcionNN As String, vExentas As String, vPercepciones As String, vIngresosBrutos As String, vMunicipales As String, vInternos As String, vCodigoMoneda As String, vCambio As String, vCantidadAlicuotas As String, vCodOperacion As String, vTributos As String, vVencimiento As String
Dim vPrimerDia As Date, vUltimoDia As Date, str As String, i As Long, sNombreFile2 As String, sNombreFile As String, sCarpeta As String, sArchivoCompleto As String, sArchivoCompleto2 As String
Dim vNeto As String, vAlicuota As String, vImpuestoLiquidado As String, nnNeto As Double, a As Integer
Dim rs1 As New ADODB.Recordset

    vPrimerDia = "01/" & Month(dtPeriodo) & "/" & Year(dtPeriodo)
    vUltimoDia = ultimoDiaDelMes(dtPeriodo)
    str = "select * from facturaventa where fecha>=" & ssFecha(vPrimerDia) & " and fecha<=" & ssFecha(vUltimoDia) & " and tipodoc in ('NDB','NDA','FAE','FAA','FAB','NCA','NCB') order by fecha,tipodoc,nrofactura"
    rs1.Open str, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly

    With rs1
        If .EOF And .BOF Then
            MsgBox "No se ha encontrado datos para este periodo.", vbInformation, "ATENCION"
            Exit Function
        Else
            sCarpeta = "C:\"
            sCarpeta = Trim(VentanaCarpeta("Carpeta Destino txt VENTAS", sCarpeta))
            If Trim(sCarpeta) = "" Then Exit Function
            If CORTO(sCarpeta, Len(sCarpeta) - 1, 0) <> "\" Then sCarpeta = sCarpeta & "\"
            sNombreFile = "VENTAS_" & Year(dtPeriodo) & Format(Month(dtPeriodo), "00") & ".txt"
            sNombreFile2 = "VENTAS_ALICUOTAS_" & Year(dtPeriodo) & Format(Month(dtPeriodo), "00") & ".txt"
            sArchivoCompleto = sCarpeta & sNombreFile
            sArchivoCompleto2 = sCarpeta & sNombreFile2
            If ExisteArchivo(sArchivoCompleto) Then
                Kill sArchivoCompleto
            End If
            If ExisteArchivo(sArchivoCompleto2) Then
                Kill sArchivoCompleto2
            End If
            Open sArchivoCompleto For Output As #1
            Open sArchivoCompleto2 For Output As #2
            
            .MoveFirst
            For i = 0 To .RecordCount - 1
                vFecha = "19000101"
                vTipo = "000"
                vPunto = "00000"
                vNro = "00000000000000000000"
                vNroHasta = "00000000000000000000"
                vCodigo = "00"
                vIdentificador = "00000000000000000000"
                vApellidoCliente = "----------SIN DATOS ----------"
                vTotal = "000000000000000"
                vNoGravado = "000000000000000"
                vPercepcionNN = "000000000000000"
                vExentas = "000000000000000"
                vPercepciones = "000000000000000"
                vIngresosBrutos = "000000000000000"
                vMunicipales = "000000000000000"
                vInternos = "000000000000000"
                vCodigoMoneda = "PES"
                vCambio = "0001000000"
                vCantidadAlicuotas = "1"
                vCodOperacion = "0"
                vTributos = "000000000000000"
                vVencimiento = "00000000"
            
                vNeto = "000000000000000"
                vAlicuota = "0000"
                vImpuestoLiquidado = "000000000000000"
                
                
                vFecha = Format(!Fecha, "YYYYMMDD")
                
                'If Trim(!TIPODOC) = "NCA" Then Stop
                
                Select Case Trim(!TIPODOC)
                    Case "FAA": vTipo = "001"
                    Case "FAB": vTipo = "006"
                    Case "NCA": vTipo = "003"
                    Case "NCB": vTipo = "008"
                    Case "NDA": vTipo = "002"
                    Case "NDB": vTipo = "007"
                    Case "FAE": vTipo = "019"
                    Case "NCE": vTipo = "021"
                End Select
                
                If vTipo <> "019" And vTipo <> "021" Then
                    vVencimiento = Format(!Vencimiento, "YYYYMMDD")
                End If
                
                'If vTipo = "019" Then Stop
                'If vTipo = "002" Then Stop
                
                vPunto = Format(!PuntoVenta, "00000")
                vNro = Format(!NroFactura, "00000000000000000000")
                vNroHasta = vNro
                'If vNro = "00000000000000000002" Then Stop
                vIdentificador = Replace(!CUIT, "-", "")
                vIdentificador = s2n(vIdentificador)
                                    
                Select Case !tipoiva
                    Case 2, 4, 7: vCodigo = 80 'cuit
                    Case 1: vCodigo = IIf(Len(vIdentificador) > 8, 86, 96) '96=dni,86=cuil
                    Case Else: vCodigo = 86
                End Select
                
                vIdentificador = Format(vIdentificador, "00000000000000000000")
                
                vApellidoCliente = Trim(!RAZONSOCIAL)
                If Len(vApellidoCliente) > 30 Then
                    vApellidoCliente = CORTO(vApellidoCliente, 0, Len(vApellidoCliente) - 30)
                Else
                    While Len(vApellidoCliente) < 30
                        vApellidoCliente = Chr(32) & vApellidoCliente
                    Wend
                End If
                
                vCantidadAlicuotas = obtenerDeSQL("select count(distinct(_iva)) from facturaventadetalle where codigofactura=" & !codigo)
                

                
                
                vTotal = Format(Replace(s2n(!Total, 2, True), ",", ""), "000000000000000")
                'vNoGravado = Format(Replace(s2n(!NoGrav, 2, True), ",", ""), "000000000000000")
                If vTipo = "019" Or vTipo = "021" Then
                    'vExentas = vTotal 'Format(!NoGrav, "000000000000000")
                    vNoGravado = Format(Replace(s2n(!NoGrav, 2, True), ",", ""), "000000000000000")
                Else
                    vNoGravado = nSinNull(obtenerDeSQL("select sum(preciototal) as preciototal from facturaventadetalle where _iva=0 and codigofactura=" & !codigo))
'                    If s2n(vNoGravado) > 0 Then
'                        vTotal = Format(Replace(s2n(!total - s2n(vNoGravado), 2, True), ",", ""), "000000000000000")
'                    End If
                    vNoGravado = Format(Replace(s2n(vNoGravado, 2, True), ",", ""), "000000000000000")
                End If
                
                If !ND_xChequeRechazado Then
                    vCodOperacion = "A"
                ElseIf CORTO(!TIPODOC, 2, 0) = "E" Then
                    vCodOperacion = "X"
                ElseIf s2n(!Total) = s2n(!NoGrav) Then
                    vCodOperacion = "E"
                End If
                
                vIngresosBrutos = Format(Replace(s2n(!IIBB, 2, True), ",", ""), "000000000000000")
                
'                nnNeto = s2n(!Neto)
'                If nnNeto = 0 Then
'                    nnNeto = s2n(!total)
'                End If
'
'                If vTipo = "019" Or vTipo = "021" Then
'                    vAlicuota = "0003"
'                ElseIf !ND_xChequeRechazado Then
'                    vAlicuota = "0003"
'                    vNoGravado = Format(Replace(s2n(!Neto, 2, True), ",", ""), "000000000000000")
'                    vNeto = "000000000000000"
'                    'vImpuestoLiquidado = Format(Replace(s2n(!Iva, 2, True), ",", ""), "000000000000000")
'                    vImpuestoLiquidado = "000000000000000"
'                ElseIf s2n(s2n(!total) / s2n(nnNeto)) > 1.2 Then
'                    vAlicuota = "0005"
'                    vNeto = Format(Replace(s2n(!Neto, 2, True), ",", ""), "000000000000000")
'                    vImpuestoLiquidado = Format(Replace(s2n(!Iva, 2, True), ",", ""), "000000000000000")
'                ElseIf s2n(s2n(!total) / s2n(nnNeto)) < 1.2 Then
'                    vAlicuota = "0004"
'                    vNeto = Format(Replace(s2n(!Neto, 2, True), ",", ""), "000000000000000")
'                    vImpuestoLiquidado = Format(Replace(s2n(!Iva, 2, True), ",", ""), "000000000000000")
'                End If


                a = 0
                If vTipo = "019" Or vTipo = "021" Then
                    vAlicuota = "0003"
                    Print #2, vTipo & vPunto & vNro & vNeto & vAlicuota & vImpuestoLiquidado
'                ElseIf vTipo = "006" Or vTipo = "007" Or vTipo = "008" Then
'                    nIva21 = nSinNull(obtenerDeSQL("select (sum(preciototal)/1.21) as preciototal from facturaventadetalle where _iva=21 and codigofactura=" & !codigo))
'                    If s2n(nIva21) > 0 Then
'                        a = a + 1
'                        vAlicuota = "0005"
'                        If !ND_xChequeRechazado Then vAlicuota = "0003"
'                        vNeto = Format(Replace(s2n(nIva21, 2, True), ",", ""), "000000000000000")
'                        vImpuestoLiquidado = Format(Replace(s2n(nIva21 * 0.21, 2, True), ",", ""), "000000000000000")
'                        Print #2, vTipo & vPunto & vNro & vNeto & vAlicuota & vImpuestoLiquidado
'                    End If
'                    nIva10 = nSinNull(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where _iva=10.5 and codigofactura=" & !codigo))
'                    If s2n(nIva10) > 0 Then
'                        a = a + 1
'                        vAlicuota = "0004"
'                        If !ND_xChequeRechazado Then vAlicuota = "0003"
'                        vNeto = Format(Replace(s2n(nIva10, 2, True), ",", ""), "000000000000000")
'                        vImpuestoLiquidado = Format(Replace(s2n(nIva10 * 0.105, 2, True), ",", ""), "000000000000000")
'                        Print #2, vTipo & vPunto & vNro & vNeto & vAlicuota & vImpuestoLiquidado
'                    End If
                Else
                    If vTipo = "006" Or vTipo = "007" Or vTipo = "008" Then
                        nIva21 = nSinNull(obtenerDeSQL("select (sum(preciototal)/1.21) as preciototal from facturaventadetalle where _iva=21 and codigofactura=" & !codigo))
                    'ElseIf vTipo = "006" Then
                    '    nIva21 = nSinNull(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where _iva=21 and codigofactura=" & !codigo))
                    Else
                        nIva21 = nSinNull(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where _iva=21 and codigofactura=" & !codigo))
                    End If
                    If s2n(nIva21) > 0 Then
                        a = a + 1
                        vAlicuota = "0005"
                        If !ND_xChequeRechazado Then vAlicuota = "0003"
                        vNeto = Format(Replace(s2n(nIva21, 2, True), ",", ""), "000000000000000")
                        vImpuestoLiquidado = Format(Replace(s2n(nIva21 * 0.21, 2, True), ",", ""), "000000000000000")
                        Print #2, vTipo & vPunto & vNro & vNeto & vAlicuota & vImpuestoLiquidado
                    End If
                    
                    If vTipo = "006" Or vTipo = "007" Or vTipo = "008" Then
                        nIva10 = nSinNull(obtenerDeSQL("select (sum(preciototal)/1.105) as preciototal  from facturaventadetalle where _iva=10.5 and codigofactura=" & !codigo))
                    Else
                        nIva10 = nSinNull(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where _iva=10.5 and codigofactura=" & !codigo))
                    End If
                    If s2n(nIva10) > 0 Then
                        a = a + 1
                        vAlicuota = "0004"
                        If !ND_xChequeRechazado Then vAlicuota = "0003"
                        vNeto = Format(Replace(s2n(nIva10, 2, True), ",", ""), "000000000000000")
                        vImpuestoLiquidado = Format(Replace(s2n(nIva10 * 0.105, 2, True), ",", ""), "000000000000000")
                        Print #2, vTipo & vPunto & vNro & vNeto & vAlicuota & vImpuestoLiquidado
                    End If
                    
                    
                    
                    If !ND_xChequeRechazado And a = 0 Or (vCodOperacion = "E") Then
                        a = a + 1
                        vAlicuota = "0003"
                        vNeto = Format(0, "000000000000000")
                        vImpuestoLiquidado = Format(0, "000000000000000")
                        Print #2, vTipo & vPunto & vNro & vNeto & vAlicuota & vImpuestoLiquidado
                    End If
                    
                End If
                vCantidadAlicuotas = a
                
                
                
                Print #1, vFecha & vTipo & vPunto & vNro & vNroHasta & vCodigo & vIdentificador & vApellidoCliente & vTotal & vNoGravado & vPercepcionNN & vExentas & vPercepciones & vIngresosBrutos & vMunicipales & vInternos & vCodigoMoneda & vCambio & vCantidadAlicuotas & vCodOperacion & vTributos & vVencimiento
                'Print #2, vTipo & vPunto & vNro & vNeto & vAlicuota & vImpuestoLiquidado
            
                .MoveNext
            Next
        End If
    End With

    
    Close #1
    Close #2
    Set rs1 = Nothing

    MsgBox "Archivo generado correctamente.", vbInformation, "Ventas"
                
            
Exit Function
ventas_err:
MsgBox "Error... " & Err.Number & " " & Err.Description
End Function

Private Function CitiCompras()
On Error GoTo compras_err
Dim cFecha As String, cTipo As String, cPunto As String, cNro As String, cNroDespacho As String, cCodigo As String, cIdentificador As String, cApellidoProveedor As String, cTotal As String, cNoGravado As String, cExentas As String, cPercepciones1 As String, cPercepciones2 As String, cIngresosBrutos As String, cMunicipales As String, cInternos As String, cCodigoMoneda As String, cCambio As String, cCantidadAlicuotas As String, cCodOperacion As String, cCreditoFiscal As String, cTributos As String, cCuitEmisor As String, cDenominacionEmisor As String, cComision As String
Dim cPrimerDia As Date, cUltimoDia As Date, str As String, i As Long, sNombreFile As String, sCarpeta As String, sArchivoCompleto As String, sArchivoCompleto2 As String, sArchivoCompleto3 As String, pLetra As String, pTipoDoc As String, a As Integer
Dim cNeto As String, cAlicuota As String, cImpuestoLiquidado As String, cImpuestos As Double
Dim dDESPACHO As String, dNETO As String, dALICUOTA As String, dIMPUESTO As String
Dim rs1 As New ADODB.Recordset

    cPrimerDia = "01/" & Month(dtPeriodo) & "/" & Year(dtPeriodo)
    cUltimoDia = ultimoDiaDelMes(dtPeriodo)
    
    If optFecha.Value = True Then
        rs1.Open "select fecha,codpr,nrodoc,iva_21,iva_27,0 as iva_9,iva_10,imp_int,razonsocialprov,cuitprov,suc,anoimp,TIPODOC,TOTAL,NETO,EXENTO,(IBCAPITAL+IBPROVINCIA) AS IIBB,LETRA,(PERCEPC+DER_EST+iva_9) AS RETIVA,(otrasretimp + OTRASRET) as retiva2,'' as nrodespacho " & _
                    " from TRANSCOM " & _
                    " where fecha>=" & ssFecha(cPrimerDia) & " and fecha<=" & ssFecha(cUltimoDia) & " and tipodoc<>'RAC' and tipodoc<>'APC' and tipodoc<>'APD' " & _
            "union " & _
                " select fecha,codpr,nrodoc,iva_21,iva_27,0 as iva_9,iva_10,imp_int,razonsocialprov,cuitprov,suc,anoimp, TIPODOC,TOTAL,NETO,EXENTO,(IBCAPITAL+IBPROVINCIA) AS IIBB,LETRA,(PERCEPC+DER_EST+iva_9) AS RETIVA,(otrasretimp + OTRASRET)as retiva2,'' as  nrodespacho " & _
                    " from compras " & _
                    " where fecha>=" & ssFecha(cPrimerDia) & " and fecha<=" & ssFecha(cUltimoDia) & " and tipodoc<>'RAC' and tipodoc<>'APC' and tipodoc<>'APD' ", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
            '"union " & _
                " select fecha,proveedor as codpr,0 as nrodoc, iva as iva21,0 as iva_27,0 as iva_9,0 as iva_10,0 as imp_int,descripcion as razonsocialprov, cuit as cuitprov, 0 as suc,year(fecha) as fecha,'66' as tipodoc,(base + total) as total,base as neto,exento,perciibb as iibb,'D' as letra,percrg3431 as  retiva,percganancia as  retiva2,nrodespacho " & _
                " from despachodeimportacion inner join prov  on despachodeimportacion.proveedor=prov.codigo " & _
                " where cuitcliente=" & ssTexto(Cuit_Empresa_Carga) & " and fecha>=" & ssFecha(cPrimerDia) & " and fecha<=" & ssFecha(cUltimoDia) & "  order by fecha,nrodespacho ", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    Else
        rs1.Open "select fecha,codpr,nrodoc,iva_21,iva_27,0 as iva_9,iva_10,imp_int,razonsocialprov,cuitprov,suc,anoimp,TIPODOC,TOTAL,NETO,EXENTO,(IBCAPITAL+IBPROVINCIA) AS IIBB,LETRA,(PERCEPC+DER_EST+iva_9) AS RETIVA,(otrasretimp + OTRASRET)as retiva2,'' as nrodespacho " & _
                    " from TRANSCOM " & _
                    " where mesimp=" & Month(dtPeriodo) & " and anoimp=" & Year(dtPeriodo) & " and tipodoc<>'RAC' and tipodoc<>'APC' and tipodoc<>'APD' " & _
            "union " & _
                " select fecha,codpr,nrodoc,iva_21,iva_27,0 as iva_9,iva_10,imp_int,razonsocialprov,cuitprov,suc,anoimp,TIPODOC,TOTAL,NETO,EXENTO,(IBCAPITAL+IBPROVINCIA) AS IIBB,LETRA,(PERCEPC+DER_EST+iva_9) AS RETIVA,(otrasretimp + OTRASRET) as retiva2,'' as  nrodespacho " & _
                    " from compras " & _
                    " where mesimp=" & Month(dtPeriodo) & " and anoimp=" & Year(dtPeriodo) & " and tipodoc<>'RAC' and tipodoc<>'APC' and tipodoc<>'APD' ", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    End If
    
    With rs1
        
        If rs1.EOF And rs1.BOF Then
            MsgBox "No se ha encontrado datos para este periodo.", vbInformation, "ATENCION"
            Exit Function
        Else
            
            sCarpeta = "C:\"
            sCarpeta = Trim(VentanaCarpeta("Carpeta Destino txt COMPRAS", sCarpeta))
            If Trim(sCarpeta) = "" Then Exit Function
            If CORTO(sCarpeta, Len(sCarpeta) - 1, 0) <> "\" Then sCarpeta = sCarpeta & "\"
            sNombreFile = "COMPRAS_" & Year(dtPeriodo) & Format(Month(dtPeriodo), "00") & "_" & Cuit_Empresa_Carga & ".txt"
            sNombreFile2 = "COMPRAS_ALICUOTAS_" & Year(dtPeriodo) & Format(Month(dtPeriodo), "00") & "_" & Cuit_Empresa_Carga & ".txt"
            sNombreFile3 = "DESPACHOS_ALICUOTAS_" & Year(dtPeriodo) & Format(Month(dtPeriodo), "00") & "_" & Cuit_Empresa_Carga & ".txt"
            sArchivoCompleto = sCarpeta & sNombreFile
            sArchivoCompleto2 = sCarpeta & sNombreFile2
            sArchivoCompleto3 = sCarpeta & sNombreFile3
            If ExisteArchivo(sArchivoCompleto) Then
                Kill sArchivoCompleto
            End If
            If ExisteArchivo(sArchivoCompleto2) Then
                Kill sArchivoCompleto2
            End If
            Open sArchivoCompleto For Output As #1
            Open sArchivoCompleto2 For Output As #2
            
            Open sArchivoCompleto3 For Output As #3
            
            .MoveFirst
            For i = 0 To .RecordCount - 1
            
                cFecha = "19000101"
                cTipo = "000"
                cPunto = "00000"
                cNro = "00000000000000000000"
                cNroDespacho = "                "
                cCodigo = "80"
                cIdentificador = "00000000000000000000"
                cApellidoProveedor = "----------SIN DATOS ----------"
                cTotal = "000000000000000"
                cNoGravado = "000000000000000"
                cExentas = "000000000000000"
                cPercepciones1 = "000000000000000"
                cPercepciones2 = "000000000000000"
                cIngresosBrutos = "000000000000000"
                cMunicipales = "000000000000000"
                cInternos = "000000000000000"
                cCodigoMoneda = "PES"
                cCambio = "0001000000"
                cCantidadAlicuotas = "1"
                cCodOperacion = "0"
                cCreditoFiscal = "000000000000000"
                cTributos = "000000000000000"
                cCuitEmisor = "00000000000"
                cDenominacionEmisor = "                              "
                cComision = "000000000000000"
            
                cNeto = "000000000000000"
                cAlicuota = "0000"
                cImpuestoLiquidado = "000000000000000"
                
                cFecha = Format(!Fecha, "YYYYMMDD")
                
                pLetra = Trim(sSinNull(!letra))
                If pLetra = "" Then
                    pLetra = Trim(obtenerDeSQL("select i.LETRA from prov p inner join ivas i on i.codigo=p.tipoiva where p.codigo=" & s2n(!CODPR)))
                End If
                pTipoDoc = Trim(!TIPODOC) & pLetra
                Select Case Trim(pTipoDoc)
                    Case "FACA": cTipo = "001"
                    Case "N/DA": cTipo = "002"
                    Case "N/CA": cTipo = "003"
                    Case "FACB": cTipo = "006"
                    Case "N/DB": cTipo = "007"
                    Case "N/CB": cTipo = "008"
                    Case "FACC": cTipo = "011"
                    Case "FACE": cTipo = "019"
                    Case "N/CE": cTipo = "021"
                End Select
                'cTipo = Format(Trim(!TIPODOC2), "000")
                
                cPunto = Format(!suc, "00000")
                If cPunto = "00000" Then cPunto = "00001"
                cNro = Format(!NroDoc, "00000000000000000000")
                cIdentificador = Format(Replace(!cuitprov, "-", ""), "00000000000000000000")
                                                
                cApellidoProveedor = Trim(!razonsocialprov)
                If Len(cApellidoProveedor) > 30 Then
                    cApellidoProveedor = CORTO(cApellidoProveedor, 0, Len(cApellidoProveedor) - 30)
                Else
                    While Len(cApellidoProveedor) < 30
                        cApellidoProveedor = Chr(32) & cApellidoProveedor
                    Wend
                End If
                
                cTotal = Format(Replace(s2n(!Total, 2, True), ",", ""), "000000000000000")
                cExentas = Format(Replace(s2n(!EXENTO, 2, True), ",", ""), "000000000000000")
                cNeto = Format(Replace(s2n(!Neto, 2, True), ",", ""), "000000000000000")
                If pLetra = "C" Or pLetra = "B" Then
                    cNeto = "000000000000000"
                    cNoGravado = "000000000000000"
                    cExentas = "000000000000000"
                End If
                cPercepciones1 = Format(Replace(s2n(!retIva, 2, True), ",", ""), "000000000000000")
                cPercepciones2 = Format(Replace(s2n(!retIva2, 2, True), ",", ""), "000000000000000")
                cIngresosBrutos = Format(Replace(s2n(!IIBB, 2, True), ",", ""), "000000000000000")
                cInternos = Format(Replace(s2n(!imp_int, 2, True), ",", ""), "000000000000000")
                
                a = 0
                cImpuestos = 0
                If pLetra = "B" Or pLetra = "C" Then
                    cAlicuota = "0003"
                    Print #2, cTipo & cPunto & cNro & cCodigo & cIdentificador & cNeto & cAlicuota & cImpuestoLiquidado
                Else
                    If pLetra <> "D" Then
                        If s2n(!IVA_21) > 0 Then
                            a = a + 1
                            cAlicuota = "0005"
                            cImpuestoLiquidado = Format(Replace(s2n(!IVA_21, 2, True), ",", ""), "000000000000000")
                            cImpuestos = cImpuestos + s2n(!IVA_21)
                            cNeto = s2n(!IVA_21 / 0.21, 2, True)
                            cNeto = Format(Replace(s2n(cNeto, 2, True), ",", ""), "000000000000000")
                            Print #2, cTipo & cPunto & cNro & cCodigo & cIdentificador & cNeto & cAlicuota & cImpuestoLiquidado
                        End If
                        If s2n(!IVA_27) > 0 Then
                            a = a + 1
                            cAlicuota = "0006"
                            cImpuestoLiquidado = Format(Replace(s2n(!IVA_27, 2, True), ",", ""), "000000000000000")
                            cImpuestos = cImpuestos + s2n(!IVA_27)
                            cNeto = s2n(!IVA_27 / 0.27, 2, True)
                            cNeto = Format(Replace(s2n(cNeto, 2, True), ",", ""), "000000000000000")
                            Print #2, cTipo & cPunto & cNro & cCodigo & cIdentificador & cNeto & cAlicuota & cImpuestoLiquidado
                        End If
                        If s2n(!iva_10) > 0 Then
                            a = a + 1
                            cAlicuota = "0004"
                            cImpuestoLiquidado = Format(Replace(s2n(!iva_10, 2, True), ",", ""), "000000000000000")
                            cImpuestos = cImpuestos + s2n(!iva_10)
                            cNeto = s2n(!iva_10 / 0.105, 2, True)
                            cNeto = Format(Replace(s2n(cNeto, 2, True), ",", ""), "000000000000000")
                            Print #2, cTipo & cPunto & cNro & cCodigo & cIdentificador & cNeto & cAlicuota & cImpuestoLiquidado
                        End If
                    End If
                End If
                cCantidadAlicuotas = a
                cCreditoFiscal = Format(Replace(s2n(s2n(!iva_10) + s2n(!IVA_21) + s2n(!IVA_27), 2, True), ",", ""), "000000000000000")
                
                If pLetra = "A" And s2n(a) = 0 Then
                    cCodOperacion = "E"
                    cAlicuota = "0003"
                    cCantidadAlicuotas = "1"
                    Print #2, cTipo & cPunto & cNro & cCodigo & cIdentificador & cNeto & cAlicuota & cImpuestoLiquidado
                End If
                
                If pLetra = "D" Then
                    cCantidadAlicuotas = 1
                    cPunto = "00000"
                    dDESPACHO = "0000000000000000"
                    dNETO = "0000"
                    dALICUOTA = "000000000000000"
                    dIMPUESTO = "000000000000000"
                    
                    dDESPACHO = sSinNull(!nrodespacho)
                    cNroDespacho = dDESPACHO
                    dNETO = Format(Replace(s2n(!Neto, 2, True), ",", ""), "000000000000000")
                    dALICUOTA = "0005"
                    dIMPUESTO = Format(Replace(s2n(!IVA_21, 2, True), ",", ""), "000000000000000")
                    
                    Print #3, dDESPACHO & dNETO & dALICUOTA & dIMPUESTO
                End If
                
                
                Select Case pTipoDoc
                    Case "N/DA": If a = 0 Then cCodOperacion = "A"
                    Case "N/DB": If a = 0 Then cCodOperacion = "A"
                    Case "FACE": cCodOperacion = "X"
                    Case "N/CE": cCodOperacion = "X"
                End Select
                cCreditoFiscal = Format(Replace(s2n(cImpuestos, 2, True), ",", ""), "000000000000000")
                
                'cCuitEmisor = Format(cIdentificador, "00000000000")
                'cDenominacionEmisor = cApellidoProveedor
                
                Print #1, cFecha & cTipo & cPunto & cNro & cNroDespacho & cCodigo & cIdentificador & cApellidoProveedor & cTotal & cNoGravado & cExentas & cPercepciones1 & cPercepciones2 & cIngresosBrutos & cMunicipales & cInternos & cCodigoMoneda & cCambio & cCantidadAlicuotas & cCodOperacion & cCreditoFiscal & cTributos & cCuitEmisor & cDenominacionEmisor & cComision
                
                .MoveNext
            Next
            
            
            

        End If
        
    End With
    
    
    Close #1
    Close #2
    Close #3
    
    Set rs1 = Nothing

    MsgBox "Archivo generado correctamente.", vbInformation, "Compras"

Exit Function
compras_err:
MsgBox "Error... " & Err.Number & " " & Err.Description

End Function

Private Function CitiComprasB()
On Error GoTo compras_err
Dim cFecha As String, cTipo As String, cPunto As String, cNro As String, cNroDespacho As String, cCodigo As String, cIdentificador As String, cApellidoProveedor As String, cTotal As String, cNoGravado As String, cExentas As String, cPercepciones1 As String, cPercepciones2 As String, cIngresosBrutos As String, cMunicipales As String, cInternos As String, cCodigoMoneda As String, cCambio As String, cCantidadAlicuotas As String, cCodOperacion As String, cCreditoFiscal As String, cTributos As String, cCuitEmisor As String, cDenominacionEmisor As String, cComision As String
Dim cPrimerDia As Date, cUltimoDia As Date, str As String, i As Long, sNombreFile As String, sCarpeta As String, sArchivoCompleto As String, sArchivoCompleto2 As String, pLetra As String, pTipoDoc As String, a As Integer
Dim cNeto As String, cAlicuota As String, cImpuestoLiquidado As String
Dim sFechaD As String, sFechaH As String
Dim rs1 As New ADODB.Recordset

    cPrimerDia = "01/" & Month(dtPeriodo) & "/" & Year(dtPeriodo)
    cUltimoDia = ultimoDiaDelMes(dtPeriodo)
    'sFechaD = Year(vPrimerDia) & Format(Month(vPrimerDia), "00") & Format(Day(vPrimerDia), "00")
    'sFechaH = Year(vUltimoDia) & Format(Month(vUltimoDia), "00") & Format(Day(vUltimoDia), "00")
    
    rs1.Open "select fecha,codpr,nrodoc,iva_21,iva_27,iva_9,iva_10,imp_int,razonsocialprov,cuitprov,suc,anoimp,iva_21,iva_27,iva_10,TIPODOC,TOTAL,NETO,EXENTO,(IBCAPITAL+IBPROVINCIA) AS IIBB ,letra,(PERCEPC+DER_EST) AS RETIVA,nogravado " & _
                " from TRANSCOM " & _
                " where fecha>=" & ssFecha(cPrimerDia) & " and fecha<=" & ssFecha(cUltimoDia) & " and tipodoc<>'RAC' and  tipodoc<>'APD' AND  tipodoc<>'APC' " & _
        "union " & _
            " select fecha,codpr,nrodoc,iva_21,iva_27,iva_9,iva_10,imp_int,razonsocialprov,cuitprov,suc,anoimp,iva_21,iva_27,iva_10,TIPODOC,TOTAL,NETO,EXENTO,(IBCAPITAL+IBPROVINCIA) AS IIBB ,letra,(PERCEPC+DER_EST) AS RETIVA,nogravado " & _
                " from compras " & _
                " where fecha>=" & ssFecha(cPrimerDia) & " and fecha<=" & ssFecha(cUltimoDia) & " and tipodoc<>'RAC' and  tipodoc<>'APD' AND  tipodoc<>'APC'  " & _
                " order by fecha,nrodoc", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly



        
    With rs1
        
        If rs1.EOF And rs1.BOF Then
            MsgBox "No se ha encontrado datos para este periodo.", vbInformation, "ATENCION"
            Exit Function
        Else
            
            sCarpeta = "C:\"
            sCarpeta = Trim(VentanaCarpeta("Carpeta Destino txt COMPRAS", sCarpeta))
            If CORTO(sCarpeta, Len(sCarpeta) - 1, 0) <> "\" Then sCarpeta = sCarpeta & "\"
            sNombreFile = "COMPRAS_" & Year(dtPeriodo) & Format(Month(dtPeriodo), "00") & ".txt"
            sNombreFile2 = "COMPRAS_ALICUOTAS_" & Year(dtPeriodo) & Format(Month(dtPeriodo), "00") & ".txt"
            sArchivoCompleto = sCarpeta & sNombreFile
            sArchivoCompleto2 = sCarpeta & sNombreFile2
            If ExisteArchivo(sArchivoCompleto) Then
                Kill sArchivoCompleto
            End If
            If ExisteArchivo(sArchivoCompleto2) Then
                Kill sArchivoCompleto2
            End If
            Open sArchivoCompleto For Output As #1
            Open sArchivoCompleto2 For Output As #2
            .MoveFirst
            For i = 0 To .RecordCount - 1
            
                cFecha = "19000101"
                cTipo = "000"
                cPunto = "00000"
                cNro = "00000000000000000000"
                cNroDespacho = "                "
                cCodigo = "80"
                cIdentificador = "00000000000000000000"
                cApellidoProveedor = "----------SIN DATOS ----------"
                cTotal = "000000000000000"
                cNoGravado = "000000000000000"
                cExentas = "000000000000000"
                cPercepciones1 = "000000000000000"
                cPercepciones2 = "000000000000000"
                cIngresosBrutos = "000000000000000"
                cMunicipales = "000000000000000"
                cInternos = "000000000000000"
                cCodigoMoneda = "PES"
                cCambio = "0001000000"
                cCantidadAlicuotas = "1"
                cCodOperacion = "0"
                cCreditoFiscal = "000000000000000"
                cTributos = "000000000000000"
                cCuitEmisor = "00000000000"
                cDenominacionEmisor = "                              "
                cComision = "000000000000000"
            
                cNeto = "000000000000000"
                cAlicuota = "0000"
                cImpuestoLiquidado = "000000000000000"
                
                cFecha = Format(!Fecha, "YYYYMMDD")
                
                
                pLetra = Trim(sSinNull(!letra))
                If pLetra = "" Then
                    pLetra = Trim(obtenerDeSQL("select i.LETRA from prov p inner join ivas i on i.codigo=p.tipoiva where p.codigo=" & s2n(!CODPR)))
                End If
                pTipoDoc = Trim(!TIPODOC) & pLetra
                
                Select Case Trim(pTipoDoc)
                    Case "FACA": cTipo = "001"
                    Case "N/DA": cTipo = "002"
                    Case "N/CA": cTipo = "003"
                    Case "FACB": cTipo = "006"
                    Case "N/DB": cTipo = "007"
                    Case "N/CB": cTipo = "008"
                    Case "FACC": cTipo = "011"
                    Case "FACE": cTipo = "019"
                    Case "N/CE": cTipo = "021"
                End Select
                
                If pLetra = "E" Then
                    GoTo siguiente
                End If
                
                cPunto = Format(!suc, "00000")
                If cPunto = "00000" Then cPunto = "00001"
                
                'If !NroDoc = 9963 Then Stop
                
                cNro = Format(!NroDoc, "00000000000000000000")
                cIdentificador = Format(Replace(!cuitprov, "-", ""), "00000000000000000000")
                
                
                cApellidoProveedor = Trim(!razonsocialprov)
                If Len(cApellidoProveedor) > 30 Then
                    cApellidoProveedor = CORTO(cApellidoProveedor, 0, Len(cApellidoProveedor) - 30)
                Else
                    While Len(cApellidoProveedor) < 30
                        cApellidoProveedor = Chr(32) & cApellidoProveedor
                    Wend
                End If
                
                cTotal = Format(Replace(s2n(!Total, 2, True), ",", ""), "000000000000000")
                cExentas = Format(Replace(s2n(!EXENTO, 2, True), ",", ""), "000000000000000")
                cNoGravado = Format(Replace(s2n(!nogravado, 2, True), ",", ""), "000000000000000")
                cNeto = Format(Replace(s2n(!Neto, 2, True), ",", ""), "000000000000000")
                If pLetra = "C" Or pLetra = "B" Then
                    cNeto = "000000000000000"
                    cNoGravado = "000000000000000"
                    cExentas = "000000000000000"
                End If
                cPercepciones1 = Format(Replace(s2n(!retIva, 2, True), ",", ""), "000000000000000")
                cIngresosBrutos = Format(Replace(s2n(!IIBB, 2, True), ",", ""), "000000000000000")
                cInternos = Format(Replace(s2n(!imp_int, 2, True), ",", ""), "000000000000000")
                
                a = 0
                If pLetra = "B" Or pLetra = "C" Then
                    cAlicuota = "0003"
                    Print #2, cTipo & cPunto & cNro & cCodigo & cIdentificador & cNeto & cAlicuota & cImpuestoLiquidado
                Else
                    If s2n(!IVA_21) > 0 Then
                        a = a + 1
                        cAlicuota = "0005"
                        cImpuestoLiquidado = Format(Replace(s2n(!IVA_21, 2, True), ",", ""), "000000000000000")
                        cNeto = s2n(!IVA_21 / 0.21, 2, True)
                        cNeto = Format(Replace(s2n(cNeto, 2, True), ",", ""), "000000000000000")
                        Print #2, cTipo & cPunto & cNro & cCodigo & cIdentificador & cNeto & cAlicuota & cImpuestoLiquidado
                    End If
                    If s2n(!IVA_27) > 0 Then
                        a = a + 1
                        cAlicuota = "0006"
                        cImpuestoLiquidado = Format(Replace(s2n(!IVA_27, 2, True), ",", ""), "000000000000000")
                        cNeto = s2n(!IVA_27 / 0.27, 2, True)
                        cNeto = Format(Replace(s2n(cNeto, 2, True), ",", ""), "000000000000000")
                        Print #2, cTipo & cPunto & cNro & cCodigo & cIdentificador & cNeto & cAlicuota & cImpuestoLiquidado
                    End If
                    If s2n(!iva_10) > 0 Then
                        a = a + 1
                        cAlicuota = "0004"
                        cImpuestoLiquidado = Format(Replace(s2n(!iva_10, 2, True), ",", ""), "000000000000000")
                        cNeto = s2n(!iva_10 / 0.105, 2, True)
                        cNeto = Format(Replace(s2n(cNeto, 2, True), ",", ""), "000000000000000")
                        Print #2, cTipo & cPunto & cNro & cCodigo & cIdentificador & cNeto & cAlicuota & cImpuestoLiquidado
                    End If
                End If
                cCantidadAlicuotas = a
                
                
                If pLetra = "A" And s2n(a) = 0 Then
                    cCodOperacion = "E"
                    cAlicuota = "0003"
                    cCantidadAlicuotas = "1"
                    Print #2, cTipo & cPunto & cNro & cCodigo & cIdentificador & cNeto & cAlicuota & cImpuestoLiquidado
                End If
                
                Select Case pTipoDoc
                    Case "N/DA": If a = 0 Then cCodOperacion = "A"
                    Case "N/DB": If a = 0 Then cCodOperacion = "A"
                    Case "FACE": cCodOperacion = "X"
                    Case "N/CE": cCodOperacion = "X"
                End Select
                
                'cCuitEmisor = Format(cIdentificador, "00000000000")
                'cDenominacionEmisor = cApellidoProveedor
                
                Print #1, cFecha & cTipo & cPunto & cNro & cNroDespacho & cCodigo & cIdentificador & cApellidoProveedor & cTotal & cNoGravado & cExentas & cPercepciones1 & cPercepciones2 & cIngresosBrutos & cMunicipales & cInternos & cCodigoMoneda & cCambio & cCantidadAlicuotas & cCodOperacion & cCreditoFiscal & cTributos & cCuitEmisor & cDenominacionEmisor & cComision
                
                
siguiente:
                .MoveNext
            Next
        End If
        
    End With
    
    Close #1
    Close #2
    Set rs1 = Nothing

    MsgBox "Archivo generado correctamente.", vbInformation, "Compras"

Exit Function
compras_err:
MsgBox "Error... " & Err.Number & " " & Err.Description

End Function


Private Function CitiVentasB()
On Error GoTo ventas_err
Dim vFecha As String, vTipo As String, vPunto As String, vNro As String, vNroHasta As String, vCodigo As String, vIdentificador As String, vApellidoCliente As String, vTotal As String, vNoGravado As String, vPercepcionNN As String, vExentas As String, vPercepciones As String, vIngresosBrutos As String, vMunicipales As String, vInternos As String, vCodigoMoneda As String, vCambio As String, vCantidadAlicuotas As String, vCodOperacion As String, vTributos As String, vVencimiento As String
Dim vPrimerDia As Date, vUltimoDia As Date, str As String, i As Long, sNombreFile2 As String, sNombreFile As String, sCarpeta As String, sArchivoCompleto As String, sArchivoCompleto2 As String, a As Integer
Dim vNeto As String, vAlicuota As String, vImpuestoLiquidado As String, nIva21 As Double, nIva10 As Double
Dim sFechaD As String, sFechaH As String
Dim rs1 As New ADODB.Recordset

    vPrimerDia = "01/" & Month(dtPeriodo) & "/" & Year(dtPeriodo)
    vUltimoDia = ultimoDiaDelMes(dtPeriodo)
    'sFechaD = Year(vPrimerDia) & Format(Month(vPrimerDia), "00") & Format(Day(vPrimerDia), "00")
    'sFechaH = Year(vUltimoDia) & Format(Month(vUltimoDia), "00") & Format(Day(vUltimoDia), "00")
    'str = "select * from facturaventa where fecha>=" & ssTexto(sFechaD) & " and fecha<=" & ssTexto(sFechaH) & " and tipodoc in ('NDB','NDA','FAE','FAA','FAB','NCA','NCB') order by fecha,tipodoc,nrofactura"
    str = "select * from facturaventa where fecha>=" & ssFecha(vPrimerDia) & " and fecha<=" & ssFecha(vUltimoDia) & " and tipodoc in ('NDB','NDA','FAE','FAA','FAB','NCA','NCB') order by fecha,tipodoc,nrofactura"
    rs1.Open str, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly

    With rs1
        If .EOF And .BOF Then
            MsgBox "No se ha encontrado datos para este periodo.", vbInformation, "ATENCION"
            Exit Function
        Else
            sCarpeta = "C:\"
            sCarpeta = Trim(VentanaCarpeta("Carpeta Destino txt VENTAS", sCarpeta))
            If CORTO(sCarpeta, Len(sCarpeta) - 1, 0) <> "\" Then sCarpeta = sCarpeta & "\"
            sNombreFile = "VENTAS_" & Year(dtPeriodo) & Format(Month(dtPeriodo), "00") & ".txt"
            sNombreFile2 = "VENTAS_ALICUOTAS_" & Year(dtPeriodo) & Format(Month(dtPeriodo), "00") & ".txt"
            sArchivoCompleto = sCarpeta & sNombreFile
            sArchivoCompleto2 = sCarpeta & sNombreFile2
            If ExisteArchivo(sArchivoCompleto) Then
                Kill sArchivoCompleto
            End If
            If ExisteArchivo(sArchivoCompleto2) Then
                Kill sArchivoCompleto2
            End If
            Open sArchivoCompleto For Output As #1
            Open sArchivoCompleto2 For Output As #2
            
            .MoveFirst
            For i = 0 To .RecordCount - 1
                vFecha = "19000101"
                vTipo = "000"
                vPunto = "00000"
                vNro = "00000000000000000000"
                vNroHasta = "00000000000000000000"
                vCodigo = "00"
                vIdentificador = "00000000000000000000"
                vApellidoCliente = "----------SIN DATOS ----------"
                vTotal = "000000000000000"
                vNoGravado = "000000000000000"
                vPercepcionNN = "000000000000000"
                vExentas = "000000000000000"
                vPercepciones = "000000000000000"
                vIngresosBrutos = "000000000000000"
                vMunicipales = "000000000000000"
                vInternos = "000000000000000"
                vCodigoMoneda = "PES"
                vCambio = "0001000000"
                vCantidadAlicuotas = "1"
                vCodOperacion = "0"
                vTributos = "000000000000000"
                vVencimiento = "00000000"
            
                vNeto = "000000000000000"
                vAlicuota = "0000"
                vImpuestoLiquidado = "000000000000000"
                
                
                vFecha = Format(!Fecha, "YYYYMMDD")
                
                
                Select Case Trim(!TIPODOC)
                    Case "FAA": vTipo = "001"
                    Case "FAB": vTipo = "006"
                    Case "NCA": vTipo = "003"
                    Case "NCB": vTipo = "008"
                    Case "NDA": vTipo = "002"
                    Case "NDB": vTipo = "007"
                    Case "FAE": vTipo = "019"
                    Case "NCE": vTipo = "021"
                End Select
                
                If vTipo <> "019" And vTipo <> "021" Then
                    vVencimiento = Format(!Vencimiento, "YYYYMMDD")
                End If
                
                'If !NroFactura = 1506 Then Stop
                'If !NroFactura = 215 Then Stop
                
                vPunto = Format(!PuntoVenta, "00000")
                vNro = Format(!NroFactura, "00000000000000000000")
                vNroHasta = vNro
                
                vIdentificador = Replace(!CUIT, "-", "")
                vIdentificador = s2n(vIdentificador)
                                    
                Select Case !tipoiva
                    Case 2, 4, 7: vCodigo = 80 'cuit
                    Case 1: vCodigo = IIf(Len(vIdentificador) > 8, 86, 96) '96=dni,86=cuil
                    Case Else: vCodigo = 86
                End Select
                
                vIdentificador = Format(vIdentificador, "00000000000000000000")
                
                vApellidoCliente = Trim(!RAZONSOCIAL)
                If Len(vApellidoCliente) > 30 Then
                    vApellidoCliente = CORTO(vApellidoCliente, 0, Len(vApellidoCliente) - 30)
                Else
                    While Len(vApellidoCliente) < 30
                        vApellidoCliente = Chr(32) & vApellidoCliente
                    Wend
                End If
                
                vCantidadAlicuotas = obtenerDeSQL("select count(distinct(_iva)) from facturaventadetalle where codigofactura=" & !codigo)
                
                If !ND_xChequeRechazado Then
                    vCodOperacion = "A"
                ElseIf CORTO(!TIPODOC, 2, 0) = "E" Then
                    vCodOperacion = "X"
                End If
                
                
                vTotal = Format(Replace(s2n(!Total, 2, True), ",", ""), "000000000000000")
                vNoGravado = Format(Replace(s2n(!NoGrav, 2, True), ",", ""), "000000000000000")
                If vTipo = "019" Or vTipo = "021" Then
                    vExentas = vTotal 'Format(!NoGrav, "000000000000000")
                End If
                vIngresosBrutos = Format(Replace(s2n(!IIBB, 2, True), ",", ""), "000000000000000")
                
                
                'vImpuestoLiquidado = Format(Replace(s2n(!Iva, 2, True), ",", ""), "000000000000000")
                
                
               
                a = 0
                If vTipo = "019" Or vTipo = "021" Then
                    vAlicuota = "0003"
                    Print #2, vTipo & vPunto & vNro & vNeto & vAlicuota & vImpuestoLiquidado
'                ElseIf vTipo = "006" Or vTipo = "007" Or vTipo = "008" Then
'                    nIva21 = nSinNull(obtenerDeSQL("select (sum(preciototal)/1.21) as preciototal from facturaventadetalle where _iva=21 and codigofactura=" & !codigo))
'                    If s2n(nIva21) > 0 Then
'                        a = a + 1
'                        vAlicuota = "0005"
'                        If !ND_xChequeRechazado Then vAlicuota = "0003"
'                        vNeto = Format(Replace(s2n(nIva21, 2, True), ",", ""), "000000000000000")
'                        vImpuestoLiquidado = Format(Replace(s2n(nIva21 * 0.21, 2, True), ",", ""), "000000000000000")
'                        Print #2, vTipo & vPunto & vNro & vNeto & vAlicuota & vImpuestoLiquidado
'                    End If
'                    nIva10 = nSinNull(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where _iva=10.5 and codigofactura=" & !codigo))
'                    If s2n(nIva10) > 0 Then
'                        a = a + 1
'                        vAlicuota = "0004"
'                        If !ND_xChequeRechazado Then vAlicuota = "0003"
'                        vNeto = Format(Replace(s2n(nIva10, 2, True), ",", ""), "000000000000000")
'                        vImpuestoLiquidado = Format(Replace(s2n(nIva10 * 0.105, 2, True), ",", ""), "000000000000000")
'                        Print #2, vTipo & vPunto & vNro & vNeto & vAlicuota & vImpuestoLiquidado
'                    End If
                Else
                    If vTipo = "006" Or vTipo = "007" Or vTipo = "008" Then
                        nIva21 = nSinNull(obtenerDeSQL("select (sum(preciototal)/1.21) as preciototal from facturaventadetalle where _iva=21 and codigofactura=" & !codigo))
                    Else
                        nIva21 = nSinNull(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where _iva=21 and codigofactura=" & !codigo))
                    End If
                    If s2n(nIva21) > 0 Then
                        a = a + 1
                        vAlicuota = "0005"
                        If !ND_xChequeRechazado Then vAlicuota = "0003"
                        vNeto = Format(Replace(s2n(nIva21, 2, True), ",", ""), "000000000000000")
                        vImpuestoLiquidado = Format(Replace(s2n(nIva21 * 0.21, 2, True), ",", ""), "000000000000000")
                        Print #2, vTipo & vPunto & vNro & vNeto & vAlicuota & vImpuestoLiquidado
                    End If
                    
                    If vTipo = "006" Or vTipo = "007" Or vTipo = "008" Then
                        nIva10 = nSinNull(obtenerDeSQL("select (sum(preciototal)/1.105) as preciototal  from facturaventadetalle where _iva=10.5 and codigofactura=" & !codigo))
                    Else
                        nIva10 = nSinNull(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where _iva=10.5 and codigofactura=" & !codigo))
                    End If
                    If s2n(nIva10) > 0 Then
                        a = a + 1
                        vAlicuota = "0004"
                        If !ND_xChequeRechazado Then vAlicuota = "0003"
                        vNeto = Format(Replace(s2n(nIva10, 2, True), ",", ""), "000000000000000")
                        vImpuestoLiquidado = Format(Replace(s2n(nIva10 * 0.105, 2, True), ",", ""), "000000000000000")
                        Print #2, vTipo & vPunto & vNro & vNeto & vAlicuota & vImpuestoLiquidado
                    End If
                End If
                vCantidadAlicuotas = a
                
                
                
                
                
                Print #1, vFecha & vTipo & vPunto & vNro & vNroHasta & vCodigo & vIdentificador & vApellidoCliente & vTotal & vNoGravado & vPercepcionNN & vExentas & vPercepciones & vIngresosBrutos & vMunicipales & vInternos & vCodigoMoneda & vCambio & vCantidadAlicuotas & vCodOperacion & vTributos & vVencimiento
                
            
                .MoveNext
            Next
        End If
    End With

    
    Close #1
    Close #2
    Set rs1 = Nothing

    MsgBox "Archivo generado correctamente.", vbInformation, "Ventas"
                
            
Exit Function
ventas_err:
MsgBox "Error... " & Err.Number & " " & Err.Description
End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FacturaElectronica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const TIPODOC_FAC_PROVEEDOR = "FAC"
Const TIPODOC_FAC_PROVGASTO = "FCG"
Const TIPODOC_FAC_BANCOGASTO = "BG"
Const TIPODOC_FAC_BOLETA = "BOL"

Const TipoDoc_FACTURA_A = "FAA"
Const TipoDoc_FACTURA_B = "FAB"
Const TipoDoc_FACTURA_C = "FAC"
Const TipoDoc_FACTURA_E = "FAE"
Const TipoDoc_FACTURA_M = "FAM"
Const TipoDoc_FACTURA_CREDITO_A = "FEA"
Const TipoDoc_FACTURA_CREDITO_B = "FEB"
Const TipoDoc_FACTURA_CREDITO_C = "FEC"

Const TipoDoc_NCREDITO_A = "NCA"
Const TipoDoc_NCREDITO_B = "NCB"
Const TipoDoc_NCREDITO_C = "NCC"
Const TipoDoc_NCREDITO_E = "NCE"
Const TipoDoc_NCREDITO_M = "NCM"

Const TipoDoc_CREDITO_ELECTRONICO_A = "CEA"
Const TipoDoc_CREDITO_ELECTRONICO_B = "CEB"
Const TipoDoc_CREDITO_ELECTRONICO_C = "CEC"

Const TipoDoc_NDEBITO_A = "NDA"
Const TipoDoc_NDEBITO_B = "NDB"
Const TipoDoc_NDEBITO_C = "NDC"
Const TipoDoc_NDEBITO_E = "NDE"
Const TipoDoc_NDEBITO_M = "NDM"
Const TipoDoc_DEBITO_ELECTRONICO_A = "DEA"
Const TipoDoc_DEBITO_ELECTRONICO_B = "DEB"
Const TipoDoc_DEBITO_ELECTRONICO_C = "DEC"




Const TipoDoc_TICKET = "TIC"

Const TipoDoc_RECIBO = "RAA"
Const TipoDoc_AJ_CREDITO = "ACC"
Const TipoDoc_AJ_DEBITO = "ACD"

Public Enum wsCodigosDeFactura
    FACTURA_A_CODIGO = 1
    FACTURA_B_CODIGO = 6
    FACTURA_C_CODIGO = 11
    FACTURA_E_CODIGO = 19
    FACTURA_M_CODIGO = 54
    FACTURA_CREDITO_A_CODIGO = 201
    FACTURA_CREDITO_B_CODIGO = 206
    FACTURA_CREDITO_C_CODIGO = 211
    NOTA_CREDITO_A_CODIGO = 3
    NOTA_CREDITO_B_CODIGO = 8
    NOTA_CREDITO_C_CODIGO = 13
    NOTA_CREDITO_E_CODIGO = 21
    NOTA_CREDITO_M_CODIGO = 53
    NOTA_CRED_ELEC_A_CODIGO = 203
    NOTA_CRED_ELEC_B_CODIGO = 208
    NOTA_CRED_ELEC_C_CODIGO = 213
    NOTA_DEBITO_A_CODIGO = 2
    NOTA_DEBITO_B_CODIGO = 7
    NOTA_DEBITO_C_CODIGO = 12
    NOTA_DEBITO_E_CODIGO = 20
    NOTA_DEBITO_M_CODIGO = 52
    NOTA_DEBI_ELEC_A_CODIGO = 202
    NOTA_DEBI_ELEC_B_CODIGO = 207
    NOTA_DEBI_ELEC_C_CODIGO = 212
End Enum

Public Enum wsSiglasComprobante
    FACTURA_TIPO_A
    FACTURA_TIPO_B
    FACTURA_TIPO_C
    FACTURA_TIPO_E
    FACTURA_TIPO_M
    FACTURA_CREDITO_TIPO_A
    FACTURA_CREDITO_TIPO_B
    FACTURA_CREDITO_TIPO_C
    NOTA_CREDITO_TIPO_A
    NOTA_CREDITO_TIPO_B
    NOTA_CREDITO_TIPO_C
    NOTA_CREDITO_TIPO_E
    NOTA_CREDITO_TIPO_M
    CREDITO_ELECTRONICA_TIPO_A
    CREDITO_ELECTRONICA_TIPO_B
    CREDITO_ELECTRONICA_TIPO_C
    NOTA_DEBITO_TIPO_A
    NOTA_DEBITO_TIPO_B
    NOTA_DEBITO_TIPO_C
    NOTA_DEBITO_TIPO_E
    NOTA_DEBITO_TIPO_M
    DEBITO_ELECTRONICA_TIPO_A
    DEBITO_ELECTRONICA_TIPO_B
    DEBITO_ELECTRONICA_TIPO_C
End Enum

Public Enum wsTipoComprobante
    FACTURA_A = 1
    Nota_de_Debito_A = 2
    Nota_de_Credito_A = 3
    FACTURA_B = 6
    Nota_de_Debito_B = 7
    Nota_de_Credito_B = 8
    Recibos_A = 4
    Notas_de_Venta_al_contado_A = 5
    Recibos_B = 9
    Notas_de_Venta_al_contado_B = 10
    liquidacion_A = 63
    liquidacion_B = 64
    Cbtes_A_RG_1415 = 34
    Cbtes_B_RG_1415 = 35
    Otro_Cbtes_A_RG_1415 = 39
    Otro_Cbtes_B_RG_1415 = 40
    Cta_de_Vta_y_Liquido_prod_A = 60
    Cta_de_Vta_y_Liquido_prod_B = 61
    FACTURA_C = 11
    Nota_de_Debito_C = 12
    Nota_de_Credito_C = 13
    Recibo_C = 15
    Comprobante_Compra_de_Bienes_Usados_CF = 49
    Factura_M = 51
    Nota_de_Debito_M = 52
    Nota_de_Credito_M = 53
    Recibo_M = 54
    FACTURA_E = 19
    Nota_de_Debito_E = 20
    Nota_de_Credito_E = 21
    Factura_de_Credito_electronica_A = 201
    Nota_de_Debito_electronica_A = 202
    Nota_de_Credito_electronica_A = 203
    Factura_de_Credito_electronica_B = 206
    Nota_de_Debito_electronica_B = 207
    Nota_de_Credito_electronica_B = 208
    Factura_de_Credito_electronica_C = 211
    Nota_de_Debito_electronica_C = 212
    Nota_de_Credito_electronica_C = 213
End Enum

Public Function LetraComprobante(docTipo As wsSiglasComprobante) As String

If docTipo = FACTURA_TIPO_A Then LetraComprobante = TipoDoc_FACTURA_A
If docTipo = FACTURA_TIPO_B Then LetraComprobante = TipoDoc_FACTURA_B
If docTipo = FACTURA_TIPO_C Then LetraComprobante = TipoDoc_FACTURA_C
If docTipo = FACTURA_TIPO_E Then LetraComprobante = TipoDoc_FACTURA_E
If docTipo = FACTURA_TIPO_M Then LetraComprobante = TipoDoc_FACTURA_M
    
If docTipo = FACTURA_CREDITO_TIPO_A Then LetraComprobante = TipoDoc_FACTURA_CREDITO_A
If docTipo = FACTURA_CREDITO_TIPO_B Then LetraComprobante = TipoDoc_FACTURA_CREDITO_B
If docTipo = FACTURA_CREDITO_TIPO_C Then LetraComprobante = TipoDoc_FACTURA_CREDITO_C

If docTipo = NOTA_CREDITO_TIPO_A Then LetraComprobante = TipoDoc_NCREDITO_A
If docTipo = NOTA_CREDITO_TIPO_B Then LetraComprobante = TipoDoc_NCREDITO_B
If docTipo = NOTA_CREDITO_TIPO_C Then LetraComprobante = TipoDoc_NCREDITO_C
If docTipo = NOTA_CREDITO_TIPO_E Then LetraComprobante = TipoDoc_NCREDITO_E
If docTipo = NOTA_CREDITO_TIPO_M Then LetraComprobante = TipoDoc_NCREDITO_M

If docTipo = CREDITO_ELECTRONICA_TIPO_A Then LetraComprobante = TipoDoc_CREDITO_ELECTRONICO_A
If docTipo = CREDITO_ELECTRONICA_TIPO_B Then LetraComprobante = TipoDoc_CREDITO_ELECTRONICO_B
If docTipo = CREDITO_ELECTRONICA_TIPO_C Then LetraComprobante = TipoDoc_CREDITO_ELECTRONICO_C

If docTipo = NOTA_DEBITO_TIPO_A Then LetraComprobante = TipoDoc_NDEBITO_A
If docTipo = NOTA_DEBITO_TIPO_B Then LetraComprobante = TipoDoc_NDEBITO_B
If docTipo = NOTA_DEBITO_TIPO_C Then LetraComprobante = TipoDoc_NDEBITO_C
If docTipo = NOTA_DEBITO_TIPO_E Then LetraComprobante = TipoDoc_NDEBITO_E
If docTipo = NOTA_DEBITO_TIPO_M Then LetraComprobante = TipoDoc_NDEBITO_M

If docTipo = DEBITO_ELECTRONICA_TIPO_A Then LetraComprobante = TipoDoc_DEBITO_ELECTRONICO_A
If docTipo = DEBITO_ELECTRONICA_TIPO_B Then LetraComprobante = TipoDoc_DEBITO_ELECTRONICO_B
If docTipo = DEBITO_ELECTRONICA_TIPO_C Then LetraComprobante = TipoDoc_DEBITO_ELECTRONICO_C

End Function

Public Function CodigoComprobante(docTipo As wsSiglasComprobante) As Integer

If docTipo = FACTURA_TIPO_A Then CodigoComprobante = FACTURA_A_CODIGO
If docTipo = FACTURA_TIPO_B Then CodigoComprobante = FACTURA_B_CODIGO
If docTipo = FACTURA_TIPO_C Then CodigoComprobante = FACTURA_C_CODIGO
If docTipo = FACTURA_TIPO_E Then CodigoComprobante = FACTURA_E_CODIGO
If docTipo = FACTURA_TIPO_M Then CodigoComprobante = FACTURA_M_CODIGO
    
If docTipo = FACTURA_CREDITO_TIPO_A Then CodigoComprobante = FACTURA_CREDITO_A_CODIGO
If docTipo = FACTURA_CREDITO_TIPO_B Then CodigoComprobante = FACTURA_CREDITO_B_CODIGO
If docTipo = FACTURA_CREDITO_TIPO_C Then CodigoComprobante = FACTURA_CREDITO_C_CODIGO

If docTipo = NOTA_CREDITO_TIPO_A Then CodigoComprobante = NOTA_CREDITO_A_CODIGO
If docTipo = NOTA_CREDITO_TIPO_B Then CodigoComprobante = NOTA_CREDITO_B_CODIGO
If docTipo = NOTA_CREDITO_TIPO_C Then CodigoComprobante = NOTA_CREDITO_C_CODIGO
If docTipo = NOTA_CREDITO_TIPO_E Then CodigoComprobante = NOTA_CREDITO_E_CODIGO
If docTipo = NOTA_CREDITO_TIPO_M Then CodigoComprobante = NOTA_CREDITO_M_CODIGO

If docTipo = CREDITO_ELECTRONICA_TIPO_A Then CodigoComprobante = NOTA_CRED_ELEC_A_CODIGO
If docTipo = CREDITO_ELECTRONICA_TIPO_B Then CodigoComprobante = NOTA_CRED_ELEC_B_CODIGO
If docTipo = CREDITO_ELECTRONICA_TIPO_C Then CodigoComprobante = NOTA_CRED_ELEC_C_CODIGO

If docTipo = NOTA_DEBITO_TIPO_A Then CodigoComprobante = NOTA_DEBITO_A_CODIGO
If docTipo = NOTA_DEBITO_TIPO_B Then CodigoComprobante = NOTA_DEBITO_B_CODIGO
If docTipo = NOTA_DEBITO_TIPO_C Then CodigoComprobante = NOTA_DEBITO_C_CODIGO
If docTipo = NOTA_DEBITO_TIPO_E Then CodigoComprobante = NOTA_DEBITO_E_CODIGO
If docTipo = NOTA_DEBITO_TIPO_M Then CodigoComprobante = NOTA_DEBITO_M_CODIGO

If docTipo = DEBITO_ELECTRONICA_TIPO_A Then CodigoComprobante = NOTA_DEBI_ELEC_A_CODIGO
If docTipo = DEBITO_ELECTRONICA_TIPO_B Then CodigoComprobante = NOTA_DEBI_ELEC_B_CODIGO
If docTipo = DEBITO_ELECTRONICA_TIPO_C Then CodigoComprobante = NOTA_DEBI_ELEC_C_CODIGO

End Function

Public Function EmisionFacturaElectronica(IDFactura As Long, CPERMISO As String) As Boolean

Dim CAE As New WSAFIPFE.Factura, EstaEnLinea As Boolean, TicketGuardado As String
Dim esFacturaDeCredito As Boolean
Dim DatosFactura
Dim bResultado As Boolean, bresultado2 As Boolean, sModoFE As Long, sEMPRESA As String
Dim CuitEmpresa As String, Certificado As String, Licencia As String, rsFactura As New ADODB.Recordset, Identificador As String, IdentificadorF As String, COMPROBANTE2 As TipoComprobante, COMPROBANTE As wsTipoComprobante
Dim FFECHA As Date, fneto As Double, ftotal As Double, FCUIT As String, fTipo As String, fCAE As String, fPrDia As Date, fPuntoVenta As Long, fDetalle As New ADODB.Recordset
Dim fCBU As String, fALIAS As String
Dim bCuit As String, bCodFactura As String, bPUNTOVENTA As String, bCAE As String, bFechaCAE As String, bBARRA As String
Dim i As Long, fexNroTMP
Dim zzz As Double, ttt As Double
Dim sumItem As Double, iii As Double, ptt As Double, PIVA As Double
Dim Tipo3 As Double, Tipo4 As Double, Tipo5 As Double
Dim Tipo3base As Double, Tipo4base As Double, Tipo5base As Double

EmisionFacturaElectronica = False
sEMPRESA = obtenerDeSQL("select nombrearchivo from datosempresa")


If IDFactura = 0 Then
    MsgBox "No se puede seguir con el proceso, ID de factura invalido.", vbCritical
    Exit Function
End If

CuitEmpresa = Trim(Replace(obtenerDeSQL("select cuitempresa from datosempresa where idempresa=" & gEMPR_idEmpresa), "-", ""))
If CuitEmpresa = "" Then
    MsgBox "No se puede seguir con el proceso, el CUIT de la empresa no esta establecido.", vbCritical
    Exit Function
End If
If Len(CuitEmpresa) < 11 Then
    MsgBox "No se puede seguir con el proceso, el CUIT de la empresa no es valido.", vbCritical
    Exit Function
End If

Certificado = App.Path & "\" & sEMPRESA & ".pfx"
If ExisteArchivo(Certificado) Then
Else
    MsgBox "No se puede seguir con el proceso, el certificado no existe.", vbCritical
    Exit Function
End If

Licencia = App.Path & "\" & sEMPRESA & ".lic"
'Licencia = ""
If ExisteArchivo(Licencia) Then
Else
    MsgBox "No se puede seguir con el proceso, la licencia no existe.", vbCritical
    Exit Function
End If

fCAE = Trim(sSinNull(obtenerDeSQL("Select cae from facturaventa where codigo=" & IDFactura)))
If fCAE = "" Then
Else
    MsgBox "No se puede seguir con el proceso, la factura ya tiene CAE.", vbCritical
    Exit Function
End If

'With CAE
bPUNTOVENTA = Trim(sSinNull(obtenerDeSQL("Select puntoventa from facturaventa where codigo=" & IDFactura)))


CAE.tls = 12
sModoFE = obtenerDeSQL("select licenciacae from datosempresa where idempresa=" & gEMPR_idEmpresa)
If sModoFE = 0 Then
    bResultado = CAE.iniciar(modoFiscal_Test, CuitEmpresa, Certificado, Licencia)
Else
    bResultado = CAE.iniciar(modoFiscal_Fiscal, CuitEmpresa, Certificado, Licencia)
End If





    If bResultado Then
        rsFactura.Open "Select * from facturaventa where codigo=" & IDFactura, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
        
        If CORTO(rsFactura!TIPODOC, 0, 2) = "N" Then
            If CORTO(rsFactura!TIPODOC, 1, 1) = "C" Then
                If CORTO(rsFactura!TIPODOC, 2, 0) = "A" Then
                    COMPROBANTE = Nota_de_Credito_A
                ElseIf CORTO(rsFactura!TIPODOC, 2, 0) = "B" Then
                    COMPROBANTE = Nota_de_Credito_B
                ElseIf CORTO(rsFactura!TIPODOC, 2, 0) = "C" Then
                    COMPROBANTE = Nota_de_Credito_C
                ElseIf CORTO(rsFactura!TIPODOC, 2, 0) = "E" Then
                    COMPROBANTE = Nota_de_Credito_E
                End If
            ElseIf CORTO(rsFactura!TIPODOC, 1, 1) = "D" Then
                If CORTO(rsFactura!TIPODOC, 2, 0) = "A" Then
                    COMPROBANTE = Nota_de_Debito_A
                ElseIf CORTO(rsFactura!TIPODOC, 2, 0) = "B" Then
                    COMPROBANTE = Nota_de_Debito_B
                ElseIf CORTO(rsFactura!TIPODOC, 2, 0) = "C" Then
                    COMPROBANTE = Nota_de_Debito_C
                ElseIf CORTO(rsFactura!TIPODOC, 2, 0) = "E" Then
                    COMPROBANTE = Nota_de_Debito_E
                End If
            End If
        ElseIf CORTO(rsFactura!TIPODOC, 0, 2) = "F" Then
            If CORTO(rsFactura!TIPODOC, 1, 1) = "E" Then
               If CORTO(rsFactura!TIPODOC, 2, 0) = "A" Then
                    COMPROBANTE = Factura_de_Credito_electronica_A
                ElseIf CORTO(rsFactura!TIPODOC, 2, 0) = "B" Then
                    COMPROBANTE = Factura_de_Credito_electronica_B
                ElseIf CORTO(rsFactura!TIPODOC, 2, 0) = "C" Then
                    COMPROBANTE = Factura_de_Credito_electronica_C
                End If
            ElseIf CORTO(rsFactura!TIPODOC, 1, 1) = "A" Then
                If CORTO(rsFactura!TIPODOC, 2, 0) = "A" Then
                    COMPROBANTE = FACTURA_A
                ElseIf CORTO(rsFactura!TIPODOC, 2, 0) = "B" Then
                    COMPROBANTE = FACTURA_B
                ElseIf CORTO(rsFactura!TIPODOC, 2, 0) = "C" Then
                    COMPROBANTE = FACTURA_C
                ElseIf CORTO(rsFactura!TIPODOC, 2, 0) = "E" Then
                    COMPROBANTE = FACTURA_E
                End If
            End If
        ElseIf CORTO(rsFactura!TIPODOC, 0, 2) = "C" Then
            If CORTO(rsFactura!TIPODOC, 1, 1) = "E" Then
                If CORTO(rsFactura!TIPODOC, 2, 0) = "A" Then
                    COMPROBANTE = Nota_de_Credito_electronica_A
                ElseIf CORTO(rsFactura!TIPODOC, 2, 0) = "B" Then
                    COMPROBANTE = Nota_de_Credito_electronica_B
                ElseIf CORTO(rsFactura!TIPODOC, 2, 0) = "C" Then
                    COMPROBANTE = Nota_de_Credito_electronica_C
                End If
            End If
        ElseIf CORTO(rsFactura!TIPODOC, 0, 2) = "D" Then
            If CORTO(rsFactura!TIPODOC, 1, 1) = "E" Then
                If CORTO(rsFactura!TIPODOC, 2, 0) = "A" Then
                    COMPROBANTE = Nota_de_Debito_electronica_A
                ElseIf CORTO(rsFactura!TIPODOC, 2, 0) = "B" Then
                    COMPROBANTE = Nota_de_Debito_electronica_B
                ElseIf CORTO(rsFactura!TIPODOC, 2, 0) = "C" Then
                    COMPROBANTE = Nota_de_Debito_electronica_C
                End If
            End If
        End If
        

        
        fPuntoVenta = s2n(bPUNTOVENTA)

        If bResultado Then
            If CORTO(Trim(obtenerDeSQL("select tipodoc from facturaventa where codigo=" & IDFactura)), 2, 0) = "E" Then
                bResultado = CAE.xObtenerTicketAcceso()
                CAE.tGuardarTicketAcceso
                EstaEnLinea = CAE.xDummy
            'ElseIf CORTO(Trim(obtenerDeSQL("select tipodoc from facturaventa where codigo=" & IDFactura)), 1, 1) = "E" Then
                'bResultado = CAE.fecObtenerTicketAcceso()
                
            Else
                bResultado = CAE.f1ObtenerTicketAcceso()
                TicketGuardado = CAE.tGuardarTicketAcceso()
                EstaEnLinea = CAE.f1Dummy
                
            End If
        End If
        

        
        If EstaEnLinea = False Then
            MsgBox "El servidor no esta disponible." & CAE.UltimoMensajeError, vbInformation
            Exit Function
        End If
        
        
        If bResultado = False Then
            bResultado = CAE.f1RestaurarTicketAcceso(TicketGuardado)
        End If
        
        If CORTO(Trim(obtenerDeSQL("select tipodoc from facturaventa where codigo=" & IDFactura)), 2, 0) = "E" Then
            fexNroTMP = CAE.xFEGetLastCMP(fPuntoVenta, COMPROBANTE)
        Else
            fexNroTMP = CAE.F1CompUltimoAutorizado(fPuntoVenta, COMPROBANTE)
        End If
        
        If s2n(s2n(fexNroTMP) + 1) <> s2n(rsFactura!NroFactura) Then
            MsgBox "El ultimo comprobante autorizado es el " & s2n(fexNroTMP) & "." & Chr(13) & " El siguiente por aprobar deberia ser el " & s2n(s2n(fexNroTMP) + 1), vbInformation
            Exit Function
        End If
        
        
        'If Comprobante = TipoComprobante_FacturaEmportacion Or Comprobante = TipoComprobante_NotaDebitoExterior Or Comprobante = TipoComprobante_NotaCreditoExterior Then
        If COMPROBANTE = FACTURA_E Or COMPROBANTE = Nota_de_Debito_E Or COMPROBANTE = Nota_de_Credito_E Then
            CAE.F1CabeceraCantReg = 1
            CAE.xFecha_cbte = afipFecha(CDate(rsFactura!Fecha))
            CAE.xtipo_expo = 1
            If Trim(sSinNull(rsFactura!permisoembarque)) = "" Then
                If COMPROBANTE = TipoComprobante_NotaCreditoExterior Or COMPROBANTE = TipoComprobante_NotaDebitoExterior Then
                    CAE.xPermiso_existenteS = ""
                    CAE.xPermisoCantidad = 0
                    CAE.xPermisoNoInformar = 1
                Else
                    CAE.xPermiso_existenteS = "N"
                    CAE.xPermisoCantidad = 0
                    CAE.xPermisoNoInformar = 1
                End If
            Else
                CAE.xPermiso_existenteS = "S"
                CAE.xPermisoCantidad = 1
                CAE.xPermisoNoInformar = 0
            End If
            If Trim(sSinNull(rsFactura!paisfae)) = "" Then
                CAE.xDst_cmp = 208
            Else
                CAE.xDst_cmp = s2n(sSinNull(rsFactura!paisfae))
            End If
            CAE.xCliente = Trim(rsFactura!RAZONSOCIAL)
            CAE.xCuit_pais_clienteS = Replace(Trim(rsFactura!CUIT), "-", "") '"50000000016"
            CAE.xDomicilio_cliente = Trim(obtenerDeSQL("select direccion from clientes where codigo=" & s2n(rsFactura!cliente)))
            CAE.xMoneda_idS = Trim(nSinNull(obtenerDeSQL("select codigowsfex from monedas where codigo=" & nSinNull(rsFactura!moneda))))
            If CAE.xMoneda_idS = "" Then
                CAE.xMoneda_idS = "DOL"
            End If
            CAE.xMoneda_ctzS = s2n(rsFactura!cotizacion, 3)
            CAE.xObs_comerciales = "Sin observaciones"
            
            ttt = CDbl(rsFactura!Total)
            zzz = 1 'CDbl(rsFactura!cotizacion)' la cotizacion es siempre 1 porque las facturas electronicas se presentan siempre en pesos ante la afip
            
            CAE.xImp_total = ttt / zzz
            CAE.xImp_total = s2n(CAE.xImp_total, 3, True)
            CAE.xForma_pago = Trim(obtenerDeSQL("select descripcion from formaspago where codigo=" & rsFactura!formaPago))
            CAE.xIncoTerms = CORTO(Trim(rsFactura!incoterms), 0, Len(Trim(rsFactura!incoterms)) - 3)
            CAE.xIncoTerms_ds = Trim(CORTO(Trim(rsFactura!incoterms), 3, 0))
            CAE.xIdioma_cbte = 1
            

            fDetalle.Open "select * from facturaventadetalle where codigofactura=" & s2n(rsFactura!codigo), DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
            sumItem = 0
            If fDetalle.EOF And fDetalle.BOF Then
            Else
                CAE.xItemCantidad = fDetalle.RecordCount
                For i = 0 To fDetalle.RecordCount - 1
                    CAE.xIndiceItem = i
                    CAE.xITEMPro_codigo = Trim(fDetalle!producto)
                    CAE.xITEMPro_ds = Trim(fDetalle!DESCRIPCION)
                    CAE.xITEMPro_qty = s2n(fDetalle!cantidad)
                    CAE.xITEMPro_umed = 7 'UNIDADES
                    iii = CDbl(fDetalle!PrecioUnitario)
                    CAE.xITEMPro_precio_uni = s2n(iii / zzz, 3, True)
                    ptt = CDbl(fDetalle!PrecioTotal)
                    CAE.xITEMPro_precio_item = s2n(ptt / zzz, 3, True)
                    
                    sumItem = sumItem + ptt
                    fDetalle.MoveNext
                Next
                sumItem = sumItem / zzz
                CAE.xImp_total = s2n(sumItem, 3, True)
            End If
            
        'ElseIf Comprobante = TipoComprobante_FacturaA Or Comprobante = TipoComprobante_NotaCreditoA Or Comprobante = TipoComprobante_NotaDebitoA Or Comprobante = TipoComprobante_FacturaB Or Comprobante = TipoComprobante_NotaCreditoB Or Comprobante = TipoComprobante_NotaDebitoB Then
        ElseIf COMPROBANTE = FACTURA_A Or COMPROBANTE = Nota_de_Credito_A Or COMPROBANTE = Nota_de_Debito_A Or COMPROBANTE = FACTURA_B Or COMPROBANTE = Nota_de_Credito_B Or COMPROBANTE = Nota_de_Debito_B Or COMPROBANTE = Factura_de_Credito_electronica_A Or COMPROBANTE = Factura_de_Credito_electronica_B Or COMPROBANTE = Nota_de_Credito_electronica_A Or COMPROBANTE = Nota_de_Credito_electronica_B Or COMPROBANTE = Nota_de_Debito_electronica_A Or COMPROBANTE = Nota_de_Debito_electronica_B Then
        
            fPrDia = CDate("01/" & Month(CDate(rsFactura!Fecha)) & "/" & Year(CDate(rsFactura!Fecha)))
            CAE.F1CabeceraCantReg = 1
            
            '#RESOLUCION NOTAS CON FACTURA
            If COMPROBANTE = Nota_de_Credito_A Or COMPROBANTE = Nota_de_Credito_B Or COMPROBANTE = Nota_de_Credito_C Or COMPROBANTE = Nota_de_Debito_A Or COMPROBANTE = Nota_de_Debito_B Or COMPROBANTE = Nota_de_Debito_C Then
                If s2n(rsFactura!codfactura) > 0 Then
                    DatosFactura = obtenerDeSQL("select tipodoc,puntoventa,nrofactura,CUIT,fecha from facturaventa where codigo=" & s2n(rsFactura!codfactura))
                    CAE.F1DetalleCbtesAsocItemCantidad = 1
                    CAE.f1IndiceItem = 0
                    If CORTO(CStr(DatosFactura(0)), 0, 2) = "N" Then
                        If CORTO(CStr(DatosFactura(0)), 1, 1) = "C" Then
                            If CORTO(CStr(DatosFactura(0)), 2, 0) = "A" Then
                                CAE.F1DetalleCbtesAsocTipo = Nota_de_Credito_A
                            ElseIf CORTO(CStr(DatosFactura(0)), 2, 0) = "B" Then
                                CAE.F1DetalleCbtesAsocTipo = Nota_de_Credito_B
                            End If
                        ElseIf CORTO(CStr(DatosFactura(0)), 1, 1) = "D" Then
                            If CORTO(CStr(DatosFactura(0)), 2, 0) = "A" Then
                                CAE.F1DetalleCbtesAsocTipo = Nota_de_Debito_A
                            ElseIf CORTO(CStr(DatosFactura(0)), 2, 0) = "B" Then
                                CAE.F1DetalleCbtesAsocTipo = Nota_de_Debito_B
                            End If
                        End If
                    ElseIf CORTO(CStr(DatosFactura(0)), 0, 2) = "F" Then
                        If CORTO(CStr(DatosFactura(0)), 2, 0) = "A" Then
                            CAE.F1DetalleCbtesAsocTipo = FACTURA_A
                        ElseIf CORTO(CStr(DatosFactura(0)), 2, 0) = "B" Then
                            CAE.F1DetalleCbtesAsocTipo = FACTURA_B
                        End If
                    End If
                    CAE.F1DetalleCbtesAsocPtoVta = s2n(DatosFactura(1))
                    CAE.F1DetalleCbtesAsocNroS = DatosFactura(2)
                    CAE.F1DetalleCbtesAsocCUIT = CuitEmpresa 'Replace(DatosFactura(3), "-", "")
                    CAE.F1DetalleCbtesAsocFecha = afipFecha(CDate(DatosFactura(4)))
                End If
               End If
            '#FIN
            
            If COMPROBANTE = Nota_de_Credito_electronica_A Or COMPROBANTE = Nota_de_Credito_electronica_B Or COMPROBANTE = Nota_de_Debito_electronica_A Or COMPROBANTE = Nota_de_Debito_electronica_B Then
                If s2n(rsFactura!codfactura) > 0 Then
                    
                    DatosFactura = obtenerDeSQL("select tipodoc,puntoventa,nrofactura,CUIT,fecha from facturaventa where codigo=" & s2n(rsFactura!codfactura))
                    'DatosFactura = obtenerDeSQL("select tipodoc,puntoventa,nrofactura,CUIT,fecha from facturaventa where codigo=113")
                    CAE.F1DetalleCbtesAsocItemCantidad = 1
                    CAE.f1IndiceItem = 0
                    If CORTO(CStr(DatosFactura(0)), 2, 0) = "A" Then
                        CAE.F1DetalleCbtesAsocTipo = Factura_de_Credito_electronica_A
                    ElseIf CORTO(CStr(DatosFactura(0)), 2, 0) = "B" Then
                        CAE.F1DetalleCbtesAsocTipo = Factura_de_Credito_electronica_B
                    End If
                    
                    CAE.F1DetalleCbtesAsocPtoVta = s2n(DatosFactura(1))
                    CAE.F1DetalleCbtesAsocNroS = DatosFactura(2)
                    
                    CAE.F1DetalleCbtesAsocCUIT = CuitEmpresa 'Replace(DatosFactura(3), "-", "")
                    CAE.F1DetalleCbtesAsocFecha = afipFecha(CDate(DatosFactura(4)))
                    
                    CAE.F1DetalleOpcionalItemCantidad = 1
                    CAE.f1IndiceItem = 0
                    CAE.F1DetalleOpcionalId = 22
                    
                    If COMPROBANTE = Nota_de_Debito_electronica_A Or COMPROBANTE = Nota_de_Debito_electronica_B Then
                        CAE.F1DetalleOpcionalValor = "N"
                    Else
                        If MsgBox("El motivo de la nota electronica es porque la factura de credito fue rechazada?", vbYesNo + vbInformation) = vbYes Then
                            CAE.F1DetalleOpcionalValor = "S"
                        Else
                            CAE.F1DetalleOpcionalValor = "N" ' o "S" dependiendo si la factura fue anulada antes de ser enviada al cliente ("N") o fue anulada porque fue rechazada por el cliente ("S"). De igual forma si es cualquier otro motivo que no sea anulación (descuento, etc.) usar "N"
                        End If
                    End If
                End If
            End If
            
            fCBU = ""
            fALIAS = ""
            If COMPROBANTE = Factura_de_Credito_electronica_A Or COMPROBANTE = Factura_de_Credito_electronica_B Then
                fCBU = sSinNull((obtenerDeSQL("select cbu from datosempresa")))
                fALIAS = sSinNull((obtenerDeSQL("select alias from datosempresa")))
                CAE.F1DetalleFchVtoPago = afipFecha(CDate(rsFactura!Vencimiento))
                If Trim(fCBU) = "" Or Trim(fALIAS) = "" Then
                    MsgBox "Para emitir una factura de credito debe configurar CBU y ALIAS."
                    Exit Function
                Else
                    CAE.F1DetalleOpcionalItemCantidad = 3
                    CAE.f1IndiceItem = 0
                    CAE.F1DetalleOpcionalId = 2101
                    CAE.F1DetalleOpcionalValor = fCBU
                    CAE.f1IndiceItem = 1
                    CAE.F1DetalleOpcionalId = 2102
                    CAE.F1DetalleOpcionalValor = fALIAS
                    '#RESOLUCION RG4919/2021
                    CAE.f1IndiceItem = 2
                    CAE.F1DetalleOpcionalId = 27
                    If MsgBox("TRANSFERENCIA AL SISTEMA DE CIRCULACION ABIERTA presione SI." & Chr(13) & "AGENTE DE DEPOSITO COLECTIVO presione NO.", vbYesNo) = vbYes Then
                        CAE.F1DetalleOpcionalValor = "SCA" 'TRANSFERENCIA AL SISTEMA DE CIRCULACION ABIERTA
                    Else
                        CAE.F1DetalleOpcionalValor = "ADC" 'AGENTE DE DEPOSITO COLECTIVO
                    End If
                    '#FIN
                End If
            End If
            
            
            
            zzz = CDbl(rsFactura!cotizacion)
            If zzz = 0 Then zzz = 1
            If zzz = 1 Then
                CAE.F1DetalleMonId = "PES"
                CAE.F1DetalleMonCotiz = 1
            Else
                CAE.F1DetalleMonId = "DOL"
                CAE.F1DetalleMonCotiz = zzz
            End If
            
            CAE.F1CabeceraPtoVta = fPuntoVenta
            CAE.F1CabeceraCbteTipo = COMPROBANTE
            CAE.f1Indice = 0
            CAE.F1DetalleConcepto = 3
            CAE.F1DetalleDocNro = s2n(Replace(rsFactura!CUIT, "-", ""))
            If Len(Trim(CAE.F1DetalleDocNro)) <= 8 Then
                CAE.F1DetalleDocTipo = TipoDocumento_DNI
            ElseIf Len(Trim(CAE.F1DetalleDocNro)) = 11 Then
                CAE.F1DetalleDocTipo = TipoDocumento_CUIT
            Else
                CAE.F1DetalleDocTipo = TipoDocumento_Pasaporte
            End If
            CAE.F1DetalleCbteDesde = s2n(rsFactura!NroFactura)
            CAE.F1DetalleCbteHasta = s2n(rsFactura!NroFactura)
            CAE.F1DetalleCbteFch = afipFecha(CDate(rsFactura!Fecha))
            CAE.F1DetalleImpTotal = s2n(rsFactura!Total / zzz)
            CAE.F1DetalleImpTotalConc = 0
            CAE.F1DetalleImpNeto = s2n(rsFactura!Neto / zzz)
            CAE.F1DetalleImpOpEx = 0
            CAE.F1DetalleImpTrib = s2n(rsFactura!IIBB / zzz)
            CAE.F1DetalleImpIva = s2n(rsFactura!Iva / zzz)
            
            CAE.F1DetalleFchServDesde = afipFecha(fPrDia)
            CAE.F1DetalleFchServHasta = afipFecha(ultimoDiaDelMes(fPrDia))
            If COMPROBANTE = FACTURA_A Or COMPROBANTE = FACTURA_B Or COMPROBANTE = FACTURA_C Or COMPROBANTE = Nota_de_Credito_A Or COMPROBANTE = Nota_de_Credito_B Or COMPROBANTE = Nota_de_Credito_C Or COMPROBANTE = Nota_de_Debito_A Or COMPROBANTE = Nota_de_Debito_B Or COMPROBANTE = Nota_de_Debito_C Then
                CAE.F1DetalleFchVtoPago = afipFecha(CDate(rsFactura!Vencimiento))
            End If
            ttt = CDbl(rsFactura!Total)
                        
            
            If s2n(CAE.F1DetalleImpTrib) > 0 Then
                CAE.F1DetalleTributoItemCantidad = 1
                CAE.f1IndiceItem = 0
                CAE.F1DetalleTributoId = 2 '1=nacional,2=provincial,3=municipal,4=interno,99=otros
                CAE.F1DetalleTributoDesc = "IIBB"
                CAE.F1DetalleTributoBaseImp = s2n(rsFactura!Neto / zzz)
                CAE.F1DetalleTributoAlic = s2n(rsFactura!AlicIIBB)
                CAE.F1DetalleTributoImporte = s2n(rsFactura!IIBB / zzz)
            Else
                CAE.F1DetalleTributoItemCantidad = 0
                CAE.f1IndiceItem = 0
                CAE.F1DetalleTributoId = 0
                CAE.F1DetalleTributoDesc = ""
                CAE.F1DetalleTributoBaseImp = 0
                CAE.F1DetalleTributoAlic = 0
                CAE.F1DetalleTributoImporte = 0
            End If
            

            
            Set fDetalle = Nothing
            fDetalle.Open "select *,_iva as piva from facturaventadetalle where codigofactura=" & s2n(rsFactura!codigo), DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
            sumItem = 0
            PIVA = s2n(rsFactura!PorcentajeIva, 4)
            If fDetalle.EOF And fDetalle.BOF Then
            Else

                Tipo3 = 0
                Tipo4 = 0
                Tipo5 = 0
                Tipo3base = 0
                Tipo4base = 0
                Tipo5base = 0
                For i = 0 To fDetalle.RecordCount - 1
                    
                    If s2n(rsFactura!Descuento, 4) > 0 Then
                        ptt = CDbl(fDetalle!PrecioTotal)
                        ptt = ptt - (ptt * s2n(rsFactura!Descuento, 4))
                    Else
                        ptt = CDbl(fDetalle!PrecioTotal)
                    End If
                    PIVA = s2n(fDetalle!PIVA / 100, 4)
                    If CORTO(Trim(rsFactura!TIPODOC), 2, 0) = "B" Then
                        PIVA = 0
                    End If
                    iii = 0
                    If CORTO(Trim(rsFactura!TIPODOC), 2, 0) = "A" Then
                        iii = s2n(s2n((ptt * PIVA), 4) / zzz, 4)
                    ElseIf CORTO(Trim(rsFactura!TIPODOC), 2, 0) = "B" Then
                        iii = s2n(s2n(ptt - (ptt / (1 + PIVA)), 4) / zzz, 4)
                    End If

                    If PIVA * 100 = 21 Then
                        'CAE.F1DetalleIvaId = 5 '5 es 21 , 4 es 10.5
                        Tipo5base = Tipo5base + s2n(ptt / zzz, 4, True)
                        Tipo5 = s2n(Tipo5 + iii, 4)
                    ElseIf PIVA * 100 = 10.5 Then
                        'CAE.F1DetalleIvaId = 4 '5 es 21 , 4 es 10.5
                        Tipo4base = Tipo4base + s2n(ptt / zzz, 4, True)
                        Tipo4 = s2n(Tipo4 + iii, 4)
                    Else
                        'CAE.F1DetalleIvaId = 3
                        Tipo3base = Tipo3base + s2n(ptt / zzz, 4, True)
                        Tipo3 = s2n(Tipo3 + iii, 4)
                    End If
                    
                    sumItem = sumItem + ptt
                    fDetalle.MoveNext
                Next
                CAE.F1DetalleIvaItemCantidad = 0
                If s2n(Tipo3base) > 0 Then
                    CAE.F1DetalleIvaItemCantidad = CAE.F1DetalleIvaItemCantidad + 1
                End If
                If s2n(Tipo4base) > 0 Then
                    CAE.F1DetalleIvaItemCantidad = CAE.F1DetalleIvaItemCantidad + 1
                End If
                If s2n(Tipo5base) > 0 Then
                    CAE.F1DetalleIvaItemCantidad = CAE.F1DetalleIvaItemCantidad + 1
                End If
                
                For i = 0 To CAE.F1DetalleIvaItemCantidad - 1
                    CAE.f1IndiceItem = i
                    If s2n(Tipo3base) > 0 Then
                        CAE.F1DetalleIvaId = 3
                        CAE.F1DetalleIvaBaseImp = s2n(Tipo3base, 2)
                        CAE.F1DetalleIvaImporte = s2n(Tipo3, 2)
                        Tipo3base = 0
                    'End If
                    ElseIf s2n(Tipo4base) > 0 Then
                        CAE.F1DetalleIvaId = 4
                        CAE.F1DetalleIvaBaseImp = s2n(Tipo4base, 2)
                        CAE.F1DetalleIvaImporte = s2n(Tipo4, 2)
                        Tipo4base = 0
                    'End If
                    ElseIf s2n(Tipo5base) > 0 Then
                        CAE.F1DetalleIvaId = 5
                        CAE.F1DetalleIvaBaseImp = s2n(Tipo5base, 2)
                        CAE.F1DetalleIvaImporte = s2n(Tipo5, 2)
                        Tipo5base = 0
                    End If
                    
                Next

                sumItem = sumItem / zzz
            End If
            
            

        'ElseIf Comprobante = TipoComprobante_FacturaC Or Comprobante = TipoComprobante_NotaCreditoC Or Comprobante = TipoComprobante_NotaDebitoC Then
        ElseIf COMPROBANTE = FACTURA_C Or COMPROBANTE = Nota_de_Credito_C Or COMPROBANTE = Nota_de_Debito_C Or COMPROBANTE = Factura_de_Credito_electronica_C Or COMPROBANTE = Nota_de_Debito_electronica_C Or COMPROBANTE = Nota_de_Credito_electronica_C Then
        
            fPrDia = CDate("01/" & Month(CDate(rsFactura!Fecha)) & "/" & Year(CDate(rsFactura!Fecha)))
            
            CAE.F1CabeceraCantReg = 1
            
            If COMPROBANTE = Nota_de_Credito_electronica_C Then
                If s2n(rsFactura!codfactura) > 0 Then
                    
                    DatosFactura = obtenerDeSQL("select tipodoc,puntoventa,nrofactura,CUIT,fecha from facturaventa where codigo=" & s2n(rsFactura!codfactura))
                    'DatosFactura = obtenerDeSQL("select tipodoc,puntoventa,nrofactura,CUIT,fecha from facturaventa where codigo=113")
                    CAE.F1DetalleCbtesAsocItemCantidad = 1
                    CAE.f1IndiceItem = 0
                    If CORTO(CStr(DatosFactura(0)), 2, 0) = "A" Then
                        CAE.F1DetalleCbtesAsocTipo = Factura_de_Credito_electronica_A
                    ElseIf CORTO(CStr(DatosFactura(0)), 2, 0) = "B" Then
                        CAE.F1DetalleCbtesAsocTipo = Factura_de_Credito_electronica_B
                    ElseIf CORTO(CStr(DatosFactura(0)), 2, 0) = "B" Then
                        CAE.F1DetalleCbtesAsocTipo = Factura_de_Credito_electronica_C
                    End If
                    
                    CAE.F1DetalleCbtesAsocPtoVta = s2n(DatosFactura(1))
                    CAE.F1DetalleCbtesAsocNroS = DatosFactura(2)
                    
                    CAE.F1DetalleCbtesAsocCUIT = CuitEmpresa 'Replace(DatosFactura(3), "-", "")
                    CAE.F1DetalleCbtesAsocFecha = afipFecha(CDate(DatosFactura(4)))
                    
                    'CAE.F1DetalleOpcionalItemCantidad = 1
                    CAE.f1IndiceItem = 0
                    CAE.F1DetalleOpcionalId = 22
                    CAE.F1DetalleOpcionalValor = "N" ' o "S" dependiendo si la factura fue anulada antes de ser enviada al cliente ("N") o fue anulada porque fue rechazada por el cliente ("S"). De igual forma si es cualquier otro motivo que no sea anulación (descuento, etc.) usar "N"
                End If
            End If
            
            fCBU = ""
            fALIAS = ""
            If COMPROBANTE = Factura_de_Credito_electronica_C Or COMPROBANTE = Nota_de_Credito_electronica_C Then
                fCBU = sSinNull((obtenerDeSQL("select cbu from datosempresa")))
                fALIAS = sSinNull((obtenerDeSQL("select alias from datosempresa")))
                CAE.F1DetalleFchVtoPago = afipFecha(CDate(rsFactura!Vencimiento))
                If Trim(fCBU) = "" Or Trim(fALIAS) = "" Then
                    MsgBox "Para emitir una factura de credito debe configurar CBU y ALIAS."
                    Exit Function
                Else
                    CAE.F1DetalleOpcionalItemCantidad = 2
                    CAE.f1IndiceItem = 0
                    CAE.F1DetalleOpcionalId = 2101
                    CAE.F1DetalleOpcionalValor = fCBU
                    CAE.f1IndiceItem = 1
                    CAE.F1DetalleOpcionalId = 2102
                    CAE.F1DetalleOpcionalValor = fALIAS
                End If
            End If
            
            zzz = 1 'CDbl(rsFactura!cotizacion)
            If zzz = 0 Then zzz = 1
            If zzz = 1 Then
                CAE.F1DetalleMonId = "PES"
                CAE.F1DetalleMonCotiz = 1
            Else
                CAE.F1DetalleMonId = "DOL"
                CAE.F1DetalleMonCotiz = zzz
            End If
            
            CAE.F1CabeceraPtoVta = fPuntoVenta
            CAE.F1CabeceraCbteTipo = COMPROBANTE
            CAE.f1Indice = 0
            CAE.F1DetalleConcepto = 3
            CAE.F1DetalleDocNro = s2n(Replace(rsFactura!CUIT, "-", ""))
            If Len(Trim(CAE.F1DetalleDocNro)) <= 8 Then
                CAE.F1DetalleDocTipo = TipoDocumento_DNI
            ElseIf Len(Trim(CAE.F1DetalleDocNro)) = 11 Then
                CAE.F1DetalleDocTipo = TipoDocumento_CUIT
            Else
                CAE.F1DetalleDocTipo = TipoDocumento_Pasaporte
            End If
            CAE.F1DetalleDocNro = s2n(Replace(rsFactura!CUIT, "-", ""))
            CAE.F1DetalleCbteDesde = s2n(rsFactura!NroFactura)
            CAE.F1DetalleCbteHasta = s2n(rsFactura!NroFactura)
            CAE.F1DetalleCbteFch = afipFecha(CDate(rsFactura!Fecha))
            CAE.F1DetalleImpTotal = s2n(rsFactura!Total / zzz)
            CAE.F1DetalleImpTotalConc = 0
            CAE.F1DetalleImpNeto = s2n(rsFactura!Neto / zzz)
            CAE.F1DetalleImpOpEx = 0
            CAE.F1DetalleImpTrib = s2n(rsFactura!IIBB / zzz)
            CAE.F1DetalleImpIva = s2n(rsFactura!Iva / zzz)
            
            CAE.F1DetalleFchServDesde = afipFecha(fPrDia)
            CAE.F1DetalleFchServHasta = afipFecha(ultimoDiaDelMes(fPrDia))
            CAE.F1DetalleFchVtoPago = afipFecha(CDate(rsFactura!Vencimiento))
            ttt = CDbl(rsFactura!Total)
                        
            
            If s2n(CAE.F1DetalleImpTrib) > 0 Then
                CAE.F1DetalleTributoItemCantidad = 1
                CAE.f1IndiceItem = 0
                CAE.F1DetalleTributoId = 2 '1=nacional,2=provincial,3=municipal,4=interno,99=otros
                CAE.F1DetalleTributoDesc = "IIBB"
                CAE.F1DetalleTributoBaseImp = s2n(rsFactura!Neto / zzz)
                CAE.F1DetalleTributoAlic = s2n(rsFactura!AlicIIBB)
                CAE.F1DetalleTributoImporte = s2n(rsFactura!IIBB / zzz)
            Else
                CAE.F1DetalleTributoItemCantidad = 0
                CAE.f1IndiceItem = 0
                CAE.F1DetalleTributoId = 0
                CAE.F1DetalleTributoDesc = ""
                CAE.F1DetalleTributoBaseImp = 0
                CAE.F1DetalleTributoAlic = 0
                CAE.F1DetalleTributoImporte = 0
            End If
            

            
            Set fDetalle = Nothing
            fDetalle.Open "select *,_iva as piva from facturaventadetalle where codigofactura=" & s2n(rsFactura!codigo), DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
            sumItem = 0
            PIVA = s2n(rsFactura!PorcentajeIva, 4)
            If fDetalle.EOF And fDetalle.BOF Then
            Else

                Tipo3 = 0
                Tipo4 = 0
                Tipo5 = 0
                Tipo3base = 0
                Tipo4base = 0
                Tipo5base = 0
                For i = 0 To fDetalle.RecordCount - 1
                   
                    If s2n(rsFactura!Descuento, 4) > 0 Then
                        ptt = CDbl(fDetalle!PrecioTotal)
                        ptt = ptt - (ptt * s2n(rsFactura!Descuento, 4))
                    Else
                        ptt = CDbl(fDetalle!PrecioTotal)
                    End If
                    PIVA = s2n(fDetalle!PIVA / 100, 4)
                    If CORTO(Trim(rsFactura!TIPODOC), 2, 0) = "B" Or CORTO(Trim(rsFactura!TIPODOC), 2, 0) = "C" Then
                        PIVA = 0
                    End If
                    iii = 0
                    If CORTO(Trim(rsFactura!TIPODOC), 2, 0) = "A" Then
                        iii = s2n(s2n((ptt * PIVA), 4) / zzz, 4)
                    ElseIf CORTO(Trim(rsFactura!TIPODOC), 2, 0) = "B" Then
                        iii = s2n(s2n(ptt - (ptt / (1 + PIVA)), 4) / zzz, 4)
                    End If

                    If PIVA * 100 = 21 Then
                        'CAE.F1DetalleIvaId = 5 '5 es 21 , 4 es 10.5
                        Tipo5base = Tipo5base + s2n(ptt / zzz, 4, True)
                        Tipo5 = s2n(Tipo5 + iii, 4)
                    ElseIf PIVA * 100 = 10.5 Then
                        'CAE.F1DetalleIvaId = 4 '5 es 21 , 4 es 10.5
                        Tipo4base = Tipo4base + s2n(ptt / zzz, 4, True)
                        Tipo4 = s2n(Tipo4 + iii, 4)
                    Else
                        'CAE.F1DetalleIvaId = 3
                        Tipo3base = Tipo3base + s2n(ptt / zzz, 4, True)
                        Tipo3 = s2n(Tipo3 + iii, 4)
                    End If
                    
                    sumItem = sumItem + ptt
                    fDetalle.MoveNext
                Next
                CAE.F1DetalleIvaItemCantidad = 0
                If CORTO(Trim(rsFactura!TIPODOC), 2, 0) = "C" Then
                Else
                    If s2n(Tipo3base) > 0 Then
                        CAE.F1DetalleIvaItemCantidad = CAE.F1DetalleIvaItemCantidad + 1
                    End If
                    If s2n(Tipo4base) > 0 Then
                        CAE.F1DetalleIvaItemCantidad = CAE.F1DetalleIvaItemCantidad + 1
                    End If
                    If s2n(Tipo5base) > 0 Then
                        CAE.F1DetalleIvaItemCantidad = CAE.F1DetalleIvaItemCantidad + 1
                    End If
                
                    
                    For i = 0 To CAE.F1DetalleIvaItemCantidad - 1
                        CAE.f1IndiceItem = i
                        If s2n(Tipo3base) > 0 Then
                            CAE.F1DetalleIvaId = 3
                            CAE.F1DetalleIvaBaseImp = s2n(Tipo3base, 2)
                            CAE.F1DetalleIvaImporte = s2n(Tipo3, 2)
                            Tipo3base = 0
                        'End If
                        ElseIf s2n(Tipo4base) > 0 Then
                            CAE.F1DetalleIvaId = 4
                            CAE.F1DetalleIvaBaseImp = s2n(Tipo4base, 2)
                            CAE.F1DetalleIvaImporte = s2n(Tipo4, 2)
                            Tipo4base = 0
                        'End If
                        ElseIf s2n(Tipo5base) > 0 Then
                            CAE.F1DetalleIvaId = 5
                            CAE.F1DetalleIvaBaseImp = s2n(Tipo5base, 2)
                            CAE.F1DetalleIvaImporte = s2n(Tipo5, 2)
                            Tipo5base = 0
                        End If
                        
                    Next

                    sumItem = sumItem / zzz
                End If
            End If
            
        End If
        
        'IdentificadorF = sSinNull(obtenerDeSQL("select identificador from facturaventa where codigo=" & IDFactura))
        'If IdentificadorF = "" Or IdentificadorF = "0" Then
            IdentificadorF = CAE.Identificador
            DataEnvironment1.Sistema.Execute "update facturaventa set identificador=" & ssTexto(IdentificadorF) & " where codigo=" & IDFactura
        'Else
        '    CAE.Identificador = IdentificadorF
        'End If
   
        
        
        If Right(Trim(rsFactura!TIPODOC), 1) = "E" Then
            bResultado = CAE.xRegistrar(fPuntoVenta, COMPROBANTE, IdentificadorF)
        'ElseIf CORTO(Trim(rsFactura!TIPODOC), 1, 1) = "E" Then
            'bResultado = CAE.Registrar(fPuntoVenta, COMPROBANTE, IdentificadorF)
            'bResultado = CAE.F1CAESolicitar()
        Else
            CAE.F1DetalleDocTipo = TipoDocumento_CUIT
            CAE.ArchivoXMLEnviado = "c:\WSFEv1_enviado.xml"
            CAE.ArchivoXMLRecibido = "c:\WSFEv1_recibido.xml"
            bResultado = CAE.F1CAESolicitar()
        End If
        
        If COMPROBANTE = TipoComprobante_FacturaEmportacion Or COMPROBANTE = TipoComprobante_NotaDebitoExterior Or COMPROBANTE = TipoComprobante_NotaCreditoExterior Then
            If bResultado Then
                bCuit = Format(CuitEmpresa, "00000000000")
                bPUNTOVENTA = bPUNTOVENTA
                bCodFactura = Format(Trim(obtenerDeSQL("select codfactura from documentoscae where tipopunto='WS' and tipo=" & ssTexto(rsFactura!TIPODOC) & " and puntoventa=" & ssTexto(bPUNTOVENTA))), "00")
                
                bCAE = CAE.xRespuestaCAE
                
                bFechaCAE = CAE.xRespuestaFch_vence_cae
                
                bBARRA = bCuit & bCodFactura & bPUNTOVENTA & bCAE & bFechaCAE
                bBARRA = bBARRA & CodVerificador(bBARRA)
                
                DataEnvironment1.Sistema.Execute "update facturaventa set barra=" & ssTexto(bBARRA) & ",identificador=" & ssTexto(IdentificadorF) & ",caev=" & ssFecha(aFecha(CAE.xRespuestaFch_vence_cae)) & ", cae=" & ssTexto(CAE.xRespuestaCAE) & ",PermisoEmbarque=" & ssTexto(CPERMISO) & " where codigo=" & IDFactura
                MsgBox "CAE: " & CAE.xRespuestaCAE, vbInformation
                MsgBox "Reproceso: " & CAE.xRespuestaReproceso & Chr(10) & "Comprobante: " & CAE.xRespuestacbte_numeroS & Chr(10) & "Resultado:" & CAE.xRespuestaResultado
                            
            Else
                MsgBox ("Motivo: " & Chr(10) & " Error " & CAE.Permsg & "Detalle: " & CAE.xerrmsg & " - " & CAE.UltimoNumeroError)
                MsgBox ("Reproceso: " & CAE.xRespuestaReproceso & Chr(10) & "Comprobante: " & CAE.xRespuestacbte_numeroS & Chr(10) & "Resultado:" & CAE.xRespuestaResultado)
            End If
        Else
            If bResultado Then
                bCuit = Format(CuitEmpresa, "00000000000")
                bCodFactura = Format(Trim(obtenerDeSQL("select codfactura from documentoscae where tipopunto='WS' and tipo=" & ssTexto(rsFactura!TIPODOC) & " and puntoventa=" & ssTexto(bPUNTOVENTA))), "00")
                bPUNTOVENTA = bPUNTOVENTA
                
                fexNroTMP = CAE.fecSign
                If Trim(CAE.F1RespuestaDetalleCae) = "" Then
                    MsgBox "CAE no asignado", vbInformation
                    If FrmPrincipal.chkDesbichando Then
                        With CAE
                            MsgBox .F1RespuestaResultado
                            MsgBox .F1RespuestaReProceso
                            MsgBox .f1ErrorCode1
                            MsgBox .f1ErrorMsg1
                            MsgBox .UltimoNumeroError
                            MsgBox .UltimoMensajeError
                            MsgBox .F1RespuestaCantidadReg
                            MsgBox .F1RespuestaDetalleResultado
                            MsgBox .F1RespuestaDetalleCae
                            MsgBox .F1RespuestaDetalleCAEFchVto
                            MsgBox .F1RespuestaDetalleObservacionCode1
                            MsgBox .F1RespuestaDetalleObservacionMsg1
                            MsgBox .F1RespuestaDetalleObservacionItemCantidad
                        End With
                    End If
                    Exit Function
                End If
                
                bCAE = CAE.F1RespuestaDetalleCae
                bFechaCAE = CAE.F1RespuestaDetalleCAEFchVto
                bBARRA = bCuit & bCodFactura & bPUNTOVENTA & bCAE & bFechaCAE
                bBARRA = bBARRA & CodVerificador(bBARRA)
                DataEnvironment1.Sistema.Execute "update facturaventa set barra=" & ssTexto(bBARRA) & ",identificador=" & ssTexto(Identificador) & ",caev=" & ssFecha(aFecha(CAE.F1RespuestaDetalleCAEFchVto)) & ", cae=" & ssTexto(CAE.F1RespuestaDetalleCae) & ",PermisoEmbarque=" & ssTexto(CPERMISO) & " where codigo=" & IDFactura
                'MsgBox "CAE: " & CAE.F1RespuestaDetalleCae, vbInformation
                'MsgBox "Reproceso: " & CAE.FERespuestaReproceso & Chr(10) & "Cantidad de Reg: " & CAE.FERespuestaCantidadReg & Chr(10) & "Resultado:" & CAE.FERespuestaResultado
                MsgBox "CAE: " & CAE.F1RespuestaDetalleCae & Chr(13) & "Resultado de AFIP: " + CAE.F1RespuestaResultado & Chr(13) & "Registros procesados por AFIP: " + str(CAE.F1RespuestaCantidadReg), vbInformation
            Else
                MsgBox "Resultado de AFIP: " & IIf(Trim(CAE.F1RespuestaResultado) = "", "R", CAE.F1RespuestaResultado) & Chr(13) & "Detalle " & IIf(Trim(CAE.F1RespuestaDetalleObservacionCode1) = "", CAE.UltimoNumeroError, CAE.F1RespuestaDetalleObservacionCode1) & " : " & IIf(Trim(CAE.F1RespuestaDetalleObservacionMsg1) = "", CAE.UltimoMensajeError, CAE.F1RespuestaDetalleObservacionMsg1), vbInformation
                'MsgBox "Resultado global AFIP: " + CAE.fecErrorCodigo1 & Chr(13) & "Es reproceso? " + CAE.F1RespuestaReProceso & Chr(13) & "Registros procesados por AFIP: " + str(CAE.fecErrorItemCantidad) & Chr(13) & "ERROR genérico global:" + CAE.fecErrorDescripcion1, vbInformation
            End If
        End If
        EmisionFacturaElectronica = bResultado
    Else
        'MsgBox "Resultado global AFIP: " + CAE.F1RespuestaResultado & Chr(13) & "Es reproceso? " + CAE.F1RespuestaReProceso & Chr(13) & "Registros procesados por AFIP: " + str(CAE.F1RespuestaCantidadReg) & Chr(13) & "ERROR genérico global:" + CAE.f1ErrorMsg1, vbInformation
        MsgBox "Resultado de AFIP: " & IIf(Trim(CAE.F1RespuestaResultado) = "", "R", CAE.F1RespuestaResultado) & Chr(13) & "Detalle " & IIf(Trim(CAE.F1RespuestaDetalleObservacionCode1) = "", CAE.UltimoNumeroError, CAE.F1RespuestaDetalleObservacionCode1) & " : " & IIf(Trim(CAE.F1RespuestaDetalleObservacionMsg1) = "", CAE.UltimoMensajeError, CAE.F1RespuestaDetalleObservacionMsg1), vbInformation
    End If
    
'    CAE.f1IndiceItem
'    CAE.F1DetalleOpcionalId
'    CAE.F1DetalleOpcionalValor
'
'    CAE.F1DetalleCbtesAsocTipo = 201
'    CAE.F1DetalleCbtesAsocPtoVta = 1
'    CAE.F1DetalleCbtesAsocNroS = "1"
'    CAE.F1DetalleCbtesAsocCUIT = "1232423 cuit emisor del comprobante asociado"
'    CAE.F1DetalleCbtesAsocFecha = "20190530"
    
End Function

Public Function DatosCuit(nCuit) As Variant
Dim ws As New WSAFIPFE.misdatos
Dim bResultado As Boolean
    If ws.iniciar("licencia.betasepp@gmail.com", "betasepp6459") Then
       bResultado = ws.cuitBuscar(nCuit)
       If ws.UltimoMensajeError = "" Then
            If ws.cuitMensajeError = "" Then
                MsgBox ("nombre: " + ws.cuitDenominacion + "iva: " + ws.cuitIVA + "monotributo: " + ws.cuitMonotributo)
            Else
                MsgBox ("El cuit no existe " + ws.cuitMensajeError)
            End If
       Else
            MsgBox ("Fallo la conexión " + ws.UltimoMensajeError)
       End If
    Else
         MsgBox ("Fallo al iniciar " + ws.UltimoMensajeError)
    End If
End Function

Public Function EstaEnPadron(sCuit As String, Optional ByRef nMonto As Double = 0) As Boolean
Dim afipDatos As New WSAFIPFE.Factura
Dim bResultadoPadron As Boolean, bResultadoBase As Boolean
Dim sEMPRESA As String, CuitEmpresa As String, Certificado As String, Licencia As String
Dim DatosResultado As Variant
Dim bResultadoCliente As Integer, bResultadoIva As Integer
Dim sMonto As String, sObligado As String


sEMPRESA = obtenerDeSQL("select nombrearchivo from datosempresa")

CuitEmpresa = Trim(Replace(obtenerDeSQL("select cuitempresa from datosempresa where idempresa=" & gEMPR_idEmpresa), "-", ""))
If CuitEmpresa = "" Then
    MsgBox "No se puede seguir con el proceso, el CUIT de la empresa no esta establecido.", vbCritical
    Exit Function
End If
If Len(CuitEmpresa) < 11 Then
    MsgBox "No se puede seguir con el proceso, el CUIT de la empresa no es valido.", vbCritical
    Exit Function
End If

Certificado = App.Path & "\" & sEMPRESA & ".pfx"
If ExisteArchivo(Certificado) Then
Else
    MsgBox "No se puede seguir con el proceso, el certificado no existe.", vbCritical
    Exit Function
End If

Licencia = App.Path & "\" & sEMPRESA & ".lic"
If ExisteArchivo(Licencia) Then
Else
    MsgBox "No se puede seguir con el proceso, la licencia no existe.", vbCritical
    Exit Function
End If



If afipDatos.iniciar(1, CuitEmpresa, Certificado, Licencia) Then
    afipDatos.p1Version = 5
    afipDatos.ArchivoXMLEnviado = "d:\p1envio.xml"
    afipDatos.ArchivoXMLRecibido = "d:\p1recibo.xml"
    afipDatos.ArchivoCertificadoPassword = ""
    If afipDatos.fecObtenerTicketAcceso() Then
        sCuit = Replace(sCuit, "-", "")
        bResultadoPadron = afipDatos.fecconsultarMontoObligadoRecepcion(sCuit, afipFecha(Date))
        If bResultadoPadron Then
            If afipDatos.UltimoMensajeError = "" And afipDatos.fecErrorDescripcion1 = "" Then
                'MsgBox (afipDatos.fecLeerPropiedad("fecconsultarMontoObligadoRecepcion", "montoDesde", "", 0, 0))
                'MsgBox (afipDatos.fecLeerPropiedad("fecconsultarMontoObligadoRecepcion", "obligado", "", 0, 0))
                sMonto = afipDatos.fecLeerPropiedad("fecconsultarMontoObligadoRecepcion", "montoDesde", "", 0, 0)
                sObligado = afipDatos.fecLeerPropiedad("fecconsultarMontoObligadoRecepcion", "obligado", "", 0, 0)
                nMonto = CDbl(sMonto)
                If sObligado = "S" Then
                    bResultadoPadron = True
                Else
                    bResultadoPadron = False
                End If
            Else
                'MsgBox ("error leer" + afipDatos.UltimoMensajeError)
                'MsgBox ("error ller servidor " + afipDatos.fecErrorDescripcion1)
                bResultadoPadron = False
            End If
        Else
            'MsgBox ("fallo método " + afipDatos.UltimoMensajeError)
            bResultadoPadron = False
        End If
    Else
        bResultadoPadron = False
    End If
End If

bResultadoIva = CodigoIvaGranEmpresa()
bResultadoCliente = nSinNull(obtenerDeSQL("select iva from clientes where cuit=" & ssTexto(sCuit)))
If bResultadoCliente = bResultadoIva Then
    bResultadoBase = True
Else
    bResultadoBase = False
End If

If bResultadoPadron = True And bResultadoBase = True Then
    EstaEnPadron = True
Else
    If bResultadoPadron Then
        DataEnvironment1.Sistema.Execute "update clientes set iva=" & bResultadoIva & " where cuit=" & ssTexto(sCuit)
        EstaEnPadron = True
    Else
        EstaEnPadron = False
    End If
End If


End Function

Public Function ConsultarDatos(sCuit As String)
Dim afipDatos As New WSAFIPFE.Factura
Dim bResultado As Boolean
Dim sEMPRESA As String, CuitEmpresa As String, Certificado As String, Licencia As String
Dim ClienteApellido As String, ClienteNombre As String, ClienteRazon As String, ClienteLocalidad As String
Dim ClienteDireccion As String, ClienteTipoDomicilio As String, ClienteMonotributo As String, ClienteInscripto As String, ClienteExento As String
Dim ClienteCodPostal As String, ClienteProvincia As String
Dim DatosResultado As Variant

sEMPRESA = obtenerDeSQL("select nombrearchivo from datosempresa")

CuitEmpresa = Trim(Replace(obtenerDeSQL("select cuitempresa from datosempresa where idempresa=" & gEMPR_idEmpresa), "-", ""))
If CuitEmpresa = "" Then
    MsgBox "No se puede seguir con el proceso, el CUIT de la empresa no esta establecido.", vbCritical
    Exit Function
End If
If Len(CuitEmpresa) < 11 Then
    MsgBox "No se puede seguir con el proceso, el CUIT de la empresa no es valido.", vbCritical
    Exit Function
End If

Certificado = App.Path & "\" & sEMPRESA & ".pfx"
If ExisteArchivo(Certificado) Then
Else
    MsgBox "No se puede seguir con el proceso, el certificado no existe.", vbCritical
    Exit Function
End If

Licencia = App.Path & "\" & sEMPRESA & ".lic"
If ExisteArchivo(Licencia) Then
Else
    MsgBox "No se puede seguir con el proceso, la licencia no existe.", vbCritical
    Exit Function
End If

afipDatos.tls = 12
If afipDatos.iniciar(1, CuitEmpresa, Certificado, Licencia) Then
    afipDatos.p1Version = 5
    afipDatos.ArchivoXMLEnviado = "d:\p1envio.xml"
    afipDatos.ArchivoXMLRecibido = "d:\p1recibo.xml"
    afipDatos.ArchivoCertificadoPassword = ""
    If afipDatos.p1ObtenerTicketAcceso() Then
        bResultado = afipDatos.p1GetPersona(sCuit)
        'bResultado = afipDatos.p5GetPersona(sCuit)
        If afipDatos.UltimoMensajeError = "" Then
            ClienteApellido = afipDatos.p1LeerPropiedad("p1getPersona", "datosgenerales.apellido", "", 0, 0)
            ClienteNombre = afipDatos.p1LeerPropiedad("p1getPersona", "datosgenerales.nombre", "", 0, 0)
            ClienteRazon = afipDatos.p1LeerPropiedad("p1getPersona", "datosgenerales.razonsocial", "", 0, 0)
            If Trim(ClienteApellido) = "" Or Trim(ClienteNombre) = "" Then ClienteApellido = ClienteRazon
            If Trim(ClienteRazon) = "" Then ClienteRazon = ClienteApellido & " " & ClienteNombre
            ClienteLocalidad = afipDatos.p1LeerPropiedad("p1getPersona", "datosgenerales.domicilioFiscal.localidad", "", 0, 0)
            ClienteDireccion = afipDatos.p1LeerPropiedad("p1getPersona", "datosgenerales.domicilioFiscal.direccion", "", 0, 0)
            ClienteTipoDomicilio = afipDatos.p1LeerPropiedad("p1getPersona", "datosgenerales.domicilioFiscal.tipodomicilio", "", 0, 0)
            ClienteCodPostal = afipDatos.p1LeerPropiedad("p1getPersona", "datosgenerales.domicilioFiscal.codPostal", "", 0, 0)
            ClienteProvincia = afipDatos.p1LeerPropiedad("p1getPersona", "datosgenerales.domicilioFiscal.descripcionProvincia", "", 0, 0)
            
            If afipDatos.p1VerificarImpuesto(20, "activo") Then
                ClienteMonotributo = "Monotributo-Si"
            Else
                ClienteMonotributo = "Monotributo-No"
            End If
            If afipDatos.p1VerificarImpuesto(30, "activo") Then
                ClienteInscripto = "Inscripto-Si"
            Else
                ClienteInscripto = "Inscripto-No"
            End If
            If afipDatos.p1VerificarImpuesto(32, "activo") Then
                ClienteExento = "Exento-Si"
            Else
                ClienteExento = "Exento-No"
            End If
            ReDim DatosResultado(11)
            DatosResultado(0) = ClienteApellido
            DatosResultado(1) = ClienteNombre
            DatosResultado(2) = ClienteRazon
            DatosResultado(3) = ClienteLocalidad
            DatosResultado(4) = ClienteDireccion
            DatosResultado(5) = ClienteTipoDomicilio
            DatosResultado(6) = ClienteCodPostal
            DatosResultado(7) = ClienteProvincia
            DatosResultado(8) = ClienteMonotributo
            DatosResultado(9) = ClienteInscripto
            DatosResultado(10) = ClienteExento
            
        Else
            MsgBox ("error general" + afipDatos.UltimoMensajeError)
            DatosResultado = Array("", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
        End If
    Else
        MsgBox ("fallo acceso " + afipDatos.UltimoMensajeError)
        DatosResultado = Array("", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
    End If
Else
    MsgBox ("fallo iniciar " + afipDatos.UltimoMensajeError)
    DatosResultado = Array("", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
End If
ConsultarDatos = DatosResultado
End Function


Public Function CodigoIvaGranEmpresa() As Integer
Dim ws As New WSAFIPFE.misdatos

Dim bResultado As Integer, bResultadoIva As Integer

bResultadoIva = nSinNull(obtenerDeSQL("select codigo from ivas where descripcion='INSCRIPTO GRAN EMPRESA'"))
If bResultadoIva = 0 Then bResultadoIva = 9
CodigoIvaGranEmpresa = bResultadoIva


End Function

Public Function EstaAprobada(IDFactura As Long) As Boolean
Dim sTipo As String
Dim nFechaFactura As Date
Dim nDias As Long
sTipo = Trim(obtenerDeSQL("select tipodoc from facturaventa where codigo=" & IDFactura))
EstaAprobada = False
If sTipo = "FEA" Then
    nFechaFactura = obtenerDeSQL("select fecha from facturaventa where codigo=" & IDFactura)
    nDias = Date - nFechaFactura
    'If nDias >= 15 Then EstaAprobada = True
End If
End Function

Public Function SaldoFacturaCredito(IDFactura As Long) As Double
Dim sTipo As String
Dim nTotalFactura As Double
Dim nTotalCredito As Double
sTipo = Trim(obtenerDeSQL("select tipodoc from facturaventa where codigo=" & IDFactura))
If sTipo = "FEA" Then
    nTotalFactura = s2n(obtenerDeSQL("select sum(total) as facturas from facturaventa where codigo=" & IDFactura))
    nTotalCredito = s2n(obtenerDeSQL("select sum(total) as notas from facturaventa where codfactura=" & IDFactura))
    SaldoFacturaCredito = nTotalFactura - nTotalCredito
Else
    MsgBox "El comprobante no es una factura de credito.", vbCritical
    SaldoFacturaCredito = 0
End If
End Function

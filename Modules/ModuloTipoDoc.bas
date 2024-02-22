Attribute VB_Name = "ModuloTipoDoc"
Option Explicit

Public Const TIPODOC_FAC_PROVEEDOR = "FAC"
Public Const TIPODOC_NC_PROVEEDOR = "N/C"
Public Const TIPODOC_ND_PROVEEDOR = "N/D"
Public Const TIPODOC_FAC_PROVGASTO = "FCG"
Public Const TIPODOC_FAC_BANCOGASTO = "BG"
Public Const TIPODOC_FAC_BOLETA = "BOL"

Public Const TipoDoc_FACTURA_A = "FAA"
Public Const TipoDoc_FACTURA_B = "FAB"
Public Const TipoDoc_FACTURA_C = "FAC"
Public Const TipoDoc_FACTURA_E = "FAE"
Public Const TipoDoc_FACTURA_CREDITO_A = "FEA"
Public Const TipoDoc_FACTURA_CREDITO_B = "FEB"
Public Const TipoDoc_FACTURA_CREDITO_C = "FEC"

Public Const TipoDoc_NCREDITO_A = "NCA"
Public Const TipoDoc_NCREDITO_B = "NCB"
Public Const TipoDoc_NCREDITO_C = "NCC"
Public Const TipoDoc_CREDITO_ELECTRONICO_A = "CEA"
Public Const TipoDoc_CREDITO_ELECTRONICO_B = "CEB"
Public Const TipoDoc_CREDITO_ELECTRONICO_C = "CEC"

Public Const TipoDoc_NDEBITO_A = "NDA"
Public Const TipoDoc_NDEBITO_B = "NDB"
Public Const TipoDoc_NDEBITO_C = "NDC"
Public Const TipoDoc_DEBITO_ELECTRONICO_A = "DEA"
Public Const TipoDoc_DEBITO_ELECTRONICO_B = "DEB"
Public Const TipoDoc_DEBITO_ELECTRONICO_C = "DEC"


Public Const TipoDoc_NCREDITO_E = "NCE"
Public Const TipoDoc_NDEBITO_E = "NDE"
Public Const TipoDoc_TICKET = "TIC"

Public Const TipoDoc_RECIBO = "RAA"     'tabla facturaVenta, en Recibos no hace falta
Public Const TipoDoc_AJ_CREDITO = "ACC"
Public Const TipoDoc_AJ_DEBITO = "ACD"
'


Public Function BorroDocumento(iddoc As Long) As Boolean
'    If ON_ERROR_HABILITADO Then On Error GoTo ufaTdocb  *********** 'NOOOO ***********
    Dim re As Long
    'Dim ss As String
    
    re = s2n(obtenerDeSQL("select iddoc from RegistroDocumentos where activo = 1 and iddoc = '" & iddoc & "' "))
    'If re = 0 Then Err.Raise 55000, "RegistroDocumentos", "no se pudo borrar idDoc"
    
    
    DataEnvironment1.Sistema.Execute _
            " update RegistroDocumentos " & _
            " set activo = 0, usuario_baja = " & UsuarioActual() & " , fecha_baja = " & ssFecha(Date) & _
            " where iddoc = '" & iddoc & "' "
    AsientoBaja_idDoc iddoc
    
    BorroDocumento = True
End Function


Public Function NuevoDocumento(TIPODOC As String, NroDoc As Long, CodProveedor As Long, NroPago As Long, Optional NroCertifGan As Long = 0, Optional NroCertifIIBB As Long = 0, Optional PuntoVenta As String = "0001") As Long
'    if on_error_habilitado then On Error GoTo ufaTdoc    DEBE MANEJARLO EL Q LO LLAMA
    Dim rs As New ADODB.Recordset, re As Long
    'esto repite iddoc
    're = s2n(obtenerDeSQL("select iddoc from RegistroDocumentos where activo = 1 and TipoDoc = '" & TIPODOC & "' and NroDoc = " & NroDoc & " and codproveedor = " & CodProveedor))
   ' re = s2n(obtenerDeSQL("select iddoc from RegistroDocumentos where TipoDoc = '" & TIPODOC & "' and NroDoc = " & NroDoc & " and codproveedor = " & CodProveedor))
   re = s2n(obtenerDeSQL("select iddoc from RegistroDocumentos where activo = 1 and TipoDoc = '" & TIPODOC & "' and NroDoc = " & NroDoc & " and codproveedor = " & CodProveedor & " and puntoventa=" & ssTexto(PuntoVenta)))
    If re > 0 Then
        ufa "Err: codigo ya existe - NuevoDocumento()", "Nuevo TipoDoc: " & TIPODOC & " nro: " & NroDoc & " prov: " & CodProveedor & " id= " & re
        NuevoDocumento = 0
        Err.Raise 30001, "NuevoDocumento", "Nro y Tipo doc ya existe"
    Else
        With rs
            .Open "select * from RegistroDocumentos", DataEnvironment1.Sistema, adOpenKeyset, adLockOptimistic
            .AddNew
                !TIPODOC = Left(TIPODOC, 3)
                !NroDoc = NroDoc
                !CodProveedor = CodProveedor
                !NumeroDePago = NroPago
                !NroCertifGan = NroCertifGan
                !NroCertifIIBB = NroCertifIIBB
                !fecha_alta = Date
                '!fecha_baja =
                !usuario_alta = UsuarioActual()
            .Update
            NuevoDocumento = !iddoc
        End With
    End If
fin:
    Set rs = Nothing
    Exit Function
ufaTdoc:
    NuevoDocumento = 0
    ufa "err al grabar", "Nuevo TipoDoc: " & TIPODOC & " nro: " & NroDoc & " prov: " & CodProveedor
    Resume fin
End Function



Public Function NuevoNroPago() As Long
    NuevoNroPago = 1 + nSinNull(obtenerDeSQL("select max(NumeroDePago) from RegistroDocumentos"))
End Function
Public Function NuevoNroCertifIIBB() As Long
    NuevoNroCertifIIBB = 1 + obtenerDeSQL("select max(NroCertifIIBB) from RegistroDocumentos ")
End Function
Public Function NuevoNroCertifGan() As Long
    NuevoNroCertifGan = 1 + obtenerDeSQL("select max(NroCertifGan) from RegistroDocumentos")
End Function

Public Function VerIdDoc(TIPODOC As String, NroDoc As Long, CodProv As Long) ' sucursal as long)

    ' deberia verificar
    '       1) si tipodoc esta registrado
    '       2) si es documento de proveedor, que tenga cod proveedor (y sucusal)
    'Select Case tipoDoc
    'Case ""
    'End Select
    
    VerIdDoc = obtenerDeSQL("select iddoc from RegistroDocumentos where tipodoc = '" & TIPODOC & "' and nrodoc = '" & NroDoc & "' and coProveedor = '" & CodProv & "' ")
End Function

Public Function VerNroPago(iddoc As Long) As Long
    VerNroPago = obtenerDeSQL("select NumeroDePago from RegistroDocumentos where iddoc = " & iddoc)
End Function
Public Function VerNroCertifIIBB(iddoc As Long) As Long
    VerNroCertifIIBB = obtenerDeSQL("select NroCertifIIBB from RegistroDocumentos where iddoc = " & iddoc)
End Function
Public Function VerNroCertifGan(iddoc As Long) As Long
    VerNroCertifGan = obtenerDeSQL("select NroCertifGan from RegistroDocumentos where iddoc = " & iddoc)
End Function


Public Function rsAsiento(iddoc As Long) ' nunca la probe...
    Dim rs As New ADODB.Recordset
    Dim ss As String
        
    ss = " SELECT r.idDoc, r.TipoDoc, r.NroDoc, a.idAsiento, a.NroAsiento, a.Concepto, a.Fecha, m.Cuenta, " & _
         " m.Debe, m.Haber, m.comprobante, c.DESCRIPCION " & _
         " FROM (dbo_CUENTAS AS c " & _
         " INNER JOIN (dbo_Asientos AS a " & _
         " INNER JOIN dbo_MAYOR AS m ON a.idAsiento = m.idAsiento) ON c.Cuenta = m.Cuenta) " & _
         " INNER JOIN dbo_RegistroDocumentos AS r ON a.idDoc = r.idDoc " & _
         " Where r.Activo = 1 "

    rs.Open ss, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    
    Set rsAsiento = rs
End Function

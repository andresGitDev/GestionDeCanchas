Attribute VB_Name = "ModEmisionComprobantes"
Option Explicit

Public cierra As Boolean

Public Const SEL_IMP_PREDETERMINADO = "Predeterminada"

Private sTablaTemp As String
'Private Const CodTomKa = 366591
Public Const tt_Etiquetas_temp = _
    "([ProdCliente] [varchar] (50)," & _
    "[Letra] [varchar] (3)," & _
    "[Cantidad] [float]," & _
    "[descripcion] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
    "[Destino] [varchar] (50), " & _
    "[Remito] [float] , " & _
    "[CodBarra] [varchar] (50)) "

'tabla formas pago
Public Const FormaPago_CONTADO = 1
'
'tabla MoviCaja
Public Const MoviCaja_EFECTIVO = "E"
Public Const MoviCaja_CHEQUE = "C"  'cheque terceros
Public Const MoviCaja_P = "P"       'cheque propio
'
Public Const MoviCaja_INGRESO = "I"
Public Const MoviCaja_EGRESO = "E"
'
Private sTablaRemito As String

'Public Type utCheque
'    Codigo      As Integer
'    Numero      As String
'    CodBanco    As Integer
'    Banco       As String
'    Monto       As Double
'    fecha       As Date
'    PT          As String * 1
'End Type
'Public Type utMovEfectivo
'    Monto       As Double
'    Caja        As Integer
'    Cuenta      As String
'End Type

Public Enum ItemFRC
    ItemFRC_Alta
    ItemFRC_BajaFactura
    ItemFRC_BajaRemito
    ItemFRC_Mod
End Enum

Public Function enletras(num As String) As String

    Dim b, paso As Long
    Dim numero As String

    Dim expresion, entero, Deci, flag As String
    numero = Replace(num, ",", ".")
    flag = "N"

    For paso = 1 To Len(numero)

        If Mid(numero, paso, 1) = "." Then

            flag = "S"

        Else

            If flag = "N" Then

                entero = entero + Mid(numero, paso, 1) 'Extae la parte entera del numero

            Else

                Deci = Deci + Mid(numero, paso, 1) 'Extrae la parte decimal del numero

            End If

        End If

    Next paso
    If Len(Deci) = 1 Then

        Deci = Deci & "0"

    End If
    flag = "N"

    If val(numero) >= -999999999 And val(numero) <= 999999999 Then 'si el numero esta dentro de 0 a 999.999.999

        For paso = Len(entero) To 1 Step -1

            b = Len(entero) - (paso - 1)

            Select Case paso

            Case 3, 6, 9

                Select Case Mid(entero, b, 1)

                    Case "1"

                        If Mid(entero, b + 1, 1) = "0" And Mid(entero, b + 2, 1) = "0" Then

                            expresion = expresion & "cien "

                        Else

                            expresion = expresion & "ciento "
                            flag = "N"
                        End If

                    Case "2"

                        expresion = expresion & "doscientos "

                    Case "3"

                        expresion = expresion & "trescientos "

                    Case "4"

                        expresion = expresion & "cuatrocientos "

                    Case "5"

                        expresion = expresion & "quinientos "

                    Case "6"

                        expresion = expresion & "seiscientos "

                    Case "7"

                        expresion = expresion & "setecientos "

                    Case "8"

                        expresion = expresion & "ochocientos "

                    Case "9"

                        expresion = expresion & "novecientos "

                End Select
            Case 2, 5, 8

                Select Case Mid(entero, b, 1)

                    Case "1"

                        If Mid(entero, b + 1, 1) = "0" Then

                            flag = "S"

                            expresion = expresion & "diez "

                        End If

                        If Mid(entero, b + 1, 1) = "1" Then

                            flag = "S"

                            expresion = expresion & "once "

                        End If

                        If Mid(entero, b + 1, 1) = "2" Then

                            flag = "S"

                            expresion = expresion & "doce "

                        End If

                        If Mid(entero, b + 1, 1) = "3" Then

                            flag = "S"

                            expresion = expresion & "trece "

                        End If

                        If Mid(entero, b + 1, 1) = "4" Then

                            flag = "S"

                            expresion = expresion & "catorce "

                        End If

                        If Mid(entero, b + 1, 1) = "5" Then

                            flag = "S"

                            expresion = expresion & "quince "

                        End If

                        If Mid(entero, b + 1, 1) > "5" Then

                            flag = "N"

                            expresion = expresion & "dieci"

                        End If
                    Case "2"

                        If Mid(entero, b + 1, 1) = "0" Then

                            expresion = expresion & "veinte "

                            flag = "S"

                        Else

                            expresion = expresion & "veinti"

                            flag = "N"

                        End If
                    Case "3"
                        If Mid(entero, b + 1, 1) = "0" Then

                            expresion = expresion & "treinta "

                            flag = "S"

                        Else

                            expresion = expresion & "treinta y "

                            flag = "N"

                        End If
                    Case "4"

                        If Mid(entero, b + 1, 1) = "0" Then

                            expresion = expresion & "cuarenta "

                            flag = "S"

                        Else

                            expresion = expresion & "cuarenta y "

                            flag = "N"

                        End If
                    Case "5"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "cincuenta "
                            flag = "S"
                        Else
                            expresion = expresion & "cincuenta y "
                            flag = "N"
                        End If
                    Case "6"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "sesenta "
                            flag = "S"
                        Else
                            expresion = expresion & "sesenta y "
                            flag = "N"
                        End If
                    Case "7"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "setenta "
                            flag = "S"
                        Else
                            expresion = expresion & "setenta y "
                            flag = "N"
                        End If
                    Case "8"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "ochenta "
                            flag = "S"
                        Else
                            expresion = expresion & "ochenta y "
                            flag = "N"
                        End If
                    Case "9"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "noventa "
                            flag = "S"
                        Else
                            expresion = expresion & "noventa y "
                            flag = "N"
                        End If
                End Select
            Case 1, 4, 7
                Select Case Mid(entero, b, 1)
                    Case "1"
                        If flag = "N" Then
                            If paso = 1 Then
                                expresion = expresion & "uno "
                            Else
                                expresion = expresion & "un "
                            End If
                        End If
                    Case "2"
                        If flag = "N" Then
                            expresion = expresion & "dos "
                        End If
                    Case "3"
                        If flag = "N" Then
                            expresion = expresion & "tres "
                        End If
                    Case "4"
                        If flag = "N" Then
                            expresion = expresion & "cuatro "
                        End If
                    Case "5"
                        If flag = "N" Then
                            expresion = expresion & "cinco "
                        End If
                    Case "6"
                        If flag = "N" Then
                            expresion = expresion & "seis "
                        End If
                    Case "7"
                        If flag = "N" Then
                            expresion = expresion & "siete "
                        End If
                    Case "8"
                        If flag = "N" Then
                            expresion = expresion & "ocho "
                        End If
                    Case "9"
                        If flag = "N" Then
                            expresion = expresion & "nueve "
                        End If
                End Select
            End Select
            If paso = 4 Then
                'If (Mid(entero, 6, 1) <> "0" Or Mid(entero, 5, 1) <> "0" Or Mid(entero, 4, 1) <> "0" Or (Mid(entero, 6, 1) = "0") And Mid(entero, 5, 1) = "0" And Mid(entero, 4, 1) = "0" And Len(entero) <= 6) Then
                    expresion = expresion & "mil "
                'End If
            End If
            If paso = 7 Then
                If Len(entero) = 7 And Mid(entero, 1, 1) = "1" Then
                    expresion = expresion & "millon "
                Else
                    expresion = expresion & "millones "
                End If
            End If
        Next paso
        If Deci <> "" Then
            If Mid(entero, 1, 1) = "-" Then 'si el numero es negativo
                enletras = "menos " & UCase(expresion & "con " & Deci & "/100")
            Else
                enletras = UCase(expresion & "con " & Deci & "/100")
            End If
        Else
            If Mid(entero, 1, 1) = "-" Then 'si el numero es negativo
                enletras = "menos " & UCase(expresion)
            Else
                enletras = UCase(expresion)
            End If
        End If
    Else 'si el numero a convertir esta fuera del rango superior e inferior
        enletras = ""
    End If

End Function
Public Function ImprimirIngresoCaja(Nmov As Long, Tipo As String, Optional iddoc As Long)
Dim sql As String
 sql = "SELECT MoviCaja.CAJA, MoviCaja.IMPORTE, " & _
    " MAYOR.Cuenta, CUENTAS.DESCRIPCION, MoviCaja.Movimiento, Mayor.haber,mayor.debe, " & _
    " MoviCaja.CONCEPTO, Cajas.sector, MoviCaja.TIPO, MoviCaja.idDoc " & _
    " FROM Cajas INNER JOIN (((MoviCaja INNER JOIN Asientos ON MoviCaja.idDoc = Asientos.idDoc)" & _
    " INNER JOIN MAYOR ON Asientos.idAsiento = MAYOR.idAsiento) " & _
    " INNER JOIN CUENTAS ON MAYOR.Cuenta = CUENTAS.Cuenta) ON Cajas.codigo = MoviCaja.CAJA "

Select Case Tipo
Case "I"
    sql = sql + "WHERE MoviCaja.MOVIMIENTO = " & Trim(Nmov) & " AND MAYOR.HABER > 0"
    RptIngresoCaja.Fecha.Text = FrmIngEgrEfectivo.dtFecha
    RptIngresoCaja.txtimporte.Text = enletras(FrmIngEgrEfectivo.txtimporte)
    RptIngresoCaja.DataIngCaja.Connection = DataEnvironment1.Sistema
    RptIngresoCaja.DataIngCaja.Source = sql
    RptIngresoCaja.Restart
    RptIngresoCaja.Show vbModal
Case "E"
    sql = sql + "WHERE MoviCaja.MOVIMIENTO = " & Trim(Nmov) & " AND MAYOR.debe > 0"
    RptEgresoCaja.Fecha.Text = FrmIngEgrEfectivo.dtFecha
    RptEgresoCaja.txtimporte.Text = enletras(FrmIngEgrEfectivo.txtimporte)
    RptEgresoCaja.Txtop.Text = VerNroPago(iddoc)
    RptEgresoCaja.DataEgresoCaja.Connection = DataEnvironment1.Sistema
    RptEgresoCaja.DataEgresoCaja.Source = sql
    RptEgresoCaja.Restart
    RptEgresoCaja.Show vbModal
Case "D"
    sql = "SELECT MoviCaja.CAJA, MoviCaja.Concepto, MoviCaja.IMPORTE, MAYOR.Cuenta, CUENTAS.DESCRIPCION, " & _
    " MoviCaja.MOVIMIENTO, MAYOR.Debe, MoviCaja.CONCEPTO, Cajas.sector, CTASBANK.NUMERO,CTASBANK.Banco, " & _
    " BancosGrales.descripcion ,BancosGrales.codigo,MoviCaja.TIPO" & _
    "" & _
    " FROM BancosGrales INNER JOIN (CTASBANK INNER JOIN (Cajas INNER JOIN (((MoviCaja INNER JOIN Asientos  " & _
    " ON MoviCaja.idDoc = Asientos.idDoc) INNER JOIN MAYOR ON Asientos.idAsiento = MAYOR.idAsiento)  " & _
    " INNER JOIN CUENTAS ON MAYOR.Cuenta = CUENTAS.Cuenta) ON Cajas.codigo = MoviCaja.CAJA) ON " & _
    " CTASBANK.CUENTA_CON = CUENTAS.Cuenta) ON BancosGrales.codigo = CTASBANK.CODIGO " & _
    " WHERE MoviCaja.MOVIMIENTO = " & Trim(Nmov) & " AND MAYOR.debe > 0"

    RptDepositoCaja.Fecha.Text = FrmIngEgrEfectivo.dtFecha
    RptDepositoCaja.txtimporte.Text = enletras(FrmIngEgrEfectivo.txtimporte)
    RptDepositoCaja.DataDepositoCaja.Connection = DataEnvironment1.Sistema
    RptDepositoCaja.DataDepositoCaja.Source = sql
    RptDepositoCaja.Restart
    RptDepositoCaja.Show vbModal
End Select
End Function
Public Function Imprimir(MovBanco As Long, Tipo As String)
Dim sql As String

Select Case Tipo
Case "CB"
    sql = " SELECT MAYOR.Cuenta, MAYOR.Debe, MAYOR.Haber, MAYOR.idAsiento, MAYOR.Cuenta, CUENTAS.DESCRIPCION, MoviBanc.MOVBANCO" & _
    " FROM MoviBanc INNER JOIN (Asientos INNER JOIN (MAYOR INNER JOIN CUENTAS ON MAYOR.Cuenta = CUENTAS.Cuenta) ON Asientos.idAsiento = MAYOR.idAsiento) ON MoviBanc.idDoc = Asientos.idDoc " & _
    " WHERE (((MAYOR.Haber)>0) AND ((MoviBanc.MOVBANCO)=" & MovBanco & "));"

    RptCreditoBancario.Fecha.Text = FrmGastosBancarios.dtFecha
    RptCreditoBancario.txtconcepto = FrmGastosBancarios.txtconcepto
    RptCreditoBancario.txtnumcta = FrmGastosBancarios.txtnumcta
    RptCreditoBancario.txtbanco = FrmGastosBancarios.uCtaBanco.DESCRIPCION
    RptCreditoBancario.txtimporte = FrmGastosBancarios.txtimporte
    RptCreditoBancario.TxtImporte1 = FrmGastosBancarios.txtimporte
    RptCreditoBancario.txtMovBanc = FrmGastosBancarios.txtMovBanc
    RptCreditoBancario.DataCreditoBank.Connection = DataEnvironment1.Sistema
    RptCreditoBancario.DataCreditoBank.Source = sql
    RptCreditoBancario.Restart
    RptCreditoBancario.Show vbModal

Case "GB"
    sql = " SELECT MAYOR.Cuenta, MAYOR.Debe, MAYOR.Haber, MAYOR.idAsiento, MAYOR.Cuenta, CUENTAS.DESCRIPCION, MoviBanc.MOVBANCO " & _
    " FROM MoviBanc INNER JOIN (Asientos INNER JOIN (MAYOR INNER JOIN CUENTAS ON MAYOR.Cuenta = CUENTAS.Cuenta) ON Asientos.idAsiento = MAYOR.idAsiento) ON MoviBanc.idDoc = Asientos.idDoc " & _
    " WHERE (((MAYOR.debe)>0) AND ((MoviBanc.MOVBANCO)= " & MovBanco & ")); "
    
    rptGastoBank.Fecha.Text = FrmGastosBancarios.dtFecha
    rptGastoBank.txtconcepto = FrmGastosBancarios.txtconcepto
    rptGastoBank.txtnumcta = FrmGastosBancarios.txtnumcta
    rptGastoBank.txtbanco = FrmGastosBancarios.uCtaBanco.DESCRIPCION
    rptGastoBank.txtimporte = FrmGastosBancarios.txtimporte
    rptGastoBank.TxtImporte1 = FrmGastosBancarios.txtimporte
    rptGastoBank.txtMovBanc = FrmGastosBancarios.txtMovBanc
    rptGastoBank.DataGastoBank.Connection = DataEnvironment1.Sistema
    rptGastoBank.DataGastoBank.Source = sql
    rptGastoBank.Restart
    rptGastoBank.Show vbModal
    
End Select
End Function
Public Function ImprimirTransferenciaBanc(iddoc)
Dim sql As String
Dim rs As New ADODB.Recordset

If FrmTransfBanc.txtcuentao = "" And FrmTransfBanc.txtcuentad = "" Then Exit Function
    RptTransfBancaria.txtMovBanc = FrmTransfBanc.movbanc
    RptTransfBancaria.Fecha = FrmTransfBanc.dtFecha
    RptTransfBancaria.txtconcepto = FrmTransfBanc.txtconcepto
    If FrmTransfBanc.txtcodctao = "" Or FrmTransfBanc.txtcodctao = "0" Then
        RptTransfBancaria.txtcuentao = "0"
        RptTransfBancaria.TxtBancoOrigen = "  Sin Asignar"
    Else
        RptTransfBancaria.txtcuentao = obtenerDeSQL("select numero from ctasbank where codigo = " & FrmTransfBanc.txtcodctao)
        'Right(FrmTransfBanc.txtcuentao, (Len(FrmTransfBanc.txtcuentao) - InStr(1, FrmTransfBanc.txtcuentao, "-") - 1))
        RptTransfBancaria.TxtBancoOrigen = obtenerDeSQL("SELECT b.descripcion FROM BancosGrales b INNER JOIN CTASBANK c ON c.BANCO = b.codigo WHERE c.CODIGO = " & FrmTransfBanc.txtcodctao)
    End If
    'Left(FrmTransfBanc.txtcuentao, InStr(1, FrmTransfBanc.txtcuentao, "-") - 1)
    RptTransfBancaria.txtimporte = NroEnLetras(s2n(FrmTransfBanc.txtimporte))
    RptTransfBancaria.TxtImporte1 = FrmTransfBanc.txtimporte

If FrmTransfBanc.txtcuentad <> "" Then
    If FrmTransfBanc.txtcodctad.Text = "" Or FrmTransfBanc.txtcodctad.Text = "0" Then
        RptTransfBancaria.TxtCuentaDestino = "0"
        RptTransfBancaria.TxtBancoDestino = "    Sin Asignar"
    Else
        RptTransfBancaria.TxtCuentaDestino = obtenerDeSQL("select numero from ctasbank where codigo = " & FrmTransfBanc.txtcodctad) 'Left(FrmTransfBanc.txtcuentad, InStr(1, FrmTransfBanc.txtcuentad, "-") - 1)
        RptTransfBancaria.TxtBancoDestino = obtenerDeSQL("SELECT b.descripcion FROM BancosGrales b INNER JOIN CTASBANK c ON c.BANCO = b.codigo WHERE c.CODIGO = " & FrmTransfBanc.txtcodctad)
    End If
'    RptTransfBancaria.Txtdescripcion.Visible = False
'    RptTransfBancaria.TxtCuenta.Visible = False
'    RptTransfBancaria.TxtDebe.Visible = False
End If
If (FrmTransfBanc.txtcodctao <> "" And (FrmTransfBanc.txtcodctad = "" Or FrmTransfBanc.txtcodctad = "0")) Or (FrmTransfBanc.txtcodctad <> "" And (FrmTransfBanc.txtcodctao = "" Or FrmTransfBanc.txtcodctao = "0")) Then
    RptTransfBancaria.txtDescripcion.Visible = True
    RptTransfBancaria.txtcuenta.Visible = True
    RptTransfBancaria.TxtDebe.Visible = True
Else
    RptTransfBancaria.txtDescripcion.Visible = False
    RptTransfBancaria.txtcuenta.Visible = False
    RptTransfBancaria.TxtDebe.Visible = False
End If

If (FrmTransfBanc.txtcuentao <> "" And FrmTransfBanc.txtcuentad = "") Then
   RptTransfBancaria.Label9.Visible = False
   RptTransfBancaria.TxtCuentaDestino.Visible = False
   RptTransfBancaria.TxtBancoDestino.Visible = False
   RptTransfBancaria.Label8.Visible = False
End If
   
   sql = "SELECT a.idDoc, a.NroAsiento, a.Concepto, a.Fecha, a.idAsiento, a.activo," & _
    " m.cuenta , m.debe, m.haber, m.comprobante, c.descripcion " & _
    " FROM CUENTAS AS c INNER JOIN (Asientos AS a INNER JOIN MAYOR AS m ON a.idAsiento = m.idAsiento) ON c.Cuenta = m.Cuenta " & _
    " Where a.iddoc = " & iddoc & " And debe > 0 ORDER BY a.NroAsiento, m.Cuenta DESC; "

    RptTransfBancaria.Data.Source = sql
    RptTransfBancaria.Data.Connection = DataEnvironment1.Sistema
    RptTransfBancaria.Restart
    RptTransfBancaria.Show vbModal

End Function
Public Function ImprimirPedido(NPedido As Double) As Boolean
Dim rs As New ADODB.Recordset
Dim sql As String

    
    Dim Propio As Boolean, tempo As Variant, cliente As Long
    
    tempo = obtenerDeSQL("select codigopropio, cliente from pedidos_clientes where numero = " & NPedido)
    Propio = tempo(0)
    cliente = tempo(1)
    
    If Propio Then
       sql = "SELECT i.*, i.producto as produ, Producto.descripcion FROM ItemPedidoCliente  as i INNER JOIN Producto ON i.producto = Producto.codigo WHERE i.pedido = " & NPedido & ""
    Else
       sql = "SELECT i.*, p.descripcion, r.productocliente as produ " _
        & " FROM ItemPedidoCliente as i " _
        & " left JOIN Producto as p ON i.producto = p.codigo " _
        & " left join relacion_producto_cliente as r on i.producto = r.producto " _
        & " WHERE i.pedido = " & NPedido & " and cliente = " & cliente
    End If
    sql = sql & " Order by i.codigo "
    
    
   rs.Open sql, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
   If Not rs.EOF Then
    RptPedidoCliente.NroCliente.Text = FrmPedidosClientes.uCliente.codigo
    RptPedidoCliente.NomCliente.Text = FrmPedidosClientes.uCliente.DESCRIPCION
    RptPedidoCliente.Fecha.Text = FrmPedidosClientes.dtFecha
    RptPedidoCliente.NroPedido.Text = FrmPedidosClientes.txtNro
    RptPedidoCliente.Total.Text = FrmPedidosClientes.lblTotalPedi
    RptPedidoCliente.Vendedor.Text = FrmPedidosClientes.cmbvendedor
    RptPedidoCliente.PedCli.Text = FrmPedidosClientes.txtnropedidocli
    RptPedidoCliente.Obser.Text = FrmPedidosClientes.txtObs
    'RptPedidoCliente.FechaEnt.Text = FrmPedidosClientes.dtfechaentrega
    RptPedidoCliente.FechaEnt.Text = FrmPedidosClientes.grillaproductos.TextMatrix(1, 4)
    RptPedidoCliente.FPago.Text = FrmPedidosClientes.cmbformapago.Text
    
   End If
rs.Close
RptPedidoCliente.DControl.Connection = DataEnvironment1.Sistema
RptPedidoCliente.DControl.Source = sql
' CANTIDAD DE COPIAS A IMPRIMIR FALTA DEFINIR
'RptPedidoCliente.Printer.Copies = 3
'
RptPedidoCliente.Show vbModal

End Function

'''Public Function ImprimirRemitoVenta(codigo) As Boolean
'''
'''    Dim rs As New ADODB.Recordset
'''
'''    Dim rs2 As New ADODB.Recordset
'''    Dim rs3 As New ADODB.Recordset
'''    Dim rs4 As New ADODB.Recordset
'''    Dim CantBulto, AuxNPed As Double
'''
'''    Dim str, str2, NPedidos As String
'''    Dim cod As Long
'''
'''    Dim Propio As Boolean
'''
''' If ON_ERROR_HABILITADO Then On Error GoTo ErrImpresora
'''
'''Call setLpt2
'''
'''    rs.Open "SELECT RemitoVenta.*,remitoventa.transporte as trans, Clientes.*, Ivas.descripcion as iva" _
'''    & " FROM Ivas INNER JOIN (Clientes INNER JOIN RemitoVenta ON Clientes.codigo = RemitoVenta.Cliente) ON Ivas.codigo = clientes.IVA  where numero=" & codigo, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
'''    If Not rs.EOF Then
'''        Propio = rs!codPropio
'''        rs2.Open "SELECT direccion,descripcion FROM Transportes WHERE codigo = " & rs!trans & "", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
'''        If Not rs2.EOF Then
'''            If rs2!direccion = "" Then
'''             RptImpresionRemitoVenta.DireccTrans = " "
'''            Else
'''             RptImpresionRemitoVenta.DireccTrans = rs2!direccion
'''            End If
'''        '    PONER EN LUGAR DEL RECORDSET EL VALOR DEL FORM
'''        '**************************************************************************
'''             RptImpresionRemitoVenta.CodTrans = rs!trans
'''             RptImpresionRemitoVenta.DescTrans = rs2!descripcion
'''
'''             RptImpresionRemitoVenta.Valor = frmRemitoVenta.lblTotalRV
'''            If IsNull(rs!obs1) = True Or rs!obs1 = "" Then
'''             RptImpresionRemitoVenta.Obser1 = "-"
'''            Else
'''             RptImpresionRemitoVenta.Obser1 = rs!obs1
'''            End If
'''            If IsNull(rs!OBS2) = True Or rs!OBS2 = "" Then
'''             RptImpresionRemitoVenta.Obser2 = "-"
'''            Else
'''             RptImpresionRemitoVenta.Obser2 = rs!OBS2
'''            End If
'''        End If
'''        rs2.Close
'''
'''        RptImpresionRemitoVenta.lblcliente = sSinNull(rs!nombrefantasia)
'''        If Not IsNull(rs!Cuit) Then
'''            RptImpresionRemitoVenta.lblcuit = rs!Cuit
'''        End If
'''        cod = rs!numero
'''        RptImpresionRemitoVenta.lblcomp = "Remito"
'''        If Propio Then
'''            str = "SELECT RemitoVentaDetalle.*, Producto.descripcion" _
'''                & " FROM RemitoVentaDetalle INNER JOIN Producto ON RemitoVentaDetalle.Producto = Producto.codigo" _
'''                & " where numero=" & cod
'''        Else
'''            str = "SELECT r.*, p.descripcion, rpc.ProductoCliente as producto " _
'''                & " FROM RemitoVentaDetalle as r left JOIN Producto as p ON r.Producto = p.codigo " _
'''                & " left join relacion_producto_cliente as rpc on rpc.producto = p.codigo " _
'''                & " where numero = " & cod & " and cliente = " & rs!cliente
'''        End If
'''        rs2.Open str, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
'''        If Not rs2.EOF Then
'''            RptImpresionRemitoVenta.PedPropio = rs2!PEDIDO
'''               rs4.Open "SELECT pedido_cli FROM pedidos_clientes WHERE numero = " & rs2!PEDIDO & "", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
'''            If Not rs4.EOF Then
'''               RptImpresionRemitoVenta.PedCli = rs4!pedido_cli
'''            End If
'''        End If
''''        Do While Not rs2.EOF
'''''         rs3.Open "SELECT formula FROM producto WHERE codigo = '" & rs2!producto & "'", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
'''''         If Not rs3.EOF And rs3!formula <> True Then
'''''            CantBulto = CantBulto + 1
'''''         End If
'''''         rs3.Close
''''        rs2.MoveNext
''''        Loop
'''        rs2.Close
'''        RptImpresionRemitoVenta.TotBultos = CantBulto
'''        If frmRemitoVenta.txtPedido <> "" Then
'''           AuxNPed = frmRemitoVenta.txtPedido
'''        Else
'''
'''        End If
'''        'rs2.Open "SELECT pedido_cli FROM pedidos_clientes WHERE numero = " & AuxNPed & ""
'''        'If Not rs2.EOF Then
'''  '
'''  '           RptImpresionRemitoVenta.PedPropio = AuxNPed
'''  '       End If
'''  '       rs2.Close
'''
'''        RptImpresionRemitoVenta.lblcliente = rs!descripcion
'''        RptImpresionRemitoVenta.lbldomicilio = rs!direccion
'''        RptImpresionRemitoVenta.lblfactura = "0001-" & Format(rs!numero, "00000000")
'''        RptImpresionRemitoVenta.lblfecha = rs!Fecha
'''        RptImpresionRemitoVenta.lbliva = rs!iva
'''        RptImpresionRemitoVenta.lbllocalidad = rs!localidad
'''
'''        LlenarTemp (str)
'''        RptImpresionRemitoVenta.DataControl1.Connection = DataEnvironment1.Sistema
'''
'''        ' CAMBIAR PARA QUE EL STR QUE FIGURA SE REDIRECCIONE A LA NUEVA TABLA
'''        ' QUE CREE
'''        str = "SELECT * FROM " & sTablaRemito & " "
'''        RptImpresionRemitoVenta.DataControl1.Source = str
'''
'''    End If
'''    rs.Close
'''    Set rs = Nothing
'''    RptImpresionRemitoVenta.Printer.Copies = 3
''''    RptImpresionRemitoVenta.PrintReport True
'''    RptImpresionRemitoVenta.Show
'''
'''Call setLpt1
'''fin:
'''    Exit Function
'''ErrImpresora:
'''    ufa "error de impresión Remito Venta", ""
'''    Call setLpt1
'''    Resume fin
'''End Function

Private Sub LlenarTemp(str As String)
   Dim rs As New ADODB.Recordset
   Dim rs2 As New ADODB.Recordset
   Dim AuxDescrip, AuxCodigo, AuxCant As String
  
  sTablaRemito = TablaTempCrear("([id] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ,   [cantidad] [varchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,   [codigo] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,   [descrip] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL) ON [PRIMARY]")
   
   rs.Open str, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
   Do While Not rs.EOF
      AuxCodigo = rs!producto
      AuxCant = x2s(rs!cantidad)
      AuxDescrip = rs!DESCRIPCION
      DataEnvironment1.Sistema.Execute "INSERT INTO " & sTablaRemito & " (cantidad,codigo,descrip) VALUES( '" & AuxCant & "','" & AuxCodigo & "','" & AuxDescrip & "')"
      rs2.Open "SELECT serie FROM series WHERE producto = '" & rs!producto & "' and nrocomprobante = '" & rs!numero & "'", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
      Do While Not rs2.EOF
      If Not rs2.EOF Then
         AuxCodigo = "Serie : "
         AuxCant = " "
         AuxDescrip = rs2!Serie
         DataEnvironment1.Sistema.Execute "INSERT INTO " & sTablaRemito & " (cantidad,codigo,descrip) VALUES( '" & AuxCant & "','" & AuxCodigo & "','" & AuxDescrip & "')"
      Else
         AuxCodigo = " "
         AuxCant = " "
         AuxDescrip = " "
         DataEnvironment1.Sistema.Execute "INSERT INTO " & sTablaRemito & "  (cantidad,codigo,descrip) VALUES('" & AuxCant & "','" & AuxCodigo & "','" & AuxDescrip & "')"
      End If
      rs2.MoveNext
      Loop
      rs2.Close
   rs.MoveNext
   Loop
   rs.Close
End Sub
Public Function CancelacionPedido(codigo)
Dim sql, Sql2 As String
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim CantBulto As Double

sql = "SELECT DISTINCT i.producto, p.descripcion, s.serie, i.*, i.numero, r.*" & _
"FROM Producto AS p INNER JOIN (RemitoDiferenciaStock AS r INNER JOIN " & _
"(Series AS s RIGHT JOIN ItemRemitoDiferenciaStock AS i ON " & _
"(s.nrocomprobante = i.numero) AND (s.producto = i.producto)) " & _
"ON r.MovimientoInterno = i.numero) ON p.codigo = i.producto WHERE " & _
"(((s.comprobante)=8 Or (s.comprobante) Is Null)) AND movimientointerno=" & codigo & " ORDER BY i.producto DESC"

RptCancPedido.CancPNro = frmPedidosCancelacion.lblNumero
RptCancPedido.PedOriginal = frmPedidosCancelacion.uPedido.codigo
RptCancPedido.ClienteNro = frmPedidosCancelacion.uCliente.codigo
RptCancPedido.Empresa = frmPedidosCancelacion.uCliente.DESCRIPCION
RptCancPedido.Fecha = frmPedidosCancelacion.uFecha.strFecha
RptCancPedido.DControl.Connection = DataEnvironment1.Sistema
RptCancPedido.DControl.Source = sql

Sql2 = "SELECT Usuarios.*, Pedidos_Clientes.numero,Pedidos_Clientes.transporte FROM Pedidos_Clientes INNER JOIN Usuarios ON Pedidos_Clientes.vendedor = Usuarios.codigo WHERE numero = " & frmPedidosCancelacion.uPedido.codigo & ""
rs.Open Sql2, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
rs2.Open "Select descripcion FROM transportes WHERE codigo = " & rs!Transporte & "", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
If Not rs.EOF Then
   RptCancPedido.Vendedor.Text = rs!codigo & " " & rs!DESCRIPCION
   RptCancPedido.Transporte.Text = rs!Transporte & " " & rs2!DESCRIPCION
End If
rs.Close
rs2.Close
rs.Open sql, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
Do While Not rs.EOF
   rs2.Open "SELECT formula FROM producto WHERE codigo = '" & rs!producto & "'", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
   If Not rs2.EOF And rs2!formula <> True Then
      CantBulto = CantBulto + 1
   End If
   rs2.Close
rs.MoveNext
Loop
RptCancPedido.CBulto.Text = CantBulto

RptCancPedido.Show
RptCancPedido.Printer.Copies = 2
End Function

'''Public Function ImprimirComprobante(codigo) As Boolean
'''    Dim str1 As String
'''    Dim rs As New ADODB.Recordset
'''    Dim rs1 As New ADODB.Recordset
'''    Dim rsnro As New ADODB.Recordset
'''    Dim str As String
'''    Dim cod As Long
'''    Dim PORCENTAJE As String
'''    Dim tdoc As String, z As Double, mone As Long
'''
'''    ' CodigoPropio: OJO!!!!
'''    ' se graba propio para cada item, pero aca lo busco una sola vez por performance,
'''    ' trae el primero que encuentra y asume que todos son propios o todos de cliente
'''    Dim Propio As Boolean
'''    Dim strDetalle As String
'''    '
'''
'''If ON_ERROR_HABILITADO Then On Error GoTo ErrImpresion
'''
'''Call setLpt2
'''
'''    str1 = "SELECT FacturaVenta.*, Clientes.*, FormasPago.descripcion as pago, " & _
'''    "Ivas.descripcion as iva, Ivas.letra as letra" & _
'''    " FROM FormasPago INNER JOIN (Ivas INNER JOIN (Clientes INNER JOIN FacturaVenta  " & _
'''    " ON Clientes.codigo = FacturaVenta.Cliente) " & _
'''    " ON Ivas.codigo = FacturaVenta.TipoIVA) " & _
'''    " ON FormasPago.codigo = FacturaVenta.FormaPago " & _
'''    " WHERE facturaventa.codigo=" & codigo
'''    rs.Open str1, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
'''
'''    tdoc = Trim(rs!TIPODOC)
'''    z = s2n(rs!cotizacion, 4)
'''    If z = 0 Then z = 1
'''    mone = rs!moneda
'''    If mone = 0 Then mone = 1 '(PESOS)
'''
'''    If Not rs.EOF Then
'''
'''        RptImpresionFacturaVenta.CodPos = rs!codigopostal
'''        If Not IsNull(rs!RAZONSOCIAL) Then
'''            RptImpresionFacturaVenta.lblcliente = rs!RAZONSOCIAL
'''        End If
'''        If Not IsNull(rs!Cuit) Then
'''            RptImpresionFacturaVenta.lblcuit = rs!Cuit
'''        End If
'''        cod = codigo
'''        PORCENTAJE = s2n(rs!PORCENTAJEiva)
'''
'''        If Left(tdoc, 2) = "FA" Then
'''
'''            RptImpresionFacturaVenta.lblcomp = "Factura"
'''
'''
'''            'detalle
'''            If tdoc = "FAB" Then
'''                strDetalle = " cantidad,descripcion,(preciounitario + (preciounitario * " & ssNum(PORCENTAJE) & ")) as punit,(preciototal + (preciototal * " & ssNum(PORCENTAJE) & ")) as ptot "
'''            ElseIf tdoc = "FAE" Then
'''                strDetalle = " cantidad,descripcion,preciounitario / " & ssNum(z) & " as punit ,preciototal / " & ssNum(z) & " as ptot "
'''            Else ' tdoc = "FAA" Then
'''                strDetalle = " cantidad,descripcion,preciounitario  as punit,preciototal  as ptot "
'''                ' solo Factura A
'''                RptImpresionFacturaVenta.txtivains = Format$(s2n(rs!neto) * s2n(rs!PORCENTAJEiva), "Standard")
'''                RptImpresionFacturaVenta.txtneto = Format$(rs!neto, "Standard")
'''                RptImpresionFacturaVenta.txtsub = Format$(rs!neto, "Standard")
'''            End If
'''
'''            Propio = obtenerDeSQL("select CodPropio from FacturaVentadetalle where codigofactura = " & cod)
'''            If Propio Then
'''                str = " select producto, " & _
'''                    strDetalle & _
'''                    " from facturaventadetalle " & _
'''                    " where codigofactura=" & cod & " ORDER BY id"
'''            Else
'''                str = " select productoCliente as producto, " & _
'''                    strDetalle & _
'''                    " from facturaventadetalle as d left join Relacion_Producto_Cliente as r " & _
'''                    " on d.producto = r.producto " & _
'''                    " where codigofactura=" & cod & " and r.cliente = " & rs!cliente & _
'''                    " ORDER BY id"
'''            End If
'''        Else
'''
'''            If Left(tdoc, 2) = "NC" Then
'''                RptImpresionFacturaVenta.lblcomp = "Nota de Credito"
'''                RptImpresionFacturaVenta.lbltachar = "XXXXXXXXX"
'''                If Propio Then
'''                    str = "select producto, cantidad, descripcion, (preciounitario + (preciounitario * " & Replace(PORCENTAJE, ",", ".") & ")) as punit,(preciototal + (preciototal * " & Replace(PORCENTAJE, ",", ".") & ")) as ptot from facturaventadetalle where codigofactura=" & cod
'''                Else
'''                    str = "select productocliente as producto, cantidad, descripcion, (preciounitario + (preciounitario * " & Replace(PORCENTAJE, ",", ".") & ")) as punit,(preciototal + (preciototal * " & Replace(PORCENTAJE, ",", ".") & ")) as ptot " & _
'''                        " from facturaventadetalle as d left join Relacion_Producto_Cliente as r " & _
'''                        " on d.producto = r.producto " & _
'''                        " where codigofactura=" & cod & " and r.cliente = " & rs!cliente & _
'''                        " ORDER BY id"
'''                End If
'''
'''                If tdoc = "NCB" Then
'''
'''                ElseIf tdoc = "NCE" Then
'''
'''                ElseIf tdoc = "NCA" Then
'''               '     str = "select producto,descripcion from facturaventadetalle where codigofactura=" & cod
'''                    RptImpresionFacturaVenta.txtivains = Format$(s2n(rs!neto) * s2n(rs!PORCENTAJEiva), "Standard")
'''                    RptImpresionFacturaVenta.txtneto = Format$(rs!neto, "Standard")
'''                    RptImpresionFacturaVenta.txtsub = Format$(rs!neto, "Standard")
'''                End If
'''            Else
'''                If Left(tdoc, 2) = "ND" Then
'''                    RptImpresionFacturaVenta.lblcomp = "Nota de Debito"
'''                    RptImpresionFacturaVenta.lbltachar = "XXXXXXXXX"
'''                    If Trim(rs!TIPODOC) = "NDB" Then
'''                        str = "select producto,descripcion from facturaventadetalle where codigofactura=" & cod
'''                    ElseIf tdoc = "NDE" Then
'''                        str = "select producto, cantidad,descripcion,(preciounitario + (preciounitario * " & x2s(PORCENTAJE) & ")) / " & x2s(z) & " as punit,(preciototal + (preciototal * " & x2s(PORCENTAJE) & ")) " & x2s(z) & " as ptot from facturaventadetalle where codigofactura=" & cod & "ORDER BY id"
'''                    Else 'NDA
'''                        str = "select producto,descripcion from facturaventadetalle where codigofactura=" & cod
'''                        RptImpresionFacturaVenta.txtivains = Format$(s2n(rs!neto) * s2n(rs!PORCENTAJEiva), "Standard")
'''                        RptImpresionFacturaVenta.txtneto = Format$(rs!neto, "Standard")
'''                        RptImpresionFacturaVenta.txtsub = Format$(rs!neto, "Standard")
'''                    End If
'''                End If
'''             End If
'''        End If
'''
'''        If Not IsNull(rs!direccion) Then
'''            RptImpresionFacturaVenta.lbldomicilio = rs!direccion
'''        Else
'''            RptImpresionFacturaVenta.lbldomicilio = ""
'''        End If
'''        If Not IsNull(rs!provincia) Then
'''            RptImpresionFacturaVenta.provincia = rs!provincia
'''        Else
'''            RptImpresionFacturaVenta.provincia = ""
'''        End If
'''
'''        RptImpresionFacturaVenta.lblfactura = "0001-" & Format(rs!nrofactura, "00000000")
'''        RptImpresionFacturaVenta.lblfecha = rs!Fecha
'''        RptImpresionFacturaVenta.lbliva = rs!iva
'''        If Not IsNull(rs!localidad) Then
'''            RptImpresionFacturaVenta.lbllocalidad = rs!localidad
'''        Else
'''            RptImpresionFacturaVenta.lbllocalidad = ""
'''        End If
'''        If rs!remito <> 0 Then
'''            RptImpresionFacturaVenta.lblref = "Remito"
'''            RptImpresionFacturaVenta.lblnroref = "0001-" & Format(rs!remito, "00000000")
'''        Else
'''            If rs!PEDIDO <> 0 Then
'''                RptImpresionFacturaVenta.lblref = "Pedido"
'''                RptImpresionFacturaVenta.lblnroref = "0001-" & Format(rs!PEDIDO, "00000000")
'''            End If
'''        End If
'''
'''        RptImpresionFacturaVenta.lblpago = rs!pago
'''        RptImpresionFacturaVenta.lblimp = "Son " & ObtenerDescripcion("Monedas", mone) & ": " & enletras(s2n(rs!Total / z))
'''        RptImpresionFacturaVenta.txttotalfinal = Format$(s2n(rs!Total / z), "standard")
'''
''''*************************************************************************************
'''        Dim FormatNro, Sql As String
'''        Sql = "select distinct nroremito from facturaventadetalle where codigofactura=" & cod
'''        rsnro.Open Sql, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
'''
'''        Do While Not rsnro.EOF
'''         If rsnro!NroRemito <> 0 Then
'''            FormatNro = String(6 - Len(rsnro!NroRemito), "0") & rsnro!NroRemito
'''            RptImpresionFacturaVenta.Remitos.Text = RptImpresionFacturaVenta.Remitos.Text & FormatNro & " - "
'''         Else
'''            RptImpresionFacturaVenta.Remitos.Text = ""
'''         End If
'''        rsnro.MoveNext
'''        Loop
'''        If RptImpresionFacturaVenta.Remitos.Text <> "" Then
'''          RptImpresionFacturaVenta.Remitos.Text = Mid(RptImpresionFacturaVenta.Remitos.Text, 1, Len(RptImpresionFacturaVenta.Remitos.Text) - 2)
'''        Else
'''          RptImpresionFacturaVenta.Remitos.Text = ""
'''        End If
'''        rsnro.Close
''''*************************************************************************************
'''        If tdoc = "FAE" Then
'''            'VER CUAL ES LA LEYENDA
'''            If mone <> 1 Then
'''                RptImpresionFacturaVenta.lblleyenda = " Equivalente a " & x2s(rs!Total) & " Pesos al tipo de cambio " & x2s(z) & " pesos por " & ObtenerDescripcion("Monedas", mone)
'''            End If
'''        End If
'''
'''        'RptImpresionFacturaVenta.lblleyenda = "El pago de la presente deberá realizarse en dolares estadounidenses a su vencimiento," _
'''        & "conforme al valor en dicha moneda expresado en este formulario.El comprador asume que el precio en dolares" _
'''        & " ha sido condición esencial de esta venta renunciando a invocar el Art 119A de Código Civil." & vbCrLf _
'''        & "En caso que el pago no pueda realizarse en dicha moneda se realizará en pesos al tipo de cambio vigente para el dolar estadounidense" _
'''        & " tomando la cotización de tipo vendedor del Banco de la Nación Argentina, al cierre de operaciones del día de efectivo pago; " _
'''        & "en caso que a la fecha de pago no existiera mercado Libre de cambios en la Ciudad de Buenos Aires se tomaráan las cotizaciones" _
'''        & " en el Mercado de Nueva York o Montevideo. A opción del vendedor la falta de pago al vencimiento constituye al comprador en mora de " _
'''        & " pleno derecho y hará devengar un interés punitorio del 20% anual hasta el efectivo pago."
'''
'''
'''        RptImpresionFacturaVenta.DataControl1.Connection = DataEnvironment1.Sistema
'''        RptImpresionFacturaVenta.DataControl1.Source = str
'''    End If
'''Call setLpt1
'''
''''----------------------------------------------------------------------------------------
'''    If gEMPR_ImprimeCertCalidad Then
'''    If rs!Certificado And Left(tdoc, 1) = "F" Then
'''
'''        Dim Consulta As String
'''        Dim cCertificadoCalidad As String
'''        Dim cFecha As Date
'''        Dim cCodCli As String
'''        Dim cRZ As String
'''        Dim cCodigoProd As String
'''        Dim cDescProd As String
'''        Dim cCantidad As String
'''        Dim cNroRemito As String
'''        Dim cMuestra As String
'''
'''
'''
'''        Consulta = "SELECT FacturaVentaDetalle.*, Producto.descripcion" _
'''            & " FROM FacturaVentaDetalle INNER JOIN Producto ON FacturaVentaDetalle.Producto = Producto.codigo where facturaventadetalle.codigofactura=" & codigo
'''
'''        rs1.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
'''        If Not rs1.EOF Then
'''            DataEnvironment1.Sistema.Execute "Delete TEMP_CONTROL_CALIDAD"
'''            cFecha = Date
'''            If Not IsNull(rs!RAZONSOCIAL) Then cRZ = rs!RAZONSOCIAL
'''            Do While Not rs1.EOF
'''
''''                rsnro.Open "Select certificadocalidad from bs", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
''''                If Not rsnro.EOF Then
''''                    If Not IsNull(rsnro!certificadocalidad) Then
''''                        cCertificadoCalidad = rsnro!certificadocalidad + 1
''''                        DataEnvironment1.Sistema.Execute "update bs set certificadocalidad=certificadocalidad+1"
''''                    End If
''''                End If
''''                rsnro.Close
''''                Set rsnro = Nothing
'''                cCertificadoCalidad = nSinNull(obtenerDeSQL("Select certificadocalidad from bs")) + 1
'''
'''                If Not IsNull(rs1!codPropio) And Not IsNull(rs1!producto) Then
'''                    cCodigoProd = VerProductoCliente(rs1!producto, rs1!codPropio, rs!cliente)
'''                    cDescProd = rs1!descripcion
'''                End If
'''
'''                If Not IsNull(rs1!cantidad) Then cCantidad = rs1!cantidad
'''
'''                If rs!remito <> 0 Then
'''                    cNroRemito = "0001-" & Format(rs!remito, "00000000")
'''                End If
'''
'''                RptControldeCalidad.lblmuestra.caption = ""
'''
'''                Consulta = "Insert Into TEMP_CONTROL_CALIDAD (CERTIFICADO_CALIDAD, FECHA, RAZON_SOCIAL, CODIGO_PROD, " & _
'''                                                            "DESCRIPCION_PROD, CANTIDAD, NRO_REMITO, MUESTRA) " & _
'''                            "Values ('" & Trim(cCertificadoCalidad) & "'," & _
'''                                    ssFecha(cFecha) & ", " & _
'''                                    "'" & Trim(cRZ) & "', " & _
'''                                    "'" & Trim(cCodigoProd) & "', " & _
'''                                    "'" & Trim(cDescProd) & "', " & _
'''                                    "'" & Trim(cCantidad) & "', " & _
'''                                    "'" & Trim(cNroRemito) & "', " & _
'''                                    "'" & Trim(cMuestra) & "')"
'''                DataEnvironment1.Sistema.Execute Consulta
'''                rs1.MoveNext
'''            Loop
'''        End If
'''        rs1.Close
'''        Set rs1 = Nothing
'''
'''        Consulta = "Select * From TEMP_CONTROL_CALIDAD Order By ID"
'''        With RptControldeCalidad
'''            .Data.Connection = DataEnvironment1.Sistema
'''            .Data.Source = Consulta
'''
'''            .FieFecha.DataField = "FECHA"
'''            .fieCertificado.DataField = "CERTIFICADO_CALIDAD"
'''            .fieCliente.DataField = "RAZON_SOCIAL"
'''            .fieProducto.DataField = "DESCRIPCION_PROD"
'''            .fieCodCliente.DataField = "CODIGO_PROD"
'''
'''            .fieCantidad.DataField = "CANTIDAD"
'''            .fieNroRemito.DataField = "NRO_REMITO"
'''            .fieMuestra.DataField = "MUESTRA"
'''            .Show
'''        End With
'''        DataEnvironment1.Sistema.Execute "update bs set certificadocalidad = certificadocalidad + 1 "
'''    End If
'''    End If
'''
''''Etiquetas orbis **********************************************
'''
'''    If rs!etiqueta And Left(tdoc, 1) = "F" Then
'''     'If rs!Certificado Then
'''
'''       sTablaTemp = TablaTempCrear(tt_Etiquetas_temp)
'''       Dim rsEti As New ADODB.Recordset
'''       Dim CodBarra As String, DifCajas As Double
'''       Dim strInsert, CodProv As String
'''       Dim code As String
'''
'''       Consulta = "SELECT CodigoFactura, fvd.Producto, Cantidad, Producto.descripcion, " & _
'''       " rpc.UnidadesxCaja,rpc.letra,rpc.destino, PRODUCTOCLIENTE,CLIENTE " & _
'''       " FROM Producto INNER JOIN Relacion_Producto_Cliente rpc INNER JOIN " & _
'''       " FacturaVentaDetalle fvd ON rpc.PRODUCTO = fvd.Producto ON Producto.codigo = fvd.Producto " & _
'''       " WHERE  cliente = " & rs!cliente & " and CodigoFactura =  '" & codigo & "' "
'''
'''        rsEti.Open Consulta, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'''
'''        Do While Not rsEti.EOF
'''          CodProv = obtenerDeSQL("select proveedor from clientes where codigo = " & rsEti!cliente & "")
'''          CodBarra = rsEti!productocliente + Space(10 - Len(Trim(rsEti!productocliente))) & _
'''                     LCase(Trim(sSinNull(rsEti!destino))) + Space(1 - Len(LCase(sSinNull(rsEti!letra)))) & _
'''                     Format(rsEti!cantidad, "000000") & _
'''                     Format(CodProv, "000000") & Format(rs!remito, "00000000")
'''          Debug.Print CodBarra
'''
'''          If rsEti!cantidad > rsEti!UnidadesxCaja And rsEti!UnidadesxCaja <> 0 Then
'''                                         'tiene q ser <> 0 pq no puede restar las UnixCajas
'''            Dim cant, UnixCaja, DifCaja As Long
'''            cant = rsEti!cantidad
'''            UnixCaja = rsEti!UnidadesxCaja
'''
'''            While cant > 0
'''               DifCaja = cant - (cant - UnixCaja)
'''
'''              If UnixCaja > cant Then DifCaja = cant  'es para lo q queda en la ultima caja
'''
'''              strInsert = "INSERT INTO " & sTablaTemp & " " & _
'''              "(ProdCliente,Cantidad,letra,descripcion,destino,remito,CodBarra ) " & _
'''              " VALUES('" & rsEti!productocliente & "','" & DifCaja & "', " & _
'''              "'" & LCase(sSinNull(rsEti!letra)) & "','" & rsEti!descripcion & "', " & _
'''              "'" & LCase(sSinNull(rsEti!destino)) & "', " & _
'''              "'" & rs!remito & "','" & CodBarra & "')"
'''              DataEnvironment1.Sistema.Execute strInsert
'''              cant = cant - UnixCaja        'hasta terminar las UnixCajas
'''                                            'tiene q ser <> 0 pq no puede restar las UnixCajas
'''            Wend
'''            Else
'''                strInsert = "INSERT INTO " & sTablaTemp & " " & _
'''              "(ProdCliente,Cantidad,letra,descripcion,destino,remito,CodBarra ) " & _
'''              " VALUES('" & rsEti!productocliente & "','" & rsEti!cantidad & "', " & _
'''              "'" & LCase(sSinNull(rsEti!letra)) & "','" & rsEti!descripcion & "', " & _
'''              "'" & LCase(sSinNull(rsEti!destino)) & "', " & _
'''              "'" & rs!remito & "','" & CodBarra & "')"
'''              DataEnvironment1.Sistema.Execute strInsert
'''
'''          End If
'''
'''          rsEti.MoveNext
'''        Loop
'''        Dim RsEtiqueta As New ADODB.Recordset
'''        str = "select * from " & sTablaTemp & ""
'''        RsEtiqueta.Open str, DataEnvironment1.Sistema
'''
'''        RptImpresionEtiqueta.DataEtiqueta.Connection = DataEnvironment1.Sistema
'''        RptImpresionEtiqueta.DataEtiqueta.Source = str
'''
'''        RptImpresionEtiqueta.lblfecha = Date
'''        RptImpresionEtiqueta.LblTomka = CodProv
'''        rsEti.Close
'''        RsEtiqueta.Close
'''     'End If
'''
'''
'''    Set rsEti = Nothing
'''    Set RsEtiqueta = Nothing
'''
'''    RptImpresionEtiqueta.Show
'''  End If
'''    RptImpresionFacturaVenta.Show
'''    RptImpresionFacturaVenta.Printer.Copies = 2
'''    rs.Close
'''    Set rs = Nothing
'''fin:
'''Exit Function
'''ErrImpresion:
'''    ufa "Error de impresión en factura venta", ""
'''    Call setLpt1
'''    Resume fin
'''End Function

Public Function ItemFacturaRemitoCompra(Ope As ItemFRC, ItemRemito As Long, CodProv As Long, tdoc As String, NroDoc As Long, cantidad As Double, Optional PrecioUnitario As Double)
    'Trabaja sobre 2 tablas:
    '   RemitoCompraDetalle     acualizando  .Cantidad_a_Facturar
    '   FacturaCompraRemito     tabla de nexo,   item remito, item factura ,  cant y precio
    ' modificacion tambien resta cantidad
    
    ' sql remito es facil, id = codigo autonum
    ' sql Compras Transcom identifica solo x   CodPR+TipoDoc+NroDoc, una CArGADA pesada a la consulta
    On Error GoTo ufaErr
    Dim ssql As String
    Dim rs As New ADODB.Recordset
    
    Select Case Ope
    Case ItemFRC_Alta
        DataEnvironment1.Sistema.Execute "Update RemitoCompraDetalle set cantidad_a_facturar = cantidad_a_facturar - " & cantidad & " where codigo = " & ItemRemito
        DataEnvironment1.Sistema.Execute "Insert into FacturaCompraRemito ( TipoDoc, NroDoc, CodPr, ItemRemitoCompra, Cantidad, PrecioUnitario) values ('" & tdoc & "'," & NroDoc & ", " & CodProv & ", " & ItemRemito & " ," & x2s(cantidad) & ", " & x2s(PrecioUnitario) & ") "
    Case ItemFRC_Mod
        DataEnvironment1.Sistema.Execute "Update RemitoCompraDetalle set cantidad_a_facturar = cantidad_a_facturar - " & cantidad & " where codigo = " & ItemRemito
        DataEnvironment1.Sistema.Execute "delete from FacturaCompraRemito where  where CodPr = " & CodProv & " and  TipoDoc = '" & tdoc & "' and  NroDoc = " & NroDoc & " and  ItemRemito = " & ItemRemito
        DataEnvironment1.Sistema.Execute "Insert into FacturaCompraRemito ( TipoDoc, NroDoc, CodPr, ItemRemitoCompra, Cantidad, PrecioUnitario) values ('" & tdoc & "'," & NroDoc & ", " & CodProv & ", " & ItemRemito & " ," & x2s(cantidad) & ", " & x2s(PrecioUnitario) & ") "
    Case ItemFRC_BajaFactura
        
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''
                        'sSql = "select cant from FacturaCompraRemito where CodPr = " & CodProv & " and  TipoDoc = '" & tDoc & "' and  NroDoc = " & NroDoc
                        '
                        '      daTaenvironment1.Sistema.Execute "Update RemitoCompraDetalle set cantidad_a_facturar = cantidad_a_facturar + " &   & " where codigo = " & ItemRemito
                        '
                        ''''''''''''''''''''''''''''''''''''''
                        
                        'daTaenvironment1.Sistema.Execute daTaenvironment1.Sistema.Execute "Update RemitoCompraDetalle set cantidad_a_facturar = cantidad_a_facturar + " &   & " where codigo = " & ItemRemito
        'actualiza cant a facturar
'        daTaenvironment1.Sistema.Execute "UPDATE RemitoCompraDetalle INNER JOIN FacturaCompraRemito ON RemitoCompraDetalle.codigo = FacturaCompraRemito.ItemRemitoCompra SET cantidad_a_facturar = cantidad_a_facturar + cantidad WHERE FacturaCompraRemito.TipoDoc ='" & tDoc & "' AND FacturaCompraRemito.NroDoc = " & NroDoc & " AND FacturaCompraRemito.CodPr = " & CodProv
            
        With rs
            ssql = "select itemRemitoCompra, cantidad from FacturaCompraRemito  WHERE FacturaCompraRemito.TipoDoc ='" & tdoc & "' AND FacturaCompraRemito.NroDoc = " & NroDoc & " AND FacturaCompraRemito.CodPr = " & CodProv
            .Open ssql, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
            While Not .EOF
                DataEnvironment1.Sistema.Execute "update RemitoCompraDetalle set Cantidad_a_Facturar = Cantidad_a_Facturar + " & x2s(!cantidad) & " where codigo = " & !ItemRemitoCompra
                .MoveNext
            Wend
            .Close
        End With
        DataEnvironment1.Sistema.Execute "delete from FacturaCompraRemito where CodPr = " & CodProv & " and  TipoDoc = '" & tdoc & "' and  NroDoc = " & NroDoc
    Case ItemFRC_BajaRemito
        'el update es al reverendo inutil
        DataEnvironment1.Sistema.Execute "delete from FacturaCompraRemito where ItemRemito = " & ItemRemito
    Case Else
        ufa "err prg", "ItemFacturaRemitoCompra()-" & Ope & CodProv & " '" & tdoc & "' " & " " & NroDoc & ItemRemito & cantidad & PrecioUnitario ', Err
    End Select
   
fin:
    Set rs = Nothing
    Exit Function
ufaErr:
    ufa "err al grabar detalle ItFacRem", "ItemFacturaRemitoCompra()-" & Ope & CodProv & " '" & tdoc & "' " & " " & NroDoc & ItemRemito & cantidad & PrecioUnitario ', Err
    Resume fin
End Function


Public Function AnularFacturaVenta(codi As Long, iddoc As Long) As Boolean
    If ON_ERROR_HABILITADO Then On Error GoTo UFAelim
    
    Dim rs As New ADODB.Recordset ', codi
    Dim tmp As Variant, activ As Long, resto As Double, modSt As Long
    Dim depot As Long, nPedi, nRemi, item, tiDoc As String, nume, FPago, Fecha As Date
    
    
    If codi = 0 Then
        ufa "Err: Factura 0", "AnularFacturaVenta()"
        Exit Function
    End If
    If existe_en_recibo(codi) = True Then Exit Function
    
    'codi = s2n(txtCodigo)
'    If codi = 0 Then Exit Function
    
    tmp = obtenerDeSQL("select activo, total-saldo as resto, ActualizaStock, deposito, pedido, remito, tipoDoc, formapago, fecha  from FacturaVenta where codigo = " & codi)
    If IsNull(tmp) Then
        ufa "err leyendo factura a eliminar", codi ', Err
        Exit Function
    End If
    activ = tmp(0)
    resto = tmp(1)
    modSt = tmp(2)
    depot = tmp(3)
    nPedi = tmp(4)
    nRemi = tmp(5)
    tiDoc = Trim(tmp(6))
    FPago = tmp(7)
    Fecha = tmp(8)
    
        
    nume = obtenerDeSQL("select NroFactura from FacturaVenta where codigo = " & codi) 's2n(TxtNroFactura)
    
    
    If activ = 0 Then
        MsgBox "Factura figura como ya anulada"
        Exit Function
    End If
    If FPago <> FormaPago_CONTADO And Round(resto, 2) <> 0 Then
        MsgBox "No se puede anular factura" & vbCrLf & "Ya fue imputada"
        Exit Function
    End If
    If Not PuedoVentas(Fecha) Then
        
        Exit Function
    End If
    ' En ***TONKA*** es mensaje al pedo
'    If modSt = True Then MsgBox "Esta anulacion actualizara el Stock  "


    If Not confirma("Anular Factura " & nume) Then Exit Function
     
     
    '----TRANS------------------------------------------------
    DE_BeginTrans
    With rs
        .Open "select producto, cantidad, item_p_r, NroPedido, NroRemito  from FacturaVentaDetalle where CodigoFactura = " & codi, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        While Not .EOF
            item = s2n(!item_p_r)
            
            If modSt = True Then
                If tiDoc = "NCA" Or tiDoc = "NCB" Or tiDoc = "NCE" Then                            'devoluc
                    DataEnvironment1.dbo_SumaStock !producto, -!cantidad, depot   ' resta stock
                ElseIf tiDoc = "FAA" Or tiDoc = "FAB" Or tiDoc = "FAE" Then                         'factura
                    DataEnvironment1.dbo_SumaStock !producto, !cantidad, depot    ' suma stock
                Else
                    ufa "prg err", "intento de mod stock " & tiDoc & " AnularFacturaVenta()" ', Err
                End If
            End If
            nPedi = s2n(!NroPedido)
            nRemi = s2n(!NroRemito)
            
            If nRemi > 0 And modSt = False Then
                DataEnvironment1.Sistema.Execute "Update RemitoVentaDetalle set facturar = facturar + " & !cantidad & " where  numero = " & nRemi & " and producto='" & Trim(!producto) & "'"
                
                DataEnvironment1.Sistema.Execute "Update RemitoVenta set factura = 0 where  numero = " & nRemi
            End If
            
            'If nPedi + nRemi > 0 And Item > 0 Then
            '    DataEnvironment1.dbo_abmFacturaVentaDetalle "B", codi, 0, 0, !cantidad, 0, !producto, "", "", 0, 0, 0, nPedi, nRemi, Item, 0
            'End If
            .MoveNext
        Wend
    End With
        
    'cabecera
    DataEnvironment1.dbo_abmFacturaVenta "B", codi, 0, 0, 0, 0, 0, 0, 0, "", "", "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, UsuarioActual(), Date, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
    'series ?
    DataEnvironment1.Sistema.Execute "update series set activo=0, usuario_baja = " & UsuarioActual() & ", fecha_baja = " & ssFecha(Date) & " where nroComprobante = " & nume & " and comprobante = '" & ObtenerCodigo("TipoComprobantesGrales", tiDoc) & "'"
    'movi ?
    If FPago = FormaPago_CONTADO Then
        'cheques
        DataEnvironment1.Sistema.Execute "update cheques set activo = 0, fecha_baja = " & ssFecha(Date) & ", usuario_baja = " & UsuarioActual() & " where ndoc = '" & nume & "' and tdoc = '" & tiDoc & "'"
        'movicaja
        DataEnvironment1.dbo_MOVICAJASdoc "B", 0, 0, 0, "", "", 0, "", Date, "", 0, 0, tiDoc, nume, Date, UsuarioActual(), 0, iddoc
    End If
    
            'Baja Doc y asiento
'            If Not BorroDocumento(iddoc) Then
'                ufa "err al borrar documento", " middoc = " & iddoc
'                DE_RollbackTrans
'                GoTo fin:
'            End If
    If siAsiento("AsientosVentas") Then
        If iddoc > 0 Then BorroDocumento (iddoc)  ' pregunto para datos migrados, sino  deberia chequear iddoc siempres
    End If
    DE_CommitTrans
    '----TRANS------------------------------------------------------
    
    grabaBitacora "B", nume, "FactVenta series movi"
'    MsgBox " Factura Anulada "
    AnularFacturaVenta = True
    
fin:
    Set rs = Nothing
    Exit Function
UFAelim:
    DE_RollbackTrans
    ufa "error al eliminar", "AnularFacturaVenta() - " & tiDoc & nume ', Err
    AnularFacturaVenta = False
    Resume fin
End Function

Private Function existe_en_recibo(codFac As Long) As Boolean
Dim existe
existe_en_recibo = False
existe = obtenerDeSQL("select codrecibo from recibosdetalle where facturaventa= " & codFac)
    If IsNull(existe) Or IsEmpty(existe) Then
    Else
        existe = obtenerDeSQL("select numero, activo from recibos where codigo= " & existe)
        If existe(1) = True Then
            MsgBox "La factura figura asociado a un recibo." & Chr(13) & "Elimine el recibo antes que la factura." & Chr(13) & "RECIBO : " & existe(0), vbCritical, "Informe"
            existe_en_recibo = True
        End If
    End If
End Function

Public Function ImprimirComprobThor(codigo) As Boolean

    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim rsnro As New ADODB.Recordset
    Dim rsnroP As New ADODB.Recordset
    Dim str As String
    Dim COD As Long
    Dim PORCENTAJE As String
    Dim tdoc As String, z As Double, mone As Long
    
    rs.Open "SELECT FacturaVenta.*, Clientes.*, FormasPago.descripcion as pago, Ivas.descripcion as iva, Ivas.letra as letra" _
        & " FROM FormasPago INNER JOIN (Ivas INNER JOIN (Clientes INNER JOIN FacturaVenta ON Clientes.codigo = FacturaVenta.Cliente) ON Ivas.codigo = FacturaVenta.TipoIVA) ON FormasPago.codigo = FacturaVenta.FormaPago where facturaventa.codigo=" & codigo, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    
    tdoc = Trim(rs!TIPODOC)
    z = s2n(rs!cotizacion, 4)
    If z = 0 Then z = 1
    mone = rs!moneda
    If mone = 0 Then mone = 1 '(PESOS)
    
    If Not rs.EOF Then
        RptImpresionFactThorVenta.lblcliente = rs!RAZONSOCIAL
        If Not IsNull(rs!CUIT) Then
            RptImpresionFactThorVenta.lblCuit = rs!CUIT
        End If
        COD = codigo
        If Not IsNull(rs!PorcentajeIva) Then
         PORCENTAJE = rs!PorcentajeIva
        End If

        If Left(tdoc, 2) = "FA" Then
            
            RptImpresionFactThorVenta.lblcomp = "Factura"
            If tdoc = "FAB" Then
                str = "select producto,cantidad,descripcion,(preciounitario + (preciounitario * " & x2s(PORCENTAJE) & ")) / " & x2s(z) & " as punit,(preciototal + (preciototal * " & x2s(PORCENTAJE) & ")) / " & x2s(z) & "  as ptot from facturaventadetalle where codigofactura=" & COD
            ElseIf tdoc = "FAE" Then
                str = "select producto,cantidad,descripcion,preciounitario / " & x2s(z) & " as punit ,preciototal / " & x2s(z) & " as ptot from facturaventadetalle where codigofactura=" & COD & "ORDER BY id"
            
            ElseIf tdoc = "FAA" Then
                str = "select producto,cantidad,descripcion,preciounitario / " & x2s(z) & " as punit ,preciototal / " & x2s(z) & " as ptot from facturaventadetalle where codigofactura=" & COD & "ORDER BY id"

                RptImpresionFactThorVenta.txtivains = Format$(s2n(s2n(rs!Neto / z, 2) - (s2n(rs!Neto / z, 2) * s2n(rs!Descuento, 2)), 2) * s2n(rs!PorcentajeIva, 4), "standard")
                RptImpresionFactThorVenta.txtneto = Format$(s2n(rs!Neto / z, 2), "standard")
                RptImpresionFactThorVenta.Descuento = Format$(s2n(rs!Neto / z, 2) * s2n(rs!Descuento, 2), "standard")
                RptImpresionFactThorVenta.txtsub = Format$(s2n(rs!Neto / z, 2) - (s2n(rs!Neto / z, 2) * s2n(rs!Descuento, 2)), "standard")
                RptImpresionFactThorVenta.txtIvaP = Format$(s2n(PORCENTAJE, 4) * 100, "standard")
            End If
        
        Else

            If Left(tdoc, 2) = "NC" Then
                RptImpresionFactThorVenta.lblcomp = "Nota de Credito"
                RptImpresionFactThorVenta.lbltachar = "XXXXXXXXX"
                str = "select producto,cantidad,descripcion,(preciounitario + (preciounitario * " & x2s(PORCENTAJE) & ")) / " & x2s(z) & " as punit,(preciototal + (preciototal * " & x2s(PORCENTAJE) & ")) / " & x2s(z) & " as ptot from facturaventadetalle where codigofactura=" & COD & "ORDER BY id"
                
                If tdoc = "NCB" Then

                ElseIf tdoc = "NCE" Then

                ElseIf tdoc = "NCA" Then

                    'str = "select descripcion from facturaventadetalle where codigofactura=" & cod
                    RptImpresionFactThorVenta.txtivains = Format$(s2n(s2n(rs!Neto / z, 2) - (s2n(rs!Neto / z, 2) * s2n(rs!Descuento, 2)), 2) * s2n(rs!PorcentajeIva, 4), "standard")
                    RptImpresionFactThorVenta.txtneto = Format$(s2n(rs!Neto / z, 2), "standard")
                    RptImpresionFactThorVenta.Descuento = Format$(s2n(rs!Neto / z, 2) * s2n(rs!Descuento, 2), "standard")
                    RptImpresionFactThorVenta.txtsub = Format$(s2n(rs!Neto / z, 2) - (s2n(rs!Neto / z, 2) * s2n(rs!Descuento, 2)), "standard")
                    RptImpresionFactThorVenta.txtIvaP = Format$(s2n(PORCENTAJE, 4) * 100, "standard")
                End If
            Else
                'If Trim(rs!TIPODOC) = "NDA" Or Trim(rs!TIPODOC) = "NDB" Then
                If Left(tdoc, 2) = "ND" Then
                    RptImpresionFactThorVenta.lblcomp = "Nota de Debito"
                    RptImpresionFactThorVenta.lbltachar = "XXXXXXXXX"
                    If tdoc = "NDB" Then
                        str = "select producto,cantidad,descripcion,(preciounitario + (preciounitario * " & x2s(PORCENTAJE) & ")) / " & x2s(z) & " as punit,(preciototal + (preciototal * " & x2s(PORCENTAJE) & ")) " & x2s(z) & " as ptot from facturaventadetalle where codigofactura=" & COD & "ORDER BY id"
                        'str = "select descripcion from facturaventadetalle where codigofactura=" & cod
                    ElseIf tdoc = "NDE" Then
                        str = "select producto,cantidad,descripcion,(preciounitario + (preciounitario * " & x2s(PORCENTAJE) & ")) / " & x2s(z) & " as punit,(preciototal + (preciototal * " & x2s(PORCENTAJE) & ")) " & x2s(z) & " as ptot from facturaventadetalle where codigofactura=" & COD & "ORDER BY id"
                        
                    ElseIf tdoc = "NDA" Then
                        str = "select producto,cantidad,descripcion,(preciounitario + (preciounitario * " & x2s(PORCENTAJE) & ")) / " & x2s(z) & " as punit,(preciototal + (preciototal * " & x2s(PORCENTAJE) & ")) " & x2s(z) & " as ptot from facturaventadetalle where codigofactura=" & COD & "ORDER BY id"
                        'str = "select descripcion from facturaventadetalle where codigofactura=" & cod
                        RptImpresionFactThorVenta.txtivains = Format$(s2n(s2n(rs!Neto / z, 2) - (s2n(rs!Neto / z, 2) * s2n(rs!Descuento, 2)), 2) * s2n(rs!PorcentajeIva, 4), "standard")
                        RptImpresionFactThorVenta.txtneto = Format$(s2n(rs!Neto / z, 2), "standard")
                        RptImpresionFactThorVenta.Descuento = Format$(s2n(rs!Neto / z, 2) * s2n(rs!Descuento, 2), "standard")
                        RptImpresionFactThorVenta.txtsub = Format$(s2n(rs!Neto / z, 2) - (s2n(rs!Neto / z, 2) * s2n(rs!Descuento, 2)), "standard")
                        RptImpresionFactThorVenta.txtIvaP = Format$(s2n(PORCENTAJE, 4) * 100, "standard")
                    End If
                End If
             End If
        End If
'*****************************************************************************************
        Dim FormatNro, sql As String
        
        
        sql = "select distinct nroremito from facturaventadetalle where codigofactura=" & COD
        rsnro.Open sql, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
        Do While Not rsnro.EOF
         If rsnro!NroRemito <> 0 Then
            FormatNro = String(6 - Len(rsnro!NroRemito), "0") & rsnro!NroRemito
            RptImpresionFactThorVenta.Remitos.Text = RptImpresionFactThorVenta.Remitos.Text & FormatNro & " - "
         Else
            RptImpresionFactThorVenta.Remitos.Text = ""
         End If
        rsnro.MoveNext
        Loop
        If RptImpresionFactThorVenta.Remitos.Text <> "" Then
          RptImpresionFactThorVenta.Remitos.Text = Mid(RptImpresionFactThorVenta.Remitos.Text, 1, Len(RptImpresionFactThorVenta.Remitos.Text) - 2)
        Else
          RptImpresionFactThorVenta.Remitos.Text = ""
        End If
        rsnro.Close
        Dim FormatNroP, sql1 As String
        
        
        sql1 = "select distinct facturaventadetalle.nropedido,pedido_cli from facturaventadetalle inner join pedidos_clientes on facturaventadetalle.nropedido=pedidos_clientes.numero where codigofactura=" & COD
        rsnroP.Open sql1, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
        Do While Not rsnroP.EOF
         If rsnroP!NroPedido <> 0 Then
            FormatNroP = String(6 - Len(rsnroP!NroPedido), "0") & rsnroP!NroPedido
            RptImpresionFactThorVenta.Pedidos.Text = RptImpresionFactThorVenta.Pedidos.Text & FormatNroP & " - "
            RptImpresionFactThorVenta.Compra.Text = rsnroP!pedido_cli
         Else
            RptImpresionFactThorVenta.Pedidos.Text = ""
            RptImpresionFactThorVenta.Compra.Text = ""
         End If
        rsnroP.MoveNext
        Loop
        If RptImpresionFactThorVenta.Pedidos.Text <> "" Then
          RptImpresionFactThorVenta.Pedidos.Text = Mid(RptImpresionFactThorVenta.Pedidos.Text, 1, Len(RptImpresionFactThorVenta.Pedidos.Text) - 2)
        Else
          RptImpresionFactThorVenta.Pedidos.Text = ""
        End If
        rsnroP.Close
                
'*****************************************************************************************
        If Not IsNull(rs!direccion) Then
            RptImpresionFactThorVenta.lbldomicilio = rs!direccion
        Else
            RptImpresionFactThorVenta.lbldomicilio = ""
        End If
        RptImpresionFactThorVenta.lblfactura = "0001-" & Format(rs!NroFactura, "00000000")
        RptImpresionFactThorVenta.lblfecha = rs!Fecha
        RptImpresionFactThorVenta.lblVencim = rs!Vencimiento
        RptImpresionFactThorVenta.LblIVA = rs!Iva
        If Not IsNull(rs!Localidad) Then
            RptImpresionFactThorVenta.lbllocalidad = rs!Localidad
        Else
            RptImpresionFactThorVenta.lbllocalidad = ""
        End If
        If rs!Remito <> 0 Then
            RptImpresionFactThorVenta.lblref = "Remito"
            RptImpresionFactThorVenta.lblnroref = "0001-" & Format(rs!Remito, "00000000")
        Else
            If rs!Pedido <> 0 Then
                RptImpresionFactThorVenta.lblref = "Pedido"
                RptImpresionFactThorVenta.lblnroref = "0001-" & Format(rs!Pedido, "00000000")
            End If
        End If
        RptImpresionFactThorVenta.lblpago = rs!pago
        RptImpresionFactThorVenta.lblimp = ObtenerDescripcion("Monedas", mone) & ": " & enletras(s2n(rs!Total / z))
        RptImpresionFactThorVenta.txttotalfinal = Format$(s2n(rs!Total / z), "standard")
        If tdoc = "FAE" Then
            'VER CUAL ES LA LEYENDA
            If mone <> 1 Then
                'RptImpresionFactThorVenta.lblleyenda = " Equivalente a " & x2s(rs!Total) & " Pesos al tipo de cambio " & x2s(z) & " pesos por " & ObtenerDescripcion("Monedas", mone)
            End If
        Else
            'RptImpresionFactThorVenta.lblleyenda = "Equivalente a dolares estadounidenses U$S          ." & Chr(13) _
            '& "Al tipo de cambio            peso/s por dolar segun clausula al pie." & Chr(13) & "El pago de la presente deberá realizarse en dolares estadounidenses a su vencimiento," _
            '& "conforme al valor en dicha moneda expresado en este formulario.El comprador asume que el precio en dolares" _
            '& " ha sido condición esencial de esta venta renunciando a invocar el Art 119A de Código Civil." & vbCrLf _
            '& "En caso que el pago no pueda realizarse en dicha moneda se realizará en pesos al tipo de cambio vigente para el dolar estadounidense" _
            '& " tomando la cotización de tipo vendedor del Banco de la Nación Argentina, al cierre de operaciones del día de efectivo pago; " _
            '& "en caso que a la fecha de pago no existiera mercado Libre de cambios en la Ciudad de Buenos Aires se tomaráan las cotizaciones" _
            '& " en el Mercado de Nueva York o Montevideo. A opción del vendedor la falta de pago al vencimiento constituye al comprador en mora de " _
            '& " pleno derecho y hará devengar un interés punitorio del 20% anual hasta el efectivo pago."
        End If
        
        RptImpresionFactThorVenta.DataControl1.Connection = DataEnvironment1.Sistema
        RptImpresionFactThorVenta.DataControl1.Source = str
        
    End If
    
    If gEMPR_ImprimeCertCalidad Then

        Dim Consulta As String
        Dim cCertificadoCalidad As String
        Dim cFecha As Date
        Dim cCodCli As String
        Dim cRZ As String
        Dim cCodigoProd As String
        Dim cDescProd As String
        Dim cCantidad As String
        Dim cNroRemito As String
        Dim cMuestra As String
        
        
        
        Consulta = "SELECT FacturaVentaDetalle.*, Producto.descripcion" _
            & " FROM FacturaVentaDetalle INNER JOIN Producto ON FacturaVentaDetalle.Producto = Producto.codigo where facturaventadetalle.codigofactura=" & codigo

        rs1.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs1.EOF Then
            DataEnvironment1.Sistema.Execute "Delete TEMP_CONTROL_CALIDAD"
            cFecha = Date
            If Not IsNull(rs!RAZONSOCIAL) Then cRZ = rs!RAZONSOCIAL
            Do While Not rs1.EOF
                
                rsnro.Open "Select certificadocalidad from bs", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
                If Not rsnro.EOF Then
                    If Not IsNull(rsnro!certificadocalidad) Then
                        cCertificadoCalidad = rsnro!certificadocalidad + 1
                        DataEnvironment1.Sistema.Execute "update bs set certificadocalidad=certificadocalidad+1"
                    End If
                End If
                rsnro.Close
                Set rsnro = Nothing
                
                If Not IsNull(rs1!codPropio) And Not IsNull(rs1!producto) Then
                    cCodigoProd = VerProductoCliente(rs1!producto, rs1!codPropio, rs!codigo)
                    cDescProd = rs1!DESCRIPCION
                End If
                
                If Not IsNull(rs1!cantidad) Then cCantidad = rs1!cantidad
                
                If rs!Remito <> 0 Then
                    cNroRemito = "0001-" & Format(rs!Remito, "00000000")
                End If
                
                RptControldeCalidad.lblmuestra.caption = ""
                
                Consulta = "Insert Into TEMP_CONTROL_CALIDAD (CERTIFICADO_CALIDAD, FECHA, RAZON_SOCIAL, CODIGO_PROD, " & _
                                                            "DESCRIPCION_PROD, CANTIDAD, NRO_REMITO, MUESTRA) " & _
                            "Values ('" & Trim(cCertificadoCalidad) & "'," & _
                                    ssFecha(cFecha) & ", " & _
                                    "'" & Trim(cRZ) & "', " & _
                                    "'" & Trim(cCodigoProd) & "', " & _
                                    "'" & Trim(cDescProd) & "', " & _
                                    "'" & Trim(cCantidad) & "', " & _
                                    "'" & Trim(cNroRemito) & "', " & _
                                    "'" & Trim(cMuestra) & "')"
                DataEnvironment1.Sistema.Execute Consulta
                rs1.MoveNext
            Loop
        End If
        rs1.Close
        Set rs1 = Nothing
                        
        Consulta = "Select * From TEMP_CONTROL_CALIDAD Order By ID"
        With RptControldeCalidad
            .Data.Connection = DataEnvironment1.Sistema
            .Data.Source = Consulta
            
            .fieFecha.DataField = "FECHA"
            .fieCertificado.DataField = "CERTIFICADO"
            .fieCliente.DataField = "RAZON_SOCIAL"
            .fieProducto.DataField = "DESCRIPCION_PROD"
            .fieCodCliente.DataField = "CODIGO_PROD"
            
            .fieCantidad.DataField = "CANTIDAD"
            .fieNroRemito.DataField = "NRO_REMITO"
            .fieMuestra.DataField = "MUESTRA"
            .Show
        End With
    End If
    rs.Close
    Set rs = Nothing
    
    RptImpresionFactThorVenta.PageSettings.TopMargin = margenTopThor_FV()
    
    RptImpresionFactThorVenta.Printer.Copies = 2
    RptImpresionFactThorVenta.Show vbModal
    
    
    
    'RptImpresionFactThorVenta.PrintReport True
    
End Function

Private Function margenTopThor_FV() As Long
    On Error GoTo fin
    Dim x As Long
    margenTopThor_FV = RptImpresionFactThorVenta.PageSettings.TopMargin
    x = s2n(VerDatoEmpresa("MargenTop_FV"))
    If x > 0 Then margenTopThor_FV = x
fin:
End Function

Private Function margenTopThor_RV() As Long
    On Error GoTo fin
    Dim x As Long
    margenTopThor_RV = RptImpresionRemitoVenta.PageSettings.TopMargin
    x = s2n(VerDatoEmpresa("margenTop_RV"))
    If x > 0 Then margenTopThor_RV = x
fin:
End Function

Public Function ImprimirAMRAT(codigo, Optional Triplicado As Boolean = False, Optional chq As Boolean = False) As Boolean

    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim rsnro As New ADODB.Recordset
    Dim str As String
    Dim COD As Long
    Dim PORCENTAJE As String
    Dim tdoc As String, z As Double, mone As Long
    Dim usuario As Long
    Dim Subtot As Double
    Dim Fact
    Dim i As Long
    Dim a As String
    Dim sig As String
    Dim FormatNro, sql As String
    
'    a = "SELECT FacturaVenta.*, Clientes.*, FormasPago.descripcion as pago, Ivas.descripcion as iva, Ivas.letra as letra,FacturaVenta.codigo as cod,FacturaVenta._docum_ve as remi,FacturaVenta._control_ve as OC,FacturaVenta.vendedor as vend" _
        & " FROM FormasPago INNER JOIN (Ivas INNER JOIN (Clientes INNER JOIN FacturaVenta ON Clientes.codigo = FacturaVenta.Cliente) ON Ivas.codigo = FacturaVenta.TipoIVA) ON FormasPago.codigo = FacturaVenta.FormaPago where facturaventa.codigo=" & codigo
    rs.Open "SELECT FacturaVenta.*, Clientes.*, FormasPago.descripcion as pago, Ivas.descripcion as iva, Ivas.letra as letra,FacturaVenta.codigo as cod,FacturaVenta._docum_ve as remi,FacturaVenta._control_ve as OC,FacturaVenta.vendedor as vend" _
        & " FROM FormasPago INNER JOIN (Ivas INNER JOIN (Clientes INNER JOIN FacturaVenta ON Clientes.codigo = FacturaVenta.Cliente) ON Ivas.codigo = FacturaVenta.TipoIVA) ON FormasPago.codigo = FacturaVenta.FormaPago where facturaventa.codigo=" & codigo, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    
    rs1.Open "SELECT leyenda FROM facturaventaleyenda where fac=" & codigo, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    While Not rs1.EOF
        If rptImpresionFacturaVentaAT.leye = "" Then
            rptImpresionFacturaVentaAT.leye = s2t(rs1, "leyenda")
        Else
            rptImpresionFacturaVentaAT.leye = rptImpresionFacturaVentaAT.leye & Chr(13) & s2t(rs1, "leyenda")
        End If
        rs1.MoveNext
    Wend
    Set rs1 = Nothing
        
    tdoc = Trim(rs!TIPODOC)
    z = s2n(rs!cotizacion, 4)
    If z = 0 Then z = 1
    If z > 1 Then
        rptImpresionFacturaVentaAT.Label6.Visible = True
        rptImpresionFacturaVentaAT.Label6.caption = "Equivale a U$S 1(dolares)= $ " & z & "(pesos)"
    Else
        rptImpresionFacturaVentaAT.Label6.Visible = False
        rptImpresionFacturaVentaAT.Label6.caption = ""
    End If
               
    mone = rs!moneda
    If mone = 0 Then mone = 1 '(PESOS)
    '************************************************
    a = ""
    Fact = Split(Replace(rs!variasfac, "#", ""), ",")
    For i = 0 To UBound(Fact)
        'If a < Fact(i) Then a = Fact(i)
        a = a & " or nrofactura=" & Fact(i)
    Next
    sig = ""
    For i = 0 To UBound(Fact)
        If rs!NroFactura < CLng(Fact(i)) Then
            sig = Fact(i)
            Exit For
        End If
    Next
    '************************************************
    If Not rs.EOF Then
        rptImpresionFacturaVentaAT.lblcliente = rs!RAZONSOCIAL
        If Not IsNull(rs!CUIT) Then
            rptImpresionFacturaVentaAT.lblCuit = rs!CUIT
            If Len(Trim(rptImpresionFacturaVentaAT.lblCuit)) < 13 Then
                rptImpresionFacturaVentaAT.lblCuit = sSinNull(rs!dni)
            End If
        End If
        COD = codigo
        If Not IsNull(rs!PorcentajeIva) Then
         PORCENTAJE = rs!PorcentajeIva
        End If

        If Left(tdoc, 2) = "FA" Then
                        
            If tdoc = "FAB" Then
                rptImpresionFacturaVentaAT.lblcomp = "Factura B"
                str = "select cantidad,descripcion,(preciounitario + (preciounitario * " & x2s(PORCENTAJE) & ")) / " & x2s(z) & " as punit,(preciototal + (preciototal * " & x2s(PORCENTAJE) & ")) / " & x2s(z) & "  as ptot from facturaventadetalle where codigofactura=" & COD & " order by id"
            ElseIf tdoc = "FAE" Then
                rptImpresionFacturaVentaAT.lblcomp = "Factura E"
                str = "select cantidad,descripcion,preciounitario / " & x2s(z) & " as punit ,preciototal / " & x2s(z) & " as ptot from facturaventadetalle where codigofactura=" & COD & " ORDER BY id"
            
            ElseIf tdoc = "FAA" Then
                rptImpresionFacturaVentaAT.lblcomp = "Factura A"
                str = "select cantidad,descripcion,preciounitario / " & x2s(z) & " as punit ,preciototal / " & x2s(z) & " as ptot from facturaventadetalle where codigofactura=" & COD & " ORDER BY id"

                'rptImpresionFacturaVentaAT.txtivains = Format$(s2n(rs!Neto / z, 2) * s2n(rs!PorcentajeIva, 4), "standard")
                rptImpresionFacturaVentaAT.txtivains = Format$(s2n(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where (codigofactura=" & rs!COD & a & ") and _iva=21") / z, 2) * s2n(0.21, 4), "standard")
                rptImpresionFacturaVentaAT.txtivains2 = Format$(s2n(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where (codigofactura=" & rs!COD & a & ") and _iva=10.5") / z, 2) * s2n(0.105, 4), "standard")
                Subtot = Format$(s2n(rs!Neto / z, 2), "standard")
                If rptImpresionFacturaVentaAT.txtivains2 = 0 Then
                    rptImpresionFacturaVentaAT.txtivains2 = ""
                    rptImpresionFacturaVentaAT.txtneto2 = ""
                    rptImpresionFacturaVentaAT.txtsub2 = ""
                Else
                    rptImpresionFacturaVentaAT.txtneto2 = Format$(s2n(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where (codigofactura=" & rs!COD & a & ") and _iva=10.5") / z, 2), "standard")
                    rptImpresionFacturaVentaAT.txtsub2 = Format$(s2n(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where (codigofactura=" & rs!COD & a & ") and _iva=10.5") / z, 2), "standard")
                    rptImpresionFacturaVentaAT.txtIvaP2 = Format$(10.5, "standard")
                End If
                rptImpresionFacturaVentaAT.txtIvaP = Format$(s2n(0.21, 4) * 100, "standard") 'Format$(s2n(PORCENTAJE, 4) * 100, "standard")
                rptImpresionFacturaVentaAT.txtsub = Format$(s2n(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where (codigofactura=" & rs!COD & a & ") and _iva=21") / z, 2), "standard") 'Format$(s2n(RS!Neto / z, 2), "standard")'Format$(s2n(RS!Neto / z, 2), "standard")
                rptImpresionFacturaVentaAT.txtneto = Format$(s2n(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where (codigofactura=" & rs!COD & a & ") and _iva=21") / z, 2), "standard") 'Format$(s2n(RS!Neto / z, 2), "standard")
                
            End If
        
        Else

            If Left(tdoc, 2) = "NC" Then
                
                rptImpresionFacturaVentaAT.lblcomp.Visible = True
'                rptImpresionFacturaVentaAT.lbltachar = "XXXXXXXXX"
'                STR = "select cantidad,descripcion,(preciounitario + (preciounitario * " & x2s(PORCENTAJE) & ")) / " & x2s(z) & " as punit,(preciototal + (preciototal * " & x2s(PORCENTAJE) & ")) / " & x2s(z) & " as ptot from facturaventadetalle where codigofactura=" & COD & "ORDER BY id"
                str = "select cantidad,descripcion,preciounitario / " & x2s(z) & " as punit,preciototal / " & x2s(z) & " as ptot from facturaventadetalle where codigofactura=" & COD & " ORDER BY id"
                
                If tdoc = "NCB" Then
                    rptImpresionFacturaVentaAT.lblcomp = "Nota de Credito B"
                    str = "select cantidad,descripcion,(preciounitario + (preciounitario * " & x2s(PORCENTAJE) & ")) / " & x2s(z) & " as punit,(preciototal + (preciototal * " & x2s(PORCENTAJE) & ")) / " & x2s(z) & " as ptot from facturaventadetalle where codigofactura=" & COD & "ORDER BY id"
                ElseIf tdoc = "NCE" Then
                    rptImpresionFacturaVentaAT.lblcomp = "Nota de Credito E"
                ElseIf tdoc = "NCA" Then
                    rptImpresionFacturaVentaAT.lblcomp = "Nota de Credito A"
'                    STR = "select cantidad,descripcion,(preciounitario + (preciounitario * 0)) / " & x2s(z) & " as punit,(preciototal + (preciototal * 0)) / " & x2s(z) & " as ptot from facturaventadetalle where codigofactura=" & COD & "ORDER BY id"
                    str = "select cantidad,descripcion,preciounitario / " & x2s(z) & " as punit,preciototal / " & x2s(z) & " as ptot from facturaventadetalle where codigofactura=" & COD & "ORDER BY id"
                    'str = "select descripcion from facturaventadetalle where codigofactura=" & cod
                    rptImpresionFacturaVentaAT.txtivains = Format$(s2n(rs!Neto / z, 2) * s2n(rs!PorcentajeIva, 4), "standard")
                    rptImpresionFacturaVentaAT.txtneto = Format$(s2n(rs!Neto / z, 2), "standard")
                    rptImpresionFacturaVentaAT.txtsub = Format$(s2n(rs!Neto / z, 2), "standard")
                    
                    rptImpresionFacturaVentaAT.txtneto2 = "" 'Format$(s2n(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where codigofactura=" & RS!COD & " and _iva=10.5") / z, 2), "standard") 'Format$(s2n(RS!Neto / z, 2), "standard")
                    rptImpresionFacturaVentaAT.txtsub2 = "" 'Format$(s2n(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where codigofactura=" & RS!COD & " and _iva=10.5") / z, 2), "standard") 'Format$(s2n(RS!Neto / z, 2), "standard")
                    
                    Subtot = Format$(s2n(rs!Neto / z, 2), "standard")
                    rptImpresionFacturaVentaAT.txtivains2 = Format$(s2n(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where (codigofactura=" & rs!COD & a & ") and _iva=10.5") / z, 2) * s2n(0.105, 4), "standard")
                    rptImpresionFacturaVentaAT.txtivains = Format$(s2n(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where (codigofactura=" & rs!COD & a & ") and _iva=21.0") / z, 2) * s2n(0.21, 4), "standard")
                    If rptImpresionFacturaVentaAT.txtivains2 = 0 Then rptImpresionFacturaVentaAT.txtivains2 = ""
                    rptImpresionFacturaVentaAT.txtIvaP = Format$(s2n(0.21, 4) * 100, "standard") 'Format$(s2n(PORCENTAJE, 4) * 100, "standard")
                    If rptImpresionFacturaVentaAT.txtivains2 = "" Then
                        rptImpresionFacturaVentaAT.txtIvaP2 = ""
                    Else
                        rptImpresionFacturaVentaAT.txtIvaP2 = Format$(10.5, "standard")
                    End If
                End If
            Else
                'If Trim(rs!TIPODOC) = "NDA" Or Trim(rs!TIPODOC) = "NDB" Then
                If Left(tdoc, 2) = "ND" Then
                    
                    rptImpresionFacturaVentaAT.lblcomp.Visible = True
'                    rptImpresionFacturaVentaAT.lbltachar = "XXXXXXXXX"
                    If tdoc = "NDB" Then
                        rptImpresionFacturaVentaAT.lblcomp = "Nota de Debito B"
                        str = "select cantidad,descripcion,(preciounitario + (preciounitario * " & x2s(PORCENTAJE) & ")) / " & x2s(z) & " as punit,(preciototal + (preciototal * " & x2s(PORCENTAJE) & ")) / " & x2s(z) & " as ptot from facturaventadetalle where codigofactura=" & COD & " ORDER BY id"
                        'str = "select descripcion from facturaventadetalle where codigofactura=" & cod
                    ElseIf tdoc = "NDE" Then
                        rptImpresionFacturaVentaAT.lblcomp = "Nota de Debito E"
                        str = "select cantidad,descripcion,(preciounitario + (preciounitario * " & x2s(PORCENTAJE) & ")) / " & x2s(z) & " as punit,(preciototal + (preciototal * " & x2s(PORCENTAJE) & ")) / " & x2s(z) & " as ptot from facturaventadetalle where codigofactura=" & COD & " ORDER BY id"
                        
                    ElseIf tdoc = "NDA" Then
                        rptImpresionFacturaVentaAT.lblcomp = "Nota de Debito A"
'                        STR = "select cantidad,descripcion,(preciounitario + (preciounitario * " & x2s(PORCENTAJE) & ")) / " & x2s(z) & " as punit,(preciototal + (preciototal * " & x2s(PORCENTAJE) & ")) / " & x2s(z) & " as ptot from facturaventadetalle where codigofactura=" & COD & "ORDER BY id"
                        str = "select cantidad,descripcion,preciounitario / " & x2s(z) & " as punit,preciototal / " & x2s(z) & " as ptot from facturaventadetalle where codigofactura=" & COD & " and producto<>'1' ORDER BY id"
                        'str = "select descripcion from facturaventadetalle where codigofactura=" & cod
                        If chq = False Then
                            rptImpresionFacturaVentaAT.txtivains = Format$(s2n(rs!Neto / z, 2) * s2n(rs!PorcentajeIva, 4), "standard")
                            rptImpresionFacturaVentaAT.txtneto = Format$(s2n(rs!Neto / z, 2), "standard")
                            rptImpresionFacturaVentaAT.txtsub = Format(s2n(rs!Neto / z, 2), "standard")
                            
                            rptImpresionFacturaVentaAT.txtneto2 = "" 'Format$(s2n(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where codigofactura=" & RS!COD & " and _iva=10.5") / z, 2), "standard") 'Format$(s2n(RS!Neto / z, 2), "standard")
                            rptImpresionFacturaVentaAT.txtsub2 = "" 'Format$(s2n(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where codigofactura=" & RS!COD & " and _iva=10.5") / z, 2), "standard")  'Format(s2n(RS!Neto / z, 2), "standard")
                            
                            Subtot = Format$(s2n(rs!Neto / z, 2), "standard")
                            rptImpresionFacturaVentaAT.txtivains2 = Format$(s2n(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where (codigofactura=" & rs!COD & a & ") and _iva=10.5") / z, 2) * s2n(0.105, 4), "standard")
                            rptImpresionFacturaVentaAT.txtivains = Format$(s2n(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where (codigofactura=" & rs!COD & a & ") and _iva=21.0") / z, 2) * s2n(0.21, 4), "standard")
                            If rptImpresionFacturaVentaAT.txtivains2 = 0 Then rptImpresionFacturaVentaAT.txtivains2 = ""
                            rptImpresionFacturaVentaAT.txtIvaP = Format$(s2n(0.21, 4) * 100, "standard") 'Format$(s2n(PORCENTAJE, 4) * 100, "standard")
                            If rptImpresionFacturaVentaAT.txtivains2 = "" Then
                                rptImpresionFacturaVentaAT.txtivains2 = ""
                            Else
                                rptImpresionFacturaVentaAT.txtIvaP2 = Format$(10.5, "standard")
                            End If
                            
                        Else
                            'es nota de debito por cheque rechazado
                            sql = "select producto,cantidad,descripcion,preciounitario as punit,preciototal as ptot,_iva as iva from facturaventadetalle where codigofactura=" & COD & " and producto<>'1' ORDER BY id"
                            rsnro.Open sql, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
                            If rsnro.RecordCount > 1 Then
                                While Not rsnro.EOF
                                    If Trim(rsnro!producto) = "2" Then 'gasto administrativo
                                        rptImpresionFacturaVentaAT.txtivains = Format$(s2n(rsnro!ptot, 2) * s2n(rsnro!Iva / 100, 4), "standard")
                                        rptImpresionFacturaVentaAT.txtneto = Format$(s2n(rsnro!ptot / z, 2), "standard")
                                        rptImpresionFacturaVentaAT.txtsub = Format(s2n(rsnro!ptot / z, 2), "standard")
                                        rptImpresionFacturaVentaAT.txtIvaP = Format$(s2n(rsnro!Iva / 100, 4) * 100, "standard") 'Format$(s2n(PORCENTAJE, 4) * 100, "standard")
                                    ElseIf CLng(rsnro!producto) > 20 Then
                                        rptImpresionFacturaVentaAT.txtneto2 = rsnro!ptot 'Format$(s2n(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where codigofactura=" & RS!COD & " and _iva=10.5") / z, 2), "standard") 'Format$(s2n(RS!Neto / z, 2), "standard")
                                        rptImpresionFacturaVentaAT.txtsub2 = rsnro!ptot 'Format$(s2n(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where codigofactura=" & RS!COD & " and _iva=10.5") / z, 2), "standard")  'Format(s2n(RS!Neto / z, 2), "standard")
                                        Subtot = Format$(s2n(rsnro!ptot / z, 2), "standard")
                                        rptImpresionFacturaVentaAT.txtivains2 = rsnro!ptot * (rsnro!Iva / 100) 'Format$(s2n(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where (codigofactura=" & rs!COD & a & ") and _iva=10.5") / z, 2) * s2n(0.105, 4), "standard")
                                        If rptImpresionFacturaVentaAT.txtivains2 = 0 Then rptImpresionFacturaVentaAT.txtivains2 = ""
                                        If rptImpresionFacturaVentaAT.txtivains2 = "" Then
                                            rptImpresionFacturaVentaAT.txtivains2 = ""
                                        Else
                                            rptImpresionFacturaVentaAT.txtIvaP2 = Format$(rsnro!Iva / 100, "standard")
                                        End If
                                    End If
                                    rsnro.MoveNext
                                Wend
                            Else
                                rptImpresionFacturaVentaAT.txtivains = Format$(s2n(rsnro!ptot, 2) * s2n(rsnro!Iva / 100, 4), "standard")
                                rptImpresionFacturaVentaAT.txtneto = Format$(s2n(rsnro!ptot / z, 2), "standard")
                                rptImpresionFacturaVentaAT.txtsub = Format(s2n(rsnro!ptot / z, 2), "standard")
                                rptImpresionFacturaVentaAT.txtIvaP = Format$(s2n(rsnro!Iva / 100, 4) * 100, "standard") 'Format$(s2n(PORCENTAJE, 4) * 100, "standard")
                                Subtot = Format$(s2n(rsnro!ptot / z, 2), "standard")
                                rptImpresionFacturaVentaAT.txtneto2 = "" 'Format$(s2n(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where codigofactura=" & RS!COD & " and _iva=10.5") / z, 2), "standard") 'Format$(s2n(RS!Neto / z, 2), "standard")
                                rptImpresionFacturaVentaAT.txtsub2 = "" 'Format$(s2n(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where codigofactura=" & RS!COD & " and _iva=10.5") / z, 2), "standard")  'Format(s2n(RS!Neto / z, 2), "standard")
                                rptImpresionFacturaVentaAT.txtIvaP2 = ""
                                rptImpresionFacturaVentaAT.txtivains2 = ""
                            End If

                        End If
                    End If
                End If
            End If
            rptImpresionFacturaVentaAT.lblfactura2.Visible = True
            rptImpresionFacturaVentaAT.lblfactura.Visible = False
        End If
'*****************************************************************************************
                
        sql = ""
        Set rsnro = Nothing
        sql = "select distinct nroremito from facturaventadetalle where codigofactura=" & COD
        rsnro.Open sql, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
        Do While Not rsnro.EOF
         If rsnro!NroRemito <> 0 Then
            FormatNro = String(6 - Len(rsnro!NroRemito), "0") & rsnro!NroRemito
            rptImpresionFacturaVentaAT.Remitos.Text = RptImpresionFacturaVenta.Remitos.Text & FormatNro & " - "
         Else
            rptImpresionFacturaVentaAT.Remitos.Text = ""
         End If
        rsnro.MoveNext
        Loop
        If rptImpresionFacturaVentaAT.Remitos.Text <> "" Then
          rptImpresionFacturaVentaAT.Remitos.Text = Mid(rptImpresionFacturaVentaAT.Remitos.Text, 1, Len(rptImpresionFacturaVentaAT.Remitos.Text) - 2)
        Else
          rptImpresionFacturaVentaAT.Remitos.Text = ""
        End If
        rsnro.Close
                
'*****************************************************************************************
        If Not IsNull(rs!direccion) Then
            rptImpresionFacturaVentaAT.lbldomicilio = rs!direccion
        Else
            rptImpresionFacturaVentaAT.lbldomicilio = ""
        End If
'        rptImpresionFacturaVentaAT.lblfactura = "0001-" & Format(rs!NroFactura, "00000000") '"0001-" & Format(rs!nrofactura, "00000000")
        rptImpresionFacturaVentaAT.lblfactura = rs!PuntoVenta & "-" & Format(rs!NroFactura, "00000000") '"0001-" & Format(rs!nrofactura, "00000000")
'        rptImpresionFacturaVentaAT.lblfactura2 = "0001-" & Format(rs!NroFactura, "00000000")
        rptImpresionFacturaVentaAT.lblfactura2 = rs!PuntoVenta & "-" & Format(rs!NroFactura, "00000000")
        rptImpresionFacturaVentaAT.lblfecha = rs!Fecha
        rptImpresionFacturaVentaAT.LblIVA = rs!Iva
        If Not IsNull(rs!Localidad) Then
            'rptImpresionFacturaVentaAT.lbllocalidad = rs!Localidad
            rptImpresionFacturaVentaAT.lbldomicilio = rptImpresionFacturaVentaAT.lbldomicilio & "-" & rs!Localidad
        Else
            'rptImpresionFacturaVentaAT.lbllocalidad = ""
            rptImpresionFacturaVentaAT.lbldomicilio = rptImpresionFacturaVentaAT.lbldomicilio
        End If
        If Not IsNull(rs!Provincia) Then
            If Len(rs!Provincia) > 1 Then
                rptImpresionFacturaVentaAT.lbldomicilio = rptImpresionFacturaVentaAT.lbldomicilio & "-" & rs!Provincia
            Else
                rptImpresionFacturaVentaAT.lbldomicilio = rptImpresionFacturaVentaAT.lbldomicilio & "-" & obtenerDeSQL("select descripcion from provincias where codigo='" & Trim(rs!Provincia) & "'")
            End If
        Else
            rptImpresionFacturaVentaAT.lbldomicilio = rptImpresionFacturaVentaAT.lbldomicilio
        End If
        If Triplicado = True Then
            rptImpresionFacturaVentaAT.lblTel.Visible = True
            rptImpresionFacturaVentaAT.lblTel = sSinNull(rs!Telefono)
        End If
        rptImpresionFacturaVentaAT.lblCli = nSinNull(rs!cliente)
        rptImpresionFacturaVentaAT.lblOc = sSinNull(rs!oc)
        rptImpresionFacturaVentaAT.Remitos = sSinNull(rs!remi)
        usuario = UsuarioActual
        rptImpresionFacturaVentaAT.lblInicial = sSinNull(obtenerDeSQL("select inicial from usuarios where codigo=" & rs!vend)) & "/" & sSinNull(obtenerDeSQL("select inicial from usuarios where codigo=" & usuario))
        If rs!Remito <> 0 Then
            rptImpresionFacturaVentaAT.lblref = "Remito"
            rptImpresionFacturaVentaAT.lblnroref = "0001-" & Format(rs!Remito, "00000000")
        Else
            If rs!Pedido <> 0 Then
                rptImpresionFacturaVentaAT.lblref = "Pedido"
                rptImpresionFacturaVentaAT.lblnroref = "0001-" & Format(rs!Pedido, "00000000")
            End If
        End If
        rptImpresionFacturaVentaAT.lblpago = rs!pago
        If sig <> "" Then
            rptImpresionFacturaVentaAT.lblimp = "CONTINUA EN LA " & tdoc & " " & sig
        Else
            rptImpresionFacturaVentaAT.lblimp = ObtenerDescripcion("Monedas", mone) & ": " & enletras(s2n(rs!Total / z))
        End If
        rptImpresionFacturaVentaAT.txttotalfinal = Format$(s2n(rs!Total / z), "standard")
        rptImpresionFacturaVentaAT.txttotalfinal = IIf(mone = 1 Or mone = -1, "$  ", "U$S  ") & rptImpresionFacturaVentaAT.txttotalfinal
        If tdoc = "FAE" Then
            'VER CUAL ES LA LEYENDA
            If mone <> 1 Then
                rptImpresionFacturaVentaAT.lblleyenda = " Equivalente a " & x2s(rs!Total) & " Pesos al tipo de cambio " & x2s(z) & " pesos por " & ObtenerDescripcion("Monedas", mone)
            End If
        Else
            rptImpresionFacturaVentaAT.lblleyenda = "Equivalente a dolares estadounidenses U$S          ." & Chr(13) _
            & "Al tipo de cambio            peso/s por dolar segun clausula al pie." & Chr(13) & "El pago de la presente deberá realizarse en dolares estadounidenses a su vencimiento," _
            & "conforme al valor en dicha moneda expresado en este formulario.El comprador asume que el precio en dolares" _
            & " ha sido condición esencial de esta venta renunciando a invocar el Art 119A de Código Civil." & vbCrLf _
            & "En caso que el pago no pueda realizarse en dicha moneda se realizará en pesos al tipo de cambio vigente para el dolar estadounidense" _
            & " tomando la cotización de tipo vendedor del Banco de la Nación Argentina, al cierre de operaciones del día de efectivo pago; " _
            & "en caso que a la fecha de pago no existiera mercado Libre de cambios en la Ciudad de Buenos Aires se tomaráan las cotizaciones" _
            & " en el Mercado de Nueva York o Montevideo. A opción del vendedor la falta de pago al vencimiento constituye al comprador en mora de " _
            & " pleno derecho y hará devengar un interés punitorio del 20% anual hasta el efectivo pago."
        End If
        
        rptImpresionFacturaVentaAT.DataControl1.Connection = DataEnvironment1.Sistema
        rptImpresionFacturaVentaAT.DataControl1.Source = str
        
    End If
    
    If gEMPR_ImprimeCertCalidad Then

        Dim Consulta As String
        Dim cCertificadoCalidad As String
        Dim cFecha As Date
        Dim cCodCli As String
        Dim cRZ As String
        Dim cCodigoProd As String
        Dim cDescProd As String
        Dim cCantidad As String
        Dim cNroRemito As String
        Dim cMuestra As String
        
        
        
        Consulta = "SELECT FacturaVentaDetalle.*, Producto.descripcion" _
            & " FROM FacturaVentaDetalle INNER JOIN Producto ON FacturaVentaDetalle.Producto = Producto.codigo where facturaventadetalle.codigofactura=" & codigo

        rs1.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs1.EOF Then
            DataEnvironment1.Sistema.Execute "Delete TEMP_CONTROL_CALIDAD"
            cFecha = Date
            If Not IsNull(rs!RAZONSOCIAL) Then cRZ = rs!RAZONSOCIAL
            Do While Not rs1.EOF
                
                rsnro.Open "Select certificadocalidad from bs", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
                If Not rsnro.EOF Then
                    If Not IsNull(rsnro!certificadocalidad) Then
                        cCertificadoCalidad = rsnro!certificadocalidad + 1
                        DataEnvironment1.Sistema.Execute "update bs set certificadocalidad=certificadocalidad+1"
                    End If
                End If
                rsnro.Close
                Set rsnro = Nothing
                
                If Not IsNull(rs1!codPropio) And Not IsNull(rs1!producto) Then
                    cCodigoProd = VerProductoCliente(rs1!producto, rs1!codPropio, rs!codigo)
                    cDescProd = rs1!DESCRIPCION
                End If
                
                If Not IsNull(rs1!cantidad) Then cCantidad = rs1!cantidad
                
                If rs!Remito <> 0 Then
                    cNroRemito = "0001-" & Format(rs!Remito, "00000000")
                End If
                
                RptControldeCalidad.lblmuestra.caption = ""
                
                Consulta = "Insert Into TEMP_CONTROL_CALIDAD (CERTIFICADO_CALIDAD, FECHA, RAZON_SOCIAL, CODIGO_PROD, " & _
                                                            "DESCRIPCION_PROD, CANTIDAD, NRO_REMITO, MUESTRA) " & _
                            "Values ('" & Trim(cCertificadoCalidad) & "'," & _
                                    ssFecha(cFecha) & ", " & _
                                    "'" & Trim(cRZ) & "', " & _
                                    "'" & Trim(cCodigoProd) & "', " & _
                                    "'" & Trim(cDescProd) & "', " & _
                                    "'" & Trim(cCantidad) & "', " & _
                                    "'" & Trim(cNroRemito) & "', " & _
                                    "'" & Trim(cMuestra) & "')"
                DataEnvironment1.Sistema.Execute Consulta
                rs1.MoveNext
            Loop
        End If
        rs1.Close
        Set rs1 = Nothing
                        
        Consulta = "Select * From TEMP_CONTROL_CALIDAD Order By ID"
        With RptControldeCalidad
            .Data.Connection = DataEnvironment1.Sistema
            .Data.Source = Consulta
            
            .fieFecha.DataField = "FECHA"
            .fieCertificado.DataField = "CERTIFICADO"
            .fieCliente.DataField = "RAZON_SOCIAL"
            .fieProducto.DataField = "DESCRIPCION_PROD"
            .fieCodCliente.DataField = "CODIGO_PROD"
            
            .fieCantidad.DataField = "CANTIDAD"
            .fieNroRemito.DataField = "NRO_REMITO"
            .fieMuestra.DataField = "MUESTRA"
            .Show
        End With
    End If
    rs.Close
    Set rs = Nothing
    
    'rptImpresionFacturaVentaAT.PageSettings.TopMargin = margenTop_FV()
    If Subtot <> 0 Then
        If rptImpresionFacturaVentaAT.txtneto <> "" And rptImpresionFacturaVentaAT.txtneto <> "0" Then rptImpresionFacturaVentaAT.Label5.caption = "SUB-TOTAL --------------------------------------------> " & s2n(Subtot, 2, True) 'rptImpresionFacturaVentaAT.txtneto
        If chq = True And rptImpresionFacturaVentaAT.txtneto = "" And rptImpresionFacturaVentaAT.txtneto2 <> "" And rptImpresionFacturaVentaAT.txtneto2 <> "0" Then rptImpresionFacturaVentaAT.Label5.caption = "SUB-TOTAL --------------------------------------------> " & s2n(Subtot, 2, True) 'rptImpresionFacturaVentaAT.txtneto
    Else
        rptImpresionFacturaVentaAT.Line1.Visible = False
    End If
    
    
    If Triplicado = True Then
        rptImpresionFacturaVentaAT.Printer.Copies = 3
    Else
        rptImpresionFacturaVentaAT.Printer.Copies = 2
    End If
    rptImpresionFacturaVentaAT.Printer.PaperSize = vbPRPSA4 'hoja A4=9
    rptImpresionFacturaVentaAT.Show vbModal
    
    
    
    'RptImpresionFacturaVenta.PrintReport True
    
End Function


Public Function ImprimirComprobanteLOC(codigo) As Boolean

    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim rsnro As New ADODB.Recordset
    Dim str As String
    Dim COD As Long
    Dim PORCENTAJE As String
    Dim tdoc As String, z As Double, mone As Long
    
    rs.Open "SELECT FacturaVenta.*, Clientes.*, FormasPago.descripcion as pago, Ivas.descripcion as iva, Ivas.letra as letra" _
        & " FROM FormasPago INNER JOIN (Ivas INNER JOIN (Clientes INNER JOIN FacturaVenta ON Clientes.codigo = FacturaVenta.Cliente) ON Ivas.codigo = FacturaVenta.TipoIVA) ON FormasPago.codigo = FacturaVenta.FormaPago where facturaventa.codigo=" & codigo, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    
    tdoc = Trim(rs!TIPODOC)
    z = s2n(rs!cotizacion, 4)
    If z = 0 Then z = 1
    mone = rs!moneda
    If mone = 0 Then mone = 1 '(PESOS)
    
    If Not rs.EOF Then
        RptImpresionFacturaVentaLoc.lblcliente = rs!RAZONSOCIAL
        If Not IsNull(rs!CUIT) Then
            RptImpresionFacturaVentaLoc.lblCuit = rs!CUIT
        End If
        COD = codigo
        If Not IsNull(rs!PorcentajeIva) Then
         PORCENTAJE = rs!PorcentajeIva
        End If

        If Left(tdoc, 2) = "FA" Then
            
            RptImpresionFacturaVentaLoc.lblcomp = "Factura"
            If tdoc = "FAB" Then
                str = "select cantidad,descripcion,(preciounitario + (preciounitario * " & x2s(PORCENTAJE) & ")) / " & x2s(z) & " as punit,(preciototal + (preciototal * " & x2s(PORCENTAJE) & ")) / " & x2s(z) & "  as ptot from facturaventadetalle where codigofactura=" & COD
            ElseIf tdoc = "FAE" Then
                str = "select cantidad,descripcion,preciounitario / " & x2s(z) & " as punit ,preciototal / " & x2s(z) & " as ptot from facturaventadetalle where codigofactura=" & COD & "ORDER BY id"
            
            ElseIf tdoc = "FAA" Then
                str = "select cantidad,descripcion,preciounitario / " & x2s(z) & " as punit ,preciototal / " & x2s(z) & " as ptot from facturaventadetalle where codigofactura=" & COD & "ORDER BY id"

                RptImpresionFacturaVentaLoc.txtivains = Format$(s2n(rs!Neto / z, 2) * s2n(rs!PorcentajeIva, 4), "standard")
                RptImpresionFacturaVentaLoc.txtneto = Format$(s2n(rs!Neto / z, 2), "standard")
                RptImpresionFacturaVentaLoc.txtsub = Format$(s2n(rs!Neto / z, 2), "standard")
                RptImpresionFacturaVentaLoc.txtIvaP = Format$(s2n(PORCENTAJE, 4) * 100, "standard")
            End If
        
        Else

            If Left(tdoc, 2) = "NC" Then
                RptImpresionFacturaVentaLoc.lblcomp = "Nota de Credito"
                RptImpresionFacturaVentaLoc.lbltachar = "XXXXXXXXX"
                str = "select cantidad,descripcion,(preciounitario + (preciounitario * " & x2s(PORCENTAJE) & ")) / " & x2s(z) & " as punit,(preciototal + (preciototal * " & x2s(PORCENTAJE) & ")) / " & x2s(z) & " as ptot from facturaventadetalle where codigofactura=" & COD & "ORDER BY id"
                
                If tdoc = "NCB" Then

                ElseIf tdoc = "NCE" Then

                ElseIf tdoc = "NCA" Then

                    'str = "select descripcion from facturaventadetalle where codigofactura=" & cod
                    RptImpresionFacturaVentaLoc.txtivains = Format$(s2n(rs!Neto / z, 2) * s2n(rs!PorcentajeIva, 4), "standard")
                    RptImpresionFacturaVentaLoc.txtneto = Format$(s2n(rs!Neto / z, 2), "standard")
                    RptImpresionFacturaVentaLoc.txtsub = Format$(s2n(rs!Neto / z, 2), "standard")
                End If
            Else
                'If Trim(rs!TIPODOC) = "NDA" Or Trim(rs!TIPODOC) = "NDB" Then
                If Left(tdoc, 2) = "ND" Then
                    RptImpresionFacturaVentaLoc.lblcomp = "Nota de Debito"
                    RptImpresionFacturaVentaLoc.lbltachar = "XXXXXXXXX"
                    If tdoc = "NDB" Then
                        str = "select cantidad,descripcion,(preciounitario + (preciounitario * " & x2s(PORCENTAJE) & ")) / " & x2s(z) & " as punit,(preciototal + (preciototal * " & x2s(PORCENTAJE) & "))/ " & x2s(z) & " as ptot from facturaventadetalle where codigofactura=" & COD & "ORDER BY id"
                        'str = "select descripcion from facturaventadetalle where codigofactura=" & cod
                    ElseIf tdoc = "NDE" Then
                        str = "select cantidad,descripcion,(preciounitario + (preciounitario * " & x2s(PORCENTAJE) & ")) / " & x2s(z) & " as punit,(preciototal + (preciototal * " & x2s(PORCENTAJE) & "))/ " & x2s(z) & " as ptot from facturaventadetalle where codigofactura=" & COD & "ORDER BY id"
                        
                    ElseIf tdoc = "NDA" Then
                        str = "select cantidad,descripcion,(preciounitario + (preciounitario * " & x2s(PORCENTAJE) & ")) / " & x2s(z) & " as punit,(preciototal + (preciototal * " & x2s(PORCENTAJE) & "))/ " & x2s(z) & " as ptot from facturaventadetalle where codigofactura=" & COD & "ORDER BY id"
                        'str = "select descripcion from facturaventadetalle where codigofactura=" & cod
                        RptImpresionFacturaVentaLoc.txtivains = Format$(s2n(rs!Neto / z, 2) * s2n(rs!PorcentajeIva, 4), "standard")
                        RptImpresionFacturaVentaLoc.txtneto = Format$(s2n(rs!Neto / z, 2), "standard")
                        RptImpresionFacturaVentaLoc.txtsub = Format(s2n(rs!Neto / z, 2), "standard")
                    End If
                End If
             End If
        End If
'*****************************************************************************************
        Dim FormatNro, sql As String
        
        
        sql = "select distinct nroremito from facturaventadetalle where codigofactura=" & COD
        rsnro.Open sql, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
        Do While Not rsnro.EOF
         If rsnro!NroRemito <> 0 Then
            FormatNro = String(6 - Len(rsnro!NroRemito), "0") & rsnro!NroRemito
            RptImpresionFacturaVentaLoc.Remitos.Text = RptImpresionFacturaVentaLoc.Remitos.Text & FormatNro & " - "
         Else
            RptImpresionFacturaVentaLoc.Remitos.Text = ""
         End If
        rsnro.MoveNext
        Loop
        If RptImpresionFacturaVentaLoc.Remitos.Text <> "" Then
          RptImpresionFacturaVentaLoc.Remitos.Text = Mid(RptImpresionFacturaVentaLoc.Remitos.Text, 1, Len(RptImpresionFacturaVentaLoc.Remitos.Text) - 2)
        Else
          RptImpresionFacturaVentaLoc.Remitos.Text = ""
        End If
        rsnro.Close
                
'*****************************************************************************************
        If Not IsNull(rs!direccion) Then
            RptImpresionFacturaVentaLoc.lbldomicilio = rs!direccion
        Else
            RptImpresionFacturaVentaLoc.lbldomicilio = ""
        End If
        RptImpresionFacturaVentaLoc.lblfactura = Format(obtenerDeSQL("select sucursal from datosempresa"), "0000") & "-" & Format(rs!NroFactura, "00000000") '"0001-" & Format(rs!nrofactura, "00000000")
        RptImpresionFacturaVentaLoc.lblfecha = rs!Fecha
        RptImpresionFacturaVentaLoc.LblIVA = rs!Iva
        If Not IsNull(rs!Localidad) Then
            RptImpresionFacturaVentaLoc.lbllocalidad = rs!Localidad
        Else
            RptImpresionFacturaVentaLoc.lbllocalidad = ""
        End If
        If rs!Remito <> 0 Then
            RptImpresionFacturaVentaLoc.lblref = "Remito"
            RptImpresionFacturaVentaLoc.lblnroref = Format(obtenerDeSQL("select sucursal from datosempresa"), "0000") & "-" & Format(rs!Remito, "00000000") '"0001-" & Format(rs!remito, "00000000")
        Else
            If rs!Pedido <> 0 Then
                RptImpresionFacturaVentaLoc.lblref = "Pedido"
                RptImpresionFacturaVentaLoc.lblnroref = Format(obtenerDeSQL("select sucursal from datosempresa"), "0000") & "-" & Format(rs!Pedido, "00000000") '"0001-" & Format(rs!Pedido, "00000000")
            End If
        End If
        RptImpresionFacturaVentaLoc.lblpago = rs!pago
        RptImpresionFacturaVentaLoc.lblimp = ObtenerDescripcion("Monedas", mone) & ": " & enletras(s2n(rs!Total / z))
        RptImpresionFacturaVentaLoc.txttotalfinal = Format$(s2n(rs!Total / z), "standard")
        If tdoc = "FAE" Then
            'VER CUAL ES LA LEYENDA
            If mone <> 1 Then
                RptImpresionFacturaVentaLoc.lblleyenda = " Equivalente a " & x2s(rs!Total) & " Pesos al tipo de cambio " & x2s(z) & " pesos por " & ObtenerDescripcion("Monedas", mone)
            End If
        Else
            RptImpresionFacturaVentaLoc.lblleyenda = "Equivalente a dolares estadounidenses U$S          ." & Chr(13) _
            & "Al tipo de cambio            peso/s por dolar segun clausula al pie." & Chr(13) & "El pago de la presente deberá realizarse en dolares estadounidenses a su vencimiento," _
            & "conforme al valor en dicha moneda expresado en este formulario.El comprador asume que el precio en dolares" _
            & " ha sido condición esencial de esta venta renunciando a invocar el Art 119A de Código Civil." & vbCrLf _
            & "En caso que el pago no pueda realizarse en dicha moneda se realizará en pesos al tipo de cambio vigente para el dolar estadounidense" _
            & " tomando la cotización de tipo vendedor del Banco de la Nación Argentina, al cierre de operaciones del día de efectivo pago; " _
            & "en caso que a la fecha de pago no existiera mercado Libre de cambios en la Ciudad de Buenos Aires se tomaráan las cotizaciones" _
            & " en el Mercado de Nueva York o Montevideo. A opción del vendedor la falta de pago al vencimiento constituye al comprador en mora de " _
            & " pleno derecho y hará devengar un interés punitorio del 20% anual hasta el efectivo pago."
        End If
        
        RptImpresionFacturaVentaLoc.DataControl1.Connection = DataEnvironment1.Sistema
        RptImpresionFacturaVentaLoc.DataControl1.Source = str
        
    End If
    
    If gEMPR_ImprimeCertCalidad Then

        Dim Consulta As String
        Dim cCertificadoCalidad As String
        Dim cFecha As Date
        Dim cCodCli As String
        Dim cRZ As String
        Dim cCodigoProd As String
        Dim cDescProd As String
        Dim cCantidad As String
        Dim cNroRemito As String
        Dim cMuestra As String
        
        
        
        Consulta = "SELECT FacturaVentaDetalle.*, Producto.descripcion" _
            & " FROM FacturaVentaDetalle INNER JOIN Producto ON FacturaVentaDetalle.Producto = Producto.codigo where facturaventadetalle.codigofactura=" & codigo

        rs1.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs1.EOF Then
            DataEnvironment1.Sistema.Execute "Delete TEMP_CONTROL_CALIDAD"
            cFecha = Date
            If Not IsNull(rs!RAZONSOCIAL) Then cRZ = rs!RAZONSOCIAL
            Do While Not rs1.EOF
                
                rsnro.Open "Select certificadocalidad from bs", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
                If Not rsnro.EOF Then
                    If Not IsNull(rsnro!certificadocalidad) Then
                        cCertificadoCalidad = rsnro!certificadocalidad + 1
                        DataEnvironment1.Sistema.Execute "update bs set certificadocalidad=certificadocalidad+1"
                    End If
                End If
                rsnro.Close
                Set rsnro = Nothing
                
                If Not IsNull(rs1!codPropio) And Not IsNull(rs1!producto) Then
                    cCodigoProd = VerProductoCliente(rs1!producto, rs1!codPropio, rs!codigo)
                    cDescProd = rs1!DESCRIPCION
                End If
                
                If Not IsNull(rs1!cantidad) Then cCantidad = rs1!cantidad
                
                If rs!Remito <> 0 Then
                    cNroRemito = Format(obtenerDeSQL("select sucursal from datosempresa"), "0000") & "-" & Format(rs!Remito, "00000000") '"0001-" & Format(rs!remito, "00000000")
                End If
                
                RptControldeCalidad.lblmuestra.caption = ""
                
                Consulta = "Insert Into TEMP_CONTROL_CALIDAD (CERTIFICADO_CALIDAD, FECHA, RAZON_SOCIAL, CODIGO_PROD, " & _
                                                            "DESCRIPCION_PROD, CANTIDAD, NRO_REMITO, MUESTRA) " & _
                            "Values ('" & Trim(cCertificadoCalidad) & "'," & _
                                    ssFecha(cFecha) & ", " & _
                                    "'" & Trim(cRZ) & "', " & _
                                    "'" & Trim(cCodigoProd) & "', " & _
                                    "'" & Trim(cDescProd) & "', " & _
                                    "'" & Trim(cCantidad) & "', " & _
                                    "'" & Trim(cNroRemito) & "', " & _
                                    "'" & Trim(cMuestra) & "')"
                DataEnvironment1.Sistema.Execute Consulta
                rs1.MoveNext
            Loop
        End If
        rs1.Close
        Set rs1 = Nothing
                        
        Consulta = "Select * From TEMP_CONTROL_CALIDAD Order By ID"
        With RptControldeCalidad
            .Data.Connection = DataEnvironment1.Sistema
            .Data.Source = Consulta
            
            .fieFecha.DataField = "FECHA"
            .fieCertificado.DataField = "CERTIFICADO"
            .fieCliente.DataField = "RAZON_SOCIAL"
            .fieProducto.DataField = "DESCRIPCION_PROD"
            .fieCodCliente.DataField = "CODIGO_PROD"
            
            .fieCantidad.DataField = "CANTIDAD"
            .fieNroRemito.DataField = "NRO_REMITO"
            .fieMuestra.DataField = "MUESTRA"
            .Show
        End With
    End If
    rs.Close
    Set rs = Nothing
    
    'RptImpresionFacturaVentaLoc.PageSettings.TopMargin = margenTop_FV()
    
    RptImpresionFacturaVentaLoc.Printer.Copies = 2
    RptImpresionFacturaVentaLoc.Printer.PaperSize = 9
    RptImpresionFacturaVentaLoc.Show vbModal
    
    
    
    'RptImpresionFacturaVentaLoc.PrintReport True
    
End Function

Public Function LeoPrinters()
'    Dim x As Printer
'    For Each x In Printers
'        MsgBox x.DeviceName
'        If confirma(x.DeviceName) Then
'        RptImpresionFacturaVenta.Printer.Port = x.Port
'        End If
'    Next
End Function

'    Sql = "SELECT MoviCaja.CAJA, MoviCaja.IMPORTE, " & _
'    " MAYOR.Cuenta, CUENTAS.DESCRIPCION, MoviCaja.MOVIMIENTO, MAYOR.haber, " & _
'    " MoviCaja.CONCEPTO, Cajas.sector, MoviCaja.TIPO " & _
'    " FROM Cajas INNER JOIN (((MoviCaja INNER JOIN Asientos ON MoviCaja.idDoc = Asientos.idDoc)" & _
'    " INNER JOIN MAYOR ON Asientos.idAsiento = MAYOR.idAsiento) " & _
'    " INNER JOIN CUENTAS ON MAYOR.Cuenta = CUENTAS.Cuenta) ON Cajas.codigo = MoviCaja.CAJA " & _
'    " WHERE MoviCaja.MOVIMIENTO = " & Trim(Nmov) & " AND MAYOR.HABER > 0"
    
Public Function ImprimirPedido2(NPedido As Double, Optional Ver As Boolean = True) As Boolean
On Error GoTo saltoIMP
Dim rs As New ADODB.Recordset
Dim sql As String
Dim rsempresa As New ADODB.Recordset
Dim x As Variant
Dim msj As String

   sql = "SELECT IPC.CODIGO,IPC.PEDIDO,IPC.PRODUCTO,IPC.CANTIDAD,IPC.FACTURAR,IPC.PRECIO,IPC.FORMULA,IPC.FECHAENTREGA,IPC.ESTADO,IPC.SALDO,IPC.TIPOITEM, dbo.rtf2txt(IPC.descripcion) AS descripcion,Producto.iva, " _
   & " (IPC.precio * IPC.cantidad)as Totaliva, " _
   & " ((IPC.precio * IPC.cantidad) + ((IPC.precio * IPC.cantidad) * producto.iva)) as TotalIvas, " _
   & " ((IPC.precio * IPC.cantidad) * producto.iva) as TotalivaItem " _
   & " FROM ItemPedidoCliente2 IPC " _
   & " left outer JOIN Producto ON IPC.producto = Producto.codigo  " _
   & " WHERE IPC.pedido = " & NPedido & " order by ipc.codigo"
   
   'sql = sql & " Union " & " SELECT IPC.CODIGO,IPC.PEDIDO,IPC.PRODUCTO,IPC.CANTIDAD,IPC.FACTURAR,IPC.PRECIO,IPC.FORMULA,IPC.FECHAENTREGA,IPC.ESTADO,IPC.SALDO,IPC.TIPOITEM, dbo.rtf2txt(IPC.DESCRIPCION) as descripcion ,0 as iva, " _
   & " 0 as Totaliva, " _
   & " 0 as TotalIvas, " _
   & " 0 as TotalivaItem " _
   & " FROM ItemPedidoCliente2 IPC " _
   & "  " _
   & " WHERE IPC.TipoItem='Pie' and IPC.pedido = " & NPedido & " order by IPC.codigo"
   
   rs.Open sql, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
   
   If Not rs.EOF Then
       
    RptPedidoCliente2.NroCliente.Text = frmPresupuesto.uCliente.codigo
    RptPedidoCliente2.NomCliente.Text = frmPresupuesto.uCliente.DESCRIPCION
    RptPedidoCliente2.Dire = frmPresupuesto.txtdireccion & Chr(13) & frmPresupuesto.txtLocalidad & Chr(13) & obtenerDeSQL("select p.descripcion from clientes c left outer join provincias p on p.codigo=c.provincia where c.codigo=" & frmPresupuesto.uCliente.codigo)
    RptPedidoCliente2.Field6.Text = frmPresupuesto.uContacto.DESCRIPCION
    RptPedidoCliente2.Fecha.Text = "Buenos Aires, " & Day(frmPresupuesto.dtFecha) & " de " & PasoMes(Month(frmPresupuesto.dtFecha)) & " de " & Year(frmPresupuesto.dtFecha)
    RptPedidoCliente2.NroPedido.Text = "BH" & Right(Year(frmPresupuesto.dtFecha), 2) & Format(Month(frmPresupuesto.dtFecha), "00") & "-" & frmPresupuesto.txtNro
    RptPedidoCliente2.Total.Text = s2n(frmPresupuesto.lblTotalPedi, 2)
    RptPedidoCliente2.Vendedor.Text = frmPresupuesto.cmbvendedor.Text
    RptPedidoCliente2.vend.Text = obtenerDeSQL("SELECT inicial from usuarios where descripcion='" & Trim(frmPresupuesto.cmbvendedor.Text) & "'")
    RptPedidoCliente2.Emisor.Text = obtenerDeSQL("SELECT letra from emisor where id=" & Trim(frmPresupuesto.uEmisor.codigo))
    RptPedidoCliente2.PedCli.Text = frmPresupuesto.txtnropedidocli
    RptPedidoCliente2.Obser.Text = frmPresupuesto.txtObs
    RptPedidoCliente2.FechaEnt.Text = frmPresupuesto.dtfechaentrega
    RptPedidoCliente2.FPago.Text = frmPresupuesto.cmbformapago.Text
    RptPedidoCliente2.Entrega.Text = frmPresupuesto.txtdireccionentrega.Text
    
    Dim Iva21, iva105 As String

    While Not rs.EOF      'NO SE SI ESTA BIEN PERO ANDA
      If Trim(rs!Iva) <> "0,105" Then
         Iva21 = "IVA 0,21"
         Else
         iva105 = "IVA 0,105"
      End If
      rs.MoveNext
    Wend
    RptPedidoCliente2.LblTipoIva = Iva21 & "  " & iva105
   End If
rs.Close
RptPedidoCliente2.DControl.Connection = DataEnvironment1.Sistema
RptPedidoCliente2.DControl.Source = sql
' CANTIDAD DE COPIAS A IMPRIMIR FALTA DEFINIR
RptPedidoCliente2.Printer.Copies = 1
'RptPedidoCliente.Printer.Copies = 10
'
cierra = False
If Ver = True Then
    RptPedidoCliente2.Show vbModal
End If
'    If gEMPR_idEmpresa = 1 Or gEMPR_idEmpresa = 3 Or gEMPR_idEmpresa = 4 Then
'        If gEMPR_idEmpresa = 4 And Date >= CDate("10/01/2009") Then
'            msj = "¿Desea crear un pdf y enviarlo por correo?"
'        Else
'            msj = "¿Desea imprimir una copia en la segunda impresora definida?"
'        End If
'        If MsgBox(msj, vbYesNo) = vbYes Then
'            If gEMPR_idEmpresa = 4 And Date >= CDate("10/01/2009") Then
'                Dim map As String
                
'                RptPedidoCliente2.Printer.DeviceName = "cutepdf writer" '"adobe pdf" '
                'RptPedidoCliente.Printer.FileName = "C:\pr.pdf"  'este esta bien pero si lo hago automatico no lo puedo abrir
                'RptPedidoCliente.documentName = "C:\pr.pdf"
'                RptPedidoCliente2.PrintReport False
                'frmMail.Show
                'de
                'frmMail.Text1 = obtenerDeSQL("select mfrom from mailfeed where id=2")
                'para
                'frmMail.Text2 = sSinNull(obtenerDeSQL("select mail from prov where codigo=" & uProv.codigo))
                'frmMail.Text3 = obtenerDeSQL("select subject from mailfeed where id=2") & " " & txtcodigo
                
                
'                Dim a
                
'                On Error Resume Next
                 
'                FrmPedidosClientes2.MAPIMail.Compose
                
'                FrmPedidosClientes2.MAPIMail.RecipIndex = 0
'                FrmPedidosClientes2.MAPIMail.RecipType = 1
                'FrmPedidosClientes.MAPIMail.RecipDisplayName = "german.sistemas@betasepp.com.ar" 'sSinNull(obtenerDeSQL("select mail from prov where codigo=" & uProv.codigo)) '"emeil@al.que.lo.envias"
                
''                    MAPIMail.RecipIndex = 1
''                    MAPIMail.RecipType = 2
''                    MAPIMail.RecipDisplayName = "german_dodge@hotmail.com" '"emeil@al.que.lo.envias"
                
''                    MAPIMail.RecipIndex = 2
''                    MAPIMail.RecipType = 3
''                    MAPIMail.RecipDisplayName = "diego@betasepp.com.ar" '"emeil@al.que.lo.envias"
                
                
'                FrmPedidosClientes2.MAPIMail.RecipAddress = "diego.book@locaire.com" 'IIf((sSinNull(obtenerDeSQL("select mail from prov where codigo=" & uProv.codigo))) = "", "correo@correo.com", sSinNull(obtenerDeSQL("select mail from prov where codigo=" & uProv.codigo))) '"emeil@al.que.lo.envias"
'                FrmPedidosClientes2.MAPIMail.AddressResolveUI = True
'                FrmPedidosClientes2.MAPIMail.ResolveName
                'MAPIMail.RecipType = mapToList
                
'                FrmPedidosClientes2.MAPIMail.MsgSubject = "PEDIDO DE CLIENTE Nº" & FrmPedidosClientes.txtNro
'                FrmPedidosClientes2.MAPIMail.MsgNoteText = "Por favor confir me la llegada de este mail, desde ya muchas gracias."
                'FrmPedidosClientes.MAPIMail.AttachmentPathName = "C:\ActiveReports Document.pdf"
                                                        
'                FrmPedidosClientes2.MAPIMail.Send True
                
'            Else
'                Dim veo, actual, segunda
'                Dim obj_Impresora As Object
'                'guardo que impresora esta como predeterminada
'                actual = Printer.DeviceName
'                'preparo el obj para setear la impresora
'                Set obj_Impresora = CreateObject("WScript.Network")
'                'obtengo la segunda impresora detectada y seteo la default
'                segunda = RptPedidoCliente2.Printer.Devices(0)
'                segunda = Trim(obtenerDeSQL("select impresorapedidos from datosempresa"))
'                obj_Impresora.setdefaultprinter segunda
'                'vasio el obj
'                Set obj_Impresora = Nothing
'                'obtengo la nueva predeterminda
'                veo = Printer.DeviceName
'                'para configurar la impresora
'                'veo = RptPedidoCliente.Printer.SetupDialog
'
'                'seteo los datos del reporte
'                RptPedidoCliente2.NroCliente.Text = FrmPedidosClientes.uCliente.codigo
'                RptPedidoCliente2.NomCliente.Text = FrmPedidosClientes.uCliente.DESCRIPCION
'                RptPedidoCliente2.Fecha.Text = FrmPedidosClientes.dtFecha
'                RptPedidoCliente2.NroPedido.Text = FrmPedidosClientes.txtNro
'                RptPedidoCliente2.Total.Text = Format$(FrmPedidosClientes.lblTotalPedi, "standard")
'                RptPedidoCliente2.Vendedor.Text = FrmPedidosClientes.cmbvendedor
'                RptPedidoCliente2.PedCli.Text = FrmPedidosClientes.txtnropedidocli
'                RptPedidoCliente2.Obser.Text = FrmPedidosClientes.txtobs
'                RptPedidoCliente2.FechaEnt.Text = FrmPedidosClientes.dtfechaentrega
'                RptPedidoCliente2.FPago.Text = FrmPedidosClientes.cmbformapago.Text
'                RptPedidoCliente2.Entrega.Text = FrmPedidosClientes.txtdireccionentrega.Text
'                RptPedidoCliente2.DControl.Connection = DataEnvironment1.Sistema
'                RptPedidoCliente2.DControl.Source = sql
'
'                RptPedidoCliente2.Printer.DeviceName = segunda
'                'muestro adonde imprimir
'                MsgBox "La siguiente copia se imprimira en :" & Chr(13) & Chr(13) & veo, vbInformation
'                'imprimo sin preview
'                RptPedidoCliente2.PrintReport False
'                'imprimo con preview
'                'RptPedidoCliente.Show vbModal
                
'                'vuelvoa setear la impresora default con la que estaba antes
'                Set obj_Impresora = CreateObject("WScript.Network")
'                obj_Impresora.setdefaultprinter actual
'                Set obj_Impresora = Nothing
'            End If
'        End If
'    End If
    If Ver = True Then
        cierra = True
        Unload RptPedidoCliente2
    End If
Exit Function
saltoIMP:
MsgBox "La impresora no esta preparada por algun motivo.", vbCritical, "Impresora Alternativa"
End Function

Public Function RemitoAjuste(codigo)
Dim sql, Sql2 As String
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim CantBulto As Double

sql = "SELECT DISTINCT i.prod as producto, p.descripcion, s.serie, i.*, i.remajuste, r.*,d.cantidad as cantidad2,d.facturar as saldo" & _
" FROM Producto AS p INNER JOIN (remitoajuste AS r INNER JOIN " & _
"(Series AS s RIGHT JOIN remitoajustedetalle AS i ON " & _
"(s.nrocomprobante = i.remajuste) AND (s.producto = i.prod)) " & _
"ON r.codigo = i.remajuste) ON p.codigo = i.prod inner join remitoventadetalle d on d.codigo=i.itemremventa  WHERE " & _
"(((s.comprobante)=9 Or (s.comprobante) Is Null)) AND r.codigo=" & codigo & " ORDER BY i.prod DESC"

RptRemAj.CancPNro = frmRemitoCancelacion.lblNumero
RptRemAj.PedOriginal = frmRemitoCancelacion.Label4 & "-" & frmRemitoCancelacion.Label5  'frmRemitoCancelacion.uRemito.codigo
RptRemAj.ClienteNro = frmRemitoCancelacion.uCliente.codigo
RptRemAj.Empresa = frmRemitoCancelacion.uCliente.DESCRIPCION
RptRemAj.Fecha = frmRemitoCancelacion.uFecha.Value
RptRemAj.Field9 = obtenerDeSQL("select fecha from remitoventa where codigo=" & frmRemitoCancelacion.uRemito.codigo)
RptRemAj.DControl.Connection = DataEnvironment1.Sistema
RptRemAj.DControl.Source = sql

'Sql2 = "SELECT remitoventa.numero,remitoventa.transporte FROM remitoventa WHERE codigo = " & frmRemitoCancelacion.uRemito.codigo & ""
'rs.Open Sql2, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
'rs2.Open "Select descripcion FROM transportes WHERE codigo = " & rs!Transporte & "", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
'If Not rs.EOF Then
'   'RptRemAj.Vendedor.Text = rs!codigo & " " & rs!DESCRIPCION
'   RptRemAj.Transporte.Text = rs!Transporte & " " & rs2!DESCRIPCION
'End If
'rs.Close
'rs2.Close
'rs.Open sql, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
'Do While Not rs.EOF
'   rs2.Open "SELECT formula FROM producto WHERE codigo = '" & rs!producto & "'", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
'   If Not rs2.EOF And rs2!formula <> True Then
'      CantBulto = CantBulto + 1
'   End If
'   rs2.Close
'rs.MoveNext
'Loop
'RptRemAj.CBulto.Text = CantBulto

RptRemAj.Show
RptRemAj.Printer.Copies = 1
End Function


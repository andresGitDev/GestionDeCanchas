Attribute VB_Name = "ModuloLiAMR"
Option Explicit ' mod 16/12/4

Public MODO_ON_ERROR_ABM_ON As Boolean
Public Const strSqlReservaCalculada = "SELECT p.codigo, p.existencia, p.formula, Sum(pi.Cantidad) AS SumaDeCantidad FROM Pedidos_Clientes AS pe RIGHT JOIN (Producto AS p LEFT JOIN ItemPedidoCliente AS pi ON p.codigo = pi.Producto) ON pe.numero = pi.PEDIDO Where (((p.activo) = 1) And ((pe.activo) = 1) And ((pe.cancelado) = 0)) GROUP BY p.codigo, p.existencia, p.formula"


'Public Type typeFormula
'    componente As String
'    cantidad As Long
'End Type
'Public ColFormula As Collection

Public Const REMITO_CON_PRECIO = True
Public Const CHAR_PROD_VIRTUAL = "V"
'Public Const PRODUCTO_CON_FORMULA_ES_VIRTUAL = True

' *** ' De Tabla TipoComprobantesGrales :  REMVTA = 5 - REMCPRA = 6
Public Const TipoComprobante_CANCELACIONPEDIDO = 8 ' Cancelpedido-remitoDifStock
Public Const TipoComprobante_REMITOVENTA = 5 ' REMVTA
Public Const TipoComprobante_REMITOCOMPRA = 6 ' REMCPRA
Public Const TipoComprobante_DIFSTOCK = 7 ' Dif Stock
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
'

' tabla Tipo documentos, BORRAR el form abm!!

Public Const TipoDoc_FACTURA_A = "FAA"
Public Const TipoDoc_FACTURA_B = "FAB"
Public Const TipoDoc_FACTURA_E = "FAE"
Public Const TipoDoc_NCREDITO_A = "NCA"
Public Const TipoDoc_NDEBITO_A = "NDA"
Public Const TipoDoc_NDEBITO_B = "NDB"
Public Const TipoDoc_NCREDITO_B = "NCB"
Public Const TipoDoc_NCREDITO_E = "NCE"
Public Const TipoDoc_NDEBITO_E = "NDE"

Public Const TipoDoc_RECIBO = "RAA"     'tabla facturaVenta, en Recibos no hace falta
Public Const TipoDoc_AJ_CREDITO = "ACC"
Public Const TipoDoc_AJ_DEBITO = "ACD"
'


'' ***********  Comprobantes ***********
'
Public Function TipoFormVenta(codigoIva) As String
    On Error GoTo fin
    TipoFormVenta = obtenerDeSQL("select letra from ivas where codigo = " & codigoIva)
fin:
    If TipoFormVenta = "" Then ufa "err: Formulario de tipo iva no definido ", "TipoFormVenta: ivas =" & codigoIva ', Err
End Function

' **************************************

' ************* Producto ************
Public Function rsFormulaComponentes(productoBase As String) As ADODB.Recordset
    'OJO cerrar rs donde lo llama
    If Not DE_EstaAbierto Then DataEnvironment1.AMR.Open
    Set rsFormulaComponentes = New ADODB.Recordset
    rsFormulaComponentes.Open "select Componente, cantidad from Formulas where activo = 1 and codigo = '" & productoBase & "'", DataEnvironment1.AMR, adOpenForwardOnly, adLockReadOnly
End Function
' Codigo de producto PROPIO
Public Function VerProductoCliente(codigo, Optional codPropio As Boolean)
    On Error Resume Next
    VerProductoCliente = ""
    
    If codigo = "" Then
        VerProductoCliente = ""
    ElseIf codPropio Then
        VerProductoCliente = codigo
    Else
        VerProductoCliente = obtenerDeSQL("select productoCliente from Relacion_Producto_Cliente where producto = '" & codigo & "'")
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
Public Function ProductoConSerie(cod As String, Optional bPropio As Boolean = True) As Boolean
    On Error Resume Next
    Dim conSerie 'variant
    conSerie = obtenerDato("Producto", "'" & VerProductoMio(cod, bPropio) & "'", "serie")
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
Public Function DepositoCod2Campo(cod)
    Dim t As Variant
    t = Array("existencia", "dep1", "dep2", "dep3", "dep4")
    DepositoCod2Campo = t(cod)
End Function
'
'************************************

Public Function AyudaProducto(codCliente As Long, codPropio As Boolean)
    If codPropio Then
        frmBuscar.MostrarSql "select codigo as [ Producto             ], descripcion  as [ Descripcion                                              ] from producto where activo = 1"
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
    rs.Open ssql, DataEnvironment1.AMR, adOpenStatic, adLockReadOnly
    obtenerParametro = rs.Fields(0)
    
    Set rs = Nothing
End Function
Public Function obtenerParametroDE(cual) As Variant 'As long
    Dim rs As New ADODB.Recordset, ssql  As String
    
    ssql = "select " & cual & " from " & TABLA_PARAMETROS
    rs.Open ssql, DataEnvironment1.AMR, adOpenStatic, adLockReadOnly
    obtenerParametroDE = rs.Fields(0)
    
    Set rs = Nothing
End Function

Public Function AumentarParametroN(cual, Nuevo) As Boolean
    Dim viejo
    
    viejo = obtenerParametro(cual)
    
    If viejo > Nuevo Then
        ufa "PrgErr: Intento grabar numero menor", "Maybe UserFault: AumentarParametro" ', Err
        Exit Function
    End If
    
    DataEnvironment1.AMR.Execute "update bs set " & cual & " = " & Nuevo
'     daTaenvironment1.AMR.Execute "update bs set " & cual & " = " & Nuevo
    AumentarParametroN = True
End Function
Public Function AumentarParametroD(cual As String, Nuevo As Date) As Boolean
    Dim viejo
    
    viejo = obtenerParametro(cual)
    
    If viejo > Nuevo Then
        ufa "PrgErr: Intento grabar fecha menor", "aum parametro" ', Err
        Exit Function
    End If
    
    DataEnvironment1.AMR.Execute "update bs set " & cual & " = " & ssFecha(Nuevo)
'    daTaenvironment1.AMR.Execute "update bs set " & cual & " = " & ssFecha(Nuevo)
    AumentarParametroD = True
End Function
Public Function CambiarParametroS(cual As String, Nuevo As String) As Boolean
    Dim viejo As String
    
    viejo = obtenerParametro(cual)
    
    If viejo = "" Then
        ufa "PrgErr: Intento grabar un vacio", "aum parametro" ', Err
        Exit Function
    End If
    
    DataEnvironment1.AMR.Execute "update " & TABLA_PARAMETROS & " set " & cual & " = '" & Nuevo & "'"
    CambiarParametroS = True
End Function
Public Function CambiarParametroN(cual As String, Nuevo As String) As Boolean
    Dim viejo
    
    viejo = obtenerParametro(cual)
    
    If IsNull(viejo) Then
        ufa "", "camb parametro n" & cual & Nuevo ', Err
        If Not confirma("dato previo vacio - Grabo?") Then Exit Function
    End If
    
    DataEnvironment1.AMR.Execute "update " & TABLA_PARAMETROS & " set " & cual & " = " & Nuevo
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

Public Function nuevoCodigo(TablaDE As String, Optional cpocodigo As String, Optional whe As String) As Long
    Dim rs As New ADODB.Recordset
    Dim ssql As String

    If cpocodigo = "" Then cpocodigo = "codigo"

    ssql = "Select max (" & cpocodigo & ")  as NN From " & TablaDE
    If whe > "" Then ssql = ssql & " where " & whe

    DE_abrir
    rs.Open ssql, DataEnvironment1.AMR, adOpenDynamic, adLockReadOnly
    
    If Not rs.EOF Then
        If IsNull(rs.Fields("NN")) Then
            nuevoCodigo = 1
        Else
            nuevoCodigo = rs.Fields("NN") + 1
        End If
    Else
        nuevoCodigo = 1
    End If

    Set rs = Nothing
End Function
Public Function nuevoCodigoDB(TablaDB As String, Optional cpocodigo As String, Optional whe As String) As Long
    Dim rs As New ADODB.Recordset
    Dim ssql As String

    If cpocodigo = "" Then cpocodigo = "codigo"

    ssql = "Select max (" & cpocodigo & ")  as NN From " & TablaDB
    If whe > "" Then ssql = ssql & " where " & whe

    DE_abrir
    rs.Open ssql, DataEnvironment1.AMR, adOpenDynamic, adLockReadOnly
    
    If Not rs.EOF Then
        If IsNull(rs.Fields("NN")) Then
            nuevoCodigoDB = 1
        Else
            nuevoCodigoDB = rs.Fields("NN") + 1
        End If
    Else
        nuevoCodigoDB = 1
    End If

    Set rs = Nothing
End Function



' -****************  TRANSACCIONES daTaenvironment1 *************************
' OJO un solo nivel
Public Function DE_BeginTrans() As Boolean 'adodb.Connection
''''   ' *** NO ATRAPAR ERROR !!!!!!!!
''''   'On Error GoTo ufaBT
''''   ' *** DEBE SALTAR error para que quien llama no siga grabando.
    If Not DE_EstaAbierto() Then DataEnvironment1.AMR.Open
    
'    Dim x as long
    'x =DataEnvironment1.AMR.BeginTrans()
    DataEnvironment1.AMR.BeginTrans
    'x = DataEnvironment1.AMR.Attributes
    DE_BeginTrans = True
    
''''fin:
''''    Exit Function
''''ufaBT:
''''    ufa "fallo intento de Comenzar transaccion", "DE_BeginTrans"
''''    End
End Function

Public Function DE_CommitTrans() As Boolean
   On Error GoTo ufaCT
   DataEnvironment1.AMR.CommitTrans
   DE_CommitTrans = True
fin:
    Exit Function
ufaCT:
    ufa "fallo intento de completar transaccion", "DE_CommitTrans"
    Resume fin
End Function
Public Function DE_RollbackTrans() As Boolean ' OJO muere silenciosamente
   On Error GoTo ufaCT
   DataEnvironment1.AMR.RollbackTrans
   DE_RollbackTrans = True
fin:
    Exit Function
ufaCT:
    ufa "", "DE_RollbackTrans Falla rollBack DE "
    Resume fin
End Function
Public Function DE_EstaAbierto() As Boolean
    On Error GoTo UfaEA
    DE_EstaAbierto = ((DataEnvironment1.AMR.State And adStateOpen) > 0)
fin:
    Exit Function
UfaEA:
    ufa "fallo revisando conexion", "DE_EstaAbierto"
    Resume fin
End Function
Public Function DE_abrir() As Boolean
    On Error GoTo UfaDA
    'DataEnvironment1.AMR.Provider
    
'    daTaenvironment1.AMR.Open
    If Not DE_EstaAbierto() Then DataEnvironment1.AMR.Open
    DE_abrir = True
fin:
    Exit Function
UfaDA:
    ufa "fallo abriendo conexion", "DE_Abrir"
    DE_abrir = False
    Resume fin
End Function
' -****************  TRANSACCIONES  daTaenvironment1 *************************
''''' -****************  TRANSACCIONES  DB  *************************
''''' OJO un solo nivel
''''Public Function DB_BeginTrans() As Boolean 'adodb.Connection
''''   On Error GoTo ufaBT
''''    daTaenvironment1.AMR.BeginTrans
''''    DE_BeginTrans = True
''''Fin:
''''    Exit Function
''''ufaBT:
''''    ufa "fallo intento de Comenzar transaccion", "Db_BeginTrans"
''''    Resume Fin
''''End Function
''''Public Function DB_CommitTrans() As Boolean
''''   On Error GoTo ufaCT
''''   daTaenvironment1.AMR.CommitTrans
''''   DE_CommitTrans = True
''''Fin:
''''    Exit Function
''''ufaCT:
''''    ufa "fallo intento de completar transaccion", "Db_CommitTrans"
''''    Resume Fin
''''End Function
''''Public Function DB_RollbackTrans() As Boolean ' OJO muere silenciosamente
''''   On Error GoTo ufaCT
''''   daTaenvironment1.AMR.RollbackTrans
''''   DE_RollbackTrans = True
''''Fin:
''''    Exit Function
''''ufaCT:
''''    ufa "", "DE_RollbackTrans Falla rollBack daTaenvironment1.amr "
''''    Resume Fin
''''End Function
''''' -****************  TRANSACCIONES  DB  *************************

Public Function leerEjercicio() As Long
    leerEjercicio = obtenerParametro("Ejercicio")
End Function

Public Function HayProdEnEdicion(strDescrProd As String) As Boolean
    If Trim$(strDescrProd) = "" Then
        HayProdEnEdicion = False
    Else
        HayProdEnEdicion = Not confirma("Hay un Producto en la linea de edicion" & vbCrLf & "Lo descarta ?")
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
'''Public Function SerieEnStock(cualSerie As String, cualProducto As String) As Boolean
'''    '    ss = "SELECT  serie as [ Serie                 ], producto as  [ Producto              ] , MAX(codigo) as  [Movimiento ] From SERIES Where (activo = 1 and producto = '" & prod & "') GROUP BY  producto, serie"
'''    Dim tempo
'''    tempo = obtenerDeSQL("select serie from series where serie = '" & cualSerie & "' and producto = '" & cualProducto & "' ")
'''    SerieEnStock = (sSinNull(tempo) > "")
'''End Function
'''Public Function SerieAfuera(cualSerie As String, cualProducto As String) As Boolean
'''    Dim tempo
'''    tempo = obtenerDeSQL("select serie from series where serie = '" & cualSerie & "' and producto = '" & cualProducto & "' ")
'''    SerieEnStock = (sSinNull(tempo) > "")
'''End Function

Public Function ProductoDescripcion(codi) As String
    codi = sSinNull(codi)
    If codi = "" Then Exit Function
    ProductoDescripcion = obtenerDeSQL("select descripcion from producto where codigo = '" & Trim(codi) & "' and activo = 1 ")
End Function
'''Public Function Buscar_SeriesEnStock(producto As String) As String
'''    On Error GoTo UfaBuscaSer
'''    Dim ss As String, tmpTablaSeries As String
'''
'''    tmpTablaSeries = TablaTempCrear(tt_SeriesEnStockTemp)
'''    ss = "INSERT INTO " & tmpTablaSeries & " ( Codigo, Producto, Serie, Descripcion ) " _
'''        & " SELECT max(series.Codigo) as UltimoCodigo, producto, Series.serie, Descripcion From Series " _
'''        & " inner join producto on producto = producto.codigo " _
'''        & " Where Series.activo = 1 and producto = '" & producto & "' " _
'''        & " GROUP BY producto, Descripcion, Series.serie order by Series.serie "
'''
'''' debug
'''    If producto = "" Then
'''        ss = "INSERT INTO " & tmpTablaSeries & " ( Codigo, Producto, Serie ) " _
'''        & " SELECT max(Codigo) as UltimoCodigo, producto, serie From Series " _
'''        & " Where activo = 1  " _
'''        & " GROUP BY producto, serie order by serie "
'''    End If
'''' '''''
'''
'''    DataEnvironment1.AMR.Execute ss
'''
'''    ss = "SELECT t.Serie as [ Serie               ], s.Producto as [ Producto                  ], t.Descripcion as [ Descripcion                                              ], s.comprobante as [c], t.codigo  as [i]" _
'''        & " FROM " & tmpTablaSeries & " AS t INNER JOIN Series AS s ON t.codigo = s.codigo left join conceptos as c on c.codigo = s.concepto " _
'''        & " WHERE s.comprobante = 6 or s.comprobante = 3 or s.comprobante = 4 or (s.comprobante = 7 and c.movimiento <> 'R' )  "
'''    Buscar_SeriesEnStock = frmBuscar.MostrarSql(ss)
'''
'''fin:
'''    Exit Function
'''UfaBuscaSer:
'''    ufa "err: buscando series", "Prod: " & producto
'''    Resume fin
'''End Function

Public Function ExistenciaCalculada(producto As String) As Long ', Optional deposito As Long = 0) As Long
    ' Nuevo concepto de existencia de productos virtuales
    ' Reemplaza a GeneraExistenciaCalculada(), ve cuantos split se pueden armar con lo que hay
    ' ** No considera deposito **
    
    Dim t As Variant, rs As New ADODB.Recordset, depo As String, s As String
    ' depo = IIf(deposito = 0, " existencia ", " dep" & Trim(CStr(deposito)))
    
    t = obtenerDeSQL("select existencia, formula from producto where codigo = '" & producto & "' ")
    If IsEmpty(t) Then
        ExistenciaCalculada = 0
        Exit Function
    ElseIf Not t(1) Then
        ExistenciaCalculada = s2n(t(0))
    Else
        s = "SELECT Min(existencia/cantidad) AS MaxArmados " _
            & " FROM  producto as p INNER JOIN Formulas as f ON p.codigo = f.Componente " _
            & " Where f.codigo = '" & producto & "' "
            
        ExistenciaCalculada = Fix(s2n(obtenerDeSQL(s)))
    End If
End Function
Public Function ReservaCalculada(producto As String) As Long
    Dim s As String, rs As New ADODB.Recordset
    s = "SELECT sum(i.Saldo) as sal FROM Pedidos_Clientes AS p INNER JOIN ItemPedidoCliente AS i " _
        & " ON p.numero = i.PEDIDO " _
        & " WHERE p.activo=1 AND p.cancelado=0 and i.producto = '" & producto & "' " 'and fechavencimiento > xxx
    'rs.Open DataEnvironment1.VistaReservas
    ReservaCalculada = s2n(obtenerDeSQL(s))
End Function


 
'Public Function GeneraExistenciaCalculada()
''    On Error GoTo UfaCalcExistencia
'
'    Dim ss0 As String, ss As String, ss1 As String, ss2 As String
'    Dim rs As New ADODB.Recordset, rs2 As New ADODB.Recordset
'    Dim CodiZZZ As String, CantMin As Double, reserv As Double, tempo
'
'    ss0 = "update producto set ExistenciaCalculada = Existencia, ReservaCalculada = 0"
'    DataEnvironment1.AMR.Execute ss0
'    ss0 = "update producto set ExistenciaCalculada = 0 where formula = 1"
'    DataEnvironment1.AMR.Execute ss0
'
'    'y el SUM() ???
'    'RESERVADOS --- ACA SE TIENE Q HACER LA VERIFICACION DE VENCIMIENTO (fecha venc pedido) ----
'    ss0 = "SELECT i.Producto, i.Saldo FROM Pedidos_Clientes AS p INNER JOIN ItemPedidoCliente AS i " _
'        & " ON p.numero = i.PEDIDO " _
'        & " WHERE (((i.Saldo)>0) AND ((p.activo)=1) AND ((p.cancelado)=0)) " 'and fechavencimiento > xxx
'    rs.Open ss0, DataEnvironment1.AMR, adOpenForwardOnly, adLockReadOnly
'    While Not rs.EOF
'        DataEnvironment1.AMR.Execute "update producto set ReservaCalculada =  ReservaCalculada + " & x2s(rs!Saldo) & " where codigo = '" & rs!producto & "' "
'        rs.MoveNext
'    Wend
'    rs.Close
'
'    If gEMPR_FormulaEsVirtual Then
'        ss = "select distinct codigo from formulas where activo = 1"
'        rs.Open ss, DataEnvironment1.AMR, adOpenForwardOnly, adLockReadOnly
'        While Not rs.EOF
'            CodiZZZ = rs!codigo
'
''            If Trim(CodiZZZ) = "ZZZZZZ553TFH0904" Then Stop
'
'            ' averiguo max cant q se pueden armar (min de componente)
'            ss1 = "SELECT Min(existenciaCalculada/cantidad) AS MaxArmados, min(ReservaCalculada) as reservado " _
'                & " FROM  producto as p INNER JOIN Formulas as f ON p.codigo = f.Componente " _
'                & " Where f.codigo = '" & CodiZZZ & "' "
'
'            tempo = obtenerDeSQL(ss1)
'            CantMin = Fix(s2n(tempo(0)))
'            reserv = s2n(tempo(1))
'
'            'update virtual cant q se pueden armar
'            DataEnvironment1.AMR.Execute "update producto set ExistenciaCalculada = " & x2s(CantMin) & " where codigo = '" & CodiZZZ & "' and activo = 1 "
'
'            'update componentes, resta de virtuales
'            If CantMin > 0 Then
'
'                ss2 = "select componente from formulas where activo = 1 and codigo = '" & CodiZZZ & "'"
'                rs2.Open ss2, DataEnvironment1.AMR, adOpenForwardOnly, adLockReadOnly
'                While Not rs2.EOF
''                    DataEnvironment1.AMR.Execute "update Producto set " _
'                        & " ExistenciaCalculada = ExistenciaCalculada - " & x2s(CantMin) _
'                        & ", ReservaCalculada = ReservaCalculada - " & x2s(reserv) _
'                        & " where codigo = '" & rs2!componente & "' "
'                    DataEnvironment1.AMR.Execute "update Producto set " _
'                        & " ExistenciaCalculada = ExistenciaCalculada - " & x2s(CantMin) _
'                        & ", ReservaCalculada = ReservaCalculada - " & x2s(reserv) _
'                        & " where codigo = '" & rs2!componente & "' "
'                    rs2.MoveNext
'                Wend
'                rs2.Close
'            End If
'            rs.MoveNext
'        Wend
'    End If
'
'fin:
'    Set rs = Nothing
'    Set rs2 = Nothing
'    Exit Function
'
'UfaCalcExistencia:
'    ufa "err: Buscando stock", "ExistenciaCalculada"
'    Resume fin
'End Function

Public Function RevisaNroYFechaOk(sTabla As String, sNum As String, sFec As String, numero As Long, fecha As Date, masWhere As String) As Boolean
    Dim ss As String, tempo As Variant
    

    RevisaNroYFechaOk = False
    
    If masWhere > "" Then masWhere = " AND " & masWhere
    
    ' reviso numero
    ss = "select " & sFec & " from " & sTabla & " where " & sNum & " = " & numero & masWhere
    tempo = obtenerDeSQL(ss)
    If Not IsEmpty(tempo) Then
        che "Numero " & numero & " existe con fecha  " & tempo
        Exit Function
    End If
    
    'Reviso Fecha anterior
    ss = "select max(" & sNum & ") from " & sTabla & " where " & sNum & " < " & numero & masWhere
    tempo = s2n(obtenerDeSQL(ss), 0)
    If tempo > 0 Then
        ss = "select " & sFec & " from " & sTabla & " where " & sNum & " = " & tempo & masWhere
        tempo = obtenerDeSQL(ss)
        If tempo > fecha Then
            che "Documento anterior tiene fecha " & tempo
            Exit Function
        End If
    End If
    
    'Reviso Fecha Posterior
    ss = "select min(" & sNum & ") from " & sTabla & " where " & sNum & " > " & numero & masWhere
    tempo = s2n(obtenerDeSQL(ss), 0)
    If tempo > 0 Then
        ss = "select " & sFec & " from " & sTabla & " where " & sNum & " = " & tempo & masWhere
        tempo = obtenerDeSQL(ss)
        If tempo < fecha Then
            che "Documento posterior tiene fecha " & tempo
            Exit Function
        End If
    End If
  
    RevisaNroYFechaOk = True
End Function

Public Function nuevoCodigoOP() As Long
    Dim tmp As Long
    
    tmp = nuevoCodigo("Rec_Comp", "Nro")
    nuevoCodigoOP = tmp
    
    tmp = nuevoCodigo("transcom", "NroDoc", "TipoDoc = 'RAC'")
    If tmp > nuevoCodigoOP Then nuevoCodigoOP = tmp
    
    tmp = nuevoCodigo("Compras", "NroDoc", "TipoDoc = 'RAC'")
    If tmp > nuevoCodigoOP Then nuevoCodigoOP = tmp
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

Public Sub capFrm(que As Form)
    que.caption = que.caption & "          " & gEMPR_NombreEmpresa
End Sub


'Public Function ClienteHabilitado(codigo As Long) As Boolean
''    Dim tempo
''    tempo = s2n(obtenerDeSQL("select PuedoFacturar from clientes where  codigo = '" & codigo & "' and activo = 1"))
''    ClienteHabilitado = (tempo = 1)
'    ClienteHabilitado = True
'End Function
Public Function ClienteCredito(codigo As Long) As Double
    Dim tempo, temp2
    
    tempo = s2n(obtenerDeSQL("select LimiteCredito from clientes where  codigo = '" & codigo & "' "))
'    If tempo = 0 Then tempo = LIMITE_CREDITO_PREDETERMINADO
    
    temp2 = 0 's2n(obtenerdesql(" ") ) 'DEUDA CLIENTE
    'OJO: Factura, NC, ND, PagosAcuenta, etc
    
    ClienteCredito = tempo - temp2
End Function

''''Public Function TablaTemp_Existencias() As String
'''''    On Error GoTo UfaBuscarExistencia
''''    Dim ss As String, ss1 As String, ss2 As String, ss0 As String
''''    Dim tmpTablaStock As String, rs As New ADODB.Recordset, rs2 As New ADODB.Recordset
''''    Dim CantMin As Double, codi As String
''''
''''    tmpTablaStock = TablaTempCopiar("Producto") 'La copia llena
'''''    ss = "INSERT INTO " & tmpTablaStock & " (codigo, descripcion, existencia, formula ) SELECT p.codigo, p.descripcion, p.existencia, p.formula " _
''''        & " From Producto as p " _
''''        & " Where activo = 1 " 'and p.existencia > 0 And p.formula = 0 "
'''''    daTaenvironment1.amr.Execute ss
''''
''''
''''    'RESTO RESERVADOS --- ACA SE TIENE Q HACER LA VERIFICACION DE VENCIMIENTO (fecha venc pedido) ----
''''    ss0 = "SELECT i.Producto, i.Saldo FROM Pedidos_Clientes AS p INNER JOIN ItemPedidoCliente AS i " _
''''        & " ON p.numero = i.PEDIDO " _
''''        & " WHERE (((i.Saldo)>0) AND ((p.activo)=1) AND ((p.cancelado)=0)) " 'and fechavencimiento > xxx
''''    rs.Open ss0, daTaenvironment1.amr, adOpenForwardOnly, adLockReadOnly
''''    While Not rs.EOF
''''        daTaenvironment1.amr.Execute "update " & tmpTablaStock & " set existencia = existencia - " & x2s(rs!Saldo) & " where codigo = '" & rs!producto & "' "
''''        rs.MoveNext
''''    Wend
''''    rs.Close
''''
''''    ss = "select distinct codigo from formulas where activo = 1"
''''    rs.Open ss, daTaenvironment1.amr, adOpenForwardOnly, adLockReadOnly
''''    While Not rs.EOF
''''        codi = rs!Codigo
''''
''''        ' averiguo max cant q se pueden armar (min de componente)
''''        ss1 = "SELECT Min(existencia/cantidad) AS MaxArmados " _
''''            & " FROM  " & tmpTablaStock & " as p INNER JOIN Formulas as f ON p.codigo = f.Componente " _
''''            & " Where f.codigo = '" & codi & "' "
''''        CantMin = Fix(s2n(obtenerDeSQL(ss1)))
''''
''''        'update virtual cant q se pueden armar
''''        daTaenvironment1.amr.Execute "update " & tmpTablaStock & " set existencia = " & x2s(CantMin) & " where codigo = '" & codi & "' and activo = 1 "
''''
''''        'update componentes, resta de virutuales
''''        If CantMin > 0 Then
''''            ss2 = "select componente from formulas where activo = 1 and codigo = '" & codi & "'"
''''            rs2.Open ss2, daTaenvironment1.amr, adOpenForwardOnly, adLockReadOnly
''''            While Not rs2.EOF
''''                daTaenvironment1.amr.Execute "update " & tmpTablaStock & " set existencia = existencia - " & x2s(CantMin) & " where codigo = '" & rs2!componente & "' and activo = 1 "
''''                rs2.MoveNext
''''            Wend
''''            rs2.Close
''''        End If
''''
''''        rs.MoveNext
''''    Wend
''''    TablaTemp_Existencias = tmpTablaStock
''''
''''Fin:
''''    Set rs = Nothing
''''    Set rs2 = Nothing
''''    Exit Function
''''
''''UfaBuscarExistencia:
''''    TablaTemp_Existencias = ""
''''    ufa "err: Buscando stock", ""
''''    Resume Fin
''''End Function
''Public Function ProductoExistencia(producto) As Double
''    On Error GoTo UfaExi
''    Dim esformu As Boolean, ss1 As String
''    esformu = (obtenerDeSQL("select formula from producto where activo = 1 and codigo = '" & producto & "' ") = 1)
''
''    If esformu Then
''        ' averiguo max cant q se pueden armar (min de componente)
''        ss1 = "SELECT  Min(existencia/cantidad) AS MaxArmados  " _
''            & " FROM  " & tmpTablaStock & " as p INNER JOIN Formulas as f ON p.codigo = f.Componente " _
''            & " Where f.codigo = '" & codi & "' "
''        ProductoExistencia = Fix(s2n(obtenerDeSQL(ss1)))
''    Else
''
''    End If
''
''fin:
''    Exit Function
''UfaExi:
''    ufa "prg: err al averiguar ewxistencia " & producto, "productoExistencia " & producto
''    Resume fin
''End Function



'Public Function DateServer() As Date
'    DateServer = obtenerDeSQL("select date() from bs")
'End Function

' pasar a modedicioncomprobantes
'Public Function LetraDoc(TipoIva As long) As String
'    Dim tmp
'    LetraDoc = ""
'    tmp = obtenerDeSQL("select letra from ivas where codigo = " & TipoIva)
'    If Not IsEmpty(tmp) Then LetraDoc = tmp
'End Function

Public Function Imputacion_YYYYMM(C_o_V As String, fecha As Date) As String()

    Dim tempo, stemp As String, sfech As String
    Dim atxt(1) As String
    
    tempo = fechaCierre(C_o_V) ' VerDatoEmpresa(sCampo)
    stemp = Left(Format(tempo, "yyyymmdd"), 6)
    sfech = Left(Format(fecha, "yyyymmdd"), 6)
    
    If stemp >= sfech Then
        If Month(tempo) = 12 Then
            atxt(1) = "1"
            atxt(0) = Year(tempo) + 1
        Else
            atxt(0) = Year(tempo)
            atxt(1) = Month(tempo) + 1
        End If
    Else
        atxt(1) = Month(fecha)
        atxt(0) = Year(fecha)
    End If
    Imputacion_YYYYMM = atxt
End Function
Public Function Imputacion_Verifica_Ok(C_o_V As String, sAnio As String, sMes As String) As Boolean
    Dim ffecha As Long, dFecha As Long, tempo
    
    tempo = fechaCierre(C_o_V)
    dFecha = Year(tempo) * 100 + Month(tempo)
    ffecha = val(sAnio) * 100 + val(sMes)
    If ffecha <= dFecha Then
        che "Fecha imputacion corresponde a periodo cerrado"
        Exit Function
    End If
    Imputacion_Verifica_Ok = True
End Function
Private Function fechaCierre(C_o_V As String) As Date
    Dim sCampo As String
    If LCase(C_o_V) = "v" Then
        sCampo = "fechaimputV"
    ElseIf LCase(C_o_V) = "c" Then
        sCampo = "fechaimputC"
    Else
        ufa "prg: Campo C_o_V mal cargado", ""
    End If
    fechaCierre = VerDatoEmpresa(sCampo)
End Function

Public Function ProvActivo(codprov) As Boolean
    On Error Resume Next
    ProvActivo = obtenerDeSQL("select activo_pr from prov where codigo = '" & codprov & "' ")
End Function
Public Function ClienteHabilitado(codClie) As Boolean
    On Error Resume Next
    ClienteHabilitado = obtenerDeSQL("select PuedoFacturar from clientes where codigo = '" & codClie & "' ")
End Function

'28/10/4    ProductoConSerie
'16/12/4    VerProductoCliente() VerProductoMio(): manegçjo err, permito codigo ""
'17/1/5      + bs.num_opago,
'           CambiarParametroN
'26/5/5
'   codigo FV desde tabla FacturaVenta no BS

'7/2/5      transacciones, para db y daTaenvironment1 / deberia ser parametro, no dupli, pero deberia haber 1 solo DE
'4/4/5      NuevoCodigo() , ahora labura con DE (daTaenvironment1.amr) para meter dentro de transacciones
'30/5/5     RevisaNroFechaOk()
'23/11/5    ProvActivo(), ClienteHabilitado() as boolean
'


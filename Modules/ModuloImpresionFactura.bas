Attribute VB_Name = "ModuloImpresionFactura"
Option Explicit

'Private Const ppppTacharY = 0
'Private Const ppppTacharX = 40

Private mpppp

Private Enum ImprFactura
   Margen_Y
   Margen_X
   tachar_Y
   tachar_X
   Nro_Y
   Nro_x
   labelFactura_Y
   labelFactura_X
   fecha_y
   fecha_X
   Encabezado_Y
   Encabezado_X
   CondIVA_Y
   CondIVA_X
   FP_Y
   FP_X
   VaCon_Y
   VaCon_X
   Detalle_Y
   Detalle_X_Cant
   Detalle_X_Prod
   Detalle_X_Desc
   Detalle_X_PreU
   Pie_Y
   Totales_x
   porc_x
   EnLetras_Y
   EnLetras_X
End Enum


Private sTablaRemito As String
Private sTablaTemp As String  'Private Const CodTomKa = 366591

Public Const tt_Etiquetas_temp = _
    "([ProdCliente] [varchar] (50)," & _
    "[Letra] [varchar] (3)," & _
    "[Cantidad] [float]," & _
    "[descripcion] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," & _
    "[Destino] [varchar] (50), " & _
    "[Remito] [float] , " & _
    "[CodBarra] [varchar] (50)) "


Public Function ImprimirRemitoVenta(codigo) As Boolean
    If ON_ERROR_HABILITADO Then On Error GoTo ErrImpresora
    
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim CantBulto, AuxNPed As Double
    Dim str, str2, NPedidos As String
    Dim COD As Long
    Dim Propio As Boolean
    Dim cadena As String
    Dim from As String
    Dim vari As Integer
    Dim trans
    Dim vende

    
    
'    str1 = "SELECT FacturaVenta.*, Clientes.*, FormasPago.descripcion as pago, " & _
'            "Ivas.descripcion as iva, Ivas.letra as letra" & _
'            " FROM FormasPago INNER JOIN (Ivas INNER JOIN (Clientes INNER JOIN FacturaVenta  " & _
'            " ON Clientes.codigo = FacturaVenta.Cliente) " & _
'            " ON Ivas.codigo = FacturaVenta.TipoIVA) " & _
'            " ON FormasPago.codigo = FacturaVenta.FormaPago " & _
'            " WHERE facturaventa.codigo=" & codigo
            
            
    
    cadena = " SELECT RemitoVenta.*,remitoventa.transporte as trans,clientes.codigo as cod, Clientes.* "
    from = " FROM  (Clientes INNER JOIN RemitoVenta ON Clientes.codigo = RemitoVenta.Cliente) "
            
    rs.Open " SELECT * FROM  RemitoVenta where numero=" & codigo, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    
    
    vari = rs!Transporte
    
    If vari > 0 Then 'para transportes
        Set rs = Nothing
        rs.Open " SELECT * FROM  transportes where codigo=" & vari, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
        If rs.EOF = True And rs.BOF = True Then
            trans = "-"
        Else
            cadena = cadena & ", transportes.descripcion as transportista, transportes.direccion as direTrans, transportes.telefono as phone "
            from = from & " inner join transportes on remitoventa.transporte=transportes.codigo "
            trans = rs!DESCRIPCION
        End If
    Else
        trans = "-"
    End If
    Set rs = Nothing
    
    rs.Open " SELECT RemitoVenta.*, Clientes.* " & _
            " FROM  (Clientes INNER JOIN RemitoVenta ON Clientes.codigo = RemitoVenta.Cliente) where numero=" & codigo, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    vari = nSinNull(rs!Vendedor)
    If vari > 0 Then 'para vendedor
        Set rs = Nothing
        rs.Open " SELECT * FROM  usuarios where codigo=" & vari, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
        If rs.EOF = True And rs.BOF Then
            vende = "-"
        Else
            cadena = cadena & ", usuarios.descripcion as vendedor "
            from = from & " inner join usuarios on clientes.vendedor=usuarios.codigo "
            vende = rs!DESCRIPCION
        End If
    Else
        vende = "-"
    End If
    Set rs = Nothing
    
    cadena = cadena & from & "where numero=" & codigo
    
    'rs.Open " SELECT RemitoVenta.*,remitoventa.transporte as trans, Clientes.*, transportes.descripcion as transportista, transportes.direccion as direTrans, transportes.telefono as phone,vendedores.nombre as vendedor, vendedores.apellido as apeVendedor " & _
    '        " FROM  (Clientes INNER JOIN RemitoVenta ON Clientes.codigo = RemitoVenta.Cliente)" & _
    '        " inner join transportes on remitoventa.transporte=transportes.codigo" & _
    '        " inner join vendedores on clientes.vendedor=vendedores.codigo where numero=" & codigo _
    '        , DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    
            
    'rs.Open " SELECT RemitoVenta.*,remitoventa.transporte as trans, Clientes.*, transportes.descripcion as transportista, transportes.direccion as direTrans, transportes.telefono as phone,usuarios.descripcion as vendedor " & _
    '        " FROM  (Clientes INNER JOIN RemitoVenta ON Clientes.codigo = RemitoVenta.Cliente)" & _
    '        " inner join transportes on remitoventa.transporte=transportes.codigo" & _
    '        " inner join usuarios on clientes.vendedor=usuarios.codigo where numero=" & codigo _
    '        , DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    
    rs.Open cadena, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly


    If Not rs.EOF Then
        Propio = rs!codPropio
        
        RptImpresionRemitoVenta.lblcliente = sSinNull(rs!nombrefantasia)
        If Not IsNull(rs!CUIT) Then
            RptImpresionRemitoVenta.lblCuit = rs!CUIT
        End If
        
        COD = rs!numero
        RptImpresionRemitoVenta.lblcomp = "Remito"
        If Propio Then
            str = "SELECT R.*, P.descripcion,u.descripcion as udes" _
                & " FROM RemitoVentaDetalle r INNER JOIN Producto p ON R.Producto = P.codigo" _
                & " inner join unidadesmedida u on u.umcodigo=p.umedida where CANCELADO=0 AND numero=" & COD _
                & " order by R.codigo "
        Else
            str = "SELECT r.*, p.descripcion, rpc.ProductoCliente as producto,u.descripcion as udes " _
                & " FROM RemitoVentaDetalle as r left JOIN Producto as p ON r.Producto = p.codigo " _
                & " left join relacion_producto_cliente as rpc on rpc.producto = p.codigo " _
                & " inner join unidadesmedida u on u.UMcodigo=producto.umedida where numero = " & COD & " and cliente = " & rs!cliente _
                & " order by r.codigo "
        End If
        
        str2 = "select codigo,numero,producto,pedido from remitoventadetalle where numero=" & rs!numero & " and" _
                & " codigo=(select max(codigo) from remitoventadetalle where numero=" & rs!numero & ") " _
                & " group by codigo,numero,producto,pedido"  'terminar,traigo de remideta... pedido para ordcomp
        rs2.Open str2, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly

        
        RptImpresionRemitoVenta.lblcliente = rs!DESCRIPCION
        RptImpresionRemitoVenta.NroCli = rs!COD
        
        Dim valu As New ADODB.Recordset
        valu.Open "select * from provincias where codigo='" & rs!Provincia & "'", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
        
        RptImpresionRemitoVenta.lblNroProv = valu!DESCRIPCION
        Set valu = Nothing
        
        RptImpresionRemitoVenta.CP = rs!codigopostal
        RptImpresionRemitoVenta.lbldomicilio = rs!direccion
        RptImpresionRemitoVenta.lblfactura = "0001-" & Format(rs!numero, "00000000")
        RptImpresionRemitoVenta.lblfecha = rs!Fecha
        RptImpresionRemitoVenta.lbllocalidad = rs!Localidad
        RptImpresionRemitoVenta.Transportista = trans 'rs!Transportista
        RptImpresionRemitoVenta.direTRANS = rs!direTRANS
        RptImpresionRemitoVenta.telTRANS = rs!phone
        RptImpresionRemitoVenta.lblAtencion = vende 'rs!Vendedor '& rs!apevendedor
        
        If rs!obs1 <> "" Then
            RptImpresionRemitoVenta.Label23.caption = "Observacion 1: " & rs!obs1 'frmRemitoVenta.txtObs(0)
        Else
            RptImpresionRemitoVenta.Label23.caption = ""
        End If
        If rs!obs2 <> "" Then
            RptImpresionRemitoVenta.Label24.caption = "Observacion 2: " & rs!obs2 'frmRemitoVenta.txtObs(1)
        Else
            RptImpresionRemitoVenta.Label24.caption = ""
        End If
        valu.Open "select * from ivas where codigo=" & rs!Iva, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
        RptImpresionRemitoVenta.LblIVA = valu!DESCRIPCION
        Set valu = Nothing
        
        RptImpresionRemitoVenta.OrdenComp = rs2!Pedido
        
        RptImpresionRemitoVenta.Dia = Day(rs!Fecha)
        RptImpresionRemitoVenta.Mes = Month(rs!Fecha)
        RptImpresionRemitoVenta.Ano = Year(rs!Fecha)
        RptImpresionRemitoVenta.Meses = PasoMes(Month(rs!Fecha))
         
        LlenarTemp (str)
        RptImpresionRemitoVenta.DataControl1.Connection = DataEnvironment1.Sistema
        
        ' CAMBIAR PARA QUE EL STR QUE FIGURA SE REDIRECCIONE A LA NUEVA TABLA
        ' QUE CREE
        str = "SELECT * FROM " & sTablaRemito & " "
        RptImpresionRemitoVenta.DataControl1.Source = str
        
        RptImpresionRemitoVenta.Printer.Copies = 1
        If VerDatoEmpresa("idEmpresa") > 1 Then
            With RptImpresionRemitoVenta
                .CP.Visible = False
            End With
        End If
        
        If PREVIEW_IMPRESIONES Then
            RptImpresionRemitoVenta.Show
            
            'esto es para acomodar las posiciones del reporte
            Posicionar (False) 'false para remito true para factura
            
        Else
            RptImpresionRemitoVenta.PrintReport True
        End If
        RptImpresionRemitoVenta.Restart
        
    End If
    rs.Close
    Set rs = Nothing
    
    
fin:
    Exit Function
ErrImpresora:
    ufa "error de impresión Remito Venta", ""
    Resume fin
End Function

Public Function ImprimirRemitoPorte(codigo) As Boolean
    If ON_ERROR_HABILITADO Then On Error GoTo ErrImpresora
    
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim CantBulto, AuxNPed As Double
    Dim str, str2, NPedidos As String
    Dim COD As Long
    Dim Propio As Boolean
    Dim cadena As String
    Dim from As String
    Dim vari As Integer
    Dim trans
    Dim vende
    Dim formas As Long

    
    
'    str1 = "SELECT FacturaVenta.*, Clientes.*, FormasPago.descripcion as pago, " & _
'            "Ivas.descripcion as iva, Ivas.letra as letra" & _
'            " FROM FormasPago INNER JOIN (Ivas INNER JOIN (Clientes INNER JOIN FacturaVenta  " & _
'            " ON Clientes.codigo = FacturaVenta.Cliente) " & _
'            " ON Ivas.codigo = FacturaVenta.TipoIVA) " & _
'            " ON FormasPago.codigo = FacturaVenta.FormaPago " & _
'            " WHERE facturaventa.codigo=" & codigo
            
            
    remiCarta = codigo
    cadena = " SELECT R.*,r.transporte as trans,C.codigo as cod, C.* "
    from = " FROM  (Clientes as C INNER JOIN RemitoPorte as R ON C.codigo = R.Cliente) "
            
    rs.Open " SELECT * FROM  RemitoPorte where numero=" & codigo, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    
    
    vari = rs!Transporte
    
    If vari > 0 Then 'para transportes
        Set rs = Nothing
        rs.Open " SELECT * FROM  transportes where codigo=" & vari, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
        If rs.EOF = True And rs.BOF = True Then
            trans = "-"
        Else
            cadena = cadena & ", transportes.descripcion as transportista, transportes.direccion as direTrans, transportes.telefono as phone "
            from = from & " inner join transportes on r.transporte=transportes.codigo "
            trans = rs!DESCRIPCION
        End If
    Else
        trans = "-"
    End If
    Set rs = Nothing
    'STR = " SELECT R.*, Clientes.* " & _
            " FROM  (Clientes AS C INNER JOIN RemitoPorte  AS R ON C.codigo = R.Cliente) where numero=" & codigo
    rs.Open " SELECT R.*, C.* " & _
            " FROM  (Clientes AS C INNER JOIN RemitoPorte  AS R ON C.codigo = R.Cliente) where numero=" & codigo, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    vari = rs!Vendedor
    formas = nSinNull(rs!formaPago)
    If vari > 0 Then 'para vendedor
        Set rs = Nothing
        rs.Open " SELECT * FROM  usuarios where codigo=" & vari, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
        If rs.EOF = True And rs.BOF Then
            vende = "-"
        Else
            cadena = cadena & ", usuarios.descripcion as vendedor "
            from = from & " inner join usuarios on c.vendedor=usuarios.codigo "
            vende = rs!DESCRIPCION
        End If
    Else
        vende = "-"
    End If
    Set rs = Nothing
    
    cadena = cadena & from & "where numero=" & codigo
    
    'rs.Open " SELECT RemitoVenta.*,remitoventa.transporte as trans, Clientes.*, transportes.descripcion as transportista, transportes.direccion as direTrans, transportes.telefono as phone,vendedores.nombre as vendedor, vendedores.apellido as apeVendedor " & _
    '        " FROM  (Clientes INNER JOIN RemitoVenta ON Clientes.codigo = RemitoVenta.Cliente)" & _
    '        " inner join transportes on remitoventa.transporte=transportes.codigo" & _
    '        " inner join vendedores on clientes.vendedor=vendedores.codigo where numero=" & codigo _
    '        , DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    
            
    'rs.Open " SELECT RemitoVenta.*,remitoventa.transporte as trans, Clientes.*, transportes.descripcion as transportista, transportes.direccion as direTrans, transportes.telefono as phone,usuarios.descripcion as vendedor " & _
    '        " FROM  (Clientes INNER JOIN RemitoVenta ON Clientes.codigo = RemitoVenta.Cliente)" & _
    '        " inner join transportes on remitoventa.transporte=transportes.codigo" & _
    '        " inner join usuarios on clientes.vendedor=usuarios.codigo where numero=" & codigo _
    '        , DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    
    rs.Open cadena, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly


    If Not rs.EOF Then
        Propio = rs!codPropio
        
        RptImpresionRemitoVenta2.lblcliente = sSinNull(rs!nombrefantasia)
        If Not IsNull(rs!CUIT) Then
            RptImpresionRemitoVenta2.lblCuit = rs!CUIT
        End If
        
        COD = rs!numero
        RptImpresionRemitoVenta2.lblcomp = "Remito"
        If Propio Then
            str = "SELECT R.*, P.descripcion,u.descripcion as udes" _
                & " FROM RemitoPorteDetalle r INNER JOIN Producto p ON R.Producto = P.codigo" _
                & " left outer join unidadesmedida u on u.umcodigo=p.umedida where CANCELADO=0 AND numero=" & COD _
                & " order by R.codigo "
        Else
            str = "SELECT r.*, p.descripcion, rpc.ProductoCliente as producto,u.descripcion as udes " _
                & " FROM RemitoPorteDetalle as r left JOIN Producto as p ON r.Producto = p.codigo " _
                & " left join relacion_producto_cliente as rpc on rpc.producto = p.codigo " _
                & " left outer join unidadesmedida u on u.UMcodigo=producto.umedida where numero = " & COD & " and cliente = " & rs!cliente _
                & " order by r.codigo "
        End If
        
        str2 = "select codigo,numero,producto,pedido from remitoventadetalle where numero=" & rs!numero & " and" _
                & " codigo=(select max(codigo) from remitoportedetalle where numero=" & rs!numero & ") " _
                & " group by codigo,numero,producto,pedido"  'terminar,traigo de remideta... pedido para ordcomp
        rs2.Open str2, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly

        
        RptImpresionRemitoVenta2.lblcliente = rs!DESCRIPCION
        RptImpresionRemitoVenta2.NroCli = rs!COD
        
        Dim valu As New ADODB.Recordset
        valu.Open "select * from provincias where codigo='" & rs!Provincia & "'", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
        
        RptImpresionRemitoVenta2.lblNroProv = valu!DESCRIPCION
        Set valu = Nothing
        
        RptImpresionRemitoVenta2.CP = rs!codigopostal_comercial
        RptImpresionRemitoVenta2.lbldomicilio = rs!direccion_comercial
        RptImpresionRemitoVenta2.lblfactura = "0001-" & Format(rs!numero, "00000000")
        RptImpresionRemitoVenta2.lblfecha = rs!Fecha
        RptImpresionRemitoVenta2.lbllocalidad = rs!localidad_comercial & " - " & RptImpresionRemitoVenta2.lblNroProv
        RptImpresionRemitoVenta2.Transportista = trans 'rs!Transportista
        RptImpresionRemitoVenta2.direTRANS = rs!direTRANS
        RptImpresionRemitoVenta2.telTRANS = rs!phone
        RptImpresionRemitoVenta2.lblAtencion = vende 'rs!Vendedor '& rs!apevendedor
        If formas = 0 Then
            RptImpresionRemitoVenta2.forma = ""
        Else
            RptImpresionRemitoVenta2.forma = obtenerDeSQL("select descripcion from formaspago where codigo=" & formas)
        End If
        
        If rs!obs1 <> "" Then
            RptImpresionRemitoVenta2.Label23.caption = "Observacion 1: " & rs!obs1 'frmRemitoVenta.txtObs(0)
        Else
            RptImpresionRemitoVenta2.Label23.caption = ""
        End If
        If rs!obs2 <> "" Then
            RptImpresionRemitoVenta2.Label24.caption = "Observacion 2: " & rs!obs2 'frmRemitoVenta.txtObs(1)
        Else
            RptImpresionRemitoVenta2.Label24.caption = ""
        End If
        valu.Open "select * from ivas where codigo=" & rs!Iva, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
        RptImpresionRemitoVenta2.LblIVA = valu!DESCRIPCION
        Set valu = Nothing
        
        RptImpresionRemitoVenta2.OrdenComp = 0 'rs2!PEDIDO
        
        RptImpresionRemitoVenta2.Dia = Day(rs!Fecha)
        RptImpresionRemitoVenta2.Mes = Month(rs!Fecha)
        RptImpresionRemitoVenta2.Ano = Year(rs!Fecha)
        RptImpresionRemitoVenta2.Meses = PasoMes(Month(rs!Fecha))
         
        LlenarTemp (str)
        RptImpresionRemitoVenta2.DataControl1.Connection = DataEnvironment1.Sistema
        
        ' CAMBIAR PARA QUE EL STR QUE FIGURA SE REDIRECCIONE A LA NUEVA TABLA
        ' QUE CREE
        str = "SELECT * FROM " & sTablaRemito & " "
        RptImpresionRemitoVenta2.DataControl1.Source = str
        
        RptImpresionRemitoVenta2.Printer.PaperSize = pprA4
        RptImpresionRemitoVenta2.Printer.Copies = 1
        If VerDatoEmpresa("idEmpresa") > 1 Then
            With RptImpresionRemitoVenta2
                .CP.Visible = False
            End With
        End If
        
        If PREVIEW_IMPRESIONES Then
            RptImpresionRemitoVenta2.Show
            
            'esto es para acomodar las posiciones del reporte
            'Posicionar (False) 'false para remito true para factura
            
        Else
            RptImpresionRemitoVenta2.PrintReport True
        End If
        'RptImpresionRemitoVenta2.Restart
        
    End If
    rs.Close
    Set rs = Nothing
    
    
fin:
    Exit Function
ErrImpresora:
    ufa "error de impresión Remito Venta", ""
    Resume fin
End Function

Private Sub LlenarTemp(str As String)
   Dim rs As New ADODB.Recordset
   Dim rs2 As New ADODB.Recordset
   Dim AuxDescrip, AuxCodigo, AuxCant, auxudes As String
  
  sTablaRemito = TablaTempCrear("([id] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ,   [cantidad] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,   [codigo] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,   [descrip] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,   [UDES] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL) ON [PRIMARY]")
   
   rs.Open str, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
   Do While Not rs.EOF
      AuxCodigo = rs!producto
      AuxCant = x2s(rs!cantidad)
      AuxDescrip = rs!DESCRIPCION
      auxudes = sSinNull(rs!udes)
      DataEnvironment1.Sistema.Execute "INSERT INTO " & sTablaRemito & " (cantidad,codigo,descrip,udes) VALUES( '" & AuxCant & "','" & AuxCodigo & "','" & AuxDescrip & "','" & auxudes & "')"
      rs2.Open "SELECT serie FROM series WHERE producto = '" & rs!producto & "' and nrocomprobante = '" & rs!numero & "'", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
      Do While Not rs2.EOF
      If Not rs2.EOF Then
         AuxCodigo = "Serie : "
         AuxCant = " "
         AuxDescrip = rs2!Serie
         auxudes = " "
         DataEnvironment1.Sistema.Execute "INSERT INTO " & sTablaRemito & " (cantidad,codigo,descrip,udes) VALUES( '" & AuxCant & "','" & AuxCodigo & "','" & AuxDescrip & "','" & auxudes & "')"
      Else
         AuxCodigo = " "
         AuxCant = " "
         AuxDescrip = " "
         auxudes = " "
         DataEnvironment1.Sistema.Execute "INSERT INTO " & sTablaRemito & "  (cantidad,codigo,descrip,udes) VALUES('" & AuxCant & "','" & AuxCodigo & "','" & AuxDescrip & "','" & auxudes & "')"
      End If
      rs2.MoveNext
      Loop
      rs2.Close
   rs.MoveNext
   Loop
   rs.Close
End Sub


Public Function ImprimirComprobante(codigo, Optional EsReimpresion As Boolean = False, Optional VaConLeyenda As Boolean = False) As Boolean
    If ON_ERROR_HABILITADO Then On Error GoTo ErrImpresion
    
    Dim str1 As String
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim rsnro As New ADODB.Recordset
    Dim str As String
    Dim COD As Long
    Dim PORCENTAJE As String
    Dim tdoc As String, z As Double, mone As Long
    Dim tempo As String, tempo2 As String, tempo3 As String
    Dim impresoraFV As String
    Dim Valor As Long
    
    Dim mImpresoraDefecto As String
    
    Dim pieY As Long, totY As Long, cabY As Long, detY As Long, cabX As Long, prcX As Long
    
'    totY = coord(Totales_x)
    pieY = coord(Pie_Y)
    cabY = coord(Encabezado_Y)
    cabX = coord(Encabezado_X)
    detY = coord(Detalle_Y)
    prcX = coord(porc_x)
    
    ' CodigoPropio: OJO!!!!
    ' se graba propio para cada item, pero aca lo busco una sola vez por performance,
    ' trae el primero que encuentra y asume que todos son propios o todos de cliente
    Dim Propio As Boolean
    Dim strDetalle As String
    '

'    Call setLpt2

    str1 = "SELECT FacturaVenta.*, Clientes.*, FormasPago.descripcion as pago, " & _
            "Ivas.descripcion as iva, Ivas.letra as letra" & _
            " FROM FormasPago INNER JOIN (Ivas INNER JOIN (Clientes INNER JOIN FacturaVenta  " & _
            " ON Clientes.codigo = FacturaVenta.Cliente) " & _
            " ON Ivas.codigo = FacturaVenta.TipoIVA) " & _
            " ON FormasPago.codigo = FacturaVenta.FormaPago " & _
            " WHERE facturaventa.codigo=" & codigo
    rs.Open str1, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    
    tdoc = Trim(rs!TIPODOC)
    z = s2n(rs!cotizacion, 4)
    If z = 0 Then z = 1
    mone = rs!moneda
    If mone = 0 Then mone = 1 '(PESOS)
    
    If Not rs.EOF Then
    
        RptImpresionFacturaVenta.CodPos = sSinNull(rs!codigopostal)
'        pppp cabY + 0, cabX + 48, rs!codigopostal
        
        
        If Not IsNull(rs!RAZONSOCIAL) Then
            RptImpresionFacturaVenta.lblcliente = rs!RAZONSOCIAL
'            pppp cabY + 0, cabX, rs!razonsocial
        End If
        If Not IsNull(rs!CUIT) Then
            RptImpresionFacturaVenta.lblCuit = rs!CUIT
'            pppp cabY + 2, cabX, rs!Cuit
        End If
        COD = codigo
        PORCENTAJE = s2n(rs!PorcentajeIva)
        'RptImpresionFacturaVenta.txtIvaP = PORCENTAJE * 100
        
'        pppp pieY + 2, prcX, PORCENTAJE * 100

        If Not IsNull(rs!Pedido) Then
            RptImpresionFacturaVenta.OrdenComp = rs!Pedido
        End If
        
        RptImpresionFacturaVenta.NroCli = rs!cliente
        
        RptImpresionFacturaVenta.lblNroProv = nSinNull(rs!Proveedor)
        'pppp cabY + 4, cabX, "( " & rs!Proveedor & " )"
'        pppp cabY + 4, cabX, sSinNull(rs!Proveedor)
        
        
        Propio = obtenerDeSQL("select CodPropio from FacturaVentadetalle where codigofactura = " & COD)
        
        If Left(tdoc, 2) = "FA" Then
        
            RptImpresionFacturaVenta.lblcomp = "" '"Factura"
'            pppp 9, 47, "FACTURA" ' Chr(&HE) & " Factura " &
'            pppp coord(labelFactura_Y), coord(labelFactura_X), "FACTURA"
            
            RptImpresionFacturaVenta.Remitos = Format(rs!Remito, "00000000")
'            pppp 21, 36, Format(rs!remito, "00000000")
            
            'detalle
            If tdoc = "FAB" Then
                strDetalle = " d.cantidad,d.descripcion,(d.preciounitario + (d.preciounitario * " & ssNum(PORCENTAJE) & ")) as punit,(d.preciototal + (d.preciototal * " & ssNum(PORCENTAJE) & ")) as ptot "
                RptImpresionFacturaVenta.IvaInsc = ""
            ElseIf tdoc = "FAE" Then
                strDetalle = " d.cantidad,d.descripcion,d.preciounitario / " & ssNum(z) & " as punit ,d.preciototal / " & ssNum(z) & " as ptot "
            Else ' tdoc = "FAA" Then
                strDetalle = " d.cantidad,d.descripcion,d.preciounitario  as punit,d.preciototal  as ptot "
                ' solo Factura A
                If rs!Iva = "NO INSCRIPTO" Then
                    RptImpresionFacturaVenta.IvaInsc = PORCENTAJE * 100
                    RptImpresionFacturaVenta.IvaNoInsc = Format$(s2n(rs!Neto) * s2n(rs!PorcentajeIva), "Standard")
                    RptImpresionFacturaVenta.txtIvaP = ""
                    RptImpresionFacturaVenta.txtivains = ""
                ElseIf rs!Iva = "INSCRIPTO" Then
                    RptImpresionFacturaVenta.txtIvaP = PORCENTAJE * 100
                    RptImpresionFacturaVenta.txtivains = Format$(s2n(rs!Neto) * s2n(rs!PorcentajeIva), "Standard")
                    RptImpresionFacturaVenta.IvaInsc = "1"
                    RptImpresionFacturaVenta.IvaNoInsc = ""
                End If
'                pppp pie + 2, coltot, formatTOT(s2n(rs!neto) * s2n(rs!PORCENTAJEiva))  ' RptImpresionFacturaVenta.txtivains
                RptImpresionFacturaVenta.txtNeto = Format$(rs!Neto, "Standard")
'                pppp pie +

                
                RptImpresionFacturaVenta.txtsub = Format$(rs!Neto, "Standard")
                
                If rs!IIBB <> 0 Then
                    RptImpresionFacturaVenta.txtIibbP = Format$(1 / (rs!Neto / rs!IIBB) * 100, "Standard")
                    RptImpresionFacturaVenta.txtIIBB = Format$(rs!IIBB, "standard")
                    Valor = RptImpresionFacturaVenta.txtIIBB
                    
'                    pppp pie + 4, coltot, formatTOT(rs!iibb) ', "standard")
'                    pppp pie + 4, colp100, formatP100(s2n(1 / (rs!neto / rs!iibb) * 100))
                Else
                    Valor = 0
                End If

                RptImpresionFacturaVenta.Subtotal2 = (RptImpresionFacturaVenta.txtsub - Valor)
                
                
' ACA HAY ALGO QUE NO ESTOY DE ACUERDO CON EL CONTADOR DE TONKA,
' QU ME DIJO QUE PONGA SUBTOTAL, INTERES y TOTAL en vez de SUBTOTAL, INTERES, NETO,  y TOTAL
' (ademas de IIBB e IVA)

                If rs!Descuento <> 0 Then
                    
                    RptImpresionFacturaVenta.txtDctoP = Format$(rs!Descuento * 100)
'                    pppp pie + 1, colp100, formatP100(rs!descuento * 100)
                    
                    Dim t_sub, t_neto, t_coef
                    
                    t_neto = rs!Neto
                    t_coef = Sgn(rs!Descuento) * 0 + (rs!Descuento)
                    't_coef = (rs!Descuento) * 100
                    If Trim(rs!Descuento) = 1 Then
                        t_sub = "0,00"
                    Else
                        t_sub = t_neto / (1 - t_coef)
                    End If
                    'RptImpresionFacturaVenta.txtDcto = Format$(s2n(rs!neto / (Sgn(rs!descuento) + rs!descuento)), "standard")
                    RptImpresionFacturaVenta.txtsub = Format$(t_sub, "Standard")
                    RptImpresionFacturaVenta.txtDcto = Format$(t_neto - t_sub, "standard")
                    
'                    pppp pie + 0, coltot, formatTOT(t_sub)
'                    pppp pie + 1, coltot, formatTOT(t_neto - tsub)
                End If
'            Else
'                pppp pie + 0, coltot, formatTOT(rs!neto)
            End If
            
            

            If Propio Then
                If VerDatoEmpresa("idEmpresa") > 1 Then
                    str = " select d.producto, " & _
                        strDetalle & " ,p.umedida,u.abreviatura as udes " & _
                        " from facturaventadetalle d inner join producto p on p.codigo=d.producto inner join unidadesmedida u on u.umcodigo=p.umedida " & _
                        " where d.codigofactura=" & COD & "  ORDER BY d.id"
                Else
                    str = " select d.producto, " & _
                        strDetalle & " ,0 as umedida,'Serv' as udes " & _
                        " from facturaventadetalle d " & _
                        " where d.codigofactura=" & COD & "  ORDER BY d.id"
                End If
            Else
                str = " select d.productoCliente as producto, " & _
                    strDetalle & " ,p.umedida,u.descripcion as udes " & _
                    " from facturaventadetalle as d left join Relacion_Producto_Cliente as r " & _
                    " on d.producto = r.producto inner join producto p on p.codigo=d.producto inner join unidadesmedida u on u.codigo=p.umedida" & _
                    " where d.codigofactura=" & COD & " and r.cliente = " & rs!cliente & _
                    " ORDER BY d.id"
            End If
        Else
        
            If Left(tdoc, 2) = "NC" Then
                RptImpresionFacturaVenta.lblcomp = "Nota de Credito"
'                pppp 9, 47, "NOTA DE CREDITO"
                
                RptImpresionFacturaVenta.lbltachar = "XXXXXXXXX"
'                pppp 0, 40, "XXXXXXXXXXX"
                If Propio Then
                    'STR = "select d.producto, d.cantidad, d.descripcion, (d.preciounitario + (d.preciounitario * " & Replace(PORCENTAJE, ",", ".") & ")) as punit,(d.preciototal + (d.preciototal * " & Replace(PORCENTAJE, ",", ".") & ")) as ptot,p.umedida,u.descripcion from facturaventadetalle d inner join producto p on p.codigo=d.producto inner join unidadesmedida u on u.umcodigo=p.umedida where d.codigofactura=" & cod
                    str = "select d.producto, d.cantidad, d.descripcion, d.preciounitario  as punit,d.preciototal as ptot,p.umedida,u.descripcion from facturaventadetalle d inner join producto p on p.codigo=d.producto inner join unidadesmedida u on u.umcodigo=p.umedida where d.codigofactura=" & COD & _
                        " union select d.producto, d.cantidad, d.descripcion, d.preciounitario  as punit,d.preciototal as ptot,0 as umedida,'' as descripcion from facturaventadetalle d where d.producto='0' and d.codigofactura=" & COD
                Else
                    str = "select r.productocliente as producto, d.cantidad, d.descripcion, (d.preciounitario + (d.preciounitario * " & Replace(PORCENTAJE, ",", ".") & ")) as punit,(d.preciototal + (d.preciototal * " & Replace(PORCENTAJE, ",", ".") & ")) as ptot " & _
                        " ,p.umedida,u.descripcion as udes from facturaventadetalle as d left join Relacion_Producto_Cliente as r " & _
                        " on d.producto = r.producto inner join producto p on p.codigo=d.producto inner join unidadesmedida u on u.UMCodigo=p.umedida " & _
                        " where d.codigofactura=" & COD & " and r.cliente = " & rs!cliente & _
                        " ORDER BY d.id"
                End If
                
                If tdoc = "NCB" Then
                
                ElseIf tdoc = "NCE" Then
                
                ElseIf tdoc = "NCA" Then
               '     str = "select producto,descripcion from facturaventadetalle where codigofactura=" & cod
                    
                    'RptImpresionFacturaVenta.txtivains = Format$(s2n(rs!neto) * s2n(rs!PORCENTAJEiva), "Standard")
                    If rs!Iva = "NO INSCRIPTO" Then
                        RptImpresionFacturaVenta.IvaInsc = PORCENTAJE * 100
                        RptImpresionFacturaVenta.IvaNoInsc = Format$(s2n(rs!Neto) * s2n(rs!PorcentajeIva), "Standard")
                        RptImpresionFacturaVenta.txtIvaP = ""
                        RptImpresionFacturaVenta.txtivains = ""
                    ElseIf rs!Iva = "INSCRIPTO" Then
                        RptImpresionFacturaVenta.txtIvaP = PORCENTAJE * 100
                        'RptImpresionFacturaVenta.txtivains = Format$((rs!Neto * rs!PorcentajeIva) - rs!Neto, "Standard")
                        RptImpresionFacturaVenta.txtivains = Format$((rs!Neto * rs!PorcentajeIva), "Standard")
                        RptImpresionFacturaVenta.IvaInsc = ""
                        RptImpresionFacturaVenta.IvaNoInsc = ""
                    End If
                    RptImpresionFacturaVenta.txtNeto = Format$(rs!Neto, "Standard")
                    RptImpresionFacturaVenta.txtsub = Format$(rs!Neto, "Standard")
                    
                    'pppp pie+
'                    pppp pie + 2, coltot, formatTOT(s2n(rs!neto) * s2n(rs!PORCENTAJEiva))
'                    pppp pie + 0, coltot, formatTOT(rs!neto)
                End If
            Else
                If Left(tdoc, 2) = "ND" Then
                    RptImpresionFacturaVenta.lblcomp = "Nota de Debito"
                    RptImpresionFacturaVenta.lbltachar = "XXXXXXXXX"
'
'                    pppp
                    
                    If Trim(rs!TIPODOC) = "NDB" Then
                        str = "select d.id,d.producto,d.descripcion,p.umedida,u.descripcion as udes,d.preciototal as ptot from facturaventadetalle d inner join producto p on p.codigo=d.producto inner join unidadesmedida u on u.umcodigo=p.umedida where d.codigofactura=" & COD & _
                            " union select d.id,d.producto, d.descripcion, 0 as umedida,'' as udes,d.preciototal as ptot from facturaventadetalle d where D.PRODUCTO<>'1' AND d.codigofactura=" & COD
                    ElseIf tdoc = "NDE" Then
                        str = "select d.producto, d.cantidad,d.descripcion,(d.preciounitario + (d.preciounitario * " & x2s(PORCENTAJE) & ")) / " & x2s(z) & " as punit,(d.preciototal + (d.preciototal * " & x2s(PORCENTAJE) & ")) " & x2s(z) & " as ptot,p.umedida,u.descripcion as udes from facturaventadetalle d inner join producto p on p.codigo=d.producto inner join unidadesmedida u on u.umcodigo=p.umedida where d.codigofactura=" & COD & "ORDER BY d.id"
                    Else 'NDA
                        str = "select d.id,d.producto,d.descripcion,p.umedida,u.descripcion as udes,d.preciototal as ptot from facturaventadetalle d inner join producto p on p.codigo=d.producto inner join unidadesmedida u on u.umcodigo=p.umedida where d.codigofactura=" & COD & _
                            " union select d.id,d.producto,d.descripcion,0,'' as udes,d.preciototal as ptot from facturaventadetalle d where d.producto<>'1' and d.codigofactura=" & COD
                        
                        'RptImpresionFacturaVenta.txtivains = Format$(s2n(rs!neto) * s2n(rs!PORCENTAJEiva), "Standard")
                        If rs!Iva = "NO INSCRIPTO" Then
                            RptImpresionFacturaVenta.IvaInsc = PORCENTAJE * 100
                            RptImpresionFacturaVenta.IvaNoInsc = Format$(s2n(rs!Neto) * s2n(rs!PorcentajeIva), "Standard")
                            RptImpresionFacturaVenta.txtIvaP = ""
                            RptImpresionFacturaVenta.txtivains = ""
                        ElseIf rs!Iva = "INSCRIPTO" Then
                            RptImpresionFacturaVenta.txtIvaP = IIf(Tipo_NotaDebitoChRechazado = 2, s2n((PORCENTAJE - 1) * 100), PORCENTAJE * 100)
                            'RptImpresionFacturaVenta.txtivains = IIf(Tipo_NotaDebitoChRechazado = 2, (s2n(rs!neto) * s2n(rs!PorcentajeIva)) - (rs!neto), Format$(s2n(rs!neto) * s2n(rs!PorcentajeIva), "Standard"))
                            RptImpresionFacturaVenta.txtivains = IIf(Tipo_NotaDebitoChRechazado = 2, (s2n(rs!Neto) * s2n(rs!PorcentajeIva)), Format$(s2n(rs!Neto) * s2n(rs!PorcentajeIva), "Standard"))
                            RptImpresionFacturaVenta.IvaInsc = IIf(Tipo_NotaDebitoChRechazado = 2, s2n((PORCENTAJE - 1) * 100), PORCENTAJE * 100)
                            RptImpresionFacturaVenta.IvaNoInsc = IIf(Tipo_NotaDebitoChRechazado = 2, (rs!Total - (rs!Neto + rs.Fields(14))) - ((rs!Total - (rs!Neto + rs.Fields(14))) / "1,21"), "")
                            RptImpresionFacturaVenta.txtIIBB.Text = IIf(Tipo_NotaDebitoChRechazado = 2, (rs!Total - (rs!Neto + rs.Fields(14))) / "1,21", 0)
                        End If
                
                        RptImpresionFacturaVenta.txtNeto = Format$(rs!Neto, "Standard")
                        RptImpresionFacturaVenta.txtsub = Format$(rs!Neto, "Standard")
                    End If
                End If
            End If
        End If
        
        If Not IsNull(rs!direccion) Then
            
            
            RptImpresionFacturaVenta.lbldomicilio = rs!direccion
            
        
'            pppp 15, 12, rs!direccion
        Else
            RptImpresionFacturaVenta.lbldomicilio = ""
        End If
        If Not IsNull(rs!Provincia) Then
            Dim cadena As String
            Dim Aux As New ADODB.Recordset
            cadena = "select * from provincias where codigo='" & rs!Provincia & "'"
            Aux.Open cadena, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            
            
            RptImpresionFacturaVenta.Provincia = Aux!DESCRIPCION 'rs!provincia
            
            Set Aux = Nothing
        Else
            RptImpresionFacturaVenta.Provincia = ""
        End If
        
        RptImpresionFacturaVenta.lblfactura = "0001-" & Format(rs!NroFactura, "00000000")
'        pppp 3, 20, "0001-" & Format(rs!NroFactura, "00000000")
        
        RptImpresionFacturaVenta.lblfecha = rs!Fecha
'        pppp 5, 30, rs!Fecha
        RptImpresionFacturaVenta.Dia = Day(rs!Fecha)
        RptImpresionFacturaVenta.Mes = Month(rs!Fecha)
        RptImpresionFacturaVenta.Ano = Year(rs!Fecha)
        RptImpresionFacturaVenta.Meses = PasoMes(Month(rs!Fecha))
        
        RptImpresionFacturaVenta.LblIVA = rs!Iva
        
        If Not IsNull(rs!Localidad) Then
            RptImpresionFacturaVenta.lbllocalidad = rs!Localidad
        Else
            RptImpresionFacturaVenta.lbllocalidad = ""
        End If
        If rs!Remito <> 0 Then
            RptImpresionFacturaVenta.lblRef = "Remito"
            RptImpresionFacturaVenta.lblnroref = "0001-" & Format(rs!Remito, "00000000")
        Else
            If rs!Pedido <> 0 Then
                RptImpresionFacturaVenta.lblRef = "Pedido"
                RptImpresionFacturaVenta.lblnroref = "0001-" & Format(rs!Pedido, "00000000")
            End If
        End If
        
        RptImpresionFacturaVenta.lblpago = rs!pago
        'pppp
        RptImpresionFacturaVenta.lblimp = "Son " & ObtenerDescripcion("Monedas", mone) & ": " & enletras(s2n(rs!Total / z))
        RptImpresionFacturaVenta.txttotalfinal = Format$(s2n(rs!Total / z), "standard")
        
'*************************************************************************************
'        Dim FormatNro, Sql As String
'        Sql = "select distinct nroremito from facturaventadetalle where codigofactura=" & cod
'        rsnro.Open Sql, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
'
'        Do While Not rsnro.EOF
'         If rsnro!NroRemito <> 0 Then
'            FormatNro = String(6 - Len(rsnro!NroRemito), "0") & rsnro!NroRemito
'            RptImpresionFacturaVenta.Remitos.Text = RptImpresionFacturaVenta.Remitos.Text & FormatNro & " - "
'         Else
'            RptImpresionFacturaVenta.Remitos.Text = ""
'         End If
'        rsnro.MoveNext
'        Loop
'        If RptImpresionFacturaVenta.Remitos.Text <> "" Then
'          RptImpresionFacturaVenta.Remitos.Text = Mid(RptImpresionFacturaVenta.Remitos.Text, 1, Len(RptImpresionFacturaVenta.Remitos.Text) - 2)
'        Else
'          RptImpresionFacturaVenta.Remitos.Text = ""
'        End If
'        rsnro.Close
'*************************************************************************************
        
        If tdoc = "FAE" Then
            'VER CUAL ES LA LEYENDA
            If mone <> 1 Then
                RptImpresionFacturaVenta.lblleyenda = " Equivalente a " & x2s(rs!Total) & " Pesos al tipo de cambio " & x2s(z) & " pesos por " & ObtenerDescripcion("Monedas", mone)
            End If
        End If
        
        'RptImpresionFacturaVenta.lblleyenda = "El pago de la presente deberá realizarse en dolares estadounidenses a su vencimiento," _
        & "conforme al valor en dicha moneda expresado en este formulario.El comprador asume que el precio en dolares" _
        & " ha sido condición esencial de esta venta renunciando a invocar el Art 119A de Código Civil." & vbCrLf _
        & "En caso que el pago no pueda realizarse en dicha moneda se realizará en pesos al tipo de cambio vigente para el dolar estadounidense" _
        & " tomando la cotización de tipo vendedor del Banco de la Nación Argentina, al cierre de operaciones del día de efectivo pago; " _
        & "en caso que a la fecha de pago no existiera mercado Libre de cambios en la Ciudad de Buenos Aires se tomaráan las cotizaciones" _
        & " en el Mercado de Nueva York o Montevideo. A opción del vendedor la falta de pago al vencimiento constituye al comprador en mora de " _
        & " pleno derecho y hará devengar un interés punitorio del 20% anual hasta el efectivo pago."
       
        
        RptImpresionFacturaVenta.DataControl1.Connection = DataEnvironment1.Sistema
        RptImpresionFacturaVenta.DataControl1.Source = str
        

        
'        Dim rspppp As New ADODB.Recordset, dddd As Integer
'        rspppp.Open str, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'        dddd = 1
'        While Not rspppp.EOF
'            pppp 20 + dddd, 2, dddd
'            pppp 20 + dddd, 10, rspppp!cantidad
'            pppp 20 + dddd, 20, rspppp!producto
'            pppp 20 + dddd, 34, rspppp!descripcion
'            pppp 20 + dddd, 60, rspppp!punit
'             rspppp.MoveNext
'             dddd = dddd + 1
'        Wend
    End If
    
    If VerDatoEmpresa("idEmpresa") > 1 Then
        With RptImpresionFacturaVenta
            .Label21.Visible = False
            .Label22.Visible = False
            .Label25.Visible = False
            .Label20.Visible = False
            .Label23.Visible = False
            .Label24.Visible = False
        End With
    End If
    
'Call setLpt1
        
'----------------------------------------------------------------------------------------
    If gEMPR_ImprimeCertCalidad Then
    If rs!Certificado And Left(tdoc, 1) = "F" Then
    
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

'                rsnro.Open "Select certificadocalidad from bs", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
'                If Not rsnro.EOF Then
'                    If Not IsNull(rsnro!certificadocalidad) Then
'                        cCertificadoCalidad = rsnro!certificadocalidad + 1
'                        DataEnvironment1.Sistema.Execute "update bs set certificadocalidad=certificadocalidad+1"
'                    End If
'                End If
'                rsnro.Close
'                Set rsnro = Nothing
                cCertificadoCalidad = nSinNull(obtenerDeSQL("Select certificadocalidad from bs")) + 1

                If Not IsNull(rs1!codPropio) And Not IsNull(rs1!producto) Then
                    cCodigoProd = VerProductoCliente(rs1!producto, rs1!codPropio, rs!cliente)
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
            .fieCertificado.DataField = "CERTIFICADO_CALIDAD"
            .fieCliente.DataField = "RAZON_SOCIAL"
            .fieProducto.DataField = "DESCRIPCION_PROD"
            .fieCodCliente.DataField = "CODIGO_PROD"

            .fieCantidad.DataField = "CANTIDAD"
            .fieNroRemito.DataField = "NRO_REMITO"
            .fieMuestra.DataField = "MUESTRA"
            
            If PREVIEW_IMPRESIONES Or EsReimpresion Then
                .Show
            Else
                .PrintReport False
            End If
            
        End With
        DataEnvironment1.Sistema.Execute "update bs set certificadocalidad = certificadocalidad + 1 "
    End If
    End If
    
    
    
'Etiquetas orbis **********************************************

    If rs!etiqueta And Left(tdoc, 1) = "F" Then
     'If rs!Certificado Then
     
       sTablaTemp = TablaTempCrear(tt_Etiquetas_temp)
       Dim rsEti As New ADODB.Recordset
       Dim CodBarra As String, DifCajas As Double
       Dim strInsert, CodProv As String
       Dim code As String
       
       Consulta = "SELECT CodigoFactura, fvd.Producto, Cantidad, Producto.descripcion, " & _
       " rpc.UnidadesxCaja,rpc.letra,rpc.destino, PRODUCTOCLIENTE,CLIENTE " & _
       " FROM Producto INNER JOIN Relacion_Producto_Cliente rpc INNER JOIN " & _
       " FacturaVentaDetalle fvd ON rpc.PRODUCTO = fvd.Producto ON Producto.codigo = fvd.Producto " & _
       " WHERE  cliente = " & rs!cliente & " and CodigoFactura =  '" & codigo & "' "
                                   
        rsEti.Open Consulta, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        CodProv = sSinNull(obtenerDeSQL("select proveedor from clientes where codigo = " & rsEti!cliente & ""))
       
        Do While Not rsEti.EOF
'            CodBarra = rsEti!productocliente + Space(10 - Len(Trim(rsEti!productocliente))) & _
                LCase(Trim(sSinNull(rsEti!destino))) + Space(1 - Len(LCase(sSinNull(rsEti!Letra)))) & _
                Format(rsEti!cantidad, "000000") & _
                Format(CodProv, "000000") & Format(rs!remito, "00000000")

          With rsEti
              CodBarra = ArmarCodeBar(sSinNull(!productocliente), sSinNull(!letra), nSinNull(!cantidad), CodProv, nSinNull(rs!Remito))
              Debug.Print CodBarra
          End With
          
          If rsEti!cantidad > rsEti!UnidadesxCaja And rsEti!UnidadesxCaja <> 0 Then
                                         'tiene q ser <> 0 pq no puede restar las UnixCajas
            Dim cant, UnixCaja, DifCaja As Long
            cant = rsEti!cantidad
            UnixCaja = rsEti!UnidadesxCaja
          
            While cant > 0
               DifCaja = cant - (cant - UnixCaja)
               
              If UnixCaja > cant Then DifCaja = cant  'es para lo q queda en la ultima caja
             
              strInsert = "INSERT INTO " & sTablaTemp & " " & _
                "(ProdCliente,Cantidad,letra,descripcion,destino,remito,CodBarra ) " & _
                " VALUES('" & rsEti!productocliente & "','" & DifCaja & "', " & _
                "'" & LCase(sSinNull(rsEti!letra)) & "','" & rsEti!DESCRIPCION & "', " & _
                "'" & LCase(sSinNull(rsEti!destino)) & "', " & _
                "'" & rs!Remito & "','" & CodBarra & "')"
                DataEnvironment1.Sistema.Execute strInsert
                cant = cant - UnixCaja        'hasta terminar las UnixCajas
                                              'tiene q ser <> 0 pq no puede restar las UnixCajas
            Wend
            Else
                strInsert = "INSERT INTO " & sTablaTemp & " " & _
                    "(ProdCliente,Cantidad,letra,descripcion,destino,remito,CodBarra ) " & _
                    " VALUES('" & rsEti!productocliente & "','" & rsEti!cantidad & "', " & _
                    "'" & LCase(sSinNull(rsEti!letra)) & "','" & rsEti!DESCRIPCION & "', " & _
                    "'" & LCase(sSinNull(rsEti!destino)) & "', " & _
                    "'" & rs!Remito & "','" & CodBarra & "')"
                    DataEnvironment1.Sistema.Execute strInsert
          
          End If
                   
          rsEti.MoveNext
        Loop
        Dim RsEtiqueta As New ADODB.Recordset
        str = "select * from " & sTablaTemp & ""
        RsEtiqueta.Open str, DataEnvironment1.Sistema
        
        RptImpresionEtiqueta.DataEtiqueta.Connection = DataEnvironment1.Sistema
        RptImpresionEtiqueta.DataEtiqueta.Source = str
        
        RptImpresionEtiqueta.lblfecha = Date
        RptImpresionEtiqueta.LblTomka = CodProv
        rsEti.Close
        RsEtiqueta.Close
    'End If
    
    Set rsEti = Nothing
    Set RsEtiqueta = Nothing
             
    RptImpresionEtiqueta.Show '
    Posicionar (True) 'con esto reposiciono los campos del reporte
    
  End If
  
  
  '--------------------------------------------------------------------------------
    ' SI REIMPRIME, QUIZAS SOLO QUIERA REIMPRIMIR ETIQUETAS, NO LA FACTURA
    ' Si uso DATAREPORT, prefiero    if EsReimpresion then .show else .printreport
    
    If VerParametro(BS_FAC_IMPR_MATRIZ) Then
        
        'TONKA ***********************
        If EsReimpresion Then
            If confirma("Reimprimo factura ? ") Then
                ImprimirFacturaRemito codigo, VaConLeyenda
            End If
        Else
            ImprimirFacturaRemito codigo, VaConLeyenda
        End If
        
    Else
        
        'OTROS ************************
        ' COMO PERSONALIZAR !????!!!!????!!!???
        'mImpresoraDefecto = verImprNombre(obtenerParametro("FV_ImpresoraNombre"))


        'Dim pripo As Printer
        'Set pripo = GetPrinterEnPort(Trim(obtenerParametro("impresoraFactura")))
        '
        '
        'If Not pripo Is Nothing Then
        '    Set RptImpresionFacturaVenta.Printer = pripo 'GetPrinterEnPort(Trim(obtenerParametro("impresoraFactura")))
        'End If
'        RptImpresionFacturaVenta.Printer.DeviceName = "\\LAURA\HPLaserJ" 'Trim(obtenerParametro("FV_ImpresoraNombre"))
'        RptImpresionFacturaVenta.Printer.Port = "\\LAURA\HPLaserJ" 'UCase(Trim(obtenerParametro("impresoraFactura")))
        
        RptImpresionFacturaVenta.Printer.Copies = 1
        If VerParametro(BS_PREVIEW_IMPRESIONES) Or EsReimpresion Then
            RptImpresionFacturaVenta.Show
            Posicionar (True)
            RptImpresionFacturaVenta.txtivains.Visible = Not obtenerDeSQL("select nd_xchequerechazado from facturaventa where codigo=" & codigo)
            With RptImpresionFacturaVenta
            If tdoc = "FAB" Then
                .lblfactura.Visible = True
                .lblcomp.Visible = False
                .Label4.Visible = False
                .LblIVA.Visible = False
            End If
            End With
            'reposiciono los campos del reporte,uso true para factura y false para remito
        Else
            '  rptImpresionFacturaVenta.Restart
            RptImpresionFacturaVenta.PrintReport True
        End If
    
    
'        verImprNombre mImpresoraDefecto
    End If
    
  '--------------------------------------------------------------------------------
    

    
    'RptImpresionFacturaVenta.Printer.Copies = 1
    'RptImpresionFacturaVenta.PageFooter.Visible = True
    
    'SetImpresora obtenerParametro("ImpresoraFactura")
'    RptImpresionFacturaVenta.PrintReport True
    
'Prueba: REMITO princ
'        SetImpresora "FACTURACION"
            
        
'    With RptImpresionFacturaVenta
    
        'impresoraFV = sSinNull(obtenerParametro("FV_ImpresoraNombre"))
        'If impresoraFV > "" Then .Printer.DeviceName = impresoraFV
        
'        .Printer.DeviceName = "FACTURACION"        '
    
'
'
'        .PageSettings.TopMargin = s2n(obtenerParametro("FV_TopMargin"))
'        .PageSettings.PaperHeight = s2n(obtenerParametro("FV_PaperHeight"))
'        '.PageSettings.BottomMargin = s2n(obtenerParametro("FV_BottomMargin"))

        ' IMPRIME FACTURA

        
        'setLpt2
'        .PageSettings.PaperSize = vbPRPSFanfoldStdGerman
'
'    'cambio porque lo hice al reves
'            'IMPRIME REMITO
'            'cambio
'            tempo = .lblfactura
'            .lblfactura = .Remitos
'            .Remitos = tempo
'            .PageFooter.Visible = False
'
'            tempo2 = .lblcomp
'            .lblcomp = "REMITO"
'
'            tempo3 = .lbltachar
'            .lbltachar = ""
'            '.PrintReport False
'           .Show
'
'
'            'IMPRIME FACTURA
'            'cambio
'            tempo = .lblfactura
'            .lblfactura = .Remitos
'            .Remitos = tempo
'
'            .lblcomp = tempo2
'            .lbltachar = tempo3
'            .PageFooter.Visible = True
'
'            .Restart
'            '.PrintReport True
'            .Show

    'aaFormFV.Show

'    End With
        
    'SetImpresora "           " ' la default
'    setLpt1
    
'Prueba: REMITO fin


    rs.Close

fin:
'    On Error Resume Next
'    If mImpresoraDefecto > "" Then verImprNombre (mImpresoraDefecto)
'finfin:
    Set rs = Nothing
    Exit Function
ErrImpresion:
    ufa "Error de impresión en factura venta", ""
'    Call setLpt1
    Resume fin
End Function

Public Function ImprimirComprobanteFE(codigo, Optional EsReimpresion As Boolean = False, Optional VaConLeyenda As Boolean = False) As Boolean
    If ON_ERROR_HABILITADO Then On Error GoTo ErrImpresion
    Dim RptImpresionFacturaVentaFE2 As New RptImpresionFacturaVentaFE
    Dim strConsulta1 As String, strConsulta2 As String
    Dim rs As New ADODB.Recordset
    Dim strCadena As String, strDetalle As String, strLeyenda As String, rsley As New ADODB.Recordset, i As Long
    Dim tdoc As String, z As Double, nTotal As Double, NPedido As String, nOrden As String
    Dim Propio As Boolean
    

    strConsulta1 = "SELECT FacturaVenta.*, Clientes.*, FormasPago.descripcion as pago, " & _
            "Ivas.descripcion as iva, Ivas.letra as letra" & _
            " FROM FormasPago INNER JOIN (Ivas INNER JOIN (Clientes INNER JOIN FacturaVenta  " & _
            " ON Clientes.codigo = FacturaVenta.Cliente) " & _
            " ON Ivas.codigo = FacturaVenta.TipoIVA) " & _
            " ON FormasPago.codigo = FacturaVenta.FormaPago " & _
            " WHERE facturaventa.codigo=" & codigo
    rs.Open strConsulta1, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    
    tdoc = Trim(rs!TIPODOC)
    
    
    If rs.EOF And rs.BOF Then
        GoTo ErrImpresion
    Else

        'leyenda de la factura
        If Right(tdoc, 1) = "E" Then
            If MsgBox("IMPRIME LEYENDA EN FACTURA ELECTRONICA...", vbInformation + vbYesNo) = vbYes Then
                rsley.Open "select * from FacturaVentaLeyendaFae where cliente=" & s2n(rs!cliente) & " order by id", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
                strLeyenda = ""
                With rsley
                    If .EOF And .BOF Then
                        strLeyenda = "ORIGEN ARGENTINO"
                    Else
                        .MoveFirst
                        For i = 0 To .RecordCount - 1
                            strLeyenda = strLeyenda & !RENGLON & Chr(13)
                            .MoveNext
                        Next
                    End If
                End With
            End If
        End If
        nOrden = ""
        NPedido = nSinNull(obtenerDeSQL("select nropedido from facturaventadetalle where codigofactura=" & codigo))
        If s2n(NPedido) > 0 Then
            nOrden = sSinNull(obtenerDeSQL("select pedido_cli from pedidos_clientes where numero=" & NPedido))
        End If
        If Trim(nOrden) > "" Then
            'strLeyenda = strLeyenda & "ORDEN DE COMPRA: " & nOrden & Chr(13)
            RptImpresionFacturaVentaFE.lblOc = nOrden
        End If
        'RptImpresionFacturaVentaFE.lblPuntoVenta = Format(Trim(obtenerDeSQL("select puntoventaFE from datosempresa where idempresa=" & gEMPR_idEmpresa)), "0000")
        RptImpresionFacturaVentaFE.lblPuntoVenta = Format(rs!PuntoVenta, "0000")
        RptImpresionFacturaVentaFE.lblComprobanteNro = RptImpresionFacturaVentaFE.lblPuntoVenta & "-" & Format(sSinNull(rs!NroFactura), "00000000")
        RptImpresionFacturaVentaFE.lblFechaEmision = rs!Fecha
        RptImpresionFacturaVentaFE.lblCuit = Trim(obtenerDeSQL("select CUITEMPRESA from datosempresa where idempresa=" & gEMPR_idEmpresa))
        RptImpresionFacturaVentaFE.lblIngresosBrutos = Trim(obtenerDeSQL("select CUITEMPRESA from datosempresa where idempresa=" & gEMPR_idEmpresa)) & " Jur. sede: 901" 'Trim(obtenerDeSQL("select iibbempresa from datosempresa where idempresa=" & gEMPR_idEmpresa))
        RptImpresionFacturaVentaFE.lblfechainicioactividad = obtenerDeSQL("select fechainicio from datosempresa where idempresa=" & gEMPR_idEmpresa)
        RptImpresionFacturaVentaFE.lblRAZONSOCIAL01 = UCase(Trim(obtenerDeSQL("select nombre from datosempresa where idempresa=" & gEMPR_idEmpresa)))
        RptImpresionFacturaVentaFE.lblRazonSocial02 = Trim(obtenerDeSQL("select nombre from datosempresa where idempresa=" & gEMPR_idEmpresa))
        RptImpresionFacturaVentaFE.lblDomicilioComercial = Trim(obtenerDeSQL("select DIRECCION from datosempresa where idempresa=" & gEMPR_idEmpresa))
        RptImpresionFacturaVentaFE.lblCondicionIVA = "Responsable Inscripto"
        RptImpresionFacturaVentaFE.lblDestinoComprobante = sSinNull(rs!Pais)
        RptImpresionFacturaVentaFE.lblPermisoEmbarque = ""
        RptImpresionFacturaVentaFE.lblPaisDestino = sSinNull(rs!Pais)
        RptImpresionFacturaVentaFE.lblSenior = sSinNull(rs!RAZONSOCIAL)
        RptImpresionFacturaVentaFE.lblDomicilioSenior = sSinNull(rs!direccion) & " - " & sSinNull(rs!Localidad) & " - " & sSinNull(rs!Pais)
        RptImpresionFacturaVentaFE.lblCuitSenior = sSinNull(rs!CUIT)
        RptImpresionFacturaVentaFE.lblIDImpositivo = ""
        RptImpresionFacturaVentaFE.lblFormaPago = rs!pago
        RptImpresionFacturaVentaFE.lblIncoterms = sSinNull(rs!incoterms)
        z = s2n(rs!cotizacion, 4, True)
        If z = 0 Then z = 1
        RptImpresionFacturaVentaFE.lblSubTotal = s2n((s2n(rs!Neto) - s2n(rs!segurofae) - s2n(rs!fletefae)) / z, 2, True)
        RptImpresionFacturaVentaFE.lblFlete = s2n(s2n(rs!fletefae) / z, 2, True)
        RptImpresionFacturaVentaFE.lblSeguro = s2n(s2n(rs!segurofae) / z, 2, True)
        nTotal = s2n(s2n(RptImpresionFacturaVentaFE.lblSubTotal) + s2n(RptImpresionFacturaVentaFE.lblFlete) + s2n(RptImpresionFacturaVentaFE.lblSeguro))
        RptImpresionFacturaVentaFE.lblImporteTotal = s2n(nTotal, 2, True)
        RptImpresionFacturaVentaFE.lblSon = "Son " & ObtenerDescripcion("Monedas", nSinNull(rs!moneda)) & ": " & enletras(s2n(nTotal))
        RptImpresionFacturaVentaFE.lblCAE = sSinNull(rs!CAE)
        RptImpresionFacturaVentaFE.lblFechaCAE = sSinNull(rs!CAEV)
        RptImpresionFacturaVentaFE.lblLeyendaAutorizo.Visible = Trim(RptImpresionFacturaVentaFE.lblCAE) > ""
        RptImpresionFacturaVentaFE.lblBARRA = sSinNull(rs!barra)
        RptImpresionFacturaVentaFE.lblBarraNro = sSinNull(rs!barra)
        RptImpresionFacturaVentaFE.lblOrigen = strLeyenda
        'RptImpresionFacturaVentaFE.provincia = rs!provincia
        'RptImpresionFacturaVentaFE.lbllocalidad = rs!Localidad
        'RptImpresionFacturaVentaFE.lblimp = "Son " & ObtenerDescripcion("Monedas", mone) & ": " & enletras(s2n(rs!Total / z))

        
        strDetalle = " cantidad,descripcion,preciounitario / " & ssNum(z) & " as punit ,preciototal / " & ssNum(z) & " as ptot "
        'strDetalle = " cantidad,descripcion,preciounitario  as punit ,preciototal  as ptot "
        
        Propio = obtenerDeSQL("select codpropio from facturaventadetalle where codigofactura= " & codigo)
        If Propio Then
        'If True Then
            strConsulta2 = " select producto, " & _
                strDetalle & _
                " from facturaventadetalle " & _
                " where codigofactura=" & codigo & " ORDER BY id"
        Else
            strConsulta2 = " select productoCliente as producto, " & _
                strDetalle & _
                " from facturaventadetalle as d left join Relacion_Producto_Cliente as r " & _
                " on d.producto = r.producto " & _
                " where r.activo=1 and codigofactura=" & codigo & " and r.cliente = " & rs!cliente & _
                " ORDER BY id"
        End If
    

        
        RptImpresionFacturaVentaFE.DataControl1.Connection = DataEnvironment1.Sistema
        RptImpresionFacturaVentaFE.DataControl1.Source = strConsulta2
        RptImpresionFacturaVentaFE.Printer.DeviceName = "Universal Document Converter"
        RptImpresionFacturaVentaFE.Printer.Copies = 1
        RptImpresionFacturaVentaFE.Printer.FromPage = 1
        RptImpresionFacturaVentaFE.Printer.ToPage = 1
        
        Dim feCODFACTURA As String
        feCODFACTURA = sSinNull(obtenerDeSQL("SELECT CODFACTURA FROM DOCUMENTOSCAE WHERE TIPO=" & ssTexto(tdoc) & " AND PUNTOVENTA=" & ssTexto(RptImpresionFacturaVentaFE.lblPuntoVenta)))
        RptImpresionFacturaVentaFE.lblCODFACTURA.caption = "COD. " & feCODFACTURA
        RptImpresionFacturaVentaFE.documentName = "FacturaElectronicaORIGINAL_" & tdoc & "" & RptImpresionFacturaVentaFE.lblComprobanteNro & ".fe"
        
        
        Dim Letra1 As String, Letra2 As String, Letra3 As String
        Dim NombreDocumento As String
        Letra1 = CORTO(tdoc, 0, 2)
        Letra2 = CORTO(tdoc, 1, 1)
        Letra3 = CORTO(tdoc, 2, 0)
        
        If Letra3 = "A" Then
            RptImpresionFacturaVentaFE.lblLETRA = "A"
        ElseIf Letra3 = "B" Then
            RptImpresionFacturaVentaFE.lblLETRA = "B"
        ElseIf Letra3 = "E" Then
            RptImpresionFacturaVentaFE.lblLETRA = "E"
        End If
        
        If Letra1 = "F" Then
            If Letra3 = "A" Then
                RptImpresionFacturaVentaFE.lblTIPOFACTURA = "FACTURA 'A' "
            ElseIf Letra3 = "B" Then
                RptImpresionFacturaVentaFE.lblTIPOFACTURA = "FACTURA 'B' "
            ElseIf Letra3 = "E" Then
                RptImpresionFacturaVentaFE.lblTIPOFACTURA = "FACTURA DE EXPORTACIÓN"
            End If
        ElseIf Letra1 = "N" Then
            If Letra2 = "C" Then
                If Letra3 = "A" Then
                    RptImpresionFacturaVentaFE.lblTIPOFACTURA = "NOTA DE CREDITO 'A' "
                ElseIf Letra3 = "B" Then
                    RptImpresionFacturaVentaFE.lblTIPOFACTURA = "NOTA DE CREDITO 'B' "
                ElseIf Letra3 = "E" Then
                    RptImpresionFacturaVentaFE.lblTIPOFACTURA = "NOTA DE CREDITO DE EXP."
                End If
            ElseIf Letra2 = "D" Then
                If Letra3 = "A" Then
                    RptImpresionFacturaVentaFE.lblTIPOFACTURA = "NOTA DE DEBITO 'A' "
                ElseIf Letra3 = "B" Then
                    RptImpresionFacturaVentaFE.lblTIPOFACTURA = "NOTA DE DEBITO 'B' "
                ElseIf Letra3 = "E" Then
                    RptImpresionFacturaVentaFE.lblTIPOFACTURA = "NOTA DE DEBITO DE EXP."
                End If
            End If
        End If
        
        
        
        
        
        'RptImpresionFacturaVentaFE.documentName = "FacturaElectronicaORIGINAL_" & RptImpresionFacturaVentaFE.lblComprobanteNro & ".fe"
        'RptImpresionFacturaVentaFE2.DataControl1.Connection = DataEnvironment1.Sistema
        'RptImpresionFacturaVentaFE2.DataControl1.Source = strConsulta2
        'RptImpresionFacturaVentaFE2.Printer.DeviceName = "Universal Document Converter"
        'RptImpresionFacturaVentaFE2.Printer.Copies = 1
        'RptImpresionFacturaVentaFE2.documentName = "FacturaElectronicaDUPLICADO_" & RptImpresionFacturaVentaFE.lblComprobanteNro & ".fe"
        'RptImpresionFacturaVentaFE2.lblORIGINAL = "DUPLICADO"

        
        'If VerParametro(BS_PREVIEW_IMPRESIONES) Then
            'RptImpresionFE(0).Show
            'RptImpresionFE(1).Show
            'RptImpresionFacturaVentaFE.pages.InsertNew 0
            'RptImpresionFacturaVentaFE.pages.Item
            'RptImpresionFacturaVentaFE.pages.InsertNew 1
            
            'RptImpresionFacturaVentaFE2.Show
        'Else
            'RptImpresionFacturaVentaFE.pages
            'Dim pg As Canvas
            'RptImpresionFacturaVentaFE.Printer.StartJob "Factura Electronica"
            'For Each pg In RptImpresionFacturaVentaFE.pages
            '    RptImpresionFacturaVentaFE.Printer.PrintPage pg
            'Next
            'RptImpresionFacturaVentaFE.Printer.EndJob
            'RptImpresionFacturaVentaFE.PrintReport False
            'RptImpresionFacturaVentaFE2.PrintReport False
        'End If
        
        'Set rptFACTURAELECTRONICA.srpt1 = RptImpresionFacturaVentaFE
        'Set rptFACTURAELECTRONICA.srpt2 = RptImpresionFacturaVentaFE
        'rptFACTURAELECTRONICA.Show
        
        'RptImpresionFacturaVentaFE.pages.InsertNew 0
        'RptImpresionFacturaVentaFE.pages.Insert 0, RptImpresionFacturaVentaFE.Canvas
        'RptImpresionFacturaVentaFE.pages.InsertNew 1
        'RptImpresionFacturaVentaFE.pages.Insert 1, RptImpresionFacturaVentaFE.Canvas
                
        RptImpresionFacturaVentaFE.Show
        RptImpresionFacturaVentaFE.PrintReport False
        
        'Dim rptCopia As New RptImpresionFacturaVentaFE
        'rptCopies.pages.InsertNew 0
        'rptCopies.pages.InsertNew 1
        'rptCopies.Canvas = RptImpresionFacturaVentaFE
        'rptCopies.Show
    End If
    
    
Set rs = Nothing
Exit Function
ErrImpresion:
MsgBox "Error de impresión en factura electronica", vbCritical
End Function 'FIN IMPRIMIRCOMPROBANTEFE

Public Function ImprimirComprobanteFE2(codigo, Optional EsReimpresion As Boolean = False, Optional VaConLeyenda As Boolean = False) As Boolean
    If ON_ERROR_HABILITADO Then On Error GoTo ErrImpresion
    Dim RptImpresionFacturaVentaFE3 As New RptImpresionFacturaVentaFE
    Dim strConsulta1 As String, strConsulta2 As String
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim strCadena As String, strDetalle As String, strLeyenda As String, rsley As New ADODB.Recordset, i As Long
    Dim tdoc As String, z As Double, nTotal As Double, NPedido As String, nOrden As String
    Dim Propio As Boolean
    Dim provi As String
        

    strConsulta1 = "SELECT FacturaVenta.*,FacturaVenta.nrocliente as nro_cliente , Clientes.*, FormasPago.descripcion as pago, " & _
            "Ivas.descripcion as iva, Ivas.letra as letra,FacturaVenta.provincia,facturaventa._docum_ve as remi,facturaventa._control_ve as orde,FacturaVenta.codigo as cod,clientes.provincia as provi,facturaventa.usuario_alta as UsuA " & _
            " FROM FormasPago INNER JOIN (Ivas INNER JOIN (Clientes INNER JOIN FacturaVenta  " & _
            " ON Clientes.codigo = FacturaVenta.Cliente) " & _
            " ON Ivas.codigo = FacturaVenta.TipoIVA) " & _
            " ON FormasPago.codigo = FacturaVenta.FormaPago " & _
            " WHERE facturaventa.codigo=" & codigo
    rs.Open strConsulta1, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    
    tdoc = Trim(rs!TIPODOC)
    
    If rs.EOF And rs.BOF Then
        GoTo ErrImpresion
    Else

'        'leyenda de la factura
'        If Right(tdoc, 1) = "E" Then
'            If MsgBox("IMPRIME LEYENDA EN FACTURA ELECTRONICA...", vbInformation + vbYesNo) = vbYes Then
'                rsley.Open "select * from FacturaVentaLeyendaFae where cliente=" & s2n(rs!cliente) & " order by id", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
'                strLeyenda = ""
'                With rsley
'                    If .EOF And .BOF Then
'                        strLeyenda = "ORIGEN ARGENTINO"
'                    Else
'                        .MoveFirst
'                        For i = 0 To .RecordCount - 1
'                            strLeyenda = strLeyenda & !RENGLON & Chr(13)
'                            .MoveNext
'                        Next
'                    End If
'                End With
'            End If
'        End If


        nOrden = ""
'        NPedido = nSinNull(obtenerDeSQL("select nropedido from facturaventadetalle where codigofactura=" & codigo))
'        If s2n(NPedido) > 0 Then
'            nOrden = sSinNull(obtenerDeSQL("select pedido_cli from pedidos_clientes where numero=" & NPedido))
'        End If
'        If Trim(nOrden) > "" Then
            'strLeyenda = strLeyenda & "ORDEN DE COMPRA: " & nOrden & Chr(13)
            RptImpresionFacturaVentaFE2.lblOc = sSinNull(rs!orde) 'nOrden
'        End If
        rs1.Open "SELECT leyenda FROM facturaventaleyenda where activo=1 and fac=" & rs!COD, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
        While Not rs1.EOF
            If RptImpresionFacturaVentaFE2.lblleyenda = "" Then
                RptImpresionFacturaVentaFE2.lblleyenda = s2t(rs1, "leyenda")
            Else
                RptImpresionFacturaVentaFE2.lblleyenda = RptImpresionFacturaVentaFE2.lblleyenda & Chr(13) & s2t(rs1, "leyenda")
            End If
            rs1.MoveNext
        Wend
        Set rs1 = Nothing
        
        'leyenda fija
        'RptImpresionFacturaVentaFE2.lblleyendaFija = "En caso de realizar el pago via transferencia Bancaria, remitir comprobante de la misma a: cobranzas@bacigaluppi.com"
        RptImpresionFacturaVentaFE2.lblleyendaFija = "Favor realizar el pago via transferencia bancaria, remitiendo el comprobante a:   cobranzas@bacigaluppi.com"
        'If rs!valeyendacotizacion = True Then
        '    RptImpresionFacturaVentaFE2.lblleyendaFija = "MONEDA : DOLARES ESTADOUNIDESES" & Chr(13) & "A EFECTOS CONTABLES E IMPOSITIVOS EL TIPO DE CAMBIO ES A U$S 1 (DOLARES)= $ " & s2n(rs!cotizacionleyenda, 2, True)
        'Else
        '    RptImpresionFacturaVentaFE2.lblleyendaFija = ""
        'End If
    
        'RptImpresionFacturaVentaFE2.lblPuntoVenta = Format(Trim(obtenerDeSQL("select puntoventaFE from datosempresa where idempresa=" & gEMPR_idEmpresa)), "0000")
        RptImpresionFacturaVentaFE2.lblPuntoVenta = Format(rs!PuntoVenta, "0000")
        RptImpresionFacturaVentaFE2.lblComprobanteNro = RptImpresionFacturaVentaFE2.lblPuntoVenta & "-" & Format(sSinNull(rs!NroFactura), "00000000")
        RptImpresionFacturaVentaFE2.lblRemito = sSinNull(rs!remi) 'RptImpresionFacturaVentaFE2.lblComprobanteNro
        RptImpresionFacturaVentaFE2.lblNroCliente = sSinNull(rs!Nro_Cliente)
        RptImpresionFacturaVentaFE2.lblIniciales = sSinNull(obtenerDeSQL("select inicial from usuarios where codigo=" & s2n(rs!usua)))
        RptImpresionFacturaVentaFE2.lblIniciales2 = sSinNull(obtenerDeSQL("select inicial from usuarios where codigo=" & s2n(rs!Vendedor)))
        
        RptImpresionFacturaVentaFE2.lblFechaEmision = rs!Fecha
        RptImpresionFacturaVentaFE2.lblFechaVencimiento = rs!vencimiento
        RptImpresionFacturaVentaFE2.lblCuit = Trim(obtenerDeSQL("select CUITEMPRESA from datosempresa where idempresa=" & gEMPR_idEmpresa))
        RptImpresionFacturaVentaFE2.lblIngresosBrutos = Trim(obtenerDeSQL("select CUITEMPRESA from datosempresa where idempresa=" & gEMPR_idEmpresa)) & " Jur. sede: 901" 'Trim(obtenerDeSQL("select iibbempresa from datosempresa where idempresa=" & gEMPR_idEmpresa))
        RptImpresionFacturaVentaFE2.lblfechainicioactividad = obtenerDeSQL("select fechainicio from datosempresa where idempresa=" & gEMPR_idEmpresa)
        RptImpresionFacturaVentaFE2.lblRAZONSOCIAL01 = UCase(Trim(obtenerDeSQL("select nombre from datosempresa where idempresa=" & gEMPR_idEmpresa)))
        RptImpresionFacturaVentaFE2.lblRazonSocial02 = Trim(obtenerDeSQL("select nombre from datosempresa where idempresa=" & gEMPR_idEmpresa))
        RptImpresionFacturaVentaFE2.lblDomicilioComercial = Trim(obtenerDeSQL("select DIRECCION from datosempresa where idempresa=" & gEMPR_idEmpresa))
        RptImpresionFacturaVentaFE2.lblCondicionIVA = "Responsable Inscripto"
        RptImpresionFacturaVentaFE2.lblRenglon02Izq = "Tel: " & obtenerDeSQL("select telefono from datosempresa where idempresa=" & gEMPR_idEmpresa)
        'RptImpresionFacturaVentaFE2.lblDestinoComprobante = sSinNull(rs!Pais)
        'RptImpresionFacturaVentaFE2.lblPermisoEmbarque = ""
        'RptImpresionFacturaVentaFE2.lblPaisDestino = sSinNull(rs!Pais)
        RptImpresionFacturaVentaFE2.lblSenior = sSinNull(rs!RAZONSOCIAL)
        provi = ""
'        If sSinNull(rs!Provincia) <> "" Then
        If sSinNull(rs!provi) <> "" Then
            provi = " - " & sSinNull(obtenerDeSQL("select descripcion from provincias where codigo='" & Trim(rs!provi) & "'"))
        End If
        RptImpresionFacturaVentaFE2.lblDomicilioSenior = sSinNull(rs!direccion) & " - " & sSinNull(rs!Localidad) & provi & " - " & sSinNull(rs!Pais)
        RptImpresionFacturaVentaFE2.lblCuitSenior = sSinNull(rs!CUIT)
        RptImpresionFacturaVentaFE2.lblIDImpositivo = ""
        RptImpresionFacturaVentaFE2.lblFormaPago = rs!pago
        RptImpresionFacturaVentaFE2.lblIncoterms = sSinNull(rs!incoterms)
        RptImpresionFacturaVentaFE2.lblIvaC = nSinNull(obtenerDeSQL("select descripcion from ivas where codigo=" & rs!tipoiva)) 'sSinNull(rs!incoterms)
        z = s2n(rs!cotizacion, 4, True)
        If z = 0 Then z = 1
        If z > 1 Then
            RptImpresionFacturaVentaFE2.lblCotizacion.Visible = True
            RptImpresionFacturaVentaFE2.lblCotizacion.caption = "Equivale a U$S 1(dolares)= $ " & z & "(pesos)"
'            RptImpresionFacturaVentaFE2.lblCotizacion.caption = "En caso de realizar el pago via transferencia Bancaria, remitir comprobante de la misma a: cobranzas@bacigaluppi.com"
            RptImpresionFacturaVentaFE2.Field4.Text = "U$S "
            RptImpresionFacturaVentaFE2.Field5.Text = "U$S "
            RptImpresionFacturaVentaFE2.Field6.Text = "U$S "
            RptImpresionFacturaVentaFE2.Field7.Text = "U$S "
            RptImpresionFacturaVentaFE2.Field8.Text = "U$S "
            RptImpresionFacturaVentaFE2.Field9.Text = "U$S "
        Else
            RptImpresionFacturaVentaFE2.lblCotizacion.Visible = False
            RptImpresionFacturaVentaFE2.lblCotizacion.caption = ""
            RptImpresionFacturaVentaFE2.Field4.Text = "$ "
            RptImpresionFacturaVentaFE2.Field5.Text = "$ "
            RptImpresionFacturaVentaFE2.Field6.Text = "$ "
            RptImpresionFacturaVentaFE2.Field7.Text = "$ "
            RptImpresionFacturaVentaFE2.Field8.Text = "$ "
            RptImpresionFacturaVentaFE2.Field9.Text = "$ "
        End If
        RptImpresionFacturaVentaFE2.lblSubTotal = s2n((s2n(rs!Neto) - s2n(rs!segurofae) - s2n(rs!fletefae)) / z, 2, True)
        If Right(tdoc, 1) = "A" Then
            If rs!Descuento <> 0 Then
                RptImpresionFacturaVentaFE2.lblDto = s2n(s2n(rs!Neto, 2) - s2n(rs!Neto / (1 - rs!Descuento), 2)) / z
            Else
                RptImpresionFacturaVentaFE2.lblDto = 0
            End If
                
            'RptImpresionFacturaVentaFE2.lblDto = s2n(s2n(rs!fletefae) / z, 2, True)
'            RptImpresionFacturaVentaFE2.lblIva21 = Format$(s2n(rs!Neto) * s2n(rs!PorcentajeIva, 3), "Standard")
'            RptImpresionFacturaVentaFE2.LblIVA = s2n(rs!PorcentajeIva, 3) * 100
            '*************************************************
            RptImpresionFacturaVentaFE2.lblIva21 = Format$(s2n(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where (codigofactura=" & rs!COD & ") and _iva=21") / z, 2) * s2n(0.21, 4), "standard")
            RptImpresionFacturaVentaFE2.lblIva10 = Format$(s2n(obtenerDeSQL("select sum(preciototal) from facturaventadetalle where (codigofactura=" & rs!COD & ") and _iva=10.5") / z, 2) * s2n(0.105, 4), "standard")
            '**********************************************************************************
            RptImpresionFacturaVentaFE2.lblIIBB = s2n(rs!IIBB / z, 2, True)
            nTotal = s2n((s2n(RptImpresionFacturaVentaFE2.lblSubTotal.caption) + s2n(RptImpresionFacturaVentaFE2.lblDto.caption) + s2n(RptImpresionFacturaVentaFE2.lblIva21.caption) + s2n(RptImpresionFacturaVentaFE2.lblIIBB.caption) + s2n(RptImpresionFacturaVentaFE2.lblIva10.caption)), 8)
        ElseIf Right(tdoc, 1) = "B" Then
            RptImpresionFacturaVentaFE2.lblIIBB.Visible = False
            RptImpresionFacturaVentaFE2.Label30.Visible = False
            RptImpresionFacturaVentaFE2.LblIVA.Visible = False
            RptImpresionFacturaVentaFE2.Label31.Visible = False
            RptImpresionFacturaVentaFE2.Label32.Visible = False
            RptImpresionFacturaVentaFE2.lblIva10.Visible = False
            RptImpresionFacturaVentaFE2.lblIva21.Visible = False
            RptImpresionFacturaVentaFE2.Label28.Visible = False
            RptImpresionFacturaVentaFE2.lblSubTotal.Visible = False
            RptImpresionFacturaVentaFE2.lblDto.Visible = False
            RptImpresionFacturaVentaFE2.lblRenglon10Der.Visible = False
            RptImpresionFacturaVentaFE2.lblRenglon15Der.Visible = False
            If rs!Descuento <> 0 Then
                RptImpresionFacturaVentaFE2.lblDto = s2n(s2n(rs!Neto, 2) - s2n(rs!Neto / (1 - rs!Descuento), 2) / z)
            Else
                RptImpresionFacturaVentaFE2.lblDto = 0
            End If
            nTotal = s2n(s2n(rs!Total) / z)
        End If
                
        RptImpresionFacturaVentaFE2.lblImporteTotal = s2n(nTotal, 2, True)  's2n(nTotal * z, 2, True)
        RptImpresionFacturaVentaFE2.lblSon = "Son " & ObtenerDescripcion("Monedas", nSinNull(rs!moneda)) & ": " & enletras(s2n(RptImpresionFacturaVentaFE2.lblImporteTotal))
        RptImpresionFacturaVentaFE2.lblCAE = sSinNull(rs!CAE)
        RptImpresionFacturaVentaFE2.lblFechaCAE = sSinNull(rs!CAEV)
        RptImpresionFacturaVentaFE2.lblLeyendaAutorizo.Visible = Trim(RptImpresionFacturaVentaFE2.lblCAE) > ""
        RptImpresionFacturaVentaFE2.lblBARRA = sSinNull(rs!barra)
        RptImpresionFacturaVentaFE2.lblBarraNro = sSinNull(rs!barra)
'        RptImpresionFacturaVentaFE2.lblOrigen = strLeyenda
        'RptImpresionFacturaVentaFE2.provincia = rs!provincia
        'RptImpresionFacturaVentaFE2.lbllocalidad = rs!Localidad
        'RptImpresionFacturaVentaFE2.lblimp = "Son " & ObtenerDescripcion("Monedas", mone) & ": " & enletras(s2n(rs!Total / z))

        
        If tdoc = "FAB" Then
            'PORCENTAJE = s2n(rs!porcentajeiva)
            strDetalle = " _iva as iva,cantidad,descripcion,(preciounitario + (preciounitario * " & ssNum(rs!PorcentajeIva) & "))  / " & ssNum(z) & "  as punit,(preciototal + (preciototal * " & ssNum(rs!PorcentajeIva) & "))  / " & ssNum(z) & "  as ptot "
        Else
            strDetalle = " _iva as iva,cantidad,descripcion,preciounitario / " & ssNum(z) & " as punit ,preciototal / " & ssNum(z) & " as ptot "
            'strDetalle = " cantidad,descripcion,preciounitario  as punit ,preciototal  as ptot "
        End If
        
        Propio = obtenerDeSQL("select codpropio from facturaventadetalle where codigofactura= " & codigo)
        If Propio Then
        'If True Then
            strConsulta2 = " select producto, " & _
                strDetalle & _
                " from facturaventadetalle " & _
                " where codigofactura=" & codigo & " ORDER BY id"
        Else
            strConsulta2 = " select productoCliente as producto, " & _
                strDetalle & _
                " from facturaventadetalle as d left join Relacion_Producto_Cliente as r " & _
                " on d.producto = r.producto " & _
                " where r.activo=1 and codigofactura=" & codigo & " and r.cliente = " & rs!cliente & _
                " ORDER BY id"
        End If
    

        Dim Impresora As String
        RptImpresionFacturaVentaFE2.DataControl1.Connection = DataEnvironment1.Sistema
        RptImpresionFacturaVentaFE2.DataControl1.Source = strConsulta2
        Impresora = RptImpresionFacturaVentaFE2.Printer.DeviceName
        RptImpresionFacturaVentaFE2.Printer.DeviceName = "Universal Document Converter"
        RptImpresionFacturaVentaFE2.Printer.Copies = 1
        RptImpresionFacturaVentaFE2.Printer.FromPage = 1
        RptImpresionFacturaVentaFE2.Printer.ToPage = 1
        
        Dim feCODFACTURA As String
        feCODFACTURA = sSinNull(obtenerDeSQL("SELECT CODFACTURA FROM DOCUMENTOSCAE WHERE TIPO=" & ssTexto(tdoc) & " AND PUNTOVENTA=" & ssTexto(RptImpresionFacturaVentaFE2.lblPuntoVenta)))
        RptImpresionFacturaVentaFE2.lblCODFACTURA.caption = "COD. " & feCODFACTURA
        RptImpresionFacturaVentaFE2.documentName = "FacturaElectronicaORIGINAL_" & tdoc & "" & RptImpresionFacturaVentaFE2.lblComprobanteNro & ".fe"
        
        
        Dim Letra1 As String, Letra2 As String, Letra3 As String
        Dim NombreDocumento As String
        Letra1 = CORTO(tdoc, 0, 2)
        Letra2 = CORTO(tdoc, 1, 1)
        Letra3 = CORTO(tdoc, 2, 0)
        
        If Letra3 = "A" Then
            RptImpresionFacturaVentaFE2.lblLETRA = "A"
        ElseIf Letra3 = "B" Then
            RptImpresionFacturaVentaFE2.lblLETRA = "B"
        ElseIf Letra3 = "E" Then
            RptImpresionFacturaVentaFE2.lblLETRA = "E"
        End If
        
        If Letra1 = "F" Then
            If Letra2 = "A" Then
                If Letra3 = "A" Then
                    RptImpresionFacturaVentaFE2.lblTIPOFACTURA = "FACTURA" '"FACTURA 'A' "
                ElseIf Letra3 = "B" Then
                    RptImpresionFacturaVentaFE2.lblTIPOFACTURA = "FACTURA" '"FACTURA 'B' "
                ElseIf Letra3 = "E" Then
                    RptImpresionFacturaVentaFE2.lblTIPOFACTURA = "FACTURA DE EXPORTACIÓN"
                End If
            ElseIf Letra2 = "E" Then
                If Letra3 = "A" Then
                    RptImpresionFacturaVentaFE2.lblTIPOFACTURA = "FACTURA DE CREDITO" '"FACTURA 'A' "
                ElseIf Letra3 = "B" Then
                    RptImpresionFacturaVentaFE2.lblTIPOFACTURA = "FACTURA DE CREDITO" '"FACTURA 'B' "
                ElseIf Letra3 = "C" Then
                    RptImpresionFacturaVentaFE2.lblTIPOFACTURA = "FACTURA DE CREDITO"
                End If
            End If
        ElseIf Letra1 = "C" Then
            If Letra3 = "A" Then
                RptImpresionFacturaVentaFE2.lblTIPOFACTURA = "CREDITO ELECTRONICO"
            ElseIf Letra3 = "B" Then
                RptImpresionFacturaVentaFE2.lblTIPOFACTURA = "CREDITO ELECTRONICO"
            ElseIf Letra3 = "C" Then
                RptImpresionFacturaVentaFE2.lblTIPOFACTURA = "CREDITO ELECTRONICO"
            End If
        ElseIf Letra1 = "D" Then
            If Letra3 = "A" Then
                RptImpresionFacturaVentaFE2.lblTIPOFACTURA = "DEBITO ELECTRONICO"
            ElseIf Letra3 = "B" Then
                RptImpresionFacturaVentaFE2.lblTIPOFACTURA = "DEBITO ELECTRONICO"
            ElseIf Letra3 = "C" Then
                RptImpresionFacturaVentaFE2.lblTIPOFACTURA = "DEBITO ELECTRONICO"
            End If
        ElseIf Letra1 = "N" Then
            RptImpresionFacturaVentaFE2.lblRemito = ""
            If Letra2 = "C" Then
                If Letra3 = "A" Then
                    RptImpresionFacturaVentaFE2.lblTIPOFACTURA = "NOTA DE CREDITO" '"NOTA DE CREDITO 'A' "
                ElseIf Letra3 = "B" Then
                    RptImpresionFacturaVentaFE2.lblTIPOFACTURA = "NOTA DE CREDITO" '"NOTA DE CREDITO 'B' "
                ElseIf Letra3 = "E" Then
                    RptImpresionFacturaVentaFE2.lblTIPOFACTURA = "NOTA DE CREDITO DE EXP."
                End If
            ElseIf Letra2 = "D" Then
                If Letra3 = "A" Then
                    RptImpresionFacturaVentaFE2.lblTIPOFACTURA = "NOTA DE DEBITO" '"NOTA DE DEBITO 'A' "
                ElseIf Letra3 = "B" Then
                    RptImpresionFacturaVentaFE2.lblTIPOFACTURA = "NOTA DE DEBITO" '"NOTA DE DEBITO 'B' "
                ElseIf Letra3 = "E" Then
                    RptImpresionFacturaVentaFE2.lblTIPOFACTURA = "NOTA DE DEBITO DE EXP."
                End If
            End If
        End If
        If rs!nc_xdevolucion = True Then
            If s2n(nSinNull(rs!codfactura)) > 0 Then
                RptImpresionFacturaVentaFE2.Label29.caption = "Factura :"
                'RptImpresionFacturaVentaFE2.lblRemito.caption = sSinNull(rs!documve) & " - " & sSinNull(rs!origenve)
                RptImpresionFacturaVentaFE2.lblRemito.caption = sSinNull(obtenerDeSQL("select  (tipodoc  + '-' + cast(nrofactura as varchar) ) as ref   from facturaventa where codigo=" & rs!codfactura))
            End If
        End If
        
        'RptImpresionFacturaVentaFE.documentName = "FacturaElectronicaORIGINAL_" & RptImpresionFacturaVentaFE.lblComprobanteNro & ".fe"
        'RptImpresionFacturaVentaFE2.DataControl1.Connection = DataEnvironment1.Sistema
        'RptImpresionFacturaVentaFE2.DataControl1.Source = strConsulta2
        'RptImpresionFacturaVentaFE2.Printer.DeviceName = "Universal Document Converter"
        'RptImpresionFacturaVentaFE2.Printer.Copies = 1
        'RptImpresionFacturaVentaFE2.documentName = "FacturaElectronicaDUPLICADO_" & RptImpresionFacturaVentaFE.lblComprobanteNro & ".fe"
        'RptImpresionFacturaVentaFE2.lblORIGINAL = "DUPLICADO"

        RptImpresionFacturaVentaFE2.codigoFactura = codigo
               
        RptImpresionFacturaVentaFE2.Show
        RptImpresionFacturaVentaFE2.PrintReport False
        RptImpresionFacturaVentaFE2.Printer.DeviceName = Impresora
'        If Letra1 = "N" Then
'        Else
'            If MsgBox("Desea obtener tambien el remito?", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
'                ImprimirRemitoFE2 codigo
'                Unload RptImpresionRemitoFE2
'            End If
'        End If
        'Dim rptCopia As New RptImpresionFacturaVentaFE
        'rptCopies.pages.InsertNew 0
        'rptCopies.pages.InsertNew 1
        'rptCopies.Canvas = RptImpresionFacturaVentaFE
        'rptCopies.Show
    End If
    
    
Set rs = Nothing
Exit Function
ErrImpresion:
MsgBox "Error de impresión en factura electronica", vbCritical
End Function 'FIN IMPRIMIRCOMPROBANTEFE




Private Function ArmarCodeBar(ProdClie As String, letra As String, cant As Long, CodProvTK As String, NroRemito As Long)
    Dim sProd As String, sLetr As String, sCant As String, sProv As String, sRemi As String
    
    sProd = Left(ProdClie & Space(20), 10)
    sLetr = LCase(Left(letra & " ", 1))
    sCant = Format(cant, "000000")
    sProv = Right("       " & CodProvTK, 6)
    sRemi = Format(NroRemito, "00000000")
    
    ArmarCodeBar = sProd & sLetr & sCant & sProv & sRemi
End Function

Private Function pppp(Y, x, txt)
Dim ss
'    ss = aaFormFV.gri.TextMatrix(Y, 0)
    ss = mpppp(Y)
    ss = Left(ss, x) & txt & Right(ss, Len(ss) - Len(txt) - x)
'    aaFormFV.gri.TextMatrix(Y, 0) = ss
    mpppp(Y) = ss
End Function
 
Private Function formatTOT(n)
    formatTOT = alinear13(Format(n, "0.00")) '"##########0.00")
End Function
Private Function formatTOT4(n)
    formatTOT4 = alinear13(Format(n, "0.0000")) '"##########0.00")
End Function

Public Function formatP100(n)
    formatP100 = Format(n, "##.##")
End Function
Private Function alinear13(s)
    alinear13 = Space(13 - Len(s)) & s
End Function

Private Function coord(cual As ImprFactura)
    'OJO mimsmo orden que enum
    Dim x
    x = Array( _
       "Margen_Y", "Margen_X", _
       "tachar_Y", "tachar_X", _
       "Nro_Y", "Nro_X", _
       "labelFactura_Y", "labelFactura_X", _
       "fecha_Y", "fecha_X", _
       "Encabezado_Y", "Encabezado_X", _
       "CondIVA_Y", "CondIVA_X", _
       "FP_Y", "FP_X", _
       "VaCon_Y", "VaCon_X", _
       "Detalle_Y", _
       "Detalle_X_Cant", "Detalle_X_Prod", "Detalle_X_Desc", "Detalle_X_PreU", _
       "Pie_Y", _
       "Totales_X", _
       "Porc_X", _
       "EnLetras_Y", "EnLetras_X")
    coord = s2n(obtenerDeSQL("select " & x(cual) & " from impr_Factura ")) ' 1er y unico registro
End Function

Public Function ImprimirFacturaRemito(codigo, Optional conleyenda As Boolean = False) As Boolean
    ' OJO esta muy para tonka, no se puede considerar generico para matricial
    Dim rsF As New ADODB.Recordset, rsD As New ADODB.Recordset, i As Long
    Dim sf As String, sd As String
    Dim ss As String

'    Dim strPedidos As String, aPedidos
'    Dim ii As Long, ss As String
    
    
    Dim z As Double, PORCENTAJE As Double, mone As String, ProdClie As String, tdoc As String
    Dim t_sub, t_neto, t_coef
    Dim NroRem As Long
    Dim NroFac As Long
    
    Dim pieY As Long, totX As Long, cabY As Long, detY As Long, cabX As Long, prcX As Long, o As Long, leyendaY As Long
    Dim xCant As Long, xProd As Long, xDesc As Long, xPuni As Long
    
'    aPedidos = Array()
    
    totX = coord(Totales_x)
    pieY = coord(Pie_Y)
    cabY = coord(Encabezado_Y)
    cabX = coord(Encabezado_X)
    detY = coord(Detalle_Y)
    prcX = coord(porc_x)
    xCant = coord(Detalle_X_Cant)
    xProd = coord(Detalle_X_Prod)
    xDesc = coord(Detalle_X_Desc)
    xPuni = coord(Detalle_X_PreU)
    leyendaY = pieY - 21   ' ****  21 sobre pie PORQUE TIENE 20 renglones *****
    
    ReDim mpppp(77)
    For i = 0 To 76: mpppp(i) = Space(150): Next i
    
    sf = "SELECT FacturaVenta.*, Clientes.*, FormasPago.descripcion as pago, " & _
            " Ivas.descripcion as ivades, Ivas.letra as letra, clientes.iva as cliva, FacturaVenta.iva as facIva " & _
            " FROM FormasPago INNER JOIN (Ivas INNER JOIN (Clientes INNER JOIN FacturaVenta  " & _
            " ON Clientes.codigo = FacturaVenta.Cliente) " & _
            " ON Ivas.codigo = FacturaVenta.TipoIVA) " & _
            " ON FormasPago.codigo = FacturaVenta.FormaPago " & _
            " WHERE facturaventa.codigo=" & codigo
            
    'sd = "select fvd.*, p.descripcion  from FacturaVentaDetalle as fvd left join producto as p on  fvd.producto = p.codigo  where codigoFactura =  " & codigo
    sd = "select fvd.*  from FacturaVentaDetalle as fvd where codigoFactura =  " & codigo
            
    rsF.Open sf, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    rsD.Open sd, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    

    
'REMITO - FACTURA

    'REMITO head ---------------------------------------------
    With rsF
        NroFac = !NroFactura
        NroRem = !NroFactura - 2000
    
        pppp coord(fecha_y), coord(fecha_X), !Fecha
        pppp coord(Nro_Y), coord(Nro_x), "0001-" & Format(NroRem, "00000000")
        pppp coord(labelFactura_Y), coord(labelFactura_X), "REMITO"
        
        pppp cabY + 0, cabX, sSinNull(!RAZONSOCIAL)
        pppp cabY + 0, coord(fecha_X), "( " & !cliente & " )"
        pppp cabY + 1, cabX, ssStr(!direccion) & " - " & ssStr(!Localidad) & " - " & ssStr(!Provincia)
        pppp cabY + 2, cabX, ssStr(!CUIT)
        pppp cabY + 3, cabX, ssStr(!Proveedor)
        pppp coord(CondIVA_Y), coord(CondIVA_X), ssStr(!ivades)
        pppp coord(FP_Y), coord(FP_X), ssStr(!pago)
        pppp coord(VaCon_Y), coord(VaCon_X), "0001-" & Format(NroFac, "00000000")
    End With
    
    ' REMITO detalle ----------------------------------------
    With rsD
        o = 0
        While Not .EOF
            ProdClie = VerProdClie(sSinNull(!producto), rsF!cliente)
            pppp detY + o, xCant, SinZero(!cantidad)
            pppp detY + o, xProd, ProdClie
            
            If Trim(!producto) > "" Then ' TONKA esto para que en remito solo imprima solo descripciones de producto, no escritas por usuario
                pppp detY + o, xDesc, sSinNull(!DESCRIPCION)
            End If
            
            If !NroPedido > 0 Then
                ss = sSinNull(obtenerDeSQL("select pedido_cli from pedidos_clientes where numero = '" & !NroPedido & "' "))
                ss = IIf(ss = "", !NroPedido, ss)
                pppp detY + o, xPuni, ss
            End If
            
            .MoveNext
            o = o + 1
        Wend
    End With

    '  IMPRIMO SI NO ES EXTERIOR  -------- preguntar por la E
    If Trim(rsF!TIPODOC) <> "FAE" And Trim(rsF!TIPODOC) <> "NCE" Then
        ImprimirMpppp
    End If
    
    'FACTURA head -------------------------------------------------------
    With rsF
        tdoc = !TIPODOC
        z = !cotizacion: If z = 0 Then z = 1
        mone = ObtenerDescripcion("Monedas", !moneda)
        
        '  IMPRIMO SI NO ES EXTERIOR  -------- preguntar por la E
        If Trim(rsF!TIPODOC) <> "FAE" And Trim(rsF!TIPODOC) <> "NCE" Then
            pppp coord(Nro_Y), coord(Nro_x), "0001-" & Format(NroFac, "00000000")
'            pppp coord(VaCon_Y), coord(VaCon_X), "0001-" & Format(!remito, "00000000")
            '**********  OJO TRUCHADA PARA TONKA
            pppp coord(VaCon_Y), coord(VaCon_X), "0001-" & Format(NroRem, "00000000")
            '
        End If

        Select Case !TIPODOC
         Case "FAA" To "FZZ"
                pppp coord(labelFactura_Y), coord(labelFactura_X), "FACTURA"
         Case "NCA" To "NCZ"
                pppp coord(tachar_Y), coord(tachar_X), "XXXXXXXXXXXX"
                pppp coord(labelFactura_Y), coord(labelFactura_X), "NOTA DE CREDITO"
         Case "NDA" To "NDZ"
                pppp coord(tachar_Y), coord(tachar_X), "XXXXXXXXXXXX"
                pppp coord(labelFactura_Y), coord(labelFactura_X), "NOTA DE DEBITO"
        End Select
        
        If !IIBB <> 0 Then
            pppp pieY + 4, prcX, formatP100(s2n((1 / (!Neto / !IIBB) * 100)))
            pppp pieY + 4, totX, formatTOT(!IIBB)
        End If
        If !Descuento = 0 Then
            pppp pieY + 0, totX, formatTOT(!Neto / z)
        Else
            pppp pieY + 1, prcX, formatP100(s2n(!Descuento * 100))
            
            ' ACA HAY ALGO QUE NO ESTOY DE ACUERDO CON EL CONTADOR DE TONKA,
            ' QU ME DIJO QUE PONGA SUBTOTAL, INTERES y TOTAL en vez de SUBTOTAL, INTERES, NETO,  y TOTAL
                   
            t_neto = !Neto / z
            t_coef = Sgn(!Descuento) * 0 + !Descuento
            t_sub = t_neto / (1 - t_coef)
            
            pppp pieY + 0, totX, formatTOT(t_sub)
            pppp pieY + 1, totX, formatTOT(t_neto - t_sub)
        End If
        
        pppp pieY + 6, totX, formatTOT(!Total / z)
        
        Select Case Right(Trim(tdoc), 1)
        Case "A"
            pppp pieY + 5, coord(EnLetras_X), "Son Pesos" & enletras(Round(!Total / z))
            'iva
            pppp pieY + 3, prcX, formatP100(s2n(!PorcentajeIva) * 100)
'            pppp pieY + 3, totX, formatTOT(s2n(!neto) * s2n(!PORCENTAJEiva))
            pppp pieY + 3, totX, formatTOT(s2n(!faciva / z))
        Case "B" ' subtot, interes, dcto, (todo + iva)
            pppp pieY + 5, coord(EnLetras_X), "Son Pesos" & enletras(Round(!Total / z, 2))
        Case "E"
            pppp pieY + 5, coord(EnLetras_X), "Son " & mone & " " & enletras(!Total / z)
            'pppp coord(EnLetras_Y + 1), coord(EnLetras_X), "a  " & z   ' cotizacion
            ' subtot, dcto, interes, (todo / z)
            ' otros mensajes
        Case Else
            ufa "prg: Letra no reconocida", "FV impr: " & codigo
        End Select
    End With
    
    ' FACTURA detalle --------------------------------
    Dim descri As String
    Dim ume As String
    With rsD
        o = 0
        If Not .BOF Then .MoveFirst
        While Not .EOF
            ume = Left(sSinNull(obtenerDeSQL("select u.descripcion from unidadesmedida u inner join producto p on p.umedida = u.codigo where p.codigo = '" & !producto & "' ")), 2)
            ProdClie = VerProdClie(sSinNull(!producto), rsF!cliente)
            pppp detY + o, xCant, SinZero(!cantidad)
            
            pppp detY + o, xProd - 3, ume  ' mide 2 de prepo y le dejo 1 espacio
            pppp detY + o, xProd, ProdClie
            
            descri = Left(sSinNull(!DESCRIPCION), 33)
            descri = descri & Space(33 - Len(descri))
'            pppp detY + o, xDesc, cadenahexa("1b0f") & descri & cadenahexa("1220")
            
            If !PrecioTotal = 0 Then  ' no imprimo ceros
                    pppp detY + o, xDesc, cadenahexa("1b0f") & descri & cadenahexa("1220") & "  "
            Else
                Select Case Right(Trim(tdoc), 1)
                Case "E"  '  FAE  NCE  exportacion
                    pppp detY + o, xDesc, cadenahexa("1b0f") & descri & cadenahexa("1220") _
                        & "  " & formatTOT4(!PrecioUnitario / z) _
                        & "  " & formatTOT(!PrecioTotal / z)
                Case "B"  ' FAB detalle con iva incluido
                    pppp detY + o, xDesc, cadenahexa("1b0f") & descri & cadenahexa("1220") _
                        & "  " & formatTOT4(!PrecioUnitario * (1 + rsF!PorcentajeIva)) _
                        & "  " & formatTOT(!PrecioTotal * (1 + rsF!PorcentajeIva))
                Case "A"
                    pppp detY + o, xDesc, cadenahexa("1b0f") & descri & cadenahexa("1220") _
                        & "  " & formatTOT4(!PrecioUnitario) _
                        & "  " & formatTOT(!PrecioTotal)
                Case Else
                    ufa "prg: no se reconoce letra " & tdoc, "impresion factura " & codigo
                End Select
            End If
            .MoveNext
            o = o + 1
        Wend
    End With
    
    
    Dim rsleye As New ADODB.Recordset
    If Trim(rsF!TIPODOC) = "FAE" And conleyenda Then
        rsleye.Open "select renglon from FacturaVentaLeyendaFAE order by id", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        For i = 1 To frmLeyendaFAE.Renglones()
            pppp leyendaY + i - 1, xCant, rsleye!RENGLON
            rsleye.MoveNext
        Next i
    End If
    
    Set rsF = Nothing
    Set rsD = Nothing
    
    ImprimirMpppp
End Function

Private Function SinZero(que)
    If ON_ERROR_HABILITADO Then On Error Resume Next
    SinZero = IIf(que = 0, "", que)
End Function

Public Function VerProdClie(prod As String, clie As Long)
    Dim s
    s = ssStr(obtenerDeSQL("select ProductoCliente From Relacion_Producto_Cliente where producto = '" & ssStr(prod, True) & "' and cliente = " & clie))
    VerProdClie = IIf(Trim(s) = "", prod, s)
End Function


Private Sub ImprimirMpppp()
    Dim Impresora As String
    Dim i As Long, lineas As Long, sMargen As String
    
    Impresora = Trim(obtenerParametro("ImpresoraFactura"))
    lineas = 75
    For i = UBound(mpppp) To 0 Step -1
        'mpppp(i) = Trim(mpppp(i))
        If Trim(mpppp(i)) = "" Then
            lineas = i
        Else
            Exit For
        End If
    Next i
    
    Open Impresora For Output As #1
                            ' 12 inch, draft, 10cpi, nocondensado
        'Print #1, cadenahexa("1b43000c" & "1b7800" & "1b50" & "12")
        
        ' prueba, no seteo pulgadas.
        '  draft 10 no concd
        Print #1, cadenahexa("1b7800" & "1b50" & "12")
        
                            ' 12 inch, draft,  nocondensado
'       Print #1, cadenahexa("1b43000c" & "1b7800" & "12")
                            
                            ' reset, 12 inch, draft, 12cpi, nocondensado
'       Print #1, cadenahexa("1b40" & "1b43000c" & "1b7800" & "1b4d" & "12")
                            
                            ' 12 inch, draft, 10cpi,
'       Print #1, cadenahexa("1b43000c" & "1b7800" & "1b50" & "12")
                            
                            ' reset, 12 inch, 12cpi, Condensado
'       Print #1, cadenahexa("1b40" & "1b43000c" & "1b4d" & "1b0f")
        
       'Print #1, "0-0"
               
        sMargen = Space(coord(Margen_X))
        For i = 1 To coord(Margen_Y)
            Print #1, Chr(&HA)
        Next i
        For i = 0 To lineas - 1
            'Debug.Print i & " * " & Len(RTrim(mpppp(i))) & " *** " & RTrim(mpppp(i))
            Print #1, sMargen & RTrim(mpppp(i))
        Next i
        Print #1, Chr(&HC);
    Close #1
End Sub

Public Function cadenahexa(x As String) As String
    Dim i As Long, s As String, C As String
    For i = 1 To Len(x) Step 2
        s = Mid$(x, i, 2)
        C = C & Chr(CByte("&H" & s))
    Next i
    cadenahexa = C
End Function

Public Function array2chr(aa) As String
    Dim i As Long, C As String
    For i = 0 To UBound(aa)
        C = C & Chr(aa(i))
    Next i
End Function

Public Function ver_impresoras()
    Dim x As Printer
    For Each x In Printers
        MsgBox x.DeviceName, vbExclamation, "Informacion"
    Next
End Function


'Private Sub agregarStrPedido(arre, ByVal cual)
'    Dim i As Long
'    Dim u As Long
'    u = UBound(arre)
'
'    If Trim(cual) = "" Then Exit Sub
'
'    If u = -1 Then
'        ReDim Preserve arre(0)
'        arre(0) = cual
'        'arre = Array(cual)
'    Else
'        For i = 0 To UBound(arre)
'            If arre(i) = cual Then
'                Exit Sub
'            End If
'        Next i
'
'        ReDim Preserve arre(u + 1)
'        arre(u + 1) = cual
'
'    End If
'End Sub
Public Function ImprimirRemitoVentaAT(codigo, codRemito) As Boolean

    Dim rs As New ADODB.Recordset
    
    Dim rs2 As New ADODB.Recordset
    Dim rs3 As New ADODB.Recordset
    Dim rs4 As New ADODB.Recordset
    Dim CantBulto, AuxNPed As Double
    
    Dim str, str2, NPedidos As String
    Dim COD As Long
     
    rs.Open "SELECT RemitoVenta.*,remitoventa.transporte as trans, Clientes.*, Ivas.descripcion as iva" _
    & " FROM Ivas INNER JOIN (Clientes INNER JOIN RemitoVenta ON Clientes.codigo = RemitoVenta.Cliente) ON Ivas.codigo = clientes.IVA  where numero=" & codigo & " and remitoventa.codigo=" & codRemito, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    If Not rs.EOF Then
        rs2.Open "SELECT direccion,descripcion FROM Transportes WHERE codigo = " & rs!trans & "", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
        If Not rs2.EOF Then
        If rs2!direccion = "" Then
         RptImpresionRemitoVentaAT.DireccTrans = " "
        Else
            If Not IsNull(rs2!direccion) Then
               RptImpresionRemitoVentaAT.DireccTrans = rs2!direccion
            Else: RptImpresionRemitoVentaAT.DireccTrans = "-"
            End If
        End If
    '    PONER EN LUGAR DEL RECORDSET EL VALOR DEL FORM
    '**************************************************************************
         RptImpresionRemitoVentaAT.CodTrans = rs!trans
         RptImpresionRemitoVentaAT.DescTrans = rs2!DESCRIPCION
        
         RptImpresionRemitoVentaAT.Valor = frmRemitoVenta.lblTotalRV
        If IsNull(rs!obs1) = True Or rs!obs1 = "" Then
         RptImpresionRemitoVentaAT.Obser1 = "-"
        Else
         RptImpresionRemitoVentaAT.Obser1 = rs!obs1
        End If
        If IsNull(rs!obs2) = True Or rs!obs2 = "" Then
         RptImpresionRemitoVentaAT.Obser2 = "-"
        Else
         RptImpresionRemitoVentaAT.Obser2 = rs!obs2
        End If
         
        End If
        rs2.Close
        
        RptImpresionRemitoVentaAT.lblcliente = sSinNull(rs!nombrefantasia)
        If Not IsNull(rs!CUIT) Then
            RptImpresionRemitoVentaAT.lblCuit = rs!CUIT
        End If
        COD = rs!numero
        RptImpresionRemitoVentaAT.lblcomp = "Remito"
        
        str = "select distinct RemitoVentaDetalle.pedido,pedidos_clientes.pedido_cli from RemitoVentaDetalle inner join pedidos_clientes on RemitoVentaDetalle.pedido=pedidos_clientes.numero  where RemitoVentaDetalle.numero=" & COD & " and remitoventadetalle.codremito=" & codRemito
        rs2.Open str, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
        If Not IsNull(rs2!Pedido) Then
            If Not IsEmpty(rs2!Pedido) Then
                If rs2.BOF = True And rs2.EOF = True Then
                    RptImpresionRemitoVentaAT.Pedido = ""
                Else
                    RptImpresionRemitoVentaAT.Pedido = rs2!Pedido
                End If
            Else
                RptImpresionRemitoVentaAT.Pedido = ""
            End If
        End If
        If Not IsNull(rs2!pedido_cli) Then
            If Not IsEmpty(rs2!pedido_cli) Then
                If rs2.EOF = True And rs2.BOF = True Then
                    RptImpresionRemitoVentaAT.Compra = ""
                Else
                    RptImpresionRemitoVentaAT.Compra = rs2!pedido_cli
                End If
            Else
                RptImpresionRemitoVentaAT.Compra = ""
            End If
        End If
        
        rs2.Close
        
        str = "SELECT RemitoVentaDetalle.*, Producto.descripcion" _
        & " FROM RemitoVentaDetalle INNER JOIN Producto ON RemitoVentaDetalle.Producto = Producto.codigo" _
        & " where numero=" & COD & " and codremito=" & codRemito & " ORDER BY remitoventadetalle.CODIGO"
        
        rs2.Open str, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
        If Not rs2.EOF Then
            RptImpresionRemitoVentaAT.PedPropio = rs2!Pedido
            rs4.Open "SELECT pedido_cli FROM pedidos_clientes WHERE numero = " & rs2!Pedido & "", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
            If Not rs4.EOF Then
                RptImpresionRemitoVentaAT.PedCli = rs4!pedido_cli
            Else: RptImpresionRemitoVentaAT.PedCli = 0
            End If
        End If
'        Do While Not rs2.EOF
'            rs3.Open "SELECT formula FROM producto WHERE codigo = '" & Trim(rs2!producto) & "'", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
'             If Not rs3.EOF And rs3!formula <> True Then
'                CantBulto = CantBulto + 1
'             End If
'             rs3.Close
'            rs2.MoveNext
'        Loop
        rs2.Close
        Dim ssql As String
        
'        sSql = "SELECT SUM(RemitoVentaDetalle.Cantidad) AS SumaCant FROM RemitoVentaDetalle " & _
'                  "INNER JOIN Producto ON RemitoVentaDetalle.Producto = Producto.codigo Where " & _
'                  "(producto.formula = 0) And (RemitoVentaDetalle.numero = 6348) GROUP BY " & _
'                  "RemitoVentaDetalle.Cantidad, RemitoVentaDetalle.Numero ORDER BY " & _
'                  "RemitoVentaDetalle.Numero"
'        ssql = "SELECT SUM(RemitoVentaDetalle.Cantidad) AS SumaCant FROM RemitoVentaDetalle " & _
                  "INNER JOIN Producto ON RemitoVentaDetalle.Producto = Producto.codigo Where " & _
                  "(producto.formula = 0) And (RemitoVentaDetalle.numero = '" & codigo & "') GROUP BY " & _
                  "RemitoVentaDetalle.Cantidad, RemitoVentaDetalle.Numero ORDER BY " & _
                  "RemitoVentaDetalle.Numero" 'ver si falta filtrar si esta cancelado!!
        ssql = "SELECT SUM(RemitoVentaDetalle.Cantidad) AS SumaCant FROM RemitoVentaDetalle " & _
                  "INNER JOIN Producto ON RemitoVentaDetalle.Producto = Producto.codigo Where " & _
                  "(producto.formula = 0) And (RemitoVentaDetalle.numero = '" & codigo & "' and codremito=" & codRemito & ") " 'ver si falta filtrar si esta cancelado!!


        rs3.Open ssql, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
        If Not IsNull(rs3!sumacant) Then
            If Not IsEmpty(rs3!sumacant) Then
                If Not (rs3.BOF = True And rs3.EOF = True) Then
                    RptImpresionRemitoVentaAT.TotBultos = rs3!sumacant
                Else
                    RptImpresionRemitoVentaAT.TotBultos = 0
                End If
            End If
        End If
        If frmRemitoVenta.txtPedido <> "" Then
           AuxNPed = frmRemitoVenta.txtPedido
        Else
           
        End If
        'If frmRemitoVenta.entrega <> "" Then
        If gEMPR_idEmpresa <> 6 Then
            RptImpresionRemitoVentaAT.Entrega.Text = sSinNull(rs!Entrega) 'frmRemitoVenta.entrega
        End If
        'End If
        'rs2.Open "SELECT pedido_cli FROM pedidos_clientes WHERE numero = " & AuxNPed & ""
        'If Not rs2.EOF Then
  '
  '           RptImpresionRemitoVenta.PedPropio = AuxNPed
  '       End If
  '       rs2.Close
        
        RptImpresionRemitoVentaAT.lblcliente = rs!DESCRIPCION
        RptImpresionRemitoVentaAT.lbldomicilio = rs!direccion
        RptImpresionRemitoVentaAT.lblfactura = Format(rs!PuntoVenta, "0000") & "-" & Format(rs!numero, "00000000") '"0002-" & Format(rs!numero, "00000000")
        RptImpresionRemitoVentaAT.lblfactura.Visible = True
        RptImpresionRemitoVentaAT.lblfecha = rs!Fecha
        RptImpresionRemitoVentaAT.LblIVA = rs!Iva
        'RptImpresionRemitoVenta.lbllocalidad = rs!localidad
        If Not IsNull(rs!Localidad) Then
            If Not IsEmpty(rs!Localidad) Then
                RptImpresionRemitoVentaAT.lbllocalidad = rs!Localidad
            Else
                RptImpresionRemitoVentaAT.lbllocalidad = rs!Localidad
            End If
        End If
         
        LlenarTemp2 (str)
        RptImpresionRemitoVentaAT.DataControl1.Connection = DataEnvironment1.Sistema
        
        ' CAMBIAR PARA QUE EL STR QUE FIGURA SE REDIRECCIONE A LA NUEVA TABLA
        ' QUE CREE
        str = "SELECT * FROM " & sTablaRemito & " "
        RptImpresionRemitoVentaAT.DataControl1.Source = str
        
    End If
    rs.Close
    Set rs = Nothing
    
    RptImpresionRemitoVentaAT.Printer.Copies = 3
    
'    RptImpresionRemitoVentaAT.PageSettings.TopMargin = margenTop_RV()
    
'    RptImpresionRemitoVenta.PrintReport True
    RptImpresionRemitoVentaAT.Show vbModal
End Function

Public Function ImprimirRemitoVentaAT2(codigo, codRemito) As Boolean

    Dim rs As New ADODB.Recordset
    
    Dim rs2 As New ADODB.Recordset
    Dim rs3 As New ADODB.Recordset
    Dim rs4 As New ADODB.Recordset
    Dim CantBulto, AuxNPed As Double
    
    Dim str, str2, NPedidos As String
    Dim COD As Long
     
    rs.Open "SELECT RemitoVenta.*,remitoventa.transporte as trans " _
    & " FROM RemitoVenta where numero=" & codigo & " and remitoventa.codigo=" & codRemito, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    If Not rs.EOF Then
        rs2.Open "SELECT direccion,descripcion FROM Transportes WHERE codigo = " & rs!trans & "", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
        If Not rs2.EOF Then
        If rs2!direccion = "" Then
         RptImpresionRemitoVentaAT.DireccTrans = " "
        Else
            If Not IsNull(rs2!direccion) Then
               RptImpresionRemitoVentaAT.DireccTrans = rs2!direccion
            Else: RptImpresionRemitoVentaAT.DireccTrans = "-"
            End If
        End If
    '    PONER EN LUGAR DEL RECORDSET EL VALOR DEL FORM
    '**************************************************************************
         RptImpresionRemitoVentaAT.CodTrans = rs!trans
         RptImpresionRemitoVentaAT.DescTrans = rs2!DESCRIPCION
        
         RptImpresionRemitoVentaAT.Valor = frmRemitoVenta.lblTotalRV
        If IsNull(rs!obs1) = True Or rs!obs1 = "" Then
         RptImpresionRemitoVentaAT.Obser1 = "-"
        Else
         RptImpresionRemitoVentaAT.Obser1 = rs!obs1
        End If
        If IsNull(rs!obs2) = True Or rs!obs2 = "" Then
         RptImpresionRemitoVentaAT.Obser2 = "-"
        Else
         RptImpresionRemitoVentaAT.Obser2 = rs!obs2
        End If
         
        End If
        rs2.Close
        
        RptImpresionRemitoVentaAT.lblcliente = sSinNull(rs!descriclie)
        RptImpresionRemitoVentaAT.lblCuit = "" 'rs!CUIT
        
        COD = rs!numero
        RptImpresionRemitoVentaAT.lblcomp = "Remito"
        
        str = "select distinct RemitoVentaDetalle.pedido,pedidos_clientes.pedido_cli from RemitoVentaDetalle inner join pedidos_clientes on RemitoVentaDetalle.pedido=pedidos_clientes.numero  where RemitoVentaDetalle.numero=" & COD & " and remitoventadetalle.codremito=" & codRemito
        rs2.Open str, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
        If Not IsNull(rs2!Pedido) Then
            If Not IsEmpty(rs2!Pedido) Then
                If rs2.BOF = True And rs2.EOF = True Then
                    RptImpresionRemitoVentaAT.Pedido = ""
                Else
                    RptImpresionRemitoVentaAT.Pedido = rs2!Pedido
                End If
            Else
                RptImpresionRemitoVentaAT.Pedido = ""
            End If
        End If
        If Not IsNull(rs2!pedido_cli) Then
            If Not IsEmpty(rs2!pedido_cli) Then
                If rs2.EOF = True And rs2.BOF = True Then
                    RptImpresionRemitoVentaAT.Compra = ""
                Else
                    RptImpresionRemitoVentaAT.Compra = rs2!pedido_cli
                End If
            Else
                RptImpresionRemitoVentaAT.Compra = ""
            End If
        End If
        
        rs2.Close
        
        str = "SELECT RemitoVentaDetalle.*, Producto.descripcion" _
        & " FROM RemitoVentaDetalle INNER JOIN Producto ON RemitoVentaDetalle.Producto = Producto.codigo" _
        & " where numero=" & COD & " and codremito=" & codRemito & " ORDER BY remitoventadetalle.CODIGO"
        
        rs2.Open str, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
        If Not rs2.EOF Then
            RptImpresionRemitoVentaAT.PedPropio = rs2!Pedido
            rs4.Open "SELECT pedido_cli FROM pedidos_clientes WHERE numero = " & rs2!Pedido & "", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
            If Not rs4.EOF Then
                RptImpresionRemitoVentaAT.PedCli = rs4!pedido_cli
            Else: RptImpresionRemitoVentaAT.PedCli = 0
            End If
        End If
        
        rs2.Close
        Dim ssql As String
        
        ssql = "SELECT SUM(RemitoVentaDetalle.Cantidad) AS SumaCant FROM RemitoVentaDetalle " & _
                  "INNER JOIN Producto ON RemitoVentaDetalle.Producto = Producto.codigo Where " & _
                  "(producto.formula = 0) And (RemitoVentaDetalle.numero = '" & codigo & "' and codremito=" & codRemito & ") " 'ver si falta filtrar si esta cancelado!!

        rs3.Open ssql, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
        If Not IsNull(rs3!sumacant) Then
            If Not IsEmpty(rs3!sumacant) Then
                If Not (rs3.BOF = True And rs3.EOF = True) Then
                    RptImpresionRemitoVentaAT.TotBultos = rs3!sumacant
                Else
                    RptImpresionRemitoVentaAT.TotBultos = 0
                End If
            End If
        End If
        If frmRemitoVenta.txtPedido <> "" Then
           AuxNPed = frmRemitoVenta.txtPedido
        Else
           
        End If
        If gEMPR_idEmpresa <> 6 Then
            RptImpresionRemitoVentaAT.Entrega.Text = sSinNull(rs!Entrega) 'frmRemitoVenta.entrega
        End If
        
        RptImpresionRemitoVentaAT.lblcliente = rs!descriclie
        RptImpresionRemitoVentaAT.lbldomicilio = "" 'rs!direccion
        RptImpresionRemitoVentaAT.lblfactura = Format(rs!PuntoVenta, "0000") & "-" & Format(rs!numero, "00000000") '"0002-" & Format(rs!numero, "00000000")
        RptImpresionRemitoVentaAT.lblfactura.Visible = True
        RptImpresionRemitoVentaAT.lblfecha = rs!Fecha
        RptImpresionRemitoVentaAT.LblIVA = "" 'rs!Iva
        
        RptImpresionRemitoVentaAT.lbllocalidad = "" 'rs!Localidad
         
        LlenarTemp2 (str)
        RptImpresionRemitoVentaAT.DataControl1.Connection = DataEnvironment1.Sistema
        
        ' CAMBIAR PARA QUE EL STR QUE FIGURA SE REDIRECCIONE A LA NUEVA TABLA
        ' QUE CREE
        str = "SELECT * FROM " & sTablaRemito & " "
        RptImpresionRemitoVentaAT.DataControl1.Source = str
        
    End If
    rs.Close
    Set rs = Nothing
    
    RptImpresionRemitoVentaAT.Printer.Copies = 3
    
    RptImpresionRemitoVentaAT.Show vbModal
End Function


Private Sub LlenarTemp2(str As String)
   Dim rs As New ADODB.Recordset
   Dim rs2 As New ADODB.Recordset
   Dim AuxDescrip As String, AuxCant As String
  
  sTablaRemito = TablaTempCrear("([id] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ,   [cantidad] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,   [codigo] [varchar] (1000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,   [descrip] [varchar] (800) COLLATE SQL_Latin1_General_CP1_CI_AS NULL) ON [PRIMARY]")
   
   rs.Open str, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    Do While Not rs.EOF
        
        AuxCant = x2s(rs!cantidad)
        AuxDescrip = rs!producto & " " & rs!DESCRIPCION
        DataEnvironment1.Sistema.Execute "INSERT INTO " & sTablaRemito & " (cantidad,codigo) VALUES( '" & AuxCant & "','" & AuxDescrip & "')"
        rs2.Open "SELECT serie FROM series WHERE producto = '" & rs!producto & "' and nrocomprobante = '" & rs!numero & "'", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
        If Not rs2.EOF Then
            
            AuxCant = "Serie/s : "
            AuxDescrip = ""
            Do While Not rs2.EOF
                If Trim(AuxDescrip) <> "" Then
                    AuxDescrip = AuxDescrip & "; " & rs2!Serie
                Else
                    AuxDescrip = AuxCant & rs2!Serie
                    AuxCant = ""
                End If
'      Else
'         AuxCodigo = " "
'         AuxCant = " "
'         AuxDescrip = " "
'         DataEnvironment1.Sistema.Execute "INSERT INTO " & sTablaRemito & "  (cantidad,codigo,descrip) VALUES('" & AuxCant & "','" & AuxCodigo & "','" & AuxDescrip & "')"
                rs2.MoveNext
            Loop
            DataEnvironment1.Sistema.Execute "INSERT INTO " & sTablaRemito & " (cantidad,codigo) VALUES( '" & s2n(AuxCant) & "','" & AuxDescrip & "')"
            Dim a
            a = Len("Serie/s : K002; K003; K004; K005; K006; K007; K008; K009; K010; K011; K012; K013; K014; K015; K016; K017; K018; K019; K020; K021; K022; K023; K024; K025; K026; K027; K028; K029; K030; K031; K032; K033; K034; K035; K036; K037; K038; K039; K040; K041; K042; K043; K044; K045; K046; K047; K048; K049; K050; K051; K052; K053; K054; K055; K056; K057; K058; K059; K060; K061; K062; K063; K064; K065; K066; K067; K068; K069; K070; K071; K072; K073; K074; K075; K076; K077; K078; K079; K080; K081; K082; K083; K084; K085; K086; K087; K088; K089; K090; K091; K092; K093; K094; K095; K096; K097; K098; K099; K100; K101; K102; K103; K104; K105; K106; K107; K108; K109; K110; K111; K112; K113; K114; K115; K116; K117; K118; K119; K120; K121; K122; K123; K124")
        End If
        
        rs2.Close
        rs.MoveNext
   Loop
   rs.Close
   Set rs = Nothing
End Sub

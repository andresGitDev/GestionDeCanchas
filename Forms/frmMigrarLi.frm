VERSION 5.00
Begin VB.Form frmMigrarLi 
   Caption         =   "AJUSTES  POS - MIGRACION"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   8175
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   615
      Left            =   5760
      TabIndex        =   27
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton cmdComprasActualiz200607 
      Caption         =   "ComprasActualiz200607"
      Height          =   630
      Left            =   5670
      TabIndex        =   26
      Top             =   2130
      Width           =   2025
   End
   Begin VB.CommandButton cmdMaeART 
      Caption         =   "mae_art modelo (solo nimisan)"
      Height          =   465
      Left            =   5670
      TabIndex        =   25
      Top             =   1500
      Width           =   1875
   End
   Begin VB.CommandButton Command3 
      Caption         =   "iibb prov de dbf"
      Height          =   480
      Left            =   5625
      TabIndex        =   24
      Top             =   885
      Width           =   1815
   End
   Begin VB.CommandButton cmdRemigraMovicaja 
      Caption         =   "movicaja dbf contado"
      Height          =   495
      Left            =   5640
      TabIndex        =   23
      Top             =   210
      Width           =   1830
   End
   Begin VB.CommandButton cmdSacarComitaProducto 
      Caption         =   "Sacar Comita Producto"
      Height          =   450
      Left            =   3540
      TabIndex        =   22
      Top             =   4875
      Width           =   1770
   End
   Begin VB.CommandButton cmd_idDoc_FV 
      Caption         =   "idDoc para FV"
      Height          =   495
      Left            =   3765
      TabIndex        =   21
      Top             =   1905
      Width           =   1440
   End
   Begin VB.CommandButton cmdCuitCliente 
      Caption         =   "cuit cliente"
      Height          =   525
      Left            =   3660
      TabIndex        =   20
      Top             =   4230
      Width           =   1635
   End
   Begin VB.CommandButton cmdOC_Moneda 
      Caption         =   "OC moneda "
      Height          =   435
      Left            =   3645
      TabIndex        =   19
      Top             =   3660
      Width           =   1635
   End
   Begin VB.CommandButton cmdRetenc 
      Caption         =   "mig ret OJO 1 VEZ"
      Height          =   450
      Left            =   3735
      TabIndex        =   18
      Top             =   2550
      Width           =   1545
   End
   Begin VB.CommandButton cmdNCT 
      Caption         =   "NDR a NDA rech"
      Height          =   525
      Left            =   3660
      TabIndex        =   17
      Top             =   3060
      Width           =   1635
   End
   Begin VB.CommandButton cmdFCContado 
      Caption         =   "FC Contado"
      Height          =   390
      Left            =   3720
      TabIndex        =   16
      Top             =   1410
      Width           =   1530
   End
   Begin VB.CommandButton cmdFC_rzSoc 
      Caption         =   "FC: Razon y Cu"
      Height          =   465
      Left            =   3765
      TabIndex        =   15
      Top             =   810
      Width           =   1485
   End
   Begin VB.CommandButton cmdTipoIvaTC 
      Caption         =   "TipoIva TC"
      Height          =   375
      Left            =   3690
      TabIndex        =   14
      Top             =   225
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "tipoiva null"
      Height          =   555
      Left            =   1860
      TabIndex        =   13
      Top             =   3735
      Width           =   1515
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Series Fecha_alta"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1770
      TabIndex        =   12
      Top             =   3135
      Width           =   1635
   End
   Begin VB.CommandButton remitoconfactura 
      Caption         =   "remito con factura"
      Height          =   435
      Left            =   1800
      TabIndex        =   11
      Top             =   2595
      Width           =   1635
   End
   Begin VB.CommandButton formula 
      Caption         =   "producto.formula"
      Height          =   435
      Left            =   1800
      TabIndex        =   10
      Top             =   2055
      Width           =   1635
   End
   Begin VB.CommandButton cmdBS 
      Caption         =   "datos BS  OJO 1 SOLA VEZ"
      Height          =   675
      Left            =   1800
      TabIndex        =   9
      Top             =   1275
      Width           =   1635
   End
   Begin VB.CommandButton cmdFVmodStock 
      Caption         =   "FV modiStock"
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   735
      Width           =   1635
   End
   Begin VB.CommandButton cmdFV_Contado 
      Caption         =   "FV Contado"
      Height          =   360
      Left            =   240
      TabIndex        =   7
      Top             =   3915
      Width           =   1455
   End
   Begin VB.CommandButton cmdFV_cli_rzSoc_cuit 
      Caption         =   "FV Razon y cuit"
      Enabled         =   0   'False
      Height          =   420
      Left            =   240
      TabIndex        =   6
      Top             =   3315
      Width           =   1455
   End
   Begin VB.CommandButton cmdMigracionPrincipal 
      Caption         =   "MigracionPrincipal"
      Enabled         =   0   'False
      Height          =   525
      Left            =   225
      TabIndex        =   5
      Top             =   765
      Width           =   1515
   End
   Begin VB.CommandButton cmdImporteOC 
      Caption         =   "ImporteOC"
      Height          =   480
      Left            =   240
      TabIndex        =   4
      Top             =   2670
      Width           =   1455
   End
   Begin VB.CommandButton cmdResetNull 
      Caption         =   "cmdResetNull"
      Height          =   480
      Left            =   300
      TabIndex        =   3
      Top             =   120
      Width           =   1395
   End
   Begin VB.CommandButton cmdProdcodigo 
      Caption         =   "cambia codigo producto"
      Enabled         =   0   'False
      Height          =   525
      Left            =   195
      TabIndex        =   2
      Top             =   2055
      Width           =   1455
   End
   Begin VB.TextBox txttabula 
      Height          =   315
      Left            =   1860
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   195
      Width           =   1695
   End
   Begin VB.CommandButton cmdFV 
      Caption         =   "reset ACTIVO = 1"
      Height          =   420
      Left            =   240
      TabIndex        =   0
      Top             =   1455
      Width           =   1455
   End
End
Attribute VB_Name = "frmMigrarLi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'30/9/4
'14/10/4 agregadas tablas


' El paso de varias tablas lo hice desde access
'    "migraña.mdb"
' poner autonum a Series.Codigo


Option Explicit

Private Const setFUA = " fecha_alta, fecha_baja, Usuario_alta, usuario_baja, activo "
'Private Const valFUA = " "
Private Function valFUA()
    valFUA = CLng(Date) & ", 0, 100, 0, 1 "
End Function

Private Sub voy(donde)
    txttabula = donde
End Sub

'Private Sub cmd_idDoc_FV_Click()
'    Dim rs As New ADODB.Recordset, i As Long
'
'    With rs
'        .Open "select * from FacturaVenta", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
'        While Not .EOF
'            i = NuevoDocumento(!TIPODOC, !nroFactura, 0)
'            !iddoc = i
'            .Update
'            .MoveNext
'        Wend
'    End With
'    Set rs = Nothing
'End Sub

Private Sub cmdBS_Click()
    Dim temp, COD_FacturaVenta, NUM_FAC_A, FEC_FAC_A As Date, NUM_FAC_B, FEC_FAC_B As Date, OrdenPago, NUM_ImpPro, anticip_pr, Nro_RemitoVenta As Long
    Dim nume_apc, nume_apd
    
    temp = obtenerDeSQL("select top 1 NroFactura, fecha from FacturaVenta where activo = 1 and TipoDoc = 'FAA' order by NroFactura  desc ")
    MsgBox temp(0) & " " & temp(1)
    NUM_FAC_A = temp(0)
    FEC_FAC_A = CDate(temp(1))
    
    temp = obtenerDeSQL("select top 1 NroFactura, fecha from FacturaVenta where activo = 1 and TipoDoc = 'FAB' order by NroFactura desc ")
    MsgBox temp(0) & " " & temp(1)
    NUM_FAC_B = temp(0)
    FEC_FAC_B = CDate(temp(1))
    
    temp = obtenerDeSQL("select max(codigo) from FacturaVenta ")
    MsgBox temp
    COD_FacturaVenta = temp

    temp = obtenerDeSQL("SELECT Max(NRO) AS Nro_ImpPro FROM IMPPRO")
    NUM_ImpPro = temp

    temp = obtenerDeSQL("SELECT Max(NRO) AS Expr1 FROM REC_COMP")
    OrdenPago = temp
    temp = obtenerDeSQL("SELECT Max(NRODOC) AS Expr1 FROM COMPRAS where tipoDoc = 'RAC' ")
    If temp > OrdenPago Then OrdenPago = temp
    temp = obtenerDeSQL("SELECT Max(NRODOC) AS Expr1 FROM TRANSCOM where tipoDoc = 'RAC' ")
    If temp > OrdenPago Then OrdenPago = temp
    
    MsgBox temp
   

    temp = obtenerDeSQL("SELECT Max(NRODOC) AS Expr1 FROM COMPRAS where tipoDoc = 'APC' ")
    MsgBox temp
    nume_apc = temp
    temp = obtenerDeSQL("SELECT Max(NRODOC) AS Expr1 FROM COMPRAS where tipoDoc = 'APD' ")
    MsgBox temp
    nume_apd = temp



    temp = obtenerDeSQL("SELECT Max(Numero) AS Expr1 FROM RemitoVenta")
    MsgBox temp
    Nro_RemitoVenta = temp

'   anticip_pr  ???

'
    temp = "insert into bs ( " _
        & " COD_FacturaVenta , Num_Factura_A, FEC_Factura_A, Num_Factura_B, FEC_Factura_B " _
        & " , Num_RemitoVenta, Num_impPro, Num_opago,Cta_Caja, idEmpresa, Num_APC, Num_APD  ) " _
        & " values (" _
        & COD_FacturaVenta & ", " _
        & NUM_FAC_A & ", " & ssFecha(FEC_FAC_A) & ", " _
        & NUM_FAC_B & ", " & ssFecha(FEC_FAC_B) & ", " _
        & Nro_RemitoVenta & " , " _
        & NUM_ImpPro & " , " _
        & OrdenPago & ", 1, 1,  " _
        & nume_apc & ", " & nume_apd _
        & " )"
    
    DataEnvironment1.Sistema.Execute temp
End Sub

Private Sub cmdComprasActualiz200607_Click()
    Dim rs As New ADODB.Recordset, s As String
    With rs
        
        s = " SELECT c.*, c.id as idc, p.NumIIBB AS pnro, p.TipoRetGan AS prgan, p.TipoRetIIBB AS prib " & _
            " FROM         COMPRAS c INNER JOIN " & _
            " Prov p ON c.CODPR = p.codigo " & _
            " WHERE     (c.NroIIBB = '') AND (c.CODPR > 0) AND (p.NumIIBB > '') and c.fecha > '20060401' "
        .Open s, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    
        While Not .EOF
            s = "update compras set NroIIBB = '" & !pnro & "', TipoRetGan = " & !prgan & ", TipoRetIIBB = " & !prib & " where id = " & !IDc
            DataEnvironment1.Sistema.Execute s
            
            .MoveNext
        Wend
        .Close
        che "fin compras, empiezo transcom"
        
        
        s = " SELECT c.*, c.id as idc, p.NumIIBB AS pnro, p.TipoRetGan AS prgan, p.TipoRetIIBB AS prib " & _
            " FROM         transcom  c INNER JOIN " & _
            " Prov p ON c.CODPR = p.codigo " & _
            " WHERE     (c.NroIIBB = '') AND (c.CODPR > 0) AND (p.NumIIBB > '') and c.fecha > '20060401' "
        .Open s, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    
        While Not .EOF
            s = "update transcom set NroIIBB = '" & !pnro & "', TipoRetGan = " & !prgan & ", TipoRetIIBB = " & !prib & " where id = " & !IDc
            DataEnvironment1.Sistema.Execute s
            
            .MoveNext
        Wend
        .Close
        che "fin adaptacion transcom"
        
    End With
    Set rs = Nothing
End Sub

Private Sub cmdCuitCliente_Click()
    Dim rs As New ADODB.Recordset
    With rs
        .Open "select cuit from clientes ", DataEnvironment1.Sistema, adOpenKeyset, adLockOptimistic
        While Not .EOF
            !Cuit = Replace(!Cuit, "/", "-")
            .Update
            .MoveNext
        Wend
        .Close
        che "ok cliente..."
        
        .Open "select cuit from facturaventa", DataEnvironment1.Sistema, adOpenKeyset, adLockOptimistic
        While Not .EOF
            !Cuit = Replace(!Cuit, "/", "-")
            .Update
            .MoveNext
        Wend
        .Close
        che "ok FV"
    
    
    End With
    Set rs = Nothing
End Sub

Private Sub cmdFCContado_Click()

'    Dim rs As New ADODB.Recordset, s As String, tempo As Double, i As Long
'    Dim j As Long
'
'    's = "select id, fecha, codpr, tipodoc, nrodoc, total from compras where activo = 1 and total > 0 and fecha > '20050101' and tipodoc = 'FAC' and contado = 1"
'    s = "select id, fecha, codpr, tipodoc, nrodoc, total from compras where activo = 1 and total > 0 and tipodoc = 'FAC' and contado = 1"
'
'    With rs
'        .Open s, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'        Debug.Print .RecordCount
'        While Not .EOF
'            tempo = s2n(obtenerDeSQL("select sum(importe) as suim from movicaja where tipodoc = 'FAC' and NroDoc = '" & !NroDoc & "' and codProv = '" & !CODPR & "' "))
'            If tempo = 0 Then
'                DataEnvironment1.Sistema.Execute "update compras set contado = 0 where id = " & !ID
'                i = i + 1
''                Debug.Print !ID & "   " & !Fecha & " " & !codpr & " " & !TIPODOC & " " & !NroDoc
'            End If
'            .MoveNext
'        Wend
'        Set rs = Nothing
'        che "pipi cucu " & i & ", gracias por apretar"
'    End With
End Sub

Private Sub cmdFV_cli_rzSoc_cuit_Click()
    Dim rs As New ADODB.Recordset
    Dim X

    rs.Open "select codigo ,cliente from FacturaVenta where RazonSocial is null", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    While Not rs.EOF
        DataEnvironment1.Sistema.Execute "update FacturaVenta set RazonSocial  = '" & obtenerDato("clientes", rs!cliente, "descripcion") & "' where codigo = " & rs!codigo
        rs.MoveNext
    Wend
    Set rs = Nothing

    cuit2
    Exit Sub
    X = rs!codigo
End Sub
Private Sub cuit2()
    Dim rs As New ADODB.Recordset
    Dim X

    rs.Open "select codigo ,cliente from FacturaVenta where cuit is null and (not cliente is null)", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    While Not rs.EOF
        DataEnvironment1.Sistema.Execute "update FacturaVenta set cuit = '" & obtenerDato("clientes", rs!cliente, "cuit") & "' where codigo = " & rs!codigo
        rs.MoveNext
    Wend
    Set rs = Nothing


    Exit Sub
    X = rs!codigo
End Sub


Private Sub cmdFV_Click()

'modifica casi todas las tablas , preset activo
If Not confirma("OjO!     >:-(     ¡pone  Activo en 1 y fecha_alta  en 1900, correcto para datos heredados  PERO NO para los nuevos!" & vbCrLf & " No pierde datos, pero se habiltarian los registros pseudo-borrados y se des-anularian comprobantes") Then Exit Sub
'If Not confirma("Confirma: OJO RESETEA A  ACTIVO = 1 y  FECHA_ALTA = 0") Then Exit Sub
'If Not confirma("Confirma de nuevo y por ULTIMA vez:  seguro?") Then Exit Sub
    
z "bancosgrales"
z "cajas"
z "categorias"
'z "centrodecostos"
z "cheques"
z "chq_comp"
z "clientes"
z "compras"
z "conceptos"
z "ctasbank"
z "cuentas"
'z "FacturaVenta"
z "FormasPago"
z "Formulas"
z "GruposProducto"
z "imppro"
z "ivas"
z "listas"
z "monedas"
z "motivosAjuste"
z "motivosRechazo"
z "moviBanc"
z "movicaja"
z "OrdenesDeCompras"
z "partesProduccion"
z "Pedidos_clientes"
z "producto"
z "prov"
'z "provincias"
z "rec_comp"
z "recibos"
z "remitoCompra"
z "series"
z "subGruposProducto"
'z "TipoCompras"
z "TipoComprobantesGrales"
z "TipoDocumentos"
z "TransCom"
z "transportes"
z "unidadesMedida"
z "usuarios"

MsgBox "fin sin error"

End Sub


Private Sub cmdFV_Contado_Click()
    Dim rs As New ADODB.Recordset, i As Long
    
    che "Transforma registros FacturaVenta TipoDoc = 'CON' a campo .Contado" & vbCrLf & " Despues los borra, o sea sirve 1 sola vez"
    With rs
        .Open "select TipoDoc, NroFactura, Cliente  from FacturaVenta where tipodoc = 'CON' ", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        While Not .EOF
            DataEnvironment1.Sistema.Execute "update FacturaVenta set contado = 1 where cliente = " & rs!cliente & " and NroFactura = " & !nrofactura & " and ( TipoDoc = 'FAA' or TipoDoc = 'FAB') "
            .MoveNext
        Wend
        .Close
        .Open "SELECT COUNT(Contado) AS contadas From FacturaVenta Where (contado = 1)", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        MsgBox "contado = " & !contadas
        .Close
    End With
    
    Set rs = Nothing
    DataEnvironment1.Sistema.Execute "delete from facturaVenta where TipoDoc = 'CON'"
End Sub


Private Sub cmdFVmodStock_Click()
    Dim rs As New ADODB.Recordset, i As Long
    With rs
        .Open "SELECT DISTINCT TipoDoc, NroFactura, NroRemito FROM FacturaVentaDetalle AS d WHERE NroRemito=0", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        i = 0
        While Not .EOF
            i = i + 1
            DataEnvironment1.Sistema.Execute "update FacturaVenta set ActualizaStock = 1 where TipoDoc = '" & rs!TIPODOC & "' and NroFactura = " & s2n(rs!nrofactura)
            .MoveNext
        Wend
    End With
    MsgBox "cambiados " & i
    Set rs = Nothing
End Sub

Private Sub cmdFvMoneda_Click()
End Sub

Private Sub cmdImporteOC_Click()
    'db.Execute "update
    Dim rs As New ADODB.Recordset, i As Long
    
    With rs
        .Open "select ordenCompra, sum(costo * cantidad) as TotItems from ItemordenCompra group by OrdenCompra", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        i = 0
        While Not .EOF
            i = i + 1
            DataEnvironment1.Sistema.Execute "update OrdenesDeCompras set importe = " & x2s(!totitems) & "  where codigo = " & !ordenCompra
            .MoveNext
        Wend
    End With
    MsgBox "cambiados " & i
    Set rs = Nothing
End Sub


Private Sub cmdMaeART_Click()
    Dim rs As New ADODB.Recordset
    Dim i As Long
    
    With rs
        .Open "select * from _mae_art", DataEnvironment1.Sistema, adOpenKeyset, adLockOptimistic
        While Not .EOF
            If IsNumeric(!modelo) Then
                !cod = Trim(CLng(!modelo))
            Else
                For i = 1 To Len(!modelo)
                    If Mid(!modelo, i, 1) <> "0" Then
                        Exit For
                    End If
                Next i
                !cod = Mid(!modelo, i)
            End If
            .Update
            .MoveNext
        Wend
        .Close
        
        .Open "SELECT * FROM _MAE_ART m INNER JOIN Producto p ON m.cod = p.codigo "
        While Not .EOF
'            !codigobarra = !
'            .Update
            .MoveNext
        Wend
        Stop
    End With
    
    Set rs = Nothing
End Sub


Private Sub cmdNCT_Click()
    DataEnvironment1.Sistema.Execute "update facturaventa set tipodoc = 'NDA', nd_xChequeRechazado = 1 where tipodoc = 'NDR'"
    che "ta"
End Sub

Private Sub cmdOC_Moneda_Click()
    DataEnvironment1.Sistema.Execute "update OrdenesDeCompras set moneda = 1 where moneda is null"
    che "ta"
End Sub

Private Sub cmdRemigraMovicaja_Click()
    Dim rs As New ADODB.Recordset, s, w, i As Long, t
    
    
    Exit Sub
    
    
    
    relojito True
    With rs
        .Open "select * from zzCOMPRAS " _
                    & " where contado = 1 and  codpr_co > 0 and tipodoc_co = 'FAC'  AND total_co > 0  " _
                    , DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                    
        While Not .EOF
            w = " where tipodoc = 'FAC' and NroDoc = " & !NroDoc_co & " and codpr = " & !codpr_co & " and activo = 1 and suc = " & !suc_co
            t = s2n(obtenerDeSQL("select count(tipodoc) from compras " & w))
            If t > 0 Then
                DataEnvironment1.Sistema.Execute "update compras  set contado = 1 " & w
                i = i + 1
'            ElseIf t > 1 Then
'                che t & " NroDoc = " & !NroDoc_co & " -- codpr = " & !codpr_co & " -- fecha " & !fecha_co
            End If
            .MoveNext
        Wend
    End With
    relojito False
    che "ok " & i
    Set rs = Nothing
End Sub

Private Sub cmdResetNull_Click()
    MsgBox "elimina null de los campos programados -no destructivo-"
    
    nn "FacturaVenta", "pedido"
    nn "FacturaVenta", "remito"
    nn "FacturaVenta", "total"
    nn "FacturaVenta", "saldo"
    nb "facturaVenta", "Moneda"
    nb "facturaVenta", "Cotizacion"
    nn "facturaVenta", "iva"
    nn "facturaVenta", "iibb"
    ns "facturaVenta", "cuit"
    nn "facturaVenta", "nograv"
    nn "facturaVenta", "descuento"
    
    nb "FacturaVentaDetalle", "codpropio"
    ns "FacturaVentaDetalle", "formula"
    
    nn "Cheques", "prov"
    nn "Cheques", "ndocprov"
    nn "cheques", "Dep_Cuenta"
    nn "cheques", "ch3_tdocpr"
    
    ns "TransCom", "RazonSocialProv"
    ns "TransCom", "CuitProv"
    nn "TransCom", "Cotizacion"
    nn "TransCom", "Moneda"
    ns "TransCom", "FormaDePago"
    nn "TransCom", "IbProvincia"
    
    nn "compras", "ibProvincia"
    
    
    nn "Compras", "Moneda"
     
    ns "RelFNR_c", "totalDocu"
    
    ns "MoviCaja", "codFp"
    
    ns "clientes", "contacto"
    ns "clientes", "barrio"
    ns "clientes", "cuit"
    ns "clientes", "fax"
    ns "clientes", "direccion_comercial"
    ns "clientes", "localidad_comercial"
    ns "clientes", "zonacomercial"
    ns "clientes", "proveedor"
    ns "clientes", "certificado"
    ns "clientes", "descripcion"
    nn "clientes", "Correo"
'    nn "clientes",
    ns "clientes", "direccion"
    ns "clientes", "localidad"
    ns "clientes", "telefono"
'    ns "clientes", ""
'    ns "clientes",
'    ns "clientes",
    
    ns "chq_comp", "CuentaBancaria"
    ns "chq_comp", "importe"
    ns "chq_comp", "fecha_cheque"
    ns "chq_comp", "fechadeposito"
    ns "chq_comp", "nrodoc"
    ns "chq_comp", "tipodoc"
    nn "chq_comp", "proveedor"
    
    nn "CtasBank", "Numero"
    nn "CtasBank", "Moneda"

    ns "prov", "cuit"
    ns "ivas", "descripcion"
    nn "itemordencompra", "fechaentrega"

    nn "facturaventa", "neto"
    ns "facturaVenta", "provincia"
    nn "facturaVenta", "vencimiento"
   
    ns "Producto", "formula"
    ns "producto", "observaciones"
    ns "producto", "calcsincosto"
    ns "producto", "puedofac"
    
    
    nn "producto", "CALCSINCOSTO"
    nn "producto", "serie"
    nn "producto", "puedofac"
    nn "clientes", "formapago"
    
    nn "clientes", "vendedor"
    nn "clientes", "categoria"
    
    
    
'6/1/5
    nn "TransCom", "Vencim"
'17/1/5
    ns "RemitoVentaDetalle ", "Formula"
    
'3/5/5
    ns "RemitoVenta", "Factura"
    
    ns "movibanc", "documento"
    
    ns "cuentas", "sumariza"
    
    ns "Relacion_Producto_cliente", "Letra"
    
    ns "FacturaVenta", "RazonSocial"
    
    MsgBox "ta"
End Sub

Sub nn(tabula, campuloN)
    DataEnvironment1.Sistema.Execute "update " & tabula & " set " & campuloN & " = 0 where " & campuloN & " is null"
    txttabula = tabula & " " & campuloN
End Sub
Sub ns(tabula, campuloS)
    DataEnvironment1.Sistema.Execute "update " & tabula & " set " & campuloS & " = '' where " & campuloS & " is null"
    txttabula = tabula & " " & campuloS
End Sub
Sub nb(tabula, campuloB)
    DataEnvironment1.Sistema.Execute "update " & tabula & " set " & campuloB & " = 1 where " & campuloB & " is null"
    txttabula = tabula & " " & campuloB
End Sub
Sub z(tabula)
    txttabula = tabula
    DataEnvironment1.Sistema.Execute "update " & tabula & " set fecha_baja = 0, fecha_alta = " & ssFecha(#1/1/2000#) & " , usuario_alta = 100, usuario_baja= 0 , activo = 1"
End Sub

'Private Sub cmdRetenc_Click()
'    Dim rs As New ADODB.Recordset ', rsf As New ADODB.Recordset
'    Dim codiR, cupa 'As Long
'    Dim i As Long
'    With rs
'        .Open "select * from facturaventa where (tipodoc = 'RIB' or tipodoc = 'RBO' or tipodoc = 'RCP' or tipodoc = 'RGA' or tipodoc = 'RET' ) and activo =  1 and fecha >=  '20040101' ", DataEnvironment1.Sistema, adOpenKeyset, adLockOptimistic
'        While Not .EOF
''            codiR = obtenerDeSQL("select codigo from facturaVenta where ")
'            cupa = obtenerDeSQL("select id FROM CUENTASPARAM where codigo = '" & !TIPODOC & "' ")
'
'            DataEnvironment1.Sistema.Execute "insert into RecibosRetenciones " _
'                & " (iddoc, idCuentasParam, Numero, Fecha, Importe, nroFactura, cuenta) values " _
'                & " ('" & !iddoc & "', " & cupa & ", '" & !nroFactura & "', " & ssFecha(!Fecha) & ", " & x2s(!Total) & ",  " & x2s(0) & ", '" & 0 & "' ) "
'            !activo = False
'            !usuario_alta = 100
'            .Update
'            i = i + 1
'            .MoveNext
'        Wend
'        .Close
'        che i
'    End With
'End Sub

Private Sub cmdSacarComitaProducto_Click()
    Dim rs As New ADODB.Recordset, i As Long
    
    With rs
        .Open "select descripcion from producto where descripcion like '%''%'", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        While Not .EOF
'        MsgBox !Descripcion
            !DESCRIPCION = Replace(!DESCRIPCION, "'", "´")
            .Update
            i = i + 1
'        MsgBox !Descripcion
            .MoveNext
        Wend
        Set rs = Nothing
    End With
    che "ta, cambiados: " & i
End Sub

Private Sub cmdTipoIvaTC_Click()
    Dim rs As New ADODB.Recordset
   Dim tI As Long
    With rs
        .Open "select id, codpr from transcom", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        While Not .EOF
            tI = nSinNull(obtenerDeSQL("select tipoiva from prov where codigo = '" & !CODPR & "' "))
            DataEnvironment1.Sistema.Execute "update transcom set tipoiva = '" & tI & "' where id = '" & !ID & "' "
            .MoveNext
        Wend
    End With
    che "listo"
End Sub
Private Sub cmdFC_rzSoc_Click()
    Dim rs As New ADODB.Recordset
    Dim i
    i = 0
    rs.Open "SELECT id, codpr From TRANSCOM WHERE RAZONSOCIALPROV = ''", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    While Not rs.EOF
        DataEnvironment1.Sistema.Execute "update transcom set RAZONSOCIALPROV  = '" & obtenerDato("prov", rs!CODPR, "descripcion") & "' where id  = '" & rs!ID & "'"
        i = i + 1
        rs.MoveNext
    Wend
    rs.Close
    
    che i & " .Tá razon, va cuit..."
    i = 0
    rs.Open "SELECT id, CODPR  FROM   TRANSCOM WHERE  (CUITPROV = '')", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    While Not rs.EOF
        DataEnvironment1.Sistema.Execute "update transcom set  cuitprov = '" & obtenerDato("prov", rs!CODPR, "cuit") & "' where id  = '" & rs!ID & "'"
        rs.MoveNext
        i = i + 1
    Wend
    rs.Close
    
    che "ya tá, " & i
    
    Set rs = Nothing
End Sub
Private Sub Command1_Click()
    Dim s
    s = "UPDATE Series SET fecha_alta = " & ssFecha(#1/1/2005#) & " WHERE fecha_alta < " & ssFecha(#1/1/2000#)
    
    DataEnvironment1.Sistema.Execute s
    
    s = "update series set essalida = 1 where comprobante = 1 or comprobante = 2 or comprobante = 5 or comprobante = 8"
    DataEnvironment1.Sistema.Execute s
    
    che "listo"
End Sub

Private Sub Command2_Click()
    DataEnvironment1.Sistema.Execute "update FacturaVenta set TipoIva = 2 where TipoIva is null "
    che "fin"
End Sub

Private Sub Command3_Click()
    Dim rs As New ADODB.Recordset, X As String
    
    With rs
        .Open "select * from DBF_Prov", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        While Not .EOF
            X = sSinNull(!NUMIIBB_PR)
            If X > "" Then
                DataEnvironment1.Sistema.Execute _
                    " update prov " & _
                    " set NumIIBB = '" & X & "' " & _
                    " where codigo = " & !cod_pr
            End If
            .MoveNext
        Wend
    End With
    Set rs = Nothing
    che "ta"
    
End Sub

Private Sub Command4_Click()
    Dim rs As New ADODB.Recordset
    Dim i As Long
    
    With rs
        
        .Open "SELECT * FROM Clientes LEFT OUTER JOIN  ClieMAY ON Clientes.direccion = ClieMAY.DIRECC_CL order by idcliente", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    
        While Not .EOF
            DataEnvironment1.Sistema.Execute _
                "update clientes set " & _
                " codigopostal = '" & !codpos_cl & "', " & _
                " categoria = '" & !catego_cl & "' , " & _
                " fechacumpleanio = " & nuf(!feccum_cl) & " , " & _
                " pais = '" & !Pais & "', " & _
                " provincia = '" & !Provincia & "' " & _
                " where idcliente = " & !IDcliente
      
                i = i + 1
            .MoveNext
        Wend
        che "termineeeee MAYORISTA  "
        .Close
        
        
        
        
        
        i = 0
        
        .Open "SELECT * FROM Clientesmailing LEFT OUTER JOIN  ClieMin ON Clientesmailing.direccion = ClieMin.DIRECC_CL order by idclimin", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        While Not .EOF
        
            DataEnvironment1.Sistema.Execute _
                "update clientesmailing set " & _
                " codigopostal = '" & !codpos_cl & "', " & _
                " categoria = '" & !catego_cl & "' , " & _
                " fechacumpleanio = " & nuf(!feccum_cl) & " , " & _
                " pais = '" & !Pais & "', " & _
                " provincia = '" & !Provincia & "' " & _
                " where idclimin = " & !IDclimin
      
                i = i + 1
            .MoveNext
        Wend
        che "termineeeee MINORISTA "
        .Close
        
        
        
        
        
        
        
    End With

    
    
    
    
    
    
    Set rs = Nothing
    
    
    
End Sub
Private Function nuf(X)
    If IsNull(X) Then
        nuf = " NULL "
    Else
        nuf = ssFecha(CDate(X))
    End If
End Function


'Private Sub Command3_Click()
'    DataEnvironment1.Sistema.Execute "UPDATE Compras as c inner join Prov AS p ON p.codigo = c.CODPR SET c.CUITPROV = [p].[cuit], c.RAZONSOCIALPROV = [p].[descripcion]"
'    che "ta"
'End Sub

Private Sub formula_Click()
    Dim rs As New ADODB.Recordset, resu As String
    With rs
        .Open "select distinct producto.codigo from producto inner join formulas on producto.codigo = formulas.codigo where formulas.activo = 1", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        While Not .EOF
            DataEnvironment1.Sistema.Execute "update producto set formula = 1 where codigo = '" & rs!codigo & "'"
            .MoveNext
        Wend
    End With
    Set rs = Nothing
    che "ya ta"
End Sub

Private Sub remitoconfactura_Click()
    Dim s As String, rs As New ADODB.Recordset
    
    s = "SELECT f.id, f.TipoDoc, f.NroFactura, r.Numero, d.codigo, d.Producto " _
        & " FROM RemitoVentaDetalle AS d INNER JOIN (RemitoVenta AS r INNER JOIN FacturaVentaDetalle AS f ON r.Factura = f.NroFactura) ON (d.Producto = f.Producto) AND (d.Numero = r.Numero) " _
        & " WHERE (((f.TipoDoc)='FAA' Or (f.TipoDoc)='FAB'))"
        
    With rs
        .Open s, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        While Not .EOF
            s = "update facturaventadetalle set item_p_r = " & !codigo & " where id = " & !ID
            DataEnvironment1.Sistema.Execute s
            .MoveNext
        Wend
    End With
    che "ya ta"
End Sub


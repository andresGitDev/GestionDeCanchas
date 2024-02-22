VERSION 5.00
Begin VB.Form FrmPrincipal 
   ClientHeight    =   10650
   ClientLeft      =   165
   ClientTop       =   0
   ClientWidth     =   14520
   Icon            =   "FrmPrincipal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "FrmPrincipal.frx":08CA
   ScaleHeight     =   10650
   ScaleWidth      =   14520
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdAgenda 
      Caption         =   "Agenda"
      Height          =   960
      Left            =   3750
      Picture         =   "FrmPrincipal.frx":19B93
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   180
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Timer Timer1 
      Left            =   8565
      Top             =   135
   End
   Begin Gestion.uActualizar uActualizar1 
      Height          =   2280
      Left            =   150
      TabIndex        =   14
      Top             =   2055
      Width           =   2610
      _extentx        =   4604
      _extenty        =   4022
   End
   Begin VB.Frame fraBs 
      Height          =   5025
      Left            =   315
      TabIndex        =   0
      Top             =   5010
      Visible         =   0   'False
      Width           =   8430
      Begin VB.CommandButton Command5 
         Caption         =   "Reiniciar Sist"
         Height          =   285
         Left            =   135
         TabIndex        =   21
         Top             =   2010
         Width           =   1245
      End
      Begin VB.CommandButton cmdLogo 
         Caption         =   "Logo"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1740
         Width           =   1260
      End
      Begin VB.CheckBox chkVerTablaTemp 
         Caption         =   "Tablas temporarias sin #"
         Height          =   300
         Left            =   1455
         TabIndex        =   18
         Top             =   1725
         Width           =   2340
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Prueba1"
         Height          =   255
         Left            =   5670
         TabIndex        =   17
         Top             =   2700
         Width           =   1260
      End
      Begin VB.CheckBox chkPreviewImpresion 
         Caption         =   "Preview en impresiones ActiveReport"
         Height          =   300
         Left            =   1455
         TabIndex        =   16
         Top             =   1470
         Width           =   3135
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Prueba3"
         Height          =   255
         Left            =   5670
         TabIndex        =   15
         Top             =   3300
         Width           =   1245
      End
      Begin VB.TextBox txtCarpetaExe 
         Height          =   315
         Left            =   1800
         TabIndex        =   13
         Top             =   600
         Width           =   3630
      End
      Begin VB.Frame fraBackup 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   450
         Left            =   90
         TabIndex        =   9
         Top             =   195
         Width           =   5400
         Begin VB.CommandButton cmdBackup 
            Caption         =   "Generar Backup en:"
            Height          =   285
            Left            =   60
            TabIndex        =   11
            Top             =   30
            Width           =   1605
         End
         Begin VB.TextBox txtBackup 
            Height          =   345
            Left            =   1695
            TabIndex        =   10
            Text            =   "C:\"
            Top             =   0
            Width           =   3660
         End
      End
      Begin VB.CheckBox chkDesbichando 
         Caption         =   "Debug errores    "
         Height          =   300
         Left            =   1455
         TabIndex        =   7
         Top             =   1215
         Width           =   3135
      End
      Begin VB.TextBox txtBug 
         Height          =   1500
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Top             =   2355
         Width           =   5370
      End
      Begin VB.CommandButton cmdMigrar 
         Caption         =   "Pos-Migracion"
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1260
      End
      Begin VB.CommandButton cmdLog 
         Caption         =   "log environment"
         Height          =   450
         Left            =   120
         TabIndex        =   1
         Top             =   1275
         Width           =   1260
      End
      Begin VB.Label Label1 
         Caption         =   "ok"
         Height          =   1395
         Left            =   5580
         TabIndex        =   12
         Top             =   900
         Width           =   2910
      End
      Begin VB.Label lblSabeloquehace 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         Height          =   315
         Left            =   1770
         TabIndex        =   8
         Top             =   945
         Width           =   3540
      End
      Begin VB.Label lblUfa 
         Caption         =   "Modo Debug                          ModuloLi .    _STOP = True"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5505
         TabIndex        =   2
         Top             =   255
         Visible         =   0   'False
         Width           =   2955
      End
   End
   Begin VB.Label lblModo 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   6120
      TabIndex        =   23
      Top             =   120
      Width           =   90
   End
   Begin VB.Label lblSucursal 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3600
      TabIndex        =   20
      Top             =   600
      Width           =   3075
   End
   Begin VB.Image imgLogoSimple 
      DragMode        =   1  'Automatic
      Height          =   1635
      Left            =   105
      Picture         =   "FrmPrincipal.frx":1A45D
      Stretch         =   -1  'True
      Top             =   135
      Width           =   2640
   End
   Begin VB.Label lblEnvironment 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7035
      TabIndex        =   6
      Top             =   1215
      Width           =   3075
   End
   Begin VB.Label lblNombreEmpresa 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3645
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   3060
   End
   Begin VB.Menu COMPRAS 
      Caption         =   "&COMPRAS"
      Begin VB.Menu TABLASCOMP 
         Caption         =   "ABM Compras"
         Begin VB.Menu AbmCostos 
            Caption         =   "Centro de costos"
         End
         Begin VB.Menu mnuAjustCosto 
            Caption         =   "Ajuste de Centro de Costos"
         End
         Begin VB.Menu PROVEEDDORES 
            Caption         =   "Proveedores"
         End
         Begin VB.Menu mnuListProv 
            Caption         =   "Listado de Proveedores"
         End
         Begin VB.Menu RELPRODPROV 
            Caption         =   "Relacion Producto Proveedor"
         End
         Begin VB.Menu mnuNumeracionC 
            Caption         =   "Numeracion"
         End
         Begin VB.Menu mnuFormasPagoCompras 
            Caption         =   "Formas de pago"
         End
         Begin VB.Menu TIPOCOMP 
            Caption         =   "Tipo de Compras"
         End
         Begin VB.Menu mnuTioProv 
            Caption         =   "Tipo de Proveedores"
         End
         Begin VB.Menu mnuCuentas001 
            Caption         =   "Cuentas de proveedor"
         End
      End
      Begin VB.Menu ORDCOMPRA 
         Caption         =   "Ordenes de Compras"
      End
      Begin VB.Menu RemitoCompra 
         Caption         =   "Remito Compras"
      End
      Begin VB.Menu mnuProcesos 
         Caption         =   "Carga de Comprobantes"
         Begin VB.Menu facturas 
            Caption         =   "Factura Proveedor"
         End
         Begin VB.Menu mnuFacturaProvRemito 
            Caption         =   "Factura Proveedor sobre Remito"
         End
         Begin VB.Menu NotaCredito 
            Caption         =   "Nota de Credito"
         End
         Begin VB.Menu NotaDebito 
            Caption         =   "Nota de Debito"
         End
         Begin VB.Menu Ajustes 
            Caption         =   "Ajuste por Credito/Debito"
         End
         Begin VB.Menu PagosACuenta 
            Caption         =   "Pagos a Cuenta"
         End
         Begin VB.Menu OrdenPago 
            Caption         =   "Orden de Pago/Imputaciones"
         End
      End
      Begin VB.Menu mnuLoca 
         Caption         =   "DATOS DE COLECTOR"
      End
   End
   Begin VB.Menu VENTAS 
      Caption         =   "VENTAS"
      Begin VB.Menu TABLAS 
         Caption         =   "ABM Ventas"
         Begin VB.Menu CLIENTES 
            Caption         =   "Clientes"
         End
         Begin VB.Menu mnuCont 
            Caption         =   "Contacto"
         End
         Begin VB.Menu mnuSuc 
            Caption         =   "Sucursal"
         End
         Begin VB.Menu PRODUCTOS 
            Caption         =   "Productos"
         End
         Begin VB.Menu VENDEDORES 
            Caption         =   "Vendedores"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu FORMASPAGO 
            Caption         =   "Formas de Pago"
         End
         Begin VB.Menu CATEGORIAS 
            Caption         =   "Categorias"
         End
         Begin VB.Menu TRANSPORTES 
            Caption         =   "Transportes"
         End
         Begin VB.Menu mnuNumeracion2 
            Caption         =   "Numeracion"
         End
         Begin VB.Menu mnuRelProdCliente 
            Caption         =   "Relacion Producto Cliente"
         End
         Begin VB.Menu mnuTexto 
            Caption         =   "Textos"
         End
         Begin VB.Menu mnuLeyFact 
            Caption         =   "Leyenda de Factura"
         End
         Begin VB.Menu mnucuentas002 
            Caption         =   "Cuentas de clientes"
         End
      End
      Begin VB.Menu mnuFactura 
         Caption         =   "Emision de Comprobantes"
         Begin VB.Menu facturaLibre 
            Caption         =   "Factura Libre"
         End
         Begin VB.Menu FacturaRemito 
            Caption         =   "Factura sobre Remito"
            Visible         =   0   'False
         End
         Begin VB.Menu FacturaPedido 
            Caption         =   "Factura sobre Pedido"
         End
         Begin VB.Menu mnuAcFac 
            Caption         =   "Actualizar facturas"
         End
         Begin VB.Menu mnuNCredito 
            Caption         =   "Nota de Credito"
         End
         Begin VB.Menu mnuNCreditoDevolucion 
            Caption         =   "Nota de Credito por Devolucion"
         End
         Begin VB.Menu mnuNotaDebito 
            Caption         =   "Nota de Debito"
         End
         Begin VB.Menu mnuND_Venta_ChRechazado 
            Caption         =   "Nota de Debito por Cheque Rechazado"
         End
         Begin VB.Menu mnuAjusteVenta 
            Caption         =   "Ajuste"
         End
         Begin VB.Menu m1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecibosACuenta 
            Caption         =   "Recibo a Cuenta"
         End
         Begin VB.Menu mnuRecibos 
            Caption         =   "Recibo"
         End
         Begin VB.Menu mnuRecibosImputacion 
            Caption         =   "Imputaciones"
         End
      End
      Begin VB.Menu mnuExterior 
         Caption         =   "Exterior"
         Begin VB.Menu mnuFacturaLibreEXT 
            Caption         =   "Factura Libre Exterior"
         End
         Begin VB.Menu mnuFactRemitoEXT 
            Caption         =   "Factura sobre Remito Exterior"
         End
         Begin VB.Menu mnuFactPedidoEXT 
            Caption         =   "Factura sobre Pedido Exterior"
         End
         Begin VB.Menu mnuNCEXT 
            Caption         =   "Nota de Credito Exterior"
         End
         Begin VB.Menu mnuNCDevolucionEXT 
            Caption         =   "Nota de Credito por Devolucion Exterior"
         End
         Begin VB.Menu mnuLeyendaFAE 
            Caption         =   "Leyenda FAE"
         End
      End
      Begin VB.Menu PEDCLI 
         Caption         =   "Pedidos de Clientes"
      End
      Begin VB.Menu mnuCancelarPedido 
         Caption         =   "Cancelacion de Pedidos"
      End
      Begin VB.Menu MREMITOS 
         Caption         =   "Remitos"
         Begin VB.Menu RemitoVenta 
            Caption         =   "Mercaderia en Transito Salida"
         End
         Begin VB.Menu mCancelaRemito 
            Caption         =   "Cancelacion de Remitos"
         End
         Begin VB.Menu ESTADOREMITO 
            Caption         =   "ESTADO DEL REMITO"
         End
         Begin VB.Menu mnuMercTran 
            Caption         =   "Mercaderia en Transito Devuelta"
         End
         Begin VB.Menu mnuPRemito 
            Caption         =   "Puntos de Remito"
         End
      End
      Begin VB.Menu mnuPresupuesto 
         Caption         =   "Presupuestos"
      End
      Begin VB.Menu mnuPorte 
         Caption         =   "Carta de Porte"
      End
      Begin VB.Menu mnuStkPrecios 
         Caption         =   "Precios"
      End
      Begin VB.Menu mnuPedidosPendienteBig 
         Caption         =   "Pedidos Pendientes"
      End
      Begin VB.Menu mnuSaldoRemito 
         Caption         =   "LIS SALDO REMITOS"
      End
   End
   Begin VB.Menu mnubol 
      Caption         =   "FORMULARIOS"
      Begin VB.Menu mnuBoletas 
         Caption         =   "Formularios y Otros"
      End
      Begin VB.Menu mnuCoef 
         Caption         =   "Actualizar Coeficientes de IIBB"
      End
   End
   Begin VB.Menu CAJASYBANC 
      Caption         =   "&CAJAS Y BANCOS"
      Begin VB.Menu mnuABMCYB 
         Caption         =   "ABM"
         Begin VB.Menu CAJAS 
            Caption         =   "Cajas"
         End
         Begin VB.Menu BANCOGRAL 
            Caption         =   "Bancos"
         End
      End
      Begin VB.Menu cajasyFondos 
         Caption         =   "Cajas y Fondos Fijos"
      End
      Begin VB.Menu GastosBancarios 
         Caption         =   "Debitos/Creditos Bancarios"
      End
      Begin VB.Menu GastosBancarios2 
         Caption         =   "Gastos Bancarios Ampliado"
      End
      Begin VB.Menu CuentasBancarias 
         Caption         =   "Cuentas Bancarias"
      End
      Begin VB.Menu IngresoChequera 
         Caption         =   "Ingreso de Chequera"
      End
      Begin VB.Menu LibracionCheques 
         Caption         =   "Libracion de Cheques"
      End
      Begin VB.Menu mnuSalChTer 
         Caption         =   "Salida de Cheques de Terceros"
      End
      Begin VB.Menu DebitoCheques 
         Caption         =   "Debito de Cheques"
      End
      Begin VB.Menu AnuloCheques 
         Caption         =   "Anulacion de Cheques"
      End
      Begin VB.Menu TransfBancarias 
         Caption         =   "Transferencia Bancaria"
      End
      Begin VB.Menu IngChTerceros 
         Caption         =   "Ingreso de Cheques de Terceros"
      End
      Begin VB.Menu ProcesoCheques 
         Caption         =   "Proceso de Cheques"
      End
      Begin VB.Menu EXTRACTOBANCARIO 
         Caption         =   "Extractos Bancarios"
      End
   End
   Begin VB.Menu STOCK 
      Caption         =   "STOCK"
      Begin VB.Menu Abms 
         Caption         =   "ABMS"
         Begin VB.Menu mnuProductos 
            Caption         =   "Productos"
         End
         Begin VB.Menu GRUPOS 
            Caption         =   "Grupos de Productos"
         End
         Begin VB.Menu SUBGRUPOS 
            Caption         =   "SubGrupos de Productos"
         End
         Begin VB.Menu mnuTP 
            Caption         =   "Tipo Producto"
         End
         Begin VB.Menu COMP 
            Caption         =   "Comprobantes para Ajuste"
         End
         Begin VB.Menu SERIES 
            Caption         =   "Series"
         End
         Begin VB.Menu mnuSeries2 
            Caption         =   "Series Altas"
         End
         Begin VB.Menu mnuVerSeries 
            Caption         =   "Series Modificacion"
         End
         Begin VB.Menu mmm000 
            Caption         =   "-"
         End
         Begin VB.Menu FORMULAS 
            Caption         =   "Formulas"
         End
         Begin VB.Menu mnuCopiaFormula 
            Caption         =   "Copiar Formula"
         End
         Begin VB.Menu mnuReemplazoComponente 
            Caption         =   "Reemplazo componentes en formula"
         End
         Begin VB.Menu mmm0001 
            Caption         =   "-"
         End
         Begin VB.Menu CONCEPTOS 
            Caption         =   "Conceptos"
         End
         Begin VB.Menu PRODCLI 
            Caption         =   "Relacion Producto Cliente"
         End
         Begin VB.Menu mnuUnidadesMedida0 
            Caption         =   "-"
         End
         Begin VB.Menu mnuUnidadesMedida 
            Caption         =   "Unidades de Medida"
         End
         Begin VB.Menu mnuUnidadesMedida2 
            Caption         =   "Tipo de Medida"
         End
         Begin VB.Menu mnuUnidadesMedida3 
            Caption         =   "Factor de Conversion"
         End
      End
      Begin VB.Menu menuDespacho 
         Caption         =   "Despacho"
         Begin VB.Menu menuD1 
            Caption         =   "Generar Despacho"
         End
         Begin VB.Menu menuD2 
            Caption         =   "Informe de Despachos"
         End
      End
      Begin VB.Menu mnup001 
         Caption         =   "Partes"
         Begin VB.Menu mnuParte 
            Caption         =   "Generar Parte de Produccion"
         End
         Begin VB.Menu mnuinfpartes 
            Caption         =   "Informe de Partes de Produccion"
         End
         Begin VB.Menu mnuinfpartes2 
            Caption         =   "Informe de Movimientos de Partes"
         End
      End
      Begin VB.Menu mnua001 
         Caption         =   "Ajustes"
         Begin VB.Menu mnuDiferencia 
            Caption         =   "Diferencia de Stock"
         End
         Begin VB.Menu mnua002 
            Caption         =   "Informe de Ajustes"
         End
      End
      Begin VB.Menu mnuPrecios 
         Caption         =   "Precios Compra venta"
      End
      Begin VB.Menu mmm003 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMovStock 
         Caption         =   "Movimiento de Stock Con Parte de Produccion"
      End
      Begin VB.Menu mnuStock2 
         Caption         =   "Movimiento de Stock"
      End
      Begin VB.Menu mS01 
         Caption         =   "Stock Actual"
         Visible         =   0   'False
      End
      Begin VB.Menu LisProductos2 
         Caption         =   "Productos + Formula + Movimientos"
      End
      Begin VB.Menu LisProductos 
         Caption         =   "Productos"
      End
      Begin VB.Menu mnuFormulas 
         Caption         =   "Formulas"
      End
   End
   Begin VB.Menu INF 
      Caption         =   "INFORMES"
      Begin VB.Menu CAJBANC 
         Caption         =   "Cajas y Bancos"
         Begin VB.Menu MovCajas 
            Caption         =   "Movimiento de Caja"
         End
         Begin VB.Menu mnuMovCajasBig 
            Caption         =   "Movimiento de Caja (New)"
         End
         Begin VB.Menu mmnh 
            Caption         =   "-"
         End
         Begin VB.Menu LisCheques 
            Caption         =   "Listado de Cheques"
         End
         Begin VB.Menu LisChTerceros 
            Caption         =   "Listado de Cheques de Terceros"
         End
         Begin VB.Menu mnuhistorico 
            Caption         =   "Historico de Cheques de Terceros"
         End
         Begin VB.Menu LisMovBancario 
            Caption         =   "Listado por movimiento Bancario"
         End
         Begin VB.Menu mnuExtractoBancario 
            Caption         =   "Extracto Bancario"
         End
         Begin VB.Menu EXTRACTOSBANCARIOSOLD 
            Caption         =   "Extracto Bancario Old"
         End
      End
      Begin VB.Menu COMPR 
         Caption         =   "Compras"
         Begin VB.Menu LisProveedores 
            Caption         =   "Proveedores"
         End
         Begin VB.Menu mnuCProv2 
            Caption         =   "Cuentas de Proveedores"
         End
         Begin VB.Menu SlProveedores 
            Caption         =   "Saldo de Proveedores"
         End
         Begin VB.Menu mnuSaldoProvSinDetalle 
            Caption         =   "Saldo de Proveedores sin Detalle"
         End
         Begin VB.Menu IVACOMPRAS 
            Caption         =   "IVA Compras"
         End
         Begin VB.Menu DETALLECOMPROBPROV 
            Caption         =   "Detalle Mensual de Comprobantes de Proveedores"
         End
         Begin VB.Menu mnuMovCtaProv 
            Caption         =   "Movimiento de Cuenta de Proveedores"
         End
         Begin VB.Menu mnuCompProv 
            Caption         =   "Composicion de Proveedores sin detalle"
         End
         Begin VB.Menu TOTTIPOCOMPRA 
            Caption         =   "Totales por tipo de Compras"
         End
         Begin VB.Menu mnuLisOC 
            Caption         =   "Ordenes de Compra"
         End
         Begin VB.Menu mnuiibbjur 
            Caption         =   "Listado de IIBB por Jurisdiccion"
         End
         Begin VB.Menu mnuListPerc 
            Caption         =   "Listado de Percepciones"
         End
         Begin VB.Menu mnuVerPagos 
            Caption         =   "Pagos"
         End
         Begin VB.Menu mnuVerPagosDetalle 
            Caption         =   "Pagos/Facturacion"
         End
      End
      Begin VB.Menu VENT 
         Caption         =   "Ventas"
         Begin VB.Menu LisClientes 
            Caption         =   "Clientes"
         End
         Begin VB.Menu mnuLisVentas 
            Caption         =   "Ventas por Cliente"
         End
         Begin VB.Menu mnuLisVentasProd 
            Caption         =   "Ventas por Producto"
         End
         Begin VB.Menu mnuProdCliente 
            Caption         =   "Relaciones Producto Cliente"
         End
         Begin VB.Menu mnuVentasProductoCLiente 
            Caption         =   "Ventas Producto Cliente"
         End
         Begin VB.Menu mnuSaldoClientes 
            Caption         =   "Saldo de Clientes"
         End
         Begin VB.Menu mnuSaldoClientesSinDetalle 
            Caption         =   "Saldo de Clientes sin Detalle"
         End
         Begin VB.Menu mnuMovCtaCliente 
            Caption         =   "Movimiento de Cuenta de Clientes"
         End
         Begin VB.Menu mnuComp 
            Caption         =   "Composicion de clientes"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuCompSin 
            Caption         =   "Composicion de clientes sin detalle"
         End
         Begin VB.Menu mnuLISTIB 
            Caption         =   "Listado de Ventas por Jurisdicción"
         End
         Begin VB.Menu mnuLISTIB2 
            Caption         =   "Listado de Ventas por Jurisdiccion y por Cuentas"
         End
         Begin VB.Menu LISIVAVENTAS 
            Caption         =   "IVA Ventas"
         End
         Begin VB.Menu mnuLisRecibos 
            Caption         =   "Recibos Emitidos"
         End
         Begin VB.Menu mnuRetVta 
            Caption         =   "Retenciones de Clientes"
         End
      End
      Begin VB.Menu mnubol2 
         Caption         =   "Formularios"
      End
      Begin VB.Menu mnuPosiva 
         Caption         =   "Posicion de IVA"
      End
      Begin VB.Menu mnuListPreciosHist 
         Caption         =   "Listado de Precios historico"
      End
      Begin VB.Menu mnuCentro 
         Caption         =   "LISTADO DE CENTRO DE COSTO"
      End
      Begin VB.Menu mnuCodBarras 
         Caption         =   "IMPRESION DE CODIGOS DE BARRA"
      End
      Begin VB.Menu mnuLibroIvaDigital 
         Caption         =   "LIBRO IVA DIGITAL"
      End
   End
   Begin VB.Menu mnuContable 
      Caption         =   "CONTABILIDAD"
      Begin VB.Menu mnuEjercicios 
         Caption         =   "ABM Ejercicios"
      End
      Begin VB.Menu CTAS 
         Caption         =   "ABM Cuentas Contables"
      End
      Begin VB.Menu mnuAbmCuentas 
         Caption         =   "ABM Plan de Cuentas"
      End
      Begin VB.Menu mnuTipoCompras 
         Caption         =   "ABM Parametrizacion CUENTAS"
      End
      Begin VB.Menu mnuPuntos 
         Caption         =   "ABM Puntos de venta"
      End
      Begin VB.Menu mnuLisPlanCuentas 
         Caption         =   "Listar Plan de Cuentas"
      End
      Begin VB.Menu COTIZACIONES 
         Caption         =   "Cotizaciones"
      End
      Begin VB.Menu VERCOTIZACIONES 
         Caption         =   "Ver Cotizaciones"
      End
      Begin VB.Menu MONEDAS 
         Caption         =   "Monedas"
      End
      Begin VB.Menu sepa3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAsientoManual 
         Caption         =   "Carga Asientos Manuales"
      End
      Begin VB.Menu mnuOrdenarAsientos 
         Caption         =   "Reordenar asientos"
      End
      Begin VB.Menu sepa2 
         Caption         =   "-"
      End
      Begin VB.Menu IVAS 
         Caption         =   "IVAS"
      End
      Begin VB.Menu PORCENTAJE 
         Caption         =   "Pordentajes de IVAS"
      End
      Begin VB.Menu mnuCierreIva 
         Caption         =   "Cierre de IVAS"
      End
      Begin VB.Menu sepa1 
         Caption         =   "-"
      End
      Begin VB.Menu LIBRODIARIO 
         Caption         =   "Libro Diario"
      End
      Begin VB.Menu LIBROMAYOR 
         Caption         =   "Libro Mayor"
      End
      Begin VB.Menu SUMASYSALDOS 
         Caption         =   "Sumas y Saldos"
      End
      Begin VB.Menu BALANCE 
         Caption         =   "Balance"
      End
      Begin VB.Menu mnuAjuste0 
         Caption         =   "Ajuste Inflacion - Indices"
      End
      Begin VB.Menu mnuAjuste1 
         Caption         =   "Ajuste Inflacion - Asientos"
      End
   End
   Begin VB.Menu mnuSISTEMA 
      Caption         =   "CONFIGURACIONES"
      Begin VB.Menu USUARIOS 
         Caption         =   "Usuarios"
      End
      Begin VB.Menu mnuEmisor 
         Caption         =   "Emisores"
      End
      Begin VB.Menu CAMBIOCLAVE 
         Caption         =   "Cambiar Clave de Usuario"
      End
      Begin VB.Menu PERMISOS 
         Caption         =   "Permisos"
      End
      Begin VB.Menu TIPOUSUARIO 
         Caption         =   "Tipos de Usuarios"
      End
      Begin VB.Menu mnuDatosE 
         Caption         =   "Datos de la Empresa"
      End
      Begin VB.Menu sepa4 
         Caption         =   "-"
      End
      Begin VB.Menu ZONAS 
         Caption         =   "Zonas"
      End
      Begin VB.Menu mnuLogos 
         Caption         =   "Logo"
      End
      Begin VB.Menu sepa5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPedVenc 
         Caption         =   "SUPERVISION PEDIDOS PENDIENTES"
      End
      Begin VB.Menu mnuPorLis 
         Caption         =   "PORCENTAJES DE LISTAS"
      End
      Begin VB.Menu mnuVigenciaPedidos 
         Caption         =   "VIGENCIA PEDIDOS"
      End
      Begin VB.Menu PRODAUDITORIA 
         Caption         =   "Relacion Producto Auditoria"
      End
      Begin VB.Menu MOTIVOSRECHAZO 
         Caption         =   "Motivos de Rechazos"
      End
      Begin VB.Menu MOTIVOS 
         Caption         =   "Motivos de Ajustes"
      End
      Begin VB.Menu TIPODOC 
         Caption         =   "Tipo de Documentos"
      End
   End
   Begin VB.Menu SALIR 
      Caption         =   "SALIR"
   End
End
Attribute VB_Name = "FrmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '16/9/4
'Private Const sTab = Chr(9)
'*******************************************
Private bApretada As Boolean

Private Sub AbmCostos_Click()
    FrmAbmCentrodeCostos.Show
End Sub

Private Sub BALANCE_Click()
   frmBalance.Show
End Sub

Private Sub chkSinVerificadorActualizacion_Click()
    'uActualizar1.tI
End Sub

Private Sub cmdAgenda_Click()
    frmAgenda.Show
End Sub

Private Sub cmdBackup_Click()
    If ON_ERROR_HABILITADO Then On Error GoTo ufa
    Dim basedato As String
    
    basedato = DataEnvironment1.Sistema.Properties.item("initial catalog")
    
    BackupSql basedato, txtBackup
    'fraBackup.Visible = False
fin:
    Exit Sub
ufa:
    ufa "Fallo el backup", ""
    
    Resume fin
End Sub

Private Sub cmdLogo_Click()
    frmLogos.Show
    mnuSISTEMA.Visible = True
    mnuSISTEMA.enabled = True
End Sub

Private Sub cmdMigrar_Click()
'FrmMigrar.Show
'frmMigrarLi.Show
Dim rsPlan As New ADODB.Recordset, i As Long, sConsul As String
Dim sCUENTA As String, sCODIGO As String, sSUMARIZA As String, sMonetaria As Long, sImputable As Long, sDescripcion As String
rsPlan.Open "select * from cuentaplan order by idd", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
With rsPlan
    If .EOF And .BOF Then
    Else
        .MoveFirst
        For i = 0 To .RecordCount - 1
            sDescripcion = Trim(!DESCRIPCION)
            sCUENTA = Trim(!Cuenta)
            sSUMARIZA = Trim(!Mayoriza)
            If Len(sSUMARIZA) = 1 Then
                sCODIGO = sSUMARIZA
            Else
                'sCODIGO = Replace(sCUENTA, sSUMARIZA, "")
                sCODIGO = CORTO(sCUENTA, Len(sSUMARIZA), 0)
            End If
            sMonetaria = !MONETARIA
            sImputable = !IMPUTABLE
            'sConsul = "update cuentaplan set codigo=" & sstexto(sCODIGO) & " where idd=" & !idd
            sConsul = "INSERT INTO CUENTAS (CUENTA,_CODIGO,DESCRIPCION,IMPUTABLE,SALTO,RENGLON,SUMARIZA,MONETARIA,FECHA_ALTA,USUARIO_ALTA,ACTIVO) VALUES " _
                    & "(" & ssTexto(sCUENTA) & "," & ssTexto(sCODIGO) & "," & ssTexto(sDescripcion) & "," & sImputable & ",0,0," & ssTexto(sSUMARIZA) & "," & sMonetaria & ",'20090805',10,1)"
            'Debug.Print sCUENTA & " " & sCODIGO & " " & sSUMARIZA
            DataEnvironment1.Sistema.Execute sConsul
            
            .MoveNext
        Next
    End If
End With
Set rsPlan = Nothing


End Sub

Private Sub Command1_Click()
'    Dim a
'    che BuscarCliente()
'    che BuscarCliente(a) & vbCrLf & a(0) & vbCrLf & a(1)

    If ON_ERROR_HABILITADO Then On Error GoTo ufaChe


    Dim rs As New ADODB.Recordset
    If confirma("TRASPASA RECIBOS A REC A CUENTA") Then
        
        DE_BeginTrans
        
        With rs
            .Open "select * from recibos ", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
            While Not .EOF
                DataEnvironment1.dbo_abmFacturaVenta "A", _
                    !codigo, "RAA", !numero, !Fecha, !Fecha, 1, 0, _
                    !cliente, " ", 1, " ", 1, 0, 0, 0, 0, !Total, !Total, 0, 0, 0, 100, Date, 1, 1, 0, 0, 0, False, False, 0, 0, 0
                .MoveNext
            Wend
            .Close
        End With
        
        DataEnvironment1.Sistema.Execute " update facturaVenta set saldo = abs(saldo) where saldo < 0 "
        
        DE_CommitTrans
        
        che "listo"
    End If
    
    
fin:
    Exit Sub
ufaChe:
    che "fallo "
    DE_RollbackTrans
    Resume fin
End Sub

'Private Sub Command4_Click()
'    On Error GoTo kaka
''    frmPedidosPendientesBIG.Show
'    'frmLisMovCajasBIG.Show
''    frmLisMovCuentaProv_NEW.Show
'    'frmLisVentas.Show
'
'    'frmVerOPdetalle.Show
''    frmLisVentasGral.Show
''
''    frmVerSeries1.Show
''    frmProductosGrabarComo.Show
'    'arreglarlibraciones
'
''    frmLogos.Show
''
''    If Not confirma("intercambia precios... segura, lau ?") Then Exit Sub
'
'    Dim ss As String
'
'
'    DE_BeginTrans
'
'    ss = "update producto set precio4 = precio "
'    DataEnvironment1.Sistema.Execute ss
'
'    ss = "update producto set precio = precio2 "
'    DataEnvironment1.Sistema.Execute ss
'
'    ss = "update producto set precio2 = precio4 "
'    DataEnvironment1.Sistema.Execute ss
'
'    DE_CommitTrans
'
'
'fin:
'    che "intercambiados"
'    Exit Sub
'
'kaka:
'    DE_RollbackTrans
'    Resume fin
'End Sub
Private Sub arreglarlibraciones()
    Dim rs As New ADODB.Recordset
    Dim s As String, i As Long, ss As String
    
    s = "SELECT m.*, c.CUENTABANCARIA AS cue FROM MoviBanc m INNER JOIN CHQ_COMP c ON m.INTERNO = c.CODIGO WHERE (m.DOCUMENTO = 'P') AND (m.CUENTA = 0) AND (m.OPERACION = 'L') ORDER BY m.INTERNO "
    With rs
        .Open s, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly

        While Not .EOF
            i = i + 1
            ss = "update movibanc set cuenta = " & !cue & "  where movbanco = " & !MovBanco
            DataEnvironment1.Sistema.Execute ss
            .MoveNext
        Wend
    End With
    che "arreglados " & i
    Set rs = Nothing
End Sub
'Private Sub ArregloVencimCompras()
'    Dim ss, rs As ADODB.Recordset
'
'    'ND y NC mosma fecha
'    DataEnvironment1.Sistema.Execute "update "
'
'    '1 desde DOS
'        ss = "select * from _dbf_compras inner join compras " & _
'            " on tipodoc = tipodoc_co, nrodoc= nrodoc_co,  codpr = codpr_co " & _
'            " where "
'
'
'
'
'        rs.Open ss, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'
'
'
'    '2 sistema nuevo
'
'End Sub



Private Sub Command2_Click()
'    Dim rsDebitos As New ADODB.Recordset, rsLibrados As String, librado As Long, movmax As Long
'    rsLibrados = "select * from movibanc where (cuenta= 1 or cuenta =2) and (activo=1) and (operacion='L') AND (FECHA <= '20070822')"
'    rsDebitos.Open rsLibrados, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
'    With rsDebitos
'        If .EOF And .BOF Then
'            MsgBox "No hay movimientos de libracion."
'        Else
'            .MoveFirst
'            For librado = 0 To .RecordCount - 1
'                movmax = 0
'                movmax = obtenerDeSQL("select max(movbanco) from movibanc")
'                movmax = movmax + 1
'
'                DataEnvironment1.Sistema.Execute "INSERT INTO MOVIBANC (CUENTA, OPERACION, DESCRIPCION, FECHA, DOCUMENTO, INTERNO, IMPORTE,TIPDOC, MOVBANCO, ACTIVO, FECHA_ALTA, USUARIO_ALTA, IDDOC)  VALUES(" & !Cuenta & ", 'B', 'Debito de Cheque', '20070823', 'P'," & !interno & ", '" & x2s(!Importe) & "','DEB'," & movmax & ", 1 , " & ssFecha(Date) & "," & UsuarioActual & ", " & !iddoc & " )"
'                .MoveNext
'            Next
'            Command2.caption = "Terminado"
'            Command2.enabled = False
'        End If
'    End With
CrearAsientos
End Sub

Private Sub Command3_Click()
'Shell "c:\windows\system32\cmd.exe md" & App.Path & "\prueba2008\"
'CrearAsientos


    Dim eee
    Dim rspermisos As New ADODB.Recordset, pDescripcion As String, eee2 As Boolean
    Dim agregados As Long, quitados As Long, mm As Object
    agregados = 0
    quitados = 0
    For Each mm In Me
        If TypeOf mm Is Menu Then
            If mm.caption <> "-" Then
                If mm.caption = "SALIR" Then
                Else
                    eee = obtenerDeSQL("select codigo from permisos where descripcion=" & ssTexto(Replace(mm.caption, "&", "")))
                    If IsNull(eee) Or IsEmpty(eee) Then
                        DataEnvironment1.Sistema.Execute "insert into permisos (descripcion,activo) values ('" & Replace(mm.caption, "&", "") & "',1)"
                        agregados = agregados + 1
                    End If
                End If
            End If
        End If
    Next
    
    rspermisos.Open "select * from permisos", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
    
    With rspermisos
        .MoveFirst
        For i = 0 To .RecordCount - 1
            eee2 = False
            pDescripcion = !DESCRIPCION
            For Each mm In Me
                If TypeOf mm Is Menu Then
                    If Replace(mm.caption, "&", "") = pDescripcion Then
                        eee2 = True
                        Exit For
                    End If
                End If
            Next
            
            If Not eee2 Then
                DataEnvironment1.Sistema.Execute "delete from permisos where codigo=" & !codigo
                quitados = quitados + 1
            End If
            .MoveNext
        Next
    End With
    
    
    MsgBox "agregardos: " & agregados & " , quitados: " & quitados

End Sub


Private Sub Command6_Click()
ver_impresoras
End Sub

Private Sub Command4_Click()
Dim rsdat As New ADODB.Recordset
Dim i As Long



rsdat.Open "select * from xlsproductos$", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly

With rsdat

    .MoveFirst
    For i = 0 To .RecordCount - 1

        'DataEnvironment1.dbo_CLIENTE "A", s2n(!codigo), Trim(!DESwCRIPCION), _
            CORTO(sSinNull(!Localidad), 0, Len(sSinNull(!Localidad)) - 4), 0, 0, 1, "", Trim(!direccion), CORTO(sSinNull(!Localidad), 4, 0), _
            Trim(!barrio), !Provincia, _
            !CUIT, 0, !Telefono, Trim(!fax2), "", nSinNull(!Vendedor), _
            nSinNull(!tipoiva), !tipopago, _
            !zona, Trim(sSinNull(!direccion2)), _
            Trim(sSinNull(!localidad2)), Trim(!barrio2), !provincia2, _
            !zona, s2n(!descuento1), _
            s2n(!descuento2), nSinNull(!preciolista), _
            sSinNull(!horario), Trim(!fax2), Trim(!telefono2), _
            "", !categoria, _
            0, 0, _
            0, "", "", 0, 0, _
            0, 0, _
            Date, UsuarioSistema!codigo, 0, 0
        
        
'             DataEnvironment1.dbo_PROVEEDOR "A", Val(Trim(!codigo)), Trim(!DESCRIPCION), _
                Trim(!direccion), CORTO(sSinNull(!Localidad), 5, 0), CORTO(sSinNull(!Localidad), 0, Len(sSinNull(!Localidad)) - 4), _
                !Provincia, sSinNull(!Pais), !CUIT, _
                sSinNull(!Telefono), sSinNull(!Fax), 0, nSinNull(!tipoiva), _
                nSinNull(!sucursal), 1, 0, 0, "N", _
                nSinNull(!tipocompra), 0, _
                "", "", "", 0, 1, 0, 0, 0, "", 0, _
                Date, UsuarioActual(), 0, ""
                
    frmProductos.ABMProducto "A", !codigo2, !codigo2, !a, !b, _
        Trim(!DESCRIPCION), 1, 1, 1, 0, 0, 0, 0, _
        0, 0, 0, 0, 0, 0, 0, "", 0, 0, 0, 0, "", "", 0, "", 0, 0, 0, "", "", 0, 0, 0, "", 0, 0
        
    .MoveNext
    Next
End With
End Sub

Private Sub Command5_Click()
    frmReiniciarSist.Show
End Sub

Private Sub emisiondeetiquetas_Click()
    Rem frmemision.Show
    Rem definir variables en etiquetas
'    datacliente1.Show
    
End Sub

Private Sub ESTADOREMITO_Click()
    FrmRemitosPend.Show
End Sub



Private Sub EXTRACTOSBANCARIOSOLD_Click()
    FrmExtractoBanc.Show
End Sub


Private Sub Form_Initialize()
    Form_Resize
End Sub


Private Sub Form_Load()
    'datos empresa
    CargaParamEmpresa
    lblNombreEmpresa = gEMPR_NombreEmpresa
    lblSucursal = gEMPR_Sucursal
    Me.caption = "Gestion " & gEMPR_NombreEmpresa
    
    'version exe
    RevisaVersion
    
    'frame betasepp
    lblUfa.Visible = UFA_STOP
    lblSabeloquehace = "USUARIO_SABE_LO_QUE_HACE = " & USUARIO_SABE_LO_QUE_HACE
    
'    imagentonka.
    cargarlogo
    
    
    If UsuarioActual() = 19 Then
        modoDacceso ("Modo Contador")
    Else
        modoDacceso ("")
        CargaPermisos
    End If
    
    AverEjercicio
    

    DataEnvironment1.Sistema.Execute "update compras set formadepago=0 where formadepago is null"
    DataEnvironment1.Sistema.Execute "update transcom set formadepago=0 where formadepago is null"
    
    'tiempo
End Sub

Sub CargaPermisos()
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim mm As Object
    Dim descrip As String
    
    'deshabilito TODO
    For Each mm In Me
        If TypeOf mm Is Menu Then
            If mm.caption <> "-" Then mm.enabled = False
            If mm.caption = "SALIR" Then mm.enabled = True
        End If
    Next
    
    rs2.Open "select tipousuario from usuarios where codigo=" & UsuarioActual() & " and activo=1", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    
    'administrador
    
    'If UsuarioSistema!TIPOUSUARIO = 1 Then
    If rs2.EOF And rs2.BOF Then Exit Sub
    
'    DataEnvironment1.Sistema.Execute "delete from permisos"
    If rs2!TIPOUSUARIO = 1 Then
        For Each mm In Me
            If TypeOf mm Is Menu Then
                mm.enabled = True
'                descrip = Replace(Trim(mm.caption), "&", "")
'                DataEnvironment1.Sistema.Execute "insert into permisos (descripcion,activo) values('" & Trim(descrip) & "',1)"
            End If
        Next
        
        If gEMPR_Sucursal = "" Then
            gEMPR_Sucursal = 0 'esto es para que arranque
        End If
        
        If gEMPR_Sucursal = 1 Then
            mnuLoca.enabled = True
            mnuLoca.Visible = True
        Else
            mnuLoca.enabled = False
            mnuLoca.Visible = False
        End If

        Exit Sub
    End If
    'descrip = "select tipousuario from usuarios where codigo=" & UsuarioActual() & " and activo=1"
    
    'rs.Open "select permisosxusuario.*,permisos.descripcion from permisosxusuario inner join permisos on permisosxusuario.permiso=permisos.codigo where grupo=" & UsuarioSistema!TIPOUSUARIO & " and activo=1", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    rs.Open "select permisosxusuario.*,permisos.descripcion from permisosxusuario inner join permisos on permisosxusuario.permiso=permisos.codigo where grupo=" & rs2!TIPOUSUARIO & " and activo=1", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    Do While Not rs.EOF
        For Each mm In Me
            If TypeOf mm Is Menu Then
                descrip = Replace(Trim(mm.caption), "&", "") 'UCase(Replace(Trim(mm.caption), "&", ""))
                If (descrip = rs!DESCRIPCION) Then mm.enabled = True
            End If
        Next
        rs.MoveNext
    Loop
    
    If gEMPR_Sucursal = 1 Then
        mnuLoca.enabled = True
        mnuLoca.Visible = True
    Else
        mnuLoca.enabled = False
        mnuLoca.Visible = False
    End If
    
    rs.Close
    Set rs = Nothing
    Set rs2 = Nothing
End Sub

Private Sub tiempo()
    'supongo que se hace a las 12 del medio dia el restore
    Dim hor
    Dim min
    Dim seg
    Dim tempo
    Dim cTiempo
    Dim a As Long
    tempo = Format(Time, "HH:mm:ss")
    cTiempo = Format("11:59:59 AM", "HH:mm:ss")
    hor = (Hour(cTiempo))
    min = (Minute(cTiempo))
    seg = (Second(cTiempo))
    
    '************* ver la empresa si esta bien escrito!!!!!!!
    
    If gEMPR_NombreEmpresa = "Thores" Then
        If tempo < Format("11:59:59 AM", "HH:mm:ss") Then
            hor = Hour(cTiempo) - Hour(tempo)
            min = Minute(cTiempo) - Minute(tempo)
            seg = Second(cTiempo) - Second(tempo)
            'Ctiempo = Format("11:59:59 AM", "hh:mm:ss") - Format(tempo, "hh:mm:ss")
            hor = (hor * 3600) * 1000
            min = (min * 60) * 1000
            seg = seg * 1000
            If hor < 0 Then hor = 0
            cTiempo = (hor + min + seg) - 7000 'se ejecuta 7 seg antes
            If cTiempo > 60000 Then
                Timer1.Interval = 60000
                cTiempo = cTiempo - 60000
            Else
                
                Timer1.Interval = cTiempo
                cTiempo = 0
            End If
        Else
            Timer1.enabled = False
            Timer1.Interval = 0
        End If
    Else
        Timer1.enabled = False
        Timer1.Interval = 0
    End If
End Sub
Public Sub cargarlogo()
    On Error GoTo fin
    
    imgLogoSimple.Picture = frmLogos.loadLogoSimple()
    FrmPrincipal.Picture = frmLogos.loadLogoFull
fin:

End Sub

Private Sub RevisaVersion()
    On Error Resume Next
    Dim sexe As String
    sexe = obtenerParametro("ServerExe")
    If sexe > "" Then txtCarpetaExe = sexe
    uActualizar1.UbicacionEXE = txtCarpetaExe
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo fin
    FrmKeyPress KeyAscii, True, True, True
    
    If bApretada And KeyAscii = 19 Then '  BS
        If fraBs.Visible = False Then
            fraBs.Visible = True
        Else
            fraBs.Visible = False
        End If
        leo
    End If
    bApretada = (KeyAscii = 2)
fin:
End Sub

Public Sub leo()
    If ON_ERROR_HABILITADO Then On Error GoTo fin
    Dim s
    txtBug = ""
    Open UFA_ARCHIVO_LOG For Input As #1
    Do While Not EOF(1)
        Line Input #1, s
        txtBug = txtBug & s & vbCrLf
    Loop
    Close #1
fin:
End Sub
'*********************************************
'*******************************************

Private Sub DETALLECOMPROBPROV_Click()
    FrmLisComprasDetalle.Show
End Sub

Private Sub EXTRACTOBANCARIO_Click()
    frmExtractoBancNew.Show
End Sub

Private Sub Ajustes_Click()
    FrmAjustes.Show
End Sub

Private Sub AnuloCheques_Click()
    FrmAnulacionCheques.Show
End Sub

Private Sub BANCOGRAL_Click()
    FrmBancosGenerales.Show
End Sub

Private Sub CAJAS_Click()
    FrmCajas.Show
End Sub

Private Sub cajasyFondos_Click()
    FrmIngEgrEfectivo.Show
End Sub

Private Sub CAMBIOCLAVE_Click()
    FrmCambioClave.Show
End Sub

Private Sub CATEGORIAS_Click()
    FrmCategorias.Show
End Sub

Private Sub CLIENTES_Click()
    FormAbmClientes
End Sub

Private Sub cmdLog_Click()
    dEnvTxt
End Sub

Private Sub COMP_Click()
    FrmComprobantes.Show
End Sub



Private Sub CONCEPTOS_Click()
    FrmConceptos.Show
End Sub

Private Sub COTIZACIONES_Click()
    FrmCotizaciones.Show
End Sub

Private Sub CTAS_Click()
    FrmCtasContables.Show
End Sub

Private Sub CuentasBancarias_Click()
    FrmCtasBancarias.Show
End Sub

Private Sub DebitoCheques_Click()
    FrmDebitosCheques.Show
End Sub

Private Sub facturaLibre_Click()
    frmFacturaVenta.mostrar FacturaVenta_Libre
End Sub

Private Sub FacturaPedido_Click()
    frmFacturaVenta.mostrar FacturaVenta_Pedido
    'frmFacturaVenta.mostrar FacturaVenta_Remito
End Sub

Private Sub FacturaRemito_Click()
    frmFacturaVenta.mostrar FacturaVenta_Remito
End Sub

Private Sub facturas_Click()
    FrmFactProv.Show
End Sub

Private Sub Form_Resize()
'    Anclar lblNoCoincide, Me, anclarArriba + anclarDerecha
'    Anclar txtFechaEXE, Me, anclarArriba + anclarDerecha
'    Anclar txtFechaSQL, Me, anclarArriba + anclarDerecha
'    Anclar fraVersiones, Me, anclarArriba + anclarDerecha
    'Anclar uActualizar1, Me, anclarDerecha
    
    'Anclar imgLogoSimple, Me, anclarArriba + anclarIzquierda
    
    
End Sub

Private Sub Form_Unload(cancel As Integer)
    End
End Sub

Private Sub FORMASPAGO_Click()
    FrmFormasPagos.Show
End Sub

Private Sub FORMULAS_Click()
    frmFormulas.Show
End Sub

'Private Sub g_Click()
'    g.ColHidden(2) = Not g.ColHidden(2)
'    g.ColHidden(3) = Not g.ColHidden(3)
'    g.AutoSizeMode = flexAutoSizeRowHeight
'End Sub

Private Sub GastosBancarios_Click()
    FrmGastosBancarios.Show
End Sub


Private Sub GastosBancarios2_Click()
    FrmFactProvBanco.Show
End Sub

Private Sub IngChTerceros_Click()
    FrmIngChequesTerceros.Show
End Sub

Private Sub IngresoChequera_Click()
    FrmIngresoChequera.Show
End Sub

Private Sub IVACOMPRAS_Click()
    FrmLisIvaCompras.Show
End Sub

Private Sub LibracionCheques_Click()
    FrmLibracionCheques.Show
End Sub

Private Sub LIBRODIARIO_Click()
   FrmLibroDiario.Show
End Sub

Private Sub LIBROMAYOR_Click()
   FrmMayorCont.Show
End Sub

Private Sub LisCheques_Click()
    FrmListadoCheques.Show
End Sub

Private Sub LisChTerceros_Click()
    FrmListadoChequesTerceros.Show
End Sub

Private Sub LisClientes_Click()
    FrmLisClientes.Show
End Sub

Private Sub LISIVAVENTAS_Click()
    'FrmLisIvaVentas.Show
    FrmLisIvaVentasNew.Show
End Sub

Private Sub LisMovBancario_Click()
    FrmListadosBancarios.Show
End Sub

Private Sub LisProductos_Click()
    FrmLisProductos.Show
End Sub

Private Sub LisProductos2_Click()
frmLisProductos2.Show
End Sub

Private Sub LisProveedores_Click()
    FrmLisProveedores.Show
End Sub

Private Sub mCancelaRemito_Click()
    frmRemitoVentaCancela.Show
End Sub

Private Sub menuD1_Click()
frmDespacho.Show
End Sub

Private Sub menuD2_Click()
frmDespachoInforme.Show
End Sub

Private Sub MIGRAR_Click()
'    FrmMigrar.Show
    frmMigrarLi.Show
End Sub

Private Sub mnua002_Click()
frmDiferenciaInforme.Show
End Sub

Private Sub mnuAbmCuentas_Click()
'    FrmCtasContables.Show
    frmAbmCuentas.Show
End Sub

Private Sub mnuAcFac_Click()
    frmFactDeta.mostrar FacturaVenta_Libre2
End Sub



Private Sub mnuAjustCosto_Click()
    frmAjusteDeCosto.Show
End Sub

Private Sub mnuAjuste0_Click()
frmAjusteIndices.Show
End Sub

Private Sub mnuAjuste1_Click()
frmAjusteAsiento.Show
End Sub

Private Sub mnuAjusteVenta_Click()
    frmAjusteVentas.Show
End Sub

Private Sub mnuAnularFactVta_Click()
    frmFacturaVentaAnulacion.Show
End Sub

Private Sub mnuAsientoManual_Click()
    frmAsientoManual.Show
End Sub

Private Sub mnuBackup_Click()
    frmBackup.Show
End Sub

Private Sub mnubol2_Click()
frmLisBoletas.Show
End Sub

Private Sub mnuBoletas_Click()
FrmFactProvBoleta.Show
End Sub

'Private Sub mnuBackup_Click()
'    fraBackup.Visible = True
'End Sub

Private Sub mnuCalc_Click()
    'frmCalculin.mostrar
     Shell "c:\windows\system32\calc.exe"
End Sub

Private Sub mnuCancelarPedido_Click()
    frmPedidosCancelacion.Show
End Sub

Private Sub mnuCentro_Click()
    FrmLisCentroCosto.Show
End Sub

Private Sub mnuCierreIva_Click()
    frmCierreIVAS.Show
End Sub

Private Sub mnuCodBarras_Click()
    frmImpresionEtiquetas.Show
End Sub

Private Sub mnuCoef_Click()
    frmActClieProv.Show
End Sub



Private Sub mnuComp_Click()
    frmLismovCli.Show
End Sub

Private Sub mnuCompPerio_Click()
    frmLismovCli3.Show
End Sub

Private Sub mnuCompProv_Click()
    frmLisMovProv2.Show
End Sub

Private Sub mnuCompSin_Click()
    frmLismovCli2.Show
End Sub

Private Sub mnuCont_Click()
    frmContac.Show
End Sub

Private Sub mnuCopiaFormula_Click()
    frmProductosGrabarComo.Show
End Sub

Private Sub mnuCProv2_Click()
    FrmLisProveedores2.Show
End Sub

Private Sub mnuCuentas001_Click()
frmCuentasP.Show
End Sub

Private Sub mnucuentas002_Click()
frmCuentasC.Show
End Sub

Private Sub mnuDatosE_Click()
frmDatosEmpresa.Show
End Sub

Private Sub mnuDiferencia_Click()
    FrmDiferenciaStock.Show
End Sub

Private Sub mnuEjercicios_Click()
    frmAbmEjercicios.Show
End Sub

Private Sub mnuEmisor_Click()
    frmEmisor.Show
End Sub





Private Sub mnuExtractoBancario_Click()
    frmExtractoBancNew.Show
End Sub

Private Sub mnuFactPedidoEXT_Click()
    frmFacturaVenta.mostrar FacturaVenta_Pedido, True
End Sub

'Private Sub mnuFactRemitoEXT_Click()
'    frmFacturaVenta.mostrar FacturaVenta_Remito, True
'End Sub

Private Sub mnuFacturaGastos_Click()
    frmFactProvSoloGastos.Show
End Sub

Private Sub mnuFacturaLibreEXT_Click()
    frmFacturaVenta.mostrar FacturaVenta_Libre, True
End Sub

Private Sub mnuFacturaProvRemito_Click()
'    FrmFactProvRemito.Show
    FrmFactProv.PorRemito
End Sub

Private Sub mnuFormasPagoCompras_Click()
    FrmFormasPagos.Show

End Sub

Private Sub mnuFormulas_Click()
    frmLisFormulas.Show
End Sub

Private Sub mnuGenAsientos_Click()
    frmAsientosGeneracion.Show
End Sub

Private Sub mnuhistorico_Click()
FrmListadoChequesTerceros2.Show
End Sub

Private Sub mnuiibbjur_Click()
FrmLisIngBrutos2.Show
End Sub

Private Sub mnuinfpartes_Click()
frmPartesInforme.Show
End Sub

Private Sub mnuinfpartes2_Click()
frmPartesInformeMovi.Show
End Sub

Private Sub mnukk_Click()
    FrmACTFormaPago.Show
End Sub


Private Sub mnuLeyendaFAE_Click()
    frmLeyendaFAE.Show
End Sub

Private Sub mnuLeyFact_Click()
    frmLeyenfact.Show
End Sub

Private Sub mnuLibroIvaDigital_Click()
frmLibroIvaDigital.Show
End Sub

Private Sub mnuLisOC_Click()
    frmListOrdenCompra.Show
End Sub

Private Sub mnuLisPlanCuentas_Click()
    frmPlandeCtas.Show
End Sub

Private Sub mnuLisRecibos_Click()
    FrmLisRecibos.Show
End Sub

Private Sub mnuLISTIB_Click()
    FrmLisIngBrutos.Show
End Sub

Private Sub mnuLISTIB2_Click()
FrmLisIngBrutos3.Show
End Sub

Private Sub mnuListPerc_Click()
    FrmLisIvaPerc.Show
End Sub

Private Sub mnuListPreciosHist_Click()
    FrmListadoPreciosHistorico.Show
End Sub

Private Sub mnuListProv_Click()
FrmLisProveedores.Show
End Sub

Private Sub mnuLisVentas_Click()
    frmLisVentas.Show
End Sub

Private Sub mnuLisVentasProd_Click()
    frmLisVentasGral.Show
End Sub

Private Sub mnuLoca_Click()
    frmColector.Show
End Sub

Private Sub mnuLogos_Click()
frmLogos.Show
End Sub

Private Sub mnuMercTran_Click()
    frmRemitoCancelacion.Show
End Sub

'Private Sub mnuLogos_Click()
'    frmLogos.Show
'End Sub

Private Sub mnuMovCajasBig_Click()
    frmLisMovCajasBIG.Show
End Sub

Private Sub mnuMovCtaCliente_Click()
    frmLisMovCuentaCli.Show
End Sub

Private Sub mnuMovCtaProv_Click()
    frmLisMovCuentaProv_NEW.Show
End Sub

Private Sub mnuMovStock_Click()
    frmMovimientoStock.Show
End Sub

Private Sub mnuNCDevolucionEXT_Click()
    frmFacturaVenta.mostrar FacturaVenta_NCreditoDevolucion, True
End Sub

Private Sub mnuNCEXT_Click()
    'frmNotaCreDebVenta.mostrar Tipo_NotaCredito, True
    
    frmFacturaVenta.mostrar FacturaVenta_NCredito, True
End Sub

Private Sub mnuNCredito_Click()
    'frmNotaCreDebVenta.mostrar Tipo_NotaCredito
    frmFacturaVenta.mostrar FacturaVenta_NCredito
End Sub

Private Sub mnuNCreditoDevolucion_Click()
    frmFacturaVenta.mostrar FacturaVenta_NCreditoDevolucion
End Sub

Private Sub mnuND_Venta_ChRechazado_Click()
'    frmNotaCreDebVenta.mostrar Tipo_NotaDebitoChRechazado
    frmNotaDcheq.mostrar Tipo_NotaDebitoChRechazado
End Sub

Private Sub mnuNotaDebito_Click()
    'frmNotaCreDebVenta.mostrar Tipo_NotaDebito
    frmFacturaVenta.mostrar FacturaVenta_NDebito
End Sub

Private Sub mnuNumeracion_Click()
    frmNumeracion.Show
End Sub

Private Sub mnuNumeracion2_Click()
    frmNumeracion.Show
End Sub

Private Sub mnuNumeracionC_Click()
    frmNumeracion.Show
End Sub

Private Sub mnuOrdenarAsientos_Click()
    ReordenarAsientos
End Sub

Private Sub mnuParte_Click()
    frmParteProduccion.Show
End Sub

Private Sub mnuPedidosPendienteBig_Click()
    frmPedidosPendientesBIG.Show
End Sub

Private Sub mnuPedPend_Click()
    FrmSaldosPedidos.Show
End Sub

Private Sub mnuPedVenc_Click()
    FrmPedVenc.Show
End Sub

Private Sub mnuPorLis_Click()
    frmPorcenList.Show
End Sub

Private Sub mnuPorte_Click()
frmRemitoPorte.Show
End Sub

Private Sub mnuPosiva_Click()
FrmLisIvaPosicion.Show
End Sub

Private Sub mnuPrecios_Click()
    frmPrecios.Show
End Sub

Private Sub mnuPRemito_Click()
    frmPRemito.Show
End Sub

Private Sub mnuPresupuesto_Click()
    frmPresupuesto.Show
End Sub

Private Sub mnuProdCliente_Click()
    FrmLisRelacionProductoCliente.Show
End Sub

Private Sub mnuProductos_Click()
    frmProductos.Show
End Sub

Private Sub mnuPuntos_Click()
frmABMPuntos.Show
End Sub

Private Sub mnuRecibos_Click()
    frmRecibos.mostrar trRECIBO
End Sub

Private Sub mnuRecibosACuenta_Click()
    frmRecibosACuenta.Show
End Sub

Private Sub mnuRecibosImputacion_Click()
    frmRecibos.mostrar trIMPUTACION
End Sub

Private Sub mnuReemplazoComponente_Click()
    frmProductoRemplazoComponente.Show
End Sub

Private Sub mnuRelProdCliente_Click()
    FrmRelacionProductoCliente.Show
End Sub

Private Sub mnuRetVta_Click()
    frmVerRetVentas.Show
End Sub

Private Sub mnuSalChTer_Click()
    frmLiberarCHtercero.Show
End Sub

Private Sub mnuSaldoClientes_Click()
    frmSaldoCuentaCli.Show
End Sub

Private Sub mnuSaldoClientesSinDetalle_Click()
    FrmLisSaldoCliSinDetalle.Show
End Sub

Private Sub mnuSaldoProvSinDetalle_Click()
    FrmLisSaldoProvSinDetalle.Show
End Sub

Private Sub mnuSaldoRemito_Click()
    frmSaldoRemito.Show
End Sub

Private Sub mnuSeries2_Click()
    FrmSeries.Show
End Sub

Private Sub mnuStkPrecios_Click()
    frmPrecios.Show
End Sub

Private Sub mnuStock2_Click()
    frmMovimientoStock2.Show
End Sub

Private Sub mnuSuc_Click()
    frmsucursal.Show
End Sub



Private Sub mnuTexto_Click()
    frmTexto.Show
End Sub

Private Sub mnuTioProv_Click()
    frmAbmTipoProv.Show
End Sub

Private Sub mnuTipoCompras_Click()
    frmTipoCompras.Show
End Sub

Private Sub mnuTP_Click()
    frmTipoProd.Show
End Sub

Private Sub mnuUnidadesMedida_Click()
    FrmABMUnidadMedida.Show
End Sub

Private Sub mnuUnidadesMedida2_Click()
    FrmABMUnidadTipos.Show
End Sub

Private Sub mnuUnidadesMedida3_Click()
    FrmABMUnidadFactor.Show
End Sub

Private Sub mnuVentasProductoCLiente_Click()
frmVentasProductoCliente.Show
End Sub

Private Sub mnuVerPagos_Click()
    frmVerOP.Show
End Sub

Private Sub mnuVerPagosDetalle_Click()
    frmVerOPdetalle.Show
End Sub

Private Sub mnuVerSeries_Click()
    frmVerSeries1.Show
End Sub

Private Sub mnuVigenciaPedidos_Click()
    FrmVigenciaPed.Show
End Sub

Private Sub MONEDAS_Click()
    FrmMonedas.Show
End Sub

Private Sub MOTIVOSRECHAZO_Click()
    FrmMotivosRechazo.Show
End Sub

Private Sub MovCajas_Click()
    FrmLisCajas.Show
End Sub

Private Sub mS01_Click()
    FrmLisStockSimplificado2.Show
End Sub

Private Sub ms2_Click()
    FrmLisStockSimplificado.Show
End Sub

Private Sub NotaCredito_Click()
    'FrmNotaDeCredito.Show
    frmNotaCredDebCompra.mostrar tipoNotaCompraNC
End Sub

Private Sub NotaDebito_Click()
    'FrmNotaDeDebito.Show
    frmNotaCredDebCompra.mostrar tipoNotaCompraND
End Sub

Private Sub ORDCOMPRA_Click()
    FrmOrdenesdeCompra.Show
End Sub

Private Sub OrdenPago_Click()
    FrmOrdenPago.Show
End Sub

Private Sub PagosACuenta_Click()
    FrmPagosACuenta.Show
End Sub

Private Sub PAISES_Click()
    FrmPaises.Show
End Sub



Private Sub PEDCLI_Click()
    FrmPedidosClientes.Show
End Sub

Private Sub PERMISOS_Click()
    'FrmPermisosdeAcceso.Show
    FrmAccesosAlSistema.Show
End Sub

Private Sub PORCENTAJE_Click()
    FrmPorcentajesIva.Show
End Sub

Private Sub ProcesoCheques_Click()
    FrmProcesoChTerceros.Show
End Sub

Private Sub PRODAUDITORIA_Click()
   FrmRequisitosAud.Show
End Sub

Private Sub PRODCLI_Click()
    FrmRelacionProductoCliente.Show
End Sub

Private Sub PRODUCTOS_Click()
    frmProductos.Show
End Sub

Private Sub RELPRODPROV_Click()
    FrmRelacionProductoProveedor.Show
End Sub

Private Sub RemitoCompra_Click()
    frmRemitoCompra.Show
End Sub

Private Sub RemitoVenta_Click()
    frmRemitoVenta.Show
    frmImagen.Show
End Sub

Private Sub REQUISITOSAUDITRIA_Click()
    FrmRequisitosAuditoria.Show
End Sub

Private Sub salir_Click()
    DataEnvironment1.Sistema.Close
    End
End Sub

Private Sub GRUPOS_Click()
    frmGruposProductos.Show
End Sub

Private Sub IVAS_Click()
    FrmIvas.Show
End Sub

Private Sub MOTIVOS_Click()
    FrmMotivosAjuste.Show
End Sub

Private Sub PROVEEDDORES_Click()
    FrmProveedor.Show
End Sub

Private Sub SERIES_Click()
    frmModSerie.Show
End Sub

Private Sub SlProveedores_Click()
    FrmSaldoCuentaProv.Show
End Sub

Private Sub SUBGRUPOS_Click()
    frmSubgruposProductos.Show
End Sub

Private Sub SUMASYSALDOS_Click()
   FrmSumasySaldos.Show
End Sub

'Private Sub Timer1_Timer()
'    Dim hor As Date
'    Dim min As Date
'    Dim seg As Date
'    Dim cTiempo As Long
'
'    If Format(Time, "HH:mm:ss") < Format("11:59:59 AM", "HH:mm:ss") Then
'        If cTiempo > 60000 Then
'            Timer1.Interval = 60000
'            cTiempo = cTiempo - 60000
'        Else
'            Timer1.Interval = cTiempo
'            If cTiempo = 0 Then
'                frmTiempo.Show
'                MsgBox "Se parara el sistema por un instante para realizar la actualizacion de la base de datos."
'                frmTiempo.Visible = True
'                Dim TiempoPausa, Inicio
'                TiempoPausa = 15    ' Asigna hora de inicio.
'                Inicio = 0
'                Inicio = Timer   ' Establece la hora de inicio.
'                cTiempo = Format(Time, "hh:mm:ss") + Format("00:00:15", "hh:mm:ss")
'                FrmPrincipal.enabled = False
'
'                DataEnvironment1.Sistema.Close
'                Do While Timer < Inicio + TiempoPausa
'                    If frmTiempo.Visible = False Then
'                        frmTiempo.Visible = True
'                    End If
'                Loop
'
'                FrmPrincipal.enabled = True
'                DataEnvironment1.Sistema.Open
'                frmTiempo.Visible = False
'                MsgBox "Ya puede seguir trabajando sobre el sistema, disculpe las molestias."
'            End If
'            cTiempo = 0
'        End If
'    Else
'        Timer1.enabled = False
'        Timer1.Interval = 0
'    End If
'End Sub

'Private Sub Timer1_Timer()
'    RevisaVersion
'End Sub

Private Sub TIPOCOMP_Click()
    FrmTipoCompras2.Show
End Sub

Private Sub TIPODOC_Click()
    FrmTipoDocumentos.Show
End Sub

Private Sub TIPOUSUARIO_Click()
    FrmAbmTiposUsuarios.Show
End Sub

Private Sub TOTTIPOCOMPRA_Click()
    FrmLisTipoCompras.Show
End Sub

Private Sub TransfBancarias_Click()
    FrmTransfBanc.Show
End Sub

Private Sub TRANSPORTES_Click()
    FrmTransportes.Show
End Sub

Private Sub txtCarpetaExe_LostFocus()
    uActualizar1.UbicacionEXE = txtCarpetaExe
End Sub



'Private Sub uNum1_cambio(numero As Double)
'    MsgBox numero & " " & uNum1.Num & " " & uNum1.ssNum & " " & uNum1.suNum
'End Sub

Private Sub USUARIOS_Click()
    FrmAbmUsuarios.Show
End Sub

Private Sub VERCOTIZACIONES_Click()
    frmVerCotizaciones.Show
End Sub

Private Sub lblModo_Click()
    lblModo.Visible = False
End Sub

Private Sub ZONAS_Click()
    FrmZonas.Show
End Sub

Private Sub FormAbmClientes()
    'On Error GoTo ufa
FrmAbmClientes1.Show
    
    'Dim cual As Long
    'cual = 1 'default
    'cual = VerParametro(BS_FORMATO_ABM_CLIENTE)
    
    'Select Case cual
    'Case 1: FrmAbmClientes1.Show
    'Case 2: FrmABMClientes2.Show
    
    'End Select
    
End Sub
    
'    g.rows = 1
'    g.cols = 4
'    g.AddItem "1" & Chr(9) & "p1" & Chr(9) & "1/1/1" & Chr(9) & 10
'    g.AddItem "1" & Chr(9) & "p1" & Chr(9) & "2/1/1" & Chr(9) & 20
'    g.AddItem "1" & Chr(9) & "p1" & Chr(9) & "3/1/1" & Chr(9) & 30
'    g.AddItem "2" & Chr(9) & "p2" & Chr(9) & "1/1/1" & Chr(9) & 10
'    g.AddItem "2" & Chr(9) & "p2" & Chr(9) & "2/1/1" & Chr(9) & 20
'    g.AddItem "3" & Chr(9) & "p3" & Chr(9) & "1/1/1" & Chr(9) & 10
'    g.AddItem "4" & Chr(9) & "p4" & Chr(9) & "1/1/1" & Chr(9) & 10
'    g.AddItem "4" & Chr(9) & "p4" & Chr(9) & "2/1/1" & Chr(9) & 20
'    g.MergeCells = flexMergeFree
'    g.MergeCol(0) = True
'    g.MergeCol(1) = True
'    g.MergeCol(2) = True

' 15/10/4 Li panel BS /migrar /err-log/ se accede con ctl+B ctl+S

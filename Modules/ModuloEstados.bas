Attribute VB_Name = "ModuloEstados"
Option Explicit

' PedidoCliente, RemitoVenta,
Public Const ESTADO_ADEUDADO = "A"
Public Const ESTADO_ENTREGADO = "E"


Public Const Cheque_CARTERA = "C"
'Public Const Cheque_

Public Const Const_PESOS = "Pesos"


'----------------------------------


'''********************************
'''     DOCUMENTACION (deberia estar en otro lado, no?)
'''
'''codigos usados en ch_comp y cheques
'''
'''     IngCompraChequePropio
'''     IngCompraChequeTercero
'''
'''---------- COMPRAS -------------
'''FPROV
'''
'''        Prop    FAC
'''        Terc    FDC
'''
'''O/P
'''
'''        Prop    REC
'''        Terc    O/P
'''
'''PgCta
'''        Prop    RAC
'''        Terc    RAC
'''
'''---------- VENTAS -------------
'''
'''         siempre cheques terceros
'''
'''FV
'''         FAA
'''         FAB
'''
'''Recibo
'''         REC
'''
'''RecCuenta
'''
'''         RAA
'''
'''********************************


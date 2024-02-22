Attribute VB_Name = "ModEmpresaParam"
Option Explicit




Public gEMPR_idEmpresa As Long
Public gEMPR_NombreEmpresa As String
Public gEMPR_ImprimeCertCalidad As BookmarkEnum
'Public gEMPR_FC_CargaCalculaIVA As Boolean
Public gEMPR_Sucursal As String
Public gEMPR_Default_ProductoConSerie As Boolean  '= False ' no usado aun, mod frmProducto cdo pablo termine
Public gEMPR_EmiteFacturaConRemito As Boolean   ' = True     ' frm FacturaVenta
Public gEMPR_EmiteFacturaSobrePedido As Boolean
Public gEMPR_ConSistContable As Boolean
Public gEMPR_Maneja_series As Boolean
Public gEMPR_FormulaEsVirtual As Boolean   ' = False          ' usado en funcion de ModLisistema
                                                    ' False: Baja de Stock Solo Prod Base
                                                    '   remito pone solo prod base
                                                    ' True:  Baja de Stock Solo Prods Componentes
                                                    '   remito descompone en componentes * cantidades
Public gEMPR_DigitosPV As Integer


'Para uso publico
Public Function VerDatoEmpresa(cual As String) As Variant
    VerDatoEmpresa = obtenerDeSQL("select " & cual & " from DatosEmpresa where idEmpresa = " & gEMPR_idEmpresa)
End Function

Public Function siAsiento(campo As String) As Boolean
'AsientoCompras
'AsientoVentas
'AsientoPagos
'AsientoRecibos
siAsiento = obtenerDeSQL("select " & campo & " from DatosEmpresa where idEmpresa = " & gEMPR_idEmpresa)
End Function

' Para uso al abrir programa
Public Function CargaParamEmpresa()
    Dim temp
    
    gEMPR_idEmpresa = obtenerParametro("idEmpresa")
    
    temp = obtenerDeSQL("select " _
        & " Default_ProductoConSerie, EmiteFacturaConRemito, FormulaEsVirtual, 1, Nombre " _
        & " ,ImprimeCertCalidad, ConSistContable " _
        & " ,ManejaSeries,sucursal,digitospv " _
        & " from  DatosEmpresa where idEmpresa = " & gEMPR_idEmpresa)
    gEMPR_Default_ProductoConSerie = temp(0)
    
    gEMPR_EmiteFacturaConRemito = temp(1)
    gEMPR_EmiteFacturaSobrePedido = temp(1) '2 en 1 trucho?
    
    gEMPR_FormulaEsVirtual = temp(2)
'    gEMPR_FC_CargaCalculaIVA = temp(3)
    gEMPR_NombreEmpresa = temp(4)
    gEMPR_ImprimeCertCalidad = temp(5)
    gEMPR_ConSistContable = temp(6)
    gEMPR_Maneja_series = nSinNull(temp(7))
'    gEMP
    If IsNull(temp(8)) Or temp(8) = "" Then
        gEMPR_Sucursal = ""
    Else
        gEMPR_Sucursal = temp(8)
    End If
    gEMPR_DigitosPV = nSinNull(temp(9))
End Function

Public Function TrabaIva(Fecha As Date) As Boolean
    Dim traba As Date
    
    TrabaIva = False
    traba = obtenerDeSQL("select fechatraba from datosempresa where idempresa=" & gEMPR_idEmpresa)
    If traba >= Fecha Then
        TrabaIva = True
    Else
        TrabaIva = False
    End If
End Function

Public Function CuantosDigitosPV() As String
    Dim i As Long
    Dim cuantos As String
    If nSinNull(gEMPR_DigitosPV) > 0 Then
        i = 1
        cuantos = ""
        While i <= gEMPR_DigitosPV
            cuantos = cuantos & "0"
            i = i + 1
        Wend
        CuantosDigitosPV = Trim(cuantos)
    End If
End Function
'daTaenvironment1.Sistema.
'   Provider=SQLOLEDB.1;Password=6459;Persist Security Info=True;User ID=sistema;Initial Catalog=Tonka;Data Source=PABLO



'voy a generar una tabla de parametros distinta a BS
'         datos_empresa
'         en vez de const, variables globales

' Un registro por empresa, asi cada empresa tiene su set de parametrizacion


' **************  TONKA **************************

'Public Const gEMPR_Default_ProductoConSerie = False ' no usado aun, mod frmProducto cdo pablo termine
'Public Const gEMPR_EmiteFacturaConRemito = True     ' frm FacturaVenta
'Public Const gEMPR_FormulaEsVirtual = False          ' usado en funcion de ModLisistema
'                                                    ' False: Baja de Stock Solo Prod Base
'                                                    '   remito pone solo prod base
'                                                    ' True:  Baja de Stock Solo Prods Componentes
'                                                    '   remito descompone en componentes * cantidades
'Public Const gEMPR_FC_CargaCalculaIVA = False

'' **************  LOCAIRE **************************
'Public Const gEMPR_Default_ProductoConSerie = True ' no usado aun, mod frmProducto cdo pablo termine
'Public Const gEMPR_EmiteFacturaConRemito = False     ' frm FacturaVenta
'Public Const gEMPR_FormulaEsVirtual = True          ' usado en funcion de ModLisistema
'                                                    ' False: Baja de Stock Solo Prod Base
'                                                    '   remito pone solo prod base
'                                                    ' True:  Baja de Stock Solo Prods Componentes
'                                                    '   remito descompone en componentes * cantidades





''Public gEMPR_Default_ProductoConSerie As Boolean ' no usado aun, mod frmProducto cdo pablo termine
''Public gEMPR_EmiteFacturaConRemito As Boolean      ' frm FacturaVenta
''Public gEMPR_FormulaEsVirtual As Boolean           ' usado en funcion de ModLisistema
'                                                    ' False: Baja de Stock Solo Prod Base
'                                                    '   remito pone solo prod base
'                                                    ' True:  Baja de Stock Solo Prods Componentes
'                                                    '   remito descompone en componentes * cantidades



'    gEMPR_Default_ProductoConSerie
'    gEMPR_EmiteFacturaConRemito
'    gEMPR_FormulaEsVirtual


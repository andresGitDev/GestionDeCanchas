Attribute VB_Name = "ModuloCierreIVA"
Option Explicit


Public Function PuedoCompras(queFecha As Date, Optional sError As Boolean = True) As Boolean
    Dim dCierre As Date
    Dim dFecha As Date
    dCierre = VerDatoEmpresa("CierreCompras") 'CDate(Format(obtenerParametro("CierreIvaCompras"), "yyyy-mm-01"))
    dFecha = queFecha
    If dFecha > dCierre Then
        PuedoCompras = True
    Else
        PuedoCompras = False
        If sError Then
            Err.Raise 55000, , "periodo iva compras cerrado"
        Else
            MsgBox "Periodo IVA Compras cerrado. Ultimo cierre " & dCierre
        End If
    End If
End Function

Public Function PuedoVentas(queFecha As Date, Optional sError As Boolean = True) As Boolean
    Dim dCierre As Date
    Dim dFecha As Date
    dCierre = VerDatoEmpresa("CierreVentas") 'CDate(Format(obtenerParametro("CierreIvaVentas"), "yyyy-mm-01"))
    dFecha = queFecha
    If dFecha > dCierre Then
        PuedoVentas = True
    Else
        PuedoVentas = False
        If sError Then
            Err.Raise 55000, , "periodo iva Ventas cerrado"
        Else
            MsgBox "Periodo IVA Ventas cerrado. Ultimo cierre " & dCierre, "mm/yyyy"
        End If
    End If
End Function




Attribute VB_Name = "ModuloParametrosEmpresa"
Option Explicit
' 12/6/2006

' Unica funcion Publica = VerParametro(...)

' Quiero centralizar toda parametrizacion dependiente de empresa aca
'        (con el enum me evito errores de programador)

'   cada parametro nuevo que se agrega a la tabla DatosEmpresa
'   se registra en dos lados en este modulo:
'       en el enum BS_                  (visible al programador)
'       en la func StringParametro()    (privado al modulo)
'

Public Enum EnumParametroEmpresa
    BS_CON_CODPROD_COMPUESTO    ' si graba Producto.Codigo al codigo solo o a grupo+subgrupo+codigo
    BS_EXIGE_CARGA_CHEQUERA     ' si carga chequera entera (chq_comp) o cheques individualmente al pagar
    BS_FORMATO_ABM_CLIENTE
    BS_DIRECCION_EMPRESA
    BS_CUIT_EMPRESA
    BS_ComprasConMesImputacion
    BS_FAC_IMPR_MATRIZ
    BS_PREVIEW_IMPRESIONES
    BS_NOMBRE_EMPRESA_CORTO
    BS_CARPETA_BACKUP
    BS_PRINTCOPIASOP
    BS_ALTAPROD_GENERACODEBAR   ' 0, no genera. 1 genera codigo MB.
    BS_MOV_CTACTE_SOLO_MAYORISTAS
    

End Enum

Private Function stringParametro(deCual As EnumParametroEmpresa)
    On Error GoTo ufaChe
    Select Case deCual
      Case BS_CON_CODPROD_COMPUESTO:        stringParametro = "CodProdCompuesto"
      Case BS_EXIGE_CARGA_CHEQUERA:         stringParametro = "ExigeCargaChequera"
      Case BS_FORMATO_ABM_CLIENTE:          stringParametro = "FormatoAbmCliente"
      Case BS_ComprasConMesImputacion:      stringParametro = "ComprasConMesImputacion"
      Case BS_FAC_IMPR_MATRIZ:              stringParametro = "FactImprMatriz"
      Case BS_PREVIEW_IMPRESIONES:          stringParametro = "PreviewImpresiones"
      Case BS_NOMBRE_EMPRESA_CORTO:         stringParametro = "NombreCortoParaListados"
      Case BS_DIRECCION_EMPRESA:            stringParametro = "Direccion"
      Case BS_CUIT_EMPRESA:                 stringParametro = "CuitEmpresa"
      Case BS_CARPETA_BACKUP:               stringParametro = "CarpetaBackupServer"
      Case BS_PRINTCOPIASOP:                stringParametro = "PrintCopiasOP"
      Case BS_ALTAPROD_GENERACODEBAR:       stringParametro = "AltaProdGenCodeBar"
      Case BS_MOV_CTACTE_SOLO_MAYORISTAS:   stringParametro = "CtaCteSoloMayoristas"
      
      Case Else       ' nada
    End Select
Exit Function
ufaChe:
    ufa "", "StrigParametro(): FaltaParametro " & deCual
End Function



Public Function VerParametro(cual As EnumParametroEmpresa)
    VerParametro = obtenerDeSQL("select " & stringParametro(cual) & " from DatosEmpresa as d inner join bs on bs.idEmpresa = d.idEmpresa ")
    If VarType(VerParametro) = vbNull Then
        VerParametro = nSinNull(VerParametro)
    End If
End Function

Attribute VB_Name = "ModuloCuentas"
Option Explicit

' parametros para el ucCode uCuenta
Public Const uCuentaIni1Todas = "select descripcion from cuentas where cuenta = '###' and activo = 1"
Public Const uCuentaIni1Imput = "select descripcion from cuentas where cuenta = '###' and activo = 1 and imputable = 1 "
Public Const uCuentaIni2Todas = "select cuenta as [ Cuenta           ], Descripcion as [ Descripcion                                                    ] from cuentas where activo = 1 order by cuenta "
Public Const uCuentaIni2Imput = "select cuenta as [ Cuenta           ], Descripcion as [ Descripcion                                                    ] from cuentas where activo = 1  and imputable = 1 order by cuenta "

'campo CuentasParam.UsoCuenta
'lo que el usuario puede agregar
' *** NO CAMBIAR los codigos ni el orden *** solo agregar
Public Enum ID_UsoCuenta
    ID_UsoCuenta_SISTEMA = 1  'son los ID_CuentasParam
    ID_UsoCuenta_COMPRAS = 2  'Fact prov , manejado por usuario
    ID_UsoCuenta_VENTAS = 3   'Fact Vta  , idem
    ID_UsoCuenta_RETVTA = 4   'Fact Retenciones Venta
    ID_UsoCuenta_RETCOM = 5   'Fact Retenciones Compras
End Enum
Public Const UsoCuenta_STRING = "Sistema|Fac Prov|Fac Venta|Ret Venta|Ret Prov"

'el codigo cargarlo con el de la Tabla CUENTASPARAM, sistema=1 para que el ABM usario no lo borre


Public Enum ID_CuentasParam
    
    'compras
    ID_Cuenta_C_DEUD_A_PROV = 3
    ID_Cuenta_C_IVA_COMPRA = 2
    ID_Cuenta_C_IVA_COMPRA_RNI = 180
    ID_Cuenta_C_IVA_COMPRA_C = 181
    ID_Cuenta_C_IB_PROV = 174
    ID_Cuenta_C_IB_CAP = 163
    ID_Cuenta_C_RET_GAN_CPRA = 162
    ID_Cuenta_C_RET_IVA_CPRA = 173
    ID_Cuenta_C_RG3337 = 175
    ID_Cuenta_C_RG3431 = 176
    ID_Cuenta_C_IMP_INT = 177
    ID_Cuenta_C_EXENTO = 7
    ID_Cuenta_C_NOGRABADO = 8
    
    
    'GASTOS BANCARIOS
    ID_Cuenta_G_ImpCre = 187
    ID_Cuenta_G_ImpDeb = 188
    ID_Cuenta_G_Sircreb = 189
    ID_Cuenta_G_Sellado = 190
    ID_Cuenta_G_MantCta = 191
    ID_Cuenta_G_MantCtaSueldos = 192
    ID_Cuenta_G_Chequera = 193
    ID_Cuenta_G_Varios = 194
    ID_Cuenta_G_ImpPorSobreGiro = 195
    ID_Cuenta_G_ValNoConformados = 196
    ID_Cuenta_G_PercIIBB = 197
    
    'Pagos
    ID_Cuenta_P_RET_GAN_3ros = 164
    ID_Cuenta_P_RET_IB_3ros = 165
    ID_Cuenta_P_ANTICIP_A_PROV = 6

    
    'misc
    ID_Cuenta_M_CH_CARTERA = 120
    'ID_Cuenta_M_REMUNERACIONES_A_PAGAR = 180
    ID_Cuenta_M_CH_RECHAZADOS = 182
    ID_Cuenta_M_CH_CANJE = 198
    
    'ventas
    ID_Cuenta_V_DEUDxVENTAS = 107
    ID_Cuenta_V_DEUDxVENTAS_EXT = 156
    ID_Cuenta_V_IVA_VENTAS = 110
    ID_Cuenta_V_IVA_VENTAS105 = 111
    ID_Cuenta_V_VENTAS = 160
    ID_Cuenta_V_VENTAS_EXT = 161
    ID_Cuenta_V_Perc_IB_ProvBsAs = 166
    ID_Cuenta_V_INTERESES = 178
    ID_Cuenta_V_DESCUENTO = 179
    
    'Cobro Recibos
    ID_Cuenta_R_ANTICIP_CLIE = 109
    ID_Cuenta_R_ANTICIP_CLIE_EXT = 172
    ID_Cuenta_R_RetSegSoc = 171
    ID_Cuenta_R_IIBB = 170
    ID_Cuenta_R_IIBB_Prov = 183
    ID_Cuenta_R_BONOS_CredFiscal = 169
    ID_Cuenta_R_RET_GAN_RR2784 = 168
    ID_Cuenta_R_RET_IVA_RG3125 = 167 '146
    ID_Cuenta_R_RET_Reparo = 186

End Enum
'

Public Function CuentaParam(idcuentaparam As ID_CuentasParam) As String
    Dim re As String
    re = sSinNull(obtenerDeSQL("select cuenta from CuentasParam where activo = 1 and id = '" & idcuentaparam & "' "))
    If Trim(re) = "" Then ufa "prg err: no se encontro cuenta", "CuentaParam para id " & idcuentaparam
    CuentaParam = re
End Function
Public Function CuentaParamDesc(idcuentaparam As ID_CuentasParam) As String
    Dim re As String
    re = sSinNull(obtenerDeSQL("select descripcion from CuentasParam where activo = 1 and id = '" & idcuentaparam & "' "))
    If re = "" Then ufa "prg err: no se encontro cuenta", "CuentaParamdes para id " & idcuentaparam
    CuentaParamDesc = re
End Function
Public Function CuentaParamxCodigo(codCuentaParam As String) As String
    Dim re As String
    re = sSinNull(obtenerDeSQL("select cuenta from CuentasParam where activo = 1 and codigo= '" & codCuentaParam & "' "))
    If re = "" Then ufa "prg err: no se encontro cuenta", "CuentaParam para cod " & codCuentaParam
    CuentaParamxCodigo = re
End Function

Public Function AsientoBaja(idAsiento) As Boolean
    If ON_ERROR_HABILITADO Then On Error GoTo ufaChe
    Dim re
    re = nSinNull(obtenerDeSQL("select idasiento from asientos where idasiento = '" & idAsiento & "' ")) ' and activo = 1" ))
    If re = 0 Then Exit Function
    
    'DataEnvironment1.Sistema.Execute "delete from mayor where idasiento = '" & idAsiento & "' "
    DataEnvironment1.Sistema.Execute "update mayor set _fecha=" & ssFecha(Date) & ", _ejerc='delete'  where idAsiento = '" & idAsiento & "'"
    DataEnvironment1.Sistema.Execute "update asientos set activo = 0  where idAsiento = '" & idAsiento & "'"
    AsientoBaja = True
fin:
    Exit Function
ufaChe:
    ufa "err al borrar  asiento", "id = " & idAsiento
    Resume fin
End Function

Public Function AsientoBaja_idDoc(iddoc As Long) As Boolean
'    On Error GoTo ufache
    Dim re
    AsientoBaja_idDoc = False
    If iddoc = 0 Then
        MsgBox "No se puede borrar el asiento." & Chr(13) & "No se encontro numero de referencia (idDoc)", vbCritical
        'ufa "err al borrar asiento", "iddoc no especifico"
        AsientoBaja_idDoc = False
        Exit Function
    End If
    re = nSinNull(obtenerDeSQL("select idAsiento from asientos where idDoc = '" & iddoc & "' ")) ' and activo = 1" ))
    If re = 0 Then
        MsgBox "No se puede borrar el asiento." & Chr(13) & "No se encontro numero de referencia (idAsiento)", vbCritical
        AsientoBaja_idDoc = False
        Exit Function
    End If
    DataEnvironment1.Sistema.Execute "update mayor set _fecha=" & ssFecha(Date) & ", _ejerc='delete'  where idAsiento = '" & re & "'"
    DataEnvironment1.Sistema.Execute "update asientos set activo = 0  where idAsiento = '" & re & "'"
    
    'AsientoBaja_idDoc = AsientoBaja(re)
    AsientoBaja_idDoc = True
'fin:
'    Exit Function
'ufaCHE:
'    ufa "err al borrar  asiento", "iddoc = " & iddoc
'    Resume fin
End Function



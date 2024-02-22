Attribute VB_Name = "ModuloCtasBank"
Option Explicit

Public Function verCuentaContableBanco(CuentaBanco As Long) As String
  verCuentaContableBanco = sSinNull(obtenerDeSQL("select cuenta_con from ctasBank where codigo = " & CuentaBanco))
End Function
Public Function verCuentaContableCaja(cuentaCaja As Long) As String
    verCuentaContableCaja = sSinNull(obtenerDeSQL("select cuenta from Cajas where codigo = " & cuentaCaja))
End Function


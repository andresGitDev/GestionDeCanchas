Attribute VB_Name = "ModProcedureGestion"
Option Explicit


Public Function ABM_Cuentas(OPERACION As String, ID As Long, CUENTA As String, CODIGO As String, DESCRIPCION As String, IMPUTABLE As Long, SALTO As Long, MONETARIA As Long, RENGLON As Long, SUMARIZA As Long, FECHAALTA As Date, FECHABAJA As Date, USUALTA As Long, USUBAJA As Long) As Boolean
'If siError Then On Error GoTo Err_17
On Error GoTo Err_17
Dim cadena_iud
cadena_iud = ""
ABM_Cuentas = False
Select Case OPERACION
    Case "A":
        cadena_iud = "INSERT INTO CUENTAS (CUENTA,_CODIGO, DESCRIPCION, IMPUTABLE,SALTO,RENGLON,SUMARIZA,MONETARIA,FECHA_ALTA, USUARIO_ALTA,  ACTIVO) " _
        & " VALUES (" & sstexto(CUENTA) & "," & sstexto(CODIGO) & "," & sstexto(DESCRIPCION) & "," & IMPUTABLE & "," & SALTO & "," & RENGLON & "," & SUMARIZA & "," & MONETARIA _
        & "," & ssFecha(FECHAALTA) & "," & USUALTA & ", 1)"
    Case "M":
        cadena_iud = "UPDATE CUENTAS SET CUENTA=" & sstexto(CUENTA) & ", DESCRIPCION=" & sstexto(DESCRIPCION) & ", IMPUTABLE=" & IMPUTABLE & ",SALTO= " & SALTO & ",RENGLON= " & RENGLON & ",SUMARIZA=" & sstexto(SUMARIZA) & ", MONETARIA=" & MONETARIA & " WHERE ID=" & ID
    Case "B":
        cadena_iud = "UPDATE CUENTAS  SET ACTIVO=0, FECHA_BAJA=" & ssFecha(FECHABAJA) & ", USUARIO_BAJA= " & USUBAJA & " WHERE ID=" & ID
End Select
'If siGestion Then
    DataEnvironment1.Sistema.Execute cadena_iud
'Else
'    DataEnvironment1.BASE.Execute cadena_iud
'End If
ABM_Cuentas = True
Exit Function
Err_17:
ABM_Cuentas = False
End Function

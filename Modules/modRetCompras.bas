Attribute VB_Name = "modRetCompras"
Option Explicit
'
Public Function CalculaRetGan(prov As Long, monto As Double, fechaperiodo As Date) As Double
'    ' Busca pagos del mes, retenciones hechas y hace el calculo de proxima retencion
'
'    Dim rangoFecha As String, sWhere As String
'    Dim Anio, Mes, tempo
'    Dim baseimp As Double, coef As Double, tipo As Integer
'    Dim SumPago As Double, SumRetPago As Double
'
'    ' tipoprov
'    tempo = obtenerDeSQL("select TipoProv, Baseimponible, coeficiente from prov inner join tipoprov on prov.tipoprov = tipoprov.codigo where prov.codigo = '" & prov & "' and prov.activo = 1 and tipoprov.activo = 1 ")
'    If IsEmpty(tempo) Then
'        che "No se encontraron datos de tipoprov para "
'        CalculaRetGan = 0
'        Exit Function
'    Else
'        tipo = s2n(tempo(0))
'        baseimp = s2n(tempo(1))
'        coef = s2n(tempo(2), 6)
'    End If
'
'    'where consulta
'    Anio = Year(fechaperiodo)
'    Mes = Month(fechaperiodo)
'    rangoFecha = " fecha between '" & Format(DateSerial(Anio, Mes, 1), "yyyymmdd") & "' and '" & Format(DateSerial(Anio, Mes + 1, 0), "yyyymmdd") & "' "
'    sWhere = " " & rangoFecha & " and activo = 1 and codpr = '" & prov & "' "
'
'
'    'compras FAC contado
'    tempo = obtenerDeSQL(" select sum(total) as SumTotal, sum(RET_GAN  ) as SumRetGan " & _
'        " from compras where tipodoc = 'FAC' and contado = 1 and " & sWhere)
'    'If Not IsEmpty(tempo) Then
'        SumPago = SumPago + s2n(tempo(0))
'        SumRetPago = SumRetPago + s2n(tempo(1))
'    'End If
'
'    'compras RAC
'    tempo = obtenerDeSQL("select sum(total) as SumTotal, sum(RET_GAN  ) as SumRetGan " & _
'        " from compras where tipodoc = 'RAC' and " & sWhere)
''    If Not IsEmpty(tempo) Then
'        SumPago = SumPago + s2n(tempo(0))
'        SumRetPago = SumRetPago + s2n(tempo(1))
''    End If
'
'    'transcom RAC
'    tempo = obtenerDeSQL("select sum(total) as SumTotal, sum(RET_GAN  ) as SumRetGan " & _
'        " from transcom where tipodoc = 'RAC' and " & sWhere)
''    If Not IsEmpty(tempo) Then
'        SumPago = SumPago + s2n(tempo(0))
'        SumRetPago = SumRetPago + s2n(tempo(1))
''    End If
'
'    'rec_com
'    tempo = obtenerDeSQL("select sum(total) as SumTotal, sum(RETGAN  ) as SumRetGan " & _
'        " from rec_comp where " & sWhere)
''    If Not IsEmpty(tempo) Then
'        SumPago = SumPago + s2n(tempo(0))
'        SumRetPago = SumRetPago + s2n(tempo(1))
''    End If
'
'
'    'tempo = s2n(obtenerdesql("select sum(importe) as sumRetGan from comprasretenciones inner join  "))
'
'
'    If tipo = 1 Then
'        CalculaRetGan = CalculaH() 'prov, monto, fechaperiodo, baseimp, coef)
'    Else
'        CalculaRetGan = calculaL(monto, baseimp, coef, SumPago, SumRetPago)
'    End If
'
End Function
'
'Private Function calculaL(monto, baseimp, coef, SumPago, SumRetPago) As Double
'    Dim tmpRet As Double
'
'    If SumPago + monto > baseimp Then
'        tmpRet = (SumPago + monto - baseimp) * coef
'        calculaL = s2n(tmpRet - SumRetPago)
'    End If
'End Function
'
'Private Function CalculaH() As Double
'
'End Function

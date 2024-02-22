Attribute VB_Name = "Bacigaluppi"
Public Function aAGRI(Optional ss As String = "", Optional dd As Long = 0, Optional ff As Double, Optional nn As Long = 0, Optional ppcc As Long = 0)
Dim rr As Long, bDel As String
'With FrmPrincipal.gSinAsientos
'    bDel = Chr(9) & Chr(9)
'    If ss > "" Then
'        .AddItem ""
'        rr = .rows - 1
'        .TextMatrix(rr, 0) = ss
'        .TextMatrix(rr, 1) = dd
'        .TextMatrix(rr, 2) = ff
'        .TextMatrix(rr, 3) = nn
'        .TextMatrix(rr, 4) = ppcc
'        Print #1, ss & bDel & dd & bDel & ff & bDel & nn & bDel & ppcc
'    Else
'        .clear
'        .cols = 0
'        .rows = 1
'        .cols = 5
'        .TextMatrix(0, 0) = "DOCUMENTO        "
'        .TextMatrix(0, 1) = "ID               "
'        .TextMatrix(0, 2) = "DIFERENCIA       "
'        .TextMatrix(0, 3) = "NROASIENTO       "
'        .TextMatrix(0, 4) = "PROVEEDOR-CLIENTE"
'        .ColWidth(0) = 1500
'        .ColWidth(1) = 1500
'        .ColWidth(2) = 1500
'        .ColWidth(3) = 1500
'        .ColWidth(4) = 1500
'        Open "C:\ASIENTOS_MARZO_ABRIL.TXT" For Output As #1
'        Print #1, Trim(.TextMatrix(0, 0)) & bDel & Trim(.TextMatrix(0, 1)) & bDel & Trim(.TextMatrix(0, 2)) & bDel & Trim(.TextMatrix(0, 3)) & bDel & Trim(.TextMatrix(0, 4))
'    End If
'    'Close #1
'End With
End Function

Public Function CrearAsientos()
'Shell "c:\windows\system32\cmd.exe md" & App.Path & "\prueba2008\"
Dim sFechaD As Date, sFechaH As Date
Dim rsDoc As New ADODB.Recordset, sDoc As String, i As Long
Dim aDoc As New Asiento, txtAsiento As String, tdoc As String
Dim aContador As Long
sFechaD = CDate("03/01/2011")
sFechaH = CDate("03/01/2011")
aAGRI
'PAGOS
'orden de pago
'pago a cuenta
aContador = 0
'********************VENTAS***********************FACTURAS,NOTAS C Y D,RECIBOS A CTA
GoTo ir:
sDoc = "select * from facturaventa where activo=1 and fecha>=" & ssFecha(sFechaD) & " and fecha<=" & ssFecha(sFechaH)
With rsDoc
    .Open sDoc, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
    If .EOF And .BOF Then
    Else
        .MoveFirst
        For i = 0 To .RecordCount - 1
            tdoc = CORTO(Trim(!TIPODOC), 0, 1)

            Select Case tdoc
                Case "NC":
                    txtAsiento = "M-NC " & !NroFactura
                    aDoc.nuevo "NC " & !RAZONSOCIAL, CDate(!Fecha), "NCv"
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_V_IVA_VENTAS), s2n(!Iva), 0, txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_V_Perc_IB_ProvBsAs), s2n(!IIBB), 0, txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_V_VENTAS), (s2n(!Neto)), 0, txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_V_DEUDxVENTAS), 0, !Total, txtAsiento

                Case "FA":
                    txtAsiento = "M-FAV " & !NroFactura
                    aDoc.nuevo "FV " & !RAZONSOCIAL, CDate(!Fecha), "FAV"
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_V_IVA_VENTAS), 0, s2n(!Iva), txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_V_Perc_IB_ProvBsAs), 0, s2n(!IIBB), txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_V_VENTAS), 0, s2n(!Neto), txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_V_DEUDxVENTAS), s2n(!Total), 0, txtAsiento

                Case "ND":
                    txtAsiento = "M-ND " & !NroFactura
                    aDoc.nuevo "ND " & !RAZONSOCIAL, CDate(!Fecha), "NDv"
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_V_IVA_VENTAS), 0, s2n(!Iva), txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_V_Perc_IB_ProvBsAs), 0, s2n(!IIBB), txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_V_VENTAS), 0, (s2n(!Neto)), txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_V_DEUDxVENTAS), !Total, 0, txtAsiento

                Case "RA":
                    txtAsiento = "M-REC " & !NroFactura
                    aDoc.nuevo "Rec " & !RAZONSOCIAL, CDate(!Fecha), "RECV"
                    aDoc.AcumularItem "11010002", s2n(!Total), 0, txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_R_ANTICIP_CLIE), 0, s2n(!Total), txtAsiento

            End Select
            If Not exAsiento(s2n(!iddoc)) Then
                If aDoc.Grabar(s2n(!iddoc)) = 0 Then
                    aAGRI txtAsiento, s2n(!iddoc), aDoc.Diferencia
                Else
                    aAGRI txtAsiento, s2n(!iddoc), aDoc.Diferencia, obtenerDeSQL("select nroasiento from asientos where activo=1 and  iddoc=" & s2n(!iddoc)), s2n(!cliente)
                End If
                aContador = aContador + 1
            End If
            Set aDoc = Nothing
            
            .MoveNext
        Next
    End If
    Set rsDoc = Nothing
End With

'************************************COBROS**********RECIBOS
sDoc = "select * from RECIBOS where activo=1 and fecha>=" & ssFecha(sFechaD) & " and fecha<=" & ssFecha(sFechaH)
With rsDoc
    .Open sDoc, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
    If .EOF And .BOF Then
    Else
        .MoveFirst
        For i = 0 To .RecordCount - 1
            tdoc = Trim(!TIPODOC)
            
            Select Case tdoc
                Case "REC":
                    txtAsiento = "M-REC " & !numero
                    aDoc.nuevo "Rec " & obtenerDeSQL("select descripcion from clientes where codigo=" & !cliente), CDate(!Fecha), "NCv"
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_R_ANTICIP_CLIE), (s2n(!Total)), 0, txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_V_DEUDxVENTAS), 0, s2n(!Total), txtAsiento
                    
            End Select
            If Not exAsiento(s2n(!iddoc)) Then
                If aDoc.Grabar(s2n(!iddoc)) = 0 Then
                    aAGRI txtAsiento, s2n(!iddoc), aDoc.Diferencia
                Else
                    aAGRI txtAsiento, s2n(!iddoc), aDoc.Diferencia, obtenerDeSQL("select nroasiento from asientos where activo=1 and  iddoc=" & s2n(!iddoc)), s2n(!cliente)
                End If
                aContador = aContador + 1
            End If
            Set aDoc = Nothing

            .MoveNext
        Next
    End If
    Set rsDoc = Nothing
End With
ir:
'************************************COMPRAS**********FACTURAS, NOTAS DE C Y D, PAGOS A CTA
'sDoc = "select * from COMPRAS where activo=1 and fecha>=" & ssFecha(sFechaD) & " and fecha<=" & ssFecha(sFechaH)
sDoc = "SELECT     iddoc,codpr,tipodoc,nrodoc,fecha,codpr,razonsocialprov,total,neto,iva_21,percepc,ibprovincia,iva_27,iva_9,iva_10,imp_int,retgan,nogravado,ibcapital,exento,retganpago,ibpago FROM         COMPRAS where activo=1 and fecha_alta>=" & ssFecha(sFechaD) & " and fecha_alta<=" & ssFecha(sFechaH) & " " _
     & " Union " _
     & " SELECT    iddoc,codpr,tipodoc,nrodoc,fecha,codpr,razonsocialprov,total,neto,iva_21,percepc,ibprovincia,iva_27,iva_9,iva_10,imp_int,retgan,nogravado,ibcapital,exento,retganpago,ibpago FROM      transcom  where activo=1 and fecha_alta>=" & ssFecha(sFechaD) & " and fecha_alta<=" & ssFecha(sFechaH) & " " _
     & " order by tipodoc"
With rsDoc
    .Open sDoc, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
    If .EOF And .BOF Then
    Else
        .MoveFirst
        For i = 0 To .RecordCount - 1
            tdoc = Trim(!TIPODOC)
            
            Select Case tdoc
                Case "FAC":
                    txtAsiento = "M-FC " & !NroDoc
                    aDoc.nuevo "Fac " & !razonsocialprov, CDate(!Fecha), "FAC"
                    'aDoc.AcumularItem CuentaParam(ID_Cuenta_C_EXENTO), s2n(!EXENTO), 0, txtAsiento
                    'aDoc.AcumularItem CuentaParam(ID_Cuenta_C_NOGRABADO), s2n(!NoGravado), 0, txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_C_IVA_COMPRA), s2n(!IVA_21), 0, txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_C_IVA_COMPRA_RNI), s2n(!iva_10), 0, txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_C_IVA_COMPRA_C), s2n(!IVA_27), 0, txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_C_IB_CAP), s2n(!ibcapital), 0, txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_C_IB_PROV), s2n(!ibprovincia), 0, txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_C_RET_GAN_CPRA), s2n(!retgan), 0, txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_C_RG3337), s2n(!percepc), 0, txtAsiento
                    aDoc.AcumularItem CtaProv(s2n(!CODPR)), s2n(s2n(!Neto) + s2n(!nogravado) + s2n(!EXENTO)), 0, txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_C_DEUD_A_PROV), 0, s2n(!Total), txtAsiento
                    
                Case "N/C":
                    txtAsiento = "M-N/C " & !NroDoc
                    aDoc.nuevo "N/C " & !razonsocialprov, CDate(!Fecha), "N/C"
                    'aDoc.AcumularItem CuentaParam(ID_Cuenta_C_EXENTO), 0, s2n(!EXENTO), txtAsiento
                    'aDoc.AcumularItem CuentaParam(ID_Cuenta_C_NOGRABADO), 0, s2n(!NoGravado), txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_C_IVA_COMPRA), 0, s2n(!IVA_21), txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_C_IVA_COMPRA_RNI), 0, s2n(!iva_10), txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_C_IVA_COMPRA_C), 0, s2n(!IVA_27), txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_C_IB_CAP), 0, s2n(!ibcapital), txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_C_IB_PROV), 0, s2n(!ibprovincia), txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_C_RET_GAN_CPRA), 0, s2n(!retgan), txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_C_RG3337), 0, s2n(!percepc), txtAsiento
                    aDoc.AcumularItem CtaProv(s2n(!CODPR)), 0, s2n(s2n(!Neto) + s2n(!nogravado) + s2n(!EXENTO)), txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_C_DEUD_A_PROV), s2n(!Total), 0, txtAsiento
                Case "N/D":
                    txtAsiento = "M-N/D " & !NroDoc
                    aDoc.nuevo "N/D " & !razonsocialprov, CDate(!Fecha), "N/D"
                    'aDoc.AcumularItem CuentaParam(ID_Cuenta_C_EXENTO), s2n(!EXENTO) , 0, txtAsiento
                    'aDoc.AcumularItem CuentaParam(ID_Cuenta_C_NOGRABADO), s2n(!NoGravado), 0, txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_C_IVA_COMPRA), s2n(!IVA_21), 0, txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_C_IVA_COMPRA_RNI), s2n(!iva_10), 0, txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_C_IVA_COMPRA_C), s2n(!IVA_27), 0, txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_C_IB_CAP), s2n(!ibcapital), 0, txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_C_IB_PROV), s2n(!ibprovincia), 0, txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_C_RET_GAN_CPRA), s2n(!retgan), 0, txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_C_RG3337), s2n(!percepc), 0, txtAsiento
                    aDoc.AcumularItem CtaProv(s2n(!CODPR)), s2n(s2n(!Neto) + s2n(!nogravado) + s2n(!EXENTO)), 0, txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_C_DEUD_A_PROV), 0, s2n(!Total), txtAsiento
                                
                Case "RAC":
                    txtAsiento = "M-RAC " & !NroDoc
                    aDoc.nuevo "RAC " & !razonsocialprov, CDate(!Fecha), "PAC"
                    aDoc.AgregarItem CuentaParam(ID_Cuenta_P_RET_GAN_3ros), 0, s2n(!retganpago), txtAsiento
                    aDoc.AgregarItem CuentaParam(ID_Cuenta_P_RET_IB_3ros), 0, s2n(!IBPAGO), txtAsiento
                    aDoc.AcumularItem "11010002", 0, s2n(s2n(!Total) - s2n(!retganpago) - s2n(!IBPAGO)), txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_C_DEUD_A_PROV), s2n(!Total), 0, txtAsiento
                    
            End Select
            If Not exAsiento(s2n(!iddoc)) Then
                If aDoc.Grabar(s2n(!iddoc)) = 0 Then
                    aAGRI txtAsiento, s2n(!iddoc), aDoc.Diferencia
                Else
                    aAGRI txtAsiento, s2n(!iddoc), aDoc.Diferencia, obtenerDeSQL("select nroasiento from asientos where activo=1 and  iddoc=" & s2n(!iddoc)), s2n(!CODPR)
                End If
                aContador = aContador + 1
            End If
            Set aDoc = Nothing

            .MoveNext
        Next
    End If
    Set rsDoc = Nothing
End With

'************************************PAGOS**********ORDEN DE PAGO
sDoc = "select * from REC_COMP where activo=1 and fecha_alta>=" & ssFecha(sFechaD) & " and fecha_alta<=" & ssFecha(sFechaH)
With rsDoc
    .Open sDoc, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
    If .EOF And .BOF Then
    Else
        .MoveFirst
        For i = 0 To .RecordCount - 1
            tdoc = "RAC" 'Trim(!TIPODOC)
            
            Select Case tdoc
               Case "RAC":
                    txtAsiento = "M-OP " & !Nro
                    aDoc.nuevo "OP " & obtenerDeSQL("select descripcion from prov where codigo=" & !CODPR), CDate(!Fecha), "O/P"
                    aDoc.AgregarItem CuentaParam(ID_Cuenta_P_RET_GAN_3ros), 0, s2n(!retganpago), txtAsiento
                    aDoc.AgregarItem CuentaParam(ID_Cuenta_P_RET_IB_3ros), 0, s2n(!IBPAGO), txtAsiento
                    aDoc.AcumularItem "11010002", 0, s2n(s2n(!Total) - s2n(!IBPAGO) - s2n(!retganpago)), txtAsiento
                    aDoc.AcumularItem CuentaParam(ID_Cuenta_C_DEUD_A_PROV), s2n(!Total), 0, txtAsiento
                 
            End Select
            If Not exAsiento(s2n(!iddoc)) Then
                If aDoc.Grabar(s2n(!iddoc)) = 0 Then
                    aAGRI txtAsiento, s2n(!iddoc), aDoc.Diferencia
                Else
                    aAGRI txtAsiento, s2n(!iddoc), aDoc.Diferencia, obtenerDeSQL("select nroasiento from asientos where activo=1 and iddoc=" & s2n(!iddoc)), s2n(!CODPR)
                End If
                aContador = aContador + 1
            End If
            Set aDoc = Nothing

            .MoveNext
        Next
    End If
    Set rsDoc = Nothing
End With
FrmPrincipal.Label1.caption = aContador & " ASIENTOS"
End Function

Public Function exAsiento(aIdDOC As Long) As Boolean
Dim esta
esta = obtenerDeSQL("select nroasiento from asientos where activo=1 and iddoc=" & aIdDOC)
If IsNull(esta) Or IsEmpty(esta) Then
    exAsiento = False
Else
    exAsiento = True
End If
End Function

Public Function CtaProv(CodProv) As String
Dim pCta As String, pCtas As String, C As String, CT '42060002

pCta = "42060002"
pCtas = sSinNull(obtenerDeSQL("select cuentascompras from prov where codigo=" & CodProv))

C = Trim(pCtas)
    If C > "" Then
        CT = Split(Replace(C, "#", ""), ",")
        CtaProv = CT(0)
    Else
        CtaProv = pCta
    End If


End Function

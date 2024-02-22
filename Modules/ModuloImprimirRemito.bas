Attribute VB_Name = "ModuloImprimirRemito"
Option Explicit

Public remiCarta As Long

'************************ REMITO ****************************
'******************cabecera de pagina*********************************
Public Sub seniorRemito() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionRemitoVenta.lblcliente.Left = PasarTamano(cargarPos("senor", "REMITO", "posx"))
    RptImpresionRemitoVenta.lblcliente.Top = PasarTamano(cargarPos("senor", "REMITO", "posy"))
    RptImpresionRemitoVenta.lblcliente.Visible = cargarPos("senor", "REMITO", "visible")
    a = PasarTamano(cargarPos("senor", "REMITO", "largo"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.lblcliente.Width = a
    End If
    a = PasarTamano(cargarPos("senor", "REMITO", "ancho"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.lblcliente.Height = a
    End If
    a = cargarColor("senor", "REMITO", "alineahorizontal")
    If Not a = "" Then
        If Not IsNull(a) Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionRemitoVenta.lblcliente.Alignment = a
        End If
    End If
    a = cargarPos("senor", "REMITO", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then
            RptImpresionRemitoVenta.lblcliente.Font.Size = a
        End If
    End If
    a = cargarColor("senor", "REMITO", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionRemitoVenta.lblcliente.BackColor = a
        End If
    End If
    a = cargarColor("senor", "REMITO", "backstyle")
    If Not a = "" Then
        If Not IsNull(a) Then
            If Not a = 0 Then RptImpresionRemitoVenta.lblcliente.BackStyle = PasoColor(a)
        End If
    End If
End Sub
Public Sub direccionRemito() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionRemitoVenta.lbldomicilio.Left = PasarTamano(cargarPos("domicilio", "REMITO", "posx"))
    RptImpresionRemitoVenta.lbldomicilio.Top = PasarTamano(cargarPos("domicilio", "REMITO", "posy"))
    RptImpresionRemitoVenta.lbldomicilio.Visible = cargarPos("domicilio", "REMITO", "visible")
    a = PasarTamano(cargarPos("domicilio", "REMITO", "largo"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.lbldomicilio.Width = a
    End If
    a = PasarTamano(cargarPos("domicilio", "REMITO", "ancho"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.lbldomicilio.Height = a
    End If
    a = cargarColor("domicilio", "REMITO", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionRemitoVenta.lbldomicilio.Alignment = a
        End If
    End If
    a = cargarPos("domicilio", "REMITO", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionRemitoVenta.lbldomicilio.Font.Size = a
    End If
    a = cargarColor("domicilio", "REMITO", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionRemitoVenta.lbldomicilio.BackColor = a
        End If
    End If
    a = cargarColor("domicilio", "REMITO", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionRemitoVenta.lbldomicilio.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub CUITremito() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionRemitoVenta.lblcuit.Left = PasarTamano(cargarPos("cuit", "REMITO", "posx"))
    RptImpresionRemitoVenta.lblcuit.Top = PasarTamano(cargarPos("cuit", "REMITO", "posy"))
    RptImpresionRemitoVenta.lblcuit.Visible = cargarPos("cuit", "REMITO", "visible")
    a = PasarTamano(cargarPos("cuit", "REMITO", "largo"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.lblcuit.Width = a
    End If
    a = PasarTamano(cargarPos("cuit", "REMITO", "ancho"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.lblcuit.Height = a
    End If
    a = cargarColor("cuit", "REMITO", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionRemitoVenta.lblcuit.Alignment = a
        End If
    End If
    a = cargarPos("cuit", "REMITO", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionRemitoVenta.lblcuit.Font.Size = a
    End If
    a = cargarColor("cuit", "REMITO", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionRemitoVenta.lblcuit.BackColor = a
        End If
    End If
    a = cargarColor("cuit", "REMITO", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionRemitoVenta.lblcuit.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub DiaRemito() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionRemitoVenta.Dia.Left = PasarTamano(cargarPos("dia", "REMITO", "posx"))
    RptImpresionRemitoVenta.Dia.Top = PasarTamano(cargarPos("dia", "REMITO", "posy"))
    RptImpresionRemitoVenta.Dia.Visible = cargarPos("dia", "REMITO", "visible")
    a = PasarTamano(cargarPos("dia", "REMITO", "largo"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.Dia.Width = a
    End If
    a = PasarTamano(cargarPos("dia", "REMITO", "ancho"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.Dia.Height = a
    End If
    a = cargarColor("dia", "REMITO", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionRemitoVenta.Dia.Alignment = a
        End If
    End If
    a = cargarPos("dia", "REMITO", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionRemitoVenta.Dia.Font.Size = a
    End If
    a = cargarColor("dia", "REMITO", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionRemitoVenta.Dia.BackColor = a
        End If
    End If
    a = cargarColor("dia", "REMITO", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionRemitoVenta.Dia.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub MesRemito() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionRemitoVenta.Mes.Left = PasarTamano(cargarPos("mes", "REMITO", "posx"))
    RptImpresionRemitoVenta.Mes.Top = PasarTamano(cargarPos("mes", "REMITO", "posy"))
    RptImpresionRemitoVenta.Mes.Visible = cargarPos("mes", "REMITO", "visible")
    a = PasarTamano(cargarPos("mes", "REMITO", "largo"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.Mes.Width = a
    End If
    a = PasarTamano(cargarPos("mes", "REMITO", "ancho"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.Mes.Height = a
    End If
    a = cargarColor("mes", "REMITO", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionRemitoVenta.Mes.Alignment = a
        End If
    End If
    a = cargarPos("mes", "REMITO", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionRemitoVenta.Mes.Font.Size = a
    End If
    a = cargarColor("mes", "REMITO", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionRemitoVenta.Mes.BackColor = a
        End If
    End If
    a = cargarColor("mes", "REMITO", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionRemitoVenta.Mes.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub AnoRemito() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionRemitoVenta.Ano.Left = PasarTamano(cargarPos("ano", "REMITO", "posx"))
    RptImpresionRemitoVenta.Ano.Top = PasarTamano(cargarPos("ano", "REMITO", "posy"))
    RptImpresionRemitoVenta.Ano.Visible = cargarPos("ano", "REMITO", "visible")
    a = PasarTamano(cargarPos("ano", "REMITO", "largo"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.Ano.Width = a
    End If
    a = PasarTamano(cargarPos("ano", "REMITO", "ancho"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.Ano.Height = a
    End If
    a = cargarColor("ano", "REMITO", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionRemitoVenta.Ano.Alignment = a
        End If
    End If
    a = cargarPos("ano", "REMITO", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionRemitoVenta.Ano.Font.Size = a
    End If
    a = cargarColor("ano", "REMITO", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionRemitoVenta.Ano.BackColor = a
        End If
    End If
    a = cargarColor("ano", "REMITO", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionRemitoVenta.Ano.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub MesesRemito()
    Dim a As Variant
    RptImpresionRemitoVenta.Meses.Left = PasarTamano(cargarPos("meses", "REMITO", "posx"))
    RptImpresionRemitoVenta.Meses.Top = PasarTamano(cargarPos("meses", "REMITO", "posy"))
    RptImpresionRemitoVenta.Meses.Visible = cargarPos("meses", "REMITO", "visible")
    a = PasarTamano(cargarPos("meses", "REMITO", "largo"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.Meses.Width = a
    End If
    a = PasarTamano(cargarPos("meses", "REMITO", "ancho"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.Meses.Height = a
    End If
    a = cargarColor("meses", "REMITO", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionRemitoVenta.Meses.Alignment = a
        End If
    End If
    a = cargarPos("meses", "REMITO", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionRemitoVenta.Meses.Font.Size = a
    End If
    a = cargarColor("meses", "REMITO", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionRemitoVenta.Meses.BackColor = a
        End If
    End If
    a = cargarColor("meses", "REMITO", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionRemitoVenta.Meses.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub Transporte() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionRemitoVenta.Transportista.Left = PasarTamano(cargarPos("transpo", "REMITO", "posx"))
    RptImpresionRemitoVenta.Transportista.Top = PasarTamano(cargarPos("transpo", "REMITO", "posy"))
    RptImpresionRemitoVenta.Transportista.Visible = cargarPos("transpo", "REMITO", "visible")
    a = PasarTamano(cargarPos("transpo", "REMITO", "largo"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.Transportista.Width = a
    End If
    a = PasarTamano(cargarPos("transpo", "REMITO", "ancho"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.Transportista.Height = a
    End If
    a = cargarColor("transpo", "REMITO", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionRemitoVenta.Transportista.Alignment = a
        End If
    End If
    a = cargarPos("transpo", "REMITO", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionRemitoVenta.Transportista.Font.Size = a
    End If
    a = cargarColor("transpo", "REMITO", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionRemitoVenta.Transportista.BackColor = a
        End If
    End If
    a = cargarColor("transpo", "REMITO", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionRemitoVenta.Transportista.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub presupuestoRemito() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionRemitoVenta.Presupu.Left = PasarTamano(cargarPos("Presupu", "REMITO", "posx"))
    RptImpresionRemitoVenta.Presupu.Top = PasarTamano(cargarPos("Presupu", "REMITO", "posy"))
    RptImpresionRemitoVenta.Presupu.Visible = cargarPos("Presupu", "REMITO", "visible")
    a = PasarTamano(cargarPos("Presupu", "REMITO", "largo"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.Presupu.Width = a
    End If
    a = PasarTamano(cargarPos("Presupu", "REMITO", "ancho"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.Presupu.Height = a
    End If
    a = cargarColor("Presupu", "REMITO", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionRemitoVenta.Presupu.Alignment = a
        End If
    End If
    a = cargarPos("Presupu", "REMITO", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionRemitoVenta.Presupu.Font.Size = a
    End If
    a = cargarColor("Presupu", "REMITO", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionRemitoVenta.Presupu.BackColor = a
        End If
    End If
    a = cargarColor("Presupu", "REMITO", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionRemitoVenta.Presupu.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub ordenCompraRemito() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionRemitoVenta.OrdenComp.Left = PasarTamano(cargarPos("OrdenComp", "REMITO", "posx"))
    RptImpresionRemitoVenta.OrdenComp.Top = PasarTamano(cargarPos("OrdenComp", "REMITO", "posy"))
    RptImpresionRemitoVenta.OrdenComp.Visible = cargarPos("OrdenComp", "REMITO", "visible")
    a = PasarTamano(cargarPos("OrdenComp", "REMITO", "largo"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.OrdenComp.Width = a
    End If
    a = PasarTamano(cargarPos("OrdenComp", "REMITO", "ancho"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.OrdenComp.Height = a
    End If
    a = cargarColor("OrdenComp", "REMITO", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionRemitoVenta.OrdenComp.Alignment = a
        End If
    End If
    a = cargarPos("OrdenComp", "REMITO", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionRemitoVenta.OrdenComp.Font.Size = a
    End If
    a = cargarColor("OrdenComp", "REMITO", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionRemitoVenta.OrdenComp.BackColor = a
        End If
    End If
    a = cargarColor("OrdenComp", "REMITO", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionRemitoVenta.OrdenComp.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub AtencionRemito() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionRemitoVenta.lblAtencion.Left = PasarTamano(cargarPos("lblAtencion", "REMITO", "posx"))
    RptImpresionRemitoVenta.lblAtencion.Top = PasarTamano(cargarPos("lblAtencion", "REMITO", "posy"))
    RptImpresionRemitoVenta.lblAtencion.Visible = cargarPos("lblAtencion", "REMITO", "visible")
    a = PasarTamano(cargarPos("lblAtencion", "REMITO", "largo"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.lblAtencion.Width = a
    End If
    a = PasarTamano(cargarPos("lblAtencion", "REMITO", "ancho"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.lblAtencion.Height = a
    End If
    a = cargarColor("lblAtencion", "REMITO", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionRemitoVenta.lblAtencion.Alignment = a
        End If
    End If
    a = cargarPos("lblAtencion", "REMITO", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionRemitoVenta.lblAtencion.Font.Size = a
    End If
    a = cargarColor("lblAtencion", "REMITO", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionRemitoVenta.lblAtencion.BackColor = a
        End If
    End If
    a = cargarColor("lblAtencion", "REMITO", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionRemitoVenta.lblAtencion.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub LocalidadRemito() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionRemitoVenta.lbllocalidad.Left = PasarTamano(cargarPos("lbllocalidad", "REMITO", "posx"))
    RptImpresionRemitoVenta.lbllocalidad.Top = PasarTamano(cargarPos("lbllocalidad", "REMITO", "posy"))
    RptImpresionRemitoVenta.lbllocalidad.Visible = cargarPos("lbllocalidad", "REMITO", "visible")
    a = PasarTamano(cargarPos("lbllocalidad", "REMITO", "largo"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.lbllocalidad.Width = a
    End If
    a = PasarTamano(cargarPos("lbllocalidad", "REMITO", "ancho"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.lbllocalidad.Height = a
    End If
    a = cargarColor("lbllocalidad", "REMITO", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionRemitoVenta.lbllocalidad.Alignment = a
        End If
    End If
    a = cargarPos("lbllocalidad", "REMITO", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionRemitoVenta.lbllocalidad.Font.Size = a
    End If
    a = cargarColor("lbllocalidad", "REMITO", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionRemitoVenta.lbllocalidad.BackColor = a
        End If
    End If
    a = cargarColor("lbllocalidad", "REMITO", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionRemitoVenta.lbllocalidad.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub FacturaRemito() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionRemitoVenta.lblfactura.Left = PasarTamano(cargarPos("lblfactura", "REMITO", "posx"))
    RptImpresionRemitoVenta.lblfactura.Top = PasarTamano(cargarPos("lblfactura", "REMITO", "posy"))
    RptImpresionRemitoVenta.lblfactura.Visible = cargarPos("lblfactura", "REMITO", "visible")
    a = PasarTamano(cargarPos("lblfactura", "REMITO", "largo"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.lblfactura.Width = a
    End If
    a = PasarTamano(cargarPos("lblfactura", "REMITO", "ancho"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.lblfactura.Height = a
    End If
    a = cargarColor("lblfactura", "REMITO", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionRemitoVenta.lblfactura.Alignment = a
        End If
    End If
    a = cargarPos("lblfactura", "REMITO", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionRemitoVenta.lblfactura.Font.Size = a
    End If
    a = cargarColor("lblfactura", "REMITO", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionRemitoVenta.lblfactura.BackColor = a
        End If
    End If
    a = cargarColor("lblfactura", "REMITO", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionRemitoVenta.lblfactura.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub FechaRemito() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionRemitoVenta.lblfecha.Left = PasarTamano(cargarPos("lblfecha", "REMITO", "posx"))
    RptImpresionRemitoVenta.lblfecha.Top = PasarTamano(cargarPos("lblfecha", "REMITO", "posy"))
    RptImpresionRemitoVenta.lblfecha.Visible = cargarPos("lblfecha", "REMITO", "visible")
    a = PasarTamano(cargarPos("lblfecha", "REMITO", "largo"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.lblfecha.Width = a
    End If
    a = PasarTamano(cargarPos("lblfecha", "REMITO", "ancho"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.lblfecha.Height = a
    End If
    a = cargarColor("lblfecha", "REMITO", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionRemitoVenta.lblfecha.Alignment = a
        End If
    End If
    a = cargarPos("lblfecha", "REMITO", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionRemitoVenta.lblfecha.Font.Size = a
    End If
    a = cargarColor("lblfecha", "REMITO", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionRemitoVenta.lblfecha.BackColor = a
        End If
    End If
    a = cargarColor("lblfecha", "REMITO", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionRemitoVenta.lblfecha.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub ComprobanteRemito() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionRemitoVenta.lblcomp.Left = PasarTamano(cargarPos("lblcomp", "REMITO", "posx"))
    RptImpresionRemitoVenta.lblcomp.Top = PasarTamano(cargarPos("lblcomp", "REMITO", "posy"))
    RptImpresionRemitoVenta.lblcomp.Visible = cargarPos("lblcomp", "REMITO", "visible")
    a = PasarTamano(cargarPos("lblcomp", "REMITO", "largo"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.lblcomp.Width = a
    End If
    a = PasarTamano(cargarPos("lblcomp", "REMITO", "ancho"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.lblcomp.Height = a
    End If
    a = cargarColor("lblcomp", "REMITO", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionRemitoVenta.lblcomp.Alignment = a
        End If
    End If
    a = cargarPos("lblcomp", "REMITO", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionRemitoVenta.lblcomp.Font.Size = a
    End If
    a = cargarColor("lblcomp", "REMITO", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionRemitoVenta.lblcomp.BackColor = a
        End If
    End If
    a = cargarColor("lblcomp", "REMITO", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionRemitoVenta.lblcomp.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub IvaRemito() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionRemitoVenta.LblIva.Left = PasarTamano(cargarPos("LblIva", "REMITO", "posx"))
    RptImpresionRemitoVenta.LblIva.Top = PasarTamano(cargarPos("LblIva", "REMITO", "posy"))
    RptImpresionRemitoVenta.LblIva.Visible = cargarPos("LblIva", "REMITO", "visible")
    a = PasarTamano(cargarPos("LblIva", "REMITO", "largo"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.LblIva.Width = a
    End If
    a = PasarTamano(cargarPos("LblIva", "REMITO", "ancho"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.LblIva.Height = a
    End If
    a = cargarColor("LblIva", "REMITO", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionRemitoVenta.LblIva.Alignment = a
        End If
    End If
    a = cargarPos("LblIva", "REMITO", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionRemitoVenta.LblIva.Font.Size = a
    End If
    a = cargarColor("LblIva", "REMITO", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionRemitoVenta.LblIva.BackColor = a
        End If
    End If
    a = cargarColor("LblIva", "REMITO", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionRemitoVenta.LblIva.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub ReferenciaRemito() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionRemitoVenta.lblref.Left = PasarTamano(cargarPos("lblref", "REMITO", "posx"))
    RptImpresionRemitoVenta.lblref.Top = PasarTamano(cargarPos("lblref", "REMITO", "posy"))
    RptImpresionRemitoVenta.lblref.Visible = cargarPos("lblref", "REMITO", "visible")
    a = PasarTamano(cargarPos("lblref", "REMITO", "largo"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.lblref.Width = a
    End If
    a = PasarTamano(cargarPos("lblref", "REMITO", "ancho"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.lblref.Height = a
    End If
    a = cargarColor("lblref", "REMITO", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionRemitoVenta.lblref.Alignment = a
        End If
    End If
    a = cargarPos("lblref", "REMITO", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionRemitoVenta.lblref.Font.Size = a
    End If
    a = cargarColor("lblref", "REMITO", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionRemitoVenta.lblref.BackColor = a
        End If
    End If
    a = cargarColor("lblref", "REMITO", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionRemitoVenta.lblref.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub NroReferenciaRemito() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionRemitoVenta.lblnroref.Left = PasarTamano(cargarPos("lblnroref", "REMITO", "posx"))
    RptImpresionRemitoVenta.lblnroref.Top = PasarTamano(cargarPos("lblnroref", "REMITO", "posy"))
    RptImpresionRemitoVenta.lblnroref.Visible = cargarPos("lblnroref", "REMITO", "visible")
    a = PasarTamano(cargarPos("lblnroref", "REMITO", "largo"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.lblnroref.Width = a
    End If
    a = PasarTamano(cargarPos("lblnroref", "REMITO", "ancho"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.lblnroref.Height = a
    End If
    a = cargarColor("lblnroref", "REMITO", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionRemitoVenta.lblnroref.Alignment = a
        End If
    End If
    a = cargarPos("lblnroref", "REMITO", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionRemitoVenta.lblnroref.Font.Size = a
    End If
    a = cargarColor("lblnroref", "REMITO", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionRemitoVenta.lblnroref.BackColor = a
        End If
    End If
    a = cargarColor("lblnroref", "REMITO", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionRemitoVenta.lblnroref.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub NroProvinciaRemito() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionRemitoVenta.lblNroProv.Left = PasarTamano(cargarPos("lblNroProv", "REMITO", "posx"))
    RptImpresionRemitoVenta.lblNroProv.Top = PasarTamano(cargarPos("lblNroProv", "REMITO", "posy"))
    RptImpresionRemitoVenta.lblNroProv.Visible = cargarPos("lblNroProv", "REMITO", "visible")
    a = PasarTamano(cargarPos("lblNroProv", "REMITO", "largo"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.lblNroProv.Width = a
    End If
    a = PasarTamano(cargarPos("lblNroProv", "REMITO", "ancho"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.lblNroProv.Height = a
    End If
    a = cargarColor("lblNroProv", "REMITO", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionRemitoVenta.lblNroProv.Alignment = a
        End If
    End If
    a = cargarPos("lblNroProv", "REMITO", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionRemitoVenta.lblNroProv.Font.Size = a
    End If
    a = cargarColor("lblNroProv", "REMITO", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionRemitoVenta.lblNroProv.BackColor = a
        End If
    End If
    a = cargarColor("lblNroProv", "REMITO", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionRemitoVenta.lblNroProv.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub TacharRemito() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionRemitoVenta.lbltachar.Left = PasarTamano(cargarPos("lbltachar", "REMITO", "posx"))
    RptImpresionRemitoVenta.lbltachar.Top = PasarTamano(cargarPos("lbltachar", "REMITO", "posy"))
    RptImpresionRemitoVenta.lbltachar.Visible = cargarPos("lbltachar", "REMITO", "visible")
    a = PasarTamano(cargarPos("lbltachar", "REMITO", "largo"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.lbltachar.Width = a
    End If
    a = PasarTamano(cargarPos("lbltachar", "REMITO", "ancho"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.lbltachar.Height = a
    End If
    a = cargarColor("lbltachar", "REMITO", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionRemitoVenta.lbltachar.Alignment = a
        End If
    End If
    a = cargarPos("lbltachar", "REMITO", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionRemitoVenta.lbltachar.Font.Size = a
    End If
    a = cargarColor("lbltachar", "REMITO", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionRemitoVenta.lbltachar.BackColor = a
        End If
    End If
    a = cargarColor("lbltachar", "REMITO", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionRemitoVenta.lbltachar.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub

'*********************** detalle **************************
Public Sub cantidadRemito() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionRemitoVenta.txtcantidad.Left = PasarTamano(cargarPos("lblCantidad", "REMITO", "posx"))
    RptImpresionRemitoVenta.txtcantidad.Top = PasarTamano(cargarPos("lblCantidad", "REMITO", "posy"))
    RptImpresionRemitoVenta.txtcantidad.Visible = cargarPos("lblCantidad", "REMITO", "visible")
    a = PasarTamano(cargarPos("lblCantidad", "REMITO", "largo"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.txtcantidad.Width = a
    End If
    a = PasarTamano(cargarPos("lblCantidad", "REMITO", "ancho"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.txtcantidad.Height = a
    End If
    a = cargarColor("lblCantidad", "REMITO", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionRemitoVenta.txtcantidad.Alignment = a
        End If
    End If
    a = cargarPos("lblCantidad", "REMITO", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionRemitoVenta.txtcantidad.Font.Size = a
    End If
    a = cargarColor("lblCantidad", "REMITO", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionRemitoVenta.txtcantidad.BackColor = a
        End If
    End If
    a = cargarColor("lblCantidad", "REMITO", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionRemitoVenta.txtcantidad.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub ArticuloRemito() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionRemitoVenta.txtdescripcion.Left = PasarTamano(cargarPos("lblArticulo", "REMITO", "posx"))
    RptImpresionRemitoVenta.txtdescripcion.Top = PasarTamano(cargarPos("lblArticulo", "REMITO", "posy"))
    RptImpresionRemitoVenta.txtdescripcion.Visible = cargarPos("lblArticulo", "REMITO", "visible")
    a = PasarTamano(cargarPos("lblArticulo", "REMITO", "largo"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.txtdescripcion.Width = a
    End If
    a = PasarTamano(cargarPos("lblArticulo", "REMITO", "ancho"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.txtdescripcion.Height = a
    End If
    a = cargarColor("lblArticulo", "REMITO", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionRemitoVenta.txtdescripcion.Alignment = a
        End If
    End If
    a = cargarPos("lblArticulo", "REMITO", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionRemitoVenta.txtdescripcion.Font.Size = a
    End If
    a = cargarColor("lblArticulo", "REMITO", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionRemitoVenta.txtdescripcion.BackColor = a
        End If
    End If
    a = cargarColor("lblArticulo", "REMITO", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionRemitoVenta.txtdescripcion.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub DescripcionRemito() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionRemitoVenta.Field1.Left = PasarTamano(cargarPos("lblDescripcion", "REMITO", "posx"))
    RptImpresionRemitoVenta.Field1.Top = PasarTamano(cargarPos("lblDescripcion", "REMITO", "posy"))
    RptImpresionRemitoVenta.Field1.Visible = cargarPos("lblDescripcion", "REMITO", "visible")
    a = PasarTamano(cargarPos("lblDescripcion", "REMITO", "largo"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.Field1.Width = a
    End If
    a = PasarTamano(cargarPos("lblDescripcion", "REMITO", "ancho"))
    If Not a = 0 Then
        RptImpresionRemitoVenta.Field1.Height = a
    End If
    a = cargarColor("lblDescripcion", "REMITO", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionRemitoVenta.Field1.Alignment = a
        End If
    End If
    a = cargarPos("lblDescripcion", "REMITO", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionRemitoVenta.Field1.Font.Size = a
    End If
    a = cargarColor("lblDescripcion", "REMITO", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionRemitoVenta.Field1.BackColor = a
        End If
    End If
    a = cargarColor("lblDescripcion", "REMITO", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionRemitoVenta.Field1.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub

'***********************************************************


'Public Sub UnitarioRemito() ' el tamaño de letra toma un maximo de 16
'    Dim a As Variant
'    RptImpresionRemitoVenta.lblPrecUnitario.Left = PasarTamano(cargarPos("lblPrecUnitario", "REMITO", "posx"))
'    rptRemito.lblPrecUnitario.Top = PasarTamano(cargarPos("lblPrecUnitario", "REMITO", "posy"))
'    rptRemito.lblPrecUnitario.Visible = cargarPos("lblPrecUnitario", "REMITO", "visible")
'    a = PasarTamano(cargarPos("lblPrecUnitario", "REMITO", "largo"))
'    If Not a = 0 Then
'        rptRemito.lblPrecUnitario.Width = a
'    End If
'    a = PasarTamano(cargarPos("lblPrecUnitario", "REMITO", "ancho"))
'    If Not a = 0 Then
'        rptRemito.lblPrecUnitario.Height = a
'    End If
'    a = cargarColor("lblPrecUnitario", "REMITO", "alineahorizontal")
'    If Not IsNull(a) Then
'        If Not a = "" Then
'            If a = "Derecha" Then
'                a = ddTXRight
'            ElseIf a = "Izquierda" Then
'                a = ddTXLeft
'            ElseIf a = "Centro" Then
'                a = ddTXCenter
'            End If
'            rptRemito.lblPrecUnitario.Alignment = a
'        End If
'    End If
'    a = cargarPos("lblPrecUnitario", "REMITO", "sizeletra")
'    If Not IsNull(a) Then
'        If Not a = 0 Then rptRemito.lblPrecUnitario.Font.Size = a
'    End If
'    a = cargarColor("lblPrecUnitario", "REMITO", "backcolor")
'    If Not IsNull(a) Then
'        If Not a = "" Then
'            If Not a = 0 Then rptRemito.lblPrecUnitario.BackColor = a
'        End If
'    End If
'    a = cargarColor("lblPrecUnitario", "REMITO", "backstyle")
'    If Not IsNull(a) Then
'        If Not a = "" Then
'            If Not a = 0 Then
'                rptRemito.lblPrecUnitario.BackStyle = PasoColor(a)
'            End If
'        End If
'    End If
'End Sub
'Public Sub TotalRemito() ' el tamaño de letra toma un maximo de 16
'    Dim a As Variant
'    rptRemito.lblPrecTotal.Left = PasarTamano(cargarPos("lblPrecTotal", "REMITO", "posx"))
'    rptRemito.lblPrecTotal.Top = PasarTamano(cargarPos("lblPrecTotal", "REMITO", "posy"))
'    rptRemito.lblPrecTotal.Visible = cargarPos("lblPrecTotal", "REMITO", "visible")
'    a = PasarTamano(cargarPos("lblPrecTotal", "REMITO", "largo"))
'    If Not a = 0 Then
'        rptRemito.lblPrecTotal.Width = a
'    End If
'    a = PasarTamano(cargarPos("lblPrecTotal", "REMITO", "ancho"))
'    If Not a = 0 Then
'        rptRemito.lblPrecTotal.Height = a
'    End If
'    a = cargarColor("lblPrecTotal", "REMITO", "alineahorizontal")
'    If Not IsNull(a) Then
'        If Not a = "" Then
'            If a = "Derecha" Then
'                a = ddTXRight
'            ElseIf a = "Izquierda" Then
'                a = ddTXLeft
'            ElseIf a = "Centro" Then
'                a = ddTXCenter
'            End If
'            rptRemito.lblPrecTotal.Alignment = a
'        End If
'    End If
'    a = cargarPos("lblPrecTotal", "REMITO", "sizeletra")
'    If Not IsNull(a) Then
'        If Not a = 0 Then rptRemito.lblPrecTotal.Font.Size = a
'    End If
'    a = cargarColor("lblPrecTotal", "REMITO", "backcolor")
'    If Not IsNull(a) Then
'        If Not a = "" Then
'            If Not a = 0 Then rptRemito.lblPrecTotal.BackColor = a
'        End If
'    End If
'    a = cargarColor("lblPrecTotal", "REMITO", "backstyle")
'    If Not IsNull(a) Then
'        If Not a = "" Then
'            If Not a = 0 Then
'                rptRemito.lblPrecTotal.BackStyle = PasoColor(a)
'            End If
'        End If
'    End If
'End Sub

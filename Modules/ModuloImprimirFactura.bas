Attribute VB_Name = "ModuloImprimirFactura"
Option Explicit

Public leye As Integer

'*************************** FACTURA ***************************
'******************cabecera de pagina*********************************
Public Sub senior() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.lblcliente.Left = PasarTamano(cargarPos("senor", "FACTURA", "posx"))
    RptImpresionFacturaVenta.lblcliente.Top = PasarTamano(cargarPos("senor", "FACTURA", "posy"))
    RptImpresionFacturaVenta.lblcliente.Visible = cargarPos("senor", "FACTURA", "visible")
    a = PasarTamano(cargarPos("senor", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.lblcliente.Width = a
    End If
    a = PasarTamano(cargarPos("senor", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.lblcliente.Height = a
    End If
    a = cargarColor("senor", "FACTURA", "alineahorizontal")
    If Not a = "" Then
        If Not IsNull(a) Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.lblcliente.Alignment = a
        End If
    End If
    a = cargarPos("senor", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.lblcliente.Font.Size = a
        End If
    End If
    a = cargarColor("senor", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.lblcliente.BackColor = a
    End If
    a = cargarColor("senor", "FACTURA", "backstyle")
    If Not a = "" Then
        If Not IsNull(a) Then
            If Not a = 0 Then RptImpresionFacturaVenta.lblcliente.BackStyle = PasoColor(a)
        End If
    End If
End Sub
Public Sub direccion() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.lbldomicilio.Left = PasarTamano(cargarPos("domicilio", "FACTURA", "posx"))
    RptImpresionFacturaVenta.lbldomicilio.Top = PasarTamano(cargarPos("domicilio", "FACTURA", "posy"))
    RptImpresionFacturaVenta.lbldomicilio.Visible = cargarPos("domicilio", "FACTURA", "visible")
    a = PasarTamano(cargarPos("domicilio", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.lbldomicilio.Width = a
    End If
    a = PasarTamano(cargarPos("domicilio", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.lbldomicilio.Height = a
    End If
    a = cargarColor("domicilio", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.lbldomicilio.Alignment = a
        End If
    End If
    a = cargarPos("domicilio", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.lbldomicilio.Font.Size = a
    End If
    a = cargarColor("domicilio", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.lbldomicilio.BackColor = a
        End If
    End If
    a = cargarColor("domicilio", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.lbldomicilio.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub CUIT() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.lblcuit.Left = PasarTamano(cargarPos("cuit", "FACTURA", "posx"))
    RptImpresionFacturaVenta.lblcuit.Top = PasarTamano(cargarPos("cuit", "FACTURA", "posy"))
    RptImpresionFacturaVenta.lblcuit.Visible = cargarPos("cuit", "FACTURA", "visible")
    a = PasarTamano(cargarPos("cuit", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.lblcuit.Width = a
    End If
    a = PasarTamano(cargarPos("cuit", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.lblcuit.Height = a
    End If
    a = cargarColor("cuit", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.lblcuit.Alignment = a
        End If
    End If
    a = cargarPos("cuit", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.lblcuit.Font.Size = a
    End If
    a = cargarColor("cuit", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.lblcuit.BackColor = a
        End If
    End If
    a = cargarColor("cuit", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.lblcuit.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub Dia() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.Dia.Left = PasarTamano(cargarPos("dia", "FACTURA", "posx"))
    RptImpresionFacturaVenta.Dia.Top = PasarTamano(cargarPos("dia", "FACTURA", "posy"))
    RptImpresionFacturaVenta.Dia.Visible = cargarPos("dia", "FACTURA", "visible")
    a = PasarTamano(cargarPos("dia", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.Dia.Width = a
    End If
    a = PasarTamano(cargarPos("dia", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.Dia.Height = a
    End If
    a = cargarColor("dia", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.Dia.Alignment = a
        End If
    End If
    a = cargarPos("dia", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.Dia.Font.Size = a
    End If
    a = cargarColor("dia", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.Dia.BackColor = a
        End If
    End If
    a = cargarColor("dia", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.Dia.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub Mes() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.Mes.Left = PasarTamano(cargarPos("mes", "FACTURA", "posx"))
    RptImpresionFacturaVenta.Mes.Top = PasarTamano(cargarPos("mes", "FACTURA", "posy"))
    RptImpresionFacturaVenta.Mes.Visible = cargarPos("mes", "FACTURA", "visible")
    a = PasarTamano(cargarPos("mes", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.Mes.Width = a
    End If
    a = PasarTamano(cargarPos("mes", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.Mes.Height = a
    End If
    a = cargarColor("mes", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.Mes.Alignment = a
        End If
    End If
    a = cargarPos("mes", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.Mes.Font.Size = a
    End If
    a = cargarColor("mes", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.Mes.BackColor = a
        End If
    End If
    a = cargarColor("mes", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.Mes.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub Ano() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.Ano.Left = PasarTamano(cargarPos("ano", "FACTURA", "posx"))
    RptImpresionFacturaVenta.Ano.Top = PasarTamano(cargarPos("ano", "FACTURA", "posy"))
    RptImpresionFacturaVenta.Ano.Visible = cargarPos("ano", "FACTURA", "visible")
    a = PasarTamano(cargarPos("ano", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.Ano.Width = a
    End If
    a = PasarTamano(cargarPos("ano", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.Ano.Height = a
    End If
    a = cargarColor("ano", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.Ano.Alignment = a
        End If
    End If
    a = cargarPos("ano", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.Ano.Font.Size = a
    End If
    a = cargarColor("ano", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.Ano.BackColor = a
        End If
    End If
    a = cargarColor("ano", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.Ano.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub Meses()
    Dim a As Variant
    RptImpresionFacturaVenta.Meses.Left = PasarTamano(cargarPos("meses", "FACTURA", "posx"))
    RptImpresionFacturaVenta.Meses.Top = PasarTamano(cargarPos("meses", "FACTURA", "posy"))
    RptImpresionFacturaVenta.Meses.Visible = cargarPos("meses", "FACTURA", "visible")
    a = PasarTamano(cargarPos("meses", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.Meses.Width = a
    End If
    a = PasarTamano(cargarPos("meses", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.Meses.Height = a
    End If
    a = cargarColor("meses", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.Meses.Alignment = a
        End If
    End If
    a = cargarPos("meses", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.Meses.Font.Size = a
    End If
    a = cargarColor("meses", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.Meses.BackColor = a
        End If
    End If
    a = cargarColor("meses", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.Meses.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub Iva() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.lbliva.Left = PasarTamano(cargarPos("iva", "FACTURA", "posx"))
    RptImpresionFacturaVenta.lbliva.Top = PasarTamano(cargarPos("iva", "FACTURA", "posy"))
    RptImpresionFacturaVenta.lbliva.Visible = cargarPos("iva", "FACTURA", "visible")
    a = PasarTamano(cargarPos("iva", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.lbliva.Width = a
    End If
    a = PasarTamano(cargarPos("iva", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.lbliva.Height = a
    End If
    a = cargarColor("iva", "FACTURA", "alineahorizontal")
    If Not a = "" Then
        If Not IsNull(a) Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.lbliva.Alignment = a
        End If
    End If
    a = cargarPos("iva", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.lbliva.Font.Size = a
    End If
    a = cargarColor("iva", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.lbliva.BackColor = a
        End If
    End If
    a = cargarColor("iva", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.lbliva.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub condicion() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.lblpago.Left = PasarTamano(cargarPos("CondiVenta", "FACTURA", "posx"))
    RptImpresionFacturaVenta.lblpago.Top = PasarTamano(cargarPos("CondiVenta", "FACTURA", "posy"))
    RptImpresionFacturaVenta.lblpago.Visible = cargarPos("CondiVenta", "FACTURA", "visible")
    a = PasarTamano(cargarPos("CondiVenta", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.lblpago.Width = a
    End If
    a = PasarTamano(cargarPos("CondiVenta", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.lblpago.Height = a
    End If
    a = cargarColor("CondiVenta", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.lblpago.Alignment = a
        End If
    End If
    a = cargarPos("CondiVenta", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.lblpago.Font.Size = a
    End If
    a = cargarColor("CondiVenta", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.lblpago.BackColor = a
        End If
    End If
    a = cargarColor("CondiVenta", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.lblpago.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub bruto() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.IngBruto.Left = PasarTamano(cargarPos("IngBruto", "FACTURA", "posx"))
    RptImpresionFacturaVenta.IngBruto.Top = PasarTamano(cargarPos("IngBruto", "FACTURA", "posy"))
    RptImpresionFacturaVenta.IngBruto.Visible = cargarPos("IngBruto", "FACTURA", "visible")
    a = PasarTamano(cargarPos("IngBruto", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.IngBruto.Width = a
    End If
    a = PasarTamano(cargarPos("IngBruto", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.IngBruto.Height = a
    End If
    a = cargarColor("IngBruto", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.IngBruto.Alignment = a
        End If
    End If
    a = cargarPos("IngBruto", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.IngBruto.Font.Size = a
    End If
    a = cargarColor("IngBruto", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.IngBruto.BackColor = a
        End If
    End If
    a = cargarColor("IngBruto", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.IngBruto.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub cliente() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.NroCli.Left = PasarTamano(cargarPos("NroCli", "FACTURA", "posx"))
    RptImpresionFacturaVenta.NroCli.Top = PasarTamano(cargarPos("NroCli", "FACTURA", "posy"))
    RptImpresionFacturaVenta.NroCli.Visible = cargarPos("NroCli", "FACTURA", "visible")
    a = PasarTamano(cargarPos("NroCli", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.NroCli.Width = a
    End If
    a = PasarTamano(cargarPos("NroCli", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.NroCli.Height = a
    End If
    a = cargarColor("NroCli", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.NroCli.Alignment = a
        End If
    End If
    a = cargarPos("NroCli", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.NroCli.Font.Size = a
    End If
    a = cargarColor("NroCli", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.NroCli.BackColor = a
        End If
    End If
    a = cargarColor("NroCli", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.NroCli.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub presupuesto() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.Presupu.Left = PasarTamano(cargarPos("Presupu", "FACTURA", "posx"))
    RptImpresionFacturaVenta.Presupu.Top = PasarTamano(cargarPos("Presupu", "FACTURA", "posy"))
    RptImpresionFacturaVenta.Presupu.Visible = cargarPos("Presupu", "FACTURA", "visible")
    a = PasarTamano(cargarPos("Presupu", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.Presupu.Width = a
    End If
    a = PasarTamano(cargarPos("Presupu", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.Presupu.Height = a
    End If
    a = cargarColor("Presupu", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.Presupu.Alignment = a
        End If
    End If
    a = cargarPos("Presupu", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.Presupu.Font.Size = a
    End If
    a = cargarColor("Presupu", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.Presupu.BackColor = a
        End If
    End If
    a = cargarColor("Presupu", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.Presupu.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub ordenCompra() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.OrdenComp.Left = PasarTamano(cargarPos("OrdenComp", "FACTURA", "posx"))
    RptImpresionFacturaVenta.OrdenComp.Top = PasarTamano(cargarPos("OrdenComp", "FACTURA", "posy"))
    RptImpresionFacturaVenta.OrdenComp.Visible = cargarPos("OrdenComp", "FACTURA", "visible")
    a = PasarTamano(cargarPos("OrdenComp", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.OrdenComp.Width = a
    End If
    a = PasarTamano(cargarPos("OrdenComp", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.OrdenComp.Height = a
    End If
    a = cargarColor("OrdenComp", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.OrdenComp.Alignment = a
        End If
    End If
    a = cargarPos("OrdenComp", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.OrdenComp.Font.Size = a
    End If
    a = cargarColor("OrdenComp", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.OrdenComp.BackColor = a
        End If
    End If
    a = cargarColor("OrdenComp", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.OrdenComp.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub Debe() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.Debe.Left = PasarTamano(cargarPos("Debe", "FACTURA", "posx"))
    RptImpresionFacturaVenta.Debe.Top = PasarTamano(cargarPos("Debe", "FACTURA", "posy"))
    RptImpresionFacturaVenta.Debe.Visible = cargarPos("Debe", "FACTURA", "visible")
    a = PasarTamano(cargarPos("Debe", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.Debe.Width = a
    End If
    a = PasarTamano(cargarPos("Debe", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.Debe.Height = a
    End If
    a = cargarColor("Debe", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.Debe.Alignment = a
        End If
    End If
    a = cargarPos("Debe", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.Debe.Font.Size = a
    End If
    a = cargarColor("Debe", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.Debe.BackColor = a
        End If
    End If
    a = cargarColor("Debe", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.Debe.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub Producto() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.Producto.Left = PasarTamano(cargarPos("producto", "FACTURA", "posx"))
    RptImpresionFacturaVenta.Producto.Top = PasarTamano(cargarPos("producto", "FACTURA", "posy"))
    RptImpresionFacturaVenta.Producto.Visible = cargarPos("producto", "FACTURA", "visible")
    a = PasarTamano(cargarPos("producto", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.Producto.Width = a
    End If
    a = PasarTamano(cargarPos("producto", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.Producto.Height = a
    End If
    a = cargarColor("producto", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.Producto.Alignment = a
        End If
    End If
    a = cargarPos("producto", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.Producto.Font.Size = a
    End If
    a = cargarColor("producto", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.Producto.BackColor = a
        End If
    End If
    a = cargarColor("producto", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.Producto.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub NumeroRemito() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.Remitos.Left = PasarTamano(cargarPos("NroRemito", "FACTURA", "posx"))
    RptImpresionFacturaVenta.Remitos.Top = PasarTamano(cargarPos("NroRemito", "FACTURA", "posy"))
    RptImpresionFacturaVenta.Remitos.Visible = cargarPos("NroRemito", "FACTURA", "visible")
    a = PasarTamano(cargarPos("NroRemito", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.Remitos.Width = a
    End If
    a = PasarTamano(cargarPos("NroRemito", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.Remitos.Height = a
    End If
    a = cargarColor("NroRemito", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.Remitos.Alignment = a
        End If
    End If
    a = cargarPos("NroRemito", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.Remitos.Font.Size = a
    End If
    a = cargarColor("NroRemito", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.Remitos.BackColor = a
        End If
    End If
    a = cargarColor("NroRemito", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.Remitos.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub ResponsableInsc() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.lblRespInsc.Left = PasarTamano(cargarPos("lblRespInsc", "FACTURA", "posx"))
    RptImpresionFacturaVenta.lblRespInsc.Top = PasarTamano(cargarPos("lblRespInsc", "FACTURA", "posy"))
    RptImpresionFacturaVenta.lblRespInsc.Visible = cargarPos("lblRespInsc", "FACTURA", "visible")
End Sub
Public Sub ResponsableNoInsc() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.lblRespNoInsc.Left = PasarTamano(cargarPos("lblRespNoInsc", "FACTURA", "posx"))
    RptImpresionFacturaVenta.lblRespNoInsc.Top = PasarTamano(cargarPos("lblRespNoInsc", "FACTURA", "posy"))
    RptImpresionFacturaVenta.lblRespNoInsc.Visible = cargarPos("lblRespNoInsc", "FACTURA", "visible")
End Sub
Public Sub Localidad() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.lbllocalidad.Left = PasarTamano(cargarPos("Localidad", "FACTURA", "posx"))
    RptImpresionFacturaVenta.lbllocalidad.Top = PasarTamano(cargarPos("Localidad", "FACTURA", "posy"))
    RptImpresionFacturaVenta.lbllocalidad.Visible = cargarPos("Localidad", "FACTURA", "visible")
    a = PasarTamano(cargarPos("Localidad", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.lbllocalidad.Width = a
    End If
    a = PasarTamano(cargarPos("Localidad", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.lbllocalidad.Height = a
    End If
    a = cargarColor("Localidad", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.lbllocalidad.Alignment = a
        End If
    End If
    a = cargarPos("Localidad", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.lbllocalidad.Font.Size = a
    End If
    a = cargarColor("Localidad", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.lbllocalidad.BackColor = a
        End If
    End If
    a = cargarColor("Localidad", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.lbllocalidad.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub

Public Sub Factura() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.lblfactura.Left = PasarTamano(cargarPos("lblfactura", "FACTURA", "posx"))
    RptImpresionFacturaVenta.lblfactura.Top = PasarTamano(cargarPos("lblfactura", "FACTURA", "posy"))
    RptImpresionFacturaVenta.lblfactura.Visible = cargarPos("lblfactura", "FACTURA", "visible")
    a = PasarTamano(cargarPos("lblfactura", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.lblfactura.Width = a
    End If
    a = PasarTamano(cargarPos("lblfactura", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.lblfactura.Height = a
    End If
    a = cargarColor("lblfactura", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.lblfactura.Alignment = a
        End If
    End If
    a = cargarPos("lblfactura", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.lblfactura.Font.Size = a
    End If
    a = cargarColor("lblfactura", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.lblfactura.BackColor = a
        End If
    End If
    a = cargarColor("lblfactura", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.lbllocalidad.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub

Public Sub Provincia() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.Provincia.Left = PasarTamano(cargarPos("Provincia", "FACTURA", "posx"))
    RptImpresionFacturaVenta.Provincia.Top = PasarTamano(cargarPos("Provincia", "FACTURA", "posy"))
    RptImpresionFacturaVenta.Provincia.Visible = cargarPos("Provincia", "FACTURA", "visible")
    a = PasarTamano(cargarPos("Provincia", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.Provincia.Width = a
    End If
    a = PasarTamano(cargarPos("Provincia", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.Provincia.Height = a
    End If
    a = cargarColor("Provincia", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.Provincia.Alignment = a
        End If
    End If
    a = cargarPos("Provincia", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.Provincia.Font.Size = a
    End If
    a = cargarColor("Provincia", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.Provincia.BackColor = a
        End If
    End If
    a = cargarColor("Provincia", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.Provincia.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub Pais() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.Pais.Left = PasarTamano(cargarPos("Pais", "FACTURA", "posx"))
    RptImpresionFacturaVenta.Pais.Top = PasarTamano(cargarPos("Pais", "FACTURA", "posy"))
    RptImpresionFacturaVenta.Pais.Visible = cargarPos("Pais", "FACTURA", "visible")
    a = PasarTamano(cargarPos("Pais", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.Pais.Width = a
    End If
    a = PasarTamano(cargarPos("Pais", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.Pais.Height = a
    End If
    a = cargarColor("Pais", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.Pais.Alignment = a
        End If
    End If
    a = cargarPos("Pais", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.Pais.Font.Size = a
    End If
    a = cargarColor("Pais", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.Pais.BackColor = a
        End If
    End If
    a = cargarColor("Pais", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.Pais.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub Fecha() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.lblfecha.Left = PasarTamano(cargarPos("lblfecha", "FACTURA", "posx"))
    RptImpresionFacturaVenta.lblfecha.Top = PasarTamano(cargarPos("lblfecha", "FACTURA", "posy"))
    RptImpresionFacturaVenta.lblfecha.Visible = cargarPos("lblfecha", "FACTURA", "visible")
    a = PasarTamano(cargarPos("lblfecha", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.lblfecha.Width = a
    End If
    a = PasarTamano(cargarPos("lblfecha", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.lblfecha.Height = a
    End If
    a = cargarColor("lblfecha", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.lblfecha.Alignment = a
        End If
    End If
    a = cargarPos("lblfecha", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.lblfecha.Font.Size = a
    End If
    a = cargarColor("lblfecha", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.lblfecha.BackColor = a
        End If
    End If
    a = cargarColor("lblfecha", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.lblfecha.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub Postal() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.CodPos.Left = PasarTamano(cargarPos("CodPos", "FACTURA", "posx"))
    RptImpresionFacturaVenta.CodPos.Top = PasarTamano(cargarPos("CodPos", "FACTURA", "posy"))
    RptImpresionFacturaVenta.CodPos.Visible = cargarPos("CodPos", "FACTURA", "visible")
    a = PasarTamano(cargarPos("CodPos", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.CodPos.Width = a
    End If
    a = PasarTamano(cargarPos("CodPos", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.CodPos.Height = a
    End If
    a = cargarColor("CodPos", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.CodPos.Alignment = a
        End If
    End If
    a = cargarPos("CodPos", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.CodPos.Font.Size = a
    End If
    a = cargarColor("CodPos", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.CodPos.BackColor = a
        End If
    End If
    a = cargarColor("CodPos", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.CodPos.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub

'****************** detalle *************************************
Public Sub Cantidad() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.txtcantidad.Left = PasarTamano(cargarPos("lblCantidad", "FACTURA", "posx"))
    RptImpresionFacturaVenta.txtcantidad.Top = PasarTamano(cargarPos("lblCantidad", "FACTURA", "posy"))
    RptImpresionFacturaVenta.txtcantidad.Visible = cargarPos("lblCantidad", "FACTURA", "visible")
    a = PasarTamano(cargarPos("lblCantidad", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.txtcantidad.Width = a
    End If
    a = PasarTamano(cargarPos("lblCantidad", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.txtcantidad.Height = a
    End If
    a = cargarColor("lblCantidad", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.txtcantidad.Alignment = a
        End If
    End If
    a = cargarPos("lblCantidad", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.txtcantidad.Font.Size = a
    End If
    a = cargarColor("lblCantidad", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.txtcantidad.BackColor = a
        End If
    End If
    a = cargarColor("lblCantidad", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.txtcantidad.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub Articulo() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.codigo.Left = PasarTamano(cargarPos("lblArticulo", "FACTURA", "posx"))
    RptImpresionFacturaVenta.codigo.Top = PasarTamano(cargarPos("lblArticulo", "FACTURA", "posy"))
    RptImpresionFacturaVenta.codigo.Visible = cargarPos("lblArticulo", "FACTURA", "visible")
    a = PasarTamano(cargarPos("lblArticulo", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.codigo.Width = a
    End If
    a = PasarTamano(cargarPos("lblArticulo", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.codigo.Height = a
    End If
    a = cargarColor("lblArticulo", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.codigo.Alignment = a
        End If
    End If
    a = cargarPos("lblArticulo", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.codigo.Font.Size = a
    End If
    a = cargarColor("lblArticulo", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.codigo.BackColor = a
        End If
    End If
    a = cargarColor("lblArticulo", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.codigo.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub DESCRIPCION() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.txtDescripcion.Left = PasarTamano(cargarPos("lblDescripcion", "FACTURA", "posx"))
    RptImpresionFacturaVenta.txtDescripcion.Top = PasarTamano(cargarPos("lblDescripcion", "FACTURA", "posy"))
    RptImpresionFacturaVenta.txtDescripcion.Visible = cargarPos("lblDescripcion", "FACTURA", "visible")
    a = PasarTamano(cargarPos("lblDescripcion", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.txtDescripcion.Width = a
    End If
    a = PasarTamano(cargarPos("lblDescripcion", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.txtDescripcion.Height = a
    End If
    a = cargarColor("lblDescripcion", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.txtDescripcion.Alignment = a
        End If
    End If
    a = cargarPos("lblDescripcion", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.txtDescripcion.Font.Size = a
    End If
    a = cargarColor("lblDescripcion", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.txtDescripcion.BackColor = a
        End If
    End If
    a = cargarColor("lblDescripcion", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.txtDescripcion.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub Unitario() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.txtprecio.Left = PasarTamano(cargarPos("lblPrecUnitario", "FACTURA", "posx"))
    RptImpresionFacturaVenta.txtprecio.Top = PasarTamano(cargarPos("lblPrecUnitario", "FACTURA", "posy"))
    RptImpresionFacturaVenta.txtprecio.Visible = cargarPos("lblPrecUnitario", "FACTURA", "visible")
    a = PasarTamano(cargarPos("lblPrecUnitario", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.txtprecio.Width = a
    End If
    a = PasarTamano(cargarPos("lblPrecUnitario", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.txtprecio.Height = a
    End If
    a = cargarColor("lblPrecUnitario", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.txtprecio.Alignment = a
        End If
    End If
    a = cargarPos("lblPrecUnitario", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.txtprecio.Font.Size = a
    End If
    a = cargarColor("lblPrecUnitario", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.txtprecio.BackColor = a
        End If
    End If
    a = cargarColor("lblPrecUnitario", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.txtprecio.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub PrecioTotal() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.txttotal.Left = PasarTamano(cargarPos("lblPrecTotal", "FACTURA", "posx"))
    RptImpresionFacturaVenta.txttotal.Top = PasarTamano(cargarPos("lblPrecTotal", "FACTURA", "posy"))
    RptImpresionFacturaVenta.txttotal.Visible = cargarPos("lblPrecTotal", "FACTURA", "visible")
    a = PasarTamano(cargarPos("lblPrecTotal", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.txttotal.Width = a
    End If
    a = PasarTamano(cargarPos("lblPrecTotal", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.txttotal.Height = a
    End If
    a = cargarColor("lblPrecTotal", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.txttotal.Alignment = a
        End If
    End If
    a = cargarPos("lblPrecTotal", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.txttotal.Font.Size = a
    End If
    a = cargarColor("lblPrecTotal", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.txttotal.BackColor = a
        End If
    End If
    a = cargarColor("lblPrecTotal", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.txttotal.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
'******************pie de pagina*********************************

Public Sub subtotal() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.txtsub.Left = PasarTamano(cargarPos("subtotal", "FACTURA", "posx"))
    RptImpresionFacturaVenta.txtsub.Top = PasarTamano(cargarPos("subtotal", "FACTURA", "posy"))
    RptImpresionFacturaVenta.txtsub.Visible = cargarPos("subtotal", "FACTURA", "visible")
    a = PasarTamano(cargarPos("subtotal", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.txtsub.Width = a
    End If
    a = PasarTamano(cargarPos("subtotal", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.txtsub.Height = a
    End If
    a = cargarColor("subtotal", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.txtsub.Alignment = a
        End If
    End If
    a = cargarPos("subtotal", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.txtsub.Font.Size = a
    End If
    a = cargarColor("subtotal", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.txtsub.BackColor = a
        End If
    End If
    a = cargarColor("subtotal", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.txtsub.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub

Public Sub Impuesto() ' el tamaño de letra toma un maximo de 16

    Dim a As Variant
    RptImpresionFacturaVenta.Impuesto.Left = PasarTamano(cargarPos("impuesto", "FACTURA", "posx"))
    RptImpresionFacturaVenta.Impuesto.Top = PasarTamano(cargarPos("impuesto", "FACTURA", "posy"))
    RptImpresionFacturaVenta.Impuesto.Visible = cargarPos("impuesto", "FACTURA", "visible")
    a = PasarTamano(cargarPos("impuesto", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.Impuesto.Width = a
    End If
    a = PasarTamano(cargarPos("impuesto", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.Impuesto.Height = a
    End If
    a = cargarColor("impuesto", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.Impuesto.Alignment = a
        End If
    End If
    a = cargarPos("impuesto", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.Impuesto.Font.Size = a
    End If
    a = cargarColor("impuesto", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.Impuesto.BackColor = a
        End If
    End If
    a = cargarColor("impuesto", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.Impuesto.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub

Public Sub Subtotal2() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.Subtotal2.Left = PasarTamano(cargarPos("subtotal2", "FACTURA", "posx"))
    RptImpresionFacturaVenta.Subtotal2.Top = PasarTamano(cargarPos("subtotal2", "FACTURA", "posy"))
    RptImpresionFacturaVenta.Subtotal2.Visible = cargarPos("subtotal2", "FACTURA", "visible")
    a = PasarTamano(cargarPos("subtotal2", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.Subtotal2.Width = a
    End If
    a = PasarTamano(cargarPos("subtotal2", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.Subtotal2.Height = a
    End If
    a = cargarColor("subtotal2", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.Subtotal2.Alignment = a
        End If
    End If
    a = cargarPos("subtotal2", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.Subtotal2.Font.Size = a
    End If
    a = cargarColor("subtotal2", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.Subtotal2.BackColor = a
        End If
    End If
    a = cargarColor("subtotal2", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.Subtotal2.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub ivainscripto() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.IvaInsc.Left = PasarTamano(cargarPos("IvaInsc", "FACTURA", "posx"))
    RptImpresionFacturaVenta.IvaInsc.Top = PasarTamano(cargarPos("IvaInsc", "FACTURA", "posy"))
    RptImpresionFacturaVenta.IvaInsc.Visible = cargarPos("IvaInsc", "FACTURA", "visible")
    a = PasarTamano(cargarPos("IvaInsc", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.IvaInsc.Width = a
    End If
    a = PasarTamano(cargarPos("IvaInsc", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.IvaInsc.Height = a
    End If
    a = cargarColor("IvaInsc", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.IvaInsc.Alignment = a
        End If
    End If
    a = cargarPos("IvaInsc", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.IvaInsc.Font.Size = a
    End If
    a = cargarColor("IvaInsc", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.IvaInsc.BackColor = a
        End If
    End If
    a = cargarColor("IvaInsc", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.IvaInsc.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub

Public Sub ivaNoinscripto() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.IvaNoInsc.Left = PasarTamano(cargarPos("IvaNoInsc", "FACTURA", "posx"))
    RptImpresionFacturaVenta.IvaNoInsc.Top = PasarTamano(cargarPos("IvaNoInsc", "FACTURA", "posy"))
    RptImpresionFacturaVenta.IvaNoInsc.Visible = cargarPos("IvaNoInsc", "FACTURA", "visible")
    a = PasarTamano(cargarPos("IvaNoInsc", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.IvaNoInsc.Width = a
    End If
    a = PasarTamano(cargarPos("IvaNoInsc", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.IvaNoInsc.Height = a
    End If
    a = cargarColor("IvaNoInsc", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.IvaNoInsc.Alignment = a
        End If
    End If
    a = cargarPos("IvaNoInsc", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.IvaNoInsc.Font.Size = a
    End If
    a = cargarColor("IvaNoInsc", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.IvaNoInsc.BackColor = a
        End If
    End If
    a = cargarColor("IvaNoInsc", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.IvaNoInsc.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub Total() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.txttotalfinal.Left = PasarTamano(cargarPos("total", "FACTURA", "posx"))
    RptImpresionFacturaVenta.txttotalfinal.Top = PasarTamano(cargarPos("total", "FACTURA", "posy"))
    RptImpresionFacturaVenta.txttotalfinal.Visible = cargarPos("total", "FACTURA", "visible")
    a = PasarTamano(cargarPos("total", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.txttotalfinal.Width = a
    End If
    a = PasarTamano(cargarPos("total", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.txttotalfinal.Height = a
    End If
    a = cargarColor("total", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.txttotalfinal.Alignment = a
        End If
    End If
    a = cargarPos("total", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.txttotalfinal.Font.Size = a
    End If
    a = cargarColor("total", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.txttotalfinal.BackColor = a
        End If
    End If
    a = cargarColor("total", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.txttotalfinal.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub Descuento() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.txtDcto.Left = PasarTamano(cargarPos("txtDcto", "FACTURA", "posx"))
    RptImpresionFacturaVenta.txtDcto.Top = PasarTamano(cargarPos("txtDcto", "FACTURA", "posy"))
    RptImpresionFacturaVenta.txtDcto.Visible = cargarPos("txtDcto", "FACTURA", "visible")
    a = PasarTamano(cargarPos("txtDcto", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.txtDcto.Width = a
    End If
    a = PasarTamano(cargarPos("txtDcto", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.txtDcto.Height = a
    End If
    a = cargarColor("txtDcto", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.txtDcto.Alignment = a
        End If
    End If
    a = cargarPos("txtDcto", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.txtDcto.Font.Size = a
    End If
    a = cargarColor("txtDcto", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.txtDcto.BackColor = a
        End If
    End If
    a = cargarColor("txtDcto", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.txtDcto.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub DescuentoP() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.txtDctoP.Left = PasarTamano(cargarPos("txtDctop", "FACTURA", "posx"))
    RptImpresionFacturaVenta.txtDctoP.Top = PasarTamano(cargarPos("txtDctop", "FACTURA", "posy"))
    RptImpresionFacturaVenta.txtDctoP.Visible = cargarPos("txtDctop", "FACTURA", "visible")
    a = PasarTamano(cargarPos("txtDctop", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.txtDctoP.Width = a
    End If
    a = PasarTamano(cargarPos("txtDctop", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.txtDctoP.Height = a
    End If
    a = cargarColor("txtDctop", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.txtDctoP.Alignment = a
        End If
    End If
    a = cargarPos("txtDctop", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.txtDctoP.Font.Size = a
    End If
    a = cargarColor("txtDctop", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.txtDctoP.BackColor = a
        End If
    End If
    a = cargarColor("txtDctop", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.txtDctoP.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub IvaIn() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.txtivains.Left = PasarTamano(cargarPos("txtivains", "FACTURA", "posx"))
    RptImpresionFacturaVenta.txtivains.Top = PasarTamano(cargarPos("txtivains", "FACTURA", "posy"))
    RptImpresionFacturaVenta.txtivains.Visible = cargarPos("txtivains", "FACTURA", "visible")
    a = PasarTamano(cargarPos("txtivains", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.txtivains.Width = a
    End If
    a = PasarTamano(cargarPos("txtivains", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.txtivains.Height = a
    End If
    a = cargarColor("txtivains", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.txtivains.Alignment = a
        End If
    End If
    a = cargarPos("txtivains", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.txtivains.Font.Size = a
    End If
    a = cargarColor("txtivains", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.txtivains.BackColor = a
        End If
    End If
    a = cargarColor("txtivains", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.txtivains.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub IvaInP() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.txtIvaP.Left = PasarTamano(cargarPos("txtIvaP", "FACTURA", "posx"))
    RptImpresionFacturaVenta.txtIvaP.Top = PasarTamano(cargarPos("txtIvaP", "FACTURA", "posy"))
    RptImpresionFacturaVenta.txtIvaP.Visible = cargarPos("txtIvaP", "FACTURA", "visible")
    a = PasarTamano(cargarPos("txtIvaP", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.txtIvaP.Width = a
    End If
    a = PasarTamano(cargarPos("txtIvaP", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.txtIvaP.Height = a
    End If
    a = cargarColor("txtIvaP", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.txtIvaP.Alignment = a
        End If
    End If
    a = cargarPos("txtIvaP", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.txtIvaP.Font.Size = a
    End If
    a = cargarColor("txtIvaP", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.txtIvaP.BackColor = a
        End If
    End If
    a = cargarColor("txtIvaP", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.txtIvaP.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub IIBB() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.txtIIBB.Left = PasarTamano(cargarPos("txtIIBB", "FACTURA", "posx"))
    RptImpresionFacturaVenta.txtIIBB.Top = PasarTamano(cargarPos("txtIIBB", "FACTURA", "posy"))
    RptImpresionFacturaVenta.txtIIBB.Visible = cargarPos("txtIIBB", "FACTURA", "visible")
    a = PasarTamano(cargarPos("txtIIBB", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.txtIIBB.Width = a
    End If
    a = PasarTamano(cargarPos("txtIIBB", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.txtIIBB.Height = a
    End If
    a = cargarColor("txtIIBB", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.txtIIBB.Alignment = a
        End If
    End If
    a = cargarPos("txtIIBB", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.txtIIBB.Font.Size = a
    End If
    a = cargarColor("txtIIBB", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.txtIIBB.BackColor = a
        End If
    End If
    a = cargarColor("txtIIBB", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.txtIIBB.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub
Public Sub IIBBP() ' el tamaño de letra toma un maximo de 16
    Dim a As Variant
    RptImpresionFacturaVenta.txtIibbP.Left = PasarTamano(cargarPos("txtIibbP", "FACTURA", "posx"))
    RptImpresionFacturaVenta.txtIibbP.Top = PasarTamano(cargarPos("txtIibbP", "FACTURA", "posy"))
    RptImpresionFacturaVenta.txtIibbP.Visible = cargarPos("txtIibbP", "FACTURA", "visible")
    a = PasarTamano(cargarPos("txtIibbP", "FACTURA", "largo"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.txtIibbP.Width = a
    End If
    a = PasarTamano(cargarPos("txtIibbP", "FACTURA", "ancho"))
    If Not a = 0 Then
        RptImpresionFacturaVenta.txtIibbP.Height = a
    End If
    a = cargarColor("txtIibbP", "FACTURA", "alineahorizontal")
    If Not IsNull(a) Then
        If Not a = "" Then
            If a = "Derecha" Then
                a = ddTXRight
            ElseIf a = "Izquierda" Then
                a = ddTXLeft
            ElseIf a = "Centro" Then
                a = ddTXCenter
            End If
            RptImpresionFacturaVenta.txtIibbP.Alignment = a
        End If
    End If
    a = cargarPos("txtIibbP", "FACTURA", "sizeletra")
    If Not IsNull(a) Then
        If Not a = 0 Then RptImpresionFacturaVenta.txtIibbP.Font.Size = a
    End If
    a = cargarColor("txtIibbP", "FACTURA", "backcolor")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then RptImpresionFacturaVenta.txtIibbP.BackColor = a
        End If
    End If
    a = cargarColor("txtIibbP", "FACTURA", "backstyle")
    If Not IsNull(a) Then
        If Not a = "" Then
            If Not a = 0 Then
                RptImpresionFacturaVenta.txtIibbP.BackStyle = PasoColor(a)
            End If
        End If
    End If
End Sub

'***********************************************************

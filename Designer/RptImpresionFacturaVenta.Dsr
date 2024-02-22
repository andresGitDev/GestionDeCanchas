VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptImpresionFacturaVenta 
   Caption         =   "ActiveReport1"
   ClientHeight    =   12940
   ClientLeft      =   200
   ClientTop       =   560
   ClientWidth     =   11730
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   20690
   _ExtentY        =   22825
   SectionData     =   "RptImpresionFacturaVenta.dsx":0000
End
Attribute VB_Name = "RptImpresionFacturaVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_BeforePrint()
    If Trim(codigo.Text) = "0" Or IsNumeric(codigo.Text) Then
        UMedida.Text = ""
        txtprecio = ""
        txttotal = Format$(Me.txttotal, "Standard")
        txtCantidad = ""
        If s2n(CDbl(txttotal)) = 0 Then
            txttotal = ""
        End If
    Else
        Me.txttotal.Text = Format$(Me.txttotal, "Standard")
        txtimporte = Format$(txtimporte, "standard")
    End If
End Sub

Private Sub Encabezado_BeforePrint()
    If Me.Provincia <> "" Then
        Me.Pais.caption = "Argentina"
    Else
        Me.Pais.caption = ""
    End If
End Sub

Private Sub pagefooter_beforeprint()
    txtsub = Format$(txtsub, "standard")
    Subtotal2 = Format$(txtsub, "standard")
    txtDcto = Format$(txtDcto, "standard")
    txtivains = Format$(txtivains, "standard")
    txtIIBB = Format$(txtIIBB, "standard")
End Sub


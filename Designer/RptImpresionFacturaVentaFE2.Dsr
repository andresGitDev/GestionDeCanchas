VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptImpresionFacturaVentaFE2 
   Caption         =   "Factura electronica"
   ClientHeight    =   10950
   ClientLeft      =   170
   ClientTop       =   560
   ClientWidth     =   15240
   WindowState     =   2  'Maximized
   _ExtentX        =   26882
   _ExtentY        =   19315
   SectionData     =   "RptImpresionFacturaVentaFE2.dsx":0000
End
Attribute VB_Name = "RptImpresionFacturaVentaFE2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public codigoFactura As Long
Public Transporte As Double

Private Sub ActiveReport_PageEnd()
'    Dim pageObject As Canvas
'    Set pageObject = PageEventHandler.Canvas
'    pageObject.ForeColor = vbBlack
'    pageObject.BackStyle = ddBKTransparent
'    pageObject.PenStyle = 1
'    pageObject.PenWidth = 1
'    pageObject.DrawRect 0, 0, pageObject.Width, pageObject.Height
End Sub

Private Sub PageHeader_Format()
    Dim lW As Long, lH As Long
    With Canvas
        .Font.Size = 48
        .MeasureText "Confidential", lW, lH
        .ForeColor = &HE0E0E0
        .DrawText "Confidential", _
        (Me.PrintWidth - lW) / 2, _
        (Me.Printer.PaperHeight - lH) / 2, _
        5200, 2400
    End With
End Sub

Private Sub ActiveReport_PageStart()
'PageHeader_Format
'    Dim rpt As New ActiveReport
'    Set PageEventHandler = rpt
'    'rpt.Show
'    PrintTiled rpt, 3, 3
'    rpt.PrintReport False
End Sub

Sub PrintTiled(rpt As ActiveReport, rowCount As Integer, colCount As Integer)
'Dim row, col
'Dim pagex As Long, pagey As Long, pagew As Long, pageh As Long
'Dim margx As Long, margy As Long
'Dim pageNum As Long
'rpt.Printer.StartJob "One Page"
'rpt.Printer.StartPage
'pagew = (rpt.Printer.PaperWidth * RptImpresionFacturaVentaFE2.PageSettings.LeftMargin * RptImpresionFacturaVentaFE2.PageSettings.RightMargin) / colCount
'pageh = (rpt.Printer.PaperHeight * RptImpresionFacturaVentaFE2.PageSettings.TopMargin * RptImpresionFacturaVentaFE2.PageSettings.BottomMargin) / rowCount
'    For row = 0 To rowCount - 1
'    pagey = (pageh + SPACE_BETWEEN_PAGES) * row
'        For col = 0 To colCount - 1
'        pagex = (pagew + SPACE_BETWEEN_PAGES) * col
'        pageNum = row * colCount + col
'            If (pageNum < rpt.pages.Count) Then
'                rpt.Printer.PrintPage rpt.pages(pageNum), pagex, pagey, pagew, pageh
'            End If
'        Next
'    Next
'rpt.Printer.endpage
'rpt.Printer.EndJob
End Sub

Private Sub Detail_BeforePrint()
Me.txtSubotal.Text = s2n(Me.txtSubotal, 2, True)
Transporte = s2n(Transporte + s2n(Me.txtSubotal.Text), 2, True)
If Me.txtCantidad = 0 Then
    Me.txtCantidad = ""
    Me.txtUMedida = ""
    Me.txtPrecioU = ""
    Me.txtSubotal = ""
    Me.txtIva = ""
End If
If Me.txtSubotal = "0,00" Then
    Me.txtIva = s2n(0, 2, True)
End If
End Sub

Private Sub GroupHeader2_AfterPrint()
    If Transporte > 0 Then
        Me.Label36.Visible = True
        Me.Label35.Visible = True
        Me.Label36.caption = s2n(Transporte, 2, True)
        'Transporte = Me.Field3
    Else
        'Me.Field4.Visible = False
        Me.Label36.Visible = False
        Me.Label35.Visible = False
    End If
End Sub

Private Sub GroupHeader2_BeforePrint()
    If Transporte > 0 Then
        Me.Label36.Visible = True
        Me.Label35.Visible = True
        Me.Label36.caption = s2n(Transporte, 2, True)
        'Transporte = Me.Field3
    Else
        'Me.Field4.Visible = False
        Me.Label36.Visible = False
        Me.Label35.Visible = False
    End If
End Sub

Private Sub pagefooter_beforeprint()
    'lo saco por que ya no hacen mas de una factura
    'If CDbl(Me.Field3.Text) <> CDbl(Me.lblSubTotal) Then
    If 0 = 1 Then
        Me.lblDto.caption = 0
        Me.lblIva21.caption = 0
        Me.lblIva10.caption = 0
        Me.lblIIBB.caption = 0
        Me.lblImporteTotal.caption = 0
        Me.lblSubTotal.caption = 0
'    Else
'        If Transporte > 0 Then
'            Me.Label36.Visible = True
'            Me.Label36.caption = Transporte
'        Else
'            Me.Label36.Visible = False
'        End If
    End If

    
End Sub

Private Sub PageFooter_Format()
    Dim codigoQRfactura As New CodigoQR
    picCodigoQR.Picture = codigoQRfactura.CodigoQRRequest(codigoFactura, 120, 120)
End Sub

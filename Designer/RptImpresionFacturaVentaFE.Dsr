VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptImpresionFacturaVentaFE 
   Caption         =   "Factura electronica"
   ClientHeight    =   14790
   ClientLeft      =   170
   ClientTop       =   560
   ClientWidth     =   18960
   WindowState     =   2  'Maximized
   _ExtentX        =   33443
   _ExtentY        =   26088
   SectionData     =   "RptImpresionFacturaVentaFE.dsx":0000
End
Attribute VB_Name = "RptImpresionFacturaVentaFE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
'pagew = (rpt.Printer.PaperWidth * RptImpresionFacturaVentaFE.PageSettings.LeftMargin * RptImpresionFacturaVentaFE.PageSettings.RightMargin) / colCount
'pageh = (rpt.Printer.PaperHeight * RptImpresionFacturaVentaFE.PageSettings.TopMargin * RptImpresionFacturaVentaFE.PageSettings.BottomMargin) / rowCount
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
End Sub

''a quien corresponda
''
''este formulario esta preparado solo para facturas de exportacion osea facturas E
''si en algun momento tonka decide imprimir en una impresora laser u otra que no sea matriz de punto
''todo este formulario hay que corregirlo ya que hay muchos campos que estan ocultos
''
''pd: espero que encuentren este mensaje
''
''raul
''02/10/2008

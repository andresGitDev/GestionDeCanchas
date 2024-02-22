VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptOrdenPago 
   Caption         =   "ORDEN DE PAGO "
   ClientHeight    =   12470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10950
   Icon            =   "RptOrdenPago.dsx":0000
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   19315
   _ExtentY        =   21996
   SectionData     =   "RptOrdenPago.dsx":628A
End
Attribute VB_Name = "RptOrdenPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportStart()
Label31.caption = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
End Sub

Private Sub Detail_BeforePrint()
fieSaldo = Format$(fieSaldo, "standard")
fiePagado = Format$(fiePagado, "standard")
End Sub


Private Sub PageHeader_BeforePrint()
    Image1.Picture = FrmPrincipal.imgLogoSimple
End Sub

Private Sub ReportFooter_BeforePrint()
    Dim empr As String
    Dim Total As Double
    lblefectivo = Format$(lblefectivo, "standard")
    lblcheques = Format$(lblcheques, "standard")
    lbltransf = Format$(lbltransf, "standard")
    LblRetGanancia = Format$(LblRetGanancia, "standard")
    LblretIB = Format$(LblretIB, "standard")
    
    Total = CDbl(lblefectivo) + CDbl(lblcheques) + CDbl(lbltransf) + CDbl(LblRetGanancia) + CDbl(LblretIB)
    lbltotal = Format$(Total, "standard")
    Me.lblpie.caption = lbltotal.caption
    
    empr = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
    s = "RECIBI "
    If empr > "" Then s = s & "DE " & empr
    Me.LblRecibide.caption = s
End Sub


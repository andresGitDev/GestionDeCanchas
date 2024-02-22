VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptOrdenPagoAcuenta 
   Caption         =   "PAGO A CUENTA"
   ClientHeight    =   14790
   ClientLeft      =   230
   ClientTop       =   560
   ClientWidth     =   18960
   Icon            =   "RptOrdenPagoAcuenta.dsx":0000
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   33443
   _ExtentY        =   26088
   SectionData     =   "RptOrdenPagoAcuenta.dsx":628A
End
Attribute VB_Name = "RptOrdenPagoAcuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Detail_BeforePrint()
txtimporte = Format$(txtimporte, "standard")
End Sub



'Private Sub ActiveReport_ReportStart()
'Image1.Picture = FrmPrincipal.imgLogoSimple
'Label2 = VerParametro(BS_DIRECCION_EMPRESA)
'cuit = VerParametro(BS_CUIT_EMPRESA)
'If cuit > "" Then s = s & cuit
'Me.Label3.caption = s
'End Sub

Private Sub pagefooter_beforeprint()
'lblpie.caption = lbltotal.caption

Dim empr As String, s As String
    Me.lbltotal.caption = Format(CDbl(lblefectivo) + CDbl(lblcheques) + CDbl(lbltransf) + CDbl(LblRetGan) + CDbl(LblretIB), "#,##0.00")

    Me.lblpie.caption = lbltotal.caption
    
    empr = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
    s = "RECIBI "
    If empr > "" Then s = s & "DE " & empr
    
    Me.LblRecibide.caption = s
End Sub

Private Sub PageHeader_BeforePrint()
Dim CUIT As String, s As String
Image1.Picture = FrmPrincipal.imgLogoSimple
Label2 = VerParametro(BS_DIRECCION_EMPRESA)
CUIT = VerParametro(BS_CUIT_EMPRESA)
s = "C.U.I.T.: "
If CUIT > "" Then s = s & CUIT
Me.Label3.caption = s
lblefectivo = Format$(lblefectivo, "standard")
lblcheques = Format$(lblcheques, "standard")
lbltransf = Format$(lbltransf, "standard")
LblRetGan = Format$(LblRetGan, "standard")
LblretIB = Format$(LblretIB, "standard")
lbltotal = Format$(lbltotal, "standard")
'Me.lbltotal.caption = s2n(lblefectivo) + s2n(lblcheques) + s2n(lbltransf) + s2n(LblRetGan) + s2n(LblretIB)
End Sub



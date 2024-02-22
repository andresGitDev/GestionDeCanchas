VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptOrdenPagoFacContProv 
   Caption         =   "ORDEN DE PAGO DE FACTURA COMPRA CONTADO"
   ClientHeight    =   11110
   ClientLeft      =   60
   ClientTop       =   350
   ClientWidth     =   15240
   Icon            =   "RptOrdenPagoFacContProv.dsx":0000
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19597
   SectionData     =   "RptOrdenPagoFacContProv.dsx":628A
End
Attribute VB_Name = "RptOrdenPagoFacContProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub pagefooter_beforeprint()
    lblpie.caption = lbltotal.caption
    
    Dim empr As String, s As String
        
    empr = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
    s = "RECIBI "
    If empr > "" Then s = s & "DE " & empr
    
    Me.Label16.caption = s
End Sub
Private Sub PageHeader_BeforePrint()
    Image1.Picture = FrmPrincipal.imgLogoSimple
    
    lblefectivo = Format$(lblefectivo, "standard")
    lblcheques = Format$(lblcheques, "standard")
    lbltransf = Format$(lbltransf, "standard")
    LblRetGanancia = Format$(LblRetGanancia, "standard")
    LblretIB = Format$(LblretIB, "standard")
    lbltotal = Format$(lbltotal, "standard")
'    Me.lbltotal.caption = s2n(lblefectivo) + s2n(lblcheques) + s2n(lbltransf) + s2n(LblRetGanancia) + s2n(LblretIB)
End Sub




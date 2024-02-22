VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptOrdenPagoConstRet_IG 
   Caption         =   "CONSTANCIA DE RETENCION DE IMPUESTO A LA GANANCIA"
   ClientHeight    =   11110
   ClientLeft      =   60
   ClientTop       =   350
   ClientWidth     =   15240
   Icon            =   "RptOrdenPagoConstRet_IG.dsx":0000
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19597
   SectionData     =   "RptOrdenPagoConstRet_IG.dsx":628A
End
Attribute VB_Name = "RptOrdenPagoConstRet_IG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActiveReport_ReportStart()
Label32.caption = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
End Sub

Private Sub Detail_BeforePrint()
fieSaldo = Format$(s2n(fieSaldo), "standard")
End Sub

Private Sub pagefooter_beforeprint()
RG_PagosTotalMes = Format$(s2n(RG_PagosTotalMes), "standard")
retgan = Format$(s2n(retgan), "standard")

Dim empr As String, s As String
        
    empr = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
    s = "RECIBI "
    If empr > "" Then s = s & "DE " & empr
    
    Me.Label14.caption = s
    
End Sub



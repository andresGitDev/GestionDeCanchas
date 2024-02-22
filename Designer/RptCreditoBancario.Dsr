VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptCreditoBancario 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11110
   ClientLeft      =   120
   ClientTop       =   350
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19597
   SectionData     =   "RptCreditoBancario.dsx":0000
End
Attribute VB_Name = "RptCreditoBancario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub ActiveReport_ReportStart()
LblEmpresa.caption = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
End Sub

Private Sub Detail_BeforePrint()
Field6 = Format$(Field6, "standard")
End Sub

Private Sub PageHeader_BeforePrint()
Image1.Picture = FrmPrincipal.imgLogoSimple
End Sub

Private Sub PageHeader_Format()
 Me.LblEmpresa.caption = gEMPR_NombreEmpresa
 Me.txtMovBanc = Format(txtMovBanc, "00000000")
 Me.txtimporte = enletras(txtimporte)
End Sub


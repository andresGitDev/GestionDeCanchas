VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptReciboConImputacion 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11110
   ClientLeft      =   90
   ClientTop       =   350
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19597
   SectionData     =   "RptReciboConImputacion.dsx":0000
End
Attribute VB_Name = "RptReciboConImputacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Detail_BeforePrint()
fieSaldo = Format$(fieSaldo, "standard")
End Sub


Private Sub GroupHeader1_BeforePrint()
lblefectivo = Format$(lblefectivo, "standard")
lblcheques = Format$(lblcheques, "standard")
lbltransf = Format$(lbltransf, "standard")
lblRetenciones = Format$(lblRetenciones, "standard")
lbltotal = Format$(lbltotal, "standard")
End Sub

Private Sub PageHeader_BeforePrint()
Image1.Picture = FrmPrincipal.imgLogoSimple
lblefectivo.Font.Name = "Courier"
lblefectivo.Font.Size = 10
lblefectivo.Font.Bold = True
End Sub


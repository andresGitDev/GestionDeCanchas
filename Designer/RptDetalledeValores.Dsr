VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptDetalledeValores 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11010
   ClientLeft      =   170
   ClientTop       =   560
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "RptDetalledeValores.dsx":0000
End
Attribute VB_Name = "RptDetalledeValores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportStart()
    Label12.caption = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
End Sub

Private Sub Detail_BeforePrint()
txtimporte = Format$(txtimporte, "standard")
End Sub


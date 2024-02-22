VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptDetalledeRetenciones 
   Caption         =   "ActiveReport1"
   ClientHeight    =   6670
   ClientLeft      =   60
   ClientTop       =   350
   ClientWidth     =   10100
   Icon            =   "RptDetalledeRetenciones.dsx":0000
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   17815
   _ExtentY        =   11765
   SectionData     =   "RptDetalledeRetenciones.dsx":628A
End
Attribute VB_Name = "RptDetalledeRetenciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActiveReport_ReportStart()
    Label12.caption = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
End Sub

Private Sub Detail_BeforePrint()
Importe = Format$(Importe, "standard")
End Sub

VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptPedidoCliente 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   350
   ClientWidth     =   15240
   StartUpPosition =   1  'CenterOwner
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "RptPedidoCliente.dsx":0000
End
Attribute VB_Name = "RptPedidoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActiveReport_ReportStart()
LblEmpresa.caption = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
End Sub

Private Sub Detail_BeforePrint()
Me.Importe = Round(Me.Importe, 4)
End Sub

Private Sub pagefooter_beforeprint()
    LblEmpresa.caption = VerParametro(BS_NOMBRE_EMPRESA_CORTO) 'esto aca no hace nada
    Me.Total = Format$(Me.Total, "standard")
End Sub


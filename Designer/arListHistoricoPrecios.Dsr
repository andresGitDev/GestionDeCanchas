VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arListHistoricoPrecios 
   Caption         =   "ActiveReport1"
   ClientHeight    =   14790
   ClientLeft      =   230
   ClientTop       =   560
   ClientWidth     =   18960
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   33443
   _ExtentY        =   26088
   SectionData     =   "arListHistoricoPrecios.dsx":0000
End
Attribute VB_Name = "arListHistoricoPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Detail_BeforePrint()
fiePU.Text = Format$(fiePU, "standard")
End Sub


Private Sub PageHeader_BeforePrint()
    lblNombreEmpresa.caption = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
End Sub


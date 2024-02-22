VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptDiferenciaStock 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11110
   ClientLeft      =   60
   ClientTop       =   350
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19597
   SectionData     =   "RptDiferenciaStock.dsx":0000
End
Attribute VB_Name = "RptDiferenciaStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Detail_BeforePrint()
    Serie.Text = sSinNull(obtenerDeSQL("select serie from series where nrocomprobante=" & Field6 & " and producto='" & Trim(Field3) & "'"))
End Sub

Private Sub PageHeader_BeforePrint()
Image1.Picture = FrmPrincipal.imgLogoSimple

End Sub


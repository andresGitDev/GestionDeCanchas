VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptMovimientoStock2 
   Caption         =   "Movimiento de Stock"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   350
   ClientWidth     =   14090
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   24853
   _ExtentY        =   14764
   SectionData     =   "rptMovimientoStock2.dsx":0000
End
Attribute VB_Name = "rptMovimientoStock2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Detail_BeforePrint()
    If Len(Trim(fieTipoComprobante)) <> 3 Then
        fieTipoComprobante.Font.Bold = True
    Else
        fieTipoComprobante.Font.Bold = False
    End If
End Sub

Private Sub PageHeader_BeforePrint()
    Me.LblEmpresa.caption = gEMPR_NombreEmpresa
End Sub


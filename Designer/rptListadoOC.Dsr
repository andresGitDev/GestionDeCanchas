VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptListadoOC 
   Caption         =   "Listado de Ordenes de Compra"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   350
   ClientWidth     =   14340
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   25294
   _ExtentY        =   14764
   SectionData     =   "rptListadoOC.dsx":0000
End
Attribute VB_Name = "rptListadoOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Detail_BeforePrint()
fieImporte = Format$(s2n(fieImporte), "standard")
End Sub


Private Sub PageHeader_BeforePrint()
Image1.Picture = FrmPrincipal.imgLogoSimple
End Sub


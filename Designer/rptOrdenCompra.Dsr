VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptOrdenCompra 
   Caption         =   "Orden de Compra"
   ClientHeight    =   8410
   ClientLeft      =   170
   ClientTop       =   560
   ClientWidth     =   11730
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   20690
   _ExtentY        =   14834
   SectionData     =   "rptOrdenCompra.dsx":0000
End
Attribute VB_Name = "rptOrdenCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Detail_BeforePrint()
'fiePU = Format$(s2n(fiePU), "standard")
'fiePT = Format$(s2n(fiePT), "standard")
End Sub

Private Sub pagefooter_beforeprint()
fieTotal = Format$(s2n(fieTotal), "standard")
End Sub



Private Sub PageHeader_Format()
    lblTitulo.caption = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
End Sub

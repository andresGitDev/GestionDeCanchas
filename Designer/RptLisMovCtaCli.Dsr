VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptLisMovCtaCli 
   Caption         =   "Listado de Cta Cte Cliente"
   ClientHeight    =   11010
   ClientLeft      =   170
   ClientTop       =   560
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "RptLisMovCtaCli.dsx":0000
End
Attribute VB_Name = "RptLisMovCtaCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_BeforePrint()
TxtDebe = Format$(s2n(TxtDebe), "standard")
txthaber = Format$(s2n(txthaber), "standard")
txtsaldo = Format$(s2n(txtsaldo), "standard")
End Sub


Private Sub GroupFooter1_BeforePrint()
txttotdebe = Format$(s2n(txttotdebe), "standard")
txttothaber = Format$(s2n(txttothaber), "standard")
End Sub


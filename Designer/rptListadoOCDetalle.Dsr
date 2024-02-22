VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptListadoOCDetalle 
   Caption         =   "Listado de Ordenes de Compra Detallado"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   350
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "rptListadoOCDetalle.dsx":0000
End
Attribute VB_Name = "rptListadoOCDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ImporteGral As Double
Dim Importe As Double

Private Sub Detail_BeforePrint()

Importe = Importe + (Me.fieSaldo * Me.fieCosto)

Me.fieSubImporte = (Me.fieSaldo * Me.fieCosto)
fieSubImporte = Format$(s2n(fieSubImporte), "standard")
fieSaldo = Format$(s2n(fieSaldo), "standard")
fieCosto = Format$(s2n(fieCosto), "standard")
End Sub

Private Sub GroupFooter1_BeforePrint()
Me.fieImporte = Format$(s2n(Importe), "standard")
ImporteGral = ImporteGral + Importe
Importe = 0
End Sub
Private Sub PageHeader_BeforePrint()
Image1.Picture = FrmPrincipal.imgLogoSimple
End Sub

Private Sub ReportFooter_Format()
txtImporteGral = Format$(s2n(ImporteGral), "standard")
End Sub

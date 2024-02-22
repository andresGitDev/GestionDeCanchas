VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptMovimientoBancario 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11010
   ClientLeft      =   170
   ClientTop       =   560
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "rptMovimientoBancario.dsx":0000
End
Attribute VB_Name = "RptMovimientoBancario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub ActiveReport_ReportStart()
    Me.LblEmpresa.caption = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
    Me.lblFechaVer = Date
End Sub

Private Sub Detail_BeforePrint()
Field3 = Format$(s2n(Field3), "standard")
Field4 = Format$(s2n(Field4), "standard")
End Sub


VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptReciboAcuenta 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11010
   ClientLeft      =   170
   ClientTop       =   560
   ClientWidth     =   15240
   Icon            =   "RptReciboAcuenta.dsx":0000
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "RptReciboAcuenta.dsx":628A
End
Attribute VB_Name = "RptReciboAcuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub PageHeader_BeforePrint()

Dim CUIT As String, s As String
Image1.Picture = FrmPrincipal.imgLogoSimple
Label3 = VerParametro(BS_DIRECCION_EMPRESA)
CUIT = VerParametro(BS_CUIT_EMPRESA)
s = "C.U.I.T.: "
If CUIT > "" Then s = s & CUIT
    
Me.Label4.caption = s
'Me.lbltotal.caption = s2n(lblefectivo) + s2n(lblcheques) + s2n(lbltransf) + s2n(LblRetGan) + s2n(LblretIB)
txtefectivo = Format$(s2n(txtefectivo), "standard")
TxtCheques = Format$(s2n(TxtCheques), "standard")
txtTransferencia = Format$(s2n(txtTransferencia), "standard")
txttotal = Format$(s2n(txttotal), "standard")
End Sub

Private Sub Detail_BeforePrint()
txtimporte = Format$(s2n(txtimporte), "standard")
End Sub




VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptLibracionChequeProv 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11110
   ClientLeft      =   60
   ClientTop       =   350
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19597
   SectionData     =   "RptLibracionCheque.dsx":0000
End
Attribute VB_Name = "RptLibracionChequeProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActiveReport_ReportStart()
LblEmpresa.caption = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
End Sub

Private Sub Detail_AfterPrint()
If Me.ContraQue = "CAJA" Then
 Me.txtcuenta = ""
End If
End Sub

Private Sub Detail_BeforePrint()
Field6 = Format$(s2n(Field6), "standard")
End Sub

Private Sub PageHeader_BeforePrint()
Image1.Picture = FrmPrincipal.imgLogoSimple
End Sub

Private Sub PageHeader_Format()
 Me.LblEmpresa.caption = gEMPR_NombreEmpresa
 Me.TxtChequeInterno = Format(Me.TxtChequeInterno, "00000000")
End Sub


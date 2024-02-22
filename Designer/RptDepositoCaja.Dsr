VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptDepositoCaja 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11110
   ClientLeft      =   90
   ClientTop       =   350
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19597
   SectionData     =   "RptDepositoCaja.dsx":0000
End
Attribute VB_Name = "RptDepositoCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub PageHeader_BeforePrint()
Select Case TxtTipo1
Case "E"
     'txtTipo.Text = "EFECTIVO"
     TxtTipo1.Text = "EFECTIVO"
Case "C"
     'txtTipo.Text = "CHEQUE DE TERCERO"
     TxtTipo1.Text = "CHEQUE DE TERCERO"
Case "D"
     'txtTipo.Text = "CREDITO BANCARIO"
     TxtTipo1.Text = "CREDITO BANCARIO"
Case "G"
     'txtTipo.Text = "GASTO BANCARIO"
     TxtTipo1.Text = "GASTO BANCARIO"
Case "P"
     'txtTipo.Text = "CHEQUE PROPIO"
     TxtTipo1.Text = "CHEQUE PROPIO"
Case "T"
     'txtTipo.Text = "TRANSFERENCIA BANCARIA"
     TxtTipo1.Text = "TRANSFERENCIA BANCARIA"
End Select
Image1.Picture = FrmPrincipal.imgLogoSimple
Field2 = Format$(Field2, "standard")
End Sub
Private Sub PageHeader_Format()
 'Me.LblEmpresa.caption = gEMPR_NombreEmpresa
 Me.Txtmovimiento = Format(Txtmovimiento, "00000000")
End Sub


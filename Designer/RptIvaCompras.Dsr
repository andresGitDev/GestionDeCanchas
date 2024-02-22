VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptIvaCompras 
   Caption         =   "Iva Compras"
   ClientHeight    =   11010
   ClientLeft      =   195
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "RptIvaCompras.dsx":0000
End
Attribute VB_Name = "RptIvaCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportStart()
LblEmpresa.caption = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
End Sub

Private Sub PageHeader_BeforePrint()
Dim rsempresa As New ADODB.Recordset
On Error GoTo FaltaFoto
   rsempresa.Open "select nombrelogo,nombrelogofull from datosempresa where idempresa='" & gEMPR_idEmpresa & "'", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
   'If Not IsNull(rsempresa!nombrelogo) Then
       Me.Image1.Picture = FrmPrincipal.imgLogoSimple 'LoadPicture(App.Path & "\" & Trim(rsempresa!nombrelogo))
   'End If
   rsempresa.Close
   Set rsempresa = Nothing
   Exit Sub
FaltaFoto:
   Me.LblEmpresa.caption = gEMPR_NombreEmpresa
   Resume Next
End Sub

Private Sub ReportFooter_BeforePrint()
Field17 = Format$(Field17, "standard")
Field18 = Format$(Field18, "standard")
Field19 = Format$(Field19, "standard")
Field20 = Format$(Field20, "standard")
Field21 = Format$(Field21, "standard")
Field22 = Format$(Field22, "standard")
Field23 = Format$(Field23, "standard")
Field24 = Format$(s2n(Field24), "standard")
Field25 = Format$(Field25, "standard")
Field26 = Format$(Field26, "standard")
Field27 = Format$(Field27, "standard")
Field28 = Format$(Field28, "standard")
Field30 = Format$(Field30, "standard")
End Sub


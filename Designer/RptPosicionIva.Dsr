VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptPosicionIva 
   Caption         =   "Listado de Posicion de Iva"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "RptPosicionIva.dsx":0000
End
Attribute VB_Name = "RptPosicionIva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Detail_BeforePrint()
Dim rsempresa As New ADODB.Recordset
On Error GoTo FaltaFoto
   rsempresa.Open "select nombrelogo from datosempresa where nombre='" & FrmPrincipal.lblNombreEmpresa.caption & "'", DataEnvironment1.AMR, adOpenStatic, adLockReadOnly
   If Not IsNull(rsempresa!nombrelogo) Then
       Me.Image1.Picture = LoadPicture(App.Path & "\" & rsempresa!nombrelogo)
   End If
   rsempresa.Close
   Set rsempresa = Nothing
   Exit Sub
   
FaltaFoto:
   LblEmpresa.caption = gEMPR_NombreEmpresa
   lblSucursal.caption = gEMPR_Sucursal
   Resume Next

End Sub


VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptTotalesTipoCompras 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11110
   ClientLeft      =   60
   ClientTop       =   350
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19597
   SectionData     =   "rptTotalesTipoCompras.dsx":0000
End
Attribute VB_Name = "rptTotalesTipoCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub PageHeader_BeforePrint()
Dim rsempresa As New ADODB.Recordset
On Error GoTo FaltaFoto
   rsempresa.Open "select nombrelogo from datosempresa where nombre='" & gEMPR_NombreEmpresa & "'", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
   If Not IsNull(rsempresa!nombrelogo) Then
       Me.Image1.Picture = LoadPicture(App.Path & "\" & rsempresa!nombrelogo)
   End If
   rsempresa.Close
   Set rsempresa = Nothing
   Exit Sub
FaltaFoto:
   Me.LblEmpresa.caption = gEMPR_NombreEmpresa
   Resume Next
End Sub

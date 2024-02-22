VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptListCosto 
   Caption         =   "Listado de Centro de Costo"
   ClientHeight    =   12270
   ClientLeft      =   60
   ClientTop       =   350
   ClientWidth     =   11610
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   20479
   _ExtentY        =   21643
   SectionData     =   "rptListCosto.dsx":0000
End
Attribute VB_Name = "rptListCosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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


VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptMovimientoStock 
   Caption         =   "Movimiento de Stock"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   350
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "rptMovimientoStock.dsx":0000
End
Attribute VB_Name = "rptMovimientoStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub PageHeader_BeforePrint()
Dim rsempresa As New ADODB.Recordset
On Error GoTo FaltaFoto
   rsempresa.Open "select nombrelogo,nombrelogofull from datosempresa where idempresa='" & gEMPR_idEmpresa & "'", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
   If Not IsNull(rsempresa!nombrelogo) Then
       Me.Image1.Picture = LoadPicture(App.Path & "\" & Trim(rsempresa!nombrelogo))
   End If
   rsempresa.Close
   Set rsempresa = Nothing
   Exit Sub
FaltaFoto:
   Me.LblEmpresa.caption = gEMPR_NombreEmpresa
   Resume Next
End Sub

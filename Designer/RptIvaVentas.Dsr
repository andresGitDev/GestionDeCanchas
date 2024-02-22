VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptIvaVentas 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8200
   ClientLeft      =   170
   ClientTop       =   560
   ClientWidth     =   14550
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   25665
   _ExtentY        =   14464
   SectionData     =   "RptIvaVentas.dsx":0000
End
Attribute VB_Name = "RptIvaVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cant As Double

Private Sub ActiveReport_ReportStart()
LblEmpresa.caption = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
End Sub

'Private Sub Detail_Format()
'    cant = cant + IIf(Me.Field5 = "neto", 0, Me.Field5)
'End Sub

'Private Sub GroupHeader1_Format()
'    Me.Field18 = cant
'End Sub

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
fieTotalNeto = Format$(fieTotalNeto, "standard")
fieTotalNoGrav = Format$(fieTotalNoGrav, "standard")
fieTotalExento = Format$(fieTotalExento, "standard")
fieTotalIVARNI = Format$(fieTotalIVARNI, "standard")
fieTotalIVARI = Format$(fieTotalIVARI, "standard")
fieTotalIVACF = Format$(fieTotalIVACF, "standard")
fieTotalIVABC = Format$(fieTotalIVABC, "standard")
fieTotalRetIva = Format$(fieTotalRetIva, "standard")
fieTotalIIBB = Format$(fieTotalIIBB, "standard")
fieTotalImporte = Format$(fieTotalImporte, "standard")

End Sub



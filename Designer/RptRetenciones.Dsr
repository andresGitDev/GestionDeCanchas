VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptRetenciones 
   Caption         =   "Listado de Retenciones"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "RptRetenciones.dsx":0000
End
Attribute VB_Name = "RptRetenciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
Dim rs As New ADODB.Recordset
sql = "SELECT cuit FROM clientes WHERE descripcion='" & Fiecliente.Text & "'"
rs.Open sql, DataEnvironment1.AMR, adOpenKeyset, adLockReadOnly
If Not rs.EOF Then
    Me.FieCuit.Text = rs!Cuit
End If
rs.Close
End Sub

Private Sub PageHeader_BeforePrint()
Dim rsempresa As New ADODB.Recordset
On Error GoTo FaltaFoto
   rsempresa.Open "select nombrelogo,nombrelogofull from datosempresa where idempresa='" & gEMPR_idEmpresa & "'", DataEnvironment1.AMR, adOpenStatic, adLockReadOnly
   If Not IsNull(rsempresa!nombrelogo) Then
       Me.Image1.Picture = LoadPicture(App.Path & "\" & Trim(rsempresa!nombrelogo))
   End If
   rsempresa.Close
   Set rsempresa = Nothing
   Exit Sub
FaltaFoto:
   Me.lblempresa.caption = gEMPR_NombreEmpresa
   Me.lblSucursal.caption = gEMPR_Sucursal
   Resume Next
End Sub


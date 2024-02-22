VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptPagoACta 
   Caption         =   "Pagos a Cuentas"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   StartUpPosition =   1  'CenterOwner
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "RptPagoACta.dsx":0000
End
Attribute VB_Name = "RptPagoACta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql As String
Dim SumoImp As Double
Dim rs As New ADODB.Recordset
Dim letra As String
Private Sub Detail_BeforePrint()
   If RptPagoACta.Field2.Text = "E" Then
      letra = RptPagoACta.Field2.Text
      RptPagoACta.Field2.Text = "Efectivo"
      sql = "SELECT responsable FROM cajas WHERE codigo = '" & Me.Caja & "'"
      rs.Open sql, DataEnvironment1.AMR, adOpenDynamic, adLockReadOnly
      Me.DescCh.Text = Me.Caja & " " & rs!responsable
      rs.Close
   End If
   If RptPagoACta.Field2.Text = "T" Then
      letra = RptPagoACta.Field2.Text
      RptPagoACta.Field2.Text = "Transferencia"
   
   End If
   If RptPagoACta.Field2.Text = "C" Then
      letra = RptPagoACta.Field2.Text
      RptPagoACta.Field2.Text = "Ch.Tercero"
   
   End If
   If RptPagoACta.Field2.Text = "P" Then
      letra = RptPagoACta.Field2.Text
      RptPagoACta.Field2.Text = "Ch.Propio"
   
   End If
If Me.codprov.Text <> "" Then
   sql = "SELECT descripcion FROM prov WHERE codigo= '" & Me.codprov & "'"
      rs.Open sql, DataEnvironment1.AMR, adOpenKeyset, adLockReadOnly
   If Not rs.EOF Then
      Me.descripcion.Text = rs!descripcion
   End If
   rs.Close
End If
SumoImp = SumoImp + Me.Importe
BuscoDescCH (letra)
End Sub
Private Sub BuscoDescCH(tipo As String)
Select Case tipo
Case "C" 'cheque de terceros
   sql = "SELECT * FROM cheques where nroint= '" & Me.interno & "'"
   rs.Open sql, DataEnvironment1.AMR, adOpenDynamic, adLockReadOnly
   
   Me.DescCh.Text = Banco(rs!banco_Nro) & " -- Cheque Nº " & rs!Nro
   rs.Close
Case "P" 'cheques propios
   sql = "SELECT * FROM chq_comp where codigo= '" & Me.interno & "'"
   rs.Open sql, DataEnvironment1.AMR, adOpenDynamic, adLockReadOnly
   Me.DescCh.Text = Banco(rs!Banco) & " -- Cheque Nº" & rs!Nro
   rs.Close
Case "T" 'transferencia
   sql = "SELECT * FROM ctasbank where codigo= '" & Me.Cuenta & "'"
   rs.Open sql, DataEnvironment1.AMR, adOpenDynamic, adLockReadOnly
   Me.DescCh.Text = Banco(rs!Banco) & " -- Cuenta : " & rs!numero
   rs.Close
End Select
End Sub

Public Function Banco(Nro As Long) As String
Dim rsbco As New ADODB.Recordset
sql = "SELECT descripcion FROM bancosgrales WHERE codigo='" & Nro & "'"
rsbco.Open sql, DataEnvironment1.AMR, adOpenDynamic, adLockReadOnly
If Not rsbco.EOF Then
   Banco = rsbco!descripcion
End If
rsbco.Close
End Function

Private Sub PageFooter_BeforePrint()
   Me.Total.Text = Format$(SumoImp, "standard")
End Sub
Private Sub PageHeader_BeforePrint()
Dim rsempresa As New ADODB.Recordset
On Error GoTo FaltaFoto
   rsempresa.Open "select nombrelogo from datosempresa where nombre='" & gEMPR_NombreEmpresa & "'", DataEnvironment1.AMR, adOpenStatic, adLockReadOnly
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


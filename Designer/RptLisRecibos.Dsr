VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptLisRecibos 
   Caption         =   "Listado de Recibos Emitidos"
   ClientHeight    =   11010
   ClientLeft      =   170
   ClientTop       =   560
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "RptLisRecibos.dsx":0000
End
Attribute VB_Name = "RptLisRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsempresa As New ADODB.Recordset
Dim rsCheques As New ADODB.Recordset
Dim rsmov As New ADODB.Recordset

Private Sub Detail_Format()
Dim AuxValores As String
sql = "Select Fecha, Importe,tipo,interno From MOVICAJA Where ACTIVO = 1 And TIPODOC = 'RAA'" & _
       " And NRODOC = " & Me.txtNro & "" 'And TIPO = 'E'"
        If Me.txtNro <> "" Then
            rsmov.Open sql, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
            Do While Not rsmov.EOF
                If rsmov!Tipo = "E" Then
                    Efectivo.Text = rsmov!Importe
                Else
                    sql = "SELECT * FROM Cheques WHERE "
                    AuxValores = AuxValores & " " & Format$(rsmov!Importe, "standard")
                End If
                rsmov.MoveNext
            Loop
            Valores.Text = AuxValores
            rsmov.Close
        End If
End Sub

Private Sub PageHeader_BeforePrint()
'************    PROCESO PARA LA CARGA DE LOS LOGOS
   CargaLogo
'*******************************************************************
End Sub

Private Sub CargaLogo()
On Error GoTo FaltaFoto

   rsempresa.Open "select nombrelogo from datosempresa where nombre='" & gEMPR_NombreEmpresa & "'", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
   If Not IsNull(rsempresa!nombrelogo) Then
       'Me.Image1.Picture = LoadPicture(App.Path & "\" & rsempresa!nombrelogo)
   End If
   rsempresa.Close
   Set rsempresa = Nothing
   Exit Sub
FaltaFoto:
   Me.LblEmpresa.caption = gEMPR_NombreEmpresa
   Me.lblSucursal.caption = gEMPR_Sucursal
   Resume Next
   
End Sub


Private Sub Detail_BeforePrint()
txtimporte = Format$(s2n(txtimporte), "standard")
End Sub


VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptPedidoCliente2 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11110
   ClientLeft      =   60
   ClientTop       =   350
   ClientWidth     =   15240
   StartUpPosition =   1  'CenterOwner
   _ExtentX        =   26882
   _ExtentY        =   19597
   SectionData     =   "RptPedidoCliente2.dsx":0000
End
Attribute VB_Name = "RptPedidoCliente2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub ActiveReport_QueryClose(cancel As Integer, CloseMode As Integer)
    If cierra = False Then
        cancel = 1
        Me.Hide
    End If
End Sub

Private Sub Detail_BeforePrint()
    Dim reglon As Long
    Dim altura As Long
    If Trim(producto) = "" Then
        cantidad = ""
        Importe = ""
        Totaliva = ""
    ElseIf Trim(producto) = "SUBTOTAL" Or Trim(producto) = "TOTAL" Then
        Totaliva = obtenerDeSQL("select total from itempedidocliente2 where codigo=" & Field7)
        Importe = ""
    End If
    altura = RichEdit1.Height
    'aca paso entero
    reglon = Fix(Len(Trim(obtenerDeSQL("select dbo.rtf2txt(descripcion) from ItemPedidoCliente2 where codigo=" & Field7))) / 73)
    If (Len(Trim(obtenerDeSQL("select dbo.rtf2txt(descripcion) from ItemPedidoCliente2 where codigo=" & Field7))) / 73) - reglon = 0 Then
        RichEdit1.Height = altura * reglon
    Else
        RichEdit1.Height = altura * (reglon + 1)
    End If
    RichEdit1.Text = obtenerDeSQL("select descripcion from ItemPedidoCliente2 where codigo=" & Field7)
    If obtenerDeSQL("select item from ItemPedidoCliente2 where codigo=" & Field7) > 0 Then
        RichEdit2.Text = obtenerDeSQL("select item from ItemPedidoCliente2 where codigo=" & Field7)
    Else
        If obtenerDeSQL("select descripcion from ItemPedidoCliente2 where codigo=" & Field7 & " and tipoitem='Otro'") <> "" Then
            RichEdit2.Text = obtenerDeSQL("select descripcion from ItemPedidoCliente2 where codigo=" & Field7 & " and tipoitem='Otro'")
            RichEdit1.Text = ""
        Else
            RichEdit2.Text = ""
        End If
    End If
    If cantidad.Text = "0" Then
        cantidad.Text = ""
    End If
    
End Sub


Private Sub PageHeader_BeforePrint()
Dim rsempresa As New ADODB.Recordset
On Error GoTo FaltaFoto

   rsempresa.Open "select nombrelogo from datosempresa where nombre='" & FrmPrincipal.lblNombreEmpresa.caption & "'", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
   If Not IsNull(rsempresa!nombrelogo) Then
       'RptPedidoCliente.Image1.Picture = LoadPicture(App.Path & "\" & rsempresa!nombrelogo)
   End If
   rsempresa.Close
   Set rsempresa = Nothing
   Exit Sub
FaltaFoto:
   LblEmpresa.caption = gEMPR_NombreEmpresa
   Resume Next
End Sub

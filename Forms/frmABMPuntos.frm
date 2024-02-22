VERSION 5.00
Begin VB.Form frmABMPuntos 
   Caption         =   "Altas de puntos de venta"
   ClientHeight    =   3885
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDescripcion 
      Height          =   300
      Left            =   1770
      TabIndex        =   8
      Text            =   " "
      Top             =   1485
      Width           =   5955
   End
   Begin VB.TextBox txtNumero 
      Height          =   315
      Left            =   1770
      MaxLength       =   4
      TabIndex        =   7
      Text            =   "0000"
      Top             =   1050
      Width           =   1230
   End
   Begin VB.ComboBox cboTipoPunto 
      Height          =   315
      Left            =   1770
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   600
      Width           =   1800
   End
   Begin VB.ComboBox cboTipo 
      Height          =   315
      Left            =   1770
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   165
      Width           =   1275
   End
   Begin Gestion.ucBotonera uMenu 
      Height          =   1605
      Left            =   90
      TabIndex        =   0
      Top             =   2175
      Width           =   7650
      _ExtentX        =   13494
      _ExtentY        =   2831
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   ""
   End
   Begin VB.Label Label4 
      Caption         =   "Descripcion"
      Height          =   300
      Left            =   150
      TabIndex        =   9
      Top             =   1500
      Width           =   990
   End
   Begin VB.Label Label3 
      Caption         =   "Numero de Punto"
      Height          =   315
      Left            =   135
      TabIndex        =   6
      Top             =   1110
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo de Punto de venta"
      Height          =   465
      Left            =   120
      TabIndex        =   4
      Top             =   615
      Width           =   1335
   End
   Begin VB.Label lblCODFactura 
      Caption         =   "000"
      Height          =   300
      Left            =   3195
      TabIndex        =   3
      Top             =   195
      Width           =   435
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo comprobante"
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Top             =   210
      Width           =   1500
   End
End
Attribute VB_Name = "frmABMPuntos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private id_en_vista As Long
Private sTipo_cargado As String
Private Sub cboTipo_Click()
Dim ws As New FacturaElectronica
lblCODFactura = ws.CodigoComprobante(cboTipo.ListIndex)
ArmoDescripcion
End Sub

Private Sub cboTipoPunto_Click()
ArmoDescripcion
End Sub

Private Function InicioTiposPuntos()
Dim i As wsSiglasComprobante
Dim wd As New FacturaElectronica
cboTipo.clear

For i = 0 To 24
    cboTipo.AddItem wd.LetraComprobante(i)
Next

End Function

Private Sub Form_Load()
InicioTiposPuntos

cboTipoPunto.AddItem "Pre-Impresa"
cboTipoPunto.AddItem "On-Line"
cboTipoPunto.AddItem "Web-Service"
cboTipoPunto.AddItem "Fiscal"

uMenu.init True, True, False, False, True, "select * from documentoscae", DataEnvironment1.Sistema
Limpio
End Sub

Private Function ArmoDescripcion()
Dim sDescripcion As String
Dim sLetra1 As String, sLetra2 As String, sLetra3 As String
If Trim(cboTipo.Text) = "" Then Exit Function
sLetra1 = CORTO(cboTipo.Text, 0, 2)
sLetra2 = CORTO(cboTipo.Text, 1, 1)
sLetra3 = CORTO(cboTipo.Text, 2, 0)
Select Case sLetra1
    Case "F": sDescripcion = "Factura "
    Case "N": sDescripcion = "Nota de "
    Case "C": sDescripcion = "Nota de "
    Case "D": sDescripcion = "Nota de "
End Select

Select Case sLetra2
    Case "A": sDescripcion = sDescripcion & "tipo "
    Case "D": sDescripcion = sDescripcion & "debito "
    Case "C": sDescripcion = sDescripcion & "credito "
    Case "E":
        If sLetra1 = "F" Then sDescripcion = sDescripcion & "de credito "
        If sLetra1 = "C" Then sDescripcion = sDescripcion & "credito electronica "
        If sLetra1 = "D" Then sDescripcion = sDescripcion & "debito electronica "
End Select

sDescripcion = sDescripcion & sLetra3 & " "

Select Case cboTipoPunto.ListIndex
    Case 0: sDescripcion = sDescripcion & "- PreImpresa "
    Case 1: sDescripcion = sDescripcion & "- OnLine "
    Case 2: sDescripcion = sDescripcion & "- WebService "
End Select

sDescripcion = sDescripcion & " - Punto " & txtNumero

txtDescripcion = sDescripcion

End Function

Private Function Limpio()
cboTipo.ListIndex = 0
cboTipoPunto.ListIndex = 2
txtDescripcion = ""
id_en_vista = 0
sTipo_cargado = ""
End Function

Private Sub txtNumero_Change()
ArmoDescripcion
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)

KeyAscii = SoloNum(KeyAscii, False, False)
ArmoDescripcion
End Sub

Private Sub uMenu_Aceptar()
Dim sInsertUpdate As String
Dim sTipoPunto As String, nPermitoCae As Integer, sPunto As String
Dim nExiste


Select Case cboTipoPunto.ListIndex
    Case 0: sTipoPunto = "PI"
    Case 1: sTipoPunto = "OL"
    Case 2: sTipoPunto = "WS"
    Case 3: sTipoPunto = "FI"
End Select

If UCase(cboTipo.Text) = "FEA" Or UCase(cboTipo.Text) = "FEB" Or UCase(cboTipo.Text) = "FEC" Then
    If sTipoPunto = "PI" Or sTipoPunto = "FI" Then
        MsgBox "Las facturas de credito solo pueden ser Web-Service u On-Line"
        Exit Sub
    End If
End If

If UCase(cboTipo.Text) = "CEA" Or UCase(cboTipo.Text) = "CEB" Or UCase(cboTipo.Text) = "CEC" Then
    If sTipoPunto = "PI" Or sTipoPunto = "FI" Then
        MsgBox "Las notas de credito electronicas solo pueden ser Web-Service u On-Line"
        Exit Sub
    End If
End If

If UCase(cboTipo.Text) = "DEA" Or UCase(cboTipo.Text) = "DEB" Or UCase(cboTipo.Text) = "DEC" Then
    If sTipoPunto = "PI" Or sTipoPunto = "FI" Then
        MsgBox "Las notas de debito electronicas solo pueden ser Web-Service u On-Line"
        Exit Sub
    End If
End If


If sTipoPunto = "WS" Then
    nPermitoCae = 1
Else
    nPermitoCae = 0
End If

sPunto = Format(txtNumero, "0000")
If sTipo_cargado = "" Then
    sTipoPunto = sTipoPunto '& sPunto
Else
    sTipoPunto = sTipo_cargado
End If

nExiste = obtenerDeSQL("select idpermitido from documentoscae where tipo=" & ssTexto(cboTipo.Text) & " and puntoventa=" & ssTexto(txtNumero) & " and tipopunto=" & ssTexto(sTipoPunto))
If nExiste > 0 Then
    MsgBox "Punto de venta ya existe.", vbCritical
Else
    If id_en_vista > 0 Then
        MsgBox "Punto de venta no se puede modificar.", vbCritical
    Else
        sInsertUpdate = "INSERT INTO [dbo].[DocumentosCAE] ([TIPO],[DESCRIPCION],[PUNTOVENTA],[TIPOPUNTO],[CODFACTURA],[PERMITO_CAE]) VALUES " & _
        " (" & ssTexto(cboTipo.Text) & "," & ssTexto(txtDescripcion) & "," & ssTexto(txtNumero) & "," & ssTexto(sTipoPunto) & "," & ssTexto(lblCODFactura) & "," & nPermitoCae & ") "
        DataEnvironment1.Sistema.Execute sInsertUpdate
        MsgBox "Punto de venta guardado con exito.", vbInformation
    End If

End If
Limpio
uMenu.AceptarOk
End Sub



Private Sub uMenu_Buscar()
Dim resultado As Long, sTipo As String
Dim datoS
resultado = s2n(frmBuscar.MostrarSql("SELECT         IDPERMITIDO as ID, TIPO, DESCRIPCION, PUNTOVENTA, TIPOPUNTO, CODFACTURA, PERMITO_CAE FROM            DocumentosCAE "))
If resultado > 0 Then
    datoS = obtenerDeSQL("select * from documentoscae where idpermitido=" & resultado)
    id_en_vista = datoS(0)
    cboTipo.Text = datoS(1)
    
    txtNumero = datoS(3)
    lblCODFactura = datoS(5)
    sTipo = CORTO(CStr((datoS(4))), 0, Len(datoS(4)) - 2)
    Select Case sTipo
        Case "PI": cboTipoPunto.ListIndex = 0
        Case "OL": cboTipoPunto.ListIndex = 1
        Case "WS": cboTipoPunto.ListIndex = 2
    End Select
    txtDescripcion = datoS(2)
    sTipo_cargado = datoS(4)
    uMenu.BuscarOK
End If
End Sub

Private Sub uMenu_Cancelar()
Limpio
End Sub

Private Sub uMenu_eliminar()
Dim sDelete As String
Dim nExiste
    If id_en_vista > 0 Then
        
        nExiste = obtenerDeSQL("select count(codigo) as n from facturaventa where puntoventa=" & ssTexto(txtNumero))
        If nExiste > 0 Then
            MsgBox "Punto de venta utilizado en facturacion, no se puede eliminar.", vbCritical
        Else
            sDelete = "delete from documentoscae where idpermitido=" & id_en_vista
            DataEnvironment1.Sistema.Execute sDelete
            MsgBox "Punto de venta eliminado con exito.", vbInformation
        End If
        Limpio
        uMenu.EliminarOK
    End If
End Sub

Private Sub uMenu_Nuevo()
Limpio
End Sub

Private Sub uMenu_SALIR()
Unload Me
End Sub

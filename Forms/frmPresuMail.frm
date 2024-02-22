VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmPresuMail 
   Caption         =   "Envio de E-Mail"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5760
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton busca 
      Caption         =   "..."
      Height          =   375
      Left            =   5640
      TabIndex        =   16
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton Cancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   3240
      TabIndex        =   15
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton Enviar 
      Caption         =   "Enviar"
      Height          =   495
      Left            =   960
      TabIndex        =   14
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox txtAdjunto 
      Height          =   285
      Left            =   1440
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   4440
      Width           =   4095
   End
   Begin VB.TextBox txtMensaje 
      Height          =   1605
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "frmPresuMail.frx":0000
      Top             =   2640
      Width           =   4095
   End
   Begin VB.TextBox txtTitulo 
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   2160
      Width           =   4095
   End
   Begin VB.TextBox txtBCC 
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1680
      Width           =   4095
   End
   Begin VB.TextBox txtCC 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1200
      Width           =   4095
   End
   Begin VB.TextBox txtPara 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   720
      Width           =   4095
   End
   Begin VB.TextBox txtDe 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   4095
   End
   Begin MSMAPI.MAPIMessages MAPIMail 
      Left            =   5760
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession 
      Left            =   5760
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.Label Label7 
      Caption         =   "Adjunto :"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Mensaje :"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Titulo :"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "BCC :"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "CC :"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Para :"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "De :"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmPresuMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ruta As String

Private Sub Cancelar_Click()
    limpio
    Unload Me
End Sub

Private Function limpio()
    txtDe.Text = ""
    txtPara = ""
    txtCC.Text = ""
    txtBCC.Text = ""
    txtTitulo.Text = ""
    txtMensaje.Text = ""
    txtAdjunto.Text = ""
End Function

Private Sub Enviar_Click()
    
    Dim rs As New ADODB.Recordset
        
    'ruta = "C:\UDC Output Files\Presu.pdf" '.documentName 'App.Path
    
    
    'MsgBox "Debe crear el pdf antes de enviar.", , "ATENCION"
    'frmMail.Show
    'de
    'frmMail.Text1 = obtenerDeSQL("select mfrom from mailfeed where id=2")
    'para
    'frmMail.Text2 = sSinNull(obtenerDeSQL("select mail from prov where codigo=" & uProv.codigo))
    'frmMail.Text3 = obtenerDeSQL("select subject from mailfeed where id=2") & " " & txtcodigo
'                    MsgBox ""
    Dim a
    
    On Error Resume Next
    
    rs.Open "select mail from usuarios where codigo=" & frmPresupuesto.uEmisor.codigo, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If (rsDatosProvee.EOF = True And rsDatosProvee.BOF = True) Or IsNull(rsDatosProvee!mail) Or IsEmpty(rsDatosProvee!mail) Or rsDatosProvee!mail = "" Then
        MsgBox "No se enviara el correo debido a que el proveedor no contiene una direccion." & Chr(13) & "Para reenviar verifique los datos del proveedor y vuelva a imprimir la orden de compra.", , "ATENCION"
    Else
        MAPIMail.Compose
        
        MAPIMail.RecipIndex = 0
        MAPIMail.RecipType = 1
        'MAPIMail.RecipDisplayName = sSinNull(obtenerDeSQL("select mail from prov where codigo=" & uProv.codigo)) '"emeil@al.que.lo.envias"
        
'                    MAPIMail.RecipIndex = 1
'                    MAPIMail.RecipType = 2
'                    MAPIMail.RecipDisplayName = "german_dodge@hotmail.com" '"emeil@al.que.lo.envias"

'                    MAPIMail.RecipIndex = 2
'                    MAPIMail.RecipType = 3
'                    MAPIMail.RecipDisplayName = "diego@betasepp.com.ar" '"emeil@al.que.lo.envias"

        'el findwiondows me verifica si esta abierta la aplic. si es 0 no lo esta
        'MsgBox "" & FindWindow(vbNullString, "OP.pdf - Adobe Acrobat Professional")
        Call Cerrar_ventana("Presu.pdf - Adobe Acrobat Professional")
        If s2n(frmPresupuesto.uCliente.codigo) <> 0 Then
            MAPIMail.RecipAddress = IIf((sSinNull(obtenerDeSQL("select mail from clientes where codigo=" & frmPresupuesto.uCliente.codigo))) = "", "correo@correo.com", sSinNull(obtenerDeSQL("select mail from clientes where codigo=" & frmPresupuesto.uCliente.codigo))) '"emeil@al.que.lo.envias"
        Else
            MAPIMail.RecipAddress = IIf((sSinNull(obtenerDeSQL("select email from contacto where id=" & frmPresupuesto.uContacto.codigo))) = "", "correo@correo.com", sSinNull(obtenerDeSQL("select email from contacto where id=" & frmPresupuesto.uContacto.codigo))) '"emeil@al.que.lo.envias"
        End If
        MAPIMail.AddressResolveUI = True
        MAPIMail.ResolveName
        'MAPIMail.RecipType = mapToList
        
        MAPIMail.MsgSubject = "Presupuesto " & Trim(frmPresupuesto.TxtNro.Text)
        MAPIMail.MsgNoteText = Trim(txtMensaje.Text)
        MAPIMail.AttachmentPathName = Trim(txtAdjunto) '.Printer.filename
                                                
        MAPIMail.Send False 'con esto muestra la pantalla, con false lo envia directamente
    End If
    
    
    cierra = True
    Unload rptOrdenCompra

End Sub

Private Sub Form_Load()
    
    '************* correo(inicio variables) MAPI
    With MAPISession
        .DownLoadMail = False
        .NewSession = True
        .LogonUI = True
        .UserName = GetSetting("Programa enviador de mails", "Settings", "UserName")
        .Password = GetSetting("Programa enviador de mails", "Settings", "Password")
        .SignOn
        MAPIMail.SessionID = .SessionID
    End With

    '*******************************************************
    ruta = "C:\UDC Output Files\Presu.pdf" '.documentName 'App.Path
    limpio
    
    txtDe.Text = "presupuestos@bacigaluppi.com"
    If s2n(frmPresupuesto.uCliente.codigo) <> 0 Then
        txtPara = IIf((sSinNull(obtenerDeSQL("select mail from clientes where codigo=" & frmPresupuesto.uCliente.codigo))) = "", "correo@correo.com", sSinNull(obtenerDeSQL("select mail from clientes where codigo=" & frmPresupuesto.uCliente.codigo))) '"emeil@al.que.lo.envias"
    Else
        txtPara = IIf((sSinNull(obtenerDeSQL("select email from contacto where id=" & frmPresupuesto.uContacto.codigo))) = "", "correo@correo.com", sSinNull(obtenerDeSQL("select email from contacto where id=" & frmPresupuesto.uContacto.codigo))) '"emeil@al.que.lo.envias"
    End If
    txtCC.Text = "emilio@bacigaluppi.com"
    txtBCC.Text = obtenerDeSQL("select email from emisor where id='" & Trim(frmPresupuesto.uEmisor.codigo) & "'")
    txtTitulo.Text = "Envío Email del presupuesto requerido"
    txtMensaje.Text = "Adjuntamos un archivo pdf con el presupuesto requerido por uds." & Chr(13) & "Para visualizarlo debe contar con el utilitario Acrobat." & Chr(13) & "De no poseerlo puedebajarlo gratuitamente de la pagina web www.acrobat.com." & Chr(13) & "Sin otro particular saludamos a Ud. muy atte."
    txtAdjunto.Text = ruta
    
End Sub

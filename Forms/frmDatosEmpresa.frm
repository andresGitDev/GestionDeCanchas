VERSION 5.00
Begin VB.Form frmDatosEmpresa 
   Caption         =   "Datos de la empresa"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8040
   Icon            =   "frmDatosEmpresa.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAlias 
      Height          =   315
      Left            =   2775
      TabIndex        =   19
      Top             =   3585
      Width           =   3615
   End
   Begin VB.TextBox txtCBU 
      Height          =   315
      Left            =   2775
      TabIndex        =   17
      Top             =   3225
      Width           =   3615
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   975
      Left            =   6585
      Picture         =   "frmDatosEmpresa.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   150
      Width           =   1290
   End
   Begin VB.TextBox txtCaja 
      Height          =   315
      Left            =   2775
      TabIndex        =   15
      Top             =   2835
      Width           =   3615
   End
   Begin VB.TextBox txtIIBB 
      Height          =   315
      Left            =   2775
      TabIndex        =   14
      Top             =   2460
      Width           =   3615
   End
   Begin VB.TextBox txtIVA 
      Height          =   315
      Left            =   2775
      TabIndex        =   13
      Top             =   2085
      Width           =   3615
   End
   Begin VB.TextBox txtTel 
      Height          =   315
      Left            =   2775
      TabIndex        =   12
      Top             =   1710
      Width           =   3615
   End
   Begin VB.TextBox txtCuit 
      Height          =   315
      Left            =   2775
      TabIndex        =   11
      Top             =   1320
      Width           =   3615
   End
   Begin VB.TextBox txtDireccion 
      Height          =   315
      Left            =   2775
      TabIndex        =   10
      Top             =   930
      Width           =   3615
   End
   Begin VB.TextBox txtNombreC 
      Height          =   315
      Left            =   2775
      TabIndex        =   9
      Top             =   540
      Width           =   3615
   End
   Begin VB.TextBox txtNombre 
      Height          =   315
      Left            =   2775
      TabIndex        =   8
      Top             =   150
      Width           =   3615
   End
   Begin VB.Label Label10 
      Caption         =   "Alias"
      Height          =   300
      Left            =   195
      TabIndex        =   20
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Label Label9 
      Caption         =   "CBU"
      Height          =   300
      Left            =   195
      TabIndex        =   18
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Label Label8 
      Caption         =   "Caja Prev"
      Height          =   300
      Left            =   195
      TabIndex        =   7
      Top             =   2850
      Width           =   2535
   End
   Begin VB.Label Label7 
      Caption         =   "IIBB"
      Height          =   195
      Left            =   195
      TabIndex        =   6
      Top             =   2505
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "Iva"
      Height          =   300
      Left            =   195
      TabIndex        =   5
      Top             =   2145
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "Telefono"
      Height          =   285
      Left            =   195
      TabIndex        =   4
      Top             =   1755
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Cuit Empresa"
      Height          =   285
      Left            =   195
      TabIndex        =   3
      Top             =   1380
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Direccion"
      Height          =   285
      Left            =   180
      TabIndex        =   2
      Top             =   1005
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre corto para Listados"
      Height          =   285
      Left            =   195
      TabIndex        =   1
      Top             =   615
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre"
      Height          =   285
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frmDatosEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGuardar_Click()
g
End Sub

Private Sub Form_Load()
v
End Sub

Private Function g()
On Error GoTo ng
Dim t As String
t = "update datosempresa set " _
& "nombre=" & ssTexto(txtNombre) _
& ",nombrecortoparalistados=" & ssTexto(txtNombreC) _
& ",direccion=" & ssTexto(txtDireccion) _
& ",cuitempresa=" & ssTexto(txtCuit) _
& ",telefono=" & ssTexto(txtTel) _
& ",iva=" & ssTexto(txtIVA) _
& ",iibb=" & ssTexto(txtIIBB) _
& ",cajaprev=" & ssTexto(txtCaja) _
& ",cbu=" & ssTexto(txtCBU) _
& ",alias=" & ssTexto(txtAlias)
DataEnvironment1.Sistema.Execute t
MsgBox "Datos Guardados", vbInformation, "Datos Actulizados"
Exit Function
ng:
MsgBox "Error al guardar", vbCritical, "No se guardo los datos"
End Function

Private Function v()
Dim t
    t = obtenerDeSQL("select nombre,nombrecortoparalistados,direccion,cuitempresa,telefono,iva,iibb,cajaprev,cbu,alias from datosempresa")
    txtNombre = sSinNull(t(0))
    txtNombreC = sSinNull(t(1))
    txtDireccion = sSinNull(t(2))
    txtCuit = sSinNull(t(3))
    txtTel = sSinNull(t(4))
    txtIVA = sSinNull(t(5))
    txtIIBB = sSinNull(t(6))
    txtCaja = sSinNull(t(7))
    txtCBU = sSinNull(t(8))
    txtAlias = sSinNull(t(9))
End Function

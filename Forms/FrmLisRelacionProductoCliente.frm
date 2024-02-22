VERSION 5.00
Begin VB.Form FrmLisRelacionProductoCliente 
   Caption         =   "Listado de Relacion Producto x Cliente"
   ClientHeight    =   2025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7620
   Icon            =   "FrmLisRelacionProductoCliente.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2025
   ScaleWidth      =   7620
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcancelar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdaceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtdesc 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   4
      Tag             =   "2"
      Top             =   960
      Width           =   3975
   End
   Begin VB.CommandButton cmbayuda 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox txtcodigo 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Tag             =   "0"
      Top             =   960
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   7095
      Begin VB.OptionButton optprov 
         Alignment       =   1  'Right Justify
         Caption         =   "Por Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4080
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton OPtProd 
         Alignment       =   1  'Right Justify
         Caption         =   "Por Producto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   360
         TabIndex        =   0
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   1335
      Left            =   120
      Top             =   120
      Width           =   7455
   End
   Begin VB.Label lblhasta 
      Caption         =   "Codigo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "FrmLisRelacionProductoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbayuda_Click()
Dim resu As String
    If optprov.Value = True Then
        resu = frmBuscar.MostrarCodigoDescripcionActivo("clientes")
        
    Else
        resu = frmBuscar.MostrarCodigoDescripcionActivo("producto")
    End If
    If resu > "" Then
        txtcodigo = frmBuscar.resultado
        txtdesc = frmBuscar.resultado(2)
    End If
End Sub

Private Sub cmdAceptar_Click()
Dim STR As String
    If txtcodigo.Visible = True Then
        If optprov.Value = True Then
            STR = "SELECT Relacion_Producto_Cliente.*, Producto.descripcion,clientes.descripcion as clientes " _
            & " FROM Relacion_Producto_Cliente " _
            & " INNER JOIN Producto ON Relacion_Producto_Cliente.PRODUCTO = Producto.codigo " _
            & " inner join clientes on clientes.codigo=Relacion_Producto_Cliente.cliente " _
            & " where cliente= " & Val(txtcodigo) & " order by producto"
        Else
            If OPtProd.Value = True Then
                STR = "SELECT Relacion_Producto_Cliente.*, Producto.descripcion,clientes.descripcion as clientes " _
                & " FROM Relacion_Producto_Cliente " _
                & " INNER JOIN Producto ON Relacion_Producto_Cliente.PRODUCTO = Producto.codigo " _
                & " inner join clientes on clientes.codigo=Relacion_Producto_Cliente.cliente " _
                & " where producto='" & Trim(txtcodigo) & "'" '& " order by precio asc"
            Else
                MsgBox "Debe seleccionar un Producto o un Cliente", 48, "Atencion"
                Exit Sub
            End If
        End If
    Else
        If optprov.Value = True Then
            STR = "SELECT Relacion_Producto_Cliente.*, Producto.descripcion,clientes.descripcion as clientes " _
            & " FROM Relacion_Producto_Cliente " _
            & " INNER JOIN Producto ON Relacion_Producto_Cliente.PRODUCTO = Producto.codigo " _
            & " inner join clientes on clientes.codigo=Relacion_Producto_Cliente.cliente " _
            & " order by producto"
        Else
            If OPtProd.Value = True Then
                STR = "SELECT Relacion_Producto_Cliente.*, Producto.descripcion,clientes.descripcion as clientes " _
                & " FROM Relacion_Producto_Cliente " _
                & " INNER JOIN Producto ON Relacion_Producto_Cliente.PRODUCTO = Producto.codigo " _
                & " inner join clientes on clientes.codigo=Relacion_Producto_Cliente.cliente "
                '" order by precio asc"
            Else
                MsgBox "Debe seleccionar un Producto o un Cliente", 48, "Atencion"
                Exit Sub
            End If
        End If
    
    End If
    RptLisProductoxCliente.Data.Connection = DataEnvironment1.Sistema
    RptLisProductoxCliente.Data.Source = STR
    RptLisProductoxCliente.Show

End Sub

Private Sub cmdCancelar_Click()
    LimpioControles
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub

Private Sub Form_Load()
    LimpioControles
End Sub
Sub LimpioControles()
    optprov.Value = True
    OPtProd.Value = False
    txtcodigo = ""
    txtdesc = ""
End Sub

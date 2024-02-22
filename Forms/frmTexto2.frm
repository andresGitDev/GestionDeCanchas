VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmTexto2 
   Caption         =   "Descripcion"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton aceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton salir 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   3240
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Descripcion"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.CheckBox Check1 
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   840
         Width           =   375
      End
      Begin VB.CheckBox Check2 
         Caption         =   "K"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox Check3 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1800
         Width           =   375
      End
      Begin RichTextLib.RichTextBox txtDesc 
         Height          =   2175
         Left            =   1200
         TabIndex        =   4
         Top             =   480
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   3836
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmTexto2.frx":0000
      End
      Begin VB.Label Label2 
         Caption         =   "Descripcion :"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmTexto2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private midDoc As Long
Public Linea As Integer

Private Sub limpio()
    txtDesc.Text = ""
    Check1.Value = 0
    Check2.Value = 0
    Check3.Value = 0
End Sub

Private Function habilito(a As Boolean)
    Frame1.enabled = a
    
    aceptar.enabled = a
End Function

Private Sub aceptar_Click()
If ON_ERROR_HABILITADO Then On Error GoTo Errror
Dim str As String
    If txtDesc.Text = "" Then
        MsgBox "Debe ingresar una descripcion", , "ATENCION"
        Exit Sub
    End If
    
    DE_BeginTrans
    
'    DataEnvironment1.Sistema.Execute "insert into texto (titulo,descripcion,activo) values('" & Trim(txtTitu.TextRTF) & "','" & Trim(txtDesc.TextRTF) & "',1)"
    str = "insert into ItemPedidoCliente2Texto (pedido,codigo,descripcion) values(" & Trim(frmPresupuesto.txtNro) & "," & Linea & ",'" & Trim(txtDesc.TextRTF) & "')"
    DataEnvironment1.Sistema.Execute str
    frmPresupuesto.grillaproductos.TextMatrix(Linea, 1) = txtDesc.Text
    frmPresupuesto.grillaproductos.TextMatrix(Linea, 0) = txtDesc.Text
    frmPresupuesto.grillaproductos.TextMatrix(Linea, 2) = txtDesc.Text
    frmPresupuesto.grillaproductos.TextMatrix(Linea, 3) = txtDesc.Text
    frmPresupuesto.grillaproductos.TextMatrix(Linea, 4) = txtDesc.Text
    frmPresupuesto.grillaproductos.TextMatrix(Linea, 5) = txtDesc.Text
    frmPresupuesto.grillaproductos.TextMatrix(Linea, 6) = txtDesc.Text
    frmPresupuesto.grillaproductos.TextMatrix(Linea, 7) = txtDesc.Text
    frmPresupuesto.grillaproductos.TextMatrix(Linea, 8) = txtDesc.Text
    frmPresupuesto.grillaproductos.TextMatrix(Linea, 9) = txtDesc.Text
        
    DE_CommitTrans
        
    'limpio
    'habilito True
    Unload Me
    Exit Sub
Errror:
    DE_RollbackTrans
    MsgBox "Error al grabar texto", vbCritical
End Sub

Private Sub Check1_Click()
    fuente
End Sub

Private Sub fuente()
    If Check1.Value = 1 Then
        txtDesc.SelBold = True
    Else
        txtDesc.SelBold = False
    End If
    If Check2.Value = 1 Then
        txtDesc.SelItalic = True
    Else
        txtDesc.SelItalic = False
    End If
    If Check3.Value = 1 Then
        txtDesc.SelUnderline = True
    Else
        txtDesc.SelUnderline = False
    End If
End Sub

Private Sub Check2_Click()
    fuente
End Sub

Private Sub Check3_Click()
    fuente
End Sub

Private Sub Form_Load()
    limpio
    habilito True
End Sub

Private Sub salir_Click()
    Unload Me
End Sub

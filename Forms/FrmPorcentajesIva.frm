VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPorcentajesIva 
   Caption         =   "PORCENTAJES iVA"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6225
   Icon            =   "FrmPorcentajesIva.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1965
   ScaleWidth      =   6225
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker txtfechadesde 
      Height          =   300
      Left            =   1680
      TabIndex        =   2
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      Format          =   284229633
      CurrentDate     =   38088
   End
   Begin VB.ComboBox cmbiva 
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   2175
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
      Left            =   5115
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1515
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
      Left            =   1575
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1515
      Width           =   975
   End
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
      Left            =   2805
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1515
      Width           =   975
   End
   Begin VB.TextBox txtporc 
      Height          =   285
      Left            =   4560
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker txtfechahasta 
      Height          =   300
      Left            =   4560
      TabIndex        =   3
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   284229633
      CurrentDate     =   38088
   End
   Begin VB.Label Label4 
      Caption         =   "Fecha hasta :"
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
      Left            =   3360
      TabIndex        =   10
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha Desde :"
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
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   1335
      Left            =   120
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label Label2 
      Caption         =   "Porcentaje :"
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
      Left            =   3360
      TabIndex        =   8
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Iva :"
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
      Left            =   240
      TabIndex        =   7
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "FrmPorcentajesIva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbiva_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub cmdAceptar_Click()
Dim rsIva As New ADODB.Recordset
Dim fecha As Variant
    

    If cmbiva.ListIndex = -1 Then
        MsgBox "Debe completar un iva", 48, "Atencion"
        cmbiva.SetFocus
        Exit Sub
    Else
        If Val(txtporc) = 0 Then
            MsgBox "Debe cargar el porcentaje", 48, "Atencion"
            txtporc.SetFocus
            Exit Sub
        Else
            fecha = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
            
            rsIva.Open "Select * from PorcentajesIva where iva=" & ObtenerCodigo("Ivas", Trim(cmbiva.Text)) & " and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            If Not rsIva.EOF Then
                DataEnvironment1.Sistema.Execute "update PorcentajesIva set fecha_baja=" & ssFecha(Date) & ",usuario_baja=" & UsuarioSistema!codigo & ",activo=0 where iva=" & ObtenerCodigo("Ivas", Trim(cmbiva.Text)) & " and activo=1"
            End If
            rsIva.Close
            Set rsIva = Nothing
            DataEnvironment1.dbo_PORCENTAJE ObtenerCodigo("Ivas", Trim(cmbiva.Text)), CDbl(Replace(txtporc, ".", ",")), fecha, UsuarioSistema!codigo, 0, DateAdd("yyyy", 10, fecha)
            MsgBox "La Operacion se ha realizado con exito", 48, "Atencion"
            LimpioControles
        End If
    End If
End Sub

Private Sub cmdCancelar_Click()
   LimpioControles
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)

    FrmKeyPress KeyAscii, True, True

    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()

    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    LimpioControles
    CargaCombo cmbiva, "Ivas", "descripcion", "codigo", ""
    
End Sub

Sub LimpioControles()

    cmbiva.ListIndex = -1
    txtfechadesde = Date
    txtfechahasta = Date
    txtporc = "0.00"
    
End Sub


Private Sub txtfechadesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub


Private Sub txtfechahasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub txtporc_GotFocus()
    txtporc.SelStart = 0
    txtporc.SelLength = Len(txtporc.Text)
End Sub

Private Sub txtporc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

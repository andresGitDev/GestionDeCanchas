VERSION 5.00
Begin VB.Form FrmCambioClave 
   BorderStyle     =   0  'None
   Caption         =   "Cambio de Clave"
   ClientHeight    =   5385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3660
   Icon            =   "FrmCambioClave.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   3660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtClaveNueva 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   375
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   3045
      Width           =   2895
   End
   Begin VB.TextBox txtConfirmClave 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   375
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   4140
      Width           =   2895
   End
   Begin VB.Frame Frame2 
      Caption         =   "Clave Nueva"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   855
      Left            =   135
      TabIndex        =   6
      Top             =   2700
      Width           =   3375
   End
   Begin VB.TextBox txtClaveActual 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   375
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1620
      Width           =   2895
   End
   Begin VB.TextBox txtUsuario 
      Height          =   285
      Left            =   375
      TabIndex        =   2
      Top             =   540
      Width           =   2895
   End
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   405
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4850
      Width           =   990
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   405
      Left            =   2070
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4850
      Width           =   975
   End
   Begin VB.Frame Frame5 
      Caption         =   "Confirmación Clave Nueva"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   855
      Left            =   135
      TabIndex        =   9
      Top             =   3780
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Clave Actual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   135
      TabIndex        =   4
      Top             =   1260
      Width           =   3375
   End
   Begin VB.Frame Frame3 
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   135
      TabIndex        =   7
      Top             =   180
      Width           =   3375
   End
End
Attribute VB_Name = "FrmCambioClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'12/8/4  falta SP

Option Explicit

Dim rsUsuarioNuevo As New ADODB.Recordset

Private Sub cmdaceptar_Click()

If rsUsuarioNuevo.State = 0 Then
    rsUsuarioNuevo.Open "Select * from usuarios where usuario = '" & txtusuario & "'", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
End If

If Trim(txtClaveActual) <> rsUsuarioNuevo!Clave Then
    MsgBox "''Clave Actual'' incorrecta", vbExclamation + vbOKOnly, "Error en clave actual"
    txtClaveActual.BackColor = &H8080FF
    txtClaveActual.SetFocus
    Exit Sub
Else
    If txtClaveNueva = "" Then
        MsgBox "Debe ingresar la clave nueva", vbExclamation + vbOKOnly, "Error en clave nueva"
        txtClaveNueva.BackColor = &H8080FF
        txtClaveNueva.SetFocus
        Exit Sub
    Else
        If txtConfirmClave = "" Then
            MsgBox "Debe ingresar la confirmacion de la clave nueva", vbExclamation + vbOKOnly, "Error en la confirmacion de la clave nueva"
            txtConfirmClave.BackColor = &H8080FF
            txtConfirmClave.SetFocus
            Exit Sub
        End If
    End If
End If

If Trim(txtClaveNueva) = Trim(txtConfirmClave) Then
'    DataEnvironment1.dbo_MODIFICACIONCLAVE Trim(rsUsuarioNuevo!codigo), Trim(txtClaveNueva)
    MsgBox "La clave se cambio correctamente.", vbExclamation + vbOKOnly, "Actualizacion de clave"
    Unload Me
Else
    MsgBox "La confirmacion de la clave nueva es incorrecta", vbExclamation + vbOKOnly, "Error en la confirmacion de la clave nueva"
    txtConfirmClave.SetFocus
End If

End Sub

Private Sub cmdcancelar_Click()

    Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If

End Sub

Private Sub txtClaveActual_GotFocus()

    txtClaveActual.SelStart = 0
    txtClaveActual.SelLength = Len(txtClaveActual)

End Sub

Private Sub txtClaveNueva_GotFocus()

    txtClaveNueva.SelStart = 0
    txtClaveNueva.SelLength = Len(txtClaveNueva)

End Sub

Private Sub txtConfirmClave_GotFocus()

    txtConfirmClave.SelStart = 0
    txtConfirmClave.SelLength = Len(txtConfirmClave)

End Sub

Private Sub txtClaveActual_LostFocus()

    txtClaveActual.BackColor = &H80000005

End Sub

Private Sub txtClaveNueva_LostFocus()

    txtClaveNueva.BackColor = &H80000005

End Sub

Private Sub txtConfirmClave_LostFocus()

    txtConfirmClave.BackColor = &H80000005

End Sub

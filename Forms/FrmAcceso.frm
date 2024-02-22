VERSION 5.00
Begin VB.Form FrmAcceso 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Acceso al Sistema de Gestion"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4545
   Icon            =   "FrmAcceso.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcancelar 
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   3000
      MaskColor       =   &H8000000F&
      Picture         =   "FrmAcceso.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   630
   End
   Begin VB.TextBox txtclave 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1770
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   555
      Width           =   2535
   End
   Begin VB.TextBox txtusuario 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1770
      TabIndex        =   0
      Top             =   210
      Width           =   2535
   End
   Begin VB.CommandButton cmdaceptar 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   2370
      MaskColor       =   &H8000000F&
      Picture         =   "FrmAcceso.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   630
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   195
      Picture         =   "FrmAcceso.frx":1A5E
      ScaleHeight     =   540
      ScaleWidth      =   570
      TabIndex        =   4
      Top             =   240
      Width           =   570
   End
   Begin VB.Label Label4 
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   195
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Clave"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   855
      TabIndex        =   5
      Top             =   555
      Width           =   975
   End
End
Attribute VB_Name = "FrmAcceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Sub cmdaceptar_Click()
Dim rsUsuarioNuevo As New ADODB.Recordset
If (txtusuario.Text <> "") And (txtclave <> "") Then
    rsUsuarioNuevo.Open "Select * from usuarios where usuario = '" & txtusuario & _
                        "' and clave = '" & txtclave & _
                        "' and activo = 1", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
    If Not rsUsuarioNuevo.EOF Then
        If rsUsuarioNuevo!cambiarclave = True Then
            FrmCambioClave.Show
            FrmCambioClave.txtusuario.Locked = True
            FrmCambioClave.txtusuario = rsUsuarioNuevo!usuario
            FrmCambioClave.txtClaveActual.SetFocus
            FrmCambioClave.Hide
            FrmCambioClave.Show vbModal
            rsUsuarioNuevo.Close '
            Set rsUsuarioNuevo = Nothing
            Exit Sub
        End If

        If AccesoSistema(txtusuario, txtclave) = True Then
            FrmPrincipal.Show
            Me.Hide
        Else
            MsgBox "Clave Invalida", vbOKOnly, "Atencion"
            txtclave.SetFocus
            txtclave.SelStart = 0
            txtclave.SelLength = Len(txtclave)
        End If
    Else
    Dim hay
        hay = obtenerDeSQL("Select * from usuarios")
        If IsNull(hay) Or IsEmpty(hay) Then
            If MsgBox("No hay ningun usuario en el sistema." & Chr(13) & "¿Desea cargar uno nuevo?", vbInformation + vbYesNo) = vbYes Then
                FrmAbmUsuarios.Show vbModal
            End If
        Else
            MsgBox "Usuario o Clave Incorrecta", vbOKOnly, "Atencion"
            txtusuario.SetFocus
        End If
    End If
Else
    MsgBox "Debe Completar Todos Los Campos", vbOKOnly, "Atencion"
End If
End Sub

Private Sub cmdcancelar_Click()
    DataEnvironment1.Sistema.Close
    End
End Sub

Private Sub Form_Initialize()
    On Error Resume Next
    InitCommonControls
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub

Private Sub Form_Load()

'   No me animo, pero probar.... especialmente sistema que ejecuta de a 3.
'    If App.PrevInstance Then
'        che "ya hay una instancia ejecutandose"
'        End
'    End If
    'MsgBox Command()
    
    Dim x
    Dim q As Long, conex As String
    x = Trim(Command())
    If x > "" Then
        x = Split(Trim(Command()), " ")
        If IsArray(x) Then
            q = UBound(x)
            '   server  "Data Source"
            'DataEnvironment1.Sistema.Properties(14) = Trim(Left(x, InStr(x, " ")))
            'DataEnvironment1.Sistema.Properties("Data Source") = Trim(Left(x, InStr(x, " ")))
            If q >= 0 Then DataEnvironment1.Sistema.Properties("Data Source") = Trim(x(0))
            '   database "Initial Catalog"
            'DataEnvironment1.Sistema.Properties(19) = Trim(Mid(x, InStr(x, " ") + 1))
            'DataEnvironment1.Sistema.Properties("Initial Catalog") = Trim(Mid(x, InStr(x, " ") + 1))
            If q > 0 Then DataEnvironment1.Sistema.Properties("Initial Catalog") = Trim(x(1))
            If q > 1 Then DataEnvironment1.Sistema.Properties("User ID") = Trim(x(2))
            If q > 2 Then DataEnvironment1.Sistema.Properties("Password") = Trim(x(3))
        End If
    End If
    
    AbrirDB
    
    If s2n(txtusuario) > 0 And s2n(txtclave) > 0 Then
        cmdaceptar_Click
    End If
End Sub

Private Sub txtclave_GotFocus()
    txtclave.SelStart = 0
    txtclave.SelLength = Len(txtclave)
End Sub

Private Sub txtusuario_Change()
    Dim i As Long
    txtusuario.Text = UCase(txtusuario.Text)
    i = Len(txtusuario.Text)
    txtusuario.SelStart = i
End Sub

Private Sub txtusuario_GotFocus()
On Error Resume Next
    txtusuario.SelStart = 0
    txtusuario.SelLength = Len(txtusuario)
End Sub


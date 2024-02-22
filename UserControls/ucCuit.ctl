VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ucCuit 
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1470
   LockControls    =   -1  'True
   ScaleHeight     =   330
   ScaleWidth      =   1470
   Begin MSMask.MaskEdBox txtCuit 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   13
      Mask            =   "##-########-#"
      PromptChar      =   " "
   End
End
Attribute VB_Name = "ucCuit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Lito Explicit 16/11/4

' Verifica digito control
' .Text devuelve "" Si no es valido, no mat
' pinta rojo si incorrecto

Private mColorLetra As Long
'

Public Property Get Text()
Attribute Text.VB_UserMemId = 0
    Text = txtCuit
    'If Verificar() Then Text = txtCuit
End Property

Public Property Let Text(que)
   On Error Resume Next
   If que = "" Then txtCuit.Text = "  -        - "
   txtCuit.Text = Trim(que)
   'Verificar
End Property

Public Property Get enabled() As Boolean
    enabled = UserControl.enabled
End Property

Public Property Let enabled(como As Boolean)
    UserControl.enabled = como
    PropertyChanged "Enabled"
End Property

Private Sub txtCuit_GotFocus()
    txtCuit.ForeColor = mColorLetra
    GotFocusPinto txtCuit
End Sub

Private Sub txtCuit_LostFocus()
    Verificar
End Sub

Private Sub UserControl_Initialize()
    mColorLetra = txtCuit.ForeColor
End Sub

Private Sub UserControl_LostFocus()
    Verificar
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
     UserControl.enabled = PropBag.ReadProperty("Enabled", True)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Enabled", UserControl.enabled, True
End Sub

' ----------
Private Sub UserControl_Resize()
    txtCuit.Left = 0
    txtCuit.Top = 0
    txtCuit.Height = UserControl.Height
    txtCuit.Width = UserControl.Width
End Sub

Private Function Verificar()
    txtCuit.ForeColor = vbRed
    Verificar = False

'son numeros
    If Not IsNumeric(Trim(Mid(txtCuit, 1, 2))) Then Exit Function
    If Not IsNumeric(Trim(Mid(txtCuit, 4, 8))) Then Exit Function
    If Not IsNumeric(Trim(Mid(txtCuit, 13, 1))) Then Exit Function
    
' XX-           20, 27, 30, ... ?
    ' por las dudas no
    
' -X            digito verificador
    If digitoOK() Then
        txtCuit.ForeColor = mColorLetra
        Verificar = True
    End If
End Function

Public Function digitoOK()
    On Error GoTo ufaChe
    Dim aa As Variant, s As String, i As Long, n As Long
    digitoOK = False
    aa = Array(1000, 5, 4, 3, 2, 7, 6, 5, 4, 3, 2, 1)
    s = Left(txtCuit.Text, 2) & Mid(txtCuit.Text, 4, 8) & Right(txtCuit.Text, 1)
    n = 0
    For i = 1 To 10
        n = n + Val(Mid(s, i, 1)) * aa(i)
    Next
    If ((11 - (n Mod 11)) Mod 11) = Val(Right(s, 1)) Then digitoOK = True
    GoTo fin
ufaChe:
    digitoOK = False
fin:
End Function


'11/11/4    agregue enabled()
'16/11/4    permito cargarlo con ""
'15/12/4    r/w property enabled


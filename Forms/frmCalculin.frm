VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmCalculin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CalcuLin"
   ClientHeight    =   915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCalcular 
      Caption         =   "="
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   0
      Width           =   795
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   3720
      Top             =   420
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.CommandButton cmdEnviar 
      Caption         =   "Enviar"
      Height          =   435
      Left            =   1980
      TabIndex        =   2
      Top             =   420
      Width           =   795
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   435
      Left            =   2820
      TabIndex        =   3
      Top             =   420
      Width           =   795
   End
   Begin VB.TextBox txtEntrada 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "4 decimales"
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label lblMsg 
      Caption         =   "Ctl+Enter asigna resultado"
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   360
      Width           =   1995
   End
End
Attribute VB_Name = "frmCalculin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mResu As Variant, mCoco As Control

Private Sub cmdCalcular_Click()
    On Error GoTo fin
    Dim x, resu
    txtEntrada = s2n(ScriptControl1.Eval(x2s(txtEntrada)), 4)
    txtEntrada.SetFocus
    PintoFocoActivo
fin:
End Sub
Private Sub cmdcancelar_Click()
    mResu = Empty
    Unload Me
End Sub
Private Sub cmdEnviar_Click()
    On Error Resume Next
    mResu = s2n(txtEntrada, 4)
    If Not (mCoco Is Nothing) Then
        mCoco.Text = mResu ' no deberia ser necesario
        mCoco = mResu
    End If
    Unload Me
    mCoco.SetFocus
End Sub
Public Function mostrar(Optional que) ', Optional queControl As Control) As Variant
    On Error GoTo ufaChe
    Set mCoco = Screen.ActiveControl
    txtEntrada = s2n(mCoco) '.Text)
    If (mCoco Is Nothing) Then NoPuedo
    If (Not mCoco.enabled) Then NoPuedo
    If (Not mCoco.Visible) Then NoPuedo
    If (mCoco.Locked) Then NoPuedo
sigo:
    Me.Show vbModal
    mostrar = mResu
fin:
    Exit Function
ufaChe:
    NoPuedo
    Resume sigo
End Function
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 10 Then cmdEnviar.Value = True
End Sub
Private Sub txtEntrada_GotFocus()
    PintoFocoActivo
End Sub
Private Sub NoPuedo()
    Set mCoco = Nothing
    lblMsg.Visible = False
    cmdEnviar.Visible = False
End Sub

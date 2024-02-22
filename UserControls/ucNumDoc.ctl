VERSION 5.00
Begin VB.UserControl ucNumDoc 
   BackStyle       =   0  'Transparent
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2940
   ScaleHeight     =   330
   ScaleWidth      =   2940
   Begin VB.TextBox txtNum 
      Height          =   315
      Left            =   1350
      MaxLength       =   8
      TabIndex        =   2
      Top             =   0
      Width           =   1560
   End
   Begin VB.TextBox txtSuc 
      Height          =   315
      Left            =   720
      MaxLength       =   4
      TabIndex        =   1
      Top             =   0
      Width           =   600
   End
   Begin VB.TextBox txtLet 
      Height          =   315
      Left            =   0
      MaxLength       =   1
      TabIndex        =   0
      Top             =   0
      Width           =   390
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1365
      TabIndex        =   3
      Top             =   -45
      Width           =   180
   End
End
Attribute VB_Name = "ucNumDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mForecolor As Long
Private mVeri As String

Private mTipoDoc As eTipoDoc
Public Enum eTipoDoc
    TipoDocFactura
    TipoDocRemito
    TipoDocRecibo
End Enum


Public Property Let enabled(como As Boolean)
    txtNum.enabled = como
    txtLet.enabled = como
    txtsuc.enabled = como
End Property
Public Property Get enabled() As Boolean
    enabled = txtNum.enabled
End Property


Public Property Let TIPODOC(cual As eTipoDoc)
    mTipoDoc = cual
    Select Case mTipoDoc
     Case TipoDocFactura: mVeri = "ABCM": txtLet.Locked = False
     Case TipoDocRecibo:  mVeri = "X": txtLet = "X": txtLet.Locked = True
     Case TipoDocRemito:  mVeri = "R": txtLet = "R": txtLet.Locked = True
    End Select
    revisar
End Property
Public Property Get NumDocValido() As Boolean
    If suc = 0 Or num = 0 Then Exit Property
    If InStr(mVeri, letra) = 0 Then Exit Property
  
  
    NumDocValido = True
End Property



Public Property Get letra() As String
    letra = txtLet
End Property
Public Property Let letra(cual As String)
    'If InStr("ABCMRX", cual) = 0 Then txtLet.ForeColor = vbRed
    txtLet = cual
End Property



Public Property Get suc() As Long
    If IsNumeric(txtsuc) Then suc = CLng(txtsuc)
End Property
Public Property Get txtSucu() As String
    txtSucu = txtSucu
End Property
Public Property Let suc(cual As Long)
    If cual < 1 Or cual > 9999 Then
        txtsuc = ""
    Else
        txtsuc = Format(cual, "0000")
    End If
End Property


Public Property Get num() As Long
    If IsNumeric(txtNum) Then num = CLng(txtNum)
End Property
Public Property Get txtNumero() As Long
    txtNumero = txtNum
End Property
Public Property Let num(cual As Long)
    If cual < 1 Or cual > 99999999 Then
        txtNum = ""
    Else
        txtNum = Format(cual, "00000000")
    End If
End Property


Private Sub txtLet_LostFocus()
    txtLet = UCase(txtLet)
    If InStr(mVeri, txtLet) < 1 Then txtLet = " "
    revisar
End Sub
Private Sub txtNum_LostFocus()
    num = s2n(txtNum)
    revisar
End Sub
Private Sub txtSuc_LostFocus()
    suc = s2n(txtsuc)
    revisar
End Sub

Private Sub UserControl_Initialize()
    mForecolor = txtNum.ForeColor
    TIPODOC = TipoDocFactura
End Sub

Private Sub revisar()
    If NumDocValido Then
        txtLet.ForeColor = mForecolor
        txtNum.ForeColor = mForecolor
        txtsuc.ForeColor = mForecolor
    Else
        txtLet.ForeColor = vbRed
        txtNum.ForeColor = vbRed
        txtsuc.ForeColor = vbRed
    End If
End Sub

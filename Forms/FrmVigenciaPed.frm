VERSION 5.00
Begin VB.Form FrmVigenciaPed 
   Caption         =   "Vigenca de un Pedido"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5985
   Icon            =   "FrmVigenciaPed.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check 
      Alignment       =   1  'Right Justify
      Caption         =   "Por un periodo mayor al previsto anteriormente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1290
      TabIndex        =   1
      Top             =   1785
      Width           =   2550
   End
   Begin VB.Frame Frame1 
      Height          =   2340
      Left            =   60
      TabIndex        =   4
      Top             =   165
      Width           =   5775
      Begin VB.ComboBox CmbCant 
         Height          =   315
         Left            =   3990
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   967
         Width           =   1605
      End
      Begin VB.TextBox Mayor 
         Height          =   300
         Left            =   3990
         TabIndex        =   6
         Top             =   1710
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad de días actuales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   420
         TabIndex        =   8
         Top             =   270
         Width           =   2970
      End
      Begin VB.Label DiasAct 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   3990
         TabIndex        =   7
         Top             =   240
         Width           =   1605
      End
      Begin VB.Image Image1 
         Height          =   750
         Left            =   135
         Picture         =   "FrmVigenciaPed.frx":030A
         Stretch         =   -1  'True
         Top             =   1485
         Width           =   810
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccione las cantidad de días para la vigencia de un Pedido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   195
         TabIndex        =   5
         Top             =   825
         Width           =   3450
      End
   End
   Begin VB.CommandButton Cmd_Salir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4590
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2745
      Width           =   1215
   End
   Begin VB.CommandButton CmdAceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3045
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2745
      Width           =   1215
   End
End
Attribute VB_Name = "FrmVigenciaPed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check_Click()
If Check.Value = 1 Then
   Mayor.Visible = True
   CmbCant.enabled = False
   Mayor.SetFocus
Else: Mayor.Visible = False
      CmbCant.enabled = True
End If
End Sub

Private Sub Cmd_Salir_Click()
Unload Me
End Sub

Private Sub cmdaceptar_Click()
Dim sql, valor As String
If MsgBox("Confirma los datos ingresados", vbQuestion + vbYesNo, "Confirmacion") = vbYes Then
   If Check.Value = 1 Then
      valor = Mayor
   Else
      valor = CmbCant.Text
   End If
   sql = "UPDATE bs SET diasvenc= '" & valor & "'"
   DataEnvironment1.Sistema.Execute sql
   DiasAct = valor
End If
End Sub

Private Sub Form_Load()
Dim x As Long, sql As String, rs As New ADODB.Recordset

For x = 1 To 60
   CmbCant.AddItem x
Next
sql = "SELECT diasvenc FROM bs"
rs.Open sql, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
If IsNull(rs!diasvenc) Then
 DiasAct = 0
Else
   DiasAct = rs!diasvenc
End If
End Sub

Private Sub Mayor_KeyPress(KeyAscii As Integer)
' entre 48 y 57
 If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
   'no hago nada
 Else
   MsgBox "Solo debe ingresar digitos numericos", vbInformation, "Aviso"
   Seleccion Mayor
 End If
End Sub
Private Sub Seleccion(captura As Object)
    captura.SetFocus
    SendKeys "{home}+{end}"
End Sub


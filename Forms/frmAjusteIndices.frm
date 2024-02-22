VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAjusteIndices 
   Caption         =   "Ajuste - Indices"
   ClientHeight    =   2010
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   ScaleHeight     =   2010
   ScaleWidth      =   4785
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   345
      Left            =   3180
      TabIndex        =   6
      Top             =   1275
      Width           =   1320
   End
   Begin VB.TextBox txtIndice 
      Height          =   345
      Left            =   780
      TabIndex        =   5
      Text            =   "0"
      Top             =   1290
      Width           =   1095
   End
   Begin VB.CheckBox chkAnual 
      Caption         =   "Anual"
      Height          =   360
      Left            =   795
      TabIndex        =   3
      Top             =   810
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker dtMes 
      Height          =   330
      Left            =   795
      TabIndex        =   0
      Top             =   420
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "MMMM"
      Format          =   127336451
      CurrentDate     =   43532
   End
   Begin MSComCtl2.DTPicker dtAnio 
      Height          =   330
      Left            =   1920
      TabIndex        =   2
      Top             =   420
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "yyyy"
      Format          =   127336451
      CurrentDate     =   43532
   End
   Begin VB.Label Label2 
      Caption         =   "Indice"
      Height          =   300
      Left            =   75
      TabIndex        =   4
      Top             =   1290
      Width           =   675
   End
   Begin VB.Label Label1 
      Caption         =   "Periodo"
      Height          =   330
      Left            =   75
      TabIndex        =   1
      Top             =   450
      Width           =   975
   End
End
Attribute VB_Name = "frmAjusteIndices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub chkAnual_Click()
If chkAnual Then
    dtMes.enabled = False
Else
    dtMes.enabled = True
End If
Ver
End Sub

Private Sub cmdGuardar_Click()
Dim AjusteGuardar As New AjusteInflacion
If AjusteGuardar.Guardar(qMes(), qAnio(), s2n(txtIndice, 4), chkAnual) Then
    MsgBox "Guardado", vbInformation
End If
End Sub

Private Sub dtAnio_Change()
dtMes = dtAnio
Ver
End Sub

Private Sub dtMes_Change()
dtAnio = dtMes
Ver
End Sub

Private Sub Form_Load()
dtMes = Date
dtAnio = Date
Ver
End Sub

Private Sub txtIndice_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNum(KeyAscii, True, False)
End Sub

Private Function Ver()
Dim AjusteVer As New AjusteInflacion
txtIndice = AjusteVer.buscar(qMes(), qAnio(), chkAnual)
End Function

Private Function qMes() As Integer
qMes = Month(dtMes)
End Function

Private Function qAnio() As Integer
qAnio = Year(dtAnio)
End Function


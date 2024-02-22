VERSION 5.00
Begin VB.Form frmTiempo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestion"
   ClientHeight    =   870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   6345
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Caption         =   "Aguarde un momento mientras se actualiza la base de datos......"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   390
      TabIndex        =   0
      Top             =   330
      Width           =   5595
   End
End
Attribute VB_Name = "frmTiempo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    'DataEnvironment1.AMR.Close
    'Timer1.Interval = 1 'son los 15 seg * 1000
    'Do While Time < Ctiempo
    '    DoEvents   ' Cambia a otros procesos.
    'Loop
    'Me.Show
End Sub

'Private Sub Timer1_Timer()
    'Unload Me
    'frmTiempo.Visible = True
    'If ctiempo < Format(Time, "hh:mm:ss") Then
    '   Unload Me
    'End If
'End Sub

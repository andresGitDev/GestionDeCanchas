VERSION 5.00
Begin VB.Form frmBackup 
   Caption         =   "Backup"
   ClientHeight    =   2385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2385
   ScaleWidth      =   7860
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   405
      Left            =   5295
      TabIndex        =   5
      Top             =   1770
      Width           =   2220
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ejecutar"
      Height          =   390
      Left            =   2505
      TabIndex        =   4
      Top             =   1785
      Width           =   2145
   End
   Begin VB.TextBox txtBackup 
      Height          =   375
      Left            =   2430
      TabIndex        =   2
      Top             =   1035
      Width           =   5055
   End
   Begin VB.TextBox txtBaseDato 
      Height          =   375
      Left            =   2430
      TabIndex        =   0
      Top             =   495
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   "Archivo Destino               (Disco C: del server de datos)"
      Height          =   450
      Index           =   1
      Left            =   195
      TabIndex        =   3
      Top             =   1005
      Width           =   2190
   End
   Begin VB.Label Label1 
      Caption         =   "Base de datos "
      Height          =   270
      Index           =   0
      Left            =   1215
      TabIndex        =   1
      Top             =   585
      Width           =   1155
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    If ON_ERROR_HABILITADO Then On Error GoTo ufa
    
    
    relojito True
    
    BackupSql txtBaseDato, txtBackup

fin:
    relojito False
    Exit Sub
ufa:
    ufa "Fallo el backup", ""
    Resume fin
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    txtBaseDato = DataEnvironment1.Sistema.Properties.Item("initial catalog")
    txtBackup = "C:\"
    If carpeta() > "" Then txtBackup = carpeta()
End Sub

Private Function carpeta() As String
On Error GoTo fin
carpeta = VerParametro(BS_CARPETA_BACKUP)
fin:
End Function

VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmsk 
   BackColor       =   &H00004000&
   Caption         =   "Sk"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6270
   FillColor       =   &H00C0C0C0&
   FillStyle       =   5  'Downward Diagonal
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5295
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C000C0&
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3375
      MaskColor       =   &H00800080&
      TabIndex        =   8
      Top             =   3585
      Width           =   1485
   End
   Begin VB.ListBox List1 
      BackColor       =   &H0000C000&
      Height          =   1035
      Left            =   3285
      TabIndex        =   7
      Top             =   2190
      Width           =   2475
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2025
      Left            =   210
      TabIndex        =   6
      Top             =   3000
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   3572
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   32768
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmsk.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H0000C000&
      Height          =   315
      Left            =   3300
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   1590
      Width           =   2445
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0000C000&
      Caption         =   "Check1"
      Height          =   405
      Left            =   3315
      TabIndex        =   4
      Top             =   825
      Width           =   2190
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0000C000&
      Caption         =   "Option1"
      Height          =   375
      Left            =   3300
      TabIndex        =   3
      Top             =   315
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Caption         =   "Frame1"
      Height          =   1305
      Left            =   240
      TabIndex        =   2
      Top             =   1425
      Width           =   2400
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFF80&
      Height          =   345
      Left            =   255
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   855
      Width           =   2235
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "Label1"
      Height          =   315
      Left            =   255
      TabIndex        =   0
      Top             =   330
      Width           =   2190
   End
End
Attribute VB_Name = "frmsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub getSkinForm(queform As Form)
    On Error Resume Next
    With queform
        .BackColor = Me.BackColor
        .Picture = Me.Picture
        getSkin queform
    End With
End Sub
Public Sub getSkinUsr(queUsr As Object)
    On Error Resume Next
    With queUsr
        .BackColor = Me.BackColor
        .Picture = Me.Picture
        getSkin queUsr
    End With
End Sub

'

Private Sub getSkin(que As Object)
'    On Error Resume Next
    Dim con As Control
    With que
        For Each con In que.Controls
            If TypeOf con Is TextBox Then decoroTxt con
            If TypeOf con Is Label Then decoroLbl con
            If TypeOf con Is OptionButton Then decoroOpt con
            If TypeOf con Is CheckBox Then decoroChk con
            If TypeOf con Is Frame Then decoroFra con
            If TypeOf con Is ComboBox Then decoroCbo con
            If TypeOf con Is ListBox Then decoroLst con
            If TypeOf con Is SSTab Then decoroSst con
            If TypeOf con Is CommandButton Then decoroCmd con
            'If TypeOf con Is  Then decoro (con)
        Next
    End With
End Sub

Private Sub decoroLbl(co As Label)
    With Label1
        co.BackColor = .BackColor
        co.Font = .Font
        co.ForeColor = .ForeColor
    End With
End Sub
Private Sub decoroTxt(co As TextBox)
    With Text1
        co.BackColor = .BackColor
        co.Font = .Font
        co.ForeColor = .ForeColor
    End With
End Sub
Private Sub decoroOpt(co As OptionButton)
    With Option1
        co.BackColor = .BackColor
        co.Font = .Font
        co.ForeColor = .ForeColor
    End With
End Sub
Private Sub decoroChk(co As CheckBox)
    With Check1
        co.BackColor = .BackColor
        co.Font = .Font
        co.ForeColor = .ForeColor
    End With
End Sub
Private Sub decoroFra(co As Frame)
    With Frame1
        co.BackColor = .BackColor
        co.Font = .Font
        co.ForeColor = .ForeColor
    End With

End Sub
Private Sub decoroCbo(co As ComboBox)
    With Combo1
        co.BackColor = .BackColor
        co.Font = .Font
        co.ForeColor = .ForeColor
    End With
End Sub
Private Sub decoroLst(co As ListBox)
    With List1
        co.BackColor = .BackColor
        co.Font = .Font
        co.ForeColor = .ForeColor
    End With
End Sub
Private Sub decoroSst(co As SSTab)
    With SSTab1
        co.BackColor = .BackColor
        co.Font = .Font
        co.ForeColor = .ForeColor
    End With
End Sub
Private Sub decoroCmd(co As CommandButton)
    With Command1
        co.BackColor = .BackColor
        co.Font = .Font
        'co.ForeColor = .ForeColor
        co.Picture = .Picture
        co.DisabledPicture = .DisabledPicture
        co.DownPicture = .DownPicture
    End With
End Sub



Private Sub Form_Load()
    'cargoskin
End Sub

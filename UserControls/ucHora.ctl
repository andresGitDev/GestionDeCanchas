VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ucHora 
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   660
   ScaleHeight     =   315
   ScaleWidth      =   660
   Begin MSMask.MaskEdBox hora 
      Height          =   300
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   5
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
End
Attribute VB_Name = "ucHora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form Buscar 
   Caption         =   "Buscar"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   7290
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraTool 
      BorderStyle     =   0  'None
      Caption         =   "-"
      Height          =   420
      Left            =   105
      TabIndex        =   1
      Top             =   5430
      Width           =   7095
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   300
         Left            =   6075
         TabIndex        =   4
         Top             =   45
         Width           =   990
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Default         =   -1  'True
         Height          =   300
         Left            =   4965
         TabIndex        =   3
         Top             =   60
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Pruebe el boton secundario sobre campo"
         Height          =   195
         Left            =   840
         TabIndex        =   2
         Top             =   0
         Width           =   5055
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid g 
      Align           =   1  'Align Top
      Bindings        =   "Buscar.frx":0000
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7290
      _cx             =   12859
      _cy             =   9340
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   3
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "Buscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'Public Function resultado(Optional que) 'Devuelve string campo[que] (desde 1)
'    On Error GoTo fin
'
'    Static re() As String
'    Dim I As Integer
'
'    resultado = ""
'
'    If IsMissing(que) Then que = 1
'    If que = -1 Then
'        ReDim re(pCols)
'        For I = 1 To pCols
'            re(I) = a(I)
'        Next I
'    Else
'        resultado = re(que)
'    End If
'fin:
'End Function

 
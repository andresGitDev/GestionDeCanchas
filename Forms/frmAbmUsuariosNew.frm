VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmAbmUsuariosNew 
   Caption         =   "Usuarios"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   10035
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRecargar 
      Caption         =   "refrescar grilla"
      Height          =   300
      Left            =   5265
      TabIndex        =   3
      Top             =   165
      Width           =   1590
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "Modificar"
      Height          =   330
      Left            =   3600
      TabIndex        =   2
      Top             =   135
      Width           =   1485
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   315
      Left            =   8385
      TabIndex        =   1
      Top             =   150
      Width           =   1425
   End
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   9780
      _cx             =   17251
      _cy             =   11668
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
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
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
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "frmAbmUsuariosNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum grises
    griCOD
    griDES
    griUSU
    griCLA
    griTIP
    griDIR
    griTEL
    griLOC
    griPOR
    griCLI
End Enum
    
    
    
Private Sub cmdRecargar_Click()
    recargar
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    recargar
    Form_Resize
End Sub

Private Sub Form_Resize()
    Anclar grilla, Me, anclarLadosTodos
End Sub

Private Sub recargar()
    Dim s As String
    s = " SELECT u.codigo, u.descripcion, u.usuario, u.clave, t.descripcion AS Tipo, " & _
        " u.direccion, u.telefono, u.localidad, u.porcentaje, u.ABMClientes " & _
        " FROM         Usuarios u LEFT OUTER JOIN " & _
        " TipoUsuarios t ON u.tipousuario = t.codigo "

    LlenarGrilla grilla, s, False
    grillaWidth grilla, Array(615, 1110, 1275, 750, 1545, 780, 1215, 840, 840, 990)
End Sub

Private Sub grilla_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
     Case griDES, griUSU, griCLA, griDIR, griTEL, griLOC, griPOR
     Case Else: Cancel = True
    End Select
End Sub

Private Sub grilla_Click()
    If Not grilla.Editable Then Exit Sub
    
End Sub
  8  cm
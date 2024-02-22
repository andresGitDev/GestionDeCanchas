VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmBancosGrilla 
   Caption         =   "formulario temporal..."
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   9960
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   5115
      Left            =   5940
      TabIndex        =   1
      Top             =   120
      Width           =   3915
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "GRABAR"
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   60
         Width           =   1515
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "salir"
         Height          =   435
         Left            =   0
         TabIndex        =   3
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton cmdcancelar 
         Caption         =   "cancelar"
         Height          =   435
         Left            =   0
         TabIndex        =   2
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Ordenado por CODIGO"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Width           =   1155
      End
      Begin VB.Label Label5 
         Caption         =   "Para modificar, selecciona campo y apreta F2"
         Height          =   555
         Left            =   1860
         TabIndex        =   9
         Top             =   660
         Width           =   1995
      End
      Begin VB.Label Label4 
         Caption         =   $"frmBancosGrilla.frx":0000
         Height          =   975
         Left            =   480
         TabIndex        =   8
         Top             =   3360
         Width           =   3015
      End
      Begin VB.Label Label3 
         Caption         =   "Para agregar uno nuevo, ponele codigo (que no este usado), descripcion y activo = 1"
         Height          =   1335
         Left            =   1800
         TabIndex        =   7
         Top             =   2220
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "No cambies codigo, porque el codigo original no se modificara ni borrara"
         Height          =   675
         Left            =   1800
         TabIndex        =   6
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Para desactivar un banco, poner activo en 0"
         Height          =   555
         Left            =   1860
         TabIndex        =   5
         Top             =   0
         Width           =   2115
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid gri 
      Height          =   6195
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      _cx             =   10186
      _cy             =   10927
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
      AllowUserResizing=   0
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
Attribute VB_Name = "frmBancosGrilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents g As LiGrilla
Attribute g.VB_VarHelpID = -1

Private gCODI As Long
Private gDESC As Long
Private gACTI As Long

Private Sub cmdaceptar_Click()
    Dim i As Long, C As Long, d As String, a As Long, co As Long
    
    For i = 1 To g.rows - 1
        C = s2n(g.tx(i, gCODI))
        d = g.tx(i, gDESC)
        a = s2n(g.tx(i, gACTI))
        
        If C > 0 Then
            co = nSinNull(obtenerDeSQL("select codigo from bancosgrales where codigo = '" & C & "'"))
            If co = 0 Then
                agregar C, d, a
            Else
                modificar C, d, a
            End If
        End If
    Next i
    carga
End Sub

Private Sub cmdcancelar_Click()
    carga
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set g = New LiGrilla
    g.init gri
    
    gCODI = g.AddCol(" Codigo ", "N", 0)
    gDESC = g.AddCol(" Descripcion                                                       ", "S")
    gACTI = g.AddCol("Activo", "N", 0)
    carga
    Form_Resize
End Sub

Private Sub carga()
    Dim rs As New ADODB.Recordset, i As Long
    g.Borrar
    rs.Open "select * from bancosgrales order by codigo", DataEnvironment1.AMR, adOpenForwardOnly, adLockReadOnly
    While Not rs.EOF
        i = g.addRow()
        g.tx i, gCODI, rs!codigo
        g.tx i, gDESC, rs!descripcion
        g.tx i, gACTI, IIf(rs!activo, 1, 0)
        rs.MoveNext
    Wend
    g.rows = g.rows + 20
End Sub

Private Sub agregar(co, de, ac)
    On Error GoTo ufa
    If ac = 1 Then
        DataEnvironment1.dbo_BANCOGRAL "A", co, de, Date, UsuarioActual(), 0, CDate("1/1/0")
    Else
        che "agregar en 0"
    End If
fin:
    Exit Sub
ufa:
    che "error al agregar" & co & " " & de
    Resume fin
End Sub

Private Sub modificar(co, de, ac)
    On Error GoTo ufa
    If ac = 1 Then
        DataEnvironment1.dbo_BANCOGRAL "M", co, de, Date, UsuarioActual(), 0, CDate("1/1/0")
    Else
        DataEnvironment1.dbo_BANCOGRAL "B", co, de, Date, 0, UsuarioActual(), Date
    End If
fin:
    Exit Sub
ufa:
    che "error al modificar"
    Resume fin
End Sub

Private Sub Form_Resize()
    Anclar gri, Me, anclarLadosTodos
    Anclar fra, Me, anclarDerecha + anclarArriba
End Sub



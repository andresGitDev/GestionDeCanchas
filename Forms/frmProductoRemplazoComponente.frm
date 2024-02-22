VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmProductoRemplazoComponente 
   Caption         =   "Reemplazo de componentes"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10275
   Icon            =   "frmProductoRemplazoComponente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   10275
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGo 
      Caption         =   "Procesar"
      Height          =   345
      Left            =   2505
      TabIndex        =   3
      Top             =   885
      Width           =   1215
   End
   Begin Gestion.ucCoDe uProdVie 
      Height          =   330
      Left            =   2460
      TabIndex        =   1
      Top             =   75
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   582
      CodigoWidth     =   1000
   End
   Begin VSFlex7LCtl.VSFlexGrid gri 
      Height          =   6135
      Left            =   135
      TabIndex        =   0
      Top             =   1305
      Width           =   9165
      _cx             =   16166
      _cy             =   10821
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
      Editable        =   2
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
   Begin Gestion.ucCoDe uProdNue 
      Height          =   330
      Left            =   2460
      TabIndex        =   2
      Top             =   480
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   582
      CodigoWidth     =   1000
   End
   Begin VB.Label Label1 
      Caption         =   "Reemplazar en FORMULAS por :"
      Height          =   285
      Index           =   1
      Left            =   60
      TabIndex        =   5
      Top             =   555
      Width           =   2445
   End
   Begin VB.Label Label1 
      Caption         =   "Componente original :"
      Height          =   285
      Index           =   0
      Left            =   60
      TabIndex        =   4
      Top             =   150
      Width           =   2355
   End
End
Attribute VB_Name = "frmProductoRemplazoComponente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum colu
    colCODI
    colDESC
    colCOMP
    colCant
    colchk
End Enum


Private Sub cmdGo_Click()

    'OJO al modificar este form, puede perder todas las formulas

    Dim x As Long, r As Long, i As Long
    Dim fo As String, covi As String, conu As String
    

    covi = uProdVie.codigo
    conu = uProdNue.codigo
    
    ' verifica...
    
    If covi = "" Or conu = "" Then Exit Sub
    If covi = conu Then Exit Sub
    
    For i = 1 To gri.rows - 1
        If Trim(gri.TextMatrix(i, colchk)) > "" Then x = x + 1
    Next i
    If x = 0 Then Exit Sub
    
    
    ' todo ok
    If confirma("reemplaza el componente en " & x & " formulas ") Then
        
        For i = 1 To gri.rows - 1
        
            If Trim(gri.TextMatrix(i, colchk)) > "" Then
            
                fo = gri.TextMatrix(i, colCODI)
                DataEnvironment1.Sistema.Execute _
                    " Update formulas set componente = '" & conu & "' " & _
                    " where codigo = '" & fo & "' and componente = '" & covi & "' "
        
            End If
        Next i
    
        che "hecho"
        rellena
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub

Private Sub Form_Load()
    uProdVie.ini "select descripcion from producto where codigo = '###' ", "select codigo as [ Codigo        ],descripcion as [ Descripcion                ], Estado from producto ", True
    uProdNue.ini "select descripcion from producto where codigo = '###' ", "select codigo as [ Codigo        ],descripcion as [ Descripcion                ], Estado from producto ", True
End Sub

Private Sub Form_Resize()
    Anclar gri, Me, anclarLadosTodos
'    Anclar cmdGo, Me, anclarAbajo + anclarDerecha
End Sub


Private Sub gri_Click()
    Dim c As Long, r As Long
    c = gri.Col
    r = gri.Row
    
    If c = colchk And r > 0 Then
        gri.TextMatrix(r, c) = IIf(Trim(gri.TextMatrix(r, c)) = "", "SI", "")
    End If
    
End Sub

Private Sub uProdVie_cambio(codigo As Variant)
    rellena
End Sub

Private Sub rellena()
    LlenarGrilla gri, _
        " select formulas.Codigo as [ Codigo          ], Descripcion , componente, cantidad, ' ' as [REEMPLAZA?]  " & _
        " from formulas inner join producto on formulas.codigo = producto.codigo " & _
        " where componente = '" & uProdVie.codigo & "' ", False
    grillaWidth gri, Array(1630, 2300, 1500, 800, 1300)
End Sub

VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmAbmCuentas 
   Caption         =   "Plan de Cuentas"
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9045
   Icon            =   "frmAbmCuentas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8745
   ScaleWidth      =   9045
   StartUpPosition =   2  'CenterScreen
   Begin Gestion.ucCoDe uSumariza 
      Height          =   255
      Left            =   1680
      TabIndex        =   22
      Top             =   3105
      Width           =   6225
      _extentx        =   10980
      _extenty        =   450
      codigowidth     =   1000
   End
   Begin Gestion.ucXls uXls 
      Height          =   960
      Left            =   2325
      TabIndex        =   21
      Top             =   7665
      Width           =   990
      _extentx        =   1746
      _extenty        =   1693
   End
   Begin VB.CheckBox chkSalto 
      Alignment       =   1  'Right Justify
      Caption         =   "Salto"
      Height          =   375
      Left            =   5280
      TabIndex        =   16
      Top             =   2220
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CheckBox chkMonetaria 
      Alignment       =   1  'Right Justify
      Caption         =   "Monetaria"
      Height          =   375
      Left            =   3120
      TabIndex        =   15
      Top             =   2220
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CheckBox chkImputable 
      Alignment       =   1  'Right Justify
      Caption         =   "Imputable"
      Height          =   375
      Left            =   720
      TabIndex        =   14
      Top             =   2220
      Width           =   1155
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   945
      Left            =   3315
      Picture         =   "frmAbmCuentas.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7665
      Width           =   825
   End
   Begin VB.TextBox txtRenglon 
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   2700
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "Guardar Modificación"
      Enabled         =   0   'False
      Height          =   945
      Left            =   1290
      Picture         =   "frmAbmCuentas.frx":0C54
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7665
      Width           =   1035
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   315
      Left            =   1980
      TabIndex        =   3
      Top             =   1800
      Width           =   5895
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Guardar Nueva cuenta"
      Enabled         =   0   'False
      Height          =   930
      Left            =   135
      Picture         =   "frmAbmCuentas.frx":151E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7680
      Width           =   1140
   End
   Begin VB.TextBox txtCodigo 
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Top             =   1800
      Width           =   1395
   End
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Height          =   4035
      Left            =   120
      TabIndex        =   0
      Top             =   3540
      Width           =   8835
      _cx             =   15584
      _cy             =   7117
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
   Begin VB.Label lblCodigo 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   3
      Left            =   1365
      TabIndex        =   20
      Top             =   150
      Width           =   915
   End
   Begin VB.Label lblDescripcion 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   3
      Left            =   2385
      TabIndex        =   19
      Top             =   150
      Width           =   5505
   End
   Begin VB.Label Label3 
      Caption         =   "Cuenta Raiz"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   45
      TabIndex        =   18
      Top             =   165
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Sumariza"
      Height          =   255
      Left            =   660
      TabIndex        =   17
      Top             =   3120
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Renglones en blanco despues de imprimir"
      Height          =   435
      Index           =   0
      Left            =   60
      TabIndex        =   13
      Top             =   2580
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label lblDescripcion 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   2
      Left            =   1980
      TabIndex        =   10
      Top             =   1380
      Width           =   5895
   End
   Begin VB.Label lblCodigo 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   2
      Left            =   960
      TabIndex        =   9
      Top             =   1380
      Width           =   915
   End
   Begin VB.Label lblDescripcion 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   1
      Left            =   1980
      TabIndex        =   8
      Top             =   1020
      Width           =   5895
   End
   Begin VB.Label lblCodigo 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   1
      Left            =   960
      TabIndex        =   7
      Top             =   1020
      Width           =   915
   End
   Begin VB.Label lblDescripcion 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   0
      Left            =   1980
      TabIndex        =   6
      Top             =   660
      Width           =   5895
   End
   Begin VB.Label lblCodigo 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   0
      Left            =   960
      TabIndex        =   5
      Top             =   660
      Width           =   915
   End
End
Attribute VB_Name = "frmAbmCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private g As LiGrilla
Private gID As Long
Private gCODI As Long
Private gDESC As Long
Private gIMPU As Long

Private Sub CargaCuentas()
    Dim rs As New ADODB.Recordset, i As Long
    
    g.Borrar
    With rs
            .Open "select id, cuenta , Descripcion, Imputable from cuentas order by cuenta", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        While Not .EOF
            i = g.addRow()
            g.tx i, gID, !ID
            g.tx i, gCODI, !Cuenta
            g.tx i, gDESC, !DESCRIPCION
            g.tx i, gIMPU, IIf(!IMPUTABLE, "SI", "")
            .MoveNext
        Wend
    End With
End Sub

Private Sub cmdAgregar_Click()
    Dim impu As Long, cantL As Integer, COD As String
    
    If Trim(txtCodigo) = "" Then Exit Sub
    If Trim(txtDescripcion) = "" Then Exit Sub
    
    cantL = Len(Trim(lblCodigo(2).caption)) + 1
    COD = Mid(txtCodigo, cantL, Len(Trim(txtCodigo)))
    impu = IIf(chkImputable.Value = vbChecked, 1, 0)

    DataEnvironment1.Sistema.Execute "insert into cuentas (cuenta,_codigo, descripcion, activo, salto, imputable, sumariza, Usuario_Alta, fecha_Alta, Monetaria ) values ( '" & txtCodigo & "','" & COD & "','" & txtDescripcion & "', 1 ,0,  " & x2s(impu) & ",'" & Trim(lblCodigo(2).caption) & "', " & UsuarioActual() & ", " & ssFecha(Date) & ", 1 )"
    Actualizar
    che "Agregado"
End Sub

Private Sub cmdmodificar_Click()
    Dim impu As Long
    
    If Trim(txtCodigo) = "" Then Exit Sub
    If Trim(txtDescripcion) = "" Then Exit Sub
    If Not confirma("Graba cambios") Then Exit Sub
    
    impu = IIf(chkImputable.Value = vbChecked, 1, 0)

    DataEnvironment1.Sistema.Execute "update cuentas set descripcion = '" & Left(txtDescripcion, 50) & "', imputable = " & x2s(impu) & ", sumariza  = '" & Trim(lblCodigo(2).caption) & "' where cuenta = '" & Trim(txtCodigo) & "'"

    Actualizar
    che "Modificado"
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, False, False
End Sub

Private Sub Form_Load()
    inigrilla
    uSumariza.ini "select descripcion from Cuentas where cuenta = '###' and activo = 1", "select cuenta as [ Cuenta       ], descripcion as [ Descripcion               ] from Cuentas where activo = 1 ", True
    Actualizar
    uXls.ini grilla, App.Path & "\PlanCuentas.xls", "Plan de cuentas"
End Sub

Private Sub RevisoCodigo()
    Dim co As String, i As Long
    co = Trim(txtCodigo)
    
    txtDescripcion = ""
    llenolabel co
    
    With g
        For i = 1 To .rows - 1
            If co = Trim(.TextMatrix(i, gCODI)) Then
                txtDescripcion = .tx(i, gDESC)
                chkImputable.Value = IIf(.tx(i, gIMPU) > "", vbChecked, vbUnchecked)
                
                grilla.TopRow = i
                grilla.Select i, 0
                
                botones "M"
                Exit Sub
            End If
            
        Next i
        botones "A"
        chkImputable.Value = IIf(Len(co) = 7 And s2n(Right(co, 4)) > 0, vbChecked, vbUnchecked)
    End With
End Sub
Private Sub llenolabel(COD As String)
If ON_ERROR_HABILITADO Then On Error GoTo ufa
    Dim cta As String
    Dim i As Long
    Dim r As Long
    
    r = 1
    For i = 2 To 0 Step -1
        lblCodigo(i).caption = ""
        lblDescripcion(i).caption = ""
        
        If Len(COD) > i Then
            cta = Mid(COD, 1, (Len(COD) - r))
            Do While descripcionDe(cta) = ""
                r = r + 1
                cta = Mid(COD, 1, (Len(COD) - r))
            Loop
            r = r + 1
            lblCodigo(i).caption = cta 'Left(cod, i + 1)
            lblDescripcion(i).caption = descripcionDe(cta)
        End If
    Next i
    lblCodigo(3).caption = Left(COD, 1)
    lblDescripcion(3).caption = descripcionDe(Trim(lblCodigo(3).caption))
fin:
    Exit Sub
ufa:
    Resume fin
End Sub

Private Function descripcionDe(codigo As String) As String
    On Error Resume Next
    Dim tempo
        tempo = obtenerDeSQL("select descripcion from cuentas where cuenta = '" & codigo & "'")
    descripcionDe = sSinNull(tempo)
End Function

Private Sub botones(quehago As String)
    cmdmodificar.enabled = False
    cmdAgregar.enabled = False

    Select Case quehago
     Case "A"
        cmdAgregar.enabled = True
     Case "M"
        cmdmodificar.enabled = True
    End Select
End Sub

Private Sub Actualizar()
    CargaCuentas
    RevisoCodigo
End Sub

Private Sub Form_Resize()
    Anclar cmdsalir, Me, anclarAbajo + anclarDerecha
    Anclar cmdmodificar, Me, anclarAbajo + anclarDerecha
    Anclar cmdAgregar, Me, anclarAbajo + anclarDerecha
    Anclar uXls, Me, anclarAbajo + anclarDerecha
    Anclar grilla, Me, anclarLadosTodos
End Sub

Private Sub grilla_DblClick()
    txtCodigo = g.tx(g.Row, gCODI)
End Sub

Private Sub txtCodigo_Change()
    If Len(txtCodigo) > 7 Then txtCodigo = Left(txtCodigo, 7)
    RevisoCodigo
End Sub

Private Sub inigrilla()
    Set g = New LiGrilla
    With g
        .init grilla
        gID = .AddCol("id", "H")
        gCODI = .AddCol(" Cuenta                ")
        gDESC = .AddCol(" Descripcion                          ")
        gIMPU = .AddCol(" Imputable ")
    End With
    grilla.ColAlignment(gCODI) = flexAlignLeftCenter
End Sub

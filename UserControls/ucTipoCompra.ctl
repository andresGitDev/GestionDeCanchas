VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.UserControl ucTipoCompra 
   Alignable       =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   3570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11340
   ScaleHeight     =   3570
   ScaleWidth      =   11340
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3585
      Left            =   9120
      TabIndex        =   1
      Top             =   -75
      Width           =   2175
      Begin VB.TextBox txtFalta 
         Height          =   315
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2130
         Width           =   1440
      End
      Begin VB.CommandButton cmdBorraItem 
         Caption         =   "Borra item"
         Height          =   855
         Left            =   0
         Picture         =   "ucTipoCompra.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   510
         Width           =   855
      End
      Begin VB.TextBox txtTotal 
         Height          =   315
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1770
         Width           =   1440
      End
      Begin VB.TextBox txtTotalaImputar 
         Height          =   315
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1410
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Label Label2 
         Caption         =   "Dbl clic en CODIGO trae ayuda cuentas"
         Height          =   435
         Index           =   1
         Left            =   0
         TabIndex        =   10
         Top             =   3210
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Dbl clic en importe vacio trae faltante"
         Height          =   435
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   2730
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Falta"
         Height          =   255
         Index           =   2
         Left            =   1500
         TabIndex        =   7
         Top             =   2130
         Width           =   675
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total "
         Height          =   255
         Index           =   0
         Left            =   1500
         TabIndex        =   6
         Top             =   1830
         Width           =   675
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Imputar"
         Height          =   255
         Index           =   1
         Left            =   1500
         TabIndex        =   5
         Top             =   1470
         Visible         =   0   'False
         Width           =   675
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Align           =   3  'Align Left
      Height          =   3570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9075
      _cx             =   16007
      _cy             =   6297
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
Attribute VB_Name = "ucTipoCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'NEW 15/6/5
'23/1/6 Mod Tonka, carga cta contable, no van aca las de sistema.

Private WithEvents g As LiGrilla
Attribute g.VB_VarHelpID = -1
Private mRows As Long
Private mUsoCuenta As Long
Public mProv As Long

'Private gIDTC As Long
'Private gCODI As Long
Private gCUEN As Long
Private gDESC As Long
Private gMONT As Long
Private gPROG As Long
'Private gCONC As Long

'Public Property Get imId(row As Long) As Long
'    imId = s2n(g.tx(row, gIDTC))
'End Property
Public Property Get imCuenta(Row As Long) As String
    imCuenta = g.tx(Row, gCUEN)
End Property
Public Property Get imMonto(Row As Long) As Double
    imMonto = g.tx(Row, gMONT)
End Property
'Public Property Get imCodigo(row As Long) As Integer
'    imCodigo = s2n(g.tx(row, gCODI))
'End Property
Public Property Get Total() As Double
    Total = s2n(txttotal)
End Property
Public Property Get Total_a_Imputar() As Double
    Total_a_Imputar = s2n(txtTotalaImputar)
End Property
Public Property Let Total_a_Imputar(que As Double)
   txtTotalaImputar = que
   recalculo
End Property
Public Property Get Diferencia() As Double
    Diferencia = s2n(txtFalta) 's2n(s2n(txtTotalaImputar) - s2n(txtTotal))
End Property
Public Property Let UsoCuenta(cual As Long)
    mUsoCuenta = cual
End Property

Public Property Get rows() As Long
    recalculo
    rows = mRows
End Property

Public Sub agregar(Cuenta As String, monto As Double, Sistema As Boolean, Optional Acumulo As Boolean = False)
    Dim i As Long, v
    If Trim(Cuenta) = "" Or monto = 0 Then Exit Sub
    With g
        i = .buscar(gCUEN, Cuenta)
        If i = 0 Then i = .PrimerVacio(gCUEN)
        
        .tx i, gDESC, sSinNull(obtenerDeSQL("select descripcion from cuentas where cuenta = '" & Cuenta & "' "))
        .tx i, gCUEN, Cuenta
        .tx i, gPROG, IIf(Sistema, "1", "")
        If Acumulo Then
            v = 0
            v = .tx(i, gMONT)
            v = s2n(v)
        End If
        .tx i, gMONT, monto + v
    End With
End Sub


'Public Function agregarPorId(idTipoCompra As Long, monto As Double)
'    On Error Resume Next
'    Dim i As Long, tempo
'    With g
'        i = .Buscar(gIDTC, idTipoCompra)
'        If i = 0 Then i = .PrimerVacio(gIDTC)
'
'        tempo = obtenerDeSQL("select codigo, descripcion, cuenta, sistema from CuentasParam where id =  " & idTipoCompra)
'        .tx i, gIDTC, idTipoCompra
'        .tx i, gCODI, tempo(0)
'        .tx i, gDESC, tempo(1)
'        .tx i, gCUEN, tempo(2)
'        .tx i, gPROG, IIf(tempo(3), "1", "")
'        .tx i, gMONT, monto
'    End With
'End Function



'-------------------------------------------------
Private Sub g_DblClick()
    Dim re As String

    If g.Row = 0 Then Exit Sub
    If g.tx(g.Row, gPROG) > "" Then Exit Sub
    
    
    If g.Col = gMONT Then
        If s2n(g.Text) = 0 Then g.Text = s2n(txtTotalaImputar) - s2n(txttotal)
    End If
    
    If g.Col = gCUEN Then
        re = BuscarCuenta(False, False, mProv)
        If re > "" Then g.Text = re
    End If
End Sub

Private Sub g_Validar(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
    'cuenta
    If Col = gCUEN Then
        If IsEmpty(obtenerDeSQL("select cuenta, descripcion from Cuentas where cuenta = '" & g.EditText & "' and activo = 1 and imputable = 1 ")) Then
            cancel = True
            Exit Sub
        End If
    End If
'    If col = gCODI Then
'        If IsEmpty(obtenerDeSQL("select cuenta, descripcion from CuentasParam where codigo = " & x2s(s2n(g.EditText)) & " and sistema = 0 and UsoCuenta = '" & mUsoCuenta & "' ")) Then
'            cancel = True
'            Exit Sub
'        End If
'    End If
'    Dim tx As String, re As String
'    If col = gCUEN Then
'        tx = g.EditText
'        If Trim(tx) = "" Then
'            Exit Sub
'        Else
'            re = CuentaDescripcion(tx, False, False)
'            If re = "" Then
'                che "no es cuenta activa o imputable"
'                cancel = True
'            End If
'        End If
'    End If
End Sub

Private Sub UserControl_Initialize()
    On Error Resume Next
    inigrilla
    recalculo
    mUsoCuenta = 2 'compra
End Sub


Public Function Borrar()
    g.Borrar
    txttotal = ""
    txtFalta = ""
    mRows = 0
    g.rows = 2 'TOT_ROWS
End Function
Public Property Let enabled(sino As Boolean)
    UserControl.enabled = sino
End Property

Private Sub cmdBorraItem_Click()
    If g.rows > 2 Then
        If g.tx(g.Row, gPROG) > "" Then
            che "No puedo eliminar item agregado por sistema"
            Exit Sub
        End If

        g.delRow g.Row
        recalculo
    End If
End Sub

Private Sub g_BeforeEdit(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
    If g.tx(Row, gPROG) > "" Then cancel = True
End Sub

Private Sub g_cambio(ByVal Row As Long, ByVal Col As Long, txt As String)
    Dim tempo
    If Col = gCUEN Then
        'tempo = obtenerDeSQL("select cuenta, descripcion, id from CuentasParam where codigo = '" & txt & "' and UsoCuenta = '" & mUsoCuenta & "' ")
        tempo = obtenerDeSQL("select descripcion,cuenta from Cuentas where cuenta = '" & txt & "' and activo  = 1 and imputable = 1")
        If IsEmpty(tempo) Then Exit Sub
        
        g.tx Row, gDESC, sSinNull(tempo(0))
        'g.tx row, gCUEN, sSinNull(tempo(0))
        'g.tx row, gIDTC, nSinNull(tempo(2))
    End If
'    If col = gMONTO Then
'        recalculo
'    ElseIf col = gPRoTE Then
''        g.tx row, gN_INT, 0
''        g.tx row, gMONTO, 0
        recalculo
'    End If
    If g.rows < Row + 2 Then g.rows = g.rows + 1
End Sub

Private Sub inigrilla()
Dim aux_n
    Set g = New LiGrilla
    g.init GRILLA
    With g
'        gIDTC = .AddCol("id", "H")
        gMONT = .AddCol(" Importe         ", "N", 2)
'        gCODI = .AddCol(" Codigo ", "N", 0)
        gCUEN = .AddCol(" Cuenta              ", "S") ' ,"H")
        gDESC = .AddCol(" Descripcion                                 ")
        gPROG = .AddCol("prg") ', "H")
        'gCONC = .AddCol(" Concepto                                                        ", "S")
        aux_n = .editOk(0, 0)
    End With
    
    mRows = 0
End Sub

Private Sub grilla_GotFocus()
    If GRILLA.Row = 0 Then GRILLA.Select 1, 0
    'grilla.set
End Sub

Private Sub recalculo()
    Dim i As Long, suma As Double
    For i = 1 To g.rows - 1
'        If s2n(g.tx(i, gCODI)) = 0 Or s2n(g.tx(i, gMONT)) = 0 Then Exit For
        If g.tx(i, gCUEN) = "" Then Exit For
        suma = suma + s2n(g.tx(i, gMONT))
    Next i
    mRows = i - 1
    txttotal = suma 'g.suma(gMONTO)
    txtFalta = s2n(s2n(txtTotalaImputar) - s2n(suma))
End Sub

Private Sub UserControl_Resize()
    Anclar fra, Me, anclarDerecha + anclarArriba
    Anclar GRILLA, Me, anclarLadosTodos
End Sub

Private Sub UserControl_Show()
    On Error Resume Next
    Dim i As Long
    'UserControl.BackColor = UserControl.ParentControls(0).BackColor
    For i = 0 To Label1.Count - 1
        'Label1(i).BackColor = UserControl.BackColor
    Next i
End Sub

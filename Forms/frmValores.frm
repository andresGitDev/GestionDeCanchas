VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmValores 
   Caption         =   "Valores: Efectivo y cheques"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   8865
   Icon            =   "frmValores.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5085
   ScaleWidth      =   8865
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAgregar 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   7140
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmValores.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1020
      Width           =   495
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   4440
      TabIndex        =   11
      Top             =   4500
      Width           =   795
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   3360
      TabIndex        =   10
      Top             =   4500
      Width           =   795
   End
   Begin VB.TextBox txtCuentaEfectivo 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4620
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   660
      Width           =   1395
   End
   Begin VB.TextBox txtCaja 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   320
      Left            =   3840
      TabIndex        =   3
      Text            =   "1"
      Top             =   660
      Width           =   675
   End
   Begin VB.TextBox txtEfectivo 
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   660
      Width           =   1275
   End
   Begin VB.CommandButton cmdBorraItem 
      BackColor       =   &H00E0E0E0&
      Height          =   435
      Left            =   7740
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmValores.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Borrar Item"
      Top             =   1020
      Width           =   435
   End
   Begin VB.TextBox txtTotal 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   1275
   End
   Begin VSFlex7LCtl.VSFlexGrid gCheques 
      Height          =   2805
      Left            =   0
      TabIndex        =   5
      Top             =   1590
      Width           =   8790
      _cx             =   15505
      _cy             =   4948
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
      ScrollBars      =   0
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
   Begin VB.Label Label5 
      Caption         =   "Caja :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   3
      Left            =   3000
      TabIndex        =   9
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "en Efectivo :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   1155
   End
   Begin VB.Label Label5 
      Caption         =   "en Cheques :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   6
      Top             =   120
      Width           =   1155
   End
End
Attribute VB_Name = "frmValores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' To do ?
' llenar previamente ? (para mod, o mostrar recibos a cunta)


Option Explicit

Private mOK As Boolean
Private WithEvents g As LiGrilla
Attribute g.VB_VarHelpID = -1

Private gBANCC  As Long
Private gBANCD  As Long
Private gNROCH  As Long
Private gMONTO  As Long
Private gFECHA  As Long
Private gPT     As Long
'Private gCODCH  As Long
'

Public Function mostrar(Total As Double) As Boolean '
    mOK = False
    inigrilla
    Limpiar
    txtcaja = 1
    verCajaEfectivo
    txtTotal = s2n(Total)
    Me.Show vbModal
    mostrar = mOK
End Function

Public Function Efectivo() As Double
    Efectivo = s2n(TxtEfectivo)
End Function
Public Function caja() As Long
    caja = s2n(txtcaja)
End Function
Public Function cuenta() As String
    cuenta = txtCuentaEfectivo
End Function

Public Function ChCant() As Long
    ChCant = g.PrimerVacio(gNROCH) - 1
End Function
Public Function chNumero(i As Long) As Long
    If MalI(i) Then Exit Function
    chNumero = s2n(g.tx(i, gNROCH))
End Function
Public Function chCodBanco(i As Long) As Long
    If MalI(i) Then Exit Function
    chCodBanco = s2n(g.tx(i, gBANCC))
End Function
Public Function ChBanco(i As Long) As String
    If MalI(i) Then Exit Function
    ChBanco = Trim(g.tx(i, gBANCD))
End Function
Public Function chMonto(i As Long) As Double
    If MalI(i) Then Exit Function
    chMonto = s2n(g.tx(i, gMONTO))
End Function
Public Function chFecha(i As Long) As Date
    On Error Resume Next
    If MalI(i) Then Exit Function
    chFecha = CDate(g.tx(i, gFECHA))
End Function
Public Function chPT(i As Long) As String
    If MalI(i) Then Exit Function
    chPT = Trim(g.tx(i, gPT))
End Function

Private Function MalI(i As Long) As Boolean
    MalI = (i = 0 Or i > ChCant())
    If MalI Then ufa "prg: codigo cheque fuera de rango", "frmValores utCheque fuera de rango" & i ', Err
End Function

Private Sub cmdAceptar_Click()
    If Falta Then Exit Sub
    mOK = True
    Me.Hide
End Sub

Private Sub cmdAgregar_Click()
    g.addRow
End Sub

Private Sub cmdBorraItem_Click()
    If g.Row > 1 Then g.delRow (g.Row)
End Sub

Private Sub inigrilla()
    Set g = New LiGrilla
    With g
        .init gCheques
        gBANCC = .AddCol(" Banco ", "N")
        gBANCD = .AddCol("  Banco                        ")
        gNROCH = .AddCol("  Nro Cheque      ", "S")
        gMONTO = .AddCol("  Monto     ", "N")
        gFECHA = .AddCol("  Fecha     ", "D")
        gPT = .AddCol(" P/T ", "S")
'        gCODCH = .AddCol("Cod Interno")
    End With
End Sub

Private Sub cmdCancelar_Click()
    Me.Hide
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True, True
End Sub

'Private Sub Form_Load()
'    CentrarMe Me
''    txtEfectivo.SetFocus
'End Sub

' grilla -----------------------------------
Private Sub g_cambio(ByVal Row As Long, ByVal Col As Long, txt As String)
    Select Case Col
    Case gBANCC: g.tx Row, gBANCD, ObtenerDescripcion("BancosGrales", s2n(txt))
    
    End Select
End Sub
Private Sub g_DblClick()
    If g.Col = gBANCC Then g.tx g.Row, g.Col, frmBuscar.MostrarCodigoDescripcionActivo("BancosGrales")
End Sub
Private Sub g_Validar(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case gPT: Cancel = (g.EditText <> "P" And g.EditText <> "T")
        
    End Select
End Sub

Private Sub txtcaja_GotFocus()
    If Trim$(txtcaja) = "" Then txtcaja = "1"
    PintoFocoActivo
End Sub

Private Sub txtCaja_Validate(Cancel As Boolean)
    Cancel = Not verCajaEfectivo()
End Sub
'------------------------------

Private Sub txtEfectivo_LostFocus()
    Dim sino As Boolean
    sino = (s2n(TxtEfectivo) <> 0)
    txtcaja.enabled = sino
    txtCuentaEfectivo.enabled = sino
End Sub

Private Sub txtEfectivo_Validate(Cancel As Boolean)
    If Not IsNumeric(TxtEfectivo) Then Cancel = True
    TxtEfectivo = s2n(TxtEfectivo)
    If s2n(TxtEfectivo) = 0 Then
        txtcaja.enabled = False
        txtCuentaEfectivo.enabled = False
    End If
End Sub

Private Sub Limpiar()
    g.Borrar
    g.rows = 40
    txtcaja = 1
    verCajaEfectivo
End Sub

Private Function verCajaEfectivo() As Boolean
    Dim tmp As String
    tmp = obtenerDeSQL("select cuenta from cajas where codigo = " & s2n(txtcaja))
    If Not IsEmpty(tmp) Then ' > "" Then
        verCajaEfectivo = True
        txtCuentaEfectivo = tmp
    Else
        che "No existe la caja"
        verCajaEfectivo = False
    End If
End Function

Private Sub che(que)
    MsgBox que, vbExclamation, "Aviso"
End Sub
Private Function Falta() As Boolean
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    Dim i As Long, f As Date
    Falta = True
    
    'zero
    If s2n(txtTotal) = 0 Then 'and
        'ufa "prg: Total = 0", "Se cargo total 0 -- " & Me.Name, Err ' no se porque lo puse como err de prg!,
        che " Se cargo total 0"
        Exit Function
    End If
    'grilla
    i = g.PrimerVacio(gBANCD)
    If i <> g.PrimerVacio(gNROCH) Or i <> g.PrimerVacio(gMONTO) Or i <> g.PrimerVacio(gFECHA) Or i <> g.PrimerVacio(gPT) Then
        che "revisar datos en grilla"
        Exit Function
    End If
    'montos
    If s2n(s2n(g.suma(gMONTO), 2) + s2n(TxtEfectivo)) <> s2n(txtTotal, 2) Then
        che "No coinciden los montos " & vbCrLf & " tot= " & s2n(txtTotal) & vbCrLf & " cheques " & s2n(g.suma(gMONTO), 2) & ", efectivo = " & s2n(TxtEfectivo, 2)
        Exit Function
    End If
    If s2n(TxtEfectivo) <> 0 And Trim(txtCuentaEfectivo) = "" Then
        che "revisar cuenta caja efectivo"
        Exit Function
    End If
    
    'fechas
    For i = 1 To ChCant()
        f = CDate(g.tx(i, gFECHA)) ' si error...
        'falta verificar q sea fecha razonable, no 1800, a 20 años, etc
    Next i
      
    Falta = False
    GoTo fin
    
ufaErr:
    Falta = True
    che "err: Posible problema de fechas"
fin:
End Function




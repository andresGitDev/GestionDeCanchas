VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.UserControl ucCheques 
   ClientHeight    =   2070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10860
   ScaleHeight     =   2070
   ScaleWidth      =   10860
   Begin VB.CommandButton cmdBorraItem 
      Caption         =   "Borra item"
      Height          =   915
      Left            =   9000
      Picture         =   "ucCheques.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   60
      Width           =   885
   End
   Begin VB.TextBox txtTotal 
      Height          =   315
      Left            =   9000
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1050
      Width           =   915
   End
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Align           =   3  'Align Left
      Height          =   2070
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8970
      _cx             =   15822
      _cy             =   3651
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
   Begin VB.Label Label1 
      Caption         =   "Elija Propio o Tercero  Doble clic en Nro Interno"
      Height          =   435
      Index           =   1
      Left            =   9015
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.Label Label1 
      Caption         =   "Pago con Cheques:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   9090
      TabIndex        =   2
      Top             =   90
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label Label1 
      Caption         =   "Total "
      Height          =   255
      Index           =   0
      Left            =   10050
      TabIndex        =   1
      Top             =   1095
      Visible         =   0   'False
      Width           =   705
   End
End
Attribute VB_Name = "ucCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Lito Explicit 29/3/2005 ' Cheques pago

Private WithEvents g As LiGrilla
Attribute g.VB_VarHelpID = -1
Private mRows As Long
'Private Const TOT_ROWS = 20

Private gN_INT As Long
Private gPRoTE As Long
Private gBAN_C As Long
Private gBAN_D As Long
Private gNUMER As Long
Private gMONTO As Long
Private gFECHA As Long
Private gC_CON As Long

'Private gCU_BK As Long
Public Event cambio()


Private mNroIntPropioFueModificado As Boolean
'

Public Property Get Total() As Double
    Total = s2n(txttotal)
End Property

Public Property Get rows() As Long
    recalculo
    rows = mRows
End Property

Public Function Borrar()
    g.Borrar
    txttotal = ""
    mRows = 0
    g.rows = 2 'TOT_ROWS
    g.tx 1, gPRoTE, "P"
End Function

Public Function metoCheque(Row As Long, NroInt As Long, PT As String)
    Dim tempo, ss As String
    Dim Cuenta As String
    
    Select Case PT
    Case "P"
        ss = "Select Nro, banco, BancosGrales.Descripcion, Importe, fecha_cheque, cuentabancaria from Chq_comp inner join bancosGrales on BancosGrales.codigo = banco  where Chq_comp.activo = 1 and chq_comp.codigo = " & NroInt
    Case "T"
        ss = "Select Nro, Banco_Nro, BancosGrales.Descripcion, Importe, Fecha from Cheques inner join BancosGrales on Banco_Nro = BancosGrales.codigo where cheques.activo = 1  and cheques.NroInt = " & NroInt
    Case "N"
        
    End Select
    
    
    If PT <> "N" Then
        tempo = obtenerDeSQL(ss)
        
        If IsEmpty(tempo) Then
            ufa "err al buscar cheque", "metocheque() " & NroInt
            Exit Function
        End If
        
        If PT = "P" Then
            Cuenta = sSinNull(obtenerDeSQL("select cuenta_con from CtasBank where codigo = '" & x2s(tempo(5)) & "' "))
        Else
            Cuenta = CuentaParam(ID_Cuenta_M_CH_CARTERA)
        End If
        
        g.tx Row, gPRoTE, PT
        g.tx Row, gN_INT, NroInt
        g.tx Row, gBAN_C, tempo(1)
        g.tx Row, gBAN_D, tempo(2)
        g.tx Row, gNUMER, tempo(0)
        g.tx Row, gMONTO, tempo(3)
        g.tx Row, gFECHA, tempo(4)
        g.tx Row, gC_CON, Cuenta
        
        recalculo
    Else
        'g.tx Row, gPRoTE, "P"
        g.tx Row, gN_INT, 0
        g.tx Row, gBAN_C, 1
        'g.tx Row, gBAN_D, ""
        'g.tx Row, gNUMER, tempo(0)
        'g.tx Row, gMONTO, tempo(3)
        'g.tx Row, gFECHA, tempo(4)
        If g.tx(Row, gPRoTE) = "P" Then
            g.tx Row, gC_CON, sSinNull(obtenerDeSQL("select cuenta_con from CtasBank where codigo = 1 "))
        Else
            g.tx Row, gC_CON, CuentaParam(ID_Cuenta_M_CH_CARTERA)
        End If
    End If
End Function

Public Property Get chNroInt(Row As Long) As Long
    chNroInt = s2n(g.tx(Row, gN_INT))
End Property
Public Property Get chPropio(Row As Long) As Boolean
    chPropio = (g.tx(Row, gPRoTE) = "P")
End Property
Public Property Get chBancCod(Row As Long) As Long
    chBancCod = s2n(g.tx(Row, gBAN_C))
End Property
Public Property Get chBancDes(Row As Long) As String
    chBancDes = Trim$(g.tx(Row, gBAN_D))
End Property
Public Property Get chNumero(Row As Long) As String 'long
    chNumero = s2n(g.tx(Row, gNUMER))
End Property
Public Property Get chMonto(Row As Long) As Double
    chMonto = s2n(g.tx(Row, gMONTO))
End Property
Public Property Get chFecha(Row As Long) As Date
    chFecha = CDate(g.tx(Row, gFECHA))
End Property
Public Property Get chCuenta(Row As Long) As String
    chCuenta = g.tx(Row, gC_CON)
End Property

Public Property Let enabled(sino As Boolean)
    UserControl.enabled = sino
End Property

Public Function chSetearNroInt(Row As Long, NroInt As Long)
    If s2n(g.tx(Row, gN_INT)) > 0 Then
        ufa "prg: err al grabar cheque propio", "uCheques chSetearNroInt"
    Else
        g.tx Row, gN_INT, NroInt
        mNroIntPropioFueModificado = True
    End If
End Function
Public Sub resetNroIntPropios()
    Dim i As Long
    If VerParametro(BS_EXIGE_CARGA_CHEQUERA) = False And mNroIntPropioFueModificado Then
        For i = 1 To GRILLA.rows - 1
            If chPropio(i) Then g.tx i, gN_INT, 0
        Next i
    End If
End Sub


Public Function FechasOk() As Boolean
    Dim i As Long
    
    For i = 1 To g.rows - 1
        If chNroInt(i) > 0 Then
            If chFecha(i) < #1/1/2001# Then
                che "fecha cheque fila " & i
                Exit Function
            End If
        End If
    Next
    FechasOk = True
End Function


Private Sub cmdBorraItem_Click()
    If g.rows > 2 Then
        g.delRow g.Row
    '    g.rows = TOT_ROWS
        recalculo
    End If
End Sub

Private Sub g_BeforeEdit(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
    If Col = gN_INT Or Col = gPRoTE Then Exit Sub
    'Cancel = (g.tx(Row, gPRoTE) <> "P")
End Sub

Private Sub g_cambio(ByVal Row As Long, ByVal Col As Long, txt As String)
    If Col = gMONTO Then
        recalculo
    ElseIf Col = gPRoTE Then
        g.tx Row, gN_INT, 0
        g.tx Row, gMONTO, 0
        recalculo
    End If
    If g.rows < Row + 2 Then
        g.rows = g.rows + 1
        g.tx g.rows - 1, gPRoTE, "P"
    End If
End Sub

Private Sub g_DblClick()
    Dim resu As String, ss As String
    
    If g.Col = gN_INT Or g.Col = gNUMER Then
        If g.tx(g.Row, gPRoTE) = "P" Then
            ss = "Select chq_comp.Codigo, Nro as [ Nro         ], Importe, fecha_cheque,  BancosGrales.Descripcion as [ Banco                       ] from Chq_comp inner join bancosGrales on BancosGrales.codigo = banco  where estado = 'C' and Chq_comp.activo = 1"
            resu = frmBuscar.MostrarSql(ss)
            If resu > "" Then
                If Not ChqEsta(s2n(resu), "P") Then metoCheque g.Row, s2n(resu), "P"
            End If
           
        ElseIf g.tx(g.Row, gPRoTE) = "T" Then
            ss = "Select NroInt, BancosGrales.Descripcion as [ Banco               ], Nro as [ Nro         ], Importe, Fecha from Cheques inner join BancosGrales on Banco_Nro = BancosGrales.codigo where estado = 'C' and cheques.activo = 1  "
            resu = frmBuscar.MostrarSql(ss)
            If resu > "" Then
                If Not ChqEsta(s2n(resu), "T") Then metoCheque g.Row, s2n(resu), "T"
            End If
        End If
    End If
End Sub

Private Function ChqEsta(COD As Long, letra As String) As Boolean
Dim i As Long
ChqEsta = False

    For i = 1 To GRILLA.rows - 1
        If s2n(GRILLA.TextMatrix(i, gN_INT)) = COD And GRILLA.TextMatrix(i, gPRoTE) = letra Then
            ChqEsta = True
        End If
    Next

End Function

Private Sub g_Validar(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
    Dim tempo
    Dim nin As Long
    Dim banDes As String
    
    Select Case Col
    Case gN_INT
        nin = s2n(g.EditText)
        Select Case g.tx(Row, gPRoTE)
        Case "P"
            tempo = obtenerDeSQL("Select chq_comp.Codigo, Nro, Importe, fecha_cheque,  BancosGrales.Descripcion as Banco from Chq_comp inner join bancosGrales on BancosGrales.codigo = banco  where estado = 'C' and Chq_comp.activo = 1 and chq_comp.codigo = " & nin)
        Case "T"
            tempo = obtenerDeSQL("Select NroInt, BancosGrales.Descripcion as Banco, Nro, Importe, Fecha from Cheques inner join BancosGrales on Banco_Nro = BancosGrales.codigo where estado = 'C' and cheques.activo = 1  and cheques.NroInt = " & nin)
        Case Else
        End Select
        
        If IsEmpty(tempo) Then
            cancel = True
            Exit Sub
        End If
        metoCheque Row, g.EditText, g.tx(Row, gPRoTE)
        
    'Case gPRoTE
    Case gNUMER
    

        Select Case g.tx(Row, gPRoTE)
        Case "P"
            If VerParametro(BS_EXIGE_CARGA_CHEQUERA) Then
                tempo = obtenerDeSQL("Select chq_comp.Codigo, Nro, Importe, fecha_cheque, BancosGrales.Descripcion as Banco from Chq_comp inner join bancosGrales on BancosGrales.codigo = banco  where estado = 'C' and Chq_comp.activo = 1 and chq_comp.nro = " & g.EditText) 'chq_comp.codigo = " & g.EditText)
                If IsEmpty(tempo) Then
                    cancel = True
                    Exit Sub
                End If
                metoCheque Row, s2n(tempo(0)), g.tx(Row, gPRoTE)
            Else
                metoCheque Row, 0, "N"
            End If
        Case "T"
            If VerParametro(BS_EXIGE_CARGA_CHEQUERA) Then
                tempo = obtenerDeSQL("Select NroInt, BancosGrales.Descripcion as Banco, Nro, Importe, Fecha from Cheques inner join BancosGrales on Banco_Nro = BancosGrales.codigo where estado = 'C' and cheques.activo = 1  and cheques.nro = '" & g.EditText & "' ") '.NroInt = " & g.EditText)
                If IsEmpty(tempo) Then
                    cancel = True
                    Exit Sub
                End If
                metoCheque Row, s2n(tempo(0)), g.tx(Row, gPRoTE)
            Else
                metoCheque Row, 0, "N"
            End If
        Case Else
        End Select
        
    Case gBAN_C  ' solo si no exige coontrol cheques
            If g.tx(Row, gPRoTE) = "T" Then
                banDes = sSinNull(obtenerDeSQL("select descripcion from bancosgrales where codigo = " & g.EditText))
                g.tx Row, gBAN_D, banDes
                g.tx Row, gC_CON, CuentaParam(ID_Cuenta_M_CH_CARTERA)
            Else
                ' ESTO ESTA MAL; DEBE IR CUENTA CONTABLE DE CUENTA BANCARIA
                banDes = sSinNull(obtenerDeSQL("select descripcion from bancosgrales where codigo = " & g.EditText))
                g.tx Row, gBAN_D, banDes
                g.tx Row, gC_CON, CuentaParam(ID_Cuenta_M_CH_CARTERA)
            End If
    End Select
End Sub



Private Sub grilla_GotFocus()
    If GRILLA.Row = 0 Then GRILLA.Select 1, 0
    'grilla.set
End Sub

Private Sub UserControl_Initialize()
    On Error Resume Next
    inigrilla
    recalculo
'    UserControl.BackColor = UserControl.ParentControls(0).BackColor
End Sub

'propios
'    "Select chq_comp.Codigo, Nro, Importe, fecha_cheque,  BancosGrales.Descripcion as Banco from Chq_comp inner join bancosGrales on BancosGrales.codigo = banco  where estado = 'C' and Chq_comp.activo = 1"
'3ros
'    "Select NroInt, BancosGrales.Descripcion as Banco, Nro, Importe, Fecha from Cheques inner join BancosGrales on Banco_Nro = BancosGrales.codigo where estado = 'C' and cheques.activo = 1"

Private Sub inigrilla()
    Set g = New LiGrilla
    g.init GRILLA
    
    If VerParametro(BS_EXIGE_CARGA_CHEQUERA) Then
        With g
            gPRoTE = .AddCol(" P/T ", "B", "P|T")
            gNUMER = .AddCol(" Numero     ", "N", 0)
            gMONTO = .AddCol(" Importe      ", "N", 2)
            gBAN_C = .AddCol(" Banco ") ', "N", 0)
            gBAN_D = .AddCol(" Banco                                ")
            gFECHA = .AddCol(" Fecha    ", "D")
            gC_CON = .AddCol(" Cuenta     ")
            gN_INT = .AddCol(" Cod Int ", "N", 0)
        End With
    Else
        With g
            gPRoTE = .AddCol(" P/T ", "B", "P|T")
            gNUMER = .AddCol(" Numero     ", "N", 0)
            gMONTO = .AddCol(" Importe      ", "N", 2)
            gBAN_C = .AddCol(" Banco ", "N", 0)
            gBAN_D = .AddCol(" Banco                                ")
            gFECHA = .AddCol(" Fecha    ", "D")
            gC_CON = .AddCol(" Cuenta     ")
            gN_INT = .AddCol(" Cod Int ", "N", 0)
        End With
    End If
    mRows = 0
End Sub
Private Sub recalculo()
    Dim i As Long, suma As Double
    Dim veri
    veri = VerParametro(BS_EXIGE_CARGA_CHEQUERA)
    
    If veri Then
        For i = 1 To g.rows - 1
            If s2n(g.tx(i, gN_INT)) = 0 Or s2n(g.tx(i, gMONTO)) = 0 Then Exit For
            suma = suma + s2n(g.tx(i, gMONTO))
        Next i
    Else
        For i = 1 To g.rows - 1
            If s2n(g.tx(i, gMONTO)) = 0 Then Exit For
            'If s2n(g.tx(i, gN_INT)) = 0 And g.tx(i, gPRoTE) = "T" Then Exit For
            suma = suma + s2n(g.tx(i, gMONTO))
        Next i

    End If
    mRows = i - 1
    txttotal = s2n(suma)
    RaiseEvent cambio
End Sub

Private Sub UserControl_Show()
    On Error Resume Next
    Dim i As Long
    'UserControl.BackColor = UserControl.ParentControls(0).BackColor
    For i = 0 To Label1.Count - 1
        'Label1(i).BackColor = UserControl.BackColor
    Next i
End Sub

'13/4/5 'bolu: foco select 1 fila
'

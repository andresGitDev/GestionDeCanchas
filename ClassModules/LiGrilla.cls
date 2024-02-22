VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LiGrilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum lgTipoColumna
    lgAutoNum
    lgCombo
    lgHidden
    lgString
    lgNumEdit
    lgDate
    lgCheck
    lg_NoEdit
    lgNumNoEdit
End Enum

Private WithEvents g As VSFlexGrid
Attribute g.VB_VarHelpID = -1

Private nc As Long, ac As Long
Private gt() As String
Private mDecimales As Long
'
Public Event cambio(ByVal Row As Long, ByVal Col As Long, ByRef txt As String)
Public Event DblClick()
Public Event Validar(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
Public Event BeforeEdit(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
Public Event click()
'


' dentro del validateEdit()
Public Function editOk(ByVal Row As Long, ByVal Col As Long) As Boolean
    Dim tx As String, C As Boolean
    
    tx = g.EditText
    C = False
    
    Select Case Left(gt(Col), 1)
    Case "S": C = True
    Case "K": C = True
    Case "B": C = True
    Case "D": C = IsDate(tx)
'    Case "N": tx = Replace(tx, ".", ","): C = IsNumeric(tx): g.EditText = tx
    Case "N":
        If SeparadorDecimal() = "," Then tx = Replace(tx, ".", ","):
        C = IsNumeric(tx): g.EditText = tx
    End Select
    editOk = C
End Function

Public Sub clear()
    g.clear
End Sub

Public Function esEditable(ByVal Col As Long) As Boolean
    esEditable = InStr("SNKBD", Left(gt(Col), 1))
End Function

Public Function addRow() As Long
    Dim i As Long
    g.rows = g.rows + 1
    If ac <> -1 Then
        For i = 1 To g.rows - 1: g.TextMatrix(i, ac) = i: Next i
    End If
    addRow = g.rows - 1
End Function

Public Function delRow(Optional ByVal Row)
    If IsMissing(Row) Then Row = g.rows - 1
    If Row > 0 Then g.RemoveItem Row
End Function

Public Function suma(Col As Long) As Double
    On Error Resume Next
    Dim i As Long
    
    suma = 0
    For i = 1 To g.rows
        suma = s2n(suma + s2n(g.TextMatrix(i, Col), mDecimales), mDecimales)
    Next i
End Function


Public Function PrimerVacio(Optional Col As Long) As Long
    Dim i As Long
    
    If IsMissing(Col) Then Col = IIf(gt(0) = "A", 1, 0)
    
    For i = 1 To g.rows - 1
        If g.TextMatrix(i, Col) = "" Then
            PrimerVacio = i
            Exit Function
        End If
    Next i
    PrimerVacio = addRow
End Function

'                                        Optional tipo_ADNSKBH As lgTipoColumna, optional editable as boolean, optional parametroColumna)
Public Function AddCol(Nombre As String, Optional tipo_ADNSKBH As String, Optional parametro) As Long
'                                        Optional tipo_ADNSKBH As lgTipoColumna
    Dim tI As String, ndec As Long, sCbo As String
    
    ndec = IIf(IsMissing(parametro), 2, s2n(parametro))
    sCbo = IIf(IsMissing(parametro), "", sSinNull(parametro))
    tI = UCase(tipo_ADNSKBH)
    If tI = "" Then tI = "-"
    
    nc = nc + 1
    g.cols = nc + 1
    ReDim Preserve gt(nc)

    g.TextMatrix(0, nc) = Nombre
    gt(nc) = tI
    
    If tI = "A" Then ac = nc
    If tI = "K" Then g.ColDataType(nc) = flexDTBoolean
    If tI = "B" Then g.ColComboList(nc) = sCbo
    If tI = "H" Then g.ColHidden(nc) = True
    If tI = "N" Then g.ColFormat(nc) = IIf(ndec = 0, "#,#", "#,#." & Left("########", ndec))
    If tI = "9" Then g.ColFormat(nc) = IIf(ndec = 0, "#,#", "#,#." & Left("########", ndec)) '"#,#." & Left("########", ndec)
    If tI = "D" Then g.ColFormat(nc) = "dd/mm/yy"
    If tI = "S" Then g.ColAlignment(nc) = flexAlignLeftCenter
    If tI = "-" Then g.ColAlignment(nc) = flexAlignLeftCenter
   
'    If ti = lgAutoNum Then ac = nc
'    If ti = lgCheck Then g.ColDataType(nc) = flexDTBoolean
'    If ti = lgCombo Then g.ColComboList(nc) = comboString
'    If ti = lgHidden Then g.ColHidden(nc) = True
'    If ti = lgNumEdit Then g.ColFormat(nc) = ",#.##"
'    If ti = lgNumNoEdit Then g.ColFormat(nc) = ",#.##"

'    'lgString
'    'lgDate
'    'lg_NoEdit
    
    g.AutoSize nc
    AddCol = nc
End Function


Public Sub init(GRILLA As VSFlexGrid, Optional nDecimales = 2)
    Set g = GRILLA
    nc = -1: ac = -1
    mDecimales = nDecimales
    With g
        .clear
        .Editable = True
        .FixedCols = 0
        .FixedRows = 1
        .TabBehavior = flexTabCells
        .cols = 0
        .rows = 2
        .AllowUserResizing = flexResizeColumns
    End With
End Sub

Public Function Borrar()
    g.rows = 1
End Function


Public Function contar(Col As Long, que As String) As Long
    On Error Resume Next
    Dim i As Long
    
    contar = 0
    For i = 1 To g.rows - 1
        If Trim(que) = Trim(g.TextMatrix(i, Col)) Then contar = contar + 1
    Next i
End Function


Public Function buscar(Col As Long, que As Variant, Optional desde As Long) As Long
    On Error Resume Next
    Dim i As Long
    
    If desde = 0 Then desde = 1
    If desde > g.rows Then desde = g.rows - 1
    
    'If IsNumeric(que) Then
    If EsColumnaNumerica(Col) Then
        buscar = 0
        For i = desde To g.rows - 1
            If s2n(que, mDecimales) = s2n(g.TextMatrix(i, Col), mDecimales) Then
                buscar = i
                Exit Function
            End If
        Next i
    Else
        buscar = 0
        For i = desde To g.rows - 1
            If Trim(que) = Trim(g.TextMatrix(i, Col)) Then
                buscar = i
                Exit Function
            End If
        Next i
    End If
End Function


Public Function Name(Col) As String
    Name = g.TextMatrix(0, Col)
End Function


Public Function tx(Row As Long, Col As Long, Optional que) As String
    Dim sTx As String
    tx = Trim$(g.TextMatrix(Row, Col))
    If Not IsMissing(que) Then
        'sTx = IIf(IsNull(que), "-null-", Trim(CStr(que))) ' ODIO VB
        If IsNull(que) Then
            sTx = "-null-"
        Else
            sTx = Trim$(CStr(que))
        End If
        If tx <> sTx Then
            If sTx = "Verdadero" Then sTx = "1"
            If sTx = "Falso" Then sTx = "0"
            g.TextMatrix(Row, Col) = sTx
'            RaiseEvent cambio(Row, Col, sTx)
        End If
    End If
End Function

Public Function tk(Row As Long, Col As Long, Optional que) As Boolean
    If Not IsMissing(que) Then
        g.cell(flexcpChecked, Row, Col) = IIf(que, flexChecked, flexUnchecked)
    End If
    tk = (g.cell(flexcpChecked, Row, Col) = flexChecked)
End Function

Public Function td(Row As Long, Col As Long, Optional que) As Date
    On Error GoTo UfaTd
    If Not IsMissing(que) Then
        g.TextMatrix(Row, Col) = que
    End If
    td = CDate(g.TextMatrix(Row, Col))
fin:
    Exit Function
UfaTd:
    td = #1/1/2000#
    Resume fin
End Function


'****************************************************************************************

Private Sub Class_Terminate()
    Set g = Nothing
End Sub

'Private Sub g_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'    If Row > 0 Then
'    'RaiseEvent cambio(Row, Col, g.EditText)
'    End If
'End Sub

Private Sub g_BeforeEdit(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
    cancel = Not esEditable(Col)
    If cancel Then Exit Sub
    
    RaiseEvent BeforeEdit(Row, Col, cancel)
End Sub

Private Sub g_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Row > 0 Then
        RaiseEvent cambio(Row, Col, g.TextMatrix(Row, Col))   'g.EditText)
    End If
End Sub

Private Sub g_Click()
    RaiseEvent click
End Sub

Private Sub g_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub g_ValidateEdit(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
    If cancel Then Exit Sub
    cancel = Not editOk(Row, Col)
    If cancel Then Exit Sub
    RaiseEvent Validar(Row, Col, cancel)
End Sub

Private Function EsColumnaNumerica(colu)
    EsColumnaNumerica = (InStr("9N", gt(colu)) > 0)
End Function

' //*************  Property Page  ***********************//
Public Property Get Row() As Long
    Row = g.Row
End Property
Public Property Let Row(que As Long)
    g.Row = que
End Property

Public Property Get Col() As Long
    Col = g.Col
End Property
Public Property Let Col(que As Long)
    g.Col = que
End Property

'Public Property Get tx(ByVal row As long, ByVal col As long) As String
'    tx = g.TextMatrix(row, col)
'End Property

Public Property Get TextMatrix(ByVal Row As Long, ByVal Col As Long) As String
    TextMatrix = g.TextMatrix(Row, Col)
End Property

Public Property Get Text() As String
    Text = g.Text
End Property
Public Property Let Text(que As String)
    g.Text = que
'    RaiseEvent cambio(g.Row, g.Col, que)
End Property

Public Property Get cols()
    cols = g.cols
End Property

Public Property Let cols(que)
    g.cols = que
    g.Refresh
End Property

Public Property Get rows()
    rows = g.rows
End Property

Public Property Let rows(que)
    g.rows = que
End Property

Public Property Get EditText()
    EditText = g.EditText
End Property

Public Property Let EditText(que)
    g.EditText = que
End Property

'28/7/4     buscar optional desde
'9/8/4      prop tx a func(), parametro para grabar
'           Se paso todo de Integer a Long, lo q usa la grilla
'           propertychanged: el help la recomienda, pero falla.
'10/8/4     fix buscar numerico
'3/9/4      solo comentarios
'9/9/4      fix EsColumnaNumerica(), hace fallar .Buscar()
'16/2/5     agregados param cant decimales opcional
'2/3/5      tx raiseevent solo si dato metido es distinto al anterior
'13/6/5     new: tk(), td()
'

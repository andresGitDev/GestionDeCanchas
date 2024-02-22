VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBuscar 
   Caption         =   "Buscar"
   ClientHeight    =   7125
   ClientLeft      =   90
   ClientTop       =   345
   ClientWidth     =   6780
   Icon            =   "frmBuscar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "frmBuscar.frx":1CFA
   ScaleHeight     =   7125
   ScaleWidth      =   6780
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      Begin VB.Frame fraTool 
         BorderStyle     =   0  'None
         Caption         =   "-"
         Height          =   1035
         Left            =   60
         TabIndex        =   2
         Top             =   6060
         Width           =   7095
         Begin VB.TextBox txtBuscar 
            Height          =   315
            Left            =   735
            TabIndex        =   6
            Top             =   600
            Width           =   3675
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   90
            Picture         =   "frmBuscar.frx":49F4
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   510
            Width           =   495
         End
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "&Aceptar"
            Default         =   -1  'True
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   4560
            Picture         =   "frmBuscar.frx":66EE
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   195
            Width           =   930
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "&Cancelar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   5475
            MaskColor       =   &H8000000F&
            Picture         =   "frmBuscar.frx":6C78
            Style           =   1  'Graphical
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   195
            Width           =   900
         End
         Begin VB.Label Label2 
            Caption         =   "Pruebe el boton secundario sobre campo"
            Height          =   195
            Left            =   15
            TabIndex        =   9
            Top             =   -15
            Width           =   5055
         End
         Begin VB.Label Label1 
            Caption         =   "Orden:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   15
            TabIndex        =   8
            Top             =   255
            Width           =   615
         End
         Begin VB.Label lblColumna 
            Caption         =   "- Predeterminado -"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   735
            TabIndex        =   7
            Top             =   255
            Width           =   2955
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgBu 
         CausesValidation=   0   'False
         Height          =   5895
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   6660
         _ExtentX        =   11748
         _ExtentY        =   10398
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   -2147483644
         ForeColorFixed  =   -2147483625
         TextStyleFixed  =   1
         FocusRect       =   2
         AllowUserResizing=   3
         GridLineWidthFixed=   2
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontWidthFixed  =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   465
      Picture         =   "frmBuscar.frx":7542
      Top             =   7335
      Width           =   960
   End
   Begin VB.Menu mnuColumna 
      Caption         =   "Columna"
      Visible         =   0   'False
      Begin VB.Menu mnuOrdenA 
         Caption         =   "A-Z Orden ascendente"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuOrdenZ 
         Caption         =   "Z-A Orden descendente"
      End
      Begin VB.Menu z1 
         Caption         =   ""
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBuscaTexto 
         Caption         =   "Buscar texto:"
         Shortcut        =   {F3}
         Visible         =   0   'False
      End
      Begin VB.Menu z11 
         Caption         =   ""
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFiltroIncluye 
         Caption         =   "Filtrar: Mostrar seleccion"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuFiltroExcluye 
         Caption         =   "Filtrar: Excluir seleccion"
      End
      Begin VB.Menu z2 
         Caption         =   ""
      End
      Begin VB.Menu mnuFiltroMayor 
         Caption         =   "Ver Mayores a"
      End
      Begin VB.Menu mnuFiltroMenor 
         Caption         =   "Ver Menores a"
      End
   End
End
Attribute VB_Name = "frmBuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum TipoCampo
    tipoNumero
    tipoString
    tipoFecha
    tipoBool
End Enum

Private Enum TipoSort
    flexSortNumericAscending = 3
    flexSortNumericDescending = 4
    flexSortStringNoCaseAsending = 5
    flexSortStringNoCaseDescending = 6
    flexSortCustom = 9
End Enum

Private Enum TipoSortAZ
    TIPO_ORDEN_A ' custom ascendente
    TIPO_ORDEN_Z ' custom descendente
End Enum

Private mTipoOrdenAZ As TipoSortAZ ' custom
Private mOrden As Long 'columna ordenada

Private mRow As Long
Private mCol As Long
Private mBackColor

'array columnas
Private pCols As Long 'cant columnas
Private a() As String
Private tcCol() As TipoCampo
'

Private mBuscarF3 As String

Public Function resultado(Optional que)
On Error GoTo fin2
Static re() As String
    Dim i As Long
    
    resultado = ""
    
    If IsMissing(que) Then que = 1
    If que = -1 Then
        ReDim re(pCols)
        For i = 1 To pCols
            re(i) = a(i)
        Next i
    Else
        resultado = re(que)
    End If
fin2:
End Function

Public Function MostrarSql(strSql, Optional arrayAnchos As Variant, Optional caption As String, Optional PalabraPorNull = "Null", Optional PalabraPorTrue = "Verdadero", Optional PalabraPorFalse = "Falso", Optional conDataSource As Boolean = False, Optional queConex As ADODB.Connection)
    'Devuelve string = al 1er campo
    'Para cambiar titulos, poner alias a los campos (select nombrecampo AS [Nombre del Campo],... )
    'Para agregar tamaño a campos:  frmBuscar.mostrarSql("select...", array(2,8,30,...))

    'On Error Resume Next 'GoTo FIN
    ' mod para campos null

    Dim rs As New ADODB.Recordset, f As ADODB.Field, i As Long, ancho
    Dim sf As String, sT As String ', valor
    
    Screen.ActiveForm.MousePointer = vbHourglass
    
    mOrden = 0
    mRow = 0
    mCol = 0
    lblColumna = ""
    'mBackColor
    If queConex Is Nothing Then
        Set queConex = DataEnvironment1.Sistema
    End If
    
    rs.Open strSql, queConex, adOpenStatic, adLockReadOnly
    pCols = rs.Fields.Count
    fgBu.cols = pCols
    ReDim a(pCols), tcCol(pCols)
    resultado -1
    
    If caption > "" Then Me.caption = Me.caption & " " & caption
    
    With fgBu
        .Redraw = False
        .FixedRows = 1
        .TextStyleFixed = flexTextRaised  ' flexTextRaisedLight
'       .FontFixed = "Arial"
        
        'MOD 21/3/6
        If conDataSource Then Set .DataSource = rs
        
        
        ' 1er loop columnas, determino columnas, nombre, tipo
        i = 0
        For Each f In rs.Fields
            If i > 0 Then sf = sf & "|"
            .TextMatrix(0, i) = f.Name ' titulos
            
            Select Case f.Type
             Case adChar, adChar, adVarChar, adVarWChar, adWChar, adLongVarChar, adLongVarWChar
                tcCol(i) = tipoString
             Case f.Type = adBigInt, adCurrency, adDecimal, adDouble, adInteger, adNumeric, adSingle, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt, adVarNumeric
                tcCol(i) = tipoNumero
             Case adDate, adDBDate, adDBTimeStamp, adDBTime
                tcCol(i) = tipoFecha
             Case adBoolean
                tcCol(i) = tipoBool
             'Case , , adFileTime + adIUnknown + adPropVariant + adUserDefined + adVariant
             Case Else
                If f.Type < 128 Then
                    tcCol(i) = tipoNumero
                Else
                    tcCol(i) = tipoString
                End If

            End Select

            sf = sf & f.Name
            i = i + 1
        Next f
        
        'acomodo automatico de columnas
        .FormatString = sf

        '2do loop columnas , despues del .FormatString, reacomodo anchos y alineamiento
        'alineamiento y anchos
        For i = 0 To UBound(tcCol) - 1 'rs.Fields.Count - 1
            Select Case tcCol(i)
            Case tipoFecha
                If IsArray(arrayAnchos) = False Then
                    cancho i, 1500
                Else
                    cancho i, arrayAnchos(i)
                End If
            Case tipoNumero
                If IsArray(arrayAnchos) = False Then
                    cancho i, 1000
                Else
                    cancho i, arrayAnchos(i)
                End If
            Case tipoString
                If IsArray(arrayAnchos) = False Then
                    'cancho i, 2000
                Else
                    cancho i, arrayAnchos(i)
                End If
'                cancho i, 5000
                .ColAlignment(i) = 1
            End Select
        Next i

        .CellFontBold = False
        
        'MOD 21/3/6
        If Not conDataSource Then
        
            'loop filas = registros
            While Not rs.EOF
                i = 0
                For Each f In rs.Fields
                    If IsNull(f.Value) Then
                        sT = PalabraPorNull
                    ElseIf LCase(Trim(f.Value)) = "true" Then
                        sT = PalabraPorTrue
                    ElseIf LCase(Trim(f.Value)) = "false" Then
                        sT = PalabraPorFalse
                    ElseIf LCase(Trim(f.Value)) = "verdadero" Then
                        sT = PalabraPorTrue
                    ElseIf LCase(Trim(f.Value)) = "falso" Then
                        sT = PalabraPorFalse
                    Else
                        sT = f.Value
                    End If
                    .TextMatrix(.rows - 1, i) = sT
                    i = i + 1
                Next f
                rs.MoveNext
                .rows = .rows + 1
            Wend
                    
            If .rows > 2 Then .rows = .rows - 1
            
        End If
        .Redraw = True
    End With
    Screen.ActiveForm.MousePointer = vbDefault
     Me.Show vbModal
fin:
    On Error Resume Next
    Screen.ActiveForm.MousePointer = vbDefault
    Set rs = Nothing
    MostrarSql = resultado()
End Function

Public Function MostrarRs(rs As ADODB.Recordset, Optional arrayAnchos As Variant, Optional caption As String, Optional PalabraPorNull = "Null", Optional PalabraPorTrue = "Verdadero", Optional PalabraPorFalse = "Falso") As String
    'Devuelve string = al 1er campo
    'Para cambiar titulos, poner alias a los campos (select nombrecampo AS [Nombre del Campo],... )
    'Para agregar tamaño a campos:  frmBuscar.mostrarSql("select...", array(2,8,30,...))

    'On Error Resume Next 'GoTo FIN
    ' mod para campos null

    'Dim rs As New ADODB.Recordset,
    Dim f As ADODB.Field, i As Long, ancho
    Dim sf As String, sT As String ', valor
    
    Screen.ActiveForm.MousePointer = vbHourglass
    
    mOrden = 0
    mRow = 0
    mCol = 0
    lblColumna = ""
    'mBackColor
    
   
    'rs.Open strSql, daTaenvironment1.Sistema, adOpenStatic, adLockReadOnly
    
    pCols = rs.Fields.Count
    fgBu.cols = pCols
    ReDim a(pCols), tcCol(pCols)
    resultado -1
    
    If caption > "" Then Me.caption = caption
    
    With fgBu
        .Redraw = False
        .FixedRows = 1
        .TextStyleFixed = flexTextRaised  ' flexTextRaisedLight
'       .FontFixed = "Arial"
        
        ' 1er loop columnas, determino columnas, nombre, tipo
        i = 0
        For Each f In rs.Fields
            If i > 0 Then sf = sf & "|"
            .TextMatrix(0, i) = f.Name ' titulos
            
            Select Case f.Type
             Case adChar, adChar, adVarChar, adVarWChar, adWChar, adLongVarChar, adLongVarWChar
                tcCol(i) = tipoString
             Case f.Type = adBigInt, adCurrency, adDecimal, adDouble, adInteger, adNumeric, adSingle, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt, adVarNumeric
                tcCol(i) = tipoNumero
             Case adDate, adDBDate, adDBTimeStamp, adDBTime
                tcCol(i) = tipoFecha
             Case adBoolean
                tcCol(i) = tipoBool
             'Case , , adFileTime + adIUnknown + adPropVariant + adUserDefined + adVariant
             Case Else
                If f.Type < 128 Then
                    tcCol(i) = tipoNumero
                Else
                    tcCol(i) = tipoString
                End If

            End Select

            sf = sf & f.Name
            i = i + 1
        Next f
        
        'acomodo automatico de columnas
        .FormatString = sf

        '2do loop columnas , despues del .FormatString, reacomodo anchos y alineamiento
        'alineamiento y anchos
        For i = 0 To UBound(tcCol) - 1 'rs.Fields.Count - 1
            Select Case tcCol(i)
            Case tipoFecha
                cancho i, 1500
            Case tipoNumero
                cancho i, 1000
            Case tipoString
'                cancho i, 5000
                .ColAlignment(i) = 1
            End Select
        Next i

        .CellFontBold = False
        
        
        'loop filas = registros
        While Not rs.EOF
            i = 0
            For Each f In rs.Fields
                '.TextMatrix(.rows - 1, i) = IIf(IsNull(f.Value), "Null", f.Value)
                If IsNull(f.Value) Then
                    sT = PalabraPorNull
                ElseIf LCase(Trim(f.Value)) = "true" Then
                    sT = PalabraPorTrue
                ElseIf LCase(Trim(f.Value)) = "false" Then
                    sT = PalabraPorFalse
                Else
                    sT = f.Value
                End If
                .TextMatrix(.rows - 1, i) = sT
                i = i + 1
            Next f
            rs.MoveNext
            .rows = .rows + 1
        Wend


        .rows = .rows - 1
        .Redraw = True
    End With

    Screen.ActiveForm.MousePointer = vbDefault
    Me.Show vbModal

fin:
    On Error Resume Next
    Screen.ActiveForm.MousePointer = vbDefault
    Set rs = Nothing
    MostrarRs = resultado()
End Function

Public Function MostrarArray(aDatos2c As Variant, Optional caption As String, Optional PalabraPorNull = "Null", Optional PalabraPorTrue = "Verdadero", Optional PalabraPorFalse = "Falso") As String
    ' *** ANDA!!!! ***
    ' pasar array   a(cols, rows)
    ' 3 columnas 200 filas seria:    redim a(2,199)
    ' primer fila asume titulos
    
    Dim f As ADODB.Field, i As Long, ancho
    Dim sf As String, sT As String ', valor
    
    Screen.ActiveForm.MousePointer = vbHourglass
    
    mOrden = 0
    mRow = 0
    mCol = 0
    lblColumna = ""
    'mBackColor
    
    pCols = UBound(aDatos2c) + 1
    fgBu.cols = pCols
    ReDim a(pCols), tcCol(pCols)
    resultado -1
    
    If caption > "" Then Me.caption = caption
    
    Dim tt As Variant
    With fgBu
        .Redraw = False
        .FixedRows = 1
        .TextStyleFixed = flexTextRaised
'       .FontFixed = "Arial"
        
        ' 1er loop columnas, determino columnas, nombre, tipo
        For i = 0 To pCols - 1
            If i > 0 Then sf = sf & "|"
            .TextMatrix(0, i) = aDatos2c(i, 0) ''tt(i)
            'f.Name ' titulos

            sf = sf & aDatos2c(i, 0) 'f.Name
            'i = i + 1
        Next i
        
        'acomodo automatico de columnas
        .FormatString = sf
        
        .CellFontBold = False
        
        
        Dim j
        For i = 1 To UBound(aDatos2c, 2)
            For j = 0 To pCols - 1
                .TextMatrix(.rows - 1, j) = aDatos2c(j, i) 'tt(j)  'sT
            Next j
            .rows = .rows + 1
        Next i
        .rows = .rows - 1
        .Redraw = True
    End With
    
    Screen.ActiveForm.MousePointer = vbDefault
    Me.Show vbModal

fin:
    On Error Resume Next
    Screen.ActiveForm.MousePointer = vbDefault
    MostrarArray = resultado()
End Function


Public Function MostrarCodigoDescripcionActivo(queTabla As String) ', Optional arrayAnchos As Variant)
    MostrarCodigoDescripcionActivo = MostrarSql("select Codigo as [ Codigo              ], Descripcion [ Descripcion                                                             ] from " & queTabla & " where activo = 1 order by codigo ") ', arrayAnchos)
End Function
Private Sub cancho(i, cual)
    fgBu.ColWidth(i) = IIf(cual = 0, cual, maximo(fgBu.ColWidth(i), cual))
End Sub

Private Sub fgBu_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
    'solo fechas
    On Error GoTo fin


    Dim d1 As Date, d2 As Date
    With fgBu
        If mTipoOrdenAZ = TIPO_ORDEN_A Then
            d1 = CDate(.TextMatrix(Row1, .Col))
            d2 = CDate(.TextMatrix(Row2, .Col))
            Cmp = IIf(d1 > d2, 1, -1)
        Else
            d1 = CDate(.TextMatrix(Row1, .Col))
            d2 = CDate(.TextMatrix(Row2, .Col))
            Cmp = IIf(d1 < d2, 1, -1)
        End If
    End With
fin:
End Sub

'
'End Sub

Private Sub fgBu_EnterCell()
    fgBu.CellBackColor = vbCyan
End Sub

'***************************************************************

Private Sub fgBu_LeaveCell()
    txtBuscar = ""
    fgBu.CellBackColor = mBackColor
End Sub

Private Sub fgBu_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    fgBu.CellBackColor = mBackColor
    If Button = 2 And fgBu.MouseRow > 0 Then
        fgBu.Row = fgBu.MouseRow
        fgBu.Col = fgBu.MouseCol
    End If
End Sub

Private Sub fgBu_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 And fgBu.MouseRow > 0 Then
        PopupMenu mnuColumna
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Me.ActiveControl.Name <> "txtBuscar" Then
        fgBu.CellBackColor = mBackColor
        txtBuscar.SetFocus
        txtBuscar.Text = Chr(KeyAscii)
        txtBuscar.SelStart = 1
    End If
End Sub

Private Sub mnuFiltroIncluye_Click()
    Dim txn As Double, txs As String, co As Long, i As Long
    Arena True
    With fgBu
        txs = Trim(.Text)
        txn = s2n(txs)
        co = .Col
        '.Redraw = False
        i = 1
        While i < .rows
            If tcCol(co) = tipoNumero Then
                If s2n(.TextMatrix(i, co)) <> txn Then
                    .RemoveItem (i)
                Else
                    i = i + 1
                End If
            Else
                If Trim(.TextMatrix(i, co)) <> txs Then
                    .RemoveItem (i)
                Else
                    i = i + 1
                End If
            End If
        Wend
        '.Redraw = True
    End With
   
    Arena False
End Sub

Private Sub mnuFiltroExcluye_Click()
    On Error GoTo fin
    Dim txn As Double, txs As String, co As Long, i As Long
    
    Arena True
    With fgBu
        txs = Trim(.Text)
        txn = s2n(txs)
        co = .Col

        i = 1
        While i < .rows
            If tcCol(co) = tipoNumero Then
                If s2n(.TextMatrix(i, co)) = txn Then
                    .RemoveItem (i)
                Else
                    i = i + 1
                End If
            Else
                If Trim(.TextMatrix(i, co)) = txs Then
                    .RemoveItem (i)
                Else
                    i = i + 1
                End If
            End If
        Wend
    End With
fin:
    Arena False
End Sub

Private Sub mnuFiltroMayor_Click()
    On Error Resume Next
    Dim txn As Double, txs As String, txd As Date, co As Long, i As Long
    
    Arena True
    With fgBu
        txs = Trim(.Text)
        txn = s2n(txs)
        txd = CDate(txs)
        co = .Col

        i = 1
        While i < .rows And .rows > 2
            If tcCol(co) = tipoNumero Then
                If s2n(.TextMatrix(i, co)) < txn Then
                    .RemoveItem (i)
                Else
                    i = i + 1
                End If
            ElseIf tcCol(co) = tipoFecha Then
                If CDate(.TextMatrix(i, co)) < txd Then
                    .RemoveItem (i)
                Else
                    i = i + 1
                End If
            Else
                If Trim(.TextMatrix(i, co)) < txs Then
                    .RemoveItem (i)
                Else
                    i = i + 1
                End If
            End If
        Wend
    End With
    Arena False
End Sub

Private Sub mnuFiltroMenor_Click()
    On Error Resume Next
    Dim txn As Double, txs As String, txd As Date, co As Long, i As Long
    Arena True
    
    With fgBu
        txs = Trim(.Text)
        txn = s2n(txs)
        txd = CDate(txs)
        co = .Col

        i = 1
        While i < .rows
            If tcCol(co) = tipoNumero Then
                If s2n(.TextMatrix(i, co)) > txn Then
                    .RemoveItem (i)
                Else
                    i = i + 1
                End If
            ElseIf tcCol(co) = tipoFecha Then
                If CDate(.TextMatrix(i, co)) > txd Then
                    .RemoveItem (i)
                Else
                    i = i + 1
                End If
            Else
                If Trim(.TextMatrix(i, co)) > txs Then
                    .RemoveItem (i)
                Else
                    i = i + 1
                End If
            End If
        Wend
    End With
    
    Arena False
    
End Sub

Private Sub mnuOrdenA_Click()
    Dim so As TipoSort
    
    Arena True
    
    Select Case tcCol(fgBu.Col)
    Case tipoBool
        so = flexSortStringNoCaseAsending
    Case tipoFecha
        mTipoOrdenAZ = TIPO_ORDEN_A
        so = flexSortCustom
    Case tipoNumero
        so = flexSortNumericAscending
    Case tipoString
        so = flexSortStringNoCaseAsending
    End Select
    
    fgBu.Sort = so
    mOrden = fgBu.Col
    lblColumna = "Asc - " & fgBu.TextMatrix(0, mOrden)
    
    Arena False
End Sub

Private Sub mnuOrdenZ_Click()
    Dim so As TipoSort
    
    Arena True
    
    Select Case tcCol(fgBu.Col)
    Case tipoBool
        so = flexSortStringNoCaseDescending
    Case tipoFecha
        mTipoOrdenAZ = TIPO_ORDEN_Z
        so = flexSortCustom
    Case tipoNumero
        so = flexSortNumericDescending
    Case tipoString
        so = flexSortStringNoCaseDescending
    End Select
    
    fgBu.Sort = so
    mOrden = fgBu.Col
    lblColumna = "Dsc - " & fgBu.TextMatrix(0, mOrden)
   
    Arena False
End Sub

Private Sub txtbuscar_Change()
    On Error GoTo fin
    Dim r As Long, i As Long
    Dim bRow As Long, bCol As Long
    Dim lena As Long
    Dim lenbus As Long
    Dim lcabus As String
    
    If fgBu.Col <> mOrden Then
        mOrden = fgBu.Col
        fgBu.Sort = flexSortStringNoCaseAsending
        lblColumna = fgBu.TextMatrix(0, mOrden)
    End If
    
    lenbus = Len(txtBuscar)
    lcabus = LCase(txtBuscar)
    With fgBu
        'If mCol > 0 And mRow > 0 Then
        
        bRow = .Row
        bCol = .Col
        If mRow > 0 Then
            .Row = mRow
            .Col = mCol
            .CellBackColor = mBackColor
            .Row = bRow
            .Col = bCol
        End If
        
        r = 1
        If lenbus > 0 Then
            While LCase(Mid(.TextMatrix(r, bCol), 1, lenbus)) <> lcabus And r + 1 < .rows
                r = r + 1
            Wend
            If LCase(Mid(.TextMatrix(r, bCol), 1, lenbus)) = lcabus Then
                .Row = r '- 1
                
                .CellBackColor = vbMagenta 'vbRed
                mCol = .Col
                mRow = .Row
                
                .TopRow = maximo(1, r - 1)
            Else
                .Row = bRow
                .Col = bCol
                .CellBackColor = vbCyan 'vbMagenta
            End If
        End If
    End With
fin:
End Sub


Private Sub cmdCancelar_Click()
    pCols = 0
    ReDim a(pCols)
    resultado -1
    Unload frmBuscar
End Sub


Private Sub cmdAceptar_Click()
    If fgBu.Row = 0 Then
        cmdCancelar_Click
        Exit Sub
    End If
    
    Dim i As Long
    For i = 1 To pCols
        a(i) = fgBu.TextMatrix(fgBu.Row, i - 1)
    Next i
    resultado -1
    Unload Me
End Sub

Private Sub fgBu_DblClick()
    If fgBu.MouseRow > 0 Then cmdAceptar_Click
End Sub

Private Sub Form_Load()
    With fgBu
        .FixedCols = 0
        .FixedRows = 1
        .AllowUserResizing = flexResizeColumns
        mBackColor = .CellBackColor
    End With
    mOrden = -1
End Sub

Private Sub Form_Resize()
    encajar fraMain, Me, 0, 0, 0, 120
    encajar fraTool, fraMain, , 120, 60
    encajar fgBu, fraMain, 60, 60, fraTool.Height + 180, 120
End Sub
Private Function maximo(a, b)
    maximo = IIf(a > b, a, b)
End Function
Private Function minimo(a, b)
    minimo = IIf(a < b, a, b)
End Function
Private Function encajar(que As Object, donde As Object, Optional mT, Optional mL, Optional mB, Optional mR)
    ' para anclar un control dentro de otro,
    ' se pone en resize() del control padre o del form
    On Error Resume Next
    Dim oH As Object, oP As Object
    Dim oPh As Long
    
    Set oH = que: Set oP = donde
    oPh = oP.Height
    oPh = oP.ScaleHeight - 100 'err
    
    If IsMissing(mL) And IsMissing(mR) Then
        oH.Left = (oP.Width - oH.Width) / 2
    ElseIf IsMissing(mL) Then
        oH.Left = (oP.Width - oH.Width) - mR
    ElseIf IsMissing(mR) Then
        oH.Left = mL
    Else
        oH.L = mL
        oH.Width = oP.Width - mR - mL
    End If
    
    If IsMissing(mT) And IsMissing(mB) Then
        oH.Top = (oPh - oH.Height) / 2
    ElseIf IsMissing(mT) Then
        oH.Top = (oPh - oH.Height) - mB
    ElseIf IsMissing(mB) Then
        oH.Top = mT
    Else
        oH.Top = mT
        oH.Height = oPh - mT - mB
    End If
End Function
Private Sub Arena(sino As Boolean)
    On Error Resume Next
    Me.MousePointer = IIf(sino, vbHourglass, vbDefault)
    fgBu.Redraw = Not sino
End Sub


Private Sub txtBuscar_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo fin
    With fgBu
        If KeyCode = 38 Then 'key_UP t
            If .Row > 1 Then
                .CellBackColor = mBackColor
                .Row = .Row - 1
                .CellBackColor = vbCyan
            End If
            .SetFocus
            .TopRow = .Row
        ElseIf KeyCode = 40 Then 'key_DN
            If .Row < (.rows - 1) Then
                .CellBackColor = mBackColor
                .Row = .Row + 1
                .CellBackColor = vbCyan
            End If
            .SetFocus
        End If
    End With
fin:
End Sub



'28/6/04    si hay arrayAnchos, agrega espacios al titulo para ensanchar la columna
'           si no, evalua ancho campo
'
' 22/7/4    ancho columna max(titulo, largocampo)
' 3/8/4     .MostrarSql() devuelve resultado() (el primer campo) !UF!
'           Se va con ESC
'           Las teclas ponen foco en txtbox
'           ancho columnas sigue sin solucion buena, volvere pa tras, sera lo mejor
' 20/8/4    ancho
' 23/8/4    redraw
' 25/8/4    ancho, again: ya no formatstring, ancho = maximo entre ancho campo y ancho titulo
'           pero sigue pendiente
' 26/8/4    ancho, ancho, ancho  widely spread problem
'           No tuvo exito. Ahora :
'           El ancho es el ancho del CAMPO o ALIAS.    Punto.
' 2/9/4     widht Minimo(), x overflow, aunque tal como esta el codigo no lo usa
'15/9/4     Si cancela, ahora resultado devuelve "", no empty (!!)
'           encajar = private, para no depender de otro modulo
'           Caption = parametro opcional titulo form
'30/9/4     MostrarCodigoDescripcionActivo () tamaño de campos predeterminado mas largo
'21/10/4    declaracion explicita MostrarSql field -> adodb.field
'22/10/4    MostrarCodigoDescripcionActivo = order by codigo
'27/10/4    Opcion de ordenar x columna
'           Fix colores busqueda
'           Busqueda SuperVeloz
'28/10/4    Mas criterios de orden
'29/10/4    Tamaños determinados x tipo columna
'22/11/4    fix colores, fix perdida 1er caracter al tipear
'30/11/4    fix doble caracter txtbox
'17/12/4    fix falla busq cuando venia de otra busqueda con mas registros (!!!)
'18/12/4    fix dblClic en row 0
'           fix alineamiento
'           fix ahora ordena fechas
'           popup con clic derecho para ordenar y filtrar
'           relojito arena mientras carga? masomeno
'21/12/4    fix borrado unica linea
'           excluye, incluye taban al reves
'2/5/5      Filtrar > y < ahora solo borran los > y < , no lo >= y <=
'           teclas arriba y abajo
'28/7/5     txtbuscar: Mucho mas rapido, cambio colores:
'               si existe, magenta
'               si NO existe, celeste el mas cercano, ya no salta al final
'           relojito en operaciones lentas de filtros
'21/3/6     conDataSource, si necesito rapidez y no formato (ancho anda, pero no reemplazo palabras NULL, TRUE,  FALSE  ni formato de decimales y fechas
'18/8/6     MostrarArray() ahora OK! probado en sueldos
'

VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmImpresionEtiquetas 
   Caption         =   "Impresion de Codigos de Barra"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9225
   Icon            =   "frmImpresionEtiquetas.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   9225
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Height          =   495
      Left            =   7800
      MaskColor       =   &H8000000F&
      Picture         =   "frmImpresionEtiquetas.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   840
      Width           =   495
   End
   Begin Gestion.ucCoDe uProd 
      Height          =   285
      Left            =   1200
      TabIndex        =   9
      Top             =   360
      Width           =   6495
      _extentx        =   11456
      _extenty        =   503
      codigowidth     =   1000
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5880
      TabIndex        =   7
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Left            =   7560
      TabIndex        =   6
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdAgregarItem 
      Height          =   495
      Left            =   8280
      MaskColor       =   &H8000000F&
      Picture         =   "frmImpresionEtiquetas.frx":099C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   840
      Width           =   4695
   End
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Height          =   3000
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   8715
      _cx             =   15372
      _cy             =   5292
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
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmImpresionEtiquetas.frx":0CA6
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
   Begin VB.Label Label2 
      Caption         =   "Serie :"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Producto : "
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "frmImpresionEtiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private WithEvents g As LiGrilla
Attribute g.VB_VarHelpID = -1

Private gcod As Long
Private gcodBarra As Long
Private gdes As Long
Private gserie As Long
Private gmarca As Long

Dim etiq As New ADODB.Recordset


Private Sub cmdAgregarItem_Click()
    Dim rs As New ADODB.Recordset
    Dim vari As String
    Dim vari2 As String
    
    If uProd.codigo = "" Then Exit Sub
    'If Text1.Text = "" Then Exit Sub
    
    rs.Open "select * from producto where codigo='" & uProd.codigo & "'", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    
    If rs!CODIGOBARRA = "" Then
        vari = uProd.codigo
        vari2 = ""
        'MsgBox "Se agregara el codigo del producto devido a que no posee codigo de barra."
    Else
        vari = uProd.codigo
        vari2 = sSinNull(rs!CODIGOBARRA)
    End If
    Set rs = Nothing
    
    MetoEnGrilla vari, vari2, uProd.DESCRIPCION, Text1.Text, "X"
    uProd.codigo = ""
    uProd.DESCRIPCION = ""
    Text1.Text = ""
End Sub

Private Sub MetoEnGrilla(prod, barra, desc, Serie, marca)
    On Error GoTo ufaErr
    
    Dim i As Long
    
    i = g.addRow()
    If grilla.TextMatrix(i - 1, 0) = "" Then
        g.tx i - 1, gcod, prod
        g.tx i - 1, gcodBarra, barra
        g.tx i - 1, gdes, desc
        g.tx i - 1, gserie, Serie
        g.tx i - 1, gmarca, marca
    Else
        g.tx i, gcod, prod
        g.tx i, gcodBarra, barra
        g.tx i, gdes, desc
        g.tx i, gserie, Serie
        g.tx i, gmarca, marca
    End If
       
fin:
    Exit Sub
ufaErr:
    ufa "err al poner en grilla", Me.Name ', Err
    Resume fin
End Sub

Private Sub inigrilla()
    Set g = New LiGrilla
    
    g.init grilla, 4
        
    gcod = g.AddCol("               Codigo                 ")
    gcodBarra = g.AddCol("     Codigo de barra del producto     ")
    gdes = g.AddCol("                            Descripcion                                       ", "S")
    gserie = g.AddCol("            Serie            ")
    gmarca = g.AddCol(" Marcar ", "C")
        
    'grilla.SelectionMode = flexSelectionListBox
    
End Sub

Private Sub Command1_Click()
    Dim resu
    Dim r As Long
    Dim C As Long
    
    If uProd.codigo = "" Then Exit Sub
    
    r = grilla.rows
    C = grilla.cols
    
    resu = SerieStockRepetida(uProd.codigo)
    If resu > "" Then Text1.Text = resu 'grilla.TextMatrix(r, c) = resu
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    uProd.codigo = ""
    uProd.DESCRIPCION = ""
    Text1.Text = ""
    inigrilla
End Sub

Private Sub Command4_Click()
    Dim str As String
    Dim a As Long
    Dim cortar As Long
    'Dim rs As New ADODB.Recordset
    
    a = 1
    cortar = 0
    
    If grilla.rows = 1 Then Exit Sub
    While grilla.rows > a And cortar = 0
        If grilla.TextMatrix(a, gmarca) <> "" Then
            cortar = 1
        Else
            cortar = 0
        End If
        a = a + 1
    Wend
    
    If gEMPR_idEmpresa = 11 Then
        If cortar = 0 Then
            MsgBox "Se guardaran los numeros de serie pero no imprimira nada."
            'MsgBox "Debe marcar los productos a imprimir."
            a = 1
            While grilla.rows > a
                If grilla.TextMatrix(a, gcod) <> "" Then
                    DataEnvironment1.dbo_abmSERIEs "A", 0, grilla.TextMatrix(a, gcod), grilla.TextMatrix(a, gserie), 0, 0, 0, 0, "", 0, Date, False, Date, UsuarioActual()
                End If
                a = a + 1
            Wend
            Exit Sub
        End If
    Else
        If cortar = 0 Then
            MsgBox "No hay items marcados para imprimir.", , "ATENCION"
            Exit Sub
        End If
    End If
    a = 1
        
    If gEMPR_idEmpresa = 11 Then
        While grilla.rows > a
            If grilla.TextMatrix(a, gcod) <> "" Then
                DataEnvironment1.dbo_abmSERIEs "A", 0, grilla.TextMatrix(a, gcod), grilla.TextMatrix(a, gserie), 0, 0, 0, 0, "", 0, Date, False, Date, UsuarioActual()
            End If
            a = a + 1
        Wend
    End If
    
    a = 1
    Set etiq = New ADODB.Recordset

    With etiq
        ' Establece Cod como la clave principal.
        .Fields.Append "Cod", adChar, 24, adFldRowID
        .Fields.Append "Descripcion", adChar, 65, adFldUpdatable
        .Fields.Append "serie", adChar, 24, adFldUpdatable
        .Fields.Append "marca", adChar, 1, adFldUpdatable
        .Fields.Append "Cod2", adChar, 24, adFldRowID
        .Fields.Append "Descripcion2", adChar, 65, adFldUpdatable
        .Fields.Append "serie2", adChar, 24, adFldUpdatable
        .Fields.Append "marca2", adChar, 1, adFldUpdatable
        ' Utilice el tipo de cursor Keyset para permitir la actualización
        ' de los registros.
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With
        
    While grilla.rows > a
        'guardo los datos de la grilla en un recordset
        If grilla.TextMatrix(a, gmarca) <> "" Then
            If grilla.TextMatrix(a, gcodBarra) <> "" Then
                addrs2 grilla.TextMatrix(a, gcodBarra), grilla.TextMatrix(a, gdes), grilla.TextMatrix(a, gserie), grilla.TextMatrix(a, gmarca), grilla.TextMatrix(a + 1, gcod), grilla.TextMatrix(a + 1, gdes), grilla.TextMatrix(a + 1, gserie), grilla.TextMatrix(a + 1, gmarca)
            Else
                addrs2 grilla.TextMatrix(a, gcod), grilla.TextMatrix(a, gdes), grilla.TextMatrix(a, gserie), grilla.TextMatrix(a, gmarca), grilla.TextMatrix(a + 1, gcod), grilla.TextMatrix(a + 1, gdes), grilla.TextMatrix(a + 1, gserie), grilla.TextMatrix(a + 1, gmarca)
            End If
        End If
            
        a = a + 2
    Wend
    
'    If MsgBox("Desea ver los codigos de barra de los productos?", vbYesNo) = vbYes Then
        'con codigo de producto
        etiq.MoveFirst
        Set rptimpresionEtiquetaPro.DataEtiqueta.Recordset = etiq
        rptimpresionEtiquetaPro.Show
'    Else
'        'sin codigo de producto
'        etiq.MoveFirst
'        Set RptImpresionEtiqueta.DataEtiqueta.Recordset = etiq
'        RptImpresionEtiqueta.Show
'    End If
    
End Sub

Private Function addrs2(cod As String, des As String, Serie As String, marca As String, cod2 As String, des2 As String, Serie2 As String, marca2 As String) As Boolean
    With etiq
        .AddNew
        !cod = cod
        !DESCRIPCION = Left(des, 65)
        !Serie = Serie
        !marca = marca
        !cod2 = cod2
        !descripcion2 = Left(des2, 65)
        !Serie2 = Serie2
        !marca2 = marca2
        .Update
        '.Bookmark = .LastModified
    End With
End Function

Private Sub Command5_Click()
    grilla.RemoveItem grilla.Row
End Sub

Private Sub Form_Load()
    inigrilla
    set_uProd
End Sub

Private Sub set_uProd() ' lo copie de pedido cliente
    Dim sqlbuscar As String, sqldesc As String

'    If Propio() Then    'propio
        sqldesc = "select descripcion from producto where codigo = '###' "
        sqlbuscar = "select codigo as [ Codigo                 ], descripcion as [ Descripcion                                                 ] from producto where activo = 1 order by codigo "
'    Else    'relCliente
'        sqldesc = "select descripcion from producto  " _
'            & " inner join relacion_Producto_Cliente " _
'            & " on producto.codigo = relacion_Producto_cliente.producto " _
'            & " where cliente = " & cliente.codigo & " and productoCliente = '###'"
'        sqlbuscar = "select relacion_producto_cliente.productoCliente, producto.descripcion, producto.codigo, relacion_producto_cliente.precio " _
'            & " from producto  " _
'            & " inner join relacion_Producto_Cliente " _
'            & " on producto.codigo = relacion_Producto_cliente.producto " _
'            & " where cliente = " & cliente.codigo _
'            & " and producto.activo = 1 and relacion_producto_cliente.activo = 1 " _
'            & " order by producto"
'    End If
    uProd.ini sqldesc, sqlbuscar, True
    uProd.EditaDescripcion = True
End Sub

Public Function SerieStockRepetida(producto As String) As String
'    On Error GoTo UfaBuscaSer
    Dim ss As String, tmpTablaSeries As String
    Dim rs As New ADODB.Recordset
    
    'Dim kk As long, kk2 '*******************************
    'kk2 = Timer
    
    tmpTablaSeries = TablaTempCrear(tt_SeriesEnStockTemp)

If producto = "" Then
    ss = "INSERT INTO " & tmpTablaSeries & " ( Producto, Serie ) " _
    & " SELECT producto, serie From Series " _
    & " Where activo = 1  "

    DataEnvironment1.Sistema.Execute ss

    ss = "SELECT t.Serie as [ Serie               ], s.Producto as [ Producto                  ], t.Descripcion as [ Descripcion                                              ], s.comprobante as [c], t.codigo  as [i]" _
    & " FROM " & tmpTablaSeries & " AS t INNER JOIN Series AS s ON t.codigo = s.codigo left join conceptos as c on c.codigo = s.concepto " _
    & " WHERE s.comprobante = 6 or s.comprobante = 3 or s.comprobante = 4 or (s.comprobante = 7 and c.movimiento <> 'R' )  "

    SerieStockRepetida = frmBuscar.MostrarSql(ss)
    Exit Function
End If

ss = "SELECT DISTINCT producto,Serie FROM Series " _
     & " Where Series.activo = 1 and producto = '" & producto & "'" _
     & " ORDER BY Series.serie "
    
'ss = "SELECT DISTINCT producto,Serie FROM Series " _
'     & " Where Series.activo = 1 and producto = '" & producto & "' and fecha > '20000101'" _
'     & " ORDER BY Series.serie "
'
    
rs.Open (ss), DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly

    'kk = 0

Do While rs.EOF = False
    'kk = kk + 1
     If SerieEnStock((rs!Serie), producto) Then
           
      ss = "INSERT INTO " & tmpTablaSeries & " (Producto, Serie) " _
       & " VALUES ('" & producto & "','" & sSinNull(rs!Serie) & "')"
    
       DataEnvironment1.Sistema.Execute ss
         
     End If
     rs.MoveNext
Loop
        

ss = "SELECT Serie as [ Serie               ], Producto as [ Producto                  ]" _
    & " FROM " & tmpTablaSeries & " "
        
'        MsgBox kk & "   " & Timer - kk2

SerieStockRepetida = frmBuscar.MostrarSql(ss)
    
fin:
    Exit Function
UfaBuscaSer:
    ufa "err: buscando series", "Prod: " & producto, Err.Description
    Resume fin
End Function

Private Sub grilla_Click()
    Dim v As Long
    
    v = grilla.Row
    
    If grilla.ColSel = gmarca Then
        'MsgBox "col " & grilla.ColSel & " linea " & v
        If grilla.TextMatrix(v, gmarca) = "X" Then
            grilla.TextMatrix(v, gmarca) = ""
        Else
            grilla.TextMatrix(v, gmarca) = "X"
        End If
    End If
End Sub


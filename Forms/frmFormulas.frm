VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmFormulas 
   Caption         =   "Formula de producto"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7755
   Icon            =   "frmFormulas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   7755
   StartUpPosition =   2  'CenterScreen
   Begin Gestion.ucBotonera ucBotonera 
      Height          =   1575
      Left            =   210
      TabIndex        =   6
      Top             =   6075
      Width           =   7410
      _ExtentX        =   13070
      _ExtentY        =   2778
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   ""
   End
   Begin VB.Frame fraProducto 
      Height          =   1155
      Left            =   165
      TabIndex        =   1
      Top             =   0
      Width           =   7380
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   180
         Width           =   1830
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   300
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   630
         Width           =   4800
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   5
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "Descripcion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmdBorraItem 
      BackColor       =   &H00E0E0E0&
      DisabledPicture =   "frmFormulas.frx":08CA
      Height          =   255
      Left            =   6690
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmFormulas.frx":0E54
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Borrar Item"
      Top             =   1305
      Width           =   795
   End
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Height          =   4335
      Left            =   195
      TabIndex        =   7
      Top             =   1635
      Width           =   7395
      _cx             =   13044
      _cy             =   7646
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
Attribute VB_Name = "frmFormulas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents g As LiGrilla
Attribute g.VB_VarHelpID = -1
Private Const mDecimales = 4

Private gCODI As Long
Private gCOMP As Long
Private gCANT As Long
Private gDESC As Long
Private gUnid As Long
Private gConv As Long

Private Const MAXROWS = 50


Private Sub cmdBorraItem_Click()
    If g.Row > 0 Then g.delRow (g.Row)
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
        FrmKeyPress KeyAscii, True, True, True
End Sub

Private Sub Form_Load()
    gIni
    cmdBorraItem.enabled = False
    CargaBotonera

    ucBotonera.MsgConfirmaEliminar = "Confirma Eliminacion"
    Form_Resize
End Sub


Private Sub cargaGrilla()
On Error GoTo fin
    Dim rs As New ADODB.Recordset, i As Long, des As String, uni
   
    With rs
        .Open "select f.codigo, f.componente, f.cantidad from formulas as f where f.activo = 1 and f.codigo = '" & Trim(txtCodigo) & "'", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
        g.Borrar
        While Not .EOF
            i = g.addRow()
            g.tx i, gCOMP, !componente
            g.tx i, gCANT, !cantidad
            des = obtenerDeSQL("select descripcion from producto where codigo = '" & !componente & "'")
            g.tx i, gDESC, des
            .MoveNext
        Wend
    End With
    Set rs = Nothing
    With g
        For i = 1 To .rows - 1
            uni = obtenerDeSQL("select umedida from producto where codigo=" & sstexto(g.tx(i, gCOMP)))
            g.tx i, gUnid, obtenerDeSQL("select abreviatura from unidadesmedida where umcodigo=" & uni)
        Next
    End With
Exit Sub
fin:
    MsgBox "Error en carga.", vbCritical
End Sub


Private Sub gIni()
    Set g = New LiGrilla
    g.init grilla, 4
    g.rows = MAXROWS
    gCOMP = g.AddCol("      Componente             ")
    gCANT = g.AddCol(" Cantidad ", "N", mDecimales)
    gDESC = g.AddCol(" Descripcion Componente                                ")
    gUnid = g.AddCol(" Medida  ")
    'gConv = g.AddCol(" Conversion  ")
End Sub


Private Function FormulaAlta() As Boolean
    If ON_ERROR_HABILITADO Then On Error GoTo UfaAlta
    
    Dim pr As String, i As Long
    Dim compo As String, cant As Double
    
    If Trim(txtCodigo) = "" Then
        MsgBox "Ingrese un codigo", vbExclamation
        Exit Function
    End If
    If g.suma(gCANT) = 0 Then
        MsgBox "Ingrese Componentes", vbExclamation
        Exit Function
    End If
    
    For i = 1 To g.rows - 1
        If g.tx(i, gCOMP) > "" And s2n(g.tx(i, gCANT), mDecimales) = 0 Then
            MsgBox "Algun Productos estan sin cantidad", vbExclamation
            Exit Function
        End If
    Next i
    
    FormulaBaja
    For i = 1 To g.rows - 1
        compo = Trim(g.tx(i, gCOMP))
        cant = s2n(g.tx(i, gCANT), mDecimales)
        If compo > "" And cant > 0 Then
            ABMFormula "A", txtCodigo, compo, cant
        End If
    Next i
    
    FormulaAlta = True
Exit Function
UfaAlta:
    DE_RollbackTrans
    FormulaAlta = False
    MsgBox "Error al grabar", vbCritical
End Function

Private Function FormulaBaja() As Boolean
On Error GoTo UFAbaja
    ABMFormula "B", txtCodigo, "", 0
    FormulaBaja = True
Exit Function
UFAbaja:
    MsgBox "Error al eliminar", vbCritical
End Function

Private Sub Form_Resize()
    'Anclar grilla, me, anclarLadosTodos
End Sub

Private Sub g_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If ucBotonera.Estado <> ucbEditando Then Exit Sub
End Sub

Private Sub g_DblClick()
    Dim re As String, mAnchos
    If ucBotonera.Estado <> ucbEditando Or g.Row = 0 Or g.Col <> gCOMP Then Exit Sub
    mAnchos = Array(1500, 2500, 500)
    re = frmBuscar.MostrarSql("Select p.Codigo, p.Descripcion, m.Abreviatura as Medida from producto as p inner join unidadesmedida as m on p.umedida=m.umcodigo where p.activo=1", mAnchos)
    If re = "" Then Exit Sub
    
    g.tx g.Row, gCOMP, re
    g.tx g.Row, gDESC, frmBuscar.resultado(2)
    g.tx g.Row, gUnid, frmBuscar.resultado(3)
End Sub

Private Sub BorrarTodo()
    txtCodigo = ""
    txtDescripcion = ""
    g.Borrar
    g.rows = MAXROWS
End Sub


Private Sub CargaBotonera()
    Dim sqlRS As String
    
    sqlRS = "select distinct producto.codigo, producto.descripcion from producto inner join formulas on producto.codigo = formulas.codigo where formulas.activo = 1"
    ucBotonera.init True, True, True, False, True, sqlRS, DataEnvironment1.Sistema
End Sub

Private Sub ucBotonera_Aceptar()
    If FormulaAlta() Then
        MsgBox "Formula Guardada.", vbInformation
        ucBotonera.AceptarOk
    End If
End Sub

Private Sub ucBotonera_BorrarControles()
    BorrarTodo
End Sub

Private Sub ucBotonera_Buscar()
    Dim resu As String
    resu = frmBuscar.MostrarSql("select distinct producto.codigo as [ Codigo                          ], producto.descripcion as [  Descripcion                                             ] from producto inner join formulas on producto.codigo = formulas.codigo where formulas.activo = 1")
    If resu > "" Then
        txtCodigo = resu
        txtDescripcion = frmBuscar.resultado(2)
        cargaGrilla
        ucBotonera.BuscarOK "codigo = '" & resu & "'"
    End If
End Sub
Private Sub ucBotonera_Eliminar()
    If FormulaBaja() Then
        ucBotonera.EliminarOK
        MsgBox "Formula eliminada.", vbInformation
    End If
End Sub
Private Sub ucBotonera_HabilitarEdicion(sino As Boolean)
    cmdBorraItem.enabled = sino
    grilla.Editable = IIf(sino, flexEDKbdMouse, flexEDNone)
End Sub
Private Sub ucBotonera_Modificar()
    g.rows = MAXROWS
End Sub
Private Sub ucBotonera_Nuevo()
    Dim re As String
    re = frmBuscar.MostrarCodigoDescripcionActivo("producto")
    If re = "" Then Exit Sub
    txtCodigo = re
    txtDescripcion = frmBuscar.resultado(2)

    cargaGrilla
    g.rows = MAXROWS
End Sub
Private Sub ucBotonera_SALIR()
    Unload Me
End Sub
Private Sub ucBotonera_SeMovio()
    txtCodigo = ucBotonera.rs!codigo
    txtDescripcion = ucBotonera.rs!DESCRIPCION
    cargaGrilla
End Sub

Private Function ABMFormula(fOpe As String, fcodigo As String, fComponente As String, fCantidad As Double) As Boolean
On Error GoTo MAL
Dim ABMF As String
ABMFormula = True
Select Case fOpe
    Case "A":
        DE_BeginTrans
        ABMF = "INSERT INTO FORMULAS (CODIGO,COMPONENTE,CANTIDAD,FECHA_ALTA,ACTIVO) " _
            & " Values (" & sstexto(fcodigo) & "," & sstexto(fComponente) & "," & x2s(fCantidad) & "," & ssFecha(Date) & ",1)"
        DataEnvironment1.Sistema.Execute ABMF
        ABMF = " UPDATE PRODUCTO SET FORMULA=1 WHERE CODIGO=" & sstexto(fcodigo)
        DataEnvironment1.Sistema.Execute ABMF
        'MsgBox "Formula Guardada", vbInformation
        DE_CommitTrans
    Case "B":
        DE_BeginTrans
        ABMF = " UPDATE FORMULAS  SET ACTIVO=0, FECHA_BAJA=" & ssFecha(Date) & ", USUARIO_BAJA= 2 WHERE CODIGO=" & sstexto(fcodigo)
        DataEnvironment1.Sistema.Execute ABMF
        ABMF = " UPDATE PRODUCTO SET FORMULA = 0 WHERE CODIGO=" & sstexto(fcodigo)
        DataEnvironment1.Sistema.Execute ABMF
        'MsgBox "Formula Eliminada", vbInformation
        DE_CommitTrans
End Select
Exit Function
MAL:
ABMFormula = False
    MsgBox "Error en carga de formula.", vbCritical, "Error en formula"
End Function

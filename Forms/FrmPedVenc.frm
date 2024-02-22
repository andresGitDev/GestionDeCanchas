VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPedVenc 
   Caption         =   "Estado de Pedidos"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   11880
   Icon            =   "FrmPedVenc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdProcesar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Procesar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10500
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5100
      Width           =   1215
   End
   Begin VB.CommandButton CmdEjecutar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Recalcular"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10500
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1695
      Width           =   1215
   End
   Begin VB.ComboBox CmbProx 
      Height          =   315
      Left            =   10500
      TabIndex        =   9
      Top             =   1185
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker ProxVenc 
      Height          =   330
      Left            =   10455
      TabIndex        =   7
      Top             =   4590
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   582
      _Version        =   393216
      Format          =   17235969
      CurrentDate     =   38547
   End
   Begin VSFlex7LCtl.VSFlexGrid GVencidos 
      Height          =   3225
      Left            =   105
      TabIndex        =   2
      Top             =   4215
      Width           =   10275
      _cx             =   18124
      _cy             =   5689
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
      ForeColor       =   255
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483633
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
      Rows            =   1
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
   Begin VSFlex7LCtl.VSFlexGrid GPorVencer 
      Height          =   3225
      Left            =   105
      TabIndex        =   1
      Top             =   615
      Width           =   10275
      _cx             =   18124
      _cy             =   5689
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
      Rows            =   1
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
   Begin VB.CommandButton CmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10500
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7065
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Proximidad de Vencimiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10500
      TabIndex        =   10
      Top             =   705
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Proximo Vencimiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   10380
      TabIndex        =   8
      Top             =   4065
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Pedidos Vencidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   6
      Top             =   3930
      Width           =   2430
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Pedidos por Vencer"
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
      Left            =   150
      TabIndex        =   5
      Top             =   360
      Width           =   2430
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Actual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   9420
      TabIndex        =   4
      Top             =   135
      Width           =   855
   End
   Begin VB.Label LblFecha 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   10365
      TabIndex        =   3
      Top             =   105
      Width           =   1380
   End
End
Attribute VB_Name = "FrmPedVenc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rsClie As New ADODB.Recordset
Dim CalcFecha, DiaPost As Date
Dim x As Long


Private Sub CmbProx_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   CmdEjecutar_Click
End If
End Sub

Private Sub CmdEjecutar_Click()
Dim x As Long
GVencidos.clear
GPorVencer.clear
If GVencidos.rows > 1 Then
   For x = (GVencidos.rows - 1) To 1 Step -1
    GVencidos.RemoveItem (x)
   Next
End If
If GPorVencer.rows > 1 Then
   For x = (GPorVencer.rows - 1) To 1 Step -1
      GPorVencer.RemoveItem (x)
   Next
End If
Armado
End Sub

Private Sub CmdProcesar_Click()
Dim x As Long
Dim Check As Boolean
Check = ControlPrevio(4)
'If Check = True Then PonerFecha
If Check = True Then
For x = 1 To GVencidos.rows - 1
  Select Case GVencidos.TextMatrix(x, 4)
  Case "Bajar"
    DataEnvironment1.dbo_PEDIDOCLIENTE "B", GVencidos.TextMatrix(x, 0), 0, 0, 0, 0, 0, 0, "", 0, 0, 0, 0, 0, UsuarioActual(), Date
'
   DataEnvironment1.Sistema.Execute "UPDATE PEDIDOS_CLIENTES SET ANULAxSUPERVISOR=1 WHERE numero = '" & GVencidos.TextMatrix(x, 0) & "'"
   grabaBitacora "B", s2n(TxtNro), "Pedidos_Clientes"

  Case "Postergar"
     DataEnvironment1.dbo_PEDIDOCLIENTE "R", GVencidos.TextMatrix(x, 0), 0, 0, 0, 0, 0, 0, "", 0, 0, GVencidos.TextMatrix(x, 5), 0, 0, 0, 0
  Case Else
      '
  End Select
Next
CmdEjecutar_Click
Else
    MsgBox "Debe seleccionar una Accion Postergar/Bajar", vbInformation, "Aviso"
End If
End Sub
Private Sub PonerFecha()

End Sub
Private Function ControlPrevio(NColumna As Long) As Boolean
Dim x As Long
ControlPrevio = False

For x = 1 To GVencidos.rows - 1
   If GVencidos.TextMatrix(x, NColumna) <> "" Then
     ControlPrevio = True
     
   Else
'      ControlPrevio = False
 '     If NColumna = 4 Then
 '        MsgBox "Hay pedidos sin Proceso Elegido", vbExclamation, "Falta de Opcion"
  '    End If
'      If NColumna = 5 Then
'         MsgBox "Hay pedidos sin Fecha de Postergacion", vbExclamation, "Falta de Fecha"
'      End If
'      Exit Function
   End If
Next
End Function
Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
lblfecha = Date
CargoProxVenc
GVencidos.Editable = flexEDKbdMouse
For x = 1 To 30
   CmbProx.AddItem x
Next
x = 0
CmbProx.Text = 3
Armado
End Sub

Private Sub Armado()
DiaPost = Date + CmbProx.Text
InicioGrilla

sql = "SELECT DISTINCT Pedidos_Clientes.numero, Pedidos_Clientes.fecha, Clientes.descripcion, " _
& "Pedidos_Clientes.Vencimiento, Sum(ItemPedidoCliente.Saldo) AS SumaDeSaldo, Sum(ItemPedidoCliente.Cantidad) " _
& "AS SumaDeCantidad FROM Clientes INNER JOIN (ItemPedidoCliente INNER JOIN Pedidos_Clientes " _
& "ON ItemPedidoCliente.PEDIDO = Pedidos_Clientes.numero) ON Clientes.codigo = Pedidos_Clientes.cliente " _
& "GROUP BY Pedidos_Clientes.numero, Pedidos_Clientes.fecha, Clientes.descripcion, Pedidos_Clientes.Vencimiento, " _
& "Pedidos_Clientes.activo Having (((Pedidos_Clientes.Vencimiento) < " & ssFecha(Date) & " Or " _
& "(Pedidos_Clientes.Vencimiento) <= " & ssFecha(DiaPost) & ") And ((Sum(ItemPedidoCliente.Saldo)) = " _
& "Sum([cantidad])) And ((Pedidos_Clientes.activo) = 1))ORDER BY Pedidos_Clientes.numero"

rs.Open sql, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
Do While Not rs.EOF
   If rs!Vencimiento < Date Then
      GVencidos.AddItem rs!numero & Chr(9) & rs!DESCRIPCION & Chr(9) & rs!fecha & Chr(9) & rs!Vencimiento
   Else
      GPorVencer.AddItem rs!numero & Chr(9) & rs!DESCRIPCION & Chr(9) & rs!fecha & Chr(9) & rs!Vencimiento
   End If
   rs.MoveNext
Loop
rs.Close
End Sub

Private Sub CargoProxVenc()
Dim CantDias As Long
   sql = "Select diasvenc from bs"
   rs.Open sql, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
   CantDias = rs!diasvenc
   rs.Close
ProxVenc.Value = Date + CantDias
End Sub
Private Sub InicioGrilla()
   GPorVencer.cols = 4
   GPorVencer.TextMatrix(0, 0) = "Pedido Nº"
   GPorVencer.TextMatrix(0, 1) = "Cliente"
   GPorVencer.ColWidth(1) = 3000
   GPorVencer.TextMatrix(0, 2) = "Fecha Pedido"
   GPorVencer.ColWidth(2) = 1500
   GPorVencer.TextMatrix(0, 3) = "Fecha de Vencimiento"
   GPorVencer.ColWidth(3) = 1700
   
   GVencidos.cols = 6
   GVencidos.TextMatrix(0, 0) = "Pedido Nº"
   GVencidos.TextMatrix(0, 1) = "Cliente"
   GVencidos.ColWidth(1) = 3000
   GVencidos.TextMatrix(0, 2) = "Fecha Pedido"
   GVencidos.ColWidth(2) = 1500
   GVencidos.TextMatrix(0, 3) = "Fecha de Vencimiento"
   GVencidos.ColWidth(3) = 1700
   GVencidos.TextMatrix(0, 4) = "Postergar/Bajar"
   GVencidos.ColWidth(4) = 1500
   GVencidos.ColComboList(4) = "Postergar|Bajar"
   GVencidos.TextMatrix(0, 5) = "Prox. Vencimiento"
   GVencidos.ColWidth(5) = 1500
End Sub

Private Sub GPorVencer_DblClick()
If GPorVencer.Row <> 0 Then
 GVencidos.AddItem GPorVencer.TextMatrix(GPorVencer.Row, 0) & Chr(9) & _
 GPorVencer.TextMatrix(GPorVencer.Row, 1) & Chr(9) & GPorVencer.TextMatrix(GPorVencer.Row, 2) & _
 Chr(9) & GPorVencer.TextMatrix(GPorVencer.Row, 3)
 GPorVencer.RemoveItem (GPorVencer.Row)
End If
End Sub

Private Sub GVencidos_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Col = 4 Then
   If GVencidos.TextMatrix(GVencidos.Row, 4) = "Postergar" Then
      GVencidos.TextMatrix(GVencidos.Row, 5) = ProxVenc.Value
   End If
End If
End Sub

Private Sub GVencidos_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = (Col <> 4)
End Sub

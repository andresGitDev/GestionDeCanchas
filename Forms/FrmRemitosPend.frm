VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form FrmRemitosPend 
   Caption         =   "Remitos Pendientes / Cancelados"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11370
   Icon            =   "FrmRemitosPend.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   11370
   Begin Gestion.ucEntreFechas EntreFechas 
      Height          =   285
      Left            =   1920
      TabIndex        =   8
      Top             =   240
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   503
   End
   Begin VB.OptionButton OptTodos 
      Caption         =   "Todos"
      Height          =   300
      Index           =   0
      Left            =   7260
      TabIndex        =   7
      Top             =   150
      Width           =   1290
   End
   Begin VB.OptionButton OptTodos 
      Caption         =   "Totalmente Facturado"
      Height          =   300
      Index           =   3
      Left            =   8835
      TabIndex        =   6
      Top             =   465
      Width           =   2370
   End
   Begin VB.OptionButton OptTodos 
      Caption         =   "Parcialmente Facturado"
      Height          =   300
      Index           =   2
      Left            =   8835
      TabIndex        =   5
      Top             =   150
      Width           =   2355
   End
   Begin VB.OptionButton OptTodos 
      Caption         =   "Sin Facturar"
      Height          =   300
      Index           =   1
      Left            =   7260
      TabIndex        =   4
      Top             =   465
      Width           =   1275
   End
   Begin VB.CommandButton CmdEjecutar 
      Caption         =   "Ejecutar"
      Height          =   435
      Left            =   5910
      TabIndex        =   3
      Top             =   210
      Width           =   900
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   435
      Left            =   10305
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   900
   End
   Begin VSFlex7LCtl.VSFlexGrid Grilla 
      Height          =   4185
      Left            =   120
      TabIndex        =   0
      Top             =   1125
      Width           =   11115
      _cx             =   19606
      _cy             =   7382
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
      Rows            =   1
      Cols            =   7
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
      AutoSearch      =   2
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
   Begin VB.Shape Shape1 
      Height          =   945
      Left            =   7110
      Top             =   45
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rango de Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   300
      Left            =   180
      TabIndex        =   1
      Top             =   240
      Width           =   1875
   End
End
Attribute VB_Name = "FrmRemitosPend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim ssql As String
Dim TipoOpt As Long
Dim TipoOpt2 As Long

Private Function andWhere() As String
    
    
End Function

Private Sub CmdEjecutar_Click()
    Dim rsRem As New ADODB.Recordset
    Dim corte As Single
    Dim AEstado As String
    
    Cargogrilla

    ssql = "SELECT RemitoVenta.Numero, Clientes.descripcion, RemitoVentaDetalle.Producto, " & _
     "RemitoVentaDetalle.Cantidad,RemitoVentaDetalle.Facturar , RemitoVenta.fecha FROM Clientes " & _
     "INNER JOIN RemitoVenta ON Clientes.codigo = RemitoVenta.Cliente INNER JOIN " & _
     "RemitoVentaDetalle ON RemitoVenta.Numero = RemitoVentaDetalle.Numero " & _
     "WHERE RemitoVenta.Fecha " & EntreFechas.ssBetween & _
     " order by RemitoVenta.Numero, RemitoVentaDetalle.codigo "
     'BETWEEN '" & EntreFechas.desde & "' AND '" & EntreFechas.hasta & "')"
     
    rsRem.Open ssql, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
    If rsRem.EOF = True And rsRem.BOF = True Then
        MsgBox "No se encontraron registros para mostrar."
        Exit Sub
    End If
    corte = rsRem!numero
    AEstado = Estados(corte)
    
    
    If OptTodos(0).Value = True Then 'todos
        Do While Not rsRem.EOF
            grilla.AddItem AEstado & Chr(9) & rsRem!numero & Chr(9) & rsRem!DESCRIPCION & Chr(9) & rsRem!producto & Chr(9) & rsRem!cantidad & Chr(9) & s0(rsRem!facturar) & Chr(9) & rsRem!fecha
            AEstado = ""
        

            rsRem.MoveNext
            If Not rsRem.EOF Then
                If corte <> rsRem!numero Then
                    corte = rsRem!numero
                    AEstado = Estados(corte)
                    grilla.AddItem ""
                End If
            End If
        Loop
    ElseIf OptTodos(1).Value = True Then 'sin facturar
        Do While Not rsRem.EOF
            If AEstado = "Sin Facturar" Then
                grilla.AddItem AEstado & Chr(9) & rsRem!numero & Chr(9) & rsRem!DESCRIPCION & Chr(9) & rsRem!producto & Chr(9) & rsRem!cantidad & Chr(9) & s0(rsRem!facturar) & Chr(9) & rsRem!fecha
                grilla.AddItem ""
            End If
            AEstado = ""
        

            rsRem.MoveNext
            If Not rsRem.EOF Then
                If corte <> rsRem!numero Then
                    corte = rsRem!numero
                    AEstado = Estados(corte)
                End If
            End If
        Loop
    ElseIf OptTodos(2).Value = True Then 'parcialmente facturado
        Do While Not rsRem.EOF
            If AEstado = "Parcialmente Facturarado" Then
                grilla.AddItem AEstado & Chr(9) & rsRem!numero & Chr(9) & rsRem!DESCRIPCION & Chr(9) & rsRem!producto & Chr(9) & rsRem!cantidad & Chr(9) & s0(rsRem!facturar) & Chr(9) & rsRem!fecha
                grilla.AddItem ""
            End If
            AEstado = ""
        

            rsRem.MoveNext
            If Not rsRem.EOF Then
                If corte <> rsRem!numero Then
                    corte = rsRem!numero
                    AEstado = Estados(corte)
                End If
            End If
        Loop
    ElseIf OptTodos(3).Value = True Then 'totalmente facturado
        Do While Not rsRem.EOF
            If AEstado = "Totalmente Facturado" Then
                grilla.AddItem AEstado & Chr(9) & rsRem!numero & Chr(9) & rsRem!DESCRIPCION & Chr(9) & rsRem!producto & Chr(9) & rsRem!cantidad & Chr(9) & s0(rsRem!facturar) & Chr(9) & rsRem!fecha
                grilla.AddItem ""
            End If
            AEstado = ""
        

            rsRem.MoveNext
            If Not rsRem.EOF Then
                If corte <> rsRem!numero Then
                    corte = rsRem!numero
                    AEstado = Estados(corte)
                End If
            End If
        Loop
    End If
    

'    rsRem.Open ssql, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
'    If rsRem.EOF = True And rsRem.BOF = True Then
'        MsgBox "No se encontraron registros para mostrar."
'        Exit Sub
'    End If
'    corte = rsRem!numero
'    AEstado = Estados(corte)
    
'    Do While Not rsRem.EOF
'        grilla.AddItem AEstado & Chr(9) & rsRem!numero & Chr(9) & rsRem!descripcion & Chr(9) & rsRem!producto & Chr(9) & rsRem!cantidad & Chr(9) & s0(rsRem!facturar) & Chr(9) & rsRem!fecha
'        AEstado = ""
        

'        rsRem.MoveNext
'        If Not rsRem.EOF Then
'            If corte <> rsRem!numero Then
'                corte = rsRem!numero
'                AEstado = Estados(corte)
'                grilla.AddItem ""
'            End If
'        End If
'    Loop

    rsRem.Close
End Sub
Private Function s0(que)
    s0 = IIf(que = 0, "", que)
End Function
Function Estados(valor As Single) As String
Dim rsSuma As New ADODB.Recordset
'    ssql = "SELECT SUM(remitoventadetalle.cantidad) as SC, SUM(remitoventadetalle.facturar) as SF " & _
'                "FROM remitoventadetalle WHERE numero=" & Valor & " " & _
'                "GROUP BY cantidad,facturar"
'    rsSuma.Open ssql, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
'    If Not rsSuma.EOF Then
'        If rsSuma!sc > rsSuma!sf Then
'            Estados = "Parcialmente Facturarado"
'            TipoOpt2 = 2
'        End If
'        If rsSuma!sc = rsSuma!sf Then
'            Estados = "Sin Facturar"
'            TipoOpt2 = 1
'        End If
'        If rsSuma!sf = 0 Then
'            Estados = "Totalmente Facturado"
'            TipoOpt2 = 3
'        End If
'    End If
    ssql = "SELECT sum(cantidad) as sc, SUM(facturar) as SF " & _
                "FROM remitoventadetalle WHERE numero=" & valor & " " & _
                "GROUP BY numero"
    rsSuma.Open ssql, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
    If Not rsSuma.EOF Then
        If rsSuma!sf = 0 Then
            Estados = "Totalmente Facturado"
            TipoOpt2 = 3
        ElseIf rsSuma!sc = rsSuma!sf Then
            Estados = "Sin Facturar"
            TipoOpt2 = 1
        Else
        'rsSuma!sc > rsSuma!sf Then
            Estados = "Parcialmente Facturarado"
            TipoOpt2 = 2
        End If
    End If

    If OptTodos(0).Value = True Then TipoOpt2 = 0
    rsSuma.Close
End Function
Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
    OptTodos(0).Value = True
    TipoOpt = 1
    Cargogrilla
End Sub

Sub Cargogrilla()
    grilla.clear
    grilla.rows = 1
    grilla.TextMatrix(0, 0) = " Estado "
    grilla.TextMatrix(0, 1) = " Remito Numero "
    grilla.TextMatrix(0, 2) = "Cliente"
    grilla.TextMatrix(0, 3) = "Producto"
    grilla.TextMatrix(0, 4) = "Cantidad"
    grilla.TextMatrix(0, 5) = "Sin Facturar"
    grilla.TextMatrix(0, 6) = "Fecha"
    grilla.ColWidth(0) = 1000
    grilla.ColWidth(1) = 1500
    grilla.ColWidth(2) = 3200
    grilla.ColWidth(3) = 1500
    grilla.ColWidth(4) = 1200
    grilla.ColWidth(5) = 1000
    grilla.ColWidth(6) = 1300

End Sub

Private Sub Form_Resize()
    Anclar grilla, Me, anclarLadosTodos
    Anclar cmdsalir, Me, anclarAbajo + anclarDerecha
End Sub

Private Sub opttodos_Click(Index As Integer)
 TipoOpt = Index
End Sub

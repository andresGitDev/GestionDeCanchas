VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmListadosBancarios 
   Caption         =   "Listado por Nº de Movimiento"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10710
   Icon            =   "FrmListadosBancarios.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   10710
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Height          =   5655
      Left            =   105
      TabIndex        =   10
      Top             =   1200
      Width           =   10395
      _cx             =   18336
      _cy             =   9975
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
      Rows            =   50
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
   Begin VB.Frame fraMenu 
      BorderStyle     =   0  'None
      Height          =   1125
      Left            =   5145
      TabIndex        =   9
      Top             =   225
      Width           =   5475
      Begin VB.CommandButton cmdImprimir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Imprimir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3315
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "3"
         Top             =   40
         Width           =   975
      End
      Begin Gestion.ucXls uXls 
         Height          =   840
         Left            =   2100
         TabIndex        =   4
         Top             =   60
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   1482
      End
      Begin VB.CommandButton cmdsalir 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4470
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "4"
         Top             =   40
         Width           =   975
      End
      Begin VB.CommandButton cmdaceptar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Mostrar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "2"
         Top             =   40
         Width           =   975
      End
      Begin VB.CommandButton cmdcancelar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1050
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "3"
         Top             =   40
         Width           =   975
      End
   End
   Begin MSComCtl2.DTPicker fechadesde 
      Height          =   330
      Left            =   1050
      TabIndex        =   0
      Tag             =   "0"
      Top             =   330
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   582
      _Version        =   393216
      Format          =   284229633
      CurrentDate     =   38052
   End
   Begin MSComCtl2.DTPicker fechahasta 
      Height          =   300
      Left            =   3360
      TabIndex        =   1
      Tag             =   "1"
      Top             =   345
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   529
      _Version        =   393216
      Format          =   284229633
      CurrentDate     =   38052
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   735
      Left            =   120
      Top             =   120
      Width           =   4890
   End
   Begin VB.Label Label5 
      Caption         =   "Desde"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Hasta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2790
      TabIndex        =   7
      Top             =   345
      Width           =   615
   End
End
Attribute VB_Name = "FrmListadosBancarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' 5/10/4

Private Sub cmdAceptar_Click()
           
    Dim rs As New ADODB.Recordset ', rs1 As New ADODB.Recordset
    Dim fecha As String, Banco As String, nrocheque As String
    Dim tempo
    Dim debito As Double, credito As Double
    Dim descri As String, inte As Long
    Dim ss As String

    relojito
    DataEnvironment1.dbo_LISTADOXMOVIMIENTO "B", 0, "", 0, 0, "", 0, "", "", 0
    
    With rs
        ss = "select * from movibanc where activo = 1 and fecha " & ssBetween(FechaDesde, FechaHasta) & " "
        .Open ss, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        
        While Not .EOF
             
            nrocheque = ""
            Banco = ""
            If !documento = "P" Then
                tempo = obtenerDeSQL("select chq_comp.nro, bancosgrales.descripcion from chq_comp inner join bancosgrales on chq_comp.banco = bancosgrales.codigo where chq_comp.codigo = " & !interno & " ")
                If Not IsEmpty(tempo) Then
                    nrocheque = sSinNull(tempo(0))
                    Banco = sSinNull(tempo(1))
                End If
            ElseIf rs!documento = "C" Then
                tempo = obtenerDeSQL("select cheques.nro, bancosgrales.descripcion from cheques inner join bancosgrales on cheques.banco_nro = bancosgrales.codigo where cheques.nroint = " & !interno & " ")
                If Not IsEmpty(tempo) Then
                    nrocheque = sSinNull(tempo(0))
                    Banco = sSinNull(tempo(1))
                End If
            End If
                   
            debito = 0
            credito = 0
            If InStr("AEDT", !OPERACION) > 0 Then
                credito = !importe
            Else
                debito = !importe
            End If
            
            descri = ssStr(!DESCRIPCION)
            inte = s2n(!interno)
            
            DataEnvironment1.dbo_LISTADOXMOVIMIENTO "A", !fecha, descri, debito, credito, !documento, inte, Banco, nrocheque, !MovBanco
            
            .MoveNext
        Wend
        .Close
        
    End With
    LlenarGrilla grilla, "select * from LISTADOMOVIMIENTO order by movbanco", False
    
    
    
    'DataEnvironment1.dbo_LISTADOXMOVIMIENTO "B", 0, "", 0, 0, "", 0, "", "", 0
    
    relojito False
    Set rs = Nothing
End Sub



Private Sub cmdCancelar_Click()
    fechin
End Sub

Private Sub cmdImprimir_Click()
    If grilla.rows < 2 Then Exit Sub
    '****REPORTE ANTERIOR DE VISUAL****
    'DataEnvironment1.LisPorMovimiento
    'rptPorMovimiento.Sections("Section4").Controls("Label31").caption = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
    'rptPorMovimiento.Show vbModal
    'DataEnvironment1.rsLisPorMovimiento.Close
    
    '****REPORTE NUEVO DE CRISTAL****
    Dim rs As New ADODB.Recordset ' PARA SACAR LA CANTIDAD
    rs.Open "select * from listadomovimiento order by fecha", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    RptMovimientoBancario.data1.Connection = DataEnvironment1.Sistema
    RptMovimientoBancario.data1.Source = "select * from listadomovimiento order by fecha"
    RptMovimientoBancario.Field10 = rs.RecordCount 'CUENTO LA CANTIDAD DE REGISTRO
    RptMovimientoBancario.Show vbModal

End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True
End Sub
Private Sub Form_Load()
    fechin
    uXls.ini grilla, "LisMovBanco", "Listado de movimientos bancarios " & Date
    Form_Resize
End Sub
Private Sub fechin()
    FechaDesde = CDate("01/" & Month(Date) & "/" & Year(Date))
    FechaHasta = CDate("31/12/" & Year(Date))
    grilla.rows = 1
End Sub
Private Sub Form_Resize()
    Anclar grilla, Me, anclarLadosTodos
'    Anclar fraMenu, Me, anclarIzquierda + anclarAbajo
End Sub

VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form FrmSaldoCuentaProv 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Saldo de cuenta de Proveedores"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11910
   Icon            =   "FrmSaldoCuentaProv.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   6915
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraBoton 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "PorReporte"
      Height          =   1155
      Left            =   120
      TabIndex        =   11
      Top             =   5640
      Width           =   11655
      Begin VB.OptionButton oPantalla 
         Caption         =   "Pantalla"
         Height          =   330
         Left            =   7110
         TabIndex        =   18
         Top             =   60
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton oPorReporte 
         Caption         =   "Imprimir"
         Height          =   330
         Left            =   5970
         TabIndex        =   17
         Top             =   75
         Width           =   1125
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
         Left            =   9375
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   30
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
         Left            =   8400
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   30
         Width           =   975
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
         Left            =   10590
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   30
         Width           =   975
      End
      Begin VB.CommandButton cmdexcel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Enviar a Excel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   90
         Visible         =   0   'False
         Width           =   810
      End
      Begin Gestion.ucXls uXls 
         Height          =   945
         Left            =   1080
         TabIndex        =   12
         Top             =   90
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1667
      End
   End
   Begin VB.Frame Framedevol 
      BackColor       =   &H00E0E0E0&
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
      Height          =   735
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   4815
      Begin VB.OptionButton optuno 
         Caption         =   "Elegir Proveedor"
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
         Height          =   240
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1950
      End
      Begin VB.OptionButton opttodos 
         Caption         =   "Todos los Proveedores"
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
         Height          =   240
         Left            =   2070
         TabIndex        =   1
         Tag             =   "1"
         Top             =   255
         Width           =   2625
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   4695
      Left            =   0
      TabIndex        =   7
      Top             =   840
      Width           =   11895
      Begin VSFlex7LCtl.VSFlexGrid Grilla 
         Height          =   4335
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   11655
         _cx             =   20558
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
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
         MergeCells      =   2
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
         AllowUserFreezing=   2
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
   Begin VB.Frame fraFiltro 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   4920
      TabIndex        =   19
      Top             =   15
      Visible         =   0   'False
      Width           =   6840
      Begin VB.OptionButton optFiltro 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Saldo Negativo"
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
         Height          =   315
         Index           =   2
         Left            =   3615
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "1"
         Top             =   240
         Width           =   1590
      End
      Begin VB.OptionButton optFiltro 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Saldo Positivo"
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
         Height          =   315
         Index           =   1
         Left            =   1980
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "1"
         Top             =   240
         Width           =   1590
      End
      Begin VB.OptionButton optFiltro 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Todos"
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
         Height          =   315
         Index           =   0
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "1"
         Top             =   240
         Value           =   -1  'True
         Width           =   1590
      End
   End
   Begin VB.Frame fraProveedor 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   4920
      TabIndex        =   6
      Top             =   15
      Width           =   6855
      Begin Gestion.ucCoDe uprov 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   240
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   556
         CodigoWidth     =   1000
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmSaldoCuentaProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
    
    Dim rs As New ADODB.Recordset
    Dim saldo As Double, provant, numint As Long
    Dim saldo1 As Double
    Dim provDesc As String
    Dim FechaVencimiento As String
    Dim titux As String
    
    If opttodos = True Or optuno = True Then
    
        DataEnvironment1.dbo_SALDOPROVTEMP "B", 0, "", 0, "", 0, 0, 0, 0, 0, 0
    
        If opttodos = True Then
            rs.Open "select * from transcom where activo = 1 order by codpr, fecha", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
        Else
            rs.Open "select * from transcom where codpr = " & UpROV.codigo & " and activo = 1 order by fecha", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
        End If
        
        numint = 0
        If Not rs.EOF Then
            provant = rs!CODPR
            provDesc = ObtenerDescripcion("Prov", rs!CODPR)
            While Not rs.EOF
                
                If rs!CODPR <> provant Then
                    AplicoFiltro provant, saldo
                
                    saldo = 0
                    provant = rs!CODPR
                    provDesc = ObtenerDescripcion("Prov", rs!CODPR)
                
                End If
                
                
                numint = numint + 1
                If rs!TIPODOC = "FAC" Or rs!TIPODOC = "N/D" Or rs!TIPODOC = "APD" Then
                    saldo1 = s2n(rs!saldo)
                    saldo = s2n(saldo - saldo1)
                    DataEnvironment1.dbo_SALDOPROVTEMP "A", rs!CODPR, provDesc, rs!Fecha, _
                                                        rs!TIPODOC, rs!NroDoc, IIf(Not IsNull(rs!vencim), rs!vencim, rs!Fecha), 0, _
                                                        saldo1, saldo, numint
                Else
                    'lito metio mano aca s2n()
                    saldo1 = s2n(rs!saldo)
                    saldo = s2n(saldo + saldo1)
                   
                   
                    DataEnvironment1.dbo_SALDOPROVTEMP _
                        "A", rs!CODPR, provDesc, rs!Fecha, _
                        rs!TIPODOC, rs!NroDoc, IIf(Not IsNull(rs!vencim), rs!vencim, rs!Fecha), _
                        saldo1, 0, saldo, numint
                        
                End If
                rs.MoveNext
            Wend
            AplicoFiltro provant, saldo
            
            DataEnvironment1.Sistema.Execute "update SALDOPROVTMP set tipodoc='PAC' where tipodoc='RAC'"
            
            If oPorReporte Then
                DataEnvironment1.LisSaldoPorProveedor
                
                
                rptSaldoPorProveedor.Sections("ReportHeader").Controls("Label31").caption = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
                rptSaldoPorProveedor.Show vbModal
                DataEnvironment1.rsLisSaldoPorProveedor.Close
                DataEnvironment1.dbo_SALDOPROVTEMP "B", 0, "", 0, "", 0, 0, 0, 0, 0, 0
'                PorReporte = False
            Else
                LimpiarGrilla grilla
                LlenarGrilla grilla, "Select CODPR AS CODIGO, RSOCIAL AS 'RAZON SOCIAL', FECHA, TIPODOC AS DOCUMENTO, " & _
                                        "NRODOC AS NUMERO, VENCIM AS 'FECHA VENC.', DEBE, HABER, SALDO, '' as Final " & _
                                     "From SALDOPROVTMP " & _
                                     "Order By CODPR, FECHA, INTERNO", False ', True ,  0
                grillaMarcoSaldosFinales grilla, 0, 9, 8
                
                titux = "Saldo Proveedores " & Date & " " & IIf(optuno, UpROV.codigo & " " & UpROV.DESCRIPCION, "")
                inigrilla titux
                grillaWidth grilla, Array(810, 1995, 1065, 1185, 900, 1230, 1395, 1395, 1395, 1395) 'Array(810, 2000, 1065, 1230, 900, 1230, 1400, 1400, 1400)

'                grilla.MergeCells = flexMergeFree
'                grilla.MergeCol(1) = True
'                grilla.MergeCol(0) = True
'
'                grilla.CellAlignment = flexAlignLeftTop
            End If
        Else
            grilla.rows = 1
            MsgBox "El proveedor no tiene saldo", vbOKOnly, "Atencion"
        End If
        rs.Close
        Set rs = Nothing
    
    Else
        MsgBox "Debe Ingresar una opción", vbOKOnly, "Atencion"
    End If
End Sub



Private Sub AplicoFiltro(prov, saldo)
    If opttodos Then
        If (optFiltro(1) And saldo < 0) Or (optFiltro(2) And saldo > 0) Then
            DataEnvironment1.Sistema.Execute "delete from saldoprovtmp where codpr = " & prov
        End If
    End If
End Sub
Private Sub cmdCancelar_Click()
    LimpioControles
    HabilitoProv (False)
End Sub

Private Sub inigrilla(que As String)
    uXls.ini grilla, "c:\SaldoProv", que
End Sub
'Private Sub cmdexcel_Click()
'    Dim provDesc As String
'    Dim rs As New ADODB.Recordset, provant, saldo, numint As Long
'
'    If opttodos = True Then
'        rs.Open "select * from transcom where activo = 1 order by codpr, fecha", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
'    Else
'        rs.Open "select * from transcom where codpr = " & Val(txtcodprov) & " and activo = 1 order by fecha", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
'    End If
'
'    numint = 0
'    If Not rs.EOF Then
'        provant = rs!CODPR
'        provDesc = ObtenerDescripcion("Prov", rs!CODPR)
'        While Not rs.EOF
'
'            If rs!CODPR <> provant Then
'                saldo = 0
'                provDesc = ObtenerDescripcion("Prov", rs!CODPR)
'            End If
'
'            numint = numint + 1
'            If rs!TIPODOC = "FAC" Or rs!TIPODOC = "N/D" Or rs!TIPODOC = "APD" Then
'                saldo = saldo - rs!Total
'                DataEnvironment1.dbo_SALDOPROVTEMP "A", rs!CODPR, provDesc, rs!Fecha, rs!TIPODOC, rs!NroDoc, rs!vencim, 0, rs!Total, saldo, numint
'            Else
'                saldo = saldo + rs!Total
'                DataEnvironment1.dbo_SALDOPROVTEMP "A", rs!CODPR, provDesc, rs!Fecha, rs!TIPODOC, rs!NroDoc, rs!vencim, rs!Total, 0, saldo, numint
'            End If
'            rs.MoveNext
'        Wend
'    End If
'    rs.Close
'    Set rs = Nothing
'
'    'MANDO A EXCEL
'    rs.Open "select codpr as [Codigo de Proveedor], rsocial, fecha, tipodoc, nrodoc, vencim, debe haber, saldo from Saldoprovtmp", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
'
'    Dim XLcamino
'    XLcamino = "C:\SaldoProv.xls"
'    VinculoXl XLcamino, "Saldo a cuenta de Proveedores", , , rs
'    MsgBox XLcamino
'
'    rs.Close
'    Set rs = Nothing
'
'End Sub



'Private Sub CmdImprimir_Click()
'    PorReporte = True
'End Sub

'Private Sub cmdProv_Click()
'    FrmHelp.Show
'    CargarHelp "Prov", "Codigo", "Descripcion", "codigo", "descripcion"
'    FrmHelp.Tag = "FrmSaldoCuentaProv"
'End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub LimpioControles()
    UpROV.clear
    opttodos = False
    optuno = False
End Sub

Private Sub HabilitoProv(habilito As Boolean)
'    txtcodprov.Visible = habilito
'    txtprov.Visible = habilito
'    Label3.Visible = habilito
'    cmdprov.Visible = habilito

    fraProveedor.Visible = habilito
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub

Private Sub Form_Load()
    HabilitoProv False
    inigrilla "Saldo Proveedores"
    uXls.caption = "Grilla a XLS"
    
    UpROV.ini "select descripcion from prov where activo = 1 and codigo = '###'", "select codigo as [ Codigo ], descripcion as [ Descripcion               ] from prov where activo = 1 order by codigo ", False
    Form_Resize
End Sub

Private Sub Form_Resize()
    Anclar fraBoton, Me, anclarAbajo + anclarIzquierda
    Anclar Frame2, Me, anclarLadosTodos
    Anclar grilla, Me, anclarLadosTodos
End Sub


Private Sub opttodos_Click()
    'HabilitoProv (False)
    revisar
End Sub
Private Sub opttodos_KeyUp(KeyCode As Integer, Shift As Integer)
    revisar
End Sub
Private Sub optuno_Click()
    'HabilitoProv (True)
    revisar
    'txtcodprov.SetFocus
End Sub
Private Sub revisar()
    fraProveedor.Visible = optuno.Value
    fraFiltro.Visible = Not optuno.Value
End Sub
Private Sub optuno_KeyUp(KeyCode As Integer, Shift As Integer)
    revisar
End Sub


'Private Sub txtcodprov_GotFocus()
'    TxtCodProv.SelStart = 0
'    TxtCodProv.SelLength = Len(TxtCodProv.Text)
'End Sub
'
'Private Sub txtcodprov_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    Else
'        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
'            KeyAscii = 0
'        End If
'    End If
'End Sub
'
'Private Sub txtcodprov_LostFocus()
'    If Trim(TxtCodProv) <> "" Then
'        txtprov = ObtenerDescripcion("Prov", Val(TxtCodProv))
'        If txtprov = "" Then
'            MsgBox "Proveedor incorrecto"
'            TxtCodProv.SetFocus
'        End If
''    Else
''        MsgBox "Debe ingresar un proveedor"
''        txtcodprov.SetFocus
'    End If
'End Sub

'Public Sub CargarDatos()
'
'Dim rs As New ADODB.Recordset
'
'    rs.Open "select * from Prov where codigo = " & Val(TxtCodProv) & " and activo = 1 order by codigo", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
'    If Not rs.EOF Then
'        TxtCodProv = rs!codigo
'        txtprov = rs!descripcion
'    End If
'    rs.Close
'    Set rs = Nothing
'
'End Sub

'6/1/5 mando xls , txtbox a donde fue. deberia haber un common dialog


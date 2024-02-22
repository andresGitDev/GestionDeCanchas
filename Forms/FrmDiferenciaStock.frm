VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmDiferenciaStock 
   Caption         =   "Diferencia Stock"
   ClientHeight    =   9945
   ClientLeft      =   135
   ClientTop       =   345
   ClientWidth     =   10050
   Icon            =   "FrmDiferenciaStock.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9945
   ScaleWidth      =   10050
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRemito 
      Height          =   285
      Left            =   6420
      TabIndex        =   31
      Text            =   "0"
      Top             =   750
      Visible         =   0   'False
      Width           =   885
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   255
      Left            =   1440
      TabIndex        =   28
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      Format          =   73269249
      CurrentDate     =   39318
   End
   Begin VB.Frame fraProv 
      Height          =   2910
      Left            =   120
      TabIndex        =   13
      Top             =   5370
      Visible         =   0   'False
      Width           =   9765
      Begin VB.TextBox txtDatos 
         Height          =   300
         Index           =   4
         Left            =   1605
         TabIndex        =   27
         Top             =   2055
         Width           =   4530
      End
      Begin VB.TextBox txtDatos 
         Height          =   300
         Index           =   3
         Left            =   1590
         TabIndex        =   26
         Top             =   1710
         Width           =   4530
      End
      Begin VB.TextBox txtDatos 
         Height          =   300
         Index           =   2
         Left            =   1590
         TabIndex        =   25
         Top             =   1335
         Width           =   4530
      End
      Begin VB.TextBox txtDatos 
         Height          =   300
         Index           =   1
         Left            =   1605
         TabIndex        =   24
         Top             =   975
         Width           =   4530
      End
      Begin Gestion.ucCuit ucCuit1 
         Height          =   315
         Left            =   1620
         TabIndex        =   21
         Top             =   2430
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
      End
      Begin VB.TextBox txtDatos 
         Height          =   300
         Index           =   0
         Left            =   1590
         TabIndex        =   18
         Top             =   585
         Width           =   4530
      End
      Begin Gestion.ucCoDe uProv 
         Height          =   315
         Left            =   1575
         TabIndex        =   22
         Top             =   195
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   556
         CodigoWidth     =   1000
      End
      Begin VB.Label Label3 
         Caption         =   "Proveedor"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   23
         Top             =   210
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "CUIT"
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
         Height          =   375
         Index           =   6
         Left            =   225
         TabIndex        =   20
         Top             =   2445
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "RazonSocial:"
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
         Height          =   375
         Index           =   5
         Left            =   150
         TabIndex        =   19
         Top             =   630
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "Cat IVA"
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
         Height          =   375
         Index           =   4
         Left            =   165
         TabIndex        =   17
         Top             =   2070
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "Provincia"
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
         Height          =   375
         Index           =   3
         Left            =   135
         TabIndex        =   16
         Top             =   1755
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "Localidad"
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
         Height          =   375
         Index           =   2
         Left            =   150
         TabIndex        =   15
         Top             =   1365
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "Direccion:"
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
         Height          =   375
         Index           =   1
         Left            =   150
         TabIndex        =   14
         Top             =   975
         Width           =   1305
      End
   End
   Begin Gestion.ucBotonera uMenu 
      Align           =   2  'Align Bottom
      Height          =   1545
      Left            =   0
      TabIndex        =   0
      Top             =   8400
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   2725
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   ""
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Relacion es:  **MovimientoInterno-Numero**"
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   6255
         TabIndex        =   12
         Top             =   150
         Visible         =   0   'False
         Width           =   3225
      End
   End
   Begin VB.OptionButton optSumaResta 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Resta"
      Height          =   255
      Index           =   1
      Left            =   7980
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
   Begin VB.OptionButton optSumaResta 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Suma"
      Height          =   255
      Index           =   0
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.ComboBox cmbconcepto 
      Height          =   315
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   360
      Width           =   4695
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   4155
      Left            =   255
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1110
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   7329
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      ForeColor       =   -2147483630
      TabCaption(0)   =   "Items"
      TabPicture(0)   =   "FrmDiferenciaStock.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "uProd"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "grillaProductos"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdotro"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtcantidad"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkConver"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Series"
      TabPicture(1)   =   "FrmDiferenciaStock.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "uSeries"
      Tab(1).ControlCount=   1
      Begin VB.CheckBox chkConver 
         Alignment       =   1  'Right Justify
         Caption         =   "Convertir el valor ingresado en la unidad del producto"
         Height          =   225
         Left            =   4455
         TabIndex        =   30
         Top             =   1140
         Width           =   4140
      End
      Begin Gestion.ucSeries uSeries 
         Height          =   2940
         Left            =   -74700
         TabIndex        =   11
         Top             =   600
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   5186
      End
      Begin VB.TextBox txtcantidad 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7395
         TabIndex        =   8
         Top             =   825
         Width           =   1215
      End
      Begin VB.CommandButton cmdotro 
         BackColor       =   &H00E0E0E0&
         Height          =   330
         Left            =   8655
         MaskColor       =   &H00E0E0E0&
         Picture         =   "FrmDiferenciaStock.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   795
         Width           =   525
      End
      Begin VSFlex7LCtl.VSFlexGrid grillaProductos 
         Height          =   2640
         Left            =   120
         TabIndex        =   6
         Top             =   1410
         Width           =   9135
         _cx             =   16113
         _cy             =   4657
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
      Begin Gestion.ucCoDe uProd 
         Height          =   315
         Left            =   180
         TabIndex        =   9
         Top             =   480
         Width           =   8430
         _ExtentX        =   14870
         _ExtentY        =   556
         CodigoWidth     =   1000
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad :"
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
         Index           =   0
         Left            =   6480
         TabIndex        =   10
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Fecha :"
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
      TabIndex        =   29
      Top             =   720
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00800000&
      FillColor       =   &H00400000&
      Height          =   495
      Left            =   6390
      Top             =   225
      Width           =   3015
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   240
      X2              =   8700
      Y1              =   1140
      Y2              =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "Concepto :"
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
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   5220
      Left            =   120
      Top             =   120
      Width           =   9735
   End
End
Attribute VB_Name = "FrmDiferenciaStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   Relacion es R.MovimientoInterno = D.Numero
'   NroComprobante puede ser vacio, estar repetido, se usa solo a veces


Option Explicit ' mod 11/8/4

Private mTipo As Long

Private stblDifSotckTmp As String
Private Const tt_stblDifSotckTmp = " ([CODIGO] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,[DESCRIPCION] [varchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,[CANT] [float] NULL )"
Private g As LiGrilla
Private Enum difGrilla
    dfpProd
    dfpDesc
    dfpCant
End Enum

Private Sub cmbconcepto_LostFocus()
    Dim conc As String
    
    conc = MovimientoConcepto(ObtenerCodigo("conceptos", Trim(cmbconcepto.Text)))
    Select Case conc
    Case "A": optSumaResta(0).Value = True
    Case "R": optSumaResta(1).Value = True
    Case "S": optSumaResta(0).Value = True
    Case "N": optSumaResta(0).Value = True
    End Select
    
'    If MovimientoConcepto(ObtenerCodigo("conceptos", Trim(cmbconcepto.Text))) = "A" Then
'        'chksuma.Value = 0
'        'chkresta.Value = 0
'    Else
'        If MovimientoConcepto(ObtenerCodigo("conceptos", Trim(cmbconcepto.Text))) = "R" Then
'            'chksuma.Value = 0
'            'chkresta.Value = 1
'        Else
'            If MovimientoConcepto(ObtenerCodigo("conceptos", Trim(cmbconcepto.Text))) = "S" Then
'                'chksuma.Value = 1
'                'chkresta.Value = 0
'            End If
'        End If
'    End If
    mTipo = s2n(obtenerDeSQL("select comprobante from conceptos where descripcion = '" & Trim(cmbconcepto) & "' "))
    fraProv.Visible = (mTipo > 0)
    
End Sub

Private Sub difDel()
If ON_ERROR_HABILITADO Then On Error GoTo UfaDif
Dim prod As String, seri As String, conc As Long
Dim i As Long, Remito As Long, COMPROBANTE As Long, nrocomp As Long, essalida As Long, CodProd As String, cant As Double, Ope As String
Dim rsremitos As New ADODB.Recordset
    
If grillaproductos.rows > 1 Then
    Ope = "B"
    essalida = 0
    DE_BeginTrans
    Remito = s2n(txtRemito)
    ABMDifStock "B", DTPicker1.Value, Remito, ComprobanteConcepto(ObtenerCodigo("conceptos", Trim(cmbconcepto.Text))), nrocomp, ObtenerCodigo("conceptos", Trim(cmbconcepto.Text)), 0, UpROV.codigo
    For i = 1 To grillaproductos.rows - 1
        grillaproductos.Row = i
        grillaproductos.Col = 0
        CodProd = grillaproductos.Text
        grillaproductos.Col = 2
        cant = s2n(grillaproductos.Text)
        If ABMDSItem(Ope, CodProd, cant, Remito, 0, 1) = False Then GoTo UfaDif
    Next i
    conc = ComprobanteConcepto(ObtenerCodigo("conceptos", Trim(cmbconcepto.Text)))
    DE_CommitTrans
    MsgBox "La operación fue realizada con éxito.", vbInformation
    'ImprimirDiferenciaStock Remito
    uMenu.AceptarOk
Else
    MsgBox "Debe cargar los productos", vbCritical, "Atencion"
    Exit Sub
End If
Exit Sub
UfaDif:
    DE_RollbackTrans
    MsgBox "Error al eliminar ajuste.", vbCritical
End Sub

Private Sub difAlta()
If ON_ERROR_HABILITADO Then On Error GoTo UfaDif
Dim prod As String, seri As String, conc As Long
Dim i As Long, Remito As Long, COMPROBANTE As Long, nrocomp As Long, essalida As Long, CodProd As String, cant As Double, Ope As String
Dim rsremitos As New ADODB.Recordset
    uSeries.modoSalida = optSumaResta(1).Value
    If uSeries.FaltaSeries() Then
        MsgBox "Faltan numeros de series.", vbExclamation
        tabMain.Tab = 1
        Exit Sub
    End If
    
If grillaproductos.rows > 1 Then
    Remito = nSinNull(obtenerDeSQL("select max(movimientointerno) as maximo from RemitoDiferenciaStock")) + 1
    Select Case ComprobanteConcepto(ObtenerCodigo("conceptos", Trim(cmbconcepto.Text)))
        Case 2
            nrocomp = nSinNull(obtenerDeSQL("select max(nrocomprobante) as nrocomp from RemitoDiferenciaStock where comprobante=2"))
        Case 3
            nrocomp = nSinNull(obtenerDeSQL("select max(nrocomprobante) as nrocomp from RemitoDiferenciaStock where comprobante=3"))
        Case 4
            nrocomp = nSinNull(obtenerDeSQL("select max(nrocomprobante) as nrocomp from RemitoDiferenciaStock where comprobante=4"))
    End Select
    Ope = IIf(optSumaResta(0).Value, "S", "R")
    essalida = IIf(optSumaResta(0).Value, 0, 1)
    DE_BeginTrans
    'DataEnvironment1.dbo_DIFSTOCK DTPicker1.Value, Remito, ComprobanteConcepto(ObtenerCodigo("conceptos", Trim(cmbconcepto.Text))), nrocomp, ObtenerCodigo("conceptos", Trim(cmbconcepto.Text)), 0, uProv.CODIGO
    ABMDifStock "A", DTPicker1.Value, Remito, ComprobanteConcepto(ObtenerCodigo("conceptos", Trim(cmbconcepto.Text))), nrocomp, ObtenerCodigo("conceptos", Trim(cmbconcepto.Text)), 0, UpROV.codigo
    For i = 1 To grillaproductos.rows - 1
        grillaproductos.Row = i
        grillaproductos.Col = 0
        CodProd = grillaproductos.Text
        grillaproductos.Col = 2
        cant = s2n(grillaproductos.Text)
        'DataEnvironment1.dbo_ITEMDIFSTOCK OPE, CodProd, s2n(CANT), Remito, 0
        If ABMDSItem(Ope, CodProd, cant, Remito, 0, grillaproductos.TextMatrix(grillaproductos.Row, 3)) = False Then GoTo UfaDif
    Next i
    conc = ComprobanteConcepto(ObtenerCodigo("conceptos", Trim(cmbconcepto.Text)))
        For i = 1 To uSeries.rows - 1
            prod = Trim(uSeries.cell(i, gsProdu))
            seri = Trim(uSeries.cell(i, gsNser))
            If seri > "" Then
                'DataEnvironment1.dbo_SERIE "A", 0, prod, seri, TipoComprobante_DIFSTOCK, nrocomp, 0, conc, "", 0, Date, UsuarioActual(), 0, 0
                DataEnvironment1.dbo_abmSERIEs "A", 0, prod, seri, TipoComprobante_DIFSTOCK, Remito, 0, conc, "", 0, Date, essalida, Date, UsuarioActual()
            End If
        Next i
    DE_CommitTrans
    MsgBox "La operación fue realizada con éxito.", vbInformation
    ImprimirDiferenciaStock Remito
    uMenu.AceptarOk
Else
    MsgBox "Debe cargar los productos", vbCritical, "Atencion"
    Exit Sub
End If
'    cmdCancelar.Enabled = False
'    cmdaceptar.Enabled = False
'    cmdnuevo.Enabled = True
'    LimpioCampos
Exit Sub
UfaDif:
    DE_RollbackTrans
    MsgBox "Error al grabar ajuste.", vbCritical
End Sub


Private Sub ImprimirDiferenciaStock(NroRemito)
If ON_ERROR_HABILITADO Then On Error GoTo UfaImprime
    
Dim rs As New ADODB.Recordset
Dim sql As String
Dim str, stblDifSotckTmp As String
Dim COMPROBANTE, r As Integer

sql = "select * from remitoDiferenciaStock r inner join conceptos c on r.concepto = c.codigo " & _
      " where r.movimientointerno = " & NroRemito
rs.Open sql, DataEnvironment1.Sistema

If rs.EOF And rs.BOF Then
Else
    COMPROBANTE = rs!COMPROBANTE
    Select Case COMPROBANTE
           Case Is = 0: RptDiferenciaStock.LblTipoRemito = "COMPROBANTE NRO:"
           Case Is = 1: RptDiferenciaStock.LblTipoRemito = "REMITO OFICIAL:"
           Case Is = 2: RptDiferenciaStock.LblTipoRemito = "NOTA ENTREGA:"
           Case Is = 3: RptDiferenciaStock.LblTipoRemito = "NOTA INTERNA:"
    End Select
End If
Set rs = Nothing

stblDifSotckTmp = TablaTempCrear(tt_stblDifSotckTmp)
With grillaproductos
    For r = 1 To .rows - 1
        sql = "insert into " & stblDifSotckTmp _
        & " (codigo,descripcion,cant) values( " & ssTexto(.TextMatrix(r, 0)) & ", " & ssTexto(.TextMatrix(r, 1)) & "," & x2s(.TextMatrix(r, 2)) & ")"
        DataEnvironment1.Sistema.Execute sql
    Next r
End With
RptDiferenciaStock.lblfecha = Date
RptDiferenciaStock.txtconcepto = cmbconcepto.Text
RptDiferenciaStock.Field6 = NroRemito
RptDiferenciaStock.TxtMovInterno = Format(NroRemito, "0001-00000000")
RptDiferenciaStock.txtCliente = txtDatos(0)
RptDiferenciaStock.txtDomicilio = txtDatos(1) & "      " & txtDatos(2)
RptDiferenciaStock.txtIva = txtDatos(4)
RptDiferenciaStock.txtCuit = ucCuit1
str = "select * from " & stblDifSotckTmp
RptDiferenciaStock.Data.Connection = DataEnvironment1.Sistema
RptDiferenciaStock.Data.Source = str
RptDiferenciaStock.Printer.Copies = 2
RptDiferenciaStock.Restart

If PREVIEW_IMPRESIONES Then
    RptDiferenciaStock.Show
Else
    RptDiferenciaStock.PrintReport False
End If

fin:
    Exit Sub
UfaImprime:
    che "fallo la impresion"
    Resume fin
End Sub

Private Sub cmdotro_Click()
    Dim desProd As String, hay As Double, cant As Double, e, iFactor As Double, iCargar As Double


    'If s2n(txtcantidad) <> 0 And Trim$(txtdescripcion) <> "" Then
    cant = s2n(txtCantidad)
    If cant <> 0 And uProd.codigo <> "" Then
        'If IsNumeric(CDbl(Trim(txtcantidad))) Then
            'If s2n(txtcantidad) <> 0 Then
            
                hay = HayStock(uProd.codigo)
                
                If optSumaResta.Item(1) Then ' le agregue la pregunta: Solamente si trata de restar!!!
                    If hay = 0 Then
                        MsgBox "No hay stock disponible", 48, "Atencion"
                    Else
                        If hay < cant Then
                            che "No hay stock suficiente,el stock es de " & hay
                        End If
                    End If
                End If
                If chkConver.Value = 1 Then
                    Set e = Nothing
                    e = obtenerDeSQL("select f.factor from producto p inner join umfactor f on p.ufactor=f.ufcodigo where p.codigo=" & ssTexto(uProd.codigo))
                    If IsNull(e) Or IsEmpty(e) Then
                        iFactor = 0
                    Else
                        iFactor = e
                    End If
                    iCargar = iFactor * cant
                Else
                    iCargar = cant
                End If
                
'                desProd = ObtenerCodigoS("producto", Trim(txtdescripcion))
                
                If grillaproductos.rows = 2 Then
                    grillaproductos.Row = 1
                    grillaproductos.Col = 0
                    
                    If Trim(grillaproductos.Text) = "" Then
                        grillaproductos.Row = 1
                        grillaproductos.Col = 0
                        grillaproductos.Text = uProd.codigo ' desProd
                        grillaproductos.Col = 1
                        grillaproductos.Text = uProd.DESCRIPCION 'txtdescripcion
                        grillaproductos.Col = 2
                        grillaproductos.Text = iCargar
                        grillaproductos.Col = 3
                        grillaproductos.Text = chkConver.Value
                    Else
                        grillaproductos.AddItem uProd.codigo & Chr(9) & uProd.DESCRIPCION & Chr(9) & iCargar & Chr(9) & chkConver.Value
                    End If
                Else
                    grillaproductos.AddItem uProd.codigo & Chr(9) & uProd.DESCRIPCION & Chr(9) & iCargar & Chr(9) & chkConver.Value
                End If
                
                    
            'Else
            '    MsgBox "debe cargar una cantidad > 0", 48, "Atencion"
            '    Exit Sub
            'End If
        'Else
        '    MsgBox "Debe cargar un valor numerico en el campo Cantidad", 48, "Atencion"
        '    txtcantidad.SetFocus
        'End If
    Else
        MsgBox "Debe llenar todos los campos", 48, "Atencion"
    End If

End Sub


Private Sub Form_Load()
    Dim sqlbuscar As String, sqldesc As String
    
    UpROV.ini "Select Descripcion from prov where activo = 1 and codigo = ### ", "Select codigo as [ Proveedor ], descripcion as [ Nombre                                                                   ] from prov where activo = 1 order by codigo", False
    
    
    sqldesc = "select descripcion from producto where codigo = '###' "
    sqlbuscar = "select codigo as [ Codigo                 ], descripcion as [ Descripcion                                                 ] from producto where activo = 1 order by codigo "
    
    'LimpioCampos
    CargaCombo cmbconcepto, "conceptos", "descripcion", "codigo", ""
    InicioGrilla
    
    uProd.ini sqldesc, sqlbuscar, True
    uSeries.ini g, 0, 2, True
    uMenu.init True, True, False, True, True
    DTPicker1.Value = Date
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub


Sub InicioGrilla()
    Set g = New LiGrilla
    g.init grillaproductos
    g.AddCol "CodProducto                      "
    g.AddCol "Descripcion                                                                      "
    g.AddCol "Cantidad         "
    g.AddCol "CONVERTIDO"
    grillaproductos.ColHidden(3) = True

'    grillaProductos.FormatString = "CodProducto                      |<Descripcion                                                                      |>Cantidad         "
'    grillaProductos.rows = 2
'    grillaProductos.cols = 3
'    grillaProductos.Row = 1
'    grillaProductos.Col = 0
'    If grillaProductos.Text <> "" Then
'        grillaProductos.Col = 0
'        grillaProductos.Text = ""
'        grillaProductos.Col = 1
'        grillaProductos.Text = ""
'        grillaProductos.Col = 2
'        grillaProductos.Text = ""
'    End If
End Sub

Sub LimpioCampos()
'    txtdescripcion = ""
    cmbconcepto.ListIndex = -1
'    chksuma.Value = 0
'    chkresta.Value = 0
    'txtcantidad = "0.00"
    FrmBorrarTxt Me
    uProd.clear
    UpROV.clear
    InicioGrilla
    fraProv.Visible = False
End Sub
Sub HabilitoTxt(habilito As Boolean) ' OJO habilito es al reves
'    txtdescripcion.Locked = habilito
    uProd.enabled = Not habilito
    cmbconcepto.Locked = habilito
    txtCantidad.Locked = habilito
'    chksuma.Enabled = Not habilito
'    chkresta.Enabled = Not habilito
    optSumaResta(0).enabled = Not habilito
    optSumaResta(1).enabled = Not habilito
    tabMain.Tab = 0
End Sub


Private Sub tabMain_Click(PreviousTab As Integer)
    If tabMain.Tab = 1 And PreviousTab <> 1 Then
        uSeries.modoSalida = optSumaResta(1).Value ' boton RESTA apretado
        uSeries.FaltaSeries
    End If
End Sub

Private Sub txtcantidad_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtDescripcion_GotFocus()
    PintoFocoActivo
End Sub

Private Sub uMenu_Buscar()
Dim d
d = frmBuscar.MostrarSql("select r.movimientointerno as Ajuste,r.Fecha,r.Concepto, c.Descripcion from remitodiferenciastock r left outer join conceptos c on r.concepto=c.codigo where r.activo=1")
    If d > "" Then
        txtRemito = d
        DTPicker1 = CDate(frmBuscar.resultado(2))
        LlenarGrilla grillaproductos, "Select i.producto as CodProducto,p.Descripcion,i.Cantidad from itemremitodiferenciastock i inner join producto p on i.producto=p.codigo where i.numero=" & d, True
        uMenu.BuscarOK
    Else
        txtRemito = 0
    End If
End Sub

Private Sub uMenu_eliminar()
If MsgBox("¿Esta seguro de eliminar?", vbYesNo) = vbYes Then
    difDel
End If
End Sub

Private Sub uMenu_Imprimir()
    ImprimirDiferenciaStock txtRemito
End Sub

Private Sub uMenu_Nuevo()
InicioGrilla
End Sub

Private Sub uProv_cambio(codigo As Variant)
    Dim tempo, i As Long
    If UpROV.codigo = 0 Then Exit Sub
    
    tempo = obtenerDeSQL("select p.descripcion, direccion, localidad, a.descripcion, i.descripcion, cuit from prov as p inner join ivas as i on tipoiva = i.codigo inner join provincias as a on a.codigo = p.provincia where p.codigo = " & UpROV.codigo)
    If IsEmpty(tempo) Then Exit Sub
    
    For i = 0 To 4
        txtDatos(i) = tempo(i)
    Next i
    ucCuit1 = tempo(5)
End Sub


'*********** Menu *************
Private Sub uMenu_AceptarAlta()
    difAlta
End Sub
Private Sub uMenu_BorrarControles()
    LimpioCampos
End Sub
Private Sub uMenu_HabilitarEdicion(sino As Boolean)
    HabilitoTxt Not sino
End Sub
Private Sub uMenu_SALIR()
    Unload Me
End Sub

Public Function ABMDifStock(sOpe As String, sfecha As Date, sNROINTERNO As Long, sComprobante As Long, sNRO As Long, sCONCEPTO As Long, sNroPedido As Long, sProv As Long) As Boolean
On Error GoTo smal
ABMDifStock = True
Dim iss As String
Select Case sOpe
    Case "A":
        iss = "INSERT INTO REMITODIFERENCIASTOCK (CONCEPTO,MOVIMIENTOINTERNO,FECHA,COMPROBANTE,NROCOMPROBANTE, NroPedido, prov) " _
            & " Values (" & sCONCEPTO & "," & sNROINTERNO & "," & ssFecha(sfecha) & "," & sComprobante & "," & sNRO & "," & sNroPedido & "," & sProv & ")"
        DataEnvironment1.Sistema.Execute iss
    Case "B":
        iss = "UPDATE REMITODIFERENCIASTOCK SET ACTIVO=0 WHERE MOVIMIENTOINTERNO=" & sNROINTERNO
        DataEnvironment1.Sistema.Execute iss
    
End Select
Exit Function
smal:
ABMDifStock = False
End Function

Public Function ABMDSItem(iOPE As String, iCODIGO As String, iCANT As Double, iNROINTERNO As Long, iPRECIO As Double, iConvertido As Long) As Boolean
On Error GoTo imal:
ABMDSItem = True
Dim srf As String, e, pFactor As Double, pCargar As Double
Dim Alma As Integer
    If iConvertido = 0 Then
        If iOPE = "S" Or iOPE = "R" Or iOPE = "B" Then
            Set e = Nothing
            e = obtenerDeSQL("select f.factor from producto p inner join umfactor f on p.ufactor=f.ufcodigo where p.codigo=" & ssTexto(iCODIGO))
            If IsNull(e) Or IsEmpty(e) Then
                pFactor = 0
            Else
                pFactor = e
            End If
            If pFactor = 0 Then
                pFactor = 1
            End If
            pCargar = pFactor * iCANT
        Else
            pCargar = iCANT
        End If
    Else
        pCargar = iCANT
    End If
    
    Alma = s2n(obtenerDeSQL("select almacen from producto where codigo='" & iCODIGO & "'"))
    
    Select Case iOPE
        Case "S": ' -- suma--
            srf = " Update producto " _
                & " SET EXISTENCIA= EXISTENCIA + (" & x2s(pCargar) _
                & ") WHERE CODIGO=" & ssTexto(iCODIGO)
            DataEnvironment1.Sistema.Execute srf
            srf = " INSERT INTO ITEMREMITODIFERENCIASTOCK ( NUMERO, PRODUCTO, CANTIDAD, Precio) " _
                & " Values (" & x2s(iNROINTERNO) & "," & ssTexto(iCODIGO) & "," & x2s(pCargar) & "," & x2s(iPRECIO) & ")"
            DataEnvironment1.Sistema.Execute srf
            
            If Alma <> 0 Then DataEnvironment1.dbo_SumaStock iCODIGO, pCargar, Alma
        Case "R": ' -- resta ---
            srf = " Update producto " _
                & " SET EXISTENCIA= EXISTENCIA - (" & x2s(pCargar) _
                & ") WHERE CODIGO=" & ssTexto(iCODIGO)
            DataEnvironment1.Sistema.Execute srf
            srf = " INSERT INTO ITEMREMITODIFERENCIASTOCK ( NUMERO, PRODUCTO, CANTIDAD, Precio) " _
                & " Values (" & x2s(iNROINTERNO) & "," & ssTexto(iCODIGO) & ", - " & x2s(pCargar) & "," & x2s(iPRECIO) & ")"
            DataEnvironment1.Sistema.Execute srf
            
            If Alma <> 0 Then DataEnvironment1.dbo_SumaStock iCODIGO, -pCargar, Alma
        Case "B": ' -- baja del ajuste ---
            srf = " Update producto " _
                & " SET EXISTENCIA= EXISTENCIA - (" & x2s(pCargar) _
                & ") WHERE CODIGO=" & ssTexto(iCODIGO)
            DataEnvironment1.Sistema.Execute srf
            
            If Alma <> 0 Then DataEnvironment1.dbo_SumaStock iCODIGO, -pCargar, Alma
        Case "F": ' --Formula virtual, sin stock --
            srf = " INSERT INTO ITEMREMITODIFERENCIASTOCK ( NUMERO, PRODUCTO, CANTIDAD, Precio) " _
                & " Values (" & x2s(iNROINTERNO) & "," & ssTexto(iCODIGO) & "," & x2s(pCargar) & "," & x2s(iPRECIO) & ")"
            DataEnvironment1.Sistema.Execute srf
    End Select
Exit Function
imal:
ABMDSItem = False
End Function


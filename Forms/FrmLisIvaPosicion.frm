VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmLisIvaPosicion 
   Caption         =   "Resumen de posicion de Iva"
   ClientHeight    =   10335
   ClientLeft      =   60
   ClientTop       =   -1065
   ClientWidth     =   13005
   Icon            =   "FrmLisIvaPosicion.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10335
   ScaleWidth      =   13005
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdResfresh 
      Caption         =   "Calcular"
      Height          =   360
      Left            =   4125
      TabIndex        =   27
      Top             =   9705
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox cboTipo 
      Height          =   315
      Left            =   7215
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   570
      Visible         =   0   'False
      Width           =   2070
   End
   Begin Gestion.ucXls ucXls2 
      Height          =   840
      Left            =   2160
      TabIndex        =   23
      Top             =   90
      Width           =   840
      _extentx        =   1482
      _extenty        =   1482
   End
   Begin MSComCtl2.DTPicker dtPeriodo 
      Height          =   300
      Left            =   7215
      TabIndex        =   10
      Top             =   120
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "MMMM 'de'  yyyy"
      Format          =   177537027
      CurrentDate     =   39574
   End
   Begin VB.Frame fraResumen 
      Caption         =   "Posicion de Iva"
      Height          =   1980
      Left            =   75
      TabIndex        =   6
      Top             =   8115
      Width           =   4005
      Begin VB.TextBox txtInicial 
         Height          =   285
         Left            =   2115
         TabIndex        =   22
         Text            =   "0"
         Top             =   225
         Width           =   1350
      End
      Begin VB.Label Label12 
         Caption         =   "Saldo Inicial"
         Height          =   210
         Left            =   915
         TabIndex        =   21
         Top             =   270
         Width           =   2100
      End
      Begin VB.Label lblResumen 
         Caption         =   "0"
         Height          =   225
         Left            =   2145
         TabIndex        =   20
         Top             =   1650
         Width           =   1770
      End
      Begin VB.Label lblPercepciones 
         Caption         =   "0"
         Height          =   225
         Left            =   2145
         TabIndex        =   19
         Top             =   1290
         Width           =   1770
      End
      Begin VB.Label lblRetenciones 
         Caption         =   "0"
         Height          =   225
         Left            =   2145
         TabIndex        =   18
         Top             =   1035
         Width           =   1770
      End
      Begin VB.Label lblCompras 
         Caption         =   "0"
         Height          =   225
         Left            =   2145
         TabIndex        =   17
         Top             =   795
         Width           =   1770
      End
      Begin VB.Label lblVentas 
         Caption         =   "0"
         Height          =   225
         Left            =   2145
         TabIndex        =   16
         Top             =   540
         Width           =   1770
      End
      Begin VB.Label Label6 
         Caption         =   "Resultado Total"
         Height          =   285
         Left            =   630
         TabIndex        =   15
         Top             =   1650
         Width           =   2100
      End
      Begin VB.Label Label5 
         Caption         =   "Total IVA Percepciones +"
         Height          =   285
         Left            =   135
         TabIndex        =   14
         Top             =   1350
         Width           =   2100
      End
      Begin VB.Label Label4 
         Caption         =   "Total IVA Retenciones -"
         Height          =   285
         Left            =   135
         TabIndex        =   13
         Top             =   1095
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Total IVA Compras -"
         Height          =   285
         Left            =   135
         TabIndex        =   12
         Top             =   840
         Width           =   1560
      End
      Begin VB.Label Label2 
         Caption         =   "Total IVA Ventas +"
         Height          =   285
         Left            =   135
         TabIndex        =   11
         Top             =   585
         Width           =   1395
      End
   End
   Begin VB.Frame fraVentas 
      Caption         =   "Iva Ventas"
      Height          =   3435
      Left            =   75
      TabIndex        =   5
      Top             =   4635
      Width           =   12900
      Begin VSFlex7LCtl.VSFlexGrid gVentas 
         Height          =   3090
         Left            =   105
         TabIndex        =   8
         Top             =   240
         Width           =   12660
         _cx             =   22331
         _cy             =   5450
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
         Cols            =   1
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmLisIvaPosicion.frx":08CA
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
   Begin VB.Frame fraCompras 
      Caption         =   "Iva Compras"
      Height          =   3435
      Left            =   75
      TabIndex        =   4
      Top             =   1170
      Width           =   12900
      Begin VSFlex7LCtl.VSFlexGrid gCompras 
         Height          =   3090
         Left            =   105
         TabIndex        =   7
         Top             =   240
         Width           =   12660
         _cx             =   22331
         _cy             =   5450
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
         Cols            =   1
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmLisIvaPosicion.frx":08F3
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
   Begin VB.CommandButton cmdMostrar 
      Caption         =   "Mostrar"
      Height          =   825
      Left            =   5100
      Picture         =   "FrmLisIvaPosicion.frx":091C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   90
      Width           =   870
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   825
      Left            =   3705
      Picture         =   "FrmLisIvaPosicion.frx":11E6
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   90
      Width           =   870
   End
   Begin VB.CommandButton cmdimprimir 
      Caption         =   "Imprimir"
      Height          =   825
      Left            =   390
      Picture         =   "FrmLisIvaPosicion.frx":1AB0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   90
      Visible         =   0   'False
      Width           =   870
   End
   Begin Gestion.ucXls ucXls1 
      Height          =   840
      Left            =   1275
      TabIndex        =   0
      Top             =   90
      Width           =   885
      _extentx        =   1561
      _extenty        =   1482
   End
   Begin VB.Label Label8 
      Caption         =   "Doble click sobre el importe limpia"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   4125
      TabIndex        =   26
      Top             =   8265
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.Label Label7 
      Caption         =   "Compras       Ventas"
      Height          =   210
      Left            =   1260
      TabIndex        =   24
      Top             =   960
      Width           =   1740
   End
   Begin VB.Label Label1 
      Caption         =   "Periodo"
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
      Left            =   6375
      TabIndex        =   9
      Top             =   135
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Menu mnuVentas 
      Caption         =   "Ventas"
      Visible         =   0   'False
      Begin VB.Menu mnuAddV01 
         Caption         =   "Agregar a IVA Ventas"
      End
      Begin VB.Menu mnuAddV02 
         Caption         =   "Agregar a IVA Retenciones"
      End
   End
   Begin VB.Menu mnuCompras 
      Caption         =   "Compras"
      Visible         =   0   'False
      Begin VB.Menu mnuAddC01 
         Caption         =   "Agregar a IVA Compras"
      End
      Begin VB.Menu mnuAddC02 
         Caption         =   "Agregar a IVA Percepciones"
      End
   End
End
Attribute VB_Name = "FrmLisIvaPosicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sTablaTemp As String
Private sTmpIvaVentas As String
Private Const tt_iva_compras_temp = "([fecha] [datetime] NULL , [razonsocial] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [nrocuit] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [tipo] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [nro] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,  [neto] [float] NULL ,[ivaresp] [float] NULL , [rg3337] [float] NULL , [iva27] [float] NULL , [iva10] [float] NULL , [retenciva] [float] NULL , [impint] [float] NULL , [retencgan] [float] NULL , [rg3431] [float] NULL , [exento] [float] NULL , [IB_CAPITAL] [float] NULL , [IB_PROVINCIA] [float] NULL    , [imptotal] [float] NULL     )"
Private Const tt_Iva_Ventas_Temp = "([FECHA] [datetime] NULL , [RAZONSOCIAL] [varchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [NROCUIT] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [TIPO] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [NRO] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [NETO] [float] NULL , [NoGrav] [float] NULL , [EXENTO] [float] NULL , [IVARNI] [float] NULL, [IVAEXEN] [float] NULL , [IVARI] [float] NULL , [IVACF] [float] NULL , [IVABC] [float] NULL , [RETIVA] [float] NULL , [iibb] [float] NULL ,[RETIIBB] [float] NULL ,[RETGAN] [float] NULL,[RETSUSS] [float] NULL ,[RETREP] [float] NULL , [IMPTOTAL] [float] NULL )"
Private Const IVA_CONS_FINAL = 1
Private Const IVA_INSCRIPTO = 2
Private Const IVA_NO_INSCRIPTO = 3
Private Const IVA_MONO = 7
Private Const IVA_EXENTO = 4

Private Sub cmdMostrar_Click()
    MostrarCompras
    MostrarVentas
    MostrarResumen
End Sub

Private Sub cmdResfresh_Click()
MostrarResumen
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CDat
    cboTipo.AddItem "A"
    cboTipo.AddItem "B"
    cboTipo.AddItem "C"
    cboTipo.AddItem "G"
    cboTipo.AddItem "TODO"
    cboTipo.ListIndex = 0
    gCompras.Editable = flexEDKbdMouse
    gVentas.Editable = flexEDKbdMouse
End Sub

Private Sub CDat()
    dtPeriodo = Date - 30
    dtPeriodo = "01/" & Month(dtPeriodo) & "/" & Year(dtPeriodo)
    ucXls1.ini gCompras, "C:\Resumen_Posicion_Iva", "Resumen de posicion de Iva"
    ucXls2.ini gVentas, "C:\Resumen_Posicion_Iva", "Resumen de posicion de Iva"
    gCompras.rows = 1
    gVentas.rows = 1
End Sub

Private Sub Form_Resize()
    Anclar fraCompras, Me, anclarLadosAncho
    Anclar gCompras, fraCompras, anclarLadosAncho
    
    Anclar fraVentas, Me, anclarLadosAncho
    Anclar gVentas, fraVentas, anclarLadosAncho
    
    Anclar fraResumen, Me, anclarAbajo
End Sub

Private Sub gCompras_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 Then
    'Me.PopupMenu mnuCompras
End If
End Sub

Private Sub gVentas_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 Then
    'Me.PopupMenu mnuVentas
End If
End Sub


Private Sub mnuAddC01_Click()
Dim Importe As Double, rr As Long, cc As Long
With gCompras
    rr = .Row
    cc = .Col
    If rr <> .rows - 1 Then
        MsgBox "Seleccione el importe del ultimo renglon", vbInformation
        Exit Sub
    End If
    Importe = s2n(.TextMatrix(rr, cc))
    lblCompras = s2n(lblCompras) + Importe
End With
End Sub

Private Sub mnuAddC02_Click()
Dim Importe As Double, rr As Long, cc As Long
With gCompras
    rr = .Row
    cc = .Col
    If rr <> .rows - 1 Then
        MsgBox "Seleccione el importe del ultimo renglon", vbInformation
        Exit Sub
    End If
    Importe = s2n(.TextMatrix(rr, cc))
    lblPercepciones = s2n(lblPercepciones) + Importe
End With
End Sub

Private Sub mnuAddV01_Click()
Dim Importe As Double, rr As Long, cc As Long
With gVentas
    rr = .Row
    cc = .Col
    If rr <> .rows - 1 Then
        MsgBox "Seleccione el importe del ultimo renglon", vbInformation
        Exit Sub
    End If
    Importe = s2n(.TextMatrix(rr, cc))
    lblVentas = s2n(lblVentas) + Importe
End With
End Sub

Private Sub mnuAddV02_Click()
Dim Importe As Double, rr As Long, cc As Long
With gVentas
    rr = .Row
    cc = .Col
    If rr <> .rows - 1 Then
        MsgBox "Seleccione el importe del ultimo renglon", vbInformation
        Exit Sub
    End If
    Importe = s2n(.TextMatrix(rr, cc))
    lblRetenciones = s2n(lblRetenciones) + Importe
End With
End Sub

Private Sub ucXls1_Clic(cancel As Boolean)
    ucXls1.ini gCompras, "C:\Resumen_Iva_Compras_(" & Month(dtPeriodo) & "-" & Year(dtPeriodo) & ").xls", "Resumen de posicion de Iva Compras(" & Month(dtPeriodo) & "-" & Year(dtPeriodo) & ")"
End Sub

Private Function MostrarResumen()
Dim iArrastre
Dim iCompras As Double
Dim iVentas As Double
Dim iRet As Double
Dim iPerc As Double
Dim Aux As Double
iCompras = 0
iVentas = 0
iRet = 0
iPerc = 0
iArrastre = s2n(txtInicial)

If gCompras.rows > 1 Then
    iCompras = gCompras.TextMatrix(gCompras.rows - 1, 7) 'iva 21
    iCompras = iCompras + gCompras.TextMatrix(gCompras.rows - 1, 8) 'iva 105
    iCompras = iCompras + gCompras.TextMatrix(gCompras.rows - 1, 9) 'iva 27
    iCompras = iCompras + gCompras.TextMatrix(gCompras.rows - 1, 10) 'rg 3337
End If
lblCompras = -iCompras

If gVentas.rows > 1 Then
    iVentas = iVentas + gVentas.TextMatrix(gVentas.rows - 1, 8) 'iva exento
    iVentas = iVentas + gVentas.TextMatrix(gVentas.rows - 1, 9) 'iva 21
    iVentas = iVentas + gVentas.TextMatrix(gVentas.rows - 1, 10) 'iva cf
    iVentas = iVentas + gVentas.TextMatrix(gVentas.rows - 1, 11) 'iva bc
    iRet = gVentas.TextMatrix(gVentas.rows - 1, 12) 'ret
End If
lblVentas = iVentas
lblRetenciones = iRet
lblPercepciones = iPerc

lblResumen = s2n(iVentas - (iCompras + Abs(iRet) + iPerc + iArrastre))
relojito False
End Function

Private Function MostrarCompras()
Dim str As String, rs As New ADODB.Recordset, Consulta As String, signo As Variant, ssql As String, scampos As String, dtfechad As Date, dtfechah As Date
Dim Tpo
dtfechad = CDate("01/" & Month(dtPeriodo) & "/" & Year(dtPeriodo))
dtfechah = ultimoDiaDelMes(dtfechad)

FrmLisIvaCompras.CargoGC gCompras, dtfechad, dtfechah, Month(dtfechad), Year(dtfechad), 1

End Function

Private Function MostrarVentas()
Dim rptV As New RptIvaVentas, str As String, rs As New ADODB.Recordset, dtfechad As Date, dtfechah As Date
    dtfechad = CDate("01/" & Month(dtPeriodo) & "/" & Year(dtPeriodo))
    dtfechah = ultimoDiaDelMes(dtfechad)

    FrmLisIvaVentasNew.CargoGV gVentas, dtfechad, dtfechah, 1

End Function

Private Sub ucXls2_Clic(cancel As Boolean)
ucXls2.ini gVentas, "C:\Resumen_Iva_Ventas(" & Month(dtPeriodo) & "-" & Year(dtPeriodo) & ").xls", "Resumen de posicion de Iva Ventas(" & Month(dtPeriodo) & "-" & Year(dtPeriodo) & ")"
End Sub

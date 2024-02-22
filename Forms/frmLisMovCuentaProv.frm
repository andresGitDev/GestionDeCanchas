VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLisMovCuentaProv 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Movimientos de Cuenta de Proveedores"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11205
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7095
   ScaleWidth      =   11205
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraBoton 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   75
      TabIndex        =   16
      Top             =   6540
      Width           =   11175
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
         Height          =   495
         Left            =   10050
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   15
         Width           =   975
      End
      Begin VB.CommandButton cmdaceptar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mostrar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7290
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   15
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
         Height          =   495
         Left            =   8250
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   15
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
         Height          =   300
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   15
         Width           =   975
      End
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
         Height          =   495
         Left            =   1725
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   15
         Width           =   975
      End
      Begin GestionTonka.ucXls ucXls1 
         Height          =   495
         Left            =   2775
         TabIndex        =   17
         Top             =   15
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
      End
   End
   Begin VB.Frame fraGrilla 
      BackColor       =   &H00E0E0E0&
      Height          =   5205
      Left            =   105
      TabIndex        =   15
      Top             =   1305
      Width           =   11055
      Begin VB.Frame fraSubGrilla 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   2430
         Left            =   60
         TabIndex        =   23
         Top             =   2700
         Width           =   10845
         Begin VSFlex7LCtl.VSFlexGrid GrillaDetalle 
            Height          =   1245
            Left            =   855
            TabIndex        =   24
            Top             =   -15
            Width           =   8535
            _cx             =   15055
            _cy             =   2196
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
         Begin VSFlex7LCtl.VSFlexGrid Grillapago 
            Height          =   1125
            Left            =   855
            TabIndex        =   25
            Top             =   1275
            Width           =   8535
            _cx             =   15055
            _cy             =   1984
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Detalle:"
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
            Left            =   150
            TabIndex        =   27
            Top             =   60
            Width           =   690
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Forma de Pago:"
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
            Height          =   720
            Left            =   30
            TabIndex        =   26
            Top             =   1245
            Width           =   825
            WordWrap        =   -1  'True
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid Grilla 
         Height          =   2430
         Left            =   90
         TabIndex        =   8
         Top             =   195
         Width           =   10815
         _cx             =   19076
         _cy             =   4286
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
         Begin VB.Label Label7 
            Caption         =   "DE BAJA ?"
            Height          =   1365
            Left            =   2415
            TabIndex        =   28
            Top             =   465
            Width           =   5400
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fechas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   8730
      TabIndex        =   12
      Top             =   75
      Width           =   2415
      Begin MSComCtl2.DTPicker dtfechad 
         Height          =   375
         Left            =   945
         TabIndex        =   6
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   52822017
         CurrentDate     =   38252
      End
      Begin MSComCtl2.DTPicker dtfechah 
         Height          =   375
         Left            =   930
         TabIndex        =   7
         Top             =   705
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   52822017
         CurrentDate     =   38252
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         Height          =   240
         Left            =   90
         TabIndex        =   14
         Top             =   705
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         Height          =   240
         Left            =   105
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Proveedores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   90
      TabIndex        =   9
      Top             =   75
      Width           =   8535
      Begin VB.TextBox txtCodProvd 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   315
         Width           =   1455
      End
      Begin VB.CommandButton cmdayudaprov 
         BackColor       =   &H00FFFFFF&
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2565
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   315
         Width           =   495
      End
      Begin VB.TextBox txtproveedord 
         Height          =   285
         Left            =   3120
         TabIndex        =   2
         Top             =   315
         Width           =   5295
      End
      Begin VB.TextBox txtcodprovh 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1095
         TabIndex        =   3
         Top             =   705
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2565
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   705
         Width           =   495
      End
      Begin VB.TextBox txtproveedorH 
         Height          =   285
         Left            =   3135
         TabIndex        =   5
         Top             =   705
         Width           =   5295
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
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
         Left            =   270
         TabIndex        =   11
         Top             =   315
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
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
         Left            =   285
         TabIndex        =   10
         Top             =   705
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmLisMovCuentaProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CONST_AJUSTE_PROV_CREDITO = "APC"
Private Const CONST_AJUSTE_PROV_DEBITO = "APD"
Private Const CONST_FACTURA = "FAC"
Private Const CONST_NOTAS_DEBITOS = "N/D"
Private Const CONST_NOTAS_CREDITOS = "N/C"
Private Const CONST_RECIBOS_CUENTA = "RAC" 'ANTES RAC
                                            ' nota 6/4/6:   ?????
Private Const CONST_RECIBOS = "REC"
Private Const CONST_IMPUTACION = "IMP"
Public TablaTemp As String
Private Const CONST_CONTADO = 1


Private Function CalcularSaldoAnterior(CodigoProveedor As Long, fechahasta As Date) As Double
'CONSULTAS DE LAS DISTINTAS TABLAS PARA OBTENER EL SALDO A LA FECHA
'UNA VEZ REALIZADAS LAS CONSULTAS HAY QUE RECORRER LAS TUPLAS Y SUMAR Y/O RESTAR DE ACUERDO AL TIPO DE DOCUMENTO

'SELECT TIPODOC, FORMADEPAGO, sum(total)
'From TRANSCOM
'WHERE ACTIVO = 1 AND CODPR = 1001 AND FECHA <= CONVERT (DATETIME,'11-19-2004')
'GROUP BY TIPODOC, FORMADEPAGO

'SELECT CLIENTE, SUM(TOTAL)
'From RECIBOS
'WHERE ACTIVO = 1 AND CLIENTE = 1000 AND FECHA <= CONVERT (DATETIME,'11-19-2004')
'GROUP BY CLIENTE

'SELECT TIPODOC, sum(total)
'From COMPRAS
'WHERE ACTIVO = 1 AND CODPR = 1001 AND FECHA <= CONVERT (DATETIME,'11-19-2004')
'GROUP BY TIPODOC

Dim debe As Double
Dim haber As Double
Dim rsCuenta As New ADODB.Recordset
Dim Consulta As String

    debe = 0
    haber = 0
    'TABLA TRANSCOM
    Consulta = "Select TIPODOC, Sum(TOTAL) as Total From TRANSCOM Where ACTIVO = 1 And CODPR = " & CodigoProveedor & _
                        " And FECHA < " & ssFecha(fechahasta) & " Group By TIPODOC"
    rsCuenta.Open Consulta, DataEnvironment1.AMR, adOpenDynamic, adLockOptimistic
    If Not rsCuenta.EOF Then rsCuenta.MoveFirst
    While Not rsCuenta.EOF
        If x2s(rsCuenta!TIPODOC) = CONST_FACTURA Or x2s(rsCuenta!TIPODOC) = CONST_NOTAS_DEBITOS Or x2s(rsCuenta!TIPODOC) = CONST_AJUSTE_PROV_DEBITO Then
            haber = haber + s2n(rsCuenta!Total)
        Else
            debe = debe + s2n(rsCuenta!Total)
        End If
        rsCuenta.MoveNext
    Wend
    rsCuenta.Close
'    Set rsCuenta = Nothing
    
    'TABLA REC_COMP,
    Consulta = "Select CODPR, SUM(TOTAL) AS TOTAL, sum(retGanPago) as sumRetGan, sum (ibPago) as sumRetIIBB From REC_COMP Where ACTIVO = 1 And CODPR = " & CodigoProveedor & _
                        " And FECHA < " & ssFecha(fechahasta) & " Group By CODPR"
    rsCuenta.Open Consulta, DataEnvironment1.AMR, adOpenDynamic, adLockOptimistic
    If Not rsCuenta.EOF Then rsCuenta.MoveFirst
    While Not rsCuenta.EOF 'ridiculo, es un sum de 1 prov
        debe = debe + s2n(rsCuenta!Total) + s2n(rsCuenta!sumRetGan) + s2n(rsCuenta!sumRetIIBB)
        rsCuenta.MoveNext
    Wend
    rsCuenta.Close
'    Set rsCuenta = Nothing
    
    'TABLA COMPRAS
    Consulta = "Select TIPODOC,  SUM(TOTAL) AS TOTAL From COMPRAS Where ACTIVO = 1 And CODPR = " & CodigoProveedor & _
                        " And FECHA < " & ssFecha(fechahasta) & " and contado = 0 Group by TIPODOC" ', CONTADO"
    rsCuenta.Open Consulta, DataEnvironment1.AMR, adOpenDynamic, adLockOptimistic
    If Not rsCuenta.EOF Then rsCuenta.MoveFirst
    While Not rsCuenta.EOF
        If x2s(rsCuenta!TIPODOC) = CONST_FACTURA Or x2s(rsCuenta!TIPODOC) = CONST_NOTAS_DEBITOS Or x2s(rsCuenta!TIPODOC) = CONST_AJUSTE_PROV_DEBITO Then
            'pregunto si la forma de pago es contado, porque con esta no hago nada _
            '(ya que debe sumar en el DEBE y restar en el HABER)
            
            
            'Y YO BORRO LA LINEA, porque se suma en el debe y no en el haber ?????
            ' If Not rsCuenta!contado Then haber = haber + s2n(rsCuenta!Total)
            ' eeeehhh ???? asi :
            haber = haber + s2n(rsCuenta!Total)
            '
            
        Else
            debe = debe + s2n(rsCuenta!Total)
        End If
'        Haber = Haber + rsCuenta!Total
        rsCuenta.MoveNext
    Wend
    rsCuenta.Close

    CalcularSaldoAnterior = debe - haber
    
    Set rsCuenta = Nothing
End Function

Private Sub CalcularSaldo()
Dim rsAux As New ADODB.Recordset
Dim Consulta As String
Dim saldo As Variant
Dim CodigoProv As Long
Dim CodigoProvActual As Long

    Consulta = "Select * From " & TablaTemp & " Order By CODIGO_PROV, FECHA, ID"
    rsAux.Open Consulta, DataEnvironment1.AMR, adOpenDynamic, adLockOptimistic
    If Not rsAux.EOF Then rsAux.MoveFirst
    While Not rsAux.EOF
        CodigoProv = rsAux!CODIGO_PROV
        CodigoProvActual = CodigoProv
        saldo = 0
'        Saldo = CalcularSaldoAnterior(CodigoProv, dtfechad.Value)
        While CodigoProv = CodigoProvActual
            If Not IsNull(rsAux!debe) And Not IsNull(rsAux!haber) Then
                
                    saldo = s2n(saldo) + s2n(rsAux!debe) - s2n(rsAux!haber)
                
            End If
            Consulta = "Update " & TablaTemp & " Set SALDO = '" & s2n(saldo, 2) & "' Where ID = " & rsAux!ID
            DataEnvironment1.AMR.Execute Consulta
            rsAux.MoveNext
            If rsAux.EOF Then
                CodigoProvActual = 0
            Else
                CodigoProvActual = rsAux!CODIGO_PROV
            End If
        Wend
    Wend
End Sub

Private Sub CrearConsulta()
Dim Saldo_Cuenta As Double
Dim CodigoProv As Long
Dim DescripcionProv As String
Dim Consulta As String
Dim rs As New ADODB.Recordset
Dim rsProv As New ADODB.Recordset

    'daTaenvironment1.amr.Execute "Delete From LIST_MOV_CUENTA_PROV"
    
    
    rsProv.Open "Select CODIGO, DESCRIPCION From PROV Where CODIGO >= " & s2n(txtCodProvd) & _
                                                    " and CODIGO <= " & s2n(txtcodprovh) & _
                " Order By CODIGO", DataEnvironment1.AMR, adOpenDynamic, adLockOptimistic

'    For CodigoProv = CLng(txtCodProvd) To CLng(txtcodprovh)
    If Not rsProv.EOF Then rsProv.MoveFirst
    While Not rsProv.EOF
        CodigoProv = rsProv!codigo
        DescripcionProv = ssStr(rsProv!descripcion)

'            Saldo_Cuenta = 0
                                
        Saldo_Cuenta = CalcularSaldoAnterior(CodigoProv, dtfechad.Value)
        If Saldo_Cuenta < 0 Then
            Consulta = "Insert Into " & TablaTemp & "(CODIGO_PROV, DESCRIPCION_PROV, FECHA, TIPO_DOCUMENTO, DEBE, HABER, SALDO) " & _
                            "VALUES (" & CodigoProv & ", '" & DescripcionProv & "', " & ssFecha(dtfechad.Value) & _
                                    ", 'SI', '0', '" & Abs(s2n(Saldo_Cuenta, 2)) & "', '" & s2n(Saldo_Cuenta) & "')"
        Else
            Consulta = "Insert Into " & TablaTemp & "(CODIGO_PROV, DESCRIPCION_PROV, FECHA, TIPO_DOCUMENTO, DEBE, HABER, SALDO) " & _
                            "VALUES (" & CodigoProv & ", '" & DescripcionProv & "', " & ssFecha(dtfechad.Value) & _
                                    ", 'SI', '" & s2n(Saldo_Cuenta, 2) & "', '0',  '" & s2n(Saldo_Cuenta, 2) & "')"
        End If
        DataEnvironment1.AMR.Execute Consulta
                            
        Saldo_Cuenta = 0
        'TABLA TRANSCOM
        Consulta = "Select FECHA, TIPODOC, NRODOC, TOTAL, RAZONSOCIALPROV, FORMADEPAGO From TRANSCOM Where ACTIVO = 1 AND " & _
                    "CODPR = " & CodigoProv & " AND FECHA " & ssBetween(dtfechad.Value, dtfechah.Value)
                    
        rs.Open Consulta, DataEnvironment1.AMR, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then rs.MoveFirst
        While Not rs.EOF
            If x2s(rs!TIPODOC) = CONST_FACTURA Or x2s(rs!TIPODOC) = CONST_NOTAS_DEBITOS Or x2s(rs!TIPODOC) = CONST_AJUSTE_PROV_DEBITO Then
'                    Saldo_Cuenta = Saldo_Cuenta + rs!Total
                Consulta = "Insert Into " & TablaTemp & "(CODIGO_PROV, DESCRIPCION_PROV, FECHA, " & _
                                                            "TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO) " & _
                                    "VALUES (" & CodigoProv & ", '" & DescripcionProv & "', " & ssFecha(rs!fecha) & _
                                            ", '" & x2s(rs!TIPODOC) & "', '" & x2s(rs!NroDoc) & "', '0', '" & x2s(rs!Total) & "', '" & Saldo_Cuenta & "')"
            Else
'                    Saldo_Cuenta = Saldo_Cuenta - rs!Total
                Consulta = "Insert Into  " & TablaTemp & "(CODIGO_PROV, DESCRIPCION_PROV, FECHA, " & _
                                                            "TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO) " & _
                                    "VALUES (" & CodigoProv & ", '" & DescripcionProv & "', " & ssFecha(rs!fecha) & _
                                            ", '" & x2s(rs!TIPODOC) & "', '" & x2s(rs!NroDoc) & "', '" & x2s(rs!Total) & "', '0', '" & Saldo_Cuenta & "')"
            End If
            DataEnvironment1.AMR.Execute Consulta
            
'                'SI ES UNA FACTURA CONTADO TAMBIEN LA TENGO QUE PONER EN EL HABER
'                If rs!TIPODOC = CONST_FACTURA And rs!FormadePago = CONST_CONTADO Then
''                    Saldo_Cuenta = Saldo_Cuenta - rs!Total
'                    Consulta = "Insert Into LIST_MOV_CUENTA_PROV (CODIGO_PROV, DESCRIPCION_PROV, FECHA, " & _
'                                                                "TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO) " & _
'                                        "VALUES (" & CodigoProv & ", '" & x2s(rs!razonsocialprov) & "', " & ssFecha(rs!fecha) & _
'                                                ", '" & x2s(rs!TIPODOC) & "', '" & x2s(rs!NroDoc) & "', '" & x2s(rs!Total) & "', '0', '" & Saldo_Cuenta & "')"
'                    daTaenvironment1.amr.Execute Consulta
'                End If
            rs.MoveNext
        Wend
        rs.Close
        Set rs = Nothing
        
        
        'TABLA REC_COMP
        Consulta = "Select * From REC_COMP Where ACTIVO = 1 And " & _
                        "CODPR = " & CodigoProv & " AND FECHA " & ssBetween(dtfechad.Value, dtfechah.Value)
                                                
        rs.Open Consulta, DataEnvironment1.AMR, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then rs.MoveFirst
        While Not rs.EOF
'                Saldo_Cuenta = Saldo_Cuenta - rs!Total
            Consulta = "Insert Into " & TablaTemp & "(CODIGO_PROV, DESCRIPCION_PROV, FECHA, " & _
                    " TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO) " & _
                    " VALUES (" & CodigoProv & ", '" & DescripcionProv & "', " & ssFecha(rs!fecha) & _
                    ", '" & CONST_RECIBOS & "', '" & x2s(rs!Nro) & "', '" & x2s(s2n(rs!Total) + s2n(rs!retganpago)) & "', '0', '" & Saldo_Cuenta & "')"
            DataEnvironment1.AMR.Execute Consulta
            rs.MoveNext
        Wend
        rs.Close
        Set rs = Nothing
        
        'TABLA IMPPRO
        Consulta = "Select * From imppro Where ACTIVO = 1 And " & _
                        "CODPR = " & CodigoProv & " AND FECHA " & ssBetween(dtfechad.Value, dtfechah.Value)
                                                
        rs.Open Consulta, DataEnvironment1.AMR, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then rs.MoveFirst
        While Not rs.EOF
            Consulta = "Insert Into " & TablaTemp & "(CODIGO_PROV, DESCRIPCION_PROV, FECHA, " & _
                                                        "TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO) " & _
                                "VALUES (" & CodigoProv & ", '" & DescripcionProv & "', " & ssFecha(rs!fecha) & _
                                        ", '" & CONST_IMPUTACION & "', '" & x2s(rs!Nro) & "', '0', '0', '" & Saldo_Cuenta & "')"
            DataEnvironment1.AMR.Execute Consulta
            rs.MoveNext
        Wend
        rs.Close
        Set rs = Nothing
        
        'TABLA COMPRAS
        Consulta = "Select CODPR, RAZONSOCIALPROV, FECHA, TIPODOC, NRODOC, TOTAL, CONTADO From COMPRAS Where ACTIVO = 1 And " & _
                    "CODPR = " & CodigoProv & " And FECHA " & ssBetween(dtfechad.Value, dtfechah.Value)
                    
        rs.Open Consulta, DataEnvironment1.AMR, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then rs.MoveFirst
        While Not rs.EOF
        
            If x2s(rs!TIPODOC) = CONST_FACTURA Or x2s(rs!TIPODOC) = CONST_NOTAS_DEBITOS Or x2s(rs!TIPODOC) = CONST_AJUSTE_PROV_DEBITO Then
'                   Saldo_Cuenta = Saldo_Cuenta + rs!Total
                Consulta = "Insert Into " & TablaTemp & "(CODIGO_PROV, DESCRIPCION_PROV, FECHA, " & _
                                                            "TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO) " & _
                                    "VALUES (" & CodigoProv & ", '" & DescripcionProv & "', " & ssFecha(rs!fecha) & _
                                            ", '" & x2s(rs!TIPODOC) & "', '" & x2s(rs!NroDoc) & "', '0', '" & x2s(rs!Total) & "', '" & Saldo_Cuenta & "')"
            Else
                Consulta = "Insert Into " & TablaTemp & "(CODIGO_PROV, DESCRIPCION_PROV, FECHA, " & _
                                                            "TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO) " & _
                                    "VALUES (" & CodigoProv & ", '" & DescripcionProv & "', " & ssFecha(rs!fecha) & _
                                            ", '" & x2s(rs!TIPODOC) & "', '" & x2s(rs!NroDoc) & "', '" & x2s(rs!Total) & "', '0', '" & Saldo_Cuenta & "')"
            End If
            DataEnvironment1.AMR.Execute Consulta
            
            'SI ES UNA FACTURA CONTADO TAMBIEN LA TENGO QUE PONER EN EL HABER
            If rs!TIPODOC = CONST_FACTURA And rs!contado Then
'                    Saldo_Cuenta = Saldo_Cuenta - rs!Total
                Consulta = "Insert Into " & TablaTemp & " (CODIGO_PROV, DESCRIPCION_PROV, FECHA, " & _
                                                            "TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO) " & _
                                    "VALUES (" & CodigoProv & ", '" & DescripcionProv & "', " & ssFecha(rs!fecha) & _
                                            ", 'CON', '" & x2s(rs!NroDoc) & "', '" & x2s(rs!Total) & "', '0', '" & Saldo_Cuenta & "')"
                DataEnvironment1.AMR.Execute Consulta
            End If
            
            
            rs.MoveNext
        Wend
        rs.Close
        Set rs = Nothing
                                
'    Next CodigoProv
        rsProv.MoveNext
    Wend
    
    CalcularSaldo
        
End Sub

Private Sub cmdAceptar_Click()
    If Trim(txtCodProvd.Text) <> "" And Trim(txtcodprovh.Text) <> "" Then
        
        TablaTemp = TablaTempCrear("(" _
        & "[ID] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ," _
        & "[CODIGO_PROV] [numeric](18, 0) NULL ," _
        & "[DESCRIPCION_PROV] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[FECHA] [datetime] NULL ," _
        & "[TIPO_DOCUMENTO] [char] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[NRO_DOCUMENTO] [char] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[DEBE] [char] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[HABER] [char] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[SALDO] [char] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NULL" _
        & ") ON [PRIMARY]")

        DataEnvironment1.AMR.Execute "ALTER TABLE  " & TablaTemp & " WITH NOCHECK ADD" _
        & " CONSTRAINT [PK_" & TablaTemp & "] PRIMARY KEY  CLUSTERED" _
        & "([id])  ON [PRIMARY]"

        DataEnvironment1.AMR.Execute "ALTER TABLE " & TablaTemp & " WITH NOCHECK ADD" _
        & " CONSTRAINT [DF_ " & TablaTemp & "1] DEFAULT (0) FOR [FECHA]," _
        & " CONSTRAINT [DF_ " & TablaTemp & "2] DEFAULT (0) FOR [DEBE]," _
        & " CONSTRAINT [DF_ " & TablaTemp & "3] DEFAULT (0) FOR [HABER]," _
        & "CONSTRAINT [DF_ " & TablaTemp & "4] DEFAULT (0) FOR [SALDO]" _

        DataEnvironment1.AMR.Execute "CREATE  INDEX [IX_ " & TablaTemp & "] ON " & TablaTemp & " ([ID]) ON [PRIMARY]"


        CrearConsulta
        
        LlenarGrilla grilla, "Select CODIGO_PROV AS CODIGO, DESCRIPCION_PROV AS 'RAZON SOCIAL', FECHA, " & _
                                    "TIPO_DOCUMENTO AS DOC, NRO_DOCUMENTO AS NUMERO, DEBE, HABER, SALDO " & _
                             "From " & TablaTemp & _
                             " Order By CODIGO_PROV, FECHA, ID", True
                            
    Else
        MsgBox "Debe seleccionar un proveedor donde comenzar y otro donde terminar", vbOKOnly, "Atencione"
    End If
End Sub

Private Sub cmdexcel_Click()
Dim rs As New ADODB.Recordset
Dim Consulta As String

        
        Consulta = "Select CODIGO_PROV AS CODIGO, DESCRIPCION_PROV AS 'RAZON SOCIAL', FECHA, " & _
                                    "TIPO_DOCUMENTO AS DOC, NRO_DOCUMENTO AS NUMERO, DEBE, HABER, SALDO " & _
                             "From " & TablaTemp & _
                             " Order By CODIGO_PROV, FECHA, ID"
       
        
        rs.Open Consulta, DataEnvironment1.AMR, adOpenDynamic, adLockOptimistic
        VinculoXl "C:\MOVCTACTEPROV.xls", "LISTADO DE MOVIMIENTOS CTA CTE PROVEEDOR", , , rs
        rs.Close
        Set rs = Nothing
End Sub

Private Sub cmdImprimir_Click()
Dim Consulta As String
Dim rsempresa As New ADODB.Recordset

    If Trim(txtCodProvd.Text) <> "" And Trim(txtcodprovh.Text) <> "" Then
    
        Consulta = "Select CODIGO_PROV AS CODIGO, DESCRIPCION_PROV AS 'RAZON SOCIAL', FECHA, " & _
                    "TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO " & _
                    "From " & TablaTemp & _
                    " Order By CODIGO_PROV, FECHA, ID"
                             
        RptLisMovCtaProv.data1.Connection = DataEnvironment1.AMR
        RptLisMovCtaProv.data1.Source = Consulta
        RptLisMovCtaProv.lblfecha = Date
        RptLisMovCtaProv.LBLFECHAD = dtfechad.Value
        RptLisMovCtaProv.LBLFECHAH = dtfechah.Value
        rsempresa.Open "select nombrelogo from datosempresa where nombre='" & FrmPrincipal.lblNombreEmpresa.caption & "'", DataEnvironment1.AMR, adOpenStatic, adLockReadOnly
        'If Not IsNull(rsempresa!nombrelogo) Then
            RptLisMovCtaProv.ImageLOGO.Picture = FrmPrincipal.ImagenTonka 'LoadPicture(App.Path & "\" & rsempresa!nombrelogo)
        'End If
        rsempresa.Close
        Set rsempresa = Nothing
        RptLisMovCtaProv.Show

    Else
        MsgBox "Debe seleccionar un proveedor donde comenzar y otro donde terminar", vbOKOnly, "Atencion"
    End If
End Sub

Private Sub cmdayudaprov_Click()
Dim resu As String

    resu = frmBuscar.MostrarSql("select codigo, descripcion as [Proveedor] from prov where activo = 1")
    If resu > "" Then
        txtCodProvd = frmBuscar.resultado
        txtproveedord = frmBuscar.resultado(2)
        
   End If
End Sub

Private Sub cmdCancelar_Click()
    txtcodprovh = "9999999"
    txtproveedorH = ""
    txtCodProvd = "0"
    txtproveedord = ""
    dtfechad.Value = Date
    dtfechah.Value = Date
    ucXls1.ini grilla, "C:\LisMovCtaProv", "Listado movimiento cuenta proveedores"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
Dim resu As String

    resu = frmBuscar.MostrarSql("select codigo, descripcion as [Proveedor] from prov where activo = 1")
    If resu > "" Then
        txtcodprovh = frmBuscar.resultado
        txtproveedorH = frmBuscar.resultado(2)
        
   End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True, True
End Sub

Private Sub Form_Load()
    cmdCancelar_Click
    Form_Resize
End Sub

Private Sub Form_Resize()
    Anclar fraBoton, Me, anclarAbajo + anclarIzquierda
    Anclar fraGrilla, Me, anclarLadosTodos
    Anclar fraSubGrilla, fraGrilla, anclarAbajo + anclarIzquierda
    Anclar grilla, fraGrilla, anclarLadosTodos
End Sub

Private Sub grilla_Click()
'codigo de proveedor    columna : 0
'fecha                  columna : 2
'tipo documento         columna : 3
'nro documento          columna : 4
Dim TIPODOC As String
Dim NroDoc As Long
Dim CodInt As Long
Dim rsefec As New ADODB.Recordset

    With grilla
        If .TextMatrix(.Row, 0) <> "" And .Row <> 0 Then
            TIPODOC = Trim(.TextMatrix(.Row, 3))
            If .TextMatrix(.Row, 4) <> "" Then NroDoc = CLng(.TextMatrix(.Row, 4))
            
            LimpiarGrilla GrillaDetalle
            LimpiarGrilla Grillapago
            Select Case TIPODOC
                Case CONST_FACTURA
                     '& _INNER JOIN PRODUCTO ON RCD.PRODUCTO=PRODUCTO.CODIGO " &
                    LlenarGrilla GrillaDetalle, "select RCD.CODIGOREMITO AS 'NRO REMITO', RCD.CANTIDAD, RCD.PRODUCTO,PRODUCTO.DESCRIPCION AS 'DESCRIPCION', FCD.PRECIOUNITARIO " & _
                                                " from facturacompraremito as fcd " & _
                                                "inner join REMITOCOMPRADETALLE AS RCD ON RCD.CODIGO=FCD.ITEMREMITOCOMPRA INNER JOIN PRODUCTO ON RCD.PRODUCTO=PRODUCTO.CODIGO" & _
                                                " where fcd.TIPODOC = '" & TIPODOC & "' AND fcd.NRODOC = " & NroDoc, False
                                    
                Case CONST_RECIBOS_CUENTA
                    
                        LlenarGrilla Grillapago, "Select NRO AS 'Nro Cheque',Importe From CHEQUES " & _
                                          "Where ACTIVO = 1 And NDOCprov = " & NroDoc, True
                        If Trim(Grillapago.TextMatrix(1, 0)) = "" Then
                            LlenarGrilla Grillapago, "Select NRO AS 'Nro Cheque', Importe From CHQ_COMP " & _
                                          "Where NRODOC = " & NroDoc & " AND (TIPODOC='" & CONST_RECIBOS_CUENTA & "' AND PROVEEDOR=" & .TextMatrix(.Row, 0) & ")", True
                        Else
                            rsefec.Open "Select nro, Importe From CHQ_COMP " & _
                                          "Where NRODOC = " & NroDoc & " AND (TIPODOC='" & CONST_RECIBOS_CUENTA & "' AND PROVEEDOR=" & .TextMatrix(.Row, 0) & ")", DataEnvironment1.AMR, adOpenDynamic, adLockOptimistic
                            If Not rsefec.EOF Then
                                Do While Not rsefec.EOF
                                    Grillapago.AddItem rsefec!Nro & Chr(9) & rsefec!importe
                                    rsefec.MoveNext
                                Loop
                            End If
                            rsefec.Close
                            Set rsefec = Nothing
                        End If
                        If Trim(Grillapago.TextMatrix(1, 0)) = "" Then
                          LlenarGrilla Grillapago, "Select Fecha, Importe From MOVICAJA " & _
                                              "Where ACTIVO = 1 And TIPODOC = '" & TIPODOC & _
                                                  "' And NRODOC = " & NroDoc & " And TIPO = 'E'", True
                        Else
                            rsefec.Open "Select Fecha,Importe From MOVICAJA " & _
                                        "Where ACTIVO = 1 And TIPODOC = '" & TIPODOC & _
                                        "' And NRODOC = " & NroDoc & " And TIPO = 'E'", DataEnvironment1.AMR, adOpenDynamic, adLockOptimistic
                            If Not rsefec.EOF Then
                                Do While Not rsefec.EOF
                                    Grillapago.AddItem rsefec!fecha & Chr(9) & 0 & Chr(9) & rsefec!importe
                                    rsefec.MoveNext
                                Loop
                            End If
                            rsefec.Close
                            Set rsefec = Nothing
                        End If
                
                Case CONST_RECIBOS
                    If ExisteDato("REC_COMP", "NRO", NroDoc) Then

                        LlenarGrilla GrillaDetalle, "Select tfac AS 'TIPODOC', facT as 'NRODOC', impor as 'IMPORTE' " & _
                                                  "From RELFNR_C " & _
                                                  "Where NDOC = " & NroDoc & " AND (TDOC='" & CONST_RECIBOS & "')", True
            
                        LlenarGrilla Grillapago, "Select NRO AS 'Nro Cheque',Importe From CHEQUES " & _
                                          "Where ACTIVO = 1 And NDOCprov = " & NroDoc, True
                        If Trim(Grillapago.TextMatrix(1, 0)) = "" Then
                            LlenarGrilla Grillapago, "Select NRO AS 'Nro Cheque', Importe From CHQ_COMP " & _
                                          "Where NRODOC = " & NroDoc & " AND (TIPODOC='" & CONST_RECIBOS & "' AND PROVEEDOR=" & .TextMatrix(.Row, 0) & ")", True
                        Else
                            rsefec.Open "Select nro, Importe From CHQ_COMP " & _
                                          "Where NRODOC = " & NroDoc & " AND (TIPODOC='" & CONST_RECIBOS & "' AND PROVEEDOR=" & .TextMatrix(.Row, 0) & ")", DataEnvironment1.AMR, adOpenDynamic, adLockOptimistic
                            If Not rsefec.EOF Then
                                Do While Not rsefec.EOF
                                    Grillapago.AddItem rsefec!Nro & Chr(9) & rsefec!importe
                                    rsefec.MoveNext
                                Loop
                            End If
                            rsefec.Close
                            Set rsefec = Nothing
                        End If
                        If Trim(Grillapago.TextMatrix(1, 0)) = "" Then
                          LlenarGrilla Grillapago, "Select Fecha, Importe From MOVICAJA " & _
                                              "Where ACTIVO = 1 And TIPODOC = '" & TIPODOC & _
                                                  "' And NRODOC = " & NroDoc & " And TIPO = 'E'", True
                        Else
                            rsefec.Open "Select Fecha,Importe From MOVICAJA " & _
                                        "Where ACTIVO = 1 And TIPODOC = '" & TIPODOC & _
                                        "' And NRODOC = " & NroDoc & " And TIPO = 'E'", DataEnvironment1.AMR, adOpenDynamic, adLockOptimistic
                            If Not rsefec.EOF Then
                                Do While Not rsefec.EOF
                                    Grillapago.AddItem rsefec!fecha & Chr(9) & 0 & Chr(9) & rsefec!importe
                                    rsefec.MoveNext
                                Loop
                            End If
                            rsefec.Close
                            Set rsefec = Nothing
                        End If
                    End If
                Case CONST_IMPUTACION
                    LlenarGrilla GrillaDetalle, "Select tfac AS 'TIPODOC', facT as 'NRODOC', impor as 'IMPORTE' " & _
                                                    "From RELFNR_C " & _
                                                    "Where NDOC = " & NroDoc & " AND TDOC='" & CONST_IMPUTACION & "'", True
                
            End Select
        End If
    End With

End Sub

Private Sub txtCodProvd_GotFocus()
    txtCodProvd.SelStart = 0
    txtCodProvd.SelLength = Len(txtCodProvd.Text)
End Sub

Private Sub txtCodProvd_LostFocus()
    If Trim(txtCodProvd) <> "" Then
        txtproveedord = ObtenerDescripcion("Prov", Val(txtCodProvd))
    End If
End Sub

Private Sub txtCodProvh_GotFocus()
    txtcodprovh.SelStart = 0
    txtcodprovh.SelLength = Len(txtcodprovh.Text)
End Sub

Private Sub txtcodprovh_LostFocus()
    If Trim(txtcodprovh) <> "" Then
        txtproveedorH = ObtenerDescripcion("Prov", Val(txtcodprovh))
    End If
End Sub

Private Sub ucXls1_Clic(Cancel As Boolean)
    Dim p As String, fe As String
    fe = " entre " & dtfechad & " y " & dtfechah & ""
    If txtCodProvd = txtcodprovh Then
        p = " para " & txtproveedord
    Else
        p = "prov " & txtproveedord & " a " & txtproveedorH
    End If
    ucXls1.aTitulo = "Listado mov cuenta proveedores " & p & fe
End Sub

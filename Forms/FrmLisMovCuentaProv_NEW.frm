VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLisMovCuentaProv_NEW 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Listado de Movimientos de Cuenta de Proveedores"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11265
   Icon            =   "FrmLisMovCuentaProv_NEW.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   11265
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraBoton 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   75
      TabIndex        =   12
      Top             =   6540
      Width           =   11175
      Begin VB.CommandButton cmdResumir 
         Caption         =   "Resumir"
         Height          =   495
         Left            =   4635
         TabIndex        =   23
         Top             =   15
         Width           =   1440
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
         Height          =   495
         Left            =   10050
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   15
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
         Height          =   495
         Left            =   2115
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   0
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
         Left            =   3075
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   -15
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
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         Width           =   975
      End
      Begin Gestion.ucXls ucXls1 
         Height          =   495
         Left            =   7410
         TabIndex        =   18
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
      TabIndex        =   11
      Top             =   1305
      Width           =   11055
      Begin VB.Frame fraSubGrilla 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   2430
         Left            =   60
         TabIndex        =   15
         Top             =   2700
         Width           =   10845
         Begin VSFlex7LCtl.VSFlexGrid GrillaDetalle 
            Height          =   2175
            Left            =   60
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   210
            Width           =   5115
            _cx             =   9022
            _cy             =   3836
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
            Height          =   2160
            Left            =   5310
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   225
            Width           =   5490
            _cx             =   9684
            _cy             =   3810
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
            Left            =   90
            TabIndex        =   22
            Top             =   -15
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
            Height          =   255
            Left            =   5295
            TabIndex        =   21
            Top             =   0
            Width           =   2115
            WordWrap        =   -1  'True
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid Grilla 
         Height          =   2430
         Left            =   90
         TabIndex        =   4
         TabStop         =   0   'False
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
      TabIndex        =   8
      Top             =   75
      Width           =   2415
      Begin MSComCtl2.DTPicker dtfechad 
         Height          =   375
         Left            =   945
         TabIndex        =   2
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   47841281
         CurrentDate     =   38252
      End
      Begin MSComCtl2.DTPicker dtfechah 
         Height          =   375
         Left            =   930
         TabIndex        =   3
         Top             =   705
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   99155969
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
         TabIndex        =   10
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
         TabIndex        =   9
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
      TabIndex        =   5
      Top             =   75
      Width           =   8535
      Begin Gestion.ucCoDe uProvD 
         Height          =   315
         Left            =   1200
         TabIndex        =   0
         Top             =   315
         Width           =   7080
         _ExtentX        =   12488
         _ExtentY        =   556
         CodigoWidth     =   1000
      End
      Begin Gestion.ucCoDe uProvH 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   720
         Width           =   7080
         _ExtentX        =   12488
         _ExtentY        =   556
         CodigoWidth     =   1000
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   705
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmLisMovCuentaProv_NEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'NOTA  8/6/6
'   La tabla tiene los campos numericos como STRING,
'   asi que con el sql NO hay que transformar el numero
'   al formato yanki (con punto)  " SET SALDO =   12.34  "
'   sino que se trabaja con coma, " SET SALDO =  '12,34' "


Private Const CONST_AJUSTE_PROV_CREDITO = "APC"
Private Const CONST_AJUSTE_PROV_DEBITO = "APD"
Private Const CONST_FACTURA = "FAC"
Private Const CONST_BOLETA = "BOL"
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

    Dim Debe As Double
    Dim haber As Double
    Dim rsCuenta As New ADODB.Recordset
    Dim Consulta As String

    Debe = 0
    haber = 0
    With rsCuenta
        'TABLA TRANSCOM
        Consulta = "Select TIPODOC, Sum(TOTAL) as Total " & _
                " From TRANSCOM " & _
                " Where ACTIVO = 1 And CODPR = " & CodigoProveedor & _
                " And FECHA < " & ssFecha(fechahasta) & _
                " Group By TIPODOC"

        .Open Consulta, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        While Not .EOF
            Select Case x2s(!TIPODOC)
            Case CONST_FACTURA, CONST_NOTAS_DEBITOS, CONST_AJUSTE_PROV_DEBITO
                haber = haber + s2n(!Total, 4)
            Case Else
                Debe = Debe + s2n(!Total, 4)
            End Select
            .MoveNext
        Wend
        .Close

        
        'TABLA REC_COMP  ' grabadas con el nuevo
        Consulta = "Select SUM(TOTAL) AS TOTAL " & _
                    " From REC_COMP " & _
                    " Where ACTIVO = 1 and iddoc >0 " & _
                    " And CODPR = " & CodigoProveedor & _
                    " And FECHA < " & ssFecha(fechahasta)
        Debe = Debe + s2n(obtenerDeSQL(Consulta), 4)
        'TABLA REC_COMP ' grabadas con el viejo
        Consulta = "Select SUM(TOTAL + RetGanPago + IBPago ) AS TOTAL " & _
                    " From REC_COMP " & _
                    " Where ACTIVO = 1 and iddoc < 1 " & _
                    " And CODPR = " & CodigoProveedor & _
                    " And FECHA < " & ssFecha(fechahasta)
        Debe = Debe + s2n(obtenerDeSQL(Consulta), 4)
        
    
        'TABLA COMPRAS
        Consulta = "Select TIPODOC,  SUM(TOTAL) AS TOTAL " & _
                    " From COMPRAS " & _
                    " Where FORMADEPAGO<>-1 AND ACTIVO = 1 And CODPR = " & CodigoProveedor & _
                    " And FECHA < " & ssFecha(fechahasta) & " and contado = 0 " & _
                    " Group by TIPODOC"
                    
        .Open Consulta, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        While Not .EOF
            Select Case x2s(!TIPODOC)
            Case CONST_FACTURA, CONST_NOTAS_DEBITOS, CONST_AJUSTE_PROV_DEBITO
                haber = haber + s2n(!Total, 4)
            Case Else
                Debe = Debe + s2n(!Total, 4)
            End Select
            .MoveNext
        Wend
        .Close
        
        
        'TABLA GASTOSBOLETAS
        Consulta = "Select TIPODOC, Sum(TOTAL) as Total " & _
                " From GastosBoletas " & _
                " Where ACTIVO = 1 And CODPR = " & CodigoProveedor & _
                " And FECHA < " & ssFecha(fechahasta) & _
                " Group By TIPODOC"

        .Open Consulta, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        While Not .EOF
            Select Case x2s(!TIPODOC)
            Case CONST_BOLETA
                haber = haber + s2n(!Total, 4)
            Case Else
                Debe = Debe + s2n(!Total, 4)
            End Select
            .MoveNext
        Wend
        .Close
    
    End With
    CalcularSaldoAnterior = Debe - haber
    
    Set rsCuenta = Nothing
End Function

Private Sub CalcularSaldo()
    Dim rsaux As New ADODB.Recordset
    Dim Consulta As String
    Dim saldo As Variant
    Dim CodigoProv As Long
    Dim CodigoProvActual As Long

    With rsaux
        Consulta = "Select * From " & TablaTemp & " Order By CODIGO_PROV, FECHA, ID"
        
        .Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic

        While Not .EOF
            CodigoProv = !CODIGO_PROV
            CodigoProvActual = CodigoProv
            saldo = 0

            While CodigoProv = CodigoProvActual
                If Not IsNull(!Debe) And Not IsNull(!haber) Then
                        saldo = s2n(saldo) + s2n(rsaux!Debe) - s2n(rsaux!haber)
                End If
                
                Consulta = "Update " & TablaTemp & " Set SALDO = '" & s2n(saldo, 2) & "' Where ID = " & !ID
                DataEnvironment1.Sistema.Execute Consulta
                
                .MoveNext
                If .EOF Then
                    CodigoProvActual = 0
                Else
                    CodigoProvActual = !CODIGO_PROV
                End If
            Wend
            
        Wend
    End With
    Set rsaux = Nothing
End Sub

Private Sub CrearConsulta()
    Dim rsProv As New ADODB.Recordset
    
    With rsProv
        .Open " Select CODIGO, DESCRIPCION " & _
              " From PROV " & _
              " Where CODIGO >= " & uProvD.codigo & _
              "   and CODIGO <= " & uProvH.codigo & _
              " Order By CODIGO", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    
        While Not .EOF
            CrearConsulta_Prov !codigo, ssStr(!DESCRIPCION)
            .MoveNext
        Wend
    
    End With
    CalcularSaldo
    
    Set rsProv = Nothing
End Sub

Private Sub CrearConsulta_Prov(CodigoProv As Long, DescripcionProv As String)
    
    Dim Saldo_Cuenta As Double, ssExtra As String
    Dim Consulta As String
    Dim rs As New ADODB.Recordset
    Dim sDebe As String, sHaber As String
    Dim tot_ConSinRet As Double
    Dim tot As Double
                                
    Saldo_Cuenta = Round(CalcularSaldoAnterior(CodigoProv, dtfechad.Value), 2)
    If Saldo_Cuenta < 0 Then
        sDebe = "'0'"
        sHaber = " '" & Abs(Saldo_Cuenta) & "' "
    Else
        sDebe = " '" & Saldo_Cuenta & "' "
        sHaber = "'0'"
    End If
    
    Consulta = " Insert Into " & TablaTemp & "(CODIGO_PROV, DESCRIPCION_PROV, FECHA, TIPO_DOCUMENTO, DEBE, HABER, SALDO) " & _
               " VALUES (" & CodigoProv & ", '" & DescripcionProv & "', " & ssFecha(dtfechad.Value) & _
                ", 'SI', " & sDebe & ", " & sHaber & ", '" & Saldo_Cuenta & "')"
    DataEnvironment1.Sistema.Execute Consulta
                            
                            
    Saldo_Cuenta = 0
    
    
    With rs
        'TABLA GASTOSBOLETAS
        Consulta = "Select FECHA, TIPODOC, NRODOC, TOTAL, RAZONSOCIAL " & _
                " From GASTOSBOLETAS Where ACTIVO = 1 " & _
                " AND CODPR = " & CodigoProv & " AND FECHA " & ssBetween(dtfechad.Value, dtfechah.Value)
                    
        .Open Consulta, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        
        While Not .EOF
            tot = s2n(!Total) ' tot numerico, lo redondeo
            
            Select Case x2s(!TIPODOC)
                Case CONST_BOLETA
                    sDebe = "'0'"
                    sHaber = " '" & tot & "' "
                Case Else
                    sDebe = " '" & tot & "' "
                    sHaber = "'0'"
            End Select
            
            Consulta = " Insert Into  " & TablaTemp & _
                " (CODIGO_PROV, DESCRIPCION_PROV, FECHA, TIPO_DOCUMENTO, NRO_DOCUMENTO, saldo, DEBE, HABER) " & _
                " VALUES (" & CodigoProv & ", '" & DescripcionProv & "', " & ssFecha(rs!Fecha) & ", " & _
                " '" & ssStr(rs!TIPODOC) & "', '" & (rs!NroDoc) & "', '" & Saldo_Cuenta & "', " & _
                sDebe & ", " & sHaber & ") "
       
            DataEnvironment1.Sistema.Execute Consulta

            .MoveNext
        Wend
        .Close
    End With
    
    With rs
        'TABLA TRANSCOM
        Consulta = "Select FECHA, TIPODOC, NRODOC, TOTAL, RAZONSOCIALPROV, FORMADEPAGO " & _
                " From TRANSCOM Where ACTIVO = 1 " & _
                " AND CODPR = " & CodigoProv & " AND FECHA " & ssBetween(dtfechad.Value, dtfechah.Value)
                    
        .Open Consulta, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        
        While Not .EOF
            tot = s2n(!Total) ' tot numerico, lo redondeo
            
            Select Case x2s(!TIPODOC)
             Case CONST_FACTURA, CONST_NOTAS_DEBITOS, CONST_AJUSTE_PROV_DEBITO
                sDebe = "'0'"
                sHaber = " '" & tot & "' "
            Case Else
                sDebe = " '" & tot & "' "
                sHaber = "'0'"
            End Select
            
            Consulta = " Insert Into  " & TablaTemp & _
                " (CODIGO_PROV, DESCRIPCION_PROV, FECHA, TIPO_DOCUMENTO, NRO_DOCUMENTO, saldo, DEBE, HABER) " & _
                " VALUES (" & CodigoProv & ", '" & DescripcionProv & "', " & ssFecha(rs!Fecha) & ", " & _
                " '" & ssStr(rs!TIPODOC) & "', '" & (rs!NroDoc) & "', '" & Saldo_Cuenta & "', " & _
                sDebe & ", " & sHaber & ") "
       
            DataEnvironment1.Sistema.Execute Consulta
'                'SI ES UNA FACTURA CONTADO TAMBIEN LA TENGO QUE PONER EN EL HABER
'                If rs!TIPODOC = CONST_FACTURA And rs!FormadePago = CONST_CONTADO Then
''                    Saldo_Cuenta = Saldo_Cuenta - rs!Total
'                    Consulta = "Insert Into LIST_MOV_CUENTA_PROV (CODIGO_PROV, DESCRIPCION_PROV, FECHA, " & _
'                                                                "TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO) " & _
'                                        "VALUES (" & CodigoProv & ", '" & x2s(rs!razonsocialprov) & "', " & ssFecha(rs!fecha) & _
'                                                ", '" & x2s(rs!TIPODOC) & "', '" & x2s(rs!NroDoc) & "', '" & x2s(rs!Total) & "', '0', '" & Saldo_Cuenta & "')"
'                    daTaenvironment1.Sistema.Execute Consulta
'                End If
            .MoveNext
        Wend
        .Close
        
        'TABLA REC_COMP
        Consulta = " Select * From REC_COMP Where ACTIVO = 1 " & _
                " And CODPR = " & CodigoProv & " AND FECHA " & ssBetween(dtfechad.Value, dtfechah.Value)
        .Open Consulta, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        
        While Not .EOF
            'sistema viejo (iddoc=0) no sumaba ret al total
            tot_ConSinRet = IIf(nSinNull(!iddoc) > 0, !Total, !Total + nSinNull(!IBPAGO) + nSinNull(!retganpago))
            
            Consulta = "Insert Into " & TablaTemp & "(CODIGO_PROV, DESCRIPCION_PROV, FECHA, " & _
                " TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO) " & _
                " VALUES ( " & CodigoProv & ", '" & DescripcionProv & "', " & ssFecha(rs!Fecha) & _
                ", '" & CONST_RECIBOS & "', '" & !Nro & "', '" & tot_ConSinRet & "', '0', '" & Saldo_Cuenta & "')"
            
            DataEnvironment1.Sistema.Execute Consulta
            .MoveNext
        Wend
        .Close
        
        'TABLA IMPPRO
        Consulta = " Select * From imppro Where ACTIVO = 1 And " & _
                " CODPR = " & CodigoProv & " AND FECHA " & ssBetween(dtfechad.Value, dtfechah.Value)
                                                
        .Open Consulta, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        While Not .EOF
            Consulta = " Insert Into " & TablaTemp & "(CODIGO_PROV, DESCRIPCION_PROV, FECHA, " & _
                " TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO) " & _
                " VALUES (" & CodigoProv & ", '" & DescripcionProv & "', " & ssFecha(rs!Fecha) & _
                ", '" & CONST_IMPUTACION & "', '" & !Nro & "', '0', '0', '" & Saldo_Cuenta & "')"
            DataEnvironment1.Sistema.Execute Consulta
            .MoveNext
        Wend
        .Close
        
        
        'TABLA COMPRAS
        Consulta = "Select CODPR, RAZONSOCIALPROV, FECHA, TIPODOC, NRODOC, TOTAL, CONTADO,FormadePago From COMPRAS Where ACTIVO = 1 And " & _
                    "CODPR = " & CodigoProv & " And FECHA " & ssBetween(dtfechad.Value, dtfechah.Value)
                    
        .Open Consulta, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        While Not .EOF
        
            tot = s2n(!Total)
            
            If s2n(!FormadePago) = -1 Then
                tot = 0
                ssExtra = "-Canjeado"
            Else
                ssExtra = ""
            End If
            
            Select Case x2s(!TIPODOC)
            Case CONST_FACTURA, CONST_NOTAS_DEBITOS, CONST_AJUSTE_PROV_DEBITO
                sDebe = "'0'"
                sHaber = " '" & tot & "' "
            Case Else
                sDebe = " '" & tot & "' "
                sHaber = "'0'"
            End Select

            Consulta = "Insert Into " & TablaTemp & "(CODIGO_PROV, DESCRIPCION_PROV, FECHA, " & _
                " TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO) " & _
                " VALUES (" & CodigoProv & ", '" & DescripcionProv & "', " & ssFecha(rs!Fecha) & _
                ", '" & !TIPODOC & ssExtra & "', '" & !NroDoc & "', " & sDebe & ", " & sHaber & ", '" & Saldo_Cuenta & "')"
            
            DataEnvironment1.Sistema.Execute Consulta
            
            'SI ES UNA FACTURA CONTADO TAMBIEN LA TENGO QUE PONER EN EL HABER
            If !TIPODOC = CONST_FACTURA And !contado Then
                Consulta = "Insert Into " & TablaTemp & " (CODIGO_PROV, DESCRIPCION_PROV, FECHA, " & _
                    " TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO) " & _
                    " VALUES (" & CodigoProv & ", '" & DescripcionProv & "', " & ssFecha(rs!Fecha) & _
                    ", 'CON', '" & !NroDoc & "', '" & tot & "', '0', '" & Saldo_Cuenta & "')"
                DataEnvironment1.Sistema.Execute Consulta
            End If
          
            .MoveNext
        Wend
        .Close
    End With
        
    Set rs = Nothing
    
End Sub

Private Sub cmdAceptar_Click()
    If ON_ERROR_HABILITADO Then On Error GoTo ufaChe
    If uProvD.codigo = 0 Then Exit Sub
    If uProvH.codigo = 0 Then uProvH.codigo = uProvD.codigo
    'If Trim(txtCodProvd.Text) <> "" And Trim(txtcodprovh.Text) <> "" Then
        
    relojito
    
    TablaTemp = TablaTempCrear("(" _
        & "[ID] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ," _
        & "[CODIGO_PROV] [numeric](18, 0) NULL ," _
        & "[DESCRIPCION_PROV] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[FECHA] [datetime] NULL ," _
        & "[TIPO_DOCUMENTO] [varchar] (4000) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[NRO_DOCUMENTO] [char] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[DEBE] [char] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[HABER] [char] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[SALDO] [char] (15) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
        & "[Obs] [char] (10) NULL " _
        & ") ON [PRIMARY]")
    DataEnvironment1.Sistema.Execute "ALTER TABLE  " & TablaTemp & " WITH NOCHECK ADD" _
        & " CONSTRAINT [PK_" & TablaTemp & "] PRIMARY KEY  CLUSTERED" _
        & "([id])  ON [PRIMARY]"
    DataEnvironment1.Sistema.Execute "ALTER TABLE " & TablaTemp & " WITH NOCHECK ADD" _
        & " CONSTRAINT [DF_ " & TablaTemp & "1] DEFAULT (0) FOR [FECHA]," _
        & " CONSTRAINT [DF_ " & TablaTemp & "2] DEFAULT (0) FOR [DEBE]," _
        & " CONSTRAINT [DF_ " & TablaTemp & "3] DEFAULT (0) FOR [HABER]," _
        & "CONSTRAINT [DF_ " & TablaTemp & "4] DEFAULT (0) FOR [SALDO]"
    DataEnvironment1.Sistema.Execute "CREATE  INDEX [IX_ " & TablaTemp & "] ON " & TablaTemp & " ([ID]) ON [PRIMARY]"


    CrearConsulta
    
    'DataEnvironment1.Sistema.Execute "update " & TablaTemp & " set tipo_documento='PAC' where tipo_documento='RAC'"
    'DataEnvironment1.Sistema.Execute "update " & TablaTemp & " set tipo_documento='OP' where tipo_documento='REC'"
    
    LlenarGrilla grilla, _
            " Select CODIGO_PROV AS CODIGO, DESCRIPCION_PROV AS 'RAZON SOCIAL', " & _
            " FECHA, TIPO_DOCUMENTO AS DOC, NRO_DOCUMENTO AS NUMERO, " & _
            " DEBE, HABER, SALDO, '' as [SaldoFinal] " & _
            " From " & TablaTemp & _
            " Order By CODIGO_PROV, FECHA, ID", False
    grillaMarcoSaldosFinales grilla, 0, 8, 7
    
    grillaWidth grilla, Array(740, 2800, 1000, 600, 900, 900, 900, 900, 900)
        
fin:
    relojito False
    Exit Sub
ufaChe:
    che "err en consulta "
    GoTo fin
End Sub


Private Sub cmdImprimir_Click()
    Dim Consulta As String
    Dim rsempresa As New ADODB.Recordset

    If uProvD.codigo = 0 Then Exit Sub
    If uProvH.codigo = "" Then uProvH.codigo = uProvD.codigo
    
    Consulta = "Select CODIGO_PROV AS CODIGO, DESCRIPCION_PROV AS 'RAZON SOCIAL', FECHA, " & _
                    " TIPO_DOCUMENTO, NRO_DOCUMENTO, DEBE, HABER, SALDO " & _
                    " From " & TablaTemp & _
                    " Order By CODIGO_PROV, FECHA, ID"
                             
    RptLisMovCtaProv.data1.Connection = DataEnvironment1.Sistema
    RptLisMovCtaProv.data1.Source = Consulta
    RptLisMovCtaProv.lblfecha = Date
    RptLisMovCtaProv.LBLFECHAD = dtfechad.Value
    RptLisMovCtaProv.LBLFECHAH = dtfechah.Value
    rsempresa.Open "select nombrelogo from datosempresa where nombre='" & FrmPrincipal.lblNombreEmpresa.caption & "'", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    RptLisMovCtaProv.ImageLOGO.Picture = FrmPrincipal.imgLogoSimple 'LoadPicture(App.Path & "\" & rsempresa!nombrelogo)
    
    RptLisMovCtaProv.Show
    
    Set rsempresa = Nothing
End Sub


Private Sub cmdCancelar_Click()
    On Error Resume Next
    uProvH.codigo = obtenerDeSQL("select max(codigo) from prov where activo = 1")
    uProvD.codigo = 1
    dtfechad.Value = CDate("01/01/" & Year(Date))
    dtfechah.Value = Date
    ucXls1.ini grilla, "C:\LisMovCtaProv", "Listado movimiento cuenta proveedores"
    uProvD.SetFocus
End Sub

Private Sub cmdResumir_Click()
Dim i As Long, esto As Long, saldofinal As Double
With grilla
    If .rows > 0 Then
    esto = 1
    saldofinal = 0
devuelta:
        For i = esto To .rows - 1
            If s2n(.TextMatrix(i, .cols - 1)) <> 0 Then
                .TextMatrix(i, 2) = dtfechah
                .TextMatrix(i, 3) = "RESUMEN"
                .TextMatrix(i, 4) = "saldo"
                .TextMatrix(i, 5) = 0
                .TextMatrix(i, 6) = 0
                .TextMatrix(i, 7) = 0
                 saldofinal = saldofinal + s2n(.TextMatrix(i, .cols - 1))
            Else
                esto = i
                .RemoveItem i
                GoTo devuelta
            End If
        Next
        .AddItem ""
        .TextMatrix(.rows - 1, .cols - 1) = s2n(saldofinal)
    End If
End With
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True, True
End Sub

Private Sub Form_Load()
    Dim p1 As String, p2 As String
    p1 = "Select descripcion from prov where codigo = '###' and activo = 1"
    p2 = "select codigo, descripcion as [Nombre                         ] from prov where categ<>2 and activo = 1"
    uProvD.ini p1, p2, False
    uProvH.ini p1, p2, False
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
    ' deberia buscar via idDoc !!!

    'codigo de proveedor    columna : 0
    'fecha                  columna : 2
    'tipo documento         columna : 3
    'nro documento          columna : 4
    
    Dim TIPODOC As String, TipoDocMC As String
    Dim NroDoc As Long
    Dim CodInt As Long
    Dim rs As New ADODB.Recordset
    Dim prov As Long

    With grilla
        If .TextMatrix(.Row, 0) <> "" And .Row <> 0 Then
            TIPODOC = Trim(.TextMatrix(.Row, 3))
            If (TIPODOC) = "OP" Then TIPODOC = "REC"
            If (TIPODOC) = "PAC" Then TIPODOC = "RAC"
            TipoDocMC = IIf(TIPODOC = "REC" Or TIPODOC = "RAC", "O/P", TIPODOC)
            prov = s2n(.TextMatrix(.Row, 0))
            
            If Trim(.TextMatrix(.Row, 4)) <> "saldo" Then
                If .TextMatrix(.Row, 4) <> "" Then NroDoc = CLng(.TextMatrix(.Row, 4))
            End If
            
            LimpiarGrilla GrillaDetalle
            LimpiarGrilla Grillapago, 1
            Select Case TIPODOC
                Case CONST_FACTURA
                     '& _INNER JOIN PRODUCTO ON RCD.PRODUCTO=PRODUCTO.CODIGO " &
                    LlenarGrilla GrillaDetalle, "select RCD.CODIGOREMITO AS 'NRO REMITO', RCD.CANTIDAD, RCD.PRODUCTO,PRODUCTO.DESCRIPCION AS 'DESCRIPCION', FCD.PRECIOUNITARIO " & _
                                                " from facturacompraremito as fcd " & _
                                                " inner join REMITOCOMPRADETALLE AS RCD ON RCD.CODIGO=FCD.ITEMREMITOCOMPRA INNER JOIN PRODUCTO ON RCD.PRODUCTO=PRODUCTO.CODIGO" & _
                                                " where fcd.TIPODOC = '" & TIPODOC & "' AND fcd.NRODOC = " & NroDoc, False
                                    
                Case CONST_RECIBOS_CUENTA
                        'en cheques busca por Nro de OP
                        LlenarGrilla Grillapago, _
                            " Select NRO AS 'Nro Cheque',Importe From CHEQUES " & _
                            " Where ACTIVO = 1 And NDOCprov = " & NroDoc, True
                            
                        ' chq_comp busca por Nro/tipo/prov
                        rs.Open _
                            " Select nro, Importe From CHQ_COMP " & _
                            " Where NRODOC = " & NroDoc & " AND TIPODOC='" & CONST_RECIBOS_CUENTA & "' AND PROVEEDOR=" & prov _
                            , DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                            While Not rs.EOF
                                AgregarGrillaPAgo rs!Nro, rs!Importe
                                rs.MoveNext
                            Wend
                        rs.Close

                        
                        'efectivo
                        rs.Open _
                            " Select Fecha,Importe From MOVICAJA " & _
                            " Where ACTIVO = 1 And TIPODOC = '" & TipoDocMC & _
                            "' And NRODOC = " & NroDoc & " And TIPO = 'E' and codprov = " & prov _
                            , DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                            While Not rs.EOF
                                AgregarGrillaPAgo "Efectivo: ", rs!Importe
                                rs.MoveNext
                            Wend
                        rs.Close

                
                Case CONST_RECIBOS
                    If ExisteDato("REC_COMP", "NRO", NroDoc) Then

                        LlenarGrilla GrillaDetalle, _
                            " Select tfac AS 'TIPODOC', facT as 'NRODOC', impor as 'IMPORTE' " & _
                            " From RELFNR_C " & _
                            " Where NDOC = " & NroDoc & " AND (TDOC='" & CONST_RECIBOS & "') and prov = " & prov, True
            
                        'cheques
                        LlenarGrilla Grillapago, _
                            " Select NRO AS 'Nro Cheque',Importe From CHEQUES " & _
                            " Where ACTIVO = 1 And NDOCprov = " & NroDoc, False
                            
                        'chqcomp
                        rs.Open _
                            " Select nro, Importe From CHQ_COMP " & _
                            " Where NRODOC = " & NroDoc & " AND (TIPODOC='" & CONST_RECIBOS & "' AND PROVEEDOR=" & prov & " )" _
                            , DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly

                            While Not rs.EOF
                                AgregarGrillaPAgo rs!Nro, rs!Importe
                                rs.MoveNext
                            Wend
                        rs.Close
                        
                        'efectivo
                        rs.Open _
                            " Select Fecha,Importe From MOVICAJA " & _
                            " Where ACTIVO = 1 And TIPODOC = '" & TipoDocMC & _
                            "' And NRODOC = " & NroDoc & " And TIPO = 'E' and codprov = " & prov _
                            , DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly

                            While Not rs.EOF
                                AgregarGrillaPAgo "Efectivo:", rs!Importe
                                rs.MoveNext
                            Wend
                        rs.Close

                        
                    End If
                Case CONST_IMPUTACION
                    LlenarGrilla GrillaDetalle, _
                        " Select tfac AS 'TIPODOC', facT as 'NRODOC', impor as 'IMPORTE' " & _
                        " From RELFNR_C " & _
                        " Where NDOC = " & NroDoc & " AND TDOC='" & CONST_IMPUTACION & "' and prov = " & prov, True
            End Select
        End If
    End With
    grillaWidth Grillapago, Array(1600, 1600)
    grillaWidth GrillaDetalle, Array(1600, 1600, 1600)
    Set rs = Nothing
End Sub

Private Sub ucXls1_Clic(cancel As Boolean)
    Dim p As String, fe As String
    fe = " entre " & dtfechad & " y " & dtfechah & ""
    If uProvD.codigo = uProvH.codigo Then
        p = " para " & uProvD.DESCRIPCION
    Else
        p = "prov " & uProvD.DESCRIPCION & " a " & uProvH.DESCRIPCION
    End If
    ucXls1.aTitulo = "Listado mov cuenta proveedores " & p & fe
End Sub

Private Sub AgregarGrillaPAgo(que, cuanto)
    Grillapago.AddItem que & Chr(9) & cuanto
End Sub

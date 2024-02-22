VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmLisRecibos 
   Caption         =   "Listado de Recibos Emitidos"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11190
   Icon            =   "FrmLisRecibos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6765
   ScaleWidth      =   11190
   StartUpPosition =   2  'CenterScreen
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
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6240
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
      Height          =   495
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6240
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fechas "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11175
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
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin MSComCtl2.DTPicker dtfechad 
         Height          =   375
         Left            =   1320
         TabIndex        =   11
         Top             =   330
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   284229633
         CurrentDate     =   38252
      End
      Begin MSComCtl2.DTPicker dtfechah 
         Height          =   375
         Left            =   3465
         TabIndex        =   12
         Top             =   330
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   284229633
         CurrentDate     =   38252
      End
      Begin VB.Label lblTotal 
         BackStyle       =   0  'Transparent
         Height          =   315
         Left            =   5550
         TabIndex        =   18
         Top             =   390
         Width           =   3705
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
         Left            =   600
         TabIndex        =   14
         Top             =   360
         Width           =   615
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
         Left            =   2865
         TabIndex        =   13
         Top             =   390
         Width           =   540
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   5295
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   11175
      Begin VSFlex7LCtl.VSFlexGrid Grilla 
         Height          =   2775
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Haga Click para ver el Detalle de la Orden de Compra"
         Top             =   240
         Width           =   10935
         _cx             =   19288
         _cy             =   4895
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
      Begin VSFlex7LCtl.VSFlexGrid GrillaDetalle 
         Height          =   1935
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Haga Doble Click para ver el Detalle de la Orden de Compra"
         Top             =   3240
         Width           =   5415
         _cx             =   9551
         _cy             =   3413
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
      Begin VSFlex7LCtl.VSFlexGrid GrillaMoviCaja 
         Height          =   855
         Left            =   5640
         TabIndex        =   5
         ToolTipText     =   "Haga Doble Click para ver el Detalle de la Orden de Compra"
         Top             =   3240
         Width           =   5415
         _cx             =   9551
         _cy             =   1508
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
      Begin VSFlex7LCtl.VSFlexGrid GrillaEfectivo 
         Height          =   855
         Left            =   5640
         TabIndex        =   6
         ToolTipText     =   "Haga Doble Click para ver el Detalle de la Orden de Compra"
         Top             =   4320
         Width           =   5415
         _cx             =   9551
         _cy             =   1508
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
         Caption         =   "Comprobantes que imputa:"
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
         TabIndex        =   9
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Movimiento de Caja"
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
         Left            =   5640
         TabIndex        =   8
         Top             =   3000
         Width           =   1785
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Movimiento en Efectivo"
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
         Left            =   5640
         TabIndex        =   7
         Top             =   4080
         Width           =   2070
      End
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
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6240
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
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6240
      Width           =   1455
   End
End
Attribute VB_Name = "FrmLisRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CONST_RECIBOS = "RAA"
Private Const CONST_RECIBOS_IMPUTADOS = "REC"
Public TablaTemp As String

'Private Const CONST_CONTADO = True

Private Sub CrearConsulta()
    Dim Consulta As String
    Dim rs As New ADODB.Recordset
    Dim rsFac As New ADODB.Recordset
    Dim rsCHQ As New ADODB.Recordset
    Dim Total As Double

        Consulta = "Select DISTINCT F.CODIGO, F.FECHA, F.TIPODOC, F.NROFACTURA, F.TOTAL, F.REMITO, F.IVA, F.RAZONSOCIAL, F.FORMAPAGO, F.CONTADO "
        Consulta = Consulta & "From FACTURAVENTA as F "
        Consulta = Consulta & "Where ACTIVO = 1 and F.TIPODOC='RAA' And F.FECHA " & ssBetween(dtfechad.Value, dtfechah.Value)
        Consulta = Consulta & "Order By F.FECHA, F.CODIGO"
        
        rs.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then rs.MoveFirst
        
        While Not rs.EOF
'            Consulta = "Insert Into lis_recibos_temp (CODIGO_CLI, DESCRIPCION_CLI, FECHA, " & _
'                        "NRO_DOCUMENTO, IMPORTE) " & _
'                        "VALUES (" & ObtenerCodigo("clientes", x2s(rs!RAZONSOCIAL)) & ", '" & x2s(rs!RAZONSOCIAL) & "', " & ssFecha(rs!fecha) & _
'                        ",'" & x2s(rs!nrofactura) & "','" & s2n(rs!Total, 2) & "',')"
                
            Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, " & _
                        "TIPODOC,NRO_DOCUMENTO, IMPORTE) " & _
                        "VALUES (" & ObtenerCodigo("clientes", x2s(rs!RAZONSOCIAL)) & ", '" & x2s(rs!RAZONSOCIAL) & "', " & ssFecha(rs!fecha) & _
                        ",'" & rs!TIPODOC & "','" & x2s(rs!nrofactura) & "'," & Replace(s2n(rs!Total, 2), ",", ".") & ")"
            DataEnvironment1.Sistema.Execute Consulta
            Total = Total + rs!Total
                
            rs.MoveNext
        Wend
        rs.Close
        Set rs = Nothing
        
        Consulta = "Select R.* "
        Consulta = Consulta & "From RECIBOS AS R "
        Consulta = Consulta & "Where R.ACTIVO = 1 AND TIPODOC='REC' AND R.FECHA " & ssBetween(dtfechad.Value, dtfechah.Value)
        Consulta = Consulta & "Order By R.FECHA, R.CODIGO"
        
        rs.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
                If Not rs.EOF Then rs.MoveFirst
        While Not rs.EOF
'            Consulta = "Insert Into Lis_recibos_temp(CODIGO_CLI, DESCRIPCION_CLI, FECHA, " & _
'                                                        "NRO_DOCUMENTO,IMPORTE) " & _
'                                "VALUES (" & rs!cliente & ", '" & ObtenerDescripcion("Clientes", rs!cliente) & "', " & ssFecha(rs!fecha) & _
'                                        ", '" & x2s(rs!numero) & "','" & s2n(rs!Total, 2) & "')"
            Consulta = "Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, " & _
                                                        "TIPODOC,NRO_DOCUMENTO,IMPORTE) " & _
                                "VALUES (" & rs!cliente & ", '" & ObtenerDescripcion("Clientes", rs!cliente) & "', " & ssFecha(rs!fecha) & _
                                        ", '" & rs!TIPODOC & "','" & x2s(rs!numero) & "'," & x2s(rs!Total) & ")"
            DataEnvironment1.Sistema.Execute Consulta
            Total = Total + rs!Total
            
            Consulta = "Select FACTURAVENTA, IMPORTE From RECIBOSDETALLE " & _
                        "Where CODRECIBO = " & ObtenerDatoDB("RECIBOS", "NUMERO", rs!numero, "CODIGO") & " Order By CODIGO"
            rsFac.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            Consulta = "Select NRO,CLIENTE,IMPORTE From CHEQUES Where ACTIVO = 1 And TDOC = '" & CONST_RECIBOS & "' And NDOC = " & s2n(rs!numero)
            rsCHQ.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            While Not rsFac.EOF Or Not rsCHQ.EOF
'                Consulta = "Insert Into Lis_recibos_temp(CODIGO_CLI, DESCRIPCION_CLI, FECHA, " & _
'                                                "NRO_DOCUMENTO,IMPORTE,FACTURAS, CHEQUES) " & _
'                                "VALUES (" & rs!cliente & ", '" & ObtenerDescripcion("Clientes", rs!cliente) & "', " & ssFecha(rs!fecha) & _
'                                        ",'', '',"
            
                Consulta = " Insert Into " & TablaTemp & "(CODIGO_CLI, DESCRIPCION_CLI, FECHA, " & _
                            " TIPODOC,NRO_DOCUMENTO,IMPORTE,FACTURAS, CHEQUES) " & _
                            " VALUES (" & rs!cliente & ", '" & ObtenerDescripcion("Clientes", rs!cliente) & "', " & ssFecha(rs!fecha) & _
                                        ",'', '','',"
                If Not rsFac.EOF Then
                    Consulta = Consulta & "'FAC NRO " & x2s(rsFac!FACTURAVENTA) & " - " & x2s(rsFac!importe) & "',"
                    rsFac.MoveNext
                Else
                    Consulta = Consulta & "'',"
                End If
                
                If Not rsCHQ.EOF Then
                    Consulta = Consulta & "'CHQ NRO " & x2s(rsCHQ!Nro) & " - " & x2s(rsCHQ!importe) & "')"
                    rsCHQ.MoveNext
                Else
                    Consulta = Consulta & "'')"
                End If
                
                DataEnvironment1.Sistema.Execute Consulta
                'total = total + rs!total ? creo que no
            Wend
            rsFac.Close
            Set rsFac = Nothing
            rsCHQ.Close
            Set rsCHQ = Nothing
            rs.MoveNext
        Wend
        rs.Close
        Set rs = Nothing
    lbltotal = "Total " & Total
End Sub

Private Sub cmdAceptar_Click()

    relojito True

    TablaTemp = TablaTempCrear("([ID] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ," _
                & "[CODIGO_CLI] [numeric](18, 0) NULL ," _
                & "[DESCRIPCION_CLI] [varchar] (60) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
                & "[FECHA] [datetime] NULL ," _
                & "[TIPODOC] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
                & "[NRO_DOCUMENTO] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
                & "[IMPORTE] [float] NULL ," _
                & "[FACTURAS] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ," _
                & "[CHEQUES] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL" _
                & ") ON [PRIMARY]")
                
    DataEnvironment1.Sistema.Execute "ALTER TABLE " & TablaTemp & " WITH NOCHECK ADD" _
    & " CONSTRAINT [DF_" & TablaTemp & "] DEFAULT ('') FOR [FACTURAS]," _
    & " CONSTRAINT [DF_" & TablaTemp & "1] DEFAULT ('') FOR [CHEQUES]"

    
        CrearConsulta
        LimpiarGrilla grilla
        LimpiarGrilla GrillaDetalle
        LimpiarGrilla GrillaEfectivo
        LimpiarGrilla GrillaMoviCaja
        
        LlenarGrilla grilla, _
            " Select CODIGO_CLI AS CODIGO,L.DESCRIPCION_CLI AS DESCRIPCION, L.FECHA," & _
            " L.TIPODOC,L.NRO_DOCUMENTO,L.IMPORTE" & _
            " From " & TablaTemp & " AS L " & _
            " Where l.facturas = '' and l.cheques = '' " & _
            " Order By  FECHA, ID", True
            
        relojito False

End Sub

Private Sub cmdexcel_Click()
    Dim rs As New ADODB.Recordset
    Dim Consulta As String

        'CrearConsulta False

        If MsgBox("¿Desea incluir los detalles de los comprobantes?", vbYesNo, "Atencion") = vbYes Then
            Consulta = "Select CODIGO_CLI as 'CLIENTE', DESCRIPCION_CLI as 'RAZON SOCIAL', FECHA, " & _
                " TIPODOC AS 'DOCUMENTO', NRO_DOCUMENTO as 'NUMERO DOC.', FACTURAS, CHEQUES, " & _
                " IMPORTE " & _
                " From " & TablaTemp & _
                " Order By  FECHA, ID"
                                
        Else
            Consulta = "Select CODIGO_CLI as 'CLIENTE', DESCRIPCION_CLI as 'RAZON SOCIAL', FECHA, " & _
                                "TIPODOC AS 'DOCUMENTO', NRO_DOCUMENTO as 'NUMERO DOC.', IMPORTE " & _
                        "From " & TablaTemp & _
                        " Where TIPODOC <> '' " & _
                        "Order By  FECHA, ID"
        End If
        
        rs.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        VinculoXl "C:\RECIBOS.xls", "LISTADO DE RECIBOS", , , rs
        rs.Close
        Set rs = Nothing
    End Sub

Private Sub cmdImprimir_Click()
Dim Consulta As String
    
        If MsgBox("¿Desea incluir los detalles de los comprobantes?", vbYesNo, "Atencion") = vbYes Then
            Consulta = "Select CODIGO_CLI as 'CLIENTE', DESCRIPCION_CLI as 'RAZONSOCIAL', FECHA, " & _
                                "TIPODOC AS 'DOCUMENTO', NRO_DOCUMENTO as 'NUMERODOC', FACTURAS, " & _
                                "IMPORTE " & _
                        "From " & TablaTemp & _
                        " Order By  FECHA, ID"
        Else
            Consulta = "Select CODIGO_CLI as 'CLIENTE', DESCRIPCION_CLI as 'RAZONSOCIAL', FECHA, " & _
                                "TIPODOC AS 'DOCUMENTO', NRO_DOCUMENTO as 'NUMERODOC', IMPORTE " & _
                        "From " & TablaTemp & _
                        " Where TIPODOC <> '' " & _
                        "Order By  FECHA, ID"
        End If
            
        RptLisRecibos.data1.Connection = DataEnvironment1.Sistema
        RptLisRecibos.data1.Source = Consulta
        RptLisRecibos.lblFecha = Date
        RptLisRecibos.LBLFECHAD = dtfechad.Value
        RptLisRecibos.LBLFECHAH = dtfechah.Value
        RptLisRecibos.Show
End Sub

Private Sub cmdCancelar_Click()
    dtfechad.Value = Date
    dtfechah.Value = Date
    lbltotal = ""
    grilla.clear
    GrillaDetalle.clear
    GrillaMoviCaja.clear
    GrillaEfectivo.clear
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub

Private Sub Form_Load()
    cmdCancelar_Click
End Sub

Private Sub grilla_Click()
Dim TIPODOC As String
Dim NroDoc As Long
Dim CodInt As Long

    With grilla
        If .TextMatrix(.Row, 0) <> "" And .Row <> 0 Then
            If .TextMatrix(.Row, 4) <> "" Then NroDoc = CLng(.TextMatrix(.Row, 4))
            TIPODOC = .TextMatrix(.Row, 3)
            LimpiarGrilla GrillaDetalle
            LimpiarGrilla GrillaMoviCaja
            LimpiarGrilla GrillaEfectivo
            
            

          CodInt = ObtenerDatoDB("RECIBOS", "NUMERO", NroDoc, "CODIGO")
          LlenarGrilla GrillaDetalle, "Select F.TIPODOC AS 'Tipo', F.NROFACTURA AS 'Numero', Importe " & _
                                      "From RECIBOSDETALLE AS R " & _
                                          "Inner Join FACTURAVENTA AS F on F.CODIGO = R.FACTURAVENTA " & _
                                      "Where R.CODRECIBO = " & CodInt & " Order By R.CODIGO", True

          LlenarGrilla GrillaMoviCaja, "Select Fecha, nro as 'Numero Cheque', Importe From CHEQUES " & _
                              "Where ACTIVO = 1 And TDOC = '" & TIPODOC & "' AND NDOC = " & NroDoc, True
          LlenarGrilla GrillaEfectivo, "Select Fecha, Importe From MOVICAJA " & _
                              "Where ACTIVO = 1 And TIPODOC = '" & TIPODOC & _
                                  "' And NRODOC = " & NroDoc & " And TIPO = 'E'", True
                      
        End If
    End With
End Sub



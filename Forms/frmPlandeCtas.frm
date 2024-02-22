VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmPlandeCtas 
   Caption         =   "Plan de Ctas"
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10740
   Icon            =   "frmPlandeCtas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9060
   ScaleWidth      =   10740
   StartUpPosition =   2  'CenterScreen
   Begin Gestion.ucXls ucXls1 
      Height          =   900
      Left            =   195
      TabIndex        =   3
      Top             =   8070
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1588
   End
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Height          =   7815
      Left            =   120
      TabIndex        =   2
      Top             =   90
      Width           =   10440
      _cx             =   18415
      _cy             =   13785
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
      BackColorBkg    =   16777215
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
      Rows            =   0
      Cols            =   1
      FixedRows       =   0
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
   Begin VB.CommandButton CmdImprimir 
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
      Left            =   8340
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8130
      Width           =   1320
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
      Left            =   9675
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8130
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00800000&
      Height          =   7950
      Left            =   45
      Top             =   30
      Width           =   10620
   End
End
Attribute VB_Name = "frmPlandeCtas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function BuscarCuenta(CUENTA As String) As String
Dim r As Long
Dim cta As String
cta = ""
For r = 1 To Len(Trim(CUENTA))
  If Mid(Trim(CUENTA), r, 1) <> " " Then
      cta = cta & Mid(Trim(CUENTA), r, 1)
  Else
    BuscarCuenta = Trim(cta)
    Exit For
  End If
Next r
End Function
Private Sub cmdImprimir_Click()
FrmImpresiones.VSPrinter.Orientation = orPortrait
FrmImpresiones.VSPrinter.PaperSize = pprA4
FrmImpresiones.VSPrinter.Preview = True
FrmImpresiones.VSPrinter.Font.Name = "Arial"
FrmImpresiones.VSPrinter.Footer = "||Pagina " & FrmImpresiones.VSPrinter.PageCount% & " de " & "%d"
FrmImpresiones.VSPrinter.StartDoc
FrmImpresiones.VSPrinter.Header = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
FrmImpresiones.VSPrinter.Paragraph = "Plan de Cuentas             Fecha de Impresion " & Format$(Date, "dd / mm / yyyy")
FrmImpresiones.VSPrinter.Paragraph = " "
If grilla.rows > 1 Then
   FrmImpresiones.VSPrinter.TextAlign = taLeftBottom
   FrmImpresiones.VSPrinter.RenderControl = grilla.hWnd
End If
FrmImpresiones.VSPrinter.Zoom = 100
FrmImpresiones.VSPrinter.EndDoc
FrmImpresiones.BorderStyle = False
FrmImpresiones.PrintForm
FrmImpresiones.Show
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub
Sub InicioGrilla()
Dim rscta As New ADODB.Recordset
Dim largo As Long
Dim atras As Long
Dim adelante As Long
grilla.ColWidth(0) = 7000
grilla.GridLines = flexGridNone
grilla.GridLinesFixed = flexGridNone
grilla.ColAlignment(0) = flexAlignLeftCenter
grilla.AddItem "CUENTA"         '& Chr(9) & "SALDO" & Chr(9) & "TOTAL"
rscta.Open "select * from cuentas where activo=1 order by cuenta", DataEnvironment1.Sistema, adOpenDynamic, adLockReadOnly
If Not rscta.EOF Then
   Do While Not rscta.EOF
       If rscta!SUMARIZA > 0 Then
'           grilla.cell(flexcpFontBold, grilla.rows - 1, 0) = True
           grilla.AddItem Space(Len(rscta!SUMARIZA) * s2n(3)) & rscta!CUENTA & "  " & rscta!DESCRIPCION
       Else
'           grilla.cell(flexcpFontBold, grilla.rows - 1, 0) = True
           grilla.AddItem rscta!CUENTA & "  " & rscta!DESCRIPCION
           
       End If
       rscta.MoveNext
   Loop
End If
    rscta.Close
    Set rscta = Nothing
End Sub
Sub LimpioControles()
    grilla.rows = 0
    InicioGrilla
End Sub

Private Sub Form_Load()
   LimpioControles
   ucXls1.ini grilla, App.Path, "plandectas "
End Sub


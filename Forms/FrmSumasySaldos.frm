VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmSumasySaldos 
   Caption         =   "Sumas y Saldos"
   ClientHeight    =   9510
   ClientLeft      =   90
   ClientTop       =   345
   ClientWidth     =   11835
   Icon            =   "FrmSumasySaldos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9510
   ScaleWidth      =   11835
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "No ver Saldos en Cero."
      Height          =   255
      Left            =   6120
      TabIndex        =   10
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton CmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
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
      Left            =   7275
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8115
      Width           =   975
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
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
      Left            =   10770
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8115
      Width           =   975
   End
   Begin VB.CommandButton cmdaceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Ver"
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
      Left            =   10185
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   135
      Width           =   1455
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
      Left            =   8415
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8115
      Width           =   975
   End
   Begin Gestion.ucXls ucXls1 
      Height          =   810
      Left            =   165
      TabIndex        =   0
      Top             =   8115
      Width           =   975
      _extentx        =   1720
      _extenty        =   1429
   End
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Height          =   7290
      Left            =   120
      TabIndex        =   1
      Top             =   555
      Width           =   11565
      _cx             =   20399
      _cy             =   12859
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
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
      Cols            =   5
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
   Begin MSComCtl2.DTPicker fechadesde 
      Height          =   315
      Left            =   1110
      TabIndex        =   5
      Top             =   90
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   16777215
      Format          =   60882945
      CurrentDate     =   38052
   End
   Begin MSComCtl2.DTPicker fechahasta 
      Height          =   315
      Left            =   4110
      TabIndex        =   6
      Top             =   90
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   556
      _Version        =   393216
      Format          =   60882945
      CurrentDate     =   38052
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00800000&
      Height          =   7950
      Left            =   15
      Top             =   15
      Width           =   11775
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
      Left            =   3240
      TabIndex        =   8
      Top             =   120
      Width           =   615
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
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "FrmSumasySaldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public NEGRO As Variant
Public GRIS_1 As Variant
Public GRIS_2 As Variant
Public GRIS_3 As Variant

Function BuscarCuenta(Cuenta As String) As String
Dim r As Long
Dim cta As String
    cta = ""
    For r = 1 To Len(Trim(Cuenta))
        If Mid(Trim(Cuenta), r, 1) <> " " Then
            cta = cta & Mid(Trim(Cuenta), r, 1)
        Else
            BuscarCuenta = Trim(cta)
            Exit For
        End If
    Next r
End Function
Private Sub cmdAceptar_Click()
Dim rsAS As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim r As Long
Dim fi As Long
Dim sql As String
Dim saldo As Double
Dim SUMARIZA As String
Dim i As Long
    grilla.rows = 0
    InicioGrilla
    sql = "SELECT MAYOR.Cuenta, Sum(MAYOR.Debe) AS SumaDeDebe, Sum(MAYOR.Haber) AS SumaDeHaber " _
    & "FROM Asientos INNER JOIN MAYOR ON Asientos.idAsiento = MAYOR.idAsiento Where Asientos.Activo = 1 AND (Asientos.Fecha >=" _
    & ssFecha(fechadesde) & " AND ASIENTOS.FECHA <=" & ssFecha(fechahasta) & ") GROUP BY MAYOR.Cuenta" _
    & " ORDER BY mayor.cuenta ASC"
    rsAS.Open sql, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
    
    If Not rsAS.EOF Then
        Do While Not rsAS.EOF
            
            For r = 1 To grilla.rows - 1
                grilla.Row = r
                
                If BuscarCuenta(Trim(grilla.TextMatrix(r, 0))) = Trim(rsAS!Cuenta) Then
                        grilla.TextMatrix(r, 1) = Format$(s2n(rsAS!SumadeDebe, 2), "standard")
                        grilla.TextMatrix(r, 2) = Format$(s2n(rsAS!SumadeHaber, 2), "standard")
                        saldo = 0
                        saldo = s2n(rsAS!SumadeDebe, 2) - s2n(rsAS!SumadeHaber, 2)
                        If saldo > 0 Then
                            grilla.TextMatrix(r, 3) = Format$(saldo, "standard")
                            grilla.TextMatrix(r, 4) = 0
                        Else
                            If saldo < 0 Then
                                grilla.TextMatrix(r, 3) = 0
                                grilla.TextMatrix(r, 4) = Format$(Abs(saldo), "standard")
                            Else
                                grilla.TextMatrix(r, 3) = 0
                                grilla.TextMatrix(r, 4) = 0
                            End If
                        End If
                    Exit For
                End If
            Next r
            rsAS.MoveNext
        Loop
        
        'Aca pongo todas las sumas
        For fi = grilla.rows - 1 To 1 Step -1
            grilla.Row = fi
            SUMARIZA = TraerSumariza(BuscarCuenta(Trim(grilla.TextMatrix(fi, 0))))
            For r = 1 To grilla.rows - 1
                grilla.Row = r
                If BuscarCuenta(Trim(grilla.TextMatrix(r, 0))) = Trim(SUMARIZA) Then
                    'If SUMARIZA = "1" Then Stop
                    grilla.TextMatrix(r, 1) = Format$(CDbl(grilla.TextMatrix(r, 1)) + CDbl(grilla.TextMatrix(fi, 1)), "standard")
                    grilla.TextMatrix(r, 2) = Format$(CDbl(grilla.TextMatrix(r, 2)) + CDbl(grilla.TextMatrix(fi, 2)), "standard")
                    saldo = 0
                    saldo = CDbl(grilla.TextMatrix(r, 1)) - CDbl(grilla.TextMatrix(r, 2))
                    If saldo > 0 Then
                        grilla.TextMatrix(r, 3) = Format$(saldo, "standard")
                        grilla.TextMatrix(r, 4) = 0
                    Else
                        If saldo < 0 Then
                            grilla.TextMatrix(r, 3) = 0
                            grilla.TextMatrix(r, 4) = Format$(Abs(saldo), "standard")
                        Else
                            grilla.TextMatrix(r, 3) = 0
                            grilla.TextMatrix(r, 4) = 0
                        End If
                    End If
                    Exit For
                End If
            Next r
        Next fi
        
        If Check1.Value = 1 Then
            i = 1
            While i < grilla.rows
                If grilla.TextMatrix(i, 3) = 0 And grilla.TextMatrix(i, 4) = 0 Then
                    grilla.RemoveItem (i)
                    i = i - 1
                End If
                i = i + 1
            Wend
        End If
        
    Else
        MsgBox "No hay registros para mostrar con esos parametros", 48, "Atencion"
    End If

    rsAS.Close
    Set rsAS = Nothing
    cmdImprimir.enabled = True

End Sub

Private Sub cmdCancelar_Click()
    LimpioControles
End Sub

Private Sub cmdImprimir_Click()
grilla.GridLines = flexGridNone
grilla.GridLinesFixed = flexGridNone

FrmImpresiones.VSPrinter.Orientation = orPortrait
FrmImpresiones.VSPrinter.PaperSize = pprA4
FrmImpresiones.VSPrinter.Preview = True
FrmImpresiones.VSPrinter.Font.Name = "Arial"
FrmImpresiones.VSPrinter.Footer = "||Pagina " & FrmImpresiones.VSPrinter.PageCount & " de " & "%d"
FrmImpresiones.VSPrinter.StartDoc
FrmImpresiones.VSPrinter.Header = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
FrmImpresiones.VSPrinter.Paragraph = "Sumas y Saldos                          Fecha de Impresion " & Format$(Date, "dd / mm / yyyy")
FrmImpresiones.VSPrinter.Paragraph = " "
If grilla.rows > 1 Then
   FrmImpresiones.VSPrinter.TextAlign = taLeftBottom
   FrmImpresiones.VSPrinter.RenderControl = grilla.hWnd
End If
FrmImpresiones.VSPrinter.Zoom = 100
FrmImpresiones.VSPrinter.EndDoc
FrmImpresiones.Show
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Sub InicioGrilla()
Dim rscta As New ADODB.Recordset
Dim largo As Long
Dim atras As Long
Dim adelante As Long

NEGRO = "&H00000000&"
GRIS_1 = "&H00404040&"
GRIS_2 = "&H00808080&"
GRIS_3 = "&H00C0C0C0&"

    grilla.ColWidth(0) = 5000
    grilla.ColWidth(1) = 1200
    grilla.ColWidth(2) = 1200
    grilla.ColWidth(3) = 1600
    grilla.ColWidth(4) = 1800
    grilla.ColAlignment(0) = flexAlignLeftCenter
    grilla.ColAlignment(1) = flexAlignRightCenter
    grilla.ColAlignment(2) = flexAlignRightCenter
    grilla.ColAlignment(3) = flexAlignRightCenter
    grilla.ColAlignment(4) = flexAlignRightCenter
    grilla.ColDataType(1) = flexDTCurrency
    grilla.ColDataType(2) = flexDTCurrency
    grilla.ColDataType(3) = flexDTCurrency
    grilla.ColDataType(4) = flexDTCurrency
    grilla.AddItem "CUENTA" & Chr(9) & "DEBE" & Chr(9) & "HABER" & Chr(9) & "SALDO DEUDOR" & Chr(9) & "SALDO ACREEDOR"
    grilla.cell(flexcpFontBold, grilla.rows - 1, 0, grilla.rows - 1, 4) = True
    rscta.Open "select * from cuentas where activo=1 order by cuenta", DataEnvironment1.Sistema, adOpenDynamic, adLockReadOnly
    If Not rscta.EOF Then
    Dim caract As Integer
        Do While Not rscta.EOF
            If rscta!SUMARIZA > 0 Then
                grilla.AddItem Space(Len(rscta!SUMARIZA) * s2n(6)) & rscta!Cuenta & "  " & rscta!DESCRIPCION & Chr(9) & "0" & Chr(9) & "0" & Chr(9) & "0" & Chr(9) & "0"
                caract = Len(Trim(rscta!SUMARIZA))
                Select Case caract
                    Case 1:
                        grilla.cell(flexcpFontBold, grilla.rows - 1, 0) = True
                    Case 2:
                        grilla.cell(flexcpFontBold, grilla.rows - 1, 0) = True
                    Case 3:
                        grilla.cell(15, grilla.rows - 1, 0) = True

                End Select
            Else
                grilla.AddItem rscta!Cuenta & "  " & rscta!DESCRIPCION & Chr(9) & "0" & Chr(9) & "0" & Chr(9) & "0" & Chr(9) & "0"
                grilla.cell(flexcpFontBold, grilla.rows - 1, 0) = True
            End If
            rscta.MoveNext
        Loop
    End If
    grilla.cell(flexcpFontBold) = False
'    grilla.cell(flexcpBackColor, 0, 0, 0, 4) = &HE0E0E0
'    grilla.cell(flexcpForeColor, 0, 0, 0, 4) = &HC0&
'    grilla.cell(flexcpForeColor, 1, 0, grilla.rows - 1, 4) = &H800000
    rscta.Close
    Set rscta = Nothing
End Sub
Sub LimpioControles()

    fechadesde = Date
    fechahasta = Date
    grilla.rows = 0
    InicioGrilla
End Sub

Private Sub Form_Load()
    LimpioControles
ucXls1.ini grilla, App.Path & "\SumasySaldos", "SUMAS Y SALDOS"

End Sub

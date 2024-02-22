VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBalance 
   Caption         =   "Balance"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10740
   Icon            =   "frmBalance.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9105
   ScaleWidth      =   10740
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Solo para asiento de cierre"
      Height          =   1095
      Left            =   2310
      TabIndex        =   12
      Top             =   7935
      Width           =   3945
      Begin VB.TextBox txtTitulo 
         Height          =   315
         Left            =   150
         TabIndex        =   14
         Text            =   "ASIENTO DE CIERRE"
         Top             =   645
         Width           =   3645
      End
      Begin MSComCtl2.DTPicker dtFecha 
         Height          =   315
         Left            =   150
         TabIndex        =   13
         Top             =   270
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         Format          =   62849025
         CurrentDate     =   40505
      End
   End
   Begin VB.CommandButton cmdCierre 
      Caption         =   "Asiento Cierre"
      Height          =   1005
      Left            =   1275
      Picture         =   "frmBalance.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8010
      Width           =   960
   End
   Begin VB.CheckBox Check1 
      Caption         =   "No ver Saldos en Cero."
      Height          =   255
      Left            =   6000
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
      Top             =   8430
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
      Left            =   8505
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8430
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
      Left            =   9045
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1455
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
      Left            =   9690
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8430
      Width           =   975
   End
   Begin Gestion.ucXls ucXls1 
      Height          =   990
      Left            =   165
      TabIndex        =   3
      Top             =   8025
      Width           =   975
      _extentx        =   1720
      _extenty        =   1746
   End
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Height          =   7290
      Left            =   150
      TabIndex        =   4
      Top             =   570
      Width           =   10455
      _cx             =   18441
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
      Cols            =   3
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
      Height          =   330
      Left            =   1140
      TabIndex        =   5
      Top             =   142
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   582
      _Version        =   393216
      CalendarBackColor=   16777215
      Format          =   62849025
      CurrentDate     =   38052
   End
   Begin MSComCtl2.DTPicker fechahasta 
      Height          =   330
      Left            =   4140
      TabIndex        =   6
      Top             =   142
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   582
      _Version        =   393216
      Format          =   62849025
      CurrentDate     =   38052
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
      Left            =   270
      TabIndex        =   8
      Top             =   180
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
      Left            =   3270
      TabIndex        =   7
      Top             =   180
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00800000&
      Height          =   7875
      Left            =   45
      Top             =   30
      Width           =   10665
   End
End
Attribute VB_Name = "frmBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Dim r As Long
Dim sql As String
Dim saldo As Double
Dim fi As Long
Dim SUMARIZA As String
Dim i As Long

    Grilla.rows = 0
    InicioGrilla
    sql = "SELECT MAYOR.Cuenta, Sum(MAYOR.Debe) AS SumaDeDebe, Sum(MAYOR.Haber) AS SumaDeHaber " _
    & "FROM Asientos INNER JOIN MAYOR ON Asientos.idAsiento = MAYOR.idAsiento Where Asientos.Activo = 1 AND (Asientos.Fecha >=" _
    & ssFecha(fechadesde) & " AND ASIENTOS.FECHA <=" & ssFecha(fechahasta) & ") GROUP BY MAYOR.Cuenta" _
    & " ORDER BY mayor.cuenta ASC"
    rsAS.Open sql, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
    If Not rsAS.EOF Then
        Do While Not rsAS.EOF
            saldo = 0
             saldo = s2n(rsAS!SumadeDebe, 2) - s2n(rsAS!SumadeHaber, 2)
            For r = 1 To Grilla.rows - 1
                Grilla.Row = r
                If BuscarCuenta(Trim(Grilla.TextMatrix(r, 0))) = Trim(rsAS!Cuenta) Then
                    Grilla.TextMatrix(r, 1) = Format$(saldo, "standard")
                    Debug.Print saldo
                End If
            Next r
            rsAS.MoveNext
        Loop
        
    Else
        MsgBox "No hay registros para mostrar con esos parametros", 48, "Atencion"
        Exit Sub
    End If

    rsAS.Close
    Set rsAS = Nothing
    
    'Aca pongo todas las sumas
    
        For fi = Grilla.rows - 1 To 1 Step -1
            Grilla.Row = fi
            SUMARIZA = TraerSumariza(BuscarCuenta(Trim(Grilla.TextMatrix(fi, 0))))
            For r = 1 To Grilla.rows - 1
                Grilla.Row = r
                If BuscarCuenta(Trim(Grilla.TextMatrix(r, 0))) = Trim(SUMARIZA) Then
                
                    Grilla.TextMatrix(r, 1) = Format$(CDbl(Grilla.TextMatrix(r, 1)) + CDbl(Grilla.TextMatrix(fi, 1)), "standard")
                    Grilla.cell(flexcpForeColor, r, 1) = vbWhite
                    Grilla.TextMatrix(r, 2) = Format$(CDbl(Grilla.TextMatrix(r, 2)) + CDbl(Grilla.TextMatrix(fi, 1)), "standard")
                    Grilla.cell(flexcpFontBold, r, 2) = True
                    Exit For
                End If
            Next r
        Next fi
        
        If Check1.Value = 1 Then
            i = 1
            While i < Grilla.rows
                If s2n(Grilla.TextMatrix(i, 1)) = 0 Then
                    Grilla.RemoveItem (i)
                    i = i - 1
                End If
                i = i + 1
            Wend
        End If

cmdImprimir.enabled = True
End Sub

Private Sub cmdCancelar_Click()
    LimpioControles
End Sub

Private Sub cmdCierre_Click()
Dim i As Long, aCierre As New Asiento, aEjercicio As Long, iCierre As Long, tmp
Dim aCuenta As String, aImputable As Boolean
DataEnvironment1.Sistema.BeginTrans
aCierre.nuevo txtTitulo, dtFecha, "C"
With Grilla
    If .rows > 1 Then
        For i = 1 To .rows - 1
            tmp = Split(Trim(.TextMatrix(i, 0)), " ")
            aCuenta = tmp(0)
            aImputable = obtenerDeSQL("select imputable from cuentas where cuenta=" & ssTexto(aCuenta))
            If aImputable Then
                If s2n(.TextMatrix(i, 1)) <> 0 Then
                    aCierre.AcumularItem aCuenta, 0, s2n(.TextMatrix(i, 1))
                End If
                If s2n(.TextMatrix(i, 2)) <> 0 Then
                    aCierre.AcumularItem aCuenta, s2n(.TextMatrix(i, 1)), 0
                End If
            End If
        Next
    Else
    End If
End With
aEjercicio = obtenerDeSQL("select ejercicio from ejercicio where activo=1")
iCierre = NuevoDocumento("A.C", aEjercicio, 0, 0)
If aCierre.Grabar(iCierre, False) > 0 Then
    MsgBox "ASIENTO REALIZADO...", vbInformation
    DataEnvironment1.Sistema.CommitTrans
Else
    DataEnvironment1.Sistema.RollbackTrans
    'BorroDocumento iCierre
    'BorroDocumento 12161
End If
End Sub

Private Sub cmdImprimir_Click()
Grilla.GridLines = flexGridNone
Grilla.GridLinesFixed = flexGridNone
FrmImpresiones.VSPrinter.Orientation = orPortrait
FrmImpresiones.VSPrinter.PaperSize = pprA4
FrmImpresiones.VSPrinter.Preview = True
FrmImpresiones.VSPrinter.Font.Name = "Arial"
FrmImpresiones.VSPrinter.Footer = "||Pagina " & FrmImpresiones.VSPrinter.PageCount & " de " & "%d"

FrmImpresiones.VSPrinter.StartDoc
FrmImpresiones.VSPrinter.Header = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
FrmImpresiones.VSPrinter.Paragraph = "Balance                            Fecha de Impresion " & Format$(Date, "dd / mm / yyyy")
FrmImpresiones.VSPrinter.Paragraph = " "
If Grilla.rows > 1 Then
   FrmImpresiones.VSPrinter.TextAlign = taLeftBottom
   FrmImpresiones.VSPrinter.RenderControl = Grilla.hWnd
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
    
    Grilla.ColWidth(0) = 7000
    Grilla.ColWidth(1) = 1500
    Grilla.ColWidth(2) = 1500
    Grilla.ColAlignment(0) = flexAlignLeftCenter
    Grilla.ColAlignment(1) = flexAlignRightCenter
    Grilla.ColAlignment(2) = flexAlignRightCenter
    Grilla.ColDataType(1) = flexDTCurrency
    Grilla.ColDataType(2) = flexDTCurrency
    Grilla.AddItem "CUENTA" & Chr(9) & "SALDO" & Chr(9) & "TOTAL"
    Grilla.cell(flexcpFontBold, Grilla.rows - 1, 0, Grilla.rows - 1, 2) = True
    rscta.Open "select * from cuentas where activo=1 order by cuenta", DataEnvironment1.Sistema, adOpenDynamic, adLockReadOnly
    If Not rscta.EOF Then
        Do While Not rscta.EOF
            If rscta!SUMARIZA > 0 Then
                Grilla.cell(flexcpFontBold, Grilla.rows - 1, 0) = True
                Grilla.AddItem Space(Len(rscta!SUMARIZA) * s2n(3)) & rscta!Cuenta & "  " & rscta!DESCRIPCION & Chr(9) & "0" & Chr(9) & "0"
            Else
                Grilla.cell(flexcpFontBold, Grilla.rows - 1, 0) = True
                Grilla.AddItem rscta!Cuenta & "  " & rscta!DESCRIPCION & Chr(9) & "0" & Chr(9) & "0"
            End If

            rscta.MoveNext
                        
        Loop
    End If
    'grilla.cell(flexcpBackColor, 0, 0, 0, 2) = &HE0E0E0
    'grilla.cell(flexcpForeColor, 0, 0, 0, 2) = &HC0&
    'grilla.cell(flexcpForeColor, 1, 0, grilla.rows - 1, 2) = &H800000
    rscta.Close
    Set rscta = Nothing
End Sub
Sub LimpioControles()
    fechadesde = Date
    fechahasta = Date
    Grilla.rows = 0
    InicioGrilla
End Sub

Private Sub Form_Load()
    LimpioControles
    ucXls1.ini Grilla, "C:\Balance.xls", "BALANCE "
End Sub

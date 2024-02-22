VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMayorCont 
   Caption         =   "Libro Mayor"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12510
   Icon            =   "FrmMayorCont.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   12510
   StartUpPosition =   2  'CenterScreen
   Begin Gestion.ucCoDe cueH 
      Height          =   315
      Left            =   180
      TabIndex        =   14
      Top             =   960
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   556
      CodigoWidth     =   1000
   End
   Begin Gestion.ucCoDe cueD 
      Height          =   300
      Left            =   195
      TabIndex        =   13
      Top             =   570
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   529
      CodigoWidth     =   1000
   End
   Begin Gestion.ucXls ucXls1 
      Height          =   1065
      Left            =   9945
      TabIndex        =   12
      Top             =   180
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1879
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
      Height          =   1065
      Left            =   8025
      Picture         =   "FrmMayorCont.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   165
      Width           =   975
   End
   Begin VB.CheckBox ChkSinMov 
      Caption         =   "Cuentas sin Movimiento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7050
      TabIndex        =   2
      Top             =   1335
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VSFlex7LCtl.VSFlexGrid Grilla 
      Height          =   6600
      Left            =   45
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1740
      Width           =   12450
      _cx             =   21960
      _cy             =   11642
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
      Rows            =   1
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
   Begin VB.TextBox FechaEH 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   1605
      TabIndex        =   10
      Top             =   1335
      Width           =   1335
   End
   Begin VB.TextBox FechaED 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   195
      TabIndex        =   9
      Top             =   1335
      Width           =   1335
   End
   Begin VB.CommandButton CmdEjecutar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Generar Listado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   7050
      Picture         =   "FrmMayorCont.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   165
      Width           =   975
   End
   Begin VB.CommandButton CmdSalir 
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
      Height          =   1065
      Left            =   9000
      Picture         =   "FrmMayorCont.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   165
      Width           =   945
   End
   Begin MSComCtl2.DTPicker FechaH 
      Height          =   360
      Left            =   2535
      TabIndex        =   1
      Top             =   75
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   635
      _Version        =   393216
      Format          =   58785793
      CurrentDate     =   38525
   End
   Begin MSComCtl2.DTPicker FechaD 
      Height          =   360
      Left            =   765
      TabIndex        =   0
      Top             =   75
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   635
      _Version        =   393216
      Format          =   58785793
      CurrentDate     =   38525
   End
   Begin VB.Label LblNEjer 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5520
      TabIndex        =   8
      Top             =   90
      Width           =   675
   End
   Begin VB.Label Label5 
      Caption         =   "Ejercicio Activo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4140
      TabIndex        =   7
      Top             =   105
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "Entre                                    y"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   210
      TabIndex        =   5
      Top             =   150
      Width           =   3585
   End
End
Attribute VB_Name = "FrmMayorCont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'Dim rsctas As New ADODB.Recordset    ' NO ME PONGAN COSAS ACA, quedan objetos colgados, por lo menos matenlos al cerrar form
Dim rsEjerAc As New ADODB.Recordset   ' cuando tenga tiempo saco este
Dim respuesta As Integer

'Dim rsSaldo As New ADODB.Recordset
Dim sql As String, Sql2 As String
Dim EjercicioActivo, x, z As Long
Dim SaldoDA As Double
Dim SaldoD, SaldoA, Arrastre As Double
Dim total As Double
Dim rsEjerSinCerrar As New ADODB.Recordset

Private Sub CmdEjecutar_Click()
    Dim rsctas As New ADODB.Recordset
    Dim saldoAcreedor As Double
    Dim SumadeDebe, SumadeHaber, x As Double
    Dim CadenaDejer As String
    
    If FechaD.Value > FechaH.Value Then
          MsgBox "La Fecha DESDE debe ser menor que la Fecha HASTA", vbExclamation, "Error de Ingreso"
          Exit Sub
    End If
'    If CmbCtaD.ListIndex > CmbCtaH.ListIndex Then
'       MsgBox "La Cuenta DESDE debe ser menor que la Cuenta HASTA", vbExclamation, "Error de Ingreso"
'       Exit Sub
'    End If
    
    Arrastre = 0
    'InicioGrilla
    If grilla.rows > 2 Then
        grilla.clear
        grilla.rows = 1
        'For x = (grilla.rows - 1) To 1 Step -1
        '   grilla.RemoveItem (x)
        'Next
    End If
    InicioGrilla
    CalculoEntreFechas
    If respuesta = 6 Or respuesta = 0 Then
            '*******************************************************GERMAN
            
            
            CadenaDejer = TraerEjerciciosAbiertos
            
            'CadenaDejer = "'" & EjercicioActivo & "') And ((Asientos.Activo) = 1))  ORDER BY Ejercicio.idEjercicio"
        
            
            '*******************************************************
            Dim tempo
            Dim AuxCtaD, AuxCtaH As String
            AuxCtaD = cueD.codigo 'SacarCodigo(CmbCtaD)
            AuxCtaH = cueH.codigo 'SacarCodigo(CmbCtaH)
            sql = "SELECT cuenta,descripcion FROM CUENTAS WHERE CUENTA BETWEEN '" & AuxCtaD & "' AND '" & AuxCtaH & "' order by cuenta"
            rsctas.Open sql, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
            Do While Not rsctas.EOF
                'agrupaba por ejercicio y devolvia un registro por ejercicio y se necesitaba todo en uno
                'Sql = "SELECT  Sum(MAYOR.Debe) AS SumaDeDebe, Sum(MAYOR.Haber) AS " _
                    & "SumaDeHaber FROM (Asientos INNER JOIN Ejercicio ON Asientos.idEjercicio = " _
                    & "Ejercicio.idEjercicio) INNER JOIN MAYOR ON Asientos.idAsiento = MAYOR.idAsiento " _
                    & "Where (((Asientos.Fecha) < " & ssFecha(FechaD) & "))GROUP BY MAYOR.Cuenta, " _
                    & "Ejercicio.idEjercicio, Ejercicio.Ejercicio, Asientos.Activo Having " _
                    & "(((MAYOR.Cuenta) = '" & rsctas!cuenta & "') And ((Ejercicio.idEjercicio) " _
                    & "= " & CadenaDejer & ") ORDER BY Ejercicio.idEjercicio"
                
                sql = "SELECT  Sum(MAYOR.Debe) AS SumaDeDebe, Sum(MAYOR.Haber) AS " _
                            & "SumaDeHaber ,SUM (MAYOR.Debe-MAYOR.Haber) AS SALDO " _
                            & "FROM (Asientos INNER JOIN MAYOR ON Asientos.idAsiento = MAYOR.idAsiento) " _
                            & "Where ((Asientos.Fecha) < " & ssFecha(FechaD.Value) & ") " _
                            & "GROUP BY MAYOR.Cuenta, Asientos.Activo " _
                            & "Having ((MAYOR.Cuenta) = '" & rsctas!CUENTA & "') And ((Asientos.Activo) = 1)"

                    tempo = obtenerDeSQL(sql)

            
                   'rsSaldo.Open Sql, DataEnvironment1.SISTEMA, adOpenKeyset, adLockReadOnly
                If IsEmpty(tempo) Then
                    SumadeDebe = 0
                    SumadeHaber = 0
                Else
                    SumadeDebe = s2n(tempo(0))
                    SumadeHaber = s2n(tempo(1))
                End If
               
            '   Do While Not rsSaldo.EOF
                Sql2 = "SELECT Ejercicio.idEjercicio, Ejercicio.Denominacion, Ejercicio.FechaInicio, " _
                    & "Ejercicio.FechaFin, Asientos.NroAsiento,Asientos.concepto ,MAYOR.Cuenta, Asientos.Fecha, " _
                    & "MAYOR.Debe, MAYOR.Haber, MAYOR.idMayor,MAYOR.Comprobante,Asientos.Activo FROM (Asientos INNER JOIN " _
                    & "Ejercicio ON Asientos.Ejercicio = Ejercicio.idEjercicio) INNER JOIN MAYOR ON " _
                    & "Asientos.idAsiento = MAYOR.idAsiento Where ((Ejercicio.idEjercicio) " _
                    & "= " & CadenaDejer & " And ((MAYOR.Cuenta) = '" & rsctas!CUENTA & "') " _
                    & "And ((Asientos.Fecha) >= " & ssFecha(FechaD) & " And (Asientos.Fecha) " _
                    & "<= " & ssFecha(FechaH) & ") And ((Asientos.Activo) = 1) ORDER BY Ejercicio.idEjercicio, " _
                    & "MAYOR.Cuenta, Asientos.Fecha"

                    rsEjerAc.Open Sql2, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly

                  If Not rsEjerAc.EOF Then 'And ChkSinMov.Value = 1
                        SaldoDA = 0
            '            SaldoDA = rsSaldo!SumadeDebe - rsSaldo!SumadeHaber
                         SaldoDA = SumadeDebe - SumadeHaber
                        If SaldoDA > 0 Then
                           SaldoD = s2n(SaldoDA, 2)
                           SaldoA = 0
                        Else
                           SaldoA = s2n(SaldoDA, 2)
                           SaldoD = 0
                        End If
                        If grilla.rows > 2 Then
                           grilla.AddItem ""
                        '   grilla.cell(flexcpForeColor, grilla.rows - 2, 5) = vbBlue
                           grilla.cell(flexcpFontBold, grilla.rows - 2, 5) = True
                        '   grilla.cell(flexcpForeColor, grilla.rows - 2, 6) = vbRed
                           grilla.cell(flexcpFontBold, grilla.rows - 2, 6) = True
                           grilla.cell(flexcpBackColor, grilla.rows - 1, 0, grilla.rows - 1, 7) = &H8000000F
                        End If
                        If (SaldoA < 0) Then
                            saldoAcreedor = aPositivo(SaldoA)
                        Else
                            saldoAcreedor = SaldoA
                        End If
                        grilla.AddItem "Cuenta " & Chr(9) & rsctas!CUENTA & Chr(9) & rsctas!DESCRIPCION & Chr(9) & "S.Anterior" & Chr(9) & "" & Chr(9) & Format$(SaldoD, "standard") & Chr(9) & Format$(saldoAcreedor, "standard") & Chr(9) & rsEjerAc!COMPROBANTE
                        grilla.cell(flexcpFontBold, grilla.rows - 1, 0, grilla.rows - 1, 7) = True
                      '  grilla.cell(flexcpForeColor, grilla.rows - 1, 6) = vbRed
                      '  grilla.cell(flexcpForeColor, grilla.rows - 1, 5) = vbBlue
                        total = 0
                        If SaldoD <> 0 Then
                            total = SaldoD
                        Else
                            total = SaldoA
                        End If
                        
                        Do While Not rsEjerAc.EOF
                           Cargogrilla
                           rsEjerAc.MoveNext
                        Loop
                        'SumadeDebe = 0: SumadeHaber = 0

                End If
                'rsEjerAc.Close
                Set rsEjerAc = Nothing
            '  rsSaldo.MoveNext
            '  Loop
            '  rsSaldo.Close
                rsctas.MoveNext
            Loop
            'rsctas.Close
            CmdImprimir.enabled = True
            Set rsctas = Nothing
   End If
End Sub

Function aPositivo(ByVal valor As Double) As Double 'esto convierte todo numero negativo a positivo
        If (valor < 0) Then
            aPositivo = s2n(valor - (valor * 2))
        End If
End Function

Sub Cargogrilla()
Dim ArrD, ArrA, SaldoAcree As Double
If rsEjerAc!Debe > 0 Then
   total = s2n(total + rsEjerAc!Debe)
Else
   total = s2n(total - rsEjerAc!haber)
End If
If total > 0 Then
   ArrD = total
Else
   ArrA = total
End If
If (ArrA < 0) Then
    SaldoAcree = aPositivo(ArrA)
Else
    SaldoAcree = ArrA
End If
grilla.AddItem rsEjerAc!fecha & Chr(9) & rsEjerAc!NroAsiento & Chr(9) & rsEjerAc!concepto & Chr(9) & Format$(rsEjerAc!Debe, "standard") & Chr(9) & Format$(rsEjerAc!haber, "standard") & Chr(9) & Format$(ArrD, "standard") & Chr(9) & Format$(SaldoAcree, "standard")
End Sub
Sub CargoTitulo()
'x = z + 1
grilla.AddItem "Fecha" & Chr(9) & "Asiento" & Chr(9) & "Descripcion" & Chr(9) & "Debe" & Chr(9) & "Haber" & Chr(9) & "Saldo Deudor" & Chr(9) & "Saldo Acredor"
'grilla.cell(flexcpBackColor, grilla.rows - 1, 0, grilla.rows - 1, 6) = &H8000000F
grilla.cell(flexcpFontBold, grilla.rows - 1, 0, grilla.rows - 1, 6) = True
End Sub

Sub CalculoEntreFechas()
Dim Control As Boolean

Control = True


'**************************LO SIG FUE COLOCADO EL 11/4/07 PARA QUE SIEMPRE VERIFIQUE LA FECHA DEL EJERCICIO
sql = "SELECT * from ejercicio where activo =1"

    rsEjerAc.Open sql, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly

If Not rsEjerAc.EOF Then
   LblNEjer.caption = rsEjerAc!denominacion
   EjercicioActivo = rsEjerAc!idejercicio
   FechaED.Text = rsEjerAc!FechaInicio
   FechaEH.Text = rsEjerAc!FechaFin
Else
   MsgBox "No hay ningun ejercicio activo", vbInformation, "Error"
   Exit Sub
End If
Set rsEjerAc = Nothing
'rsEjerAc.Close
'*************************
   
   If FechaD <= CDate(FechaED) Or FechaD > CDate(FechaEH) Then
      Control = False
   End If
   If FechaH < CDate(FechaED) Or FechaH >= CDate(FechaEH) Then
      Control = False
   End If
'If Control = False Then'SACADO PRA GREEN OIL
'   'MSGBOX DEVUELVE 6 SI ES SI Y DEVUELVE 7 SI NO
'   respuesta = MsgBox("El Periodo seleccionado no corresponde al ejercicio Actual. A continuacion se mostraran los asientos de los Ejercicios no cerrados. ¿Desea continuar?.", vbYesNo, "Error de Intervalo")
'
'   FechaD.SetFocus
'   Exit Sub
'End If

End Sub

Private Sub cmdImprimir_Click()
grilla.GridLines = flexGridNone
grilla.GridLinesFixed = flexGridNone

FrmImpresiones.VSPrinter.Orientation = orPortrait
FrmImpresiones.VSPrinter.PaperSize = pprA4
FrmImpresiones.VSPrinter.Preview = True
FrmImpresiones.VSPrinter.Font.Name = "Arial"
FrmImpresiones.VSPrinter.FontSize = 12
FrmImpresiones.VSPrinter.Header = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
FrmImpresiones.VSPrinter.FontSize = 10
FrmImpresiones.VSPrinter.Footer = "||Pagina " & FrmImpresiones.VSPrinter.PageCount & " de " & "%d"
FrmImpresiones.VSPrinter.StartDoc
FrmImpresiones.VSPrinter.Paragraph = "Listado Mayor al " & Format$(Date, "dd / mm / yyyy")
FrmImpresiones.VSPrinter.Paragraph = " "
'FrmImpresiones.VSPrinter.Paragraph = "Rango de Fechas : " & FechaD & " - " & FechaH & "     Rango de Cuentas : " & CmbCtaD & "  -  " & CmbCtaH
FrmImpresiones.VSPrinter.Paragraph = "Rango de Fechas : " & FechaD & " - " & FechaH & "     Rango de Cuentas : " & cueD.codigo & " " & cueD.DESCRIPCION & "  -  " & cueH.codigo & " " & cueH.DESCRIPCION
FrmImpresiones.VSPrinter.Paragraph = " "

If grilla.rows > 1 Then
   FrmImpresiones.VSPrinter.TextAlign = taLeftBottom
   grilla.cols = 7
   FrmImpresiones.VSPrinter.RenderControl = grilla.hWnd
End If
FrmImpresiones.VSPrinter.Zoom = 100
FrmImpresiones.VSPrinter.EndDoc
FrmImpresiones.Show
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
FechaD.Value = Date
FechaH.Value = Date
'**********************************************************
'**********************************************************
'**********************************************************
'                  FechaD.Value = #6/17/2007#
'                  FechaH.Value = #12/30/2007#
'                  CmbCtaD.Text = "1"
'                  CmbCtaH.Text = "5310103"
'**********************************************************
'**********************************************************
'**********************************************************

sql = "SELECT * from ejercicio where activo =1"

    rsEjerAc.Open sql, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly

If Not rsEjerAc.EOF Then
   LblNEjer.caption = rsEjerAc!denominacion
   EjercicioActivo = rsEjerAc!idejercicio
   FechaED.Text = rsEjerAc!FechaInicio
   FechaEH.Text = rsEjerAc!FechaFin
Else
   MsgBox "No hay ningun ejercicio activo", vbInformation, "Error"
   Exit Sub
End If
rsEjerAc.Close

    'Sql = "Select id,cuenta,descripcion FROM Cuentas order by cuenta ASC"
'    rsctas.Open Sql, DataEnvironment1.sistema, adOpenKeyset, adLockReadOnly
'    Do While Not rsctas.EOF
'       CmbCtaD.AddItem rsctas!cuenta & String(10 - Len(rsctas!cuenta), " ") & rsctas!descripcion
'       CmbCtaD.ItemData(CmbCtaD.NewIndex) = rsctas!ID
'
'       CmbCtaH.AddItem rsctas!cuenta & String(10 - Len(rsctas!cuenta), " ") & rsctas!descripcion
'       CmbCtaH.ItemData(CmbCtaD.NewIndex) = rsctas!ID
'       rsctas.MoveNext
'    Loop

        cueD.ini "select descripcion from cuentas where cuenta = '###'", "select cuenta as [ Cuenta       ], descripcion as  [Descripcion                            ] from cuentas where activo = 1", True
        cueH.ini "select descripcion from cuentas where cuenta = '###'", "select cuenta as [ Cuenta       ], descripcion as  [Descripcion                            ] from cuentas where activo = 1", True

    'rsctas.Close
    InicioGrilla
    ucXls1.ini grilla, "c:\Mayor", "MAYOR"
End Sub

Sub InicioGrilla()
   grilla.FixedCols = 0
   grilla.cols = 8
   grilla.clear
   grilla.FontSize = 8
   grilla.TextMatrix(0, 0) = "Fecha"
   grilla.TextMatrix(0, 1) = "Asiento"
   grilla.TextMatrix(0, 2) = "Descripcion"
   grilla.TextMatrix(0, 3) = "Debe"
   grilla.TextMatrix(0, 4) = "Haber"
   grilla.TextMatrix(0, 5) = "S.Deudor"
   grilla.TextMatrix(0, 6) = "S.Acreedor"
   grilla.TextMatrix(0, 7) = "Concepto"
   
   grilla.ColWidth(0) = 1000
   grilla.ColWidth(1) = 900
   grilla.ColWidth(2) = 3300
   grilla.ColWidth(3) = 1200
   grilla.ColWidth(4) = 1200
   grilla.ColWidth(5) = 1400
   grilla.ColWidth(6) = 1400
   grilla.ColWidth(7) = 2000
End Sub

Function SacarCodigo(dato As String) As String
   Dim i As Long
   Dim Resul As String
   For i = 1 To Len(dato)
      If Mid(dato, i, 1) <> Chr(32) Then
        Resul = Resul + Mid(dato, i, 1)
      Else
         SacarCodigo = Resul
         Exit Function
      End If
   Next
   SacarCodigo = Resul
End Function

Private Sub Form_Resize()
    Anclar grilla, Me, anclarLadosTodos
End Sub

Function TraerEjerciciosAbiertos() As String
Dim i As Integer
Dim cerrar As String
'**************************LOS SIG FUE COLOCADO EL 11/4/07 ESTO ES PORQUE UN CLIENTE QUERIA VER TODOS LOS ASIENTOS
'**************************DE EJERCICIOS ANTERIORES
            
            sql = "SELECT * from ejercicio where Cerrado =0" 'con esto vemos todos los ejecicios sin cerrar

                rsEjerSinCerrar.Open sql, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly

            cerrar = " "
            For i = 1 To rsEjerSinCerrar.RecordCount Step 1
                If cerrar = " " Then
                    cerrar = cerrar & rsEjerSinCerrar!idejercicio
                Else
                    cerrar = cerrar & " or (Ejercicio.idEjercicio) = " & rsEjerSinCerrar!idejercicio
                End If
               rsEjerSinCerrar.MoveNext
            Next i
            
            Set rsEjerSinCerrar = Nothing
cerrar = cerrar & ")"
'**************************
TraerEjerciciosAbiertos = cerrar
End Function



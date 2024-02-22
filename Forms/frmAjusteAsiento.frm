VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAjusteAsiento 
   Caption         =   "Ajuste por inflacion"
   ClientHeight    =   10155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13035
   LinkTopic       =   "Form1"
   ScaleHeight     =   10155
   ScaleWidth      =   13035
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdVista 
      Caption         =   "Vista previa"
      Height          =   480
      Left            =   6315
      TabIndex        =   20
      Top             =   9645
      Width           =   1605
   End
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "Quitar cuenta"
      Height          =   480
      Left            =   8130
      TabIndex        =   19
      Top             =   9630
      Width           =   1605
   End
   Begin Gestion.ucXls ucXls1 
      Height          =   345
      Left            =   300
      TabIndex        =   16
      Top             =   9690
      Visible         =   0   'False
      Width           =   915
      _extentx        =   1614
      _extenty        =   609
   End
   Begin MSComCtl2.DTPicker dtFechaAsiento 
      Height          =   300
      Left            =   9855
      TabIndex        =   14
      Top             =   9825
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   529
      _Version        =   393216
      Format          =   176160769
      CurrentDate     =   43535
   End
   Begin VB.CommandButton cmdAsiento 
      Caption         =   "Generar Asiento"
      Height          =   495
      Left            =   11580
      TabIndex        =   10
      Top             =   9630
      Width           =   1260
   End
   Begin VB.CheckBox chkAnual 
      Caption         =   "Anual"
      Height          =   360
      Left            =   3225
      TabIndex        =   5
      Top             =   135
      Width           =   1095
   End
   Begin VB.TextBox txtIndice 
      Enabled         =   0   'False
      Height          =   345
      Left            =   5055
      TabIndex        =   4
      Text            =   "0"
      Top             =   150
      Width           =   1095
   End
   Begin VB.CommandButton cmdMostrar 
      Caption         =   "Mostrar"
      Height          =   450
      Left            =   11595
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1275
   End
   Begin VSFlex7LCtl.VSFlexGrid gSumariza 
      Height          =   2820
      Left            =   105
      TabIndex        =   1
      Top             =   615
      Width           =   12795
      _cx             =   22569
      _cy             =   4974
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
      FixedCols       =   1
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
   Begin VSFlex7LCtl.VSFlexGrid gSumariza3 
      Height          =   2820
      Left            =   75
      TabIndex        =   2
      Top             =   3705
      Width           =   12795
      _cx             =   22569
      _cy             =   4974
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
      FixedCols       =   1
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
   Begin VSFlex7LCtl.VSFlexGrid gSumariza4 
      Height          =   2820
      Left            =   90
      TabIndex        =   3
      Top             =   6780
      Width           =   12765
      _cx             =   22516
      _cy             =   4974
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
      FixedCols       =   1
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
   Begin MSComCtl2.DTPicker dtMes 
      Height          =   330
      Left            =   855
      TabIndex        =   6
      Top             =   165
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "MMMM"
      Format          =   176160771
      CurrentDate     =   43532
   End
   Begin MSComCtl2.DTPicker dtAnio 
      Height          =   330
      Left            =   1980
      TabIndex        =   7
      Top             =   165
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "yyyy"
      Format          =   176160771
      CurrentDate     =   43532
   End
   Begin Gestion.ucXls ucXls2 
      Height          =   345
      Left            =   1455
      TabIndex        =   17
      Top             =   9705
      Visible         =   0   'False
      Width           =   915
      _extentx        =   1614
      _extenty        =   609
   End
   Begin Gestion.ucXls ucXls3 
      Height          =   345
      Left            =   2640
      TabIndex        =   18
      Top             =   9720
      Visible         =   0   'False
      Width           =   915
      _extentx        =   1614
      _extenty        =   609
   End
   Begin VB.Label Label6 
      Caption         =   "Fecha Asiento"
      Height          =   225
      Left            =   9825
      TabIndex        =   15
      Top             =   9615
      Width           =   1920
   End
   Begin VB.Label Label5 
      Caption         =   "Diferencia entre ambos"
      Height          =   360
      Left            =   9690
      TabIndex        =   13
      Top             =   6555
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Balance Actual por coeficiente"
      Height          =   360
      Left            =   9720
      TabIndex        =   12
      Top             =   3465
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Balance Actual"
      Height          =   360
      Left            =   9735
      TabIndex        =   11
      Top             =   360
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "Periodo"
      Height          =   330
      Left            =   135
      TabIndex        =   9
      Top             =   195
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Indice"
      Height          =   300
      Left            =   4395
      TabIndex        =   8
      Top             =   210
      Width           =   675
   End
End
Attribute VB_Name = "frmAjusteAsiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAsiento_Click()
Dim AjusteAsiento As New Asiento, i As Long, sCuentaAjuste As String, nIDdoc As Long, nEjercicio As Long
Dim AcumuloHaber As Double, AcumuloDebe As Double
sCuentaAjuste = sSinNull(obtenerDeSQL("select cuenta from cuentasparam where codigo ='API'"))
If sCuentaAjuste = "" Then
    MsgBox "No existe cuenta contable definida para el asiento.", vbInformation
    Exit Sub
End If

nEjercicio = nSinNull(obtenerDeSQL("select idejercicio from ejercicio where fechainicio<=" & ssFecha(dtFechaAsiento) & " and fechafin>=" & ssFecha(dtFechaAsiento)))
If nEjercicio = 0 Then
    MsgBox "No existe ejercicio para la fecha indicada.", vbInformation
    Exit Sub
End If

relojito True
AjusteAsiento.nuevo "Ajuste por inflacion", dtFechaAsiento, "A"

With gSumariza4
    AcumuloDebe = 0
    AcumuloHaber = 0
    For i = 1 To .rows - 1
        AcumuloDebe = AcumuloDebe + s2n(.TextMatrix(i, 3))
        AcumuloHaber = AcumuloHaber + s2n(.TextMatrix(i, 4))
        AjusteAsiento.AgregarItem .TextMatrix(i, 1), s2n(.TextMatrix(i, 3)), .TextMatrix(i, 4)
    Next
End With

If AcumuloDebe >= AcumuloHaber Then
    AcumuloHaber = s2n(AcumuloDebe - AcumuloHaber)
    AcumuloDebe = 0
ElseIf AcumuloHaber > AcumuloDebe Then
    AcumuloDebe = s2n(AcumuloHaber - AcumuloDebe)
    AcumuloHaber = 0
End If

AjusteAsiento.AgregarItem sCuentaAjuste, s2n(AcumuloDebe), s2n(AcumuloHaber)


nIDdoc = 0
If AjusteAsiento.Grabar(nIDdoc, , nEjercicio) > 0 Then
    MsgBox "Asiento generado correctamente.", vbInformation
End If
relojito False
End Sub

Private Sub cmdMostrar_Click()
Dim rsAsientos As New ADODB.Recordset, i As Long, cConsul As String, pActual As String, pPrueba As String, rInicio As Long, gEncontro As Boolean
Dim rsMayor As New ADODB.Recordset, x As Long, dFechaD As Date, dFechaH As Date
Dim dPeriodoActual As String, dCuentaActual As Long
Dim mCoeficiente As Double, mDebeActual As Double, mHaberActual As Double, mDebeNuevo As Double, mHaberNuevo As Double, mDebeAjuste As Double, mHaberAjuste As Double

g_ini2
relojito True
If chkAnual Then
    dFechaD = CDate("01/01/" & qAnio)
    'dFechaH = CDate("31/12/" & qAnio)
    dFechaH = CDate("30/11/" & qAnio)
Else
    dFechaD = CDate("01/" & qMes & "/" & qAnio)
    dFechaH = ultimoDiaDelMes(dFechaD)
End If


'cConsul = "SELECT * FROM ASIENTOS WHERE ACTIVO=1 AND fecha>=" & ssFecha(dFechaD) & " and fecha<= " & ssFecha(dFechaH) & " order by fecha"

'cConsul = "select a.fecha,m.cuenta,c.descripcion,m.debe,m.haber from (asientos a inner join mayor m on a.idasiento=m.idasiento) inner join cuentas c on m.cuenta=c.cuenta where a.activo=1 and fecha>=" & ssFecha(dFechaD) & " and fecha<= " & ssFecha(dFechaH) & " order by fecha"

cConsul = "select ('PERIODO  ' + CAST(datepart(MONTH,a.fecha) AS VARCHAR) + '/' + CAST(datepart(year,a.fecha) AS VARCHAR)) as PERIODO,m.cuenta AS CUENTA,c.descripcion AS DESCRIPCION,sum(round(m.debe,2)) as DEBE,SUM(round(m.haber,2)) AS HABER,datepart(month,a.fecha) as MES,datepart(year,a.fecha) as anio from (asientos a inner join mayor m on a.idasiento=m.idasiento) inner join cuentas c on m.cuenta=c.cuenta where  a.activo=1 and fecha>=" & ssFecha(dFechaD) & " and fecha<= " & ssFecha(dFechaH) & " and a.origen<>'E' GROUP BY datepart(month,a.fecha),datepart(year,a.fecha) ,m.cuenta,c.descripcion order by mes,anio,m.cuenta"


LlenarGrilla gSumariza, cConsul, False

With gSumariza
    If .rows > 1 Then
        .ColWidth(0) = 2500
        .ColWidth(1) = 1500
        .ColWidth(2) = 3500
        .ColWidth(3) = 2000
        .ColWidth(4) = 2000
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        dPeriodoActual = ""
        dCuentaActual = 1
otraves:
        For i = dCuentaActual To .rows - 1
            
            mCoeficiente = s2n(txtIndice, 4)
            If mCoeficiente = 0 Then mCoeficiente = 1
            
            mDebeActual = s2n(.TextMatrix(i, 3))
            mHaberActual = s2n(.TextMatrix(i, 4))
            
            If mDebeActual = mHaberActual Then
                mDebeActual = 0
                mHaberActual = 0
            ElseIf mDebeActual > mHaberActual Then
                mDebeActual = mDebeActual - mHaberActual
                mHaberActual = 0
            ElseIf mHaberActual > mDebeActual Then
                mHaberActual = mHaberActual - mDebeActual
                mDebeActual = 0
            End If
            
            .TextMatrix(i, 3) = s2n(mDebeActual)
            .TextMatrix(i, 4) = s2n(mHaberActual)
            
            mDebeNuevo = s2n(mDebeActual * mCoeficiente)
            mHaberNuevo = s2n(mHaberActual * mCoeficiente)
            
            mDebeAjuste = s2n(mDebeNuevo - mDebeActual)
            mHaberAjuste = s2n(mHaberNuevo - mHaberActual)
            
            If dPeriodoActual <> .TextMatrix(i, 0) Then
                g_add3 .TextMatrix(i, 0), .TextMatrix(i, 1), .TextMatrix(i, 2), s2n(mDebeNuevo), s2n(mHaberNuevo)
                g_add4 .TextMatrix(i, 0), .TextMatrix(i, 1), .TextMatrix(i, 2), s2n(mDebeAjuste), s2n(mHaberAjuste)
            Else
                g_add3 "", .TextMatrix(i, 1), .TextMatrix(i, 2), s2n(mDebeNuevo), s2n(mHaberNuevo)
                g_add4 "", .TextMatrix(i, 1), .TextMatrix(i, 2), s2n(mDebeAjuste), s2n(mHaberAjuste)
            End If
            
            If dPeriodoActual <> .TextMatrix(i, 0) Then
                dPeriodoActual = .TextMatrix(i, 0)
                .cell(flexcpFontBold, i, 0, i, 0) = True
                dCuentaActual = i + 1
                GoTo otraves
            Else
                .TextMatrix(i, 0) = ""
            End If
            
            
        Next
    Else
        MsgBox "No hay datos en la consulta.", vbInformation
    End If
End With


relojito False
End Sub

Private Function Nombre_Mes(f As Date, Optional limitador As String = "") As String
Dim Mes, Anio, mes_Letra As String
    Anio = Mid(Year(f), 3, 2)
    Mes = Month(f)
    Select Case Mes
        Case 1:
            mes_Letra = "Enero"
        Case 2:
            mes_Letra = "Febrero"
        Case 3:
            mes_Letra = "Marzo"
        Case 4:
            mes_Letra = "Abril"
        Case 5:
            mes_Letra = "Mayo"
        Case 6:
            mes_Letra = "Junio"
        Case 7:
            mes_Letra = "Julio"
        Case 8:
            mes_Letra = "Agosto"
        Case 9:
            mes_Letra = "Septiembre"
        Case 10:
            mes_Letra = "Octubre"
        Case 11:
            mes_Letra = "Noviembre"
        Case 12:
            mes_Letra = "Diciembre"
    End Select
    Nombre_Mes = mes_Letra & " / " & Anio
    If limitador > "" Then Nombre_Mes = mes_Letra & limitador & Anio
End Function


Private Function g_ini2()

With gSumariza3
    .rows = 1
    .cols = 0
    .cols = 5
    .TextMatrix(0, 0) = " PERIODO "
    .TextMatrix(0, 1) = " CUENTA "
    .TextMatrix(0, 2) = " DESCRIPCION "
    .TextMatrix(0, 3) = " DEBE "
    .TextMatrix(0, 4) = " HABER "
    .ColWidth(0) = 2500
    .ColWidth(1) = 1500
    .ColWidth(2) = 3500
    .ColWidth(3) = 2000
    .ColWidth(4) = 2000
End With

With gSumariza4
    .rows = 1
    .cols = 0
    .cols = 5
    .TextMatrix(0, 0) = " PERIODO "
    .TextMatrix(0, 1) = " CUENTA "
    .TextMatrix(0, 2) = " DESCRIPCION "
    .TextMatrix(0, 3) = " DEBE "
    .TextMatrix(0, 4) = " HABER "
    .ColWidth(0) = 2500
    .ColWidth(1) = 1500
    .ColWidth(2) = 3500
    .ColWidth(3) = 2000
    .ColWidth(4) = 2000
End With
End Function


Private Function g_add2(a As String, b As String, C As String, d As String, e As String)
Dim rou As Long
With gSumariza
    .AddItem " "
    rou = .rows - 1
    .TextMatrix(rou, 0) = a
    .TextMatrix(rou, 1) = b
    .TextMatrix(rou, 2) = C
    'If s2n(d) = 0 Then d = ""
    .TextMatrix(rou, 3) = d
    'If s2n(e) = 0 Then e = ""
    .TextMatrix(rou, 4) = e
'    If a > "" And C > "" And b = "" Then
'        .cell(flexcpFontBold, rou, 0, rou, 4) = True
'    End If
    If a > "" Then
        .cell(flexcpFontBold, rou, 0, rou, 0) = True
    End If

End With
End Function

Private Function g_add3(a As String, b As String, C As String, d As String, e As String)
Dim rou As Long
With gSumariza3
    .AddItem " "
    rou = .rows - 1
    .TextMatrix(rou, 0) = a
    .TextMatrix(rou, 1) = b
    .TextMatrix(rou, 2) = C
    'If s2n(d) = 0 Then d = ""
    .TextMatrix(rou, 3) = d
    'If s2n(e) = 0 Then e = ""
    .TextMatrix(rou, 4) = e
'    If a > "" And C > "" And b = "" Then
'        .cell(flexcpFontBold, rou, 0, rou, 4) = True
'    End If
    If a > "" Then
        .cell(flexcpFontBold, rou, 0, rou, 0) = True
    End If
    
End With
End Function

Private Function g_add4(a As String, b As String, C As String, d As String, e As String)
Dim rou As Long
With gSumariza4
    .AddItem " "
    rou = .rows - 1
    .TextMatrix(rou, 0) = a
    .TextMatrix(rou, 1) = b
    .TextMatrix(rou, 2) = C
'    If s2n(d) = 0 Then d = ""
    .TextMatrix(rou, 3) = d
'    If s2n(e) = 0 Then e = ""
    .TextMatrix(rou, 4) = e
'    If a > "" And C > "" And b = "" Then
'        .cell(flexcpFontBold, rou, 0, rou, 4) = True
'    End If
    If a > "" Then
        .cell(flexcpFontBold, rou, 0, rou, 0) = True
    End If
    
End With
End Function

Private Sub cmdQuitar_Click()
If gSumariza4.Row > 0 Then
    gSumariza4.RemoveItem gSumariza4.Row
End If
End Sub





Private Sub cmdVista_Click()
Dim AcumuloHaber As Double, AcumuloDebe As Double
Dim sCuentaAjuste As String
sCuentaAjuste = sSinNull(obtenerDeSQL("select cuenta from cuentasparam where codigo ='API'"))
If sCuentaAjuste = "" Then
    MsgBox "No existe cuenta contable definida para el asiento.", vbInformation
    Exit Sub
End If

With gSumariza4
    AcumuloDebe = 0
    AcumuloHaber = 0
    For i = 1 To .rows - 1
        AcumuloDebe = AcumuloDebe + s2n(.TextMatrix(i, 3))
        AcumuloHaber = AcumuloHaber + s2n(.TextMatrix(i, 4))
    Next
End With

If AcumuloDebe >= AcumuloHaber Then
    AcumuloHaber = s2n(AcumuloDebe - AcumuloHaber)
    AcumuloDebe = 0
ElseIf AcumuloHaber > AcumuloDebe Then
    AcumuloDebe = s2n(AcumuloHaber - AcumuloDebe)
    AcumuloHaber = 0
End If

frmAsientoManual.VistaPrevia gSumariza4, "Ajuste por inflacion", dtFechaAsiento, sCuentaAjuste, AcumuloDebe, AcumuloHaber


End Sub

Private Sub Form_Load()
dtMes = Date
dtAnio = Date
Ver
dtFechaAsiento = Date
ucXls1.ini gSumariza
ucXls2.ini gSumariza3
ucXls3.ini gSumariza4
End Sub

Private Sub chkAnual_Click()
If chkAnual Then
    dtMes.enabled = False
Else
    dtMes.enabled = True
End If
Ver
End Sub

Private Sub cmdGuardar_Click()
Dim AjusteGuardar As New AjusteInflacion
If AjusteGuardar.Guardar(qMes(), qAnio(), s2n(txtIndice, 4), chkAnual) Then
    MsgBox "Guardado", vbInformation
End If
End Sub

Private Sub dtAnio_Change()
dtMes = dtAnio
Ver
End Sub

Private Sub dtMes_Change()
dtAnio = dtMes
Ver
End Sub


Private Sub gSumariza4_KeyPress(KeyAscii As Integer)
If KeyAscii < 0 Then
    cmdQuitar_Click
End If
End Sub

Private Sub txtIndice_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNum(KeyAscii, True, False)
End Sub

Private Function Ver()
Dim AjusteVer As New AjusteInflacion
txtIndice = AjusteVer.buscar(qMes(), qAnio(), chkAnual)
End Function

Private Function qMes() As Integer
qMes = Month(dtMes)
End Function

Private Function qAnio() As Integer
qAnio = Year(dtAnio)
End Function



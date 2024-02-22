VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmLibroDiario 
   Caption         =   "Libro Diario"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11310
   Icon            =   "FrmLibroDiario.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9090
   ScaleWidth      =   11310
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdImpLibro 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir &Libro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6990
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8490
      Width           =   1380
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
      Left            =   5895
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8490
      Width           =   975
   End
   Begin Gestion.ucXls ucXls1 
      Height          =   810
      Left            =   75
      TabIndex        =   9
      Top             =   8130
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1429
   End
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Height          =   6525
      Left            =   180
      TabIndex        =   15
      Top             =   1350
      Width           =   10950
      _cx             =   19315
      _cy             =   11509
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
   Begin VB.ComboBox cmbejercicio 
      Height          =   315
      Left            =   7425
      TabIndex        =   2
      Top             =   120
      Width           =   1290
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
      Left            =   8835
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8475
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
      Left            =   7275
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   555
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
      Left            =   9945
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8475
      Width           =   975
   End
   Begin VB.TextBox txtasientod 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   615
      Width           =   1335
   End
   Begin VB.TextBox txtasientoh 
      Height          =   285
      Left            =   4845
      TabIndex        =   4
      Top             =   615
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker fechadesde 
      Height          =   300
      Left            =   1800
      TabIndex        =   0
      Top             =   127
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   529
      _Version        =   393216
      Format          =   68943873
      CurrentDate     =   38052
   End
   Begin MSComCtl2.DTPicker fechahasta 
      Height          =   300
      Left            =   4830
      TabIndex        =   1
      Top             =   120
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   529
      _Version        =   393216
      Format          =   68943873
      CurrentDate     =   38052
   End
   Begin VSFlex7LCtl.VSFlexGrid Grilla2 
      Height          =   6540
      Left            =   2265
      TabIndex        =   22
      Top             =   1290
      Visible         =   0   'False
      Width           =   8115
      _cx             =   14314
      _cy             =   11536
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
      GridColor       =   8421504
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
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
      Cols            =   6
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
   Begin VB.ComboBox CmbTipoHoja 
      Height          =   315
      Left            =   5820
      TabIndex        =   20
      Top             =   8520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Hoja"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   4560
      TabIndex        =   21
      Top             =   8535
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00800000&
      Height          =   555
      Left            =   4230
      Top             =   8415
      Visible         =   0   'False
      Width           =   3990
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Haber"
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
      Height          =   300
      Left            =   8790
      TabIndex        =   18
      Top             =   1020
      Width           =   1035
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Debe"
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
      Height          =   300
      Left            =   7395
      TabIndex        =   17
      Top             =   1020
      Width           =   1035
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Comprobante"
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
      Height          =   300
      Left            =   5850
      TabIndex        =   16
      Top             =   1020
      Width           =   1470
   End
   Begin VB.Label Label3 
      Caption         =   "Ejercicio:"
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
      Left            =   6480
      TabIndex        =   14
      Top             =   150
      Width           =   1395
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
      Left            =   285
      TabIndex        =   13
      Top             =   150
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
      Left            =   3375
      TabIndex        =   12
      Top             =   150
      Width           =   615
   End
   Begin VB.Label Label11 
      Caption         =   "Desde Asiento:"
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
      TabIndex        =   11
      Top             =   630
      Width           =   1395
   End
   Begin VB.Label Label2 
      Caption         =   "Hasta Asiento:"
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
      Left            =   3375
      TabIndex        =   10
      Top             =   630
      Width           =   1500
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00800000&
      Height          =   8010
      Left            =   60
      Top             =   45
      Width           =   11205
   End
End
Attribute VB_Name = "FrmLibroDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAceptar_Click()
Dim rsAS As New ADODB.Recordset
Dim primero As Boolean
Dim Asiento As Long
Dim r As Long
Dim TotDebe As Variant
Dim TotdebeTemp As Variant
Dim TotHaber As Variant
Dim TotHaberTemp As Variant
grilla.rows = 0
    rsAS.Open "SELECT Asientos.NroAsiento, Asientos.Fecha, Asientos.Activo, MAYOR.Cuenta, MAYOR.Debe, MAYOR.Haber,MAYOR.comprobante ,Asientos.Concepto, CUENTAS.DESCRIPCION " _
    & "FROM (Asientos INNER JOIN MAYOR ON Asientos.idAsiento = MAYOR.idAsiento) INNER JOIN CUENTAS ON MAYOR.Cuenta = CUENTAS.Cuenta " _
    & "WHERE (Asientos.NroAsiento >=" & Val(txtasientod) & " And Asientos.NroAsiento<=" & Val(txtasientoh) & ") AND (Asientos.Fecha>=" & ssFecha(FechaDesde) & " And Asientos.Fecha<=" & ssFecha(FechaHasta) & ") AND Asientos.Activo=1 order by fecha asc,nroasiento asc", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not rsAS.EOF Then
        primero = True
        Asiento = rsAS!NroAsiento
        r = 0
        TotDebe = 0
        TotHaber = 0
        TotdebeTemp = 0
        TotHaberTemp = 0
   '     grilla.AddItem Chr(9) & Chr(9) & "COMPROBANTE" & Chr(9) & "DEBE" & Chr(9) & "HABER"
'        grilla.cell(flexcpFontBold, grilla.rows - 1, 0, grilla.rows - 1, 4) = True
        r = r + 1
'        grilla.cell(flexcpForeColor, r - 1, 0) = &H80&
'        grilla.cell(flexcpBackColor, r - 1, 0) = &HE0E0E0
'        grilla.cell(flexcpForeColor, r - 1, 1) = &H80&
'        grilla.cell(flexcpBackColor, r - 1, 1) = &HE0E0E0
'        grilla.cell(flexcpForeColor, r - 1, 2) = &H80&
'        grilla.cell(flexcpBackColor, r - 1, 2) = &HE0E0E0
'        grilla.cell(flexcpForeColor, r - 1, 3) = &H80&
'        grilla.cell(flexcpBackColor, r - 1, 3) = &HE0E0E0
        
        Do While Not rsAS.EOF
            
            If primero Then
                If grilla.rows > 1 Then
                   grilla.AddItem ""
                End If
                grilla.AddItem "Asiento  " & rsAS!NroAsiento & Chr(9) & rsAS!fecha & "  " & rsAS!concepto
                grilla.cell(flexcpFontBold, grilla.rows - 1, 0, grilla.rows - 1, 3) = True
                r = r + 1
'                grilla.cell(flexcpForeColor, r - 1, 0) = &H80&
'                grilla.cell(flexcpBackColor, r - 1, 0) = &HE0E0E0
'                grilla.cell(flexcpForeColor, r - 1, 1) = &H80&
'                grilla.cell(flexcpBackColor, r - 1, 1) = &HE0E0E0
'                grilla.cell(flexcpForeColor, r - 1, 2) = &H80&
'                grilla.cell(flexcpBackColor, r - 1, 2) = &HE0E0E0
'                grilla.cell(flexcpForeColor, r - 1, 3) = &H80&
'                grilla.cell(flexcpBackColor, r - 1, 3) = &HE0E0E0
                
                grilla.AddItem rsAS!cuenta & Chr(9) & rsAS!descripcion & Chr(9) & Trim(rsAS!comprobante) & Chr(9) & Format$(rsAS!Debe, "standard") & Chr(9) & Format$(rsAS!haber, "standard")
                TotDebe = s2n(TotDebe, 2) + s2n(rsAS!Debe, 2)
                TotHaber = s2n(TotHaber, 2) + s2n(rsAS!haber, 2)
                TotdebeTemp = s2n(TotdebeTemp, 2) + s2n(rsAS!Debe, 2)
                TotHaberTemp = s2n(TotHaberTemp, 2) + s2n(rsAS!haber, 2)
                r = r + 1
'                grilla.cell(flexcpForeColor, r - 1, 0) = vbBlack
'                grilla.cell(flexcpBackColor, r - 1, 0) = vbWhite
'                grilla.cell(flexcpForeColor, r - 1, 1) = vbBlack
'                grilla.cell(flexcpBackColor, r - 1, 1) = vbWhite
'                grilla.cell(flexcpForeColor, r - 1, 2) = vbBlack
'                grilla.cell(flexcpBackColor, r - 1, 2) = vbWhite
'                grilla.cell(flexcpForeColor, r - 1, 3) = vbBlack
'                grilla.cell(flexcpBackColor, r - 1, 3) = vbWhite
                primero = False
            Else
'                grilla.cell(flexcpForeColor, r - 1, 0) = vbBlack
'                grilla.cell(flexcpBackColor, r - 1, 0) = vbWhite
'                grilla.cell(flexcpForeColor, r - 1, 1) = vbBlack
'                grilla.cell(flexcpBackColor, r - 1, 1) = vbWhite
'                grilla.cell(flexcpForeColor, r - 1, 2) = vbBlack
'                grilla.cell(flexcpBackColor, r - 1, 2) = vbWhite
'                grilla.cell(flexcpForeColor, r - 1, 3) = vbBlack
'                grilla.cell(flexcpBackColor, r - 1, 3) = vbWhite
                grilla.AddItem rsAS!cuenta & Chr(9) & rsAS!descripcion & Chr(9) & Trim(rsAS!comprobante) & Chr(9) & Format$(rsAS!Debe, "standard") & Chr(9) & Format$(rsAS!haber, "standard")
                TotDebe = s2n(TotDebe, 2) + s2n(rsAS!Debe, 2)
                TotHaber = s2n(TotHaber, 2) + s2n(rsAS!haber, 2)
                TotdebeTemp = s2n(TotdebeTemp, 2) + s2n(rsAS!Debe, 2)
                TotHaberTemp = s2n(TotHaberTemp, 2) + s2n(rsAS!haber, 2)
                r = r + 1
            End If
            rsAS.MoveNext
            If Not rsAS.EOF Then
                If Asiento = rsAS!NroAsiento Then
                    primero = False
                Else
                    Asiento = rsAS!NroAsiento
                    grilla.AddItem Chr(9) & Chr(9) & "Total" & Chr(9) & Format$(TotdebeTemp, "standard") & Chr(9) & Format$(TotHaberTemp, "standard")
                    grilla.cell(flexcpFontBold, grilla.rows - 1, 0, grilla.rows - 1, 4) = True
                    TotdebeTemp = 0
                    TotHaberTemp = 0
                    primero = True
                End If
            Else
                grilla.AddItem Chr(9) & Chr(9) & "Total" & Chr(9) & Format$(TotdebeTemp, "standard") & Chr(9) & Format$(TotHaberTemp, "standard")
                grilla.cell(flexcpFontBold, grilla.rows - 1, 0, grilla.rows - 1, 4) = True
            End If
        Loop
        grilla.AddItem ""
        grilla.AddItem "TOTAL" & Chr(9) & Chr(9) & Chr(9) & Format$(TotDebe, "standard") & Chr(9) & Format$(TotHaber, "standard")
        grilla.cell(flexcpFontBold, grilla.rows - 1, 0, grilla.rows - 1, 4) = True
        r = r + 1
'        grilla.cell(flexcpForeColor, r - 1, 0) = &H80&
'        grilla.cell(flexcpBackColor, r - 1, 0) = &HE0E0E0
'        grilla.cell(flexcpForeColor, r - 1, 1) = &H80&
'        grilla.cell(flexcpBackColor, r - 1, 1) = &HE0E0E0
'        grilla.cell(flexcpForeColor, r - 1, 2) = &H80&
'        grilla.cell(flexcpBackColor, r - 1, 2) = &HE0E0E0
'        grilla.cell(flexcpForeColor, r - 1, 3) = &H80&
'        grilla.cell(flexcpBackColor, r - 1, 3) = &HE0E0E0
    Else
        MsgBox "No hay registros pára mostrar con esos parametros", 48, "Atencion"
    End If

rsAS.Close
Set rsAS = Nothing
CmdImprimir.enabled = True
End Sub

Private Sub TranspasoGrilla()
Dim x As Long, cant As Long
Dim TotDebe As Double, TotHaber As Double
Dim ACabecera As Boolean, ATotal As Boolean
Grilla2.rows = 0
cant = 1
InicioGrilla2
For x = 0 To grilla.rows - 1
   'Falta terminar de analizar por que estira el font en la previsualizacion
   'If cant <= CmbTipoHoja.ItemData(CmbTipoHoja.ListIndex) Then                                     'And (Mid(grilla.TextMatrix(x, 0), 1, 7) <> "Asiento" Or    Mid(grilla.TextMatrix(x, 2), 1, 7) <> "Total") Then
    If cant < 56 Then
        Grilla2.AddItem grilla.TextMatrix(x, 0) & Chr(9) & grilla.TextMatrix(x, 1) & Chr(9) & grilla.TextMatrix(x, 2) & Chr(9) & grilla.TextMatrix(x, 3) & Chr(9) & Format$(grilla.TextMatrix(x, 4), "standard")
        
        If Mid(Grilla2.TextMatrix(Grilla2.rows - 1, 0), 1, 7) = "Asiento" Or Mid(Grilla2.TextMatrix(Grilla2.rows - 1, 2), 1, 7) = "Total" Or Mid(Grilla2.TextMatrix(Grilla2.rows - 1, 2), 1, 10) = "TRANSPORTE" Then
            Grilla2.cell(flexcpFontBold, Grilla2.rows - 1, 0, Grilla2.rows - 1, 4) = True
            TotDebe = 0
            TotHaber = 0
        End If
'        ESTO ES PARA PROBAR LA IMPRESION
'        If Trim(Grilla2.TextMatrix(Grilla2.rows - 1, 0)) = "Asiento  464" Then
'            MsgBox "s"
'        End If
        cant = cant + 1
    
        If Grilla2.TextMatrix(Grilla2.rows - 1, 2) <> "Total" And Grilla2.TextMatrix(Grilla2.rows - 1, 3) <> "" Then
            TotDebe = TotDebe + CDbl(Grilla2.TextMatrix(Grilla2.rows - 1, 3))
            TotHaber = TotHaber + CDbl(Grilla2.TextMatrix(Grilla2.rows - 1, 4))
        Else
            TotDebe = 0
            TotHaber = 0
        End If
    Else
        If Mid(grilla.TextMatrix(x, 0), 1, 7) = "Asiento" Then
            Grilla2.AddItem " "
            Grilla2.AddItem " "
            Grilla2.AddItem " "
            cant = 1
            Grilla2.AddItem grilla.TextMatrix(x, 0) & Chr(9) & grilla.TextMatrix(x, 1) & Chr(9) & grilla.TextMatrix(x, 2) & Chr(9) & grilla.TextMatrix(x, 3) & Chr(9) & Format$(grilla.TextMatrix(x, 4), "standard")
            Grilla2.cell(flexcpFontBold, Grilla2.rows - 1, 0, Grilla2.rows - 1, 4) = True
            cant = cant + 1
        Else
            
            If Trim(Mid(grilla.TextMatrix(x, 2), 1, 5)) = "Total" Then
                Grilla2.AddItem grilla.TextMatrix(x, 0) & Chr(9) & grilla.TextMatrix(x, 1) & Chr(9) & grilla.TextMatrix(x, 2) & Chr(9) & grilla.TextMatrix(x, 3) & Chr(9) & Format$(grilla.TextMatrix(x, 4), "standard")
                Grilla2.cell(flexcpFontBold, Grilla2.rows - 1, 0, Grilla2.rows - 1, 4) = True
                Grilla2.AddItem " "
                Grilla2.AddItem " "
                cant = 1
                TotDebe = 0
                TotHaber = 0
            Else
                If Trim(grilla.TextMatrix(x, 3)) = "" Then
'                    Grilla2.AddItem grilla.TextMatrix(x, 0) & Chr(9) & grilla.TextMatrix(x, 1) & Chr(9) & grilla.TextMatrix(x, 2) & Chr(9) & grilla.TextMatrix(x, 3) & Chr(9) & Format$(grilla.TextMatrix(x, 4), "standard") & Chr(9) & cant
                    Grilla2.AddItem " "
                    Grilla2.AddItem " "
                    Grilla2.AddItem " "
                    cant = 1
                Else
                    Grilla2.AddItem grilla.TextMatrix(x, 0) & Chr(9) & grilla.TextMatrix(x, 1) & Chr(9) & grilla.TextMatrix(x, 2) & Chr(9) & grilla.TextMatrix(x, 3) & Chr(9) & Format$(grilla.TextMatrix(x, 4), "standard")
                    TotDebe = TotDebe + CDbl(Grilla2.TextMatrix(Grilla2.rows - 1, 3))
                    TotHaber = TotHaber + CDbl(Grilla2.TextMatrix(Grilla2.rows - 1, 4))
                    
                    If TotDebe <> 0 Or TotHaber <> 0 Then
                        If grilla.TextMatrix(x + 1, 2) = "Total" Then
                            x = x + 1
                            Grilla2.AddItem grilla.TextMatrix(x, 0) & Chr(9) & grilla.TextMatrix(x, 1) & Chr(9) & grilla.TextMatrix(x, 2) & Chr(9) & grilla.TextMatrix(x, 3) & Chr(9) & Format$(grilla.TextMatrix(x, 4), "standard")
                            Grilla2.cell(flexcpFontBold, Grilla2.rows - 1, 0, Grilla2.rows - 1, 4) = True
                            Grilla2.AddItem " "
                            cant = 1
                        Else
                            Grilla2.AddItem Chr(9) & Chr(9) & "TRANSPORTE" & Chr(9) & Format$(TotDebe, "standard") & Chr(9) & Format$(TotHaber, "standard")
                            Grilla2.cell(flexcpFontBold, Grilla2.rows - 1, 0, Grilla2.rows - 1, 4) = True
                            Grilla2.AddItem " "
                            cant = 1
                            Grilla2.AddItem Chr(9) & Chr(9) & "TRANSPORTE" & Chr(9) & Format$(TotDebe, "standard") & Chr(9) & Format$(TotHaber, "standard")
                            Grilla2.cell(flexcpFontBold, Grilla2.rows - 1, 0, Grilla2.rows - 1, 4) = True
                            cant = cant + 1
                        End If
                    End If
                End If
            End If
            
        End If
    End If
Next
End Sub

Private Sub cmdCancelar_Click()
    LimpioControles
End Sub

Private Sub CmdImpLibro_Click()
Dim sTablaTempLD As String
Dim x As Long
'If CmbTipoHoja <> "" Then
    TranspasoGrilla
    Imprimir Grilla2, 2
'End If
'sTablaTempLD = TablaTempCrear(tt_LibroDiario)
'For x = 0 To Grilla2.rows - 1
'    Sql = "insert into " & sTablaTempLD & " (columna1,columna2,columna3,columna4,columna5) values ('" & Grilla2.TextMatrix(x, 0) & "','" & Grilla2.TextMatrix(x, 1) & "','" & Grilla2.TextMatrix(x, 2) & "','" & Grilla2.TextMatrix(x, 3) & "','" & Grilla2.TextMatrix(x, 4) & "' )"
'    DataEnvironment1.Sistema.Execute Sql
'Next
End Sub

Sub Imprimir(ByVal xObj As Object, Ngrilla As String)
xObj.GridLines = flexGridNone
xObj.GridLinesFixed = flexGridNone

FrmImpresiones.VSPrinter.StartDoc
FrmImpresiones.VSPrinter.PhysicalPage = True
FrmImpresiones.VSPrinter.Orientation = orPortrait
'If Ngrilla = 2 Then
'    Select Case CmbTipoHoja.Text
'    Case "A4"
'        FrmImpresiones.VSPrinter.PaperSize = pprA4
'        FrmImpresiones.VSPrinter_BeforeUserPage (pprA4)
'        FrmImpresiones.VSPrinter.PrintDialog
'    Case "Carta"
'        FrmImpresiones.VSPrinter.PaperSize = pprLetter
'        FrmImpresiones.VSPrinter_BeforeUserPage (pprLetter)
'    Case "Oficio"
'        FrmImpresiones.VSPrinter.PaperSize = pprLegal
'        FrmImpresiones.VSPrinter_BeforeUserPage (pprLegal)
'    End Select
'Else
    FrmImpresiones.VSPrinter.PaperSize = pprA4
'End If
FrmImpresiones.VSPrinter.Preview = True
FrmImpresiones.VSPrinter.Font.Name = "Courier"
FrmImpresiones.VSPrinter.FontSize = 10
FrmImpresiones.VSPrinter.Header = VerParametro(BS_NOMBRE_EMPRESA_CORTO) & "                              Hoja : " & "%d" & vbLf & "Libro Diario" & vbLf & "Rango de Fechas   : " & FechaDesde & " - " & FechaHasta & vbLf & "Rango de Asientos : " & txtasientod & "  -  " & txtasientoh & vbLf & "                               Comprobante    Debe   Haber"
FrmImpresiones.VSPrinter.FontSize = 10

If xObj.rows > 1 Then
   FrmImpresiones.VSPrinter.TextAlign = taLeftop
   FrmImpresiones.VSPrinter.RenderControl = xObj.hWnd
End If
FrmImpresiones.VSPrinter.Zoom = 100
FrmImpresiones.VSPrinter.EndDoc

FrmImpresiones.Show

End Sub

Private Sub cmdImprimir_Click()
    Imprimir grilla, 1
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Sub InicioGrilla()
    grilla.FontSize = 8
    grilla.ColWidth(0) = 1400
    grilla.ColWidth(1) = 4300
    grilla.ColWidth(2) = 1700
    grilla.ColWidth(3) = 1200
    grilla.ColWidth(4) = 1200
    grilla.ColAlignment(0) = flexAlignLeftCenter
    grilla.ColAlignment(1) = flexAlignLeftCenter
    grilla.ColAlignment(2) = flexAlignLeftCenter
    grilla.ColAlignment(3) = flexAlignRightCenter
    grilla.ColAlignment(4) = flexAlignRightCenter
    grilla.ColDataType(3) = flexDTCurrency
    grilla.ColDataType(4) = flexDTCurrency
End Sub
Sub InicioGrilla2()
    Grilla2.FontSize = 8
    Grilla2.ColWidth(0) = 1400
    Grilla2.ColWidth(1) = 4300
    Grilla2.ColWidth(2) = 1700
    Grilla2.ColWidth(3) = 1200
    Grilla2.ColWidth(4) = 1200
    Grilla2.ColAlignment(0) = flexAlignLeftCenter
    Grilla2.ColAlignment(1) = flexAlignLeftCenter
    Grilla2.ColAlignment(2) = flexAlignLeftCenter
    Grilla2.ColAlignment(3) = flexAlignRightCenter
    Grilla2.ColAlignment(4) = flexAlignRightCenter
    Grilla2.ColDataType(3) = flexDTCurrency
    Grilla2.ColDataType(4) = flexDTCurrency
End Sub

Sub LimpioControles()
Dim rs As New ADODB.Recordset

    FechaDesde = Date
    FechaHasta = Date
    txtasientod = "0"
    txtasientoh = "999999"
    CargaComboEj cmbejercicio, "ejercicio", "denominacion", "idejercicio", ""
    rs.Open "Select denominacion from ejercicio where activo=1", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    If Not rs.EOF Then
        cmbejercicio.Text = rs!denominacion
    End If
    rs.Close
    Set rs = Nothing
    grilla.rows = 0
    Grilla2.rows = 0
    InicioGrilla
End Sub
Sub CargaComboEj(Combo As Object, tabla As String, campo As String, Bound As String, wer As String)

Dim rsCargacombo As New ADODB.Recordset
Dim sqlstrCC As String
Dim i As Long
    If Bound <> "" Then
        sqlstrCC = "Select " + campo + " as NN," + Bound + " from " + tabla
    Else
        sqlstrCC = "Select " + campo + " as NN" + Bound + " from " + tabla
    End If
    If wer <> "" Then
        sqlstrCC = sqlstrCC + " and " + wer
    End If
    sqlstrCC = sqlstrCC + " order by " + campo
    rsCargacombo.Open sqlstrCC, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    Combo.clear
    If Not rsCargacombo.EOF And Not rsCargacombo.BOF Then
        rsCargacombo.MoveFirst
        i = 0
        While Not rsCargacombo.EOF
            Combo.AddItem Trim(rsCargacombo.Fields("NN"))
            Combo.ItemData(i) = i
            i = i + 1
            rsCargacombo.MoveNext
        Wend
    End If
End Sub

Private Sub Form_Load()
    LimpioControles

' SACARLO UNA VEZ TERMINADO
'fechadesde.Value = #1/1/2004#
'fechahasta.Value = #12/31/2004#
'***************************************

    CmbTipoHoja.AddItem "A4"
    CmbTipoHoja.ItemData(CmbTipoHoja.NewIndex) = 57
    CmbTipoHoja.AddItem "Carta"
    CmbTipoHoja.ItemData(CmbTipoHoja.NewIndex) = 54
    CmbTipoHoja.AddItem "Oficio"
    CmbTipoHoja.ItemData(CmbTipoHoja.NewIndex) = 71
    ucXls1.ini grilla, App.Path & "\LibroDiario" & Trim(cmbejercicio.Text)
End Sub

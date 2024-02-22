VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmSueldosAdelantos 
   Caption         =   "Pago de Anticipos, Prestamos y Sueldos"
   ClientHeight    =   8580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9855
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   9855
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraContable 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   690
      TabIndex        =   38
      Top             =   900
      Width           =   7455
      Begin VB.OptionButton optPres 
         Caption         =   "PRESTAMO"
         Height          =   360
         Left            =   2895
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   105
         Width           =   2640
      End
      Begin VB.OptionButton optRemu 
         Caption         =   "REMUNERACIONES A PAGAR"
         Height          =   345
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Value           =   -1  'True
         Width           =   2640
      End
      Begin VB.TextBox txtCuentaContable 
         Enabled         =   0   'False
         Height          =   360
         Left            =   5625
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   105
         Width           =   1650
      End
   End
   Begin VB.Frame fraOP 
      Caption         =   "Orden de Pago"
      Height          =   1860
      Left            =   60
      TabIndex        =   30
      Top             =   6030
      Width           =   9735
      Begin Gestion.ucFecha uFechaCheque 
         Height          =   300
         Left            =   6315
         TabIndex        =   19
         Top             =   855
         Width           =   1035
         _ExtentX        =   2117
         _ExtentY        =   529
         FechaInit       =   4
      End
      Begin Gestion.uCtaBanco uCtaBanc 
         Height          =   330
         Left            =   2925
         TabIndex        =   21
         Top             =   1320
         Width           =   6675
         _ExtentX        =   9234
         _ExtentY        =   582
      End
      Begin Gestion.ucCoDe uCaja 
         Height          =   330
         Left            =   2925
         TabIndex        =   16
         Top             =   345
         Width           =   5220
         _ExtentX        =   9208
         _ExtentY        =   609
         CodigoWidth     =   800
      End
      Begin Gestion.uNum numTotEfec 
         Height          =   330
         Left            =   870
         TabIndex        =   15
         Top             =   345
         Width           =   1185
         _ExtentX        =   2275
         _ExtentY        =   582
      End
      Begin Gestion.uNum numTotCheq 
         Height          =   330
         Left            =   870
         TabIndex        =   17
         Top             =   840
         Width           =   1185
         _ExtentX        =   2275
         _ExtentY        =   582
      End
      Begin Gestion.uNum numTotDepo 
         Height          =   330
         Left            =   870
         TabIndex        =   20
         Top             =   1320
         Width           =   1185
         _ExtentX        =   2275
         _ExtentY        =   582
      End
      Begin Gestion.ucCoDe uNroCheque 
         Height          =   330
         Left            =   2940
         TabIndex        =   18
         Top             =   840
         Width           =   2715
         _ExtentX        =   7699
         _ExtentY        =   582
         CodigoWidth     =   1455
         CodigoInvalido  =   0
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5715
         TabIndex        =   37
         Top             =   900
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Numero"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2220
         TabIndex        =   36
         Top             =   855
         Width           =   735
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2220
         TabIndex        =   35
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Depósito"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   180
         TabIndex        =   34
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Efectivo"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   210
         TabIndex        =   33
         Top             =   345
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cheques"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   165
         TabIndex        =   32
         Top             =   870
         Width           =   915
      End
      Begin VB.Label label26 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Caja"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2220
         TabIndex        =   31
         Top             =   390
         Width           =   675
      End
   End
   Begin VB.Frame fraPrestamo 
      Height          =   4515
      Left            =   60
      TabIndex        =   24
      Top             =   1455
      Width           =   9585
      Begin VB.TextBox txtCodTliq 
         Height          =   345
         Left            =   1665
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   285
         Width           =   1335
      End
      Begin VB.TextBox txtTliq 
         Height          =   345
         Left            =   3135
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   270
         Width           =   4545
      End
      Begin VB.CommandButton cmdTliq 
         Caption         =   "?"
         Height          =   345
         Left            =   1035
         TabIndex        =   7
         Top             =   255
         Width           =   495
      End
      Begin VB.TextBox txtMotivo 
         Height          =   375
         Left            =   1005
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   840
         Width           =   6660
      End
      Begin Gestion.uNum numTotal 
         Height          =   330
         Left            =   855
         TabIndex        =   12
         Top             =   2385
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   582
      End
      Begin Gestion.ucFecha uFecha 
         Height          =   345
         Left            =   870
         TabIndex        =   11
         Top             =   1830
         Width           =   1290
         _ExtentX        =   2223
         _ExtentY        =   609
         FechaInit       =   0
      End
      Begin VSFlex7LCtl.VSFlexGrid grilla 
         Height          =   3000
         Left            =   2445
         TabIndex        =   14
         Top             =   1335
         Width           =   5325
         _cx             =   9393
         _cy             =   5292
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
      Begin Gestion.uNum numCuotas 
         Height          =   330
         Left            =   870
         TabIndex        =   13
         Top             =   2985
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   582
         Decimales       =   0
         DecimalesCalculo=   0
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Liq"
         Height          =   330
         Index           =   1
         Left            =   135
         TabIndex        =   29
         Top             =   300
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Motivo:"
         Height          =   270
         Index           =   2
         Left            =   180
         TabIndex        =   28
         Top             =   900
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Total:"
         Height          =   270
         Index           =   8
         Left            =   225
         TabIndex        =   27
         Top             =   2445
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Inicio:"
         Height          =   270
         Index           =   9
         Left            =   240
         TabIndex        =   26
         Top             =   1935
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Cuotas:"
         Height          =   270
         Index           =   4
         Left            =   255
         TabIndex        =   25
         Top             =   3030
         Width           =   840
      End
   End
   Begin VB.TextBox txtApellido 
      Height          =   345
      Left            =   2850
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   105
      Width           =   5190
   End
   Begin VB.CommandButton cmdBuscarPersonal 
      Caption         =   "?"
      Height          =   345
      Left            =   2295
      TabIndex        =   2
      Top             =   105
      Width           =   495
   End
   Begin Gestion.ucBotonera uMenu 
      Align           =   2  'Align Bottom
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   7980
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   1058
      MsgConfirmaSalir=   ""
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   ""
      CaptionEliminar =   "&Eliminar"
   End
   Begin Gestion.ucCuit uCuil 
      Height          =   345
      Left            =   885
      TabIndex        =   1
      Top             =   90
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   609
      Enabled         =   0   'False
   End
   Begin Gestion.uNum numNumero 
      Height          =   330
      Left            =   870
      TabIndex        =   39
      Top             =   510
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   582
      Decimales       =   0
      DecimalesCalculo=   0
   End
   Begin VB.Label Label1 
      Caption         =   "CUIL"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Top             =   135
      Width           =   630
   End
   Begin VB.Label Label1 
      Caption         =   "Numero:"
      Height          =   270
      Index           =   3
      Left            =   120
      TabIndex        =   23
      Top             =   540
      Width           =   735
   End
End
Attribute VB_Name = "frmSueldosAdelantos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' AbreDBF sueldos
' carga:
'   empleado
'   anticipo o prestamo  ( 1 o varias cuotas )
'   pago: caja, cheques, transferencia
'   genera asiento ( contra remuneraciones para anticipo/ contra cuenta empleado para prestamo)
'   genera OP

Dim mRsPers As New ADODB.Recordset
Dim mRsCred As New ADODB.Recordset
Dim mRsCuot As New ADODB.Recordset
Dim mRsTLiq As New ADODB.Recordset
Dim Cajon As Long 'nro de movimiento
Dim Legajo As Long 'nro de leg de personal

Private mCtaEmpleado As String
Private midDoc       As Long

Private Const mTitCuota = "Cuota  | Periodo  | Importe    | Cobrado"
Private Enum gricolcuota
    gCUOT
    gPERI
    gIMPO
    gCOBR
End Enum

Private Sub cmdBuscarPersonal_Click()
    ' OJO ahora asume alta, ojo estoy llamando desde umenu_buscar()
    
    Dim a(), i As Long
    Dim Nro As Long
    With mRsPers
        .MoveFirst
        ReDim a(2, 0)
        a(0, 0) = "CUIL                "
        a(1, 0) = "APELLIDO                                  "
        a(2, 0) = "CUENTA PRESTAMO"
        While Not .EOF
            i = i + 1
            ReDim Preserve a(2, i)
            a(0, i) = !cuil
            a(1, i) = !apellido
            a(2, i) = sSinNull(!ctaantic)
            .MoveNext
        Wend
    End With
    With frmBuscar
        .MostrarArray (a)
        If .resultado > "" Then
            uCuil.Text = .resultado(1)
            txtApellido = .resultado(2)
            mCtaEmpleado = .resultado(3)
        End If
    End With
    
    'para traer el legajo
    With mRsPers
        .MoveFirst
        While Not .EOF
            If uCuil.Text = !cuil Then
                Legajo = !Legajo
            End If
            .MoveNext
        Wend
    End With
    
    'cargo Nro. credito
    With mRsCred
        .MoveFirst
        Nro = 0
        While Not .EOF
            If !cuil = uCuil.Text And !credito > Nro Then
                Nro = !credito
            End If
            .MoveNext
        Wend
        Nro = Nro + 1
        numNumero.num = Nro
    End With

End Sub
Private Sub cmdTliq_Click()
    Dim a(), i As Integer
    With mRsTLiq
        .MoveFirst
        ReDim a(1, 0)
        a(0, 0) = "Codigo       "
        a(1, 0) = "Denominacion                                "
        While Not .EOF
            i = i + 1
            ReDim Preserve a(1, i)
            a(0, i) = !codtip
            a(1, i) = !denom
            .MoveNext
        Wend
    End With
    With frmBuscar
        .MostrarArray (a)
        If .resultado > "" Then
            txtCodTliq = .resultado(1)
            txtTliq = .resultado(2)
        End If
    End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub

Private Sub Form_Load()
    iniDBF
    
    iniPersonal
    iniOP
    
    uMenu.init True, True, True, True, True
End Sub

Private Sub iniDBF()
    cnxAbrirDBF obtenerParametro("PathSueldos")
    
    mRsPers.Open "personal", cnxDBF, adOpenDynamic, adLockReadOnly
    mRsTLiq.Open "TipoLiq", cnxDBF, adOpenDynamic, adLockReadOnly
    
    mRsCred.Open "credito", cnxDBF, adOpenDynamic, adLockOptimistic
    mRsCuot.Open "cuotas", cnxDBF, adOpenDynamic, adLockOptimistic
End Sub


Private Sub iniPersonal()
    With grilla
        .clear
        .FixedCols = 0
        .FixedRows = 1
        .cols = 4
        .FormatString = mTitCuota
        .ColDataType(gCOBR) = flexDTBoolean
        .ColDataType(gIMPO) = flexDTDecimal
        grillaWidth grilla, Array(1000, 1000, 1111, 1111)
    End With
    numCuotas.num = 1
End Sub
Private Sub iniOP()
    uCaja.ini "select responsable from cajas where codigo = ###", "select codigo , responsable as [Descripcion                       ] from cajas"
    uCaja.codigo = 1
    
    uCtaBanc.codigo = 1

    inicheque False
    uFechaCheque.dtFecha CDate("01/01/2000")
End Sub
Private Sub inicheque(paramodi As Boolean)
    If paramodi Then
        uNroCheque.ini "select codigo from chq_comp where estado = 'T' and activo = 1 and nro  = '###' ", "select nro Numero, codigo , banco from chq_comp where activo = 1 and estado = 'T' ", True
    Else
        uNroCheque.ini "select codigo from chq_comp where estado = 'C' and activo = 1 and nro  = '###' ", "select nro Numero, codigo , banco from chq_comp where activo = 1 and estado = 'C' ", True
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    mRsPers.Close
    mRsCred.Close
    mRsCuot.Close
    mRsTLiq.Close
    
    Set mRsPers = Nothing
    Set mRsCred = Nothing
    Set mRsCuot = Nothing
    Set mRsTLiq = Nothing
End Sub

Private Sub grilla_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    ' no dejo si esa row esta cobrada
    If fueCobrado(Row) Then
        Cancel = True
        Exit Sub
    End If
    
    Select Case Col
     Case gCUOT: Cancel = True
     Case gPERI:
     Case gIMPO:
     Case gCOBR: Cancel = True
    End Select
End Sub

Private Sub grilla_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
     Case gCUOT: Cancel = True
     Case gPERI: Cancel = Not chequeomes(grilla.EditText)
     Case gIMPO: grilla.EditText = s2n(grilla.EditText, 2)
     Case gCOBR: Cancel = True
    End Select
End Sub
Private Function chequeomes(cualmes As String) As Boolean
    chequeomes = IsDate("28/" & cualmes)
End Function

Private Sub numCuotas_cambio(numero As Double)
    'OJO CONTROLAR los cobrados
    ' ahora asume ALTA solamente
    Dim i As Long, parcial As Double
    Dim Mes As Long, Anio As Long
    
    
    If numero < 1 Then numCuotas.num = 1
    parcial = Round(numTotal.num / numCuotas.num, 2)
    Mes = uFecha.Mes
    Anio = uFecha.Anio
    
    With grilla
        .rows = numero + 1
        For i = 1 To .rows - 1
            .TextMatrix(i, gCUOT) = i
            
            .TextMatrix(i, gPERI) = Format(Mes, "00") & "/" & Right(Format(Anio, "00"), 2)
            Mes = Mes + 1
            If Mes = 13 Then
                Mes = 1
                Anio = Anio + 1
            End If
            
            .TextMatrix(i, gIMPO) = parcial
        Next i
    End With
End Sub

Private Sub numTotCheq_GotFocus()
    If fraOP.enabled = False Then Exit Sub
    numTotCheq.num = numTotal.num - numTotEfec.num
End Sub
Private Sub numTotDepo_GotFocus()
    If fraOP.enabled = False Then Exit Sub
    numTotDepo = numTotal.num - numTotEfec.num - numTotCheq.num
End Sub
Private Sub numTotEfec_GotFocus()
    If fraOP.enabled = False Then Exit Sub
    numTotEfec.num = numTotal.num
End Sub

Private Sub queCuenta()
    If optRemu Then
        txtCuentaContable = CuentaParam(ID_Cuenta_M_REMUNERACIONES_A_PAGAR)
    Else
        txtCuentaContable = mCtaEmpleado
    End If
    If Trim(txtCuentaContable) = "" Then
        che "falta definir cuenta contable"
    End If
End Sub

Private Function FaltaAlgo() As Boolean
    Dim i As Long, sumGri As Double, sumOP As Double
    
    FaltaAlgo = True  ' hasta que demuestre lo contrario
    
    
    'Seteo cuenta, ya deberia estar, pero...
    queCuenta
    
    'totales
    If numTotal.num = 0 Then
    
    End If
    For i = 1 To grilla.rows - 1
        sumGri = sumGri + s2n(grilla.TextMatrix(i, gIMPO))
    Next i
    sumGri = Round(sumGri, 2)
    sumOP = Round(numTotEfec.num + numTotCheq.num + numTotDepo.num, 2)
    If sumGri <> numTotal.num Or sumGri <> sumOP Then
        che "totales no coinciden " & vbCrLf & "grilla " & sumGri & vbCrLf & "Pago " & sumOP & vbCrLf & "Tot " & numTotal.num
        Exit Function
    End If
    
    'personal
    If uCuil.Text = "" Then
        che "falta cuil"
        Exit Function
    End If
    
    If s2n(txtCodTliq) = 0 Then
        che "falta tipo liquidacion"
        Exit Function
    End If
    
    If numCuotas.num = 0 Then
        che "faltan cant cuotas"
        Exit Function
    End If
'    If numCuotas.num > 1 And Trim(txtCtaPrestamo) = "" Then
'        che "Falta definir cuenta contable Prestamo " & uCuil.Text
'        Exit Function
'    End If
    If Trim(txtCuentaContable) = "" Then
        che "Falta definir cuenta contable " & uCuil.Text
        Exit Function
    End If

    
    'porlas dudas
    If numCuotas.num <> grilla.rows - 1 Then
        che "err: cantidad de cuotas no coincide con grilla"
        Exit Function
    End If
    'op
    If numTotCheq.num <> 0 Then
        If uFechaCheque.dtFecha < CDate("01/01/2005") Or uNroCheque.codigo = 0 Then
            che "faltan datos cheque"
            Exit Function
        End If
    End If
    
    If numTotDepo.num <> 0 Then
        If uCtaBanc.codigo = 0 Then
            che "falta cuenta banco"
            Exit Function
        End If
    End If
    
    
    FaltaAlgo = False

End Function

Private Sub ImprimirOP()
    
    'terminar ver si agrego aca pasando parametros y los pongo en el reporte
    'o bien los tomo directamente...
    Dim cade As String
    Dim cade2 As String
    Dim rs As New ADODB.Recordset
        
    rptPagoAdelanto2.Label12 = uCuil
    rptPagoAdelanto2.Label36 = uCuil
    rptPagoAdelanto2.Label13 = txtApellido
    rptPagoAdelanto2.Label37 = txtApellido
    rptPagoAdelanto2.Label14 = uFecha.strFecha
    rptPagoAdelanto2.Label38 = uFecha.strFecha
    rptPagoAdelanto2.Label15 = Mid(uCuil, 4, 8)
    rptPagoAdelanto2.Label39 = Mid(uCuil, 4, 8)
    rptPagoAdelanto2.Label16 = Legajo 'numNumero
    rptPagoAdelanto2.Label40 = Legajo 'numNumero
    
    rptPagoAdelanto2.Label6 = Now
    rptPagoAdelanto2.Label30 = Now
    rptPagoAdelanto2.Label18 = enletras(numTotal)
    rptPagoAdelanto2.Label42 = enletras(numTotal)
    If Not txtCodTliq = "" Then
        cade = "en concepto de " & txtTliq & " Nº "
        cade2 = txtCodTliq
    Else
        cade = ""
        cade2 = ""
    End If
    rptPagoAdelanto2.Label20 = cade
    rptPagoAdelanto2.Label44 = cade
    rptPagoAdelanto2.Label19 = cade2
    rptPagoAdelanto2.Label43 = cade2
    rptPagoAdelanto2.Label24 = Format(Round(numTotal, 4), IIf(2 = 0, "#,#", "#,0." & Left("00000000", 2)))
    rptPagoAdelanto2.Label48 = Format(Round(numTotal, 4), IIf(2 = 0, "#,#", "#,0." & Left("00000000", 2)))
    rptPagoAdelanto2.Label22 = txtMotivo.Text
    rptPagoAdelanto2.Label46 = txtMotivo.Text
    'x = Format(mNumero, IIf(mDecimales = 0, "#,#", "#,0." & Left("00000000", mDecimales))) 'Format(mNumero)
    'txtNumero = IIf(IsNumeric(x), x, txtNumero)
    
    rptPagoAdelanto1.Label3 = Now
    If Not Cajon = 0 Then
        rptPagoAdelanto1.Label5 = Format(Cajon, "00000000")
        rs.Open "select * from movicaja where movimiento=" & Cajon, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        rptPagoAdelanto1.Label7 = rs!concepto
        rptPagoAdelanto1.Label13 = rs!cuenta
        Set rs = Nothing
    End If
    rs.Open "select * from cajas where codigo=" & uCaja.codigo, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    If IsNull(rs) Or IsEmpty(rs) Or (rs.EOF = True And rs.BOF = True) Then
        rptPagoAdelanto1.Label9 = 0
    Else
        rptPagoAdelanto1.Label9 = rs!codigo & " - " & rs!sector
    End If
    Set rs = Nothing
    rptPagoAdelanto1.Label10 = "Por la cantidad de pesos " & enletras(numTotal)
    rptPagoAdelanto1.Label12 = Format(Round(numTotal, 4), IIf(2 = 0, "#,#", "#,0." & Left("00000000", 2)))
    '************************************************
    If numTotEfec.num <> 0 Then
        rptPagoAdelanto1.Label15.caption = "En Efectivo "
        rptPagoAdelanto1.Label16.caption = "............................"
        rptPagoAdelanto1.Label17.caption = "............................................"
        rptPagoAdelanto1.Label18.caption = "......................................"
        rptPagoAdelanto1.Label11.caption = Format(Round(numTotal, 4), IIf(2 = 0, "#,#", "#,0." & Left("00000000", 2)))
    End If
    If numTotCheq.num <> 0 Then
        rs.Open "select * from chq_comp inner join bancosgrales on bancosgrales.codigo=chq_comp.banco where chq_comp.iddoc=" & midDoc, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        rptPagoAdelanto1.Label15.caption = "En Cheque "
        rptPagoAdelanto1.Label16.caption = rs!Nro
        rptPagoAdelanto1.Label17.caption = rs!descripcion
        rptPagoAdelanto1.Label18.caption = rs!fecha_cheque
        rptPagoAdelanto1.Label11.caption = Format(Round(numTotal, 4), IIf(2 = 0, "#,#", "#,0." & Left("00000000", 2)))
        Set rs = Nothing
    End If
    If numTotDepo.num <> 0 Then
        rs.Open "select movibanc.movbanco,bancosgrales.descripcion from movibanc inner join ctasbank on ctasbank.codigo=movibanc.cuenta inner join bancosgrales on bancosgrales.codigo=ctasbank.banco where movibanc.iddoc=" & midDoc, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        rptPagoAdelanto1.Label15.caption = "En Transferencia "
        rptPagoAdelanto1.Label16.caption = rs!MovBanco
        rptPagoAdelanto1.Label17.caption = rs!descripcion
        rptPagoAdelanto1.Label18.caption = "......................................"
        rptPagoAdelanto1.Label11.caption = Format(Round(numTotal, 4), IIf(2 = 0, "#,#", "#,0." & Left("00000000", 2)))
        Set rs = Nothing
    End If
    '***********************************************
    
    rptPagoAdelanto1.Show
    rptPagoAdelanto2.Show

End Sub

Private Sub optPRES_Click()
    queCuenta
End Sub
Private Sub optPRES_LostFocus()
    queCuenta
End Sub
Private Sub optRemu_Click()
    queCuenta
End Sub
Private Sub optRemu_LostFocus()
    queCuenta
End Sub

Private Sub uMenu_AceptarAlta()
'   registrodoc nuevo tipo
'   asiento

'   op
'   VERIFICAR FECHA CHEQUE!
    If ON_ERROR_HABILITADO Then On Error GoTo ufaChe

    Dim i As Long
    Dim asie    As New Asiento
    Dim iddoc   As Long
    Dim NroDoc  As Long, tdoc   As String
    Dim nroPag  As Long
    Dim ctaCaja As String
    Dim ctaEmpl As String
    Dim ctaChq  As String
    Dim movca   As Long, movba  As Long
    Dim conc    As String
    Dim inte    As Long
    Dim asse    As String 'donde fallo
    Dim fech    As Date
    Dim fecha As Date
    

    If FaltaAlgo Then Exit Sub
    
    tdoc = "SUE"
    NroDoc = nuevoCodigo("registroDocumentos", "nroDoc")
    nroPag = nuevoCodigoOP()
    ctaEmpl = Trim(txtCuentaContable) 'IIf(numCuotas.num = 1, CuentaParam(ID_Cuenta_M_REMUNERACIONES_A_PAGAR), txtCtaPrestamo)   '*^*'
    
    conc = IIf(optPres, "Prestamo ", "Remuneracion ")
    conc = conc & Left(txtApellido, 15)
    conc = conc & " " & txtTliq & " "
    conc = Left(conc, 50)    ' VER QUE QUIEREN COMO CONCEPTO
    
    inte = s2n(uNroCheque.descripcion)
    fech = uFecha.dtFecha
    
    asse = "0- Empiezo SQL"
    DE_BeginTrans
        iddoc = NuevoDocumento(tdoc, NroDoc, 0, nroPag, 0, 0)
        asie.Nuevo conc, fech, tdoc
        asie.AgregarItem ctaEmpl, numTotal.num, 0
    '    asie.
        Cajon = 0
        Cajon = NuevoMoviCaja()
        'pagos
        If numTotEfec.num <> 0 Then
            ctaCaja = verCuentaContableCaja(uCaja.codigo)
            DataEnvironment1.dbo_MOVICAJAS "A", uCaja.codigo, Cajon, _
                0, "E", "E", numTotEfec.num, conc, fech, ctaCaja, 0, 1, iddoc, Date, UsuarioActual()
            asie.AgregarItem ctaCaja, 0, numTotEfec.num
        End If
        If numTotCheq.num <> 0 Then
            ctaCaja = verCuentaContableCaja(uCaja.codigo)
            ctaChq = verCuentaContableBanco(obtenerDeSQL("select cuentabancaria from chq_comp where codigo = " & inte))
            movca = Cajon 'NuevoMoviCaja()
            movba = NuevoMovibanc()
            DataEnvironment1.dbo_MOVLIBMOVICAJA 0, movca, "P", "E", numTotCheq.num, _
                conc, fech, inte, tdoc, NroDoc, 0, _
                movba, "O", iddoc, Date, UsuarioActual()
            DataEnvironment1.dbo_MOVLIBMOVIBANC 0, "L", conc, fech, _
                "P", inte, numTotCheq.num, tdoc, NroDoc, movca, movba, _
                 iddoc, Date, UsuarioActual()
            fecha = uFechaCheque.strFecha
            DataEnvironment1.dbo_MOVLIBCHEQUES inte, fecha, numTotCheq.num, NroDoc, tdoc, "T", uFechaCheque.dtFecha, _
                fech, 0, _
                iddoc, Date, UsuarioActual()
            asie.AgregarItem ctaChq, 0, numTotCheq.num
        End If
        If numTotDepo.num <> 0 Then
            movba = NuevoMovibanc()
            DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanc.codigo, "S", conc, fech, "E", numTotDepo.num, movba, iddoc, Date, UsuarioActual()
            asie.AgregarItem uCtaBanc.CuentaContable, 0, numTotDepo.num
        End If

        asie.Grabar iddoc
        ' Recien ahora trato de grabar DBF que estan FUERA de la transaccion , si falla puedo volver atras sqlserver
        
        asse = "1- Empiezo DBF CRED # " & numNumero.num & " " & uCuil.Text
        With mRsCred
            .AddNew
                !iddoc = iddoc
                !credito = numNumero.num
                !cuil = uCuil.Text
                !Total = numTotal.num
                !inicio = fech
                !cuotas = numCuotas.num
                !motivo = txtMotivo
                On Error Resume Next
                !movcaja = Cajon 'aca tenia 0 y lo cambie, ver si esta bien
                !lprest = optPres.Value '(numCuotas.num > 1)
            .Update
        End With
        asse = "2- Empiezo DBF CUOTAS. (DBF.CRED # " & numNumero.num & " " & uCuil.Text & " )"
        With mRsCuot
            For i = 1 To numCuotas.num
                .AddNew
                    !credito = numNumero.num
                    !cuil = uCuil.Text
                    !cuota = i
                    !periodo = grilla.TextMatrix(i, gPERI)
                    !codtip = s2n(txtCodTliq)
                    !importe = s2n(grilla.TextMatrix(i, gIMPO))
                    !cobrado = False
                    !iddoc = iddoc
                .Update
            Next i
        End With
        
    DE_CommitTrans
    
    ' y ahora como joraca imprimo OP...
    ImprimirOP
    
    uMenu.AceptarOk

fin:
    Set asie = Nothing
    Exit Sub
ufaChe:
    DE_RollbackTrans
    ufa "err de grabacion", "adelantos sueldos: " & asse
    Resume fin
End Sub

Private Sub uMenu_AceptarModi()
'
    If FaltaAlgo Then Exit Sub
    
End Sub
Private Sub uMenu_BorrarControles()
    FrmBorrarTxt Me
    
    uFecha.dtFecha Date
    uFechaCheque.dtFecha CDate("2000-01-01")  ' para verificar si lo puso
    uCuil.Text = ""
    grilla.rows = 1
    
    FrmBorrarNum Me
    numCuotas.num = 1

    midDoc = 0
    uCaja.codigo = 1
    uNroCheque.clear
    uCtaBanc.codigo = 0
    mCtaEmpleado = ""
    optRemu = True: queCuenta
    
    inicheque False
End Sub
Private Sub uMenu_Buscar()
' se pone jodido...

    'si tiene algo mostrando, que no pise y deje cambio por la mitad
    If uMenu.Estado = ucbMostrando Then Exit Sub
    
    inicheque True
    
    'buscar persona
    cmdBuscarPersonal_Click
    If uCuil.Text = "" Then Exit Sub
    
    'buscar nro
    Dim a(), i As Integer
    With mRsCred
        .MoveFirst
        ReDim a(3, 0)
        a(0, 0) = "Numero          "
        a(1, 0) = "Fecha           "
        a(2, 0) = "Total           "
        a(3, 0) = "id   "
        
        While Not .EOF
            If !cuil = uCuil.Text And nSinNull(!iddoc) > 0 Then
                i = i + 1
                ReDim Preserve a(3, i)
                
                a(0, i) = !credito
                a(1, i) = !inicio
                a(2, i) = !Total
                a(3, i) = nSinNull(!iddoc)
            End If
            .MoveNext
        Wend
    End With
    With frmBuscar
        .MostrarArray (a)
        If .resultado > "" Then
            midDoc = .resultado(4)
            
            CargoCredito
            CargoCuotas
            CargoOP
            
            uMenu.BuscarOK
        End If
    End With
   

End Sub

Private Sub CargoCredito()
    Dim rs As New ADODB.Recordset
    On Error GoTo ufaChe
    
    With mRsCred
        .MoveFirst
        While Not .EOF
            If s2n(!iddoc) = midDoc Then
                If !lprest Then optPres = True Else optRemu = True
                numTotal = !Total
                uFecha.dtFecha !inicio
                numCuotas = !cuotas
                txtMotivo = !motivo
                Cajon = !movcaja 'esto es para obtener el movimiento para imprimir
                Liquida
                Exit Sub
            End If
            .MoveNext
        Wend
        midDoc = 0
        che "no pude cargar el pago"
    End With
fin:
    Exit Sub
ufaChe:
    midDoc = 0
    che "no pude cargar el pago"
    Resume fin
End Sub
Private Sub CargoCuotas()
    On Error GoTo ufaChe
    Dim i As Long
    
    iniPersonal
    
    With mRsCuot
        .MoveFirst
        While Not .EOF
            If s2n(!iddoc) = midDoc Then
                i = i + 1
                grilla.rows = i + 1
                grilla.TextMatrix(i, gPERI) = !periodo
                grilla.TextMatrix(i, gIMPO) = !importe
                grilla.TextMatrix(i, gCUOT) = !cuota
                grilla.TextMatrix(i, gCOBR) = !cobrado
            End If
            .MoveNext
        Wend
        If i = 0 Then
            midDoc = 0
            che "no pude cargar cuotas"
        End If
    End With
fin:
    Exit Sub
ufaChe:
    midDoc = 0
    che "no pude cargar cuotas"
    Resume fin
End Sub
Private Sub CargoOP()
    Dim tempo
    Dim rs1 As New ADODB.Recordset
    rs1.Open "select * from movicaja where movimiento=" & Cajon, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    'efectico
    If rs1!tipo = "E" Then
        tempo = obtenerDeSQL("select caja,importe from movicaja where movimiento= " & Cajon)
        If Not IsEmpty(tempo) Then
            uCaja.codigo = tempo(0)
            numTotEfec = tempo(1)
        End If
    End If
    'chequico
    If rs1!tipo = "P" Then
        tempo = obtenerDeSQL("select nro, importe,fecha_cheque from chq_comp where iddoc = " & midDoc)
        If Not IsEmpty(tempo) Then
            uNroCheque.codigo = tempo(0)
            numTotCheq = tempo(1)
            'uFechaCheque = uFechaCheque.strFecha(tempo(2))
            uFechaCheque.contenido (tempo(2))
        End If
    End If
    
    'transfica
    If rs1!tipo = "T" Then
        tempo = obtenerDeSQL("select cuenta,importe from movibanc where iddoc = " & midDoc)
        If Not IsEmpty(tempo) Then
            uCtaBanc.codigo = tempo(0)
            numTotDepo = tempo(1)
        End If
    End If

End Sub

Private Function fueCobrado(i As Long) As Boolean
    If i < 1 Then Exit Function
    fueCobrado = (grilla.cell(flexcpChecked, i, gCOBR) = flexChecked)
End Function

Private Function puedoEliminar() As Boolean
    Dim i As Long
    Dim e As String
    
    'barrer la grilla y ver si alguno fue cobrado
    For i = 1 To grilla.rows - 1
        If fueCobrado(i) Then
            che "Ya fue descontado" & vbCrLf & "no se puede eliminar"
            Exit Function
        End If
    Next i

    If midDoc = 0 Then Exit Function  ' solo los que genere yo

    If numTotCheq > 0 Then
        e = obtenerDeSQL("select estado from chq_comp where codigo = " & uNroCheque.descripcion)
        If e <> "T" Then
            che "El cheque no se puede recuperar"
            Exit Function
        End If
    End If
    

    puedoEliminar = True
End Function

Private Sub uMenu_eliminar()
    If ON_ERROR_HABILITADO Then On Error GoTo ufaChe
    Dim i As Long
    Dim loborre As Boolean
    Dim asse As String
    
    If Not puedoEliminar() Then Exit Sub
    
    
    DE_BeginTrans
        'borrar regdoc
        'cheque
        DataEnvironment1.Sistema.Execute "update chq_comp set estado = 'C', importe = 0 where codigo = " & uNroCheque.descripcion
        'movis
        DataEnvironment1.Sistema.Execute "delete from movicaja where iddoc = " & midDoc
        DataEnvironment1.Sistema.Execute "delete from movibanc where iddoc = " & midDoc
        'asiento, registro
        BorroDocumento midDoc
        
        'borrar registros dbf
        
        asse = "empiezo elim cred " & midDoc
        With mRsCred
            loborre = False
            .MoveFirst
            Do While Not .EOF
                If s2n(!iddoc) = midDoc Then
                    loborre = True
                    .Delete
                    .Update ' no se si es necesario
                    Exit Do
                End If
                .MoveNext
            Loop
            If Not loborre Then Err.Raise 555000, "No pude borrar tabla credito"
        End With
        
        asse = "empiezo elim cuotas: " & midDoc
        With mRsCuot
            loborre = False
            .MoveFirst
            Do While Not .EOF
                If s2n(!iddoc) = midDoc Then
                    loborre = True
                    .Delete
                    .Update ' no se si es necesario
                End If
                .MoveNext
            Loop
            If Not loborre Then Err.Raise 555000, "No pude borrar tabla cuotas"
        End With
        
    DE_CommitTrans
    uMenu.EliminarOK
fin:
    Exit Sub
ufaChe:
    DE_RollbackTrans
    ufa "err al eliminar", asse
    Resume fin
End Sub
Private Sub uMenu_HabilitarEdicionAM(sino As Boolean, SiNoAlta As Boolean)
    ' en modi, no dejo cambiar clave cuit + numero
    ' ni total, ni cuenta prestamo-anticipo
    
    cmdBuscarPersonal.enabled = SiNoAlta
    numTotal.enabled = SiNoAlta
    optPres.enabled = SiNoAlta
    optRemu.enabled = SiNoAlta
    fraOP.enabled = SiNoAlta
    
    cmdTliq.enabled = sino
    grilla.Editable = IIf(sino, flexEDKbdMouse, flexEDNone)

End Sub

Private Sub uMenu_Imprimir()
    ImprimirOP
End Sub

Private Sub uMenu_Nuevo()
    On Error Resume Next
    cmdBuscarPersonal.SetFocus
End Sub

Private Sub uMenu_SALIR()
    Unload Me
End Sub

Private Function Liquida() 'ver si agrego cajon
    Dim cadena As String
    Dim cadena2 As String
    Dim cod As Long
    Dim pos As Long
    Dim rs2 As New ADODB.Recordset
    rs2.Open "select * from movicaja where movimiento=" & Cajon, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    cadena = rs2!concepto
        
    With mRsTLiq
        .MoveFirst
        While Not .EOF
            cadena2 = !denom
            pos = InStr(1, cadena, cadena2, vbTextCompare)
            If pos > 0 Then
                cod = !codtip
                txtCodTliq.Text = cod
                txtTliq.Text = cadena2
                Exit Function
            End If
            .MoveNext
        Wend
    End With
End Function


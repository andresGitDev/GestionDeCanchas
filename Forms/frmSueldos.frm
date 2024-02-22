VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmSueldos 
   Caption         =   "Pago de Anticipos, Prestamos y Sueldos"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9900
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   9900
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCtaPrestamo 
      Enabled         =   0   'False
      Height          =   360
      Left            =   7620
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   105
      Width           =   2220
   End
   Begin VB.Frame fraOP 
      Caption         =   "Orden de Pago"
      Height          =   1860
      Left            =   105
      TabIndex        =   23
      Top             =   5535
      Width           =   9705
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   360
         Left            =   8100
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   855
         Width           =   1500
      End
      Begin Gestion.ucFecha uFechaCheque 
         Height          =   300
         Left            =   6345
         TabIndex        =   18
         Top             =   885
         Width           =   1035
         _ExtentX        =   2117
         _ExtentY        =   529
         FechaInit       =   4
      End
      Begin Gestion.uCtaBanco uCtaBanc 
         Height          =   330
         Left            =   2925
         TabIndex        =   20
         Top             =   1320
         Width           =   6675
         _ExtentX        =   9234
         _ExtentY        =   582
      End
      Begin Gestion.ucCoDe uCaja 
         Height          =   330
         Left            =   2925
         TabIndex        =   15
         Top             =   345
         Width           =   5220
         _ExtentX        =   9208
         _ExtentY        =   609
         CodigoWidth     =   800
      End
      Begin Gestion.uNum numTotEfec 
         Height          =   330
         Left            =   870
         TabIndex        =   14
         Top             =   345
         Width           =   1185
         _ExtentX        =   2275
         _ExtentY        =   582
      End
      Begin Gestion.uNum numTotCheq 
         Height          =   330
         Left            =   870
         TabIndex        =   16
         Top             =   870
         Width           =   1185
         _ExtentX        =   2275
         _ExtentY        =   582
      End
      Begin Gestion.uNum numTotDepo 
         Height          =   330
         Left            =   870
         TabIndex        =   19
         Top             =   1320
         Width           =   1185
         _ExtentX        =   2275
         _ExtentY        =   582
      End
      Begin Gestion.ucCoDe uNroCheque 
         Height          =   330
         Left            =   2910
         TabIndex        =   17
         Top             =   870
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
         Left            =   5730
         TabIndex        =   35
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
         Left            =   2160
         TabIndex        =   34
         Top             =   915
         Width           =   735
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2220
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
         Top             =   345
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cheques"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   165
         TabIndex        =   30
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
         TabIndex        =   29
         Top             =   390
         Visible         =   0   'False
         Width           =   675
      End
   End
   Begin VB.Frame fraPrestamo 
      Height          =   4515
      Left            =   135
      TabIndex        =   22
      Top             =   990
      Width           =   9630
      Begin VB.TextBox txtCodTliq 
         Height          =   345
         Left            =   1665
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   285
         Width           =   1335
      End
      Begin VB.TextBox txtTliq 
         Height          =   345
         Left            =   3135
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   270
         Width           =   4545
      End
      Begin VB.CommandButton cmdTliq 
         Caption         =   "?"
         Height          =   345
         Left            =   1035
         TabIndex        =   6
         Top             =   255
         Width           =   495
      End
      Begin VB.TextBox txtMotivo 
         Height          =   375
         Left            =   1005
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   840
         Width           =   6660
      End
      Begin Gestion.uNum numTotal 
         Height          =   330
         Left            =   855
         TabIndex        =   11
         Top             =   2385
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   582
      End
      Begin Gestion.ucFecha uFecha 
         Height          =   345
         Left            =   885
         TabIndex        =   10
         Top             =   1830
         Width           =   1290
         _ExtentX        =   2223
         _ExtentY        =   609
         FechaInit       =   0
      End
      Begin VSFlex7LCtl.VSFlexGrid grilla 
         Height          =   3000
         Left            =   2445
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   28
         Top             =   300
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Motivo:"
         Height          =   270
         Index           =   2
         Left            =   180
         TabIndex        =   27
         Top             =   900
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Total:"
         Height          =   270
         Index           =   8
         Left            =   225
         TabIndex        =   26
         Top             =   2445
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Inicio:"
         Height          =   270
         Index           =   9
         Left            =   240
         TabIndex        =   25
         Top             =   1935
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Cuotas:"
         Height          =   270
         Index           =   4
         Left            =   255
         TabIndex        =   24
         Top             =   3030
         Width           =   840
      End
   End
   Begin VB.TextBox txtApellido 
      Height          =   345
      Left            =   2955
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   105
      Width           =   4560
   End
   Begin VB.CommandButton cmdBuscarPersonal 
      Caption         =   "?"
      Height          =   345
      Left            =   2385
      TabIndex        =   1
      Top             =   105
      Width           =   495
   End
   Begin Gestion.ucBotonera uMenu 
      Align           =   2  'Align Bottom
      Height          =   600
      Left            =   0
      TabIndex        =   21
      Top             =   7455
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   1058
      MsgConfirmaSalir=   ""
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   ""
      CaptionEliminar =   "&Eliminar"
   End
   Begin Gestion.ucCuit uCuil 
      Height          =   345
      Left            =   975
      TabIndex        =   2
      Top             =   90
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   609
      Enabled         =   0   'False
   End
   Begin Gestion.uNum numNumero 
      Height          =   330
      Left            =   960
      TabIndex        =   5
      Top             =   540
      Width           =   1020
      _ExtentX        =   2275
      _ExtentY        =   582
      Enabled         =   0   'False
      Locked          =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "CUIL"
      Height          =   255
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   135
      Width           =   630
   End
   Begin VB.Label Label1 
      Caption         =   "Numero:"
      Height          =   270
      Index           =   3
      Left            =   210
      TabIndex        =   4
      Top             =   540
      Width           =   735
   End
End
Attribute VB_Name = "frmSueldos"
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

Private Const mTitCuota = "Cuota  | Periodo  | Importe    | Cobrado"
Private Enum gricolcuota
'    gCUIL ' invis
'    gCred ' invis
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
            txtCtaPrestamo = .resultado(3)
        End If
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
    
    uMenu.init True, True, True, False, True
End Sub

Private Sub iniDBF()
    cnxAbrirDBF obtenerParametro("PathSueldos")
    
    mRsPers.Open "personal", cnxDBF, adOpenDynamic, adLockReadOnly
    mRsTLiq.Open "TipoLiq ", cnxDBF, adOpenDynamic, adLockReadOnly
    
    mRsCred.Open "credito ", cnxDBF, adOpenDynamic, adLockOptimistic
    mRsCuot.Open "cuotas  ", cnxDBF, adOpenDynamic, adLockOptimistic
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
    uNroCheque.ini "select codigo from chq_comp where estado = 'C' and activo = 1 and nro  = '###' ", "select nro Numero, codigo , banco from chq_comp where activo = 1 and estado = 'C' ", True
    
    uFechaCheque.dtfecha CDate("01/01/2000")
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
    If grilla.cell(flexcpChecked, Row, Col) = flexChecked Then
        Cancel = True
        Exit Sub
    End If
    
    Select Case Col
'     Case gCUIL: Cancel = True
     Case gCUOT: Cancel = True
'     Case gCred: Cancel = True
     Case gPERI:
     Case gIMPO:
     Case gCOBR: Cancel = True
    End Select
End Sub

Private Sub grilla_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
'     Case gCUIL: Cancel = True
     Case gCUOT: Cancel = True
'     Case gCred: Cancel = True
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
    Dim mes As Long, Anio As Long
    
    
    If numero < 1 Then numCuotas.num = 1
    parcial = Round(numTotal.num / numCuotas.num, 2)
    mes = uFecha.mes
    Anio = uFecha.Anio
    
    With grilla
        .rows = numero + 1
        For i = 1 To .rows - 1
            .TextMatrix(i, gCUOT) = i
            
            .TextMatrix(i, gPERI) = Format(mes, "00") & "/" & Right(Format(Anio, "00"), 2)
            mes = mes + 1
            If mes = 13 Then
                mes = 1
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

Private Function FaltaAlgo() As Boolean
    Dim i As Long, sumGri As Double, sumOP As Double
    
    FaltaAlgo = True  ' hasta que demuestre lo contrario
    
    
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
    If numCuotas.num > 1 And Trim(txtCtaPrestamo) = "" Then
        che "Falta definir cuenta contable Prestamo " & uCuil.Text
        Exit Function
    End If
    
    'porlas dudas
    If numCuotas.num <> grilla.rows - 1 Then
        che "err: cantidad de cuotas no coincide con grilla"
        Exit Function
    End If
    'op
    If numTotCheq.num <> 0 Then
        If uFechaCheque.dtfecha < CDate("01/01/2005") Or uNroCheque.codigo = 0 Then
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
    
    

    If FaltaAlgo Then Exit Sub
    
    tdoc = "SUE"
    NroDoc = nuevoCodigo("registroDocumentos", "nroDoc")
    nroPag = nuevoCodigoOP()
    ctaEmpl = "1111111" ' IIf(numCuotas.num = 1, CuentaParam(ID_Cuenta_M_ANTICIPOS_PERSONAL), txtCtaPrestamo)   '*^*'
    conc = Left("suel/antic/prest " & uCuil.Text & " " & txtTliq, 50) ' VER QUE QUIEREN COMO CONCEPTO
    inte = s2n(uNroCheque.descripcion)
    
    DE_BeginTrans
        iddoc = NuevoDocumento(tdoc, NroDoc, 0, nroPag, 0, 0)
        asie.Nuevo conc, uFecha.dtfecha, tdoc
        asie.AgregarItem ctaEmpl, numTotal.num, 0
    '    asie.
        
        'pagos
        If numTotEfec.num > 0 Then
            ctaCaja = verCuentaContableCaja(uCaja.codigo)
            DataEnvironment1.dbo_MOVICAJAS "A", uCaja.codigo, NuevoMoviCaja(), _
                0, "E", "E", numTotEfec.num, txtTliq, uFecha.dtfecha, ctaCaja, 0, 1, iddoc, Date, UsuarioActual()
            asie.AgregarItem ctaCaja, 0, numTotEfec.num
        End If
        If numTotCheq.num > 0 Then
            ctaCaja = verCuentaContableCaja(uCaja.codigo)
            ctaChq = verCuentaContableBanco(obtenerDeSQL("select cuentabancaria from chq_comp where codigo = " & inte))
            movca = NuevoMoviCaja()
            movba = NuevoMovibanc()
            DataEnvironment1.dbo_MOVLIBMOVICAJA 0, movca, "P", "E", numTotCheq.num, _
                conc, uFecha.dtfecha, inte, tdoc, NroDoc, 0, _
                movba, "O", iddoc, Date, UsuarioActual()
            DataEnvironment1.dbo_MOVLIBMOVIBANC 0, "L", conc, uFecha.dtfecha, _
                "P", inte, numTotCheq.num, tdoc, NroDoc, movca, movba, _
                 iddoc, Date, UsuarioActual()
            DataEnvironment1.dbo_MOVLIBCHEQUES inte, uFechaCheque.dtfecha, numTotCheq.num, NroDoc, tdoc, "T", uFechaCheque.dtfecha, _
                uFecha.dtfecha, 0, _
                iddoc, Date, UsuarioActual()
            asie.AgregarItem ctaChq, 0, numTotCheq.num
        End If
        If numTotDepo.num > 0 Then
            movba = NuevoMovibanc()
            DataEnvironment1.dbo_MOVIBANCOS "A", uCtaBanc.codigo, "E", conc, uFecha.dtfecha, "E", numTotDepo.num, movba, iddoc, Date, UsuarioActual()
            asie.AgregarItem uCtaBanc.CuentaContable, 0, numTotDepo.num
        End If

        asie.Grabar iddoc
        ' Recien ahora trato de grabar DBF que estan FUERA de la transaccion , si falla puedo volver atras sqlserver
        
        
        
        With mRsCred
            .AddNew
            !credito = numNumero.num
            !cuil = uCuil.Text
            !Total = numTotal.num
            !Inicio = uFecha.dtfecha
            !cuotas = numCuotas.num
            !motivo = txtMotivo
            !movcaja = 0
            !lprest = (numCuotas.num > 1)
            .Update
        End With
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
    ufa "err de grabacion", ""
    Resume fin
End Sub

Private Sub uMenu_AceptarModi()
'
    If FaltaAlgo Then Exit Sub
    
    
    
    
    
    
    
End Sub
Private Sub uMenu_BorrarControles()
    FrmBorrarTxt Me
    uFecha.dtfecha Date
    uFechaCheque.dtfecha CDate("2000-01-01")  ' para verificar si lo puso
    uCuil.Text = ""
    grilla.rows = 1
    numCuotas.num = 1
    numTotal.num = 0
    numNumero.num = 0
End Sub

Private Sub uMenu_Buscar()
' se pone jodido...

    'si tiene algo mostrando, que no pise y deje cambio por la mitad
    If uMenu.Estado = ucbMostrando Then Exit Sub
    
    'buscar persona
    cmdBuscarPersonal_Click
    If uCuil.Text = "" Then Exit Sub
    
    'buscar nro
    Dim a(), i As Integer
    With mRsCred
        .MoveFirst
        ReDim a(2, 0)
        a(0, 0) = "Numero          "
        a(1, 0) = "Fecha           "
        While Not .EOF
            If !cuil = uCuil.Text Then
                i = i + 1
                ReDim Preserve a(2, i)
                
                a(0, i) = !credito
                a(1, i) = !Inicio
            End If
            .MoveNext
        Wend
    End With
    With frmBuscar
        .MostrarArray (a)
        If .resultado > "" Then
        
        End If
    End With
    
    ' cargo credito y cuotas y op !!!
    ' un lio
    
    

End Sub

Private Sub uMenu_eliminar()
    'barrer la grilla y ver si alguno fue cobrado
    'borrar regdoc
    'anular op
    'borrar registros dbf
    
    
    
    
    
    
End Sub
Private Sub uMenu_HabilitarEdicionAM(sino As Boolean, SiNoAlta As Boolean)
    ' en modi, no dejo cambiar clave cuit + numero
    cmdBuscarPersonal.enabled = SiNoAlta
    numTotal.enabled = SiNoAlta
    
    cmdTliq.enabled = sino
    grilla.Editable = IIf(sino, flexEDKbdMouse, flexEDNone)
    
    fraOP.enabled = SiNoAlta

End Sub
Private Sub uMenu_SALIR()
    Unload Me
End Sub


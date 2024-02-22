VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmAsientoManual 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asientos Manuales"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8835
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   8835
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraExcluir 
      Height          =   720
      Left            =   4500
      TabIndex        =   25
      Top             =   6615
      Width           =   4275
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Guardar opcion"
         Height          =   315
         Left            =   2730
         TabIndex        =   27
         Top             =   240
         Width           =   1440
      End
      Begin VB.CheckBox chkExcluir 
         Caption         =   "Excluir para Ajuste por inflacion"
         Height          =   240
         Left            =   135
         TabIndex        =   26
         Top             =   285
         Width           =   2715
      End
   End
   Begin Gestion.ucBotonera uMenu 
      Height          =   1395
      Left            =   1305
      TabIndex        =   24
      Top             =   7470
      Width           =   7455
      _extentx        =   13150
      _extenty        =   2461
      msgconfirmasalir=   ""
      msgconfirmaeliminar=   ""
      msgconfirmacancelar=   ""
      captioneliminar =   "&Eliminar"
   End
   Begin VB.ComboBox cboEjercicio 
      Height          =   315
      Left            =   120
      TabIndex        =   20
      Text            =   "Ejercicio"
      Top             =   6240
      Width           =   990
   End
   Begin Gestion.ucEntreFechas ucFechas 
      Height          =   360
      Left            =   75
      TabIndex        =   19
      Top             =   6690
      Width           =   2655
      _extentx        =   4683
      _extenty        =   635
   End
   Begin VB.Frame fraAsiento 
      Height          =   6615
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   8745
      Begin Gestion.ucFecha uFecha 
         Height          =   315
         Left            =   2340
         TabIndex        =   18
         Top             =   465
         Width           =   1095
         _extentx        =   1931
         _extenty        =   556
         fechainit       =   0
      End
      Begin VB.TextBox txtRegDoc 
         Height          =   330
         Left            =   7095
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   17
         Top             =   825
         Width           =   1560
      End
      Begin VB.TextBox txtConcepto 
         Height          =   315
         Left            =   900
         MaxLength       =   50
         TabIndex        =   9
         Top             =   840
         Width           =   5415
      End
      Begin VB.TextBox txtOrigen 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   4110
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   465
         Width           =   555
      End
      Begin VSFlex7LCtl.VSFlexGrid griMayor 
         Height          =   4605
         Left            =   90
         TabIndex        =   12
         ToolTipText     =   "Doble-clic en campo Cuenta"
         Top             =   1575
         Width           =   8535
         _cx             =   15055
         _cy             =   8123
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
      Begin VB.Label lblEAsiento 
         Height          =   210
         Left            =   4260
         TabIndex        =   23
         Top             =   6285
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.Label Label3 
         Caption         =   " - Ejercicio del asiento"
         Height          =   210
         Left            =   2625
         TabIndex        =   22
         Top             =   6285
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.Label Label2 
         Caption         =   "Ejercicio Actual"
         Height          =   210
         Left            =   1110
         TabIndex        =   21
         Top             =   6285
         Width           =   1470
      End
      Begin VB.Label lblSumaDebe 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5820
         TabIndex        =   16
         Top             =   1215
         Width           =   1395
      End
      Begin VB.Label lblSumaHaber 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   7245
         TabIndex        =   15
         Top             =   1215
         Width           =   1395
      End
      Begin VB.Label lblIdDoc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   7065
         TabIndex        =   14
         Top             =   510
         Width           =   855
      End
      Begin VB.Label label1 
         BackStyle       =   0  'Transparent
         Caption         =   "idDoc:"
         Height          =   315
         Index           =   6
         Left            =   6540
         TabIndex        =   13
         Top             =   525
         Width           =   555
      End
      Begin VB.Label label1 
         BackStyle       =   0  'Transparent
         Caption         =   "idAsiento:"
         Height          =   315
         Index           =   5
         Left            =   4845
         TabIndex        =   11
         Top             =   510
         Width           =   705
      End
      Begin VB.Label lblId 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   5550
         TabIndex        =   10
         Top             =   510
         Width           =   855
      End
      Begin VB.Label label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Origen:"
         Height          =   255
         Index           =   4
         Left            =   3510
         TabIndex        =   7
         Top             =   525
         Width           =   795
      End
      Begin VB.Label label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Concepto:"
         Height          =   315
         Index           =   3
         Left            =   75
         TabIndex        =   6
         Top             =   870
         Width           =   915
      End
      Begin VB.Label label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         Height          =   315
         Index           =   2
         Left            =   1725
         TabIndex        =   5
         Top             =   525
         Width           =   735
      End
      Begin VB.Label lblEjercicio 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2004"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   810
         TabIndex        =   4
         Top             =   210
         Width           =   795
      End
      Begin VB.Label label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ejercicio:"
         Height          =   315
         Index           =   1
         Left            =   105
         TabIndex        =   3
         Top             =   225
         Width           =   615
      End
      Begin VB.Label lblAsiento 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   825
         TabIndex        =   2
         Top             =   495
         Width           =   795
      End
      Begin VB.Label label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Asiento:"
         Height          =   315
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   495
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmAsientoManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private EsVista As Boolean
Private WithEvents g As LiGrilla
Attribute g.VB_VarHelpID = -1
Private Asiento As Asiento
Private mIdEjercicioActivo As Long
Private Const STR_MOVMANUAL = "M"
'Private mIdDoc As Long
Private gCUEN As Long
Private gDESC As Long
Private gDEBE As Long
Private gHABE As Long
Private gORIG As Long 'comprobante origen


Public Sub mostrar(que_IdDoc As Long)
    Dim idasi As Long
    If que_IdDoc = 0 Then Exit Sub
    idasi = s2n(obtenerDeSQL("select idasiento from asientos where iddoc = " & que_IdDoc))
    If idasi > 0 Then
        'uMenu.BuscarOK "idAsiento = " & idasi
        CargaAsiento idasi
        Me.Show
    Else
        che "asiento no encontrado"
    End If
    
End Sub

Private Sub cboEjercicio_Click()
Dim denominacion As String
    mIdEjercicioActivo = leerEjercicioId(cboEjercicio)
    denominacion = "select denominacion from ejercicio where idejercicio= " & leerEjercicioId(cboEjercicio)
    lblEjercicio = obtenerDeSQL(denominacion)
    inimenu
End Sub

Private Sub chkExcluir_Click()
If chkExcluir.Value = 1 Then
    txtOrigen = "E"
Else
    txtOrigen = "M"
End If
End Sub

Private Sub cmdGuardar_Click()
Dim sUpdate As String
If s2n(lblID) > 0 Then
    If chkExcluir.Value = 1 Then
        sUpdate = "update asientos set origen='E' where idasiento=" & s2n(lblID)
    Else
        sUpdate = "update asientos set origen='M' where idasiento=" & s2n(lblID)
    End If
    DataEnvironment1.Sistema.Execute sUpdate
    If chkExcluir.Value = 1 Then
        MsgBox "Asiento excluido para Ajuste por inflacion.", vbInformation
    Else
        MsgBox "Asiento incluido para Ajuste por inflacion.", vbInformation
    End If
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, False, True
End Sub

Public Function VistaPrevia(GRILLA As Control, Titulo As String, Fecha As Date, CuentaFinal As String, ImporteD As Double, ImporteH As Double)
Dim x As Long, i As Long

Me.Show
inimenu True
uFecha.dtFecha Fecha
txtconcepto = Titulo
fraAsiento.enabled = True
txtconcepto.enabled = False
uFecha.enabled = False
griMayor.rows = 1
griMayor.AddItem ""
For i = 1 To GRILLA.rows - 1
    x = griMayor.rows - 1
    griMayor.TextMatrix(x, 0) = GRILLA.TextMatrix(i, 1)
    griMayor.TextMatrix(x, 2) = GRILLA.TextMatrix(i, 3)
    griMayor.TextMatrix(x, 3) = GRILLA.TextMatrix(i, 4)
Next

If Trim(CuentaFinal) > "" Then
    x = griMayor.rows - 1
    griMayor.TextMatrix(x, 0) = CuentaFinal
    griMayor.TextMatrix(x, 2) = s2n(ImporteD)
    griMayor.TextMatrix(x, 3) = s2n(ImporteH)
    
End If
griMayor.Editable = flexEDNone
EsVista = True
fraExcluir.Visible = False
End Function

Private Sub Form_Load()
    Dim rsEjercicio As New ADODB.Recordset
    Dim EjerA As New ADODB.Recordset
    EjerA.Open "select * from ejercicio", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    EjerA.MoveFirst
    While Not EjerA.EOF
        cboEjercicio.AddItem EjerA!denominacion 'EjerA!idejercicio
        EjerA.MoveNext
    Wend
    'obtener ejercicio antes de menu
    lblEjercicio = leerEjercicioDenominacion()
    mIdEjercicioActivo = leerEjercicioId()
    cboEjercicio = lblEjercicio ' mIdEjercicioActivo
    inimenu
    inigrilla
    If UsuarioActual() <> 19 Then cboEjercicio.enabled = False
    Set Asiento = New Asiento
    '    uXls.ini griMayor, ".\AsientoIndividual"
    
    rsEjercicio.Open "SELECT * From Ejercicio WHERE activo =1", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
   
    'ucFechas.ini CDate(rsEjercicio!FechaInicio), CDate(rsEjercicio!FechaFin), ucefHorizontal, ucefFormatoSqlServer
    ucFechas.ini Date - 30, Date, ucefHorizontal, ucefFormatoSqlServer
    rsEjercicio.Close
    
    Form_Resize
    EsVista = False
End Sub

Private Function CargaAsiento(idAsiento As Long) As Boolean
    If ON_ERROR_HABILITADO Then On Error GoTo UfaCarga
    Dim rsM As New ADODB.Recordset
    Dim i As Long, aCerrado As Boolean, aBloqueado As Boolean
    Dim tempo, tmp
    'lblEAsiento
    g.Borrar
    g.rows = 2
    txtRegDoc = ""
    
    If idAsiento > 0 Then
'        mIdDoc = !idDoc
        tempo = obtenerDeSQL("select idasiento,iddoc,nroasiento,fecha,origen,concepto,ejercicio from asientos where idasiento=" & idAsiento)
        lblId = tempo(0)
        lblIdDoc = tempo(1)
        lblAsiento = tempo(2)
        uFecha.dtFecha tempo(3)
        txtOrigen = tempo(4)
        txtConcepto = tempo(5)
        lblEAsiento = obtenerDeSQL("select denominacion from ejercicio where idejercicio=" & tempo(6))
'        tmp = obtenerDeSQL("select cerrado,bloqueado from ejercicio where idejercicio=" & tempo(6))
'        aCerrado = nSinNull(tmp(0))
'        aBloqueado = nSinNull(tmp(1))
        aCerrado = False
        aBloqueado = False
        lblEjercicio = lblEAsiento
        Set tempo = Nothing
        
        tempo = sSinNull(obtenerDeSQL("select tipodoc, NroDoc from registrodocumentos where iddoc = " & lblIdDoc))
        If Not IsEmpty(tempo) Then txtRegDoc = tempo(0) & " " & tempo(1)
        Set rsM = Nothing
        rsM.Open "select * from mayor where idAsiento = " & lblId, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
        rsM.MoveFirst
        griMayor.rows = 1
        For i = 1 To rsM.RecordCount
            g.addRow
            g.tx i, gCUEN, rsM!Cuenta
            g.tx i, gDESC, DescrCuenta(rsM!Cuenta)  '
            g.tx i, gDEBE, rsM!Debe
            g.tx i, gHABE, rsM!haber
            g.tx i, gORIG, rsM!COMPROBANTE
            
            rsM.MoveNext
        Next

    End If
    CargaAsiento = True
    If aCerrado Or aBloqueado Then
        MsgBox "El asiento no se puede modificar porque el ejercicio se encuentra cerrado o bloqueado.", vbInformation
        CargaAsiento = False
    End If
    
    PermitoExcluir
    
fin:
    Set rsM = Nothing
    Exit Function
UfaCarga:
    CargaAsiento = False
    Resume fin
End Function

Private Function PermitoExcluir()
If Trim(txtOrigen) = "M" Or Trim(txtOrigen) = "E" Or Trim(txtOrigen) = "" Then
    fraExcluir.enabled = True
Else
    fraExcluir.enabled = False
End If
If Trim(txtOrigen) = "E" Then
    chkExcluir = 1
Else
    chkExcluir = 0
End If
End Function

Private Sub inimenu(Optional EsVista As Boolean = False)
If EsVista Then
    uMenu.init False, False, False, False, False, , DataEnvironment1.Sistema
Else
    uMenu.init True, True, True, False, True, , DataEnvironment1.Sistema
End If
End Sub

Private Sub inigrilla()
    Set g = New LiGrilla
    With g
        .init griMayor
'        gIDDO = .AddCol("iddoc", "H")
        gCUEN = .AddCol(" Cuenta                    ", "S")
        gDESC = .AddCol(" Descripcion Cuenta                             ")
        gDEBE = .AddCol(" DEBE                      ", "N")
        gHABE = .AddCol(" HABER                     ", "N")
        gORIG = .AddCol(" Origen                                          ", "S")
        .rows = 30
    End With
End Sub

Private Sub Form_Resize()
    'encajar fraAsiento, Me, 0, 0, uMenu.Height, 0
    Anclar fraAsiento, Me, anclarLadosTodos
    'encajar griMayor, fraAsiento, 1800, 60, 60, 60
    Anclar griMayor, fraAsiento, anclarLadosTodos
'    Anclar lblSumaDebe, fraAsiento, anclarDerecha + anclarAbajo
'    Anclar lblSumaHaber, fraAsiento, anclarDerecha + anclarAbajo
End Sub

Private Sub Form_Unload(cancel As Integer)
    Set Asiento = Nothing
    Set g = Nothing
End Sub

Private Sub g_cambio(ByVal Row As Long, ByVal Col As Long, txt As String)

    If Row = 0 Then Exit Sub
    'select case col
    If Col = gCUEN Then
        If txt > "" Then
            g.tx Row, gDESC, DescrCuenta(txt)
        Else
            g.tx Row, gDESC, ""
        End If
    End If
    
    If Col = gDEBE Or Col = gHABE Then
        lblSumaDebe = g.suma(gDEBE)
        lblSumaHaber = g.suma(gHABE)
    End If

    'agrego lineas
    If g.rows < Row + 2 Then g.rows = g.rows + 1
End Sub

Private Sub g_DblClick()
If EsVista Then Exit Sub
    Dim i As Long
    Dim resu
    If g.Row > 0 And g.Col = gCUEN Then
        resu = frmBuscar.MostrarSql("select cuenta as [ Cuenta           ], Descripcion as [ Descripcion                                                    ] from cuentas where activo = 1  and imputable = 1 order by cuenta ")
        If resu > "" Then
'            For i = 1 To g.rows - 1
'                If Trim(resu) = Trim(g.TextMatrix(i, 0)) Then
'                    MsgBox "La cuenta ya esta ingresada.", vbInformation
'                    Exit Sub
'                End If
'            Next
            g.Text = resu
        End If
    End If
End Sub

Private Function DescrCuenta(cue As String)
'    DescrCuenta = sSinNull(obtenerDatoS("cuentas",  cue, "Descripcion"))
    DescrCuenta = sSinNull(obtenerDeSQL("select descripcion from cuentas where cuenta = '" & cue & "' "))
End Function

Private Sub gri2asiento()
    Dim i As Long
    Asiento.nuevo txtConcepto, uFecha.dtFecha, Trim(txtOrigen)
    For i = 1 To g.rows - 1
        'Asiento.AcumularItem Trim(g.tx(i, gCUEN)), s2n(g.tx(i, gDEBE)), s2n(g.tx(i, gHABE))
        If Trim(g.tx(i, gCUEN)) <> "" And (s2n(g.tx(i, gDEBE)) <> 0 Or s2n(g.tx(i, gHABE)) <> 0) Then
           'Asiento.AgregarItem Trim(g.tx(i, gCUEN)), s2n(g.tx(i, gDEBE)), s2n(g.tx(i, gHABE)), Trim(g.tx(i, gORIG))
           Asiento.AcumularItem Trim(g.tx(i, gCUEN)), s2n(g.tx(i, gDEBE)), s2n(g.tx(i, gHABE)), Trim(g.tx(i, gORIG))
        End If
    Next i
End Sub

Private Function BorrarAsientoActual() As Boolean
    If ON_ERROR_HABILITADO Then On Error GoTo UFAelim
    If s2n(lblId) = 0 Then Exit Function
    DataEnvironment1.Sistema.Execute "update asientos set activo = 0 where idAsiento = " & s2n(lblId)
'    daTaenvironment1.SISTEMA.Execute "delete from mayor  where idAsiento = " & s2n(lblId)
    BorrarAsientoActual = True
fin:
    Exit Function
UFAelim:
    ufa "Err al eliminar asiento", "id " & lblId
    Resume fin
End Function

Private Function EsManual() As Boolean
    EsManual = (Trim(txtOrigen) = STR_MOVMANUAL)
End Function

Private Function AsientoPipiCucu() As Boolean
Dim rquitar As Long, i As Integer
    If Trim(txtConcepto) = "" Or Trim(txtOrigen) = "" Then
        che "Falta llenar cabecera"
'    ElseIf s2n(g.suma(gHABE), 2) - s2n(g.suma(gDEBE), 2) <> 0 Then
    ElseIf g.suma(gHABE) - g.suma(gDEBE) <> 0 Then
        che "No cierra, " & vbCrLf & "Dif: " & g.suma(gHABE) - g.suma(gDEBE)
    ElseIf g.suma(gDEBE) = 0 Then
        che "no hay montos especificados"
    ElseIf MalRenglon() > 0 Then
        rquitar = MalRenglon()
        che "falta cuenta o monto en renglon " & rquitar
        If MsgBox("¿Desea quitarlo?", vbInformation + vbYesNo) = vbYes Then
            g.delRow rquitar
        End If
    Else
        AsientoPipiCucu = True
    End If
    
    For i = 1 To g.rows - 1
        If (g.TextMatrix(i, 0)) > "" Then
            If (s2n(g.TextMatrix(i, 2)) - s2n(g.TextMatrix(i, 3))) = 0 Then
                MsgBox "Error en la cuenta " & g.TextMatrix(i, 0) & ", Debe/Haber - No se permite el ingreso de datos en ambos campos para la misma cuenta.", vbCritical
                AsientoPipiCucu = False
                Exit For
            End If
        End If
    Next
    
End Function
Private Function MalRenglon() As Long
    Dim i As Long
    For i = 1 To g.rows - 1
        If (Trim(g.tx(i, gCUEN)) > "" And (s2n(g.tx(i, gDEBE)) + s2n(g.tx(i, gHABE))) = 0) _
          Or (Trim(g.tx(i, gCUEN)) = "" And (s2n(g.tx(i, gDEBE)) + s2n(g.tx(i, gHABE))) <> 0) _
                Then
                MalRenglon = i
        End If
    Next i
End Function

Private Sub g_Validar(ByVal Row As Long, ByVal Col As Long, cancel As Boolean)
    Dim temp
    If Row = 0 Or Col <> gCUEN Or g.EditText = "" Then Exit Sub
    
    temp = obtenerDeSQL("select descripcion from cuentas where cuenta  = '" & (g.EditText) & "' and activo = 1 and imputable = 1 ")
    cancel = Trim(sSinNull(temp)) = ""
End Sub

Private Sub txtConcepto_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtOrigen_GotFocus()
    PintoFocoActivo
End Sub

'******************** MENU **********************
Private Sub uMenu_AceptarAlta()
    If ON_ERROR_HABILITADO Then On Error GoTo UFAalta
    Dim x
    
    If Not AsientoPipiCucu() Then Exit Sub
    'DE_BeginTrans
    gri2asiento
    
    Dim EsActual
    EsActual = obtenerDeSQL("select activo,idejercicio,Denominacion,0 as bloqueado from ejercicio where fechainicio<=" & ssFecha(uFecha.dtFecha) & " and fechafin>=" & ssFecha(uFecha.dtFecha))
    
    If EsActual(3) = "True" Or EsActual(3) = "Verdadero" Then
        MsgBox "El movimiento que esta ingresando se encuentra fuera del ejercicio. Actualmente el ejercicio " & EsActual(1) & " con Denominación " & EsActual(2) & " se encuentra bloqueado", vbExclamation
        Exit Sub
    End If
    
    If EsActual(0) = "False" Or EsActual(0) = "Falso" Then
        If MsgBox("El movimiento que esta ingresando se encuentra fuera del ejercicio. De continuar, el movimiento se generara en el Ejercicio " & EsActual(1) & " con Denominación " & EsActual(2) & " ¿Desea Continuar?", vbYesNo) = vbNo Then
            GoTo fin
        End If
    End If
    
    If EsActual(0) = "False" Or EsActual(0) = "Falso" Then
        x = Asiento.Grabar(0, , s2n(EsActual(1)))
    Else
        x = Asiento.Grabar(0, , leerEjercicioId(cboEjercicio))
    End If
    
    'DE_BeginTrans
    'x = Asiento.Grabar(0, , leerEjercicioId(cboEjercicio))
    If x = 0 Then
        che "No se pudo grabar, revise datos"
        'DE_RollbackTrans
    Else
        'DE_CommitTrans
        che "grabado id= " & x
        uMenu.AceptarOk
    End If
fin:
    Exit Sub
UFAalta:
    'DE_RollbackTrans
    ufa "error al grabar alta", "id " & lblId
    Resume fin
End Sub
Private Sub uMenu_AceptarModi()
    If ON_ERROR_HABILITADO Then On Error GoTo ufamodi
    Dim x
    If Not AsientoPipiCucu Then Exit Sub
    gri2asiento
    If s2n(lblId) = 0 Then
        ufa "err programa - no modificable", "falta id"
        Exit Sub
    End If
    
    DE_BeginTrans
    If BorrarAsientoActual() Then
    
    
    
        Dim EsActual
        EsActual = obtenerDeSQL("select activo,idejercicio,Denominacion,0 as bloqueado from ejercicio where fechainicio<=" & ssFecha(uFecha.dtFecha) & " and fechafin>=" & ssFecha(uFecha.dtFecha))
        
        If EsActual(3) = "True" Or EsActual(3) = "Verdadero" Then
            MsgBox "El movimiento que esta ingresando se encuentra fuera del ejercicio. Actualmente el ejercicio " & EsActual(1) & " con Denominación " & EsActual(2) & " se encuentra bloqueado", vbExclamation
            Exit Sub
        End If
        
        If EsActual(0) = "False" Or EsActual(0) = "Falso" Then
            If MsgBox("El movimiento que esta ingresando se encuentra fuera del ejercicio. De continuar, el movimiento se generara en el Ejercicio " & EsActual(1) & " con Denominación " & EsActual(2) & " ¿Desea Continuar?", vbYesNo) = vbNo Then
                GoTo fin
            End If
        End If
        
        'x = Asiento.Grabar(CLng(lblIDDOC), , leerEjercicioId(cboEjercicio))
        x = Asiento.Grabar(CLng(lblIdDoc), , s2n(EsActual(1)))
        If x > 0 Then
            che "grabado, nuevo id= " & x
            DE_CommitTrans
            uMenu.AceptarOk
        Else
            che "No se pudo grabar, revise datos"
            DE_RollbackTrans
        End If
    Else
        DE_RollbackTrans
    End If
fin:
    Exit Sub
ufamodi:
    DE_RollbackTrans
    ufa "error al grabar modificacion", "id " & lblId
    Resume fin
End Sub
Private Sub uMenu_BorrarControles()
    On Error Resume Next
    FrmBorrarTxt Me
    lblId = ""
    lblIdDoc = ""
    lblAsiento = ""
    lblSumaDebe = ""
    lblSumaHaber = ""
    g.Borrar
    g.rows = 2
End Sub
Private Sub uMenu_Buscar()
    Dim resu
    Dim idej As Long
    Dim WhereFecha As String
    With frmBuscar
        idej = leerEjercicioId(cboEjercicio) 'leerEjercicioId()
        WhereFecha = " and a.ejercicio=" & idej & " and a.fecha " & ucFechas.ssBetween()
        resu = frmBuscar.MostrarSql("select a.NroAsiento, cast(a.Fecha as datetime) as Fecha, a.Concepto, rtrim(a.Origen) as [Origen      ],e.Denominacion as Ej, a.idAsiento as ID, a.IDdoc from (asientos as a inner join ejercicio e on a.ejercicio=e.idejercicio) where a.activo = 1   " & WhereFecha & " order by a.NroAsiento desc")
        If resu > "" Then
            uMenu.AceptarOk
            If CargaAsiento(s2n(frmBuscar.resultado(6))) = True Then
                uMenu.BuscarOK
            End If
        End If
    End With
End Sub
Private Sub uMenu_eliminar()
    If s2n(lblId) = 0 Then
        che "Nada que borrar"
        Exit Sub
    End If

    If Not EsManual() Then
        If USUARIO_SABE_LO_QUE_HACE Then
            che "Asiento generado por sistema. No se puede borrar" & " -o no deberia- "
            If Not confirma("Seguro desea eliminar?") Then Exit Sub
        Else
            che "Asiento generado por sistema. No se puede borrar"
            Exit Sub
        End If
    End If
    
    If BorrarAsientoActual() Then uMenu.EliminarOK
End Sub

Private Sub uMenu_HabilitarEdicion(sino As Boolean)
    fraAsiento.enabled = sino
End Sub

Private Sub uMenu_Modificar()
        
    If Not EsManual() Then
        If USUARIO_SABE_LO_QUE_HACE Then
            che "Asiento generado por sistema, " & vbCrLf & "los cambios se perderan si se modifica o anula el documento que lo genero"
        Else
            che "Asiento generado por sistema, no modificable "
            Exit Sub
        End If
    End If
    
    txtConcepto.SetFocus
End Sub

Private Sub uMenu_Nuevo()
    lblEjercicio = leerEjercicioDenominacion()
    txtOrigen = STR_MOVMANUAL
    uFecha.ini ucHoy
    PermitoExcluir
End Sub

Private Sub uMenu_SALIR()
    Unload Me
End Sub

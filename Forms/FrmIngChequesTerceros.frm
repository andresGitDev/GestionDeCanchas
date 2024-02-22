VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmIngChequesTerceros 
   Caption         =   "Ingreso de Cheques de Terceros"
   ClientHeight    =   7785
   ClientLeft      =   90
   ClientTop       =   345
   ClientWidth     =   8025
   Icon            =   "FrmIngChequesTerceros.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   7785
   ScaleWidth      =   8025
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboEjercicio 
      Height          =   315
      Left            =   2640
      TabIndex        =   31
      Text            =   "Ejercicio"
      Top             =   0
      Width           =   990
   End
   Begin Gestion.ucCoDe uBanco 
      Height          =   300
      Left            =   1305
      TabIndex        =   0
      Top             =   405
      Width           =   6405
      _extentx        =   11298
      _extenty        =   529
      codigowidth     =   700
   End
   Begin Gestion.ucCoDe uClie 
      Height          =   330
      Left            =   1290
      TabIndex        =   30
      Top             =   1095
      Width           =   6360
      _extentx        =   11218
      _extenty        =   582
      codigowidth     =   1000
   End
   Begin VB.TextBox txtnumero 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1260
      TabIndex        =   2
      Top             =   780
      Width           =   2235
   End
   Begin VB.ComboBox cmbprocedencia 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "FrmIngChequesTerceros.frx":08CA
      Left            =   1260
      List            =   "FrmIngChequesTerceros.frx":08D4
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1500
      Width           =   2415
   End
   Begin VB.TextBox txtconcepto 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1260
      TabIndex        =   4
      Top             =   1920
      Width           =   6375
   End
   Begin VB.TextBox txtimporte 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1260
      TabIndex        =   5
      Top             =   2340
      Width           =   1215
   End
   Begin VB.Frame fraContable 
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   0
      TabIndex        =   23
      Top             =   2820
      Width           =   8175
      Begin Gestion.ucCoDe uCuenta 
         Height          =   330
         Left            =   1290
         TabIndex        =   8
         Top             =   45
         Width           =   6330
         _extentx        =   10874
         _extenty        =   582
         codigowidth     =   1000
      End
      Begin VSFlex7LCtl.VSFlexGrid grilla 
         Height          =   2115
         Left            =   120
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1140
         Width           =   5595
         _cx             =   9869
         _cy             =   3731
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
         Cols            =   4
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
      Begin VB.TextBox txttotal 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5880
         TabIndex        =   24
         TabStop         =   0   'False
         Tag             =   "8"
         Top             =   2940
         Width           =   1095
      End
      Begin VB.TextBox txtvalor 
         Height          =   285
         Left            =   1260
         TabIndex        =   10
         Top             =   780
         Width           =   1335
      End
      Begin VB.TextBox txtconc 
         Height          =   285
         Left            =   1260
         TabIndex        =   9
         Top             =   420
         Width           =   5655
      End
      Begin VB.CommandButton cmdcargar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5880
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1140
         Width           =   975
      End
      Begin VB.CommandButton cmbeliminofila 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Eliminar Fila"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5880
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1620
         Width           =   975
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "TOTAL"
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
         Height          =   255
         Left            =   6240
         TabIndex        =   28
         Top             =   2700
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Concepto:"
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
         Left            =   120
         TabIndex        =   27
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Importe:"
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
         Left            =   120
         TabIndex        =   26
         Top             =   780
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Cuenta:"
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
         Left            =   120
         TabIndex        =   25
         Top             =   60
         Width           =   795
      End
   End
   Begin Gestion.ucBotonera uMenu 
      Align           =   2  'Align Bottom
      Height          =   1545
      Left            =   0
      TabIndex        =   1
      Top             =   6240
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   2725
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   ""
   End
   Begin VB.TextBox interno 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Tag             =   "2"
      Top             =   60
      Width           =   1035
   End
   Begin VB.TextBox cargar 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8760
      TabIndex        =   21
      TabStop         =   0   'False
      Tag             =   "2"
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComCtl2.DTPicker fechacheque 
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   2340
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   62193665
      CurrentDate     =   38052
   End
   Begin MSComCtl2.DTPicker fechaingreso 
      Height          =   255
      Left            =   6300
      TabIndex        =   7
      Top             =   2340
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   62193665
      CurrentDate     =   38052
   End
   Begin VB.Label Label34 
      Caption         =   "Ejercicio"
      Height          =   255
      Left            =   3720
      TabIndex        =   32
      Top             =   60
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   0
      X2              =   8160
      Y1              =   2700
      Y2              =   2700
   End
   Begin VB.Label Label7 
      Caption         =   "F. Ingreso:"
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
      Left            =   5280
      TabIndex        =   20
      Top             =   2340
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "F. Cheque:"
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
      Left            =   2760
      TabIndex        =   19
      Top             =   2340
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Importe:"
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
      Left            =   60
      TabIndex        =   18
      Top             =   2340
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Concepto:"
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
      Left            =   60
      TabIndex        =   17
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Procedencia:"
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
      Left            =   60
      TabIndex        =   16
      Top             =   1500
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente:"
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
      Left            =   60
      TabIndex        =   15
      Top             =   1140
      Width           =   735
   End
   Begin VB.Label Label14 
      Caption         =   "Banco:"
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
      Left            =   60
      TabIndex        =   14
      Top             =   420
      Width           =   735
   End
   Begin VB.Label Label12 
      Caption         =   "Cheque Nº:"
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
      Left            =   0
      TabIndex        =   13
      Top             =   780
      Width           =   1095
   End
End
Attribute VB_Name = "FrmIngChequesTerceros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' mod 12/4/5

Private midDoc As Long

'Private Sub cmbbanco_Click()
'    FrmHelp.Show
'    Call CargarHelp("BancosGrales", "Codigo", "Descripcion", "Codigo", "Descripcion", "Codigo")
'    FrmHelp.Tag = Me.Name
'    cargar = "Bancos"
'End Sub

'Private Sub cmbcliente_Click()
'    FrmHelp.Show
'    CargarHelp "Clientes", "Codigo", "Descripcion", "codigo", "descripcion"
'    FrmHelp.Tag = Me.Name
'    cargar = "Clientes"
'End Sub

'Private Sub cmbcuenta_Click()
'    FrmHelp.Show
'    CargarHelpCuentas "Cuentas", "Codigo", "Descripcion", "codigo", "descripcion"
'    FrmHelp.Tag = Me.Name
'    cargar = "Cuentas"
'End Sub


Private Function TaTodoGrilla() As Boolean
    If s2n(txtvalor) = 0 Or uCuenta.codigo = "" Or Trim$(txtconc) = "" Then
        che "Faltan datos"
        Exit Function
    End If
    TaTodoGrilla = True
End Function

Private Sub cmdBack_Click()
'Dim rsmov As New ADODB.Recordset
'Dim sConsul As String, i As Long
'If s2n(interno) > 0 Then
'    rsmov.Open "select * from movibanc where operacion<>'I' and interno=" & interno, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
'    With rsmov
'        If .EOF And .BOF Then
'            MsgBox "El cheque no tiene movimientos...de todas maneras se colocara en cartea...", vbInformation
'            sConsul = "update cheques set estado='C' where nroint=" & interno
'            DataEnvironment1.Sistema.Execute sConsul
'        Else
'            .MoveFirst
'            For i = 0 To .RecordCount - 1
'                sConsul = "update asientos set activo=0 where iddoc=" & !iddoc
'                DataEnvironment1.Sistema.Execute sConsul
'                .MoveNext
'            Next
'            sConsul = "update cheques set estado='C' where nroint=" & interno
'            DataEnvironment1.Sistema.Execute sConsul
'            sConsul = "update movibanc set activo=0 where operacion<>'I' and interno=" & interno
'            DataEnvironment1.Sistema.Execute sConsul
'            MsgBox "Se borraron los asientos, movimientos y el cheque fue colocado en cartera...", vbInformation
'        End If
'    End With
'    Set rsmov = Nothing
'Else
'    MsgBox "Seleccione un cheque...", vbInformation
'End If
End Sub

'Private Sub cmbbanco_Click()
'
'End Sub

'Private Sub cmbCliente_Click()
'
'End Sub

Private Sub cmdcargar_Click()
    On Error Resume Next
    Dim Valor As Double
    Dim totalgrilla As Double

    If Not TaTodoGrilla() Then Exit Sub
    grilla.AddItem uCuenta.codigo & Chr(9) & uCuenta.DESCRIPCION & Chr(9) & txtconc & Chr(9) & txtvalor
    Limpiotextosgrilla
    sumogrilla
    uCuenta.SetFocus
End Sub

Function sumogrilla() As Double
    Dim x As Long
    Dim Total As Double
    Total = 0
    For x = 1 To grilla.rows - 1
        Total = Total + s2n(grilla.TextMatrix(x, 3))
    Next
    txttotal = s2n(Total)
    sumogrilla = Total
End Function

Private Sub MuestroGrilla()
    'txtcuentacod = grilla.TextMatrix(grilla.row, 0)
    uCuenta.codigo = grilla.TextMatrix(grilla.Row, 0)
'    txtcuenta = grilla.TextMatrix(grilla.row, 1)
    txtconc = grilla.TextMatrix(grilla.Row, 2)
    txtvalor = grilla.TextMatrix(grilla.Row, 3)
End Sub

Private Sub Form_Load()
    uClie.ini "select descripcion from clientes where activo = 1 and codigo = '###' ", "Select codigo, descripcion as [ Cliente                           ] from clientes where activo = 1", False, True
    uBanco.ini "select descripcion from BancosGrales where codigo = ### ", "select codigo, descripcion as [ Banco                                              ] from BancosGrales where activo = 1"
    uMenu.init True, True, False, False, True
    uCuenta.ini "select descripcion from cuentas where cuenta = '###' and imputable = 1 and activo = 1", "select cuenta as [ Cuenta          ], descripcion as [ Descripcion                                   ] from cuentas where activo = 1 and imputable = 1 order by cuenta ", True
    
    Dim EjerA As New ADODB.Recordset
    EjerA.Open "select * from ejercicio", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    EjerA.MoveFirst
    While Not EjerA.EOF
        cboEjercicio.AddItem EjerA!denominacion 'EjerA!idejercicio
        EjerA.MoveNext
    Wend
    cboEjercicio = leerEjercicioDenominacion() ' mIdEjercicioActivo
    If UsuarioActual() <> 19 Then
        cboEjercicio.Visible = False
        Label34.Visible = False
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub

Private Sub grilla_DblClick()
    MuestroGrilla
    grilla.RemoveItem grilla.Row
End Sub

Private Sub txtconc_Click()
    If txtconcepto <> "" Then
        txtconc.Text = txtconcepto.Text
    End If
End Sub

'Private Sub txtcodbanco_Change()
'
'End Sub

'Private Sub txtcliente_Change()
'
'End Sub
'
'Private Sub txtcodbanco_GotFocus()
'    PintoFocoActivo
'End Sub

'Private Sub txtcodbanco_LostFocus()
'    If txtcodbanco <> "" Then
'        txtdesbanco = ObtenerDescripcion("BancosGrales", Val(txtcodbanco))
'        If txtdesbanco = "" Then
'            MsgBox "El Banco es incorrecto"
'            cmbbanco.SetFocus
'        End If
'    Else
'        MsgBox "El Banco es incorrecto"
'        cmbbanco.SetFocus
'    End If
'End Sub
'Private Sub txtcodbanco_KeyPress(KeyAscii As Integer)
'    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
'        KeyAscii = 0
'    End If
'End Sub

'Private Sub txtcodcli_Change()
'
'End Sub

Private Sub txtconc_GotFocus()
    PintoFocoActivo
End Sub
Private Sub txtConcepto_GotFocus()
    PintoFocoActivo
End Sub
'Private Sub txtcodcli_GotFocus()
'    PintoFocoActivo
'End Sub
'Private Sub txtcodcli_KeyPress(KeyAscii As Integer)
'    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
'        KeyAscii = 0
'    End If
'End Sub

'Private Sub txtdesbanco_Change()
'
'End Sub

'Private Sub txtcodcli_LostFocus()
'    If txtcodcli <> "" Then
'        txtcliente = ObtenerDescripcion("Clientes", Val(txtcodcli))
'        If txtcliente = "" Then
'            MsgBox "El Cliente es incorrecto"
'            cmbcliente.SetFocus
'        End If
'    Else
'        MsgBox "El Cliente es incorrecto"
'        cmbcliente.SetFocus
'    End If
'End Sub

'Private Sub txtcuentacod_KeyPress(KeyAscii As Integer)
'        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
'            KeyAscii = 0
'        End If
'End Sub

Private Sub txtimporte_GotFocus()
    txtimporte.SelStart = 0
    txtimporte.SelLength = Len(txtimporte.Text)
End Sub

Private Sub txtimporte_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtimporte_LostFocus()
    txtimporte = s2n(txtimporte)
    pTotalEnCheques = pTotalEnCheques + s2n(txtimporte)
End Sub

Private Sub cmbeliminofila_Click()
    If grilla.Row = 0 Then Exit Sub
    grilla.RemoveItem grilla.Row
End Sub

Private Sub Cargogrilla(interno As Long)
    Dim rs1 As New ADODB.Recordset
    
    If midDoc = 0 Then Exit Sub
    
    'rs1.Open "select DetalleMovcajas.* from DetalleMovcajas inner join Movicaja on Movicaja.movimiento = DetalleMovcajas.movimiento where MoviCaja.Tipo = 'C' and  Movicaja.interno = " & interno & "", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
    rs1.Open "select * from mayor inner join asientos on asientos.idasiento = mayor.idasiento where iddoc = " & midDoc & " and haber > 0 ", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    If Not rs1.EOF Then
        InicioGrilla
'        habilitogrilla (True)
        grilla.rows = 1
        'grilla.row = 0
        While Not rs1.EOF
            grilla.AddItem rs1!Cuenta & Chr(9) & obtenerDeSQL("select descripcion from cuentas where cuenta = '" & (rs1!Cuenta) & "' ") & Chr(9) & rs1!concepto & Chr(9) & rs1!haber
        
'            grilla.row = grilla.row + 1
'            grilla.TextMatrix(grilla.row, 0) = rs1!Cuenta
'            grilla.TextMatrix(grilla.row, 1) = ObtenerDescripcion("Cuentas", Val(rs1!Cuenta))
'            grilla.TextMatrix(grilla.row, 2) = rs1!concepto
'            grilla.TextMatrix(grilla.row, 3) = rs1!importe
'            If txttotal <> "" Then
'                txttotal = s2n(txttotal) + s2n(rs1!importe)
'            Else
'                txttotal = s2n(rs1!importe)
'            End If
            rs1.MoveNext
'            If Not rs1.EOF Then
'                grilla.rows = grilla.rows + 1
'            End If
        Wend
    End If
    Set rs1 = Nothing
End Sub

Private Sub Limpiotextosgrilla()
    uCuenta.clear
    'txtcuentacod = ""
'    txtcuenta = ""
    txtconc = ""
    txtvalor = ""
End Sub

Public Sub CargarDatos()
    Dim rs As New ADODB.Recordset
    Dim codigo As Long

    codigo = Val(Trim(Me.Tag))
       
'    If cargar = "Cuentas" Then
'        If txtcuentacod = "" Then
'            txtcuentacod = Trim(str(codigo))
'        End If
'        If Not noestaenlagrillaVS(txtcuentacod, grilla) And esimputable(txtcuentacod) Then
'            rs.Open "select * from Cuentas where codigo = " & codigo & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
'            If Not rs.EOF Then
'                txtcuentacod = rs!codigo
'                txtcuenta = rs!descripcion
'            End If
'            rs.Close
'            Set rs = Nothing
'        Else
'            MsgBox "El concepto ya se encuentra cargado"
'            txtcuentacod = ""
'            txtcuentacod.SetFocus
'        End If
'    End If
'
'    If cargar = "Bancos" Then
'        rs.Open "select * from BancosGrales where codigo = " & codigo & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
'        If Not rs.EOF Then
'            txtcodbanco = rs!codigo
'            txtdesbanco = rs!descripcion
'        End If
'        rs.Close
'        Set rs = Nothing
'    End If
'
'    If cargar = "Clientes" Then
'        rs.Open "select * from Clientes where codigo = " & codigo & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
'        If Not rs.EOF Then
'            txtcodcli = rs!codigo
'            txtcliente = rs!descripcion
'        End If
'        rs.Close
'        Set rs = Nothing
'    End If
        
    If cargar = "Cheques" Then
        rs.Open "select * from Cheques where Nroint = " & codigo & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
            uBanco.codigo = rs!BANCO_NRO
            'txtdesbanco = ObtenerDescripcion("BancosGrales", rs!banco_nro)
            txtnumero = rs!Nro
            uClie.codigo = rs!cliente
'            txtcliente = ObtenerDescripcion("Clientes", rs!cliente)
            cmbprocedencia = IIf(rs!procedencia = "P", "PROPIO", "TERCERO")

            txtimporte = rs!Importe
            fechacheque = rs!Fecha
            fechaingreso = rs!fecha_ingr
        
            interno = rs!NroInt
            InicioGrilla

            
            midDoc = nSinNull(obtenerDeSQL("select iddoc_ingreso from cheques where nroint = " & Val(Trim(Me.Tag))))
            Cargogrilla (rs!NroInt)
            
'           CONCEPTO LO CARGO ALCARGAR LA GRILLA YA QUE AHI TENGO EL CONCEPTO
            'mentira, ni aparece!!
            If midDoc = 0 Then
                txtconcepto = ""
            Else
                txtconcepto = obtenerDeSQL("select descripcion from movibanc where iddoc=" & midDoc)
            End If


            uMenu.BuscarOK
            
        End If
    End If
    
End Sub

Sub InicioGrilla()
    grilla.clear

    grilla.TextMatrix(0, 0) = "Cuenta"
    grilla.TextMatrix(0, 1) = "Descripción"
    grilla.TextMatrix(0, 2) = "Concepto"
    grilla.TextMatrix(0, 3) = "Importe"
    grilla.rows = 1
End Sub


Sub HabilitoControles(habilito As Boolean)
'    txtcodbanco.enabled = habilito
    uBanco.enabled = habilito
    txtnumero.enabled = habilito
    'txtcodcli.Enabled = habilito
    uClie.enabled = habilito
    cmbprocedencia.enabled = habilito
    txtconcepto.enabled = habilito
    txtimporte.enabled = habilito
    fechacheque.enabled = habilito
    fechaingreso.enabled = habilito
'    cmbbanco.enabled = habilito
'    cmbcliente.Enabled = habilito
End Sub

Sub LimpioControles()
    uBanco.clear
'    txtcodbanco = ""
'    txtdesbanco = ""
    txtnumero = ""
'    txtcodcli = ""
'    txtcliente = ""
    uClie.clear
    cmbprocedencia.ListIndex = 0
    txtconcepto = ""
    txtimporte = ""
    fechacheque = Date
    fechaingreso = Date
'    txtcuentacod = ""
'    txtcuenta = ""
    txtconc = ""
    txtvalor = ""
    txttotal = "0"
    cargar = ""
    interno = ""
End Sub

Private Sub txtNumero_GotFocus()
    txtnumero.SelStart = 0
    txtnumero.SelLength = Len(txtnumero.Text)
End Sub

Private Sub txtvalor_Click()
    If txtimporte <> "" Then
        txtvalor.Text = txtimporte.Text
    End If
End Sub

Private Sub txtvalor_GotFocus()
    txtvalor.SelStart = 0
    txtvalor.SelLength = Len(txtvalor.Text)
End Sub
Private Sub txtvalor_KeyPress(KeyAscii As Integer)
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
End Sub
Private Sub txtvalor_LostFocus()
    txtvalor = s2n(txtvalor)
End Sub

Private Function TaTodo() As Boolean
    If s2n(txtimporte) = 0 Then
        che "Falta Importe "
        Exit Function
    End If
    If s2n(txtnumero) = 0 Then
        che "Falta Numero"
        Exit Function
    End If
    'If s2n(txtcodbanco) = 0 Then
    If uBanco.codigo = 0 Then
        che "Falta Banco"
        Exit Function
    End If
    If Trim$(txtnumero) = "" Then
        che "Falta Nro Cheque"
        Exit Function
    End If
'    If s2n(txtcodcli) = 0 Then
'        che "falta cliente"
'        Exit Function
'    End If
    
    If gEMPR_ConSistContable Then
        If s2n(txtimporte) <> s2n(txttotal) Then
            che "importes cheque y total grilla conciden"
            Exit Function
        End If
    End If
    TaTodo = True
End Function





'------------------------------------------
Private Sub uMenu_AceptarAlta()
    If ON_ERROR_HABILITADO Then On Error GoTo UFAalta
    
    If Not TaTodo() Then Exit Sub
    
    Dim rs As New ADODB.Recordset
    Dim maximobanc1 As Long, maxcheque As Long, maximocaja As Long, x As Long, valcartera As Long
    Dim valorcuentacon1 As String, valorcuentacon2 As String, enletras

    
    Dim asie As New Asiento, iddoc As Long
    
    
'    rs.Open "select max(movbanco) as maxcodigo from MOVIBANC", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'    If Not IsNull(rs!maxcodigo) Then
'        maximobanc1 = rs!maxcodigo + 1
'    Else
'        maximobanc1 = 1
'    End If
'    rs.Close
'    Set rs = Nothing

    maximobanc1 = nuevoCodigo("movibanc", "movbanco")
    maxcheque = nuevoCodigo("cheques", "nroint")
    maximocaja = nuevoCodigo("movicaja", "movimiento")
'    rs.Open "select max(nroint) as maxcodigo from Cheques", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'    If Not IsNull(rs!maxcodigo) Then
'        maxcheque = rs!maxcodigo + 1
'    Else
'        maxcheque = 1
'    End If
'    rs.Close
'    Set rs = Nothing
        
'    rs.Open "select max(movimiento) as maxcodigo from MOVICAJA", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'    If Not IsNull(rs!maxcodigo) Then
'        maximocaja = rs!maxcodigo + 1
'    Else
'        maximocaja = 1
'    End If
'    rs.Close
'    Set rs = Nothing
        
'    rs.Open "select valores_cartera from Imputaciones", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'    If Not rs.EOF Then
'        valcartera = rs!valores_cartera
'    Else
'        valcartera = 0
'    End If
'    rs.Close
'    Set rs = Nothing
        
        
    DE_BeginTrans
        iddoc = NuevoDocumento("ch3", nuevoCodigo("RegistroDocumentos", "Nrodoc", " tipodoc = 'ch3' "), 0, 0)
        midDoc = iddoc
        asie.nuevo "Ingreso cheque " & uClie.DESCRIPCION, fechaingreso, "ch3"
        asie.AgregarItem CuentaParam(ID_Cuenta_M_CH_CARTERA), s2n(txtimporte), 0
        
    
        DataEnvironment1.dbo_INGCHEQUEMOVIBANC "A", 0, "I", Trim(txtconcepto), fechacheque, "C" _
            , maxcheque, s2n(txtimporte), maximobanc1, iddoc, Date, UsuarioSistema!codigo
                                               
        DataEnvironment1.dbo_INGCHEQUEMOVICAJA "A", 0, maximocaja, "C", "I", s2n(txtimporte), Trim(txtconcepto) _
            , fechaingreso, maxcheque, valcartera, maximobanc1, iddoc, Date, UsuarioSistema!codigo
                   
        For x = 1 To grilla.rows - 1
'''''            DataEnvironment1.dbo_INGCHEQUEDETALLE "A", maximocaja, s2n(grilla.TextMatrix(x, 3)), Val(txtcodcli), Val(grilla.TextMatrix(x, 0)), IIf(Trim(grilla.TextMatrix(x, 2)) <> "", Trim(grilla.TextMatrix(x, 2)), Trim(txtConcepto)) _
                , "IC", fechaingreso
            
            'viejo
            DataEnvironment1.dbo_DETALLEGTOSTEMP "A", Val(grilla.TextMatrix(x, 0)), Trim(grilla.TextMatrix(x, 1)), Trim(grilla.TextMatrix(x, 2)), s2n(grilla.TextMatrix(x, 3))
            'nuevo
            asie.AgregarItem grilla.TextMatrix(x, 0), 0, s2n(grilla.TextMatrix(x, 3))
        Next
    
        DataEnvironment1.dbo_INGCHEQUESTERCEROS "A", maxcheque, fechacheque, txtnumero, s2n(txtimporte), maximocaja _
            , "CAJ", fechaingreso, fechaingreso, "C", uBanco.codigo, IIf(Trim(cmbprocedencia) = "PROPIO", "P", "T"), uClie.codigo _
            , Date, UsuarioSistema!codigo, iddoc, 0
            
        asie.Grabar iddoc, , leerEjercicioId(cboEjercicio)
    DE_CommitTrans
    
    On Error GoTo ufaImpr
        
    enletras = NroEnLetras(s2n(txtimporte))
        
    'DataEnvironment1.LisChequesTerceros
    'rptChequesTerceros.Sections("Encabezado").Controls("lblEmpresa").caption = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
    'rptChequesTerceros.Sections("Encabezado").Controls("lblnumero").caption = maxcheque
    'rptChequesTerceros.Sections("Encabezado").Controls("lblnumcheque").caption = txtnumero
    'rptChequesTerceros.Sections("Encabezado").Controls("lblconcepto").caption = txtconcepto
    'rptChequesTerceros.Sections("Encabezado").Controls("lblcodbanco").caption = uBanco.descripcion 'ObtenerDescripcion("BancosGrales", Val(txtcodbanco))
    'rptChequesTerceros.Sections("Encabezado").Controls("lblcliente").caption = uClie.codigo
    'rptChequesTerceros.Sections("Encabezado").Controls("lblprocedencia").caption = cmbprocedencia
    'rptChequesTerceros.Sections("Encabezado").Controls("lblfechacheque").caption = fechacheque
    'rptChequesTerceros.Sections("Encabezado").Controls("lblfechaingreso").caption = fechaingreso
    'rptChequesTerceros.Sections("Encabezado").Controls("lblmovcaja").caption = maximocaja
    'rptChequesTerceros.Sections("Encabezado").Controls("lblmovbanco").caption = maximobanc1
    
    'rptChequesTerceros.Sections("Medio").Controls("lblenpesos").caption = enletras
    'rptChequesTerceros.Sections("Medio").Controls("lbltotoperacion").caption = txtimporte

    'rptChequesTerceros.Show vbModal
    'DataEnvironment1.rsLisChequesTerceros.Close
    
    CheqTer
    Set rptCheques3.DataControl1.Recordset = RStraer
    'rptCheques3.Field6.Text = RStraer.RecordCount
    rptCheques3.Label26.caption = Date
    rptChequesTerceros.Sections("Encabezado").Controls("lblEmpresa").caption = VerParametro(BS_NOMBRE_EMPRESA_CORTO)
    rptCheques3.lblNumero.caption = maxcheque
    rptCheques3.lblnumcheque.caption = txtnumero
    rptCheques3.lblConcepto.caption = txtconcepto
    rptCheques3.lblcodbanco.caption = uBanco.DESCRIPCION 'ObtenerDescripcion("BancosGrales", Val(txtcodbanco))
    rptCheques3.lblcliente.caption = uClie.codigo
    rptCheques3.lblprocedencia.caption = cmbprocedencia
    rptCheques3.lblfechacheque.caption = fechacheque
    rptCheques3.lblfechaingreso.caption = fechaingreso
    rptCheques3.lblmovcaja.caption = maximocaja
    rptCheques3.lblmovbanco.caption = maximobanc1
    
    rptCheques3.lblenpesos.caption = enletras
    rptCheques3.lbltotoperacion.caption = txtimporte
    rptCheques3.Show vbModal
    Set RStraer = Nothing

    DataEnvironment1.dbo_DETALLEGTOSTEMP "B", 0, "", "", 0
    
    uMenu.AceptarOk
    
fin:
    Exit Sub
UFAalta:
    DE_RollbackTrans
    ufa "Err al dar alta", "alta cheque"
    Resume fin
ufaImpr:
    ufa "Fallo la impresion, el movimiento fue grabado", "impresion"
    uMenu.AceptarOk
    Resume fin
End Sub
Private Sub uMenu_BorrarControles()
    InicioGrilla
    LimpioControles
    FrmBorrarTxt Me
    midDoc = 0
End Sub

Private Sub uMenu_Buscar()
    cargar = "Cheques"
    FrmHelp.Show
    CargarHelpChequesTerceros "Cheques", "Nro. Interno", "Nro. Cheque", "Banco - Importe", "Nroint", "Nro", "Banco_Nro", "Importe", "Nroint"
    FrmHelp.Tag = Me.Name
    
End Sub

Private Sub uMenu_BuscarYa(que As Variant)
'
End Sub

Private Sub uMenu_eliminar()
    If ON_ERROR_HABILITADO Then On Error GoTo UFAelim
    Dim tempCH, tempMC, tempMB
    Dim asse As String, nInterno As Long
    
    nInterno = s2n(interno.Text)
    
    tempCH = obtenerDeSQL("select estado     from cheques  where NroInt = " & nInterno)
    tempMC = obtenerDeSQL("select movimiento from MOVICAJA where tipo = 'C' and interno = " & nInterno)
    tempMB = obtenerDeSQL("select movBanco   from movibanc where interno = " & nInterno)
    
    If IsEmpty(tempCH) Then
        ufa "prg: err al eliminar. no se encuentra cheque", "eliminar " & nInterno
        Exit Sub
    ElseIf tempCH <> Cheque_CARTERA Then
        che "Cheque no esta en cartera"
        Exit Sub
    End If
    
    
    DE_BeginTrans
        
        BorroDocumento midDoc
        
        asse = "cheques " 'cheques
        DataEnvironment1.dbo_INGCHEQUESTERCEROS "B", Val(interno), 0, "", 0, 0, "", 0, 0, "", 0, "", 0, Date, UsuarioSistema!codigo, 0, 0
        
        asse = "movibanc" 'movibanc
        DataEnvironment1.dbo_INGCHEQUEMOVIBANC "B", 0, "", "", 0, "", 0, 0, tempMB, midDoc, Date, UsuarioSistema!codigo
        
        asse = "movicaja" 'movicaja
        DataEnvironment1.dbo_INGCHEQUEMOVICAJA "B", 0, s2n(tempMC), "", "", 0, "", 0, 0, "", 0, midDoc, Date, UsuarioSistema!codigo
        
        'asse = "movcjdet" 'movcjdet detalleMovCajas
'''''        DataEnvironment1.dbo_INGCHEQUEDETALLE "B", s2n(tempMC), 0, 0, 0, "", "", 0
        
        asse = "bitacora"
        grabaBitacora "B", nInterno, "cheques, mb, mc, mcd "
    DE_CommitTrans
    uMenu.EliminarOK
'        daTaenvironment1.dbo_INGCHEQUEMOVIBANC "B", 0, "", "", 0, "", 0, 0, 0, 0, 0, Date, UsuarioSistema!Codigo
'        daTaenvironment1.dbo_INGCHEQUEMOVICAJA "B", 0, 0, "", "", 0, "", 0, Val(interno), "", 0, 0, 0, Date, UsuarioSistema!Codigo
'        If Not IsEmpty(tempMC) Then
'            daTaenvironment1.dbo_INGCHEQUEDETALLE "B", s2n(tempMC), 0, 0, 0, "", "", 0
'        End If
'        daTaenvironment1.dbo_INGCHEQUESTERCEROS "B", Val(interno), 0, "", 0, 0, "", 0, 0, "", 0, "", 0, 0, 0, Date, UsuarioSistema!Codigo
'        daTaenvironment1.dbo_GRABARBITACORA Val(interno), "Cheques", UsuarioSistema!Codigo, Date, Time, "B"
fin:
    Exit Sub
UFAelim:
    ufa "prg: err en la eliminacion", "cheques 3ros " & asse
    DE_RollbackTrans
    Resume fin
End Sub

Private Sub uMenu_HabilitarEdicion(sino As Boolean)
    HabilitoControles sino
    fraContable.enabled = gEMPR_ConSistContable And sino
End Sub

Private Sub uMenu_Nuevo()
grilla.rows = 1
End Sub

Private Sub uMenu_SALIR()
    Unload Me
End Sub
'------------------------------------------

'3/12/4 rempl fechas string x date
'12/4/5 anule contable PENDIENTE para q no de err en locaire
'      reformulo menu, simplifico, etc
'15/4/5
'   rearmo eliminacion: pregunta si en cartera, fix baja movibanc


VERSION 5.00
Begin VB.Form FrmSeries 
   BackColor       =   &H00C0C0C0&
   Caption         =   "CARGA DE SERIES"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9930
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5130
   ScaleWidth      =   9930
   StartUpPosition =   2  'CenterScreen
   Begin Gestion.ucFecha uDesde 
      Height          =   315
      Left            =   2940
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   4320
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   556
      FechaInit       =   0
   End
   Begin VB.CommandButton cmdSeriesEnStock 
      Caption         =   "?"
      Height          =   315
      Left            =   4920
      TabIndex        =   4
      Top             =   600
      Width           =   375
   End
   Begin VB.CheckBox chkEsSalida 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Es Salida:"
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
      Left            =   6540
      TabIndex        =   6
      Top             =   2400
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdNumero 
      Caption         =   "?"
      Height          =   315
      Left            =   4920
      TabIndex        =   8
      Top             =   1680
      Width           =   375
   End
   Begin VB.CommandButton cmdayudaprod 
      BackColor       =   &H00FFFFFF&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   375
   End
   Begin VB.CheckBox chkconsignacion 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Consignacion"
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
      Left            =   6540
      TabIndex        =   10
      Top             =   2880
      Width           =   1575
   End
   Begin VB.ComboBox cmbconcepto 
      Height          =   315
      Left            =   1620
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Tag             =   "5"
      Top             =   2400
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.TextBox txtobs 
      Height          =   645
      Left            =   1620
      TabIndex        =   11
      Tag             =   "14"
      Top             =   3240
      Width           =   8055
   End
   Begin VB.TextBox txtnrocomp 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2880
      TabIndex        =   7
      Tag             =   "12"
      Top             =   1680
      Width           =   1935
   End
   Begin VB.ComboBox cmbcomprobante 
      Height          =   315
      Left            =   1620
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Tag             =   "5"
      Top             =   1260
      Width           =   3195
   End
   Begin VB.TextBox txtserie 
      Height          =   285
      Left            =   1620
      TabIndex        =   3
      Tag             =   "3"
      Top             =   600
      Width           =   3195
   End
   Begin VB.TextBox txtproducto 
      Height          =   285
      Left            =   1620
      TabIndex        =   1
      Tag             =   "2"
      Top             =   240
      Width           =   3195
   End
   Begin VB.CommandButton cmdmodificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Modificar"
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4740
      Width           =   975
   End
   Begin VB.CommandButton cmdbuscar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Buscar"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4740
      Width           =   975
   End
   Begin VB.CommandButton cmdeliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Eliminar"
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4740
      Width           =   975
   End
   Begin VB.CommandButton cmdaceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4740
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
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4740
      Width           =   975
   End
   Begin VB.CommandButton cmdnuevo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Nuevo"
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
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4740
      Width           =   975
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
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4740
      Width           =   975
   End
   Begin VB.CommandButton cmdprimero 
      BackColor       =   &H00C0C0C0&
      Height          =   555
      Left            =   4200
      Picture         =   "FrmSeries.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Primero"
      Top             =   4140
      Width           =   615
   End
   Begin VB.CommandButton cmdultimo 
      BackColor       =   &H00C0C0C0&
      Height          =   555
      Left            =   6105
      Picture         =   "FrmSeries.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Ultimo"
      Top             =   4140
      Width           =   615
   End
   Begin VB.CommandButton cmdsiguiente 
      BackColor       =   &H00C0C0C0&
      Height          =   555
      Left            =   5490
      Picture         =   "FrmSeries.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Siguiente"
      Top             =   4140
      Width           =   615
   End
   Begin VB.CommandButton cmdanterior 
      BackColor       =   &H00C0C0C0&
      Height          =   555
      Left            =   4815
      Picture         =   "FrmSeries.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Anterior"
      Top             =   4140
      Width           =   675
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Comprobantes con producto seleccionado"
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
      Left            =   5340
      TabIndex        =   30
      Top             =   1680
      Width           =   3915
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Series en stock:"
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
      Left            =   5340
      TabIndex        =   29
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Busca Comprobantes Desde:"
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
      Index           =   1
      Left            =   60
      TabIndex        =   28
      Top             =   4380
      Width           =   2775
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Observaciones:"
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
      Index           =   0
      Left            =   180
      TabIndex        =   27
      Top             =   3300
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Numero:"
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
      Left            =   720
      TabIndex        =   26
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Comprobante:"
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
      Left            =   240
      TabIndex        =   25
      Top             =   1260
      Width           =   1395
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Producto:"
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
      Left            =   600
      TabIndex        =   24
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Serie:"
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
      Left            =   900
      TabIndex        =   23
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblConcepto 
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   660
      TabIndex        =   22
      Top             =   2400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   3945
      Left            =   120
      Top             =   120
      Width           =   9720
   End
End
Attribute VB_Name = "FrmSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit ' mod 12/8/4


Dim Ope As String
Dim rsserie As New ADODB.Recordset


Private Sub cmbcomprobante_LostFocus()
    Dim conConcepto As Boolean
    conConcepto = (ComboCodigo(cmbcomprobante) = 7) 'Or ComboCodigo(cmbcomprobante) = 8)
    cmbconcepto.Visible = conConcepto
    lblConcepto.Visible = conConcepto
End Sub

Private Sub cmdAceptar_Click()
    If ON_ERROR_HABILITADO Then On Error GoTo UfaGraba
    
    'Dim fecha As Variant
    Dim rs As New ADODB.Recordset
    Dim codigo As Long, nTipoCompr, essalida As Integer, n, tmpFecha
    Dim consig As Integer
    
    If FaltaAlgo() Then Exit Sub

    Call HabilitoControles(False, False, False, True, False, True)
    
    nTipoCompr = ObtenerCodigo2("Tipocomprobantesgrales", Trim(cmbcomprobante.Text))
    n = nTipoCompr
    'perdon lo trucho !!!!!!!!!!!!! me urgia, por eso
    'If InStr(trim(str(n)), "12568") > 0 Then EsSalida = 1
    If n = 1 Or n = 2 Or n = 5 Or n = 7 Or n = 8 Then essalida = 1
    'perdon lo trucho !!!!!!!!!!!!! me urgia, por eso
    
'    If Trim(txtserie) = "" Then
'        MsgBox "Debe cargar la Serie", 48, "Atencion"
'        txtserie.SetFocus
'        Exit Sub
'    Else
'        If  Then
'            consig = 1
'        Else
'            consig = 0
'        End If
        
        consig = IIf(chkconsignacion.Value = vbChecked, 1, 0)
        
        '**************************
        DE_BeginTrans
        If Ope = "A" Then
            rs.Open "Select max(codigo) as cod from Series", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
            'If Not rs.EOF Then  ' boludez
                codigo = nSinNull(rs!cod) + 1 '
            'Else
            '    Codigo = 1
            'End If
            rs.Close
            Set rs = Nothing
'***********************
'            DataEnvironment1.dbo_SERIE "A", Codigo, Trim(txtproducto), Trim(txtserie), nTipoCompr, Val(Trim(txtnrocomp)), 1, ObtenerCodigo("conceptos", Trim(cmbconcepto.Text)), Trim(txtobs), consig, Date, UsuarioSistema!Codigo, 0, 0
            DataEnvironment1.dbo_abmSERIEs "A", codigo, Trim(txtproducto), Trim(txtserie), nTipoCompr, Val(Trim(txtnrocomp)), 1, ObtenerCodigo("conceptos", Trim(cmbconcepto.Text)), Trim(txtObs), consig, Date, essalida, Date, UsuarioSistema!codigo
'***********************
            MsgBox "La Operación se ha realizado con éxito", 48, "Atencion"
        Else
            If Ope = "M" Then
'***********************
'                DataEnvironment1.dbo_SERIE "M", rsserie!Codigo, Trim(txtproducto), Trim(txtserie), ObtenerCodigo("Tipocomprobantesgrales", Trim(cmbcomprobante.Text)), Val(Trim(txtnrocomp)), RsParam!sucursal, ObtenerCodigo("conceptos", Trim(cmbconcepto.Text)), Trim(txtobs), consig, 0, 0, 0, 0
                DataEnvironment1.dbo_abmSERIEs "M", rsserie!codigo, Trim(txtproducto), Trim(txtserie), nTipoCompr, Val(Trim(txtnrocomp)), RsParam!sucursal, ObtenerCodigo("conceptos", Trim(cmbconcepto.Text)), Trim(txtObs), consig, Date, essalida, Date, UsuarioActual()
'***********************
                DataEnvironment1.dbo_GRABARBITACORA rsserie!codigo, "Series", UsuarioSistema!codigo, Date, Time, "M"
                MsgBox "La Operación se ha realizado con éxito", 48, "Atencion"
            End If
        End If
        DE_CommitTrans
        '****************************
        
        LimpioTxt
        HabilitoTxt (True)
        Call HabilitoControles(False, False, False, True, False, True)
'    End If
fin:
    Exit Sub
UfaGraba:
    DE_RollbackTrans
    ufa "err al grabar", txtproducto & " " & txtserie & " " & cmbcomprobante & " " & txtnrocomp
    Resume fin
End Sub

Private Function FaltaAlgo() As Boolean
    Dim Nro 'As Long
    
    FaltaAlgo = True
    
    If Trim(txtproducto) = "" Or Trim(txtserie) = "" Or Trim(txtnrocomp) = "" Then
        che "faltan datos"
        Exit Function
    End If
    
    Nro = Trim(txtnrocomp)
    Select Case ComboCodigo(cmbcomprobante)
    Case 1
        If ssFV("FAA", Nro) = "" Then
            che "No figura el comprobante o no tiene el producto seleccionado"
            Exit Function
        End If
    Case 2
        If ssFV("FAB", Nro) = "" Then
            che "No figura el comprobante o no tiene el producto seleccionado"
            Exit Function
        End If
    
    Case 3
        If ssFV("NCA", Nro) = "" Then
            che "No figura el comprobante o no tiene el producto seleccionado"
            Exit Function
        End If
    
    Case 4
        If ssFV("NCB", Nro) = "" Then
            che "No figura el comprobante o no tiene el producto seleccionado"
            Exit Function
        End If
    
    Case 5
        If ssRV(Nro) = "" Then
            che "No figura el comprobante o no tiene el producto seleccionado"
            Exit Function
        End If
    
    Case 6
        If ssRC(Nro) = "" Then
            che "No figura el comprobante o no tiene el producto seleccionado"
            Exit Function
        End If
    
    Case 7
        If ssDif(Nro) = "" Then
            che "No figura el comprobante o no tiene el producto seleccionado"
            Exit Function
        End If
    
    'Case 8
        
    End Select
    
    FaltaAlgo = False
End Function

Sub CargoRegistro()
    txtproducto = rsserie!producto
    txtserie = rsserie!Serie
    If rsserie!COMPROBANTE <> 0 Then
        cmbcomprobante.ListIndex = BuscarenComboS(cmbcomprobante, ObtenerDescripcion("TipoComprobantesGrales", rsserie!COMPROBANTE))
    Else
        cmbcomprobante.ListIndex = -1
    End If
    If rsserie!NroComprobante <> 0 Then
        txtnrocomp = rsserie!NroComprobante
    End If
    If rsserie!concepto <> 0 Then
        cmbconcepto.ListIndex = BuscarenComboS(cmbconcepto, ObtenerDescripcion("Conceptos", rsserie!concepto))
    Else
        cmbconcepto.ListIndex = -1
    End If
    If rsserie!consignacion = True Then
        chkconsignacion.Value = 1
    Else
        chkconsignacion.Value = 0
    End If
    If Not IsNull(rsserie!observaciones) Then
        txtObs = rsserie!observaciones
    Else
        txtObs = ""
    End If
End Sub
Private Sub cmdanterior_Click()
    rsserie.MovePrevious
    If Not rsserie.BOF Then
        
        CargoRegistro
        Call HabilitoBotonesMoverse(True, True, True, True)
    Else
        Call HabilitoBotonesMoverse(False, False, True, True)
    End If
End Sub

Private Sub cmdayudaprod_Click()
'    FrmHelp.Show
'    CargarHelp "producto", "Codigo", "Descripcion", "codigo", "descripcion"
'    FrmHelp.Tag = "Prodseries"
    
    Dim resu As String
    resu = frmBuscar.MostrarCodigoDescripcionActivo("producto")
    If resu > "" Then
        txtproducto = frmBuscar.resultado(1)
    End If

End Sub

Private Sub cmdBuscar_Click()
'    Call HabilitoControles(True, False, True, False, True, False)
'
'    FrmHelp.Show
'    CargarHelp "series", "Producto", "Serie", "producto", "serie"
'    FrmHelp.Tag = Me.Name

    Dim resu As String
    resu = frmBuscar.MostrarSql("select Producto as [  Producto                      ], Serie as [  Serie                ] from series where activo = 1 and fecha>=" & ssFecha(uDesde.strFecha))
    If resu > "" Then
        txtproducto = resu
        txtserie = frmBuscar.resultado(2)
        CargarDatos
        Call HabilitoControles(True, False, True, False, True, False)
    End If
End Sub

Private Sub cmdCancelar_Click()
    Call HabilitoControles(False, False, False, True, False, True)
    LimpioTxt
    HabilitoTxt (True)
    Call HabilitoBotonesMoverse(False, False, False, False)
End Sub

Private Sub cmdeliminar_Click()
'Dim fecha As Variant
    Dim mensaje As String
    mensaje = MsgBox("Esta seguro de querer eliminar este registro", vbYesNo, "Atencion")
    If mensaje = 6 Then
        'DataEnvironment1.dbo_SERIE "B", rsserie!Codigo, "", "", 0, 0, 0, 0, "", 0, 0, 0, UsuarioSistema!Codigo, Date
        DataEnvironment1.dbo_abmSERIEs "B", rsserie!codigo, "", "", 0, 0, 0, 0, "", 0, 0, 0, Date, UsuarioSistema!codigo
        
        DataEnvironment1.dbo_GRABARBITACORA rsserie!codigo, "series", UsuarioSistema!codigo, Date, Time, "B"
        Call HabilitoControles(True, True, False, False, False, False)
        LimpioTxt
        HabilitoTxt (True)
        Call HabilitoControles(False, False, False, True, False, True)
    End If
End Sub

Private Sub cmdNumero_Click()
    If txtproducto = "" Then Exit Sub
    
    Select Case ComboCodigo(cmbcomprobante)
    Case 1
        ssFV "FAA"
    Case 2
        ssFV "FAB"
    Case 3
        ssFV "NCA"
    Case 4
        ssFV "NCB"
    Case 5 ' rv
        ssRV
    Case 6 ' rc
        ssRC
    Case 7 'dif
        ssDif
'    Case 8 'can
'        ssCan
'    Case Else
    End Select
    
End Sub

'Private Function ssCan(Optional nro) As String
'    Dim ss As String
'    ss = "   "
'
'    If frmBuscar.MostrarSql(ss) > "" Then txtnrocomp = frmBuscar.resultado
'End Function

Private Function ssDif(Optional Nro) As String
    Dim ss As String
    If IsMissing(Nro) Then
        ss = "SELECT DISTINCT r.MovimientoInterno, r.fecha, d.producto " _
            & " FROM ItemRemitoDiferenciaStock AS d INNER JOIN RemitoDiferenciaStock AS r ON d.numero = r.MovimientoInterno " _
            & " WHERE r.fecha > " & uDesde.ConvertFecha & " AND d.producto = '" & txtproducto & "' " _
            & " order by MovimientoInterno desc "
        
        If frmBuscar.MostrarSql(ss) > "" Then txtnrocomp = frmBuscar.resultado
    Else
        ss = "SELECT r.MovimientoInterno " _
            & " FROM ItemRemitoDiferenciaStock AS d INNER JOIN RemitoDiferenciaStock AS r ON d.numero = r.MovimientoInterno " _
            & " WHERE d.producto = '" & txtproducto & "' and MovimientoInterno = '" & Nro & "' "
    
        ssDif = sSinNull(obtenerDeSQL(ss))
    End If
End Function

Private Function ssRC(Optional Nro) As String
    Dim ss As String
    If IsMissing(Nro) Then
        ss = "SELECT DISTINCT r.NroRemito, r.Fecha, d.producto " _
             & " FROM RemitoCompraDetalle AS d INNER JOIN RemitoCompra AS r ON d.CodigoRemito = r.codigo " _
             & " WHERE r.Fecha > " & uDesde.ConvertFecha & "  AND r.activo = 1 AND d.producto = '" & txtproducto & "' " _
             & " order by r.NroRemito desc "
        
        If frmBuscar.MostrarSql(ss) > "" Then txtnrocomp = frmBuscar.resultado
    Else
        ss = "SELECT  r.NroRemito " _
             & " FROM RemitoCompraDetalle AS d INNER JOIN RemitoCompra AS r ON d.CodigoRemito = r.codigo " _
             & " WHERE r.activo = 1 AND d.producto = '" & txtproducto & "' and NroRemito = '" & Nro & "' "
    
        ssRC = sSinNull(obtenerDeSQL(ss))
    End If

End Function

Private Function ssRV(Optional Nro) As String
    Dim ss As String
    If IsMissing(Nro) Then
        ss = "SELECT DISTINCT r.Numero, d.Producto, r.Fecha " _
            & " FROM RemitoVentaDetalle AS d INNER JOIN RemitoVenta AS r ON d.Numero = r.Numero " _
            & " WHERE r.Fecha > " & uDesde.ConvertFecha & "  AND r.Anulado = 0 and d.producto = '" & txtproducto & "' " _
            & " order by r.numero desc "
           
        If frmBuscar.MostrarSql(ss) > "" Then txtnrocomp = frmBuscar.resultado
    Else
        ss = "SELECT  r.Numero " _
            & " FROM RemitoVentaDetalle AS d INNER JOIN RemitoVenta AS r ON d.Numero = r.Numero " _
            & " WHERE r.Anulado = 0 and d.producto = '" & txtproducto & "' and r.numero = '" & Nro & "' "
        ssRV = sSinNull(obtenerDeSQL(ss))
    End If
End Function

Private Function ssFV(tipo As String, Optional Nro) As String
    Dim ss As String ', ss1 'As String
    If IsMissing(Nro) Then
         ss = "SELECT DISTINCT f.NroFactura, f.TipoDoc, d.Producto, f.Fecha " _
            & " FROM FacturaVentaDetalle AS d INNER JOIN FacturaVenta AS f ON d.CodigoFactura = f.Codigo " _
            & " Where f.TIPODOC = '" & tipo & "'  And d.producto = '" & txtproducto & "'  And f.fecha > " & uDesde.ConvertFecha & " " _
            & " ORDER BY f.NroFactura DESC "
        
         If frmBuscar.MostrarSql(ss) > "" Then txtnrocomp = frmBuscar.resultado
    Else
         ss = "SELECT  f.NroFactura " _
            & " FROM FacturaVentaDetalle AS d INNER JOIN FacturaVenta AS f ON d.CodigoFactura = f.Codigo " _
            & " Where f.TIPODOC = '" & tipo & "'  And d.producto = '" & txtproducto & "'  And f.NroFactura = '" & Nro & "' "
        
        ssFV = (sSinNull(obtenerDeSQL(ss)))
    End If
End Function

Private Sub cmdPrimero_Click()
    rsserie.MoveFirst
    CargoRegistro
    Call HabilitoBotonesMoverse(False, False, True, True)
End Sub

Private Sub cmdSeriesEnStock_Click()
    Dim ss  As String
    If txtproducto > "" Then
        ss = Buscar_SeriesEnStock(txtproducto)
        If ss > "" Then
            txtserie = ss
        End If
    End If
End Sub

Private Sub cmdsiguiente_Click()
    rsserie.MoveNext
    If Not rsserie.EOF Then
        CargoRegistro
        Call HabilitoBotonesMoverse(True, True, True, True)
    Else
        Call HabilitoBotonesMoverse(True, True, False, False)
    End If
End Sub

Private Sub cmdUltimo_Click()
    rsserie.MoveLast
    CargoRegistro
    Call HabilitoBotonesMoverse(True, True, False, False)
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True

    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Unload(cancel As Integer)
'    If rsserie.State = 1 Then
'        rsserie.Close
        Set rsserie = Nothing
'    End If
End Sub

'Private Sub txtobs_Change()
'Dim i As Integer
'    txtobs.Text = UCase(txtobs.Text)
'    i = Len(txtobs.Text)
'    txtobs.SelStart = i
'End Sub

Private Sub txtobs_GotFocus()
    PintoFocoActivo
End Sub

'Private Sub txtserie_Change()
'Dim i As Integer
'    txtserie.Text = UCase(txtserie.Text)
'    i = Len(txtserie.Text)
'    txtserie.SelStart = i
'End Sub

Private Sub txtSerie_GotFocus()
'    txtserie.SelStart = 0
'    txtserie.SelLength = Len(txtserie.Text)
    PintoFocoActivo
End Sub
Private Sub txtproducto_GotFocus()

'    txtproducto.SelStart = 0
'    txtproducto.SelLength = Len(txtproducto.Text)
    PintoFocoActivo
End Sub
Private Sub txtnrocomp_GotFocus()
'    txtnrocomp.SelStart = 0
'    txtnrocomp.SelLength = Len(txtnrocomp.Text)
    PintoFocoActivo
End Sub

'Private Sub txtserie_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    End If
'End Sub
'Private Sub txtproducto_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    End If
'End Sub
'Private Sub txtobs_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    End If
'End Sub
Private Sub txtnrocomp_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    Else
        If KeyAscii < 47 Or KeyAscii > 57 Then
            KeyAscii = 0
        End If
'    End If
End Sub
'Private Sub cmbcomprobante_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    End If
'End Sub
'Private Sub cmbconcepto_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    End If
'End Sub

Private Sub cmdmodificar_Click()
    Call HabilitoControles(True, True, False, False, False, False)
    HabilitoTxt (False)
    txtproducto.SetFocus
    Ope = "M"
End Sub

Private Sub cmdnuevo_Click()
    Dim rsserie As New ADODB.Recordset
    
    Ope = "A"
    LimpioTxt
    HabilitoTxt (False)
    txtproducto.SetFocus
    Call HabilitoControles(True, True, False, False, False, False)
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Sub LimpioTxt()
    txtproducto = ""
    txtserie = ""
    txtObs = ""
    txtnrocomp = "0"
    cmbcomprobante.ListIndex = 0
    'cmbconcepto.ListIndex = 0
    chkconsignacion.Value = 0
End Sub
Sub HabilitoTxt(habilito As Boolean) ' habilito deshabilita!? AL REVES
    txtproducto.Locked = habilito
    txtserie.Locked = habilito
    txtObs.Locked = habilito
    txtnrocomp.Locked = habilito
    cmbcomprobante.Locked = habilito
    cmbconcepto.Locked = habilito
    chkconsignacion.enabled = Not habilito
    cmdayudaprod.enabled = Not habilito
    cmdSeriesEnStock.enabled = Not habilito
    cmdNumero.enabled = Not habilito
End Sub
Sub HabilitoControles(hab1 As Boolean, hab2 As Boolean, hab3 As Boolean, hab4 As Boolean, hab5 As Boolean, hab6 As Boolean)
    cmdcancelar.enabled = hab1
    cmdAceptar.enabled = hab2
    cmdeliminar.enabled = hab3
    cmdnuevo.enabled = hab4
    cmdmodificar.enabled = hab5
    cmdbuscar.enabled = hab6
End Sub
Sub HabilitoBotonesMoverse(hab1 As Boolean, hab2 As Boolean, hab3 As Boolean, hab4 As Boolean)
    cmdprimero.enabled = hab2
    cmdanterior.enabled = hab1
    cmdsiguiente.enabled = hab3
    cmdultimo.enabled = hab4
End Sub

Private Sub Form_Load()
'    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2

    comboSql cmbcomprobante, "select DescripcionUsuario, Codigo from TipoComprobantesGrales where codigo <> 8"
'    CargaCombo cmbcomprobante, "TipocomprobantesGrales", "descripcionusuario", "codigo", ""
    
    CargaCombo cmbconcepto, "Conceptos", "descripcion", "codigo", ""

    LimpioTxt
    HabilitoControles False, False, False, True, False, True
    HabilitoBotonesMoverse False, False, False, False
    HabilitoTxt True  'ESTA AL REVES!!!
End Sub

Sub CargarDatos()
    If rsserie.State = 1 Then
        rsserie.Close
        Set rsserie = Nothing
    End If
    rsserie.Open "select * from series where activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockReadOnly ', adLockOptimistic
    
    If Not rsserie.EOF Then
        rsserie.MoveFirst
        rsserie.Find "serie= '" & Trim(txtserie) & "'"
        CargoRegistro
        Call HabilitoBotonesMoverse(True, True, True, True)
    End If
End Sub

'Private Sub txtserie_LostFocus()
'    Dim rs As New ADODB.Recordset
'
'    rs.Open "Select * from series where producto='" & Trim(txtproducto) & "' and serie='" & Trim(txtSerie) & "'", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
'    If Not rs.EOF Then
'        MsgBox "La serie esta repetida,verifiquelo", 48, "Atencion"
'        txtSerie.SetFocus
'    End If
'    rs.Close
'    Set rs = Nothing
'End Sub

' 12/8/4 Lito
'   busq  frmHelp  pasa a frmBUSCAR
'   inhibo mov al cargar
'21/1/5
'   fix fecha string
'23/2/5
'   deshabilite busqueda hast q veamos como limitar... 30000 registros en grilla...???
'15/3/5
'   deshabilito edicion en load
'26/5/5
'   big mod

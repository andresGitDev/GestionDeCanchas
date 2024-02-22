VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmTipoCompras 
   Caption         =   "Parametrizacion Cuentas "
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10290
   Icon            =   "frmTipoCompras.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   10290
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optVerInactivo 
      Caption         =   "Ver Inactivos"
      Height          =   315
      Left            =   8940
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   60
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Height          =   3855
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   10095
      _cx             =   17806
      _cy             =   6800
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
   Begin Gestion.ucBotonera uMenu 
      Align           =   2  'Align Bottom
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   4440
      Width           =   10290
      _extentx        =   18150
      _extenty        =   2778
      msgconfirmaeliminar=   ""
      msgconfirmacancelar=   ""
   End
   Begin VB.Label Label1 
      Caption         =   "Doble clic en CUENTA para ayuda"
      Height          =   255
      Left            =   5880
      TabIndex        =   3
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmTipoCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' pendiente  VALIDAR
' help cuentas

Option Explicit ' mod 15/6/5
'15/6/5 new sin SP

Private WithEvents g As LiGrilla
Attribute g.VB_VarHelpID = -1
Private gIDTC As Long
Private gCODI As Long
Private gDESC As Long
Private gCUEN As Long
Private gCUED As Long
Private gPROG As Long
Private gACTI As Long
Private gUSOC As Long
Private gUSOD As Long
'Private aUso


Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True
End Sub
Private Sub Form_Load()
    inigrilla

    uMenu.init False, False, True, False, False
    cargaGrilla
'    aUso = Array("Sistema", "Fac Compra", "Fac Venta", "Retenciones")
    Form_Resize
End Sub

Private Sub inigrilla()
'    Dim s, ss
'
'    For Each s In aUso
'        ss = ss & "s" & "|"
'    Next
    
    
    Set g = New LiGrilla
    With g
        .init Grilla
        gIDTC = .AddCol("id", "H")
        gCODI = .AddCol(" Codigo ", "S")
        gDESC = .AddCol(" Descripcion                             ", "S")
        gCUEN = .AddCol(" Cuenta            ", "S")
        gCUED = .AddCol(" Desc Cuenta                               ")
        gPROG = .AddCol("prg", "H")
        gUSOC = .AddCol("usoC", "H")
        gUSOD = .AddCol(" Uso en                ", "H") ', "B", UsoCuenta_STRING) ' "Sistema|Fac Compra|Fac Venta|Retenciones")
        gACTI = .AddCol(" Activo ", "K") ', "H")
        .rows = 100
    End With
End Sub
Private Sub cargaGrilla()
    On Error Resume Next
    Dim rs As New ADODB.Recordset, i As Long, s As String
    
'    If optVerInactivo.Value Then
        's = "select id, codigo, descripcion, cuenta, sistema, activo from CuentasParam order by sistema, codigo"
'        s = "select * from CuentasParam order by sistema, usocuenta, codigo"
'    Else
        's = "select id, codigo, descripcion, cuenta, sistema, activo from CuentasParam Where activo = 1 order by sistema, codigo"
        s = "select * from CuentasParam Where activo = 1 order by sistema, usocuenta, codigo"
'    End If
        
    With rs
        .Open s, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        i = 0
        g.Borrar
        g.rows = 100
        While Not .EOF
            i = i + 1
            g.tx i, gIDTC, !ID
            g.tx i, gCODI, !codigo
            g.tx i, gDESC, !DESCRIPCION
            g.tx i, gCUEN, !Cuenta
            g.tx i, gPROG, IIf(!Sistema, "1", "")
            g.tx i, gUSOD, combo_Cod2Str(UsoCuenta_STRING, !UsoCuenta)
            g.tx i, gACTI, !activo
            .MoveNext
        Wend
    End With
    uMenu.BuscarOK
    Set rs = Nothing
End Sub

Private Sub Form_Resize()
    Anclar Grilla, Me, anclarLadosTodos
End Sub

Private Sub g_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = gCUEN Then Exit Sub ' CUENTA lo dejo siempre
    
    If g.tx(Row, gPROG) = "" Then           ' > "", modificable solo por programador
'        If s2n(g.tx(row, gIDTC)) > 0 Then   ' le dejo modificar: cuenta, desc, activo
'            If col = gDESC Or col = gACTI Then Exit Sub
'        Else                                ' le dejo modificar: cuenta, desc, activo, codigo
            If Col = gDESC Or Col = gACTI Or Col = gCODI Then Exit Sub
            If Col = gUSOD Then
                Exit Sub
            End If
'        End If
    End If
    
    Cancel = True
End Sub

Private Sub g_cambio(ByVal Row As Long, ByVal Col As Long, txt As String)
    If Col = gCUEN Then
        g.tx Row, gCUED, CuentaDescripcion(txt)
    End If
End Sub

Private Sub g_DblClick()
    If uMenu.estado <> ucbEditando Then Exit Sub
    If g.Col <> gCUEN Then Exit Sub
    'help cuenta IMPUTABLES
    Dim re As String
    re = BuscarCuenta(False, False)
    If re = "" Then Exit Sub
    g.Text = re
End Sub

Private Sub g_Validar(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    ' cuenta, imputable que exista
    Dim re As String
    If Col = gCUEN Then
        re = CuentaDescripcion(g.EditText, False, False)
        If re = "" Then
            che "no es cuenta activa o imputable"
            Cancel = True
            Exit Sub
        End If
    End If
    
    If Col = gUSOD Then
        'If Left(g.EditText, 3) <> "Fac" Then cancel = True
        Cancel = (g.EditText = "Sistema")
    End If
    
    ' codigo,  nuevo no repetido
    If Col = gCODI Then
        '
    End If
End Sub

Private Function TodoOk() As Boolean
    Dim i As Long, j As Long, n As Long
    Dim s As String
    n = g.PrimerVacio(gCODI)
    For i = 1 To n - 1
        'repetido ?
        For j = 1 To n - 1
            If i <> j Then
                If (g.tx(i, gCODI) = g.tx(j, gCODI)) And (g.tx(i, gUSOD) = g.tx(j, gUSOD)) Then
                    che "Codigo repetido " & g.tx(i, gCODI) & " lineas " & i & " " & j
                    Exit Function
                End If
            End If
        Next
        'vacio
        If Trim(g.tx(i, gDESC)) = "" Or Trim(g.tx(i, gCUED)) = "" Then 'Or Trim(g.tx(i, gUSOD)) = "" Then
            che "faltan datos descripcion, cuenta o lugar de uso para " & g.tx(i, gCODI)
            If Not confirma("grabo igualmente ?") Then Exit Function
        End If
        'intento nuevo codigo
        If s2n(g.tx(i, gIDTC)) = 0 Then
            s = "%" & UCase(Trim(g.tx(i, gCODI)))
            If InStr("%ACC%ACD%FAA%FAB%FAC%FAE%NCA%NCB%NCN%NDA%RAA%RET%RGA%RIB%REC", s) > 0 Then
                che "El codigo " & g.tx(i, gCODI) & " ya esta usado por el sistema"
                Exit Function
            End If
        End If
    Next i
    
    'id sin codigo
    If n < g.PrimerVacio(gIDTC) Then ' no alcanza para verificar bien, pero limita la mayoria de probl
        che "no se puede eliminar un codigo"
        Exit Function
    End If
    
    TodoOk = True
End Function





Private Sub optVerInactivo_Click()
    g.Borrar
    cargaGrilla
End Sub

Private Sub uMenu_Aceptar()
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    
    Dim i As Long, ss As String
    Dim ID, co, de, cu, ac, us
    
    'If Not TodoOk() Then Exit Sub
    
    With g
        For i = 1 To g.PrimerVacio(gCODI) - 1
            ID = s2n(.tx(i, gIDTC))
            co = .tx(i, gCODI)
            de = .tx(i, gDESC)
            cu = .tx(i, gCUEN)
            ac = IIf(.tk(i, gACTI), 1, 0)
            us = combo_Str2Cod(UsoCuenta_STRING, .tx(i, gUSOD))
            If co > "" Then
                If ID > 0 Then
                    ss = "update CuentasParam set codigo = '" & co & "', descripcion = '" & de & "', cuenta = '" & cu & "', activo = " & ac & ", UsoCuenta = '" & us & "'   where id = " & x2s(ID)
                Else
                    ss = "insert into CuentasParam (codigo, descripcion, cuenta, usoCuenta, fecha_alta, usuario_alta ) " _
                        & " values ( '" & co & "', '" & de & "', '" & cu & "', '" & 0 & "', " & ssFecha(Date) & ", " & UsuarioActual() & ") "
                End If
                DataEnvironment1.Sistema.Execute ss
            End If
        Next i
    End With
    uMenu.AceptarOk
fin:
    Exit Sub
ufaErr:
    ufa "Err al grabar", "CuentasParam Ex-tipocompras " & Left(ss, 6) & " "
    Resume fin
End Sub

Private Sub uMenu_BorrarControles()
    cargaGrilla
End Sub

Private Sub uMenu_HabilitarEdicion(sino As Boolean)
    Grilla.Editable = IIf(sino, flexEDKbdMouse, flexEDNone)
End Sub
Private Sub uMenu_SALIR()
    Unload Me
End Sub


Function combo_Cod2Str(queString As String, quePosicion As Long)
    Dim p As Long, i As Long, j As Long
    
    If Left(queString, 1) <> "|" Then queString = "|" & queString
    For i = 1 To Len(queString)
        If Mid(queString, i, 1) = "|" Then p = p + 1
        If p = quePosicion Then
            For j = i + 1 To Len(queString)
                If Mid(queString, j, 1) = "|" Then Exit For
            Next j
            combo_Cod2Str = Mid(queString, i + 1, j - i - 1)
            Exit Function
        End If
    Next i
End Function
Function combo_Str2Cod(queString As String, str_aBuscar As String)
    Dim p, i, j
    If Left(queString, 1) <> "|" Then queString = "|" & queString
    j = InStr(queString, str_aBuscar)
    For i = 1 To Len(queString)
        If i > j Then Exit For
        If Mid(queString, i, 1) = "|" Then p = p + 1
    Next i
    combo_Str2Cod = p
End Function

'15/6/5 new, viejo abm es obsoleto, No uso SP


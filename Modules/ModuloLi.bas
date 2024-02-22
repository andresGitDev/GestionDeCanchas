Attribute VB_Name = "ModuloLi"
Option Explicit     '   ModuloLi gral
' Lito Explicit
'

  
' get para saber si debo reempl "." del tecl numerico por ","
Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Public Const LOCALE_SDECIMAL = &HE 'separador decimal
Public Const LOCALE_USER_DEFAULT = &H400 'presentar información del usuario
Public Const LOCALE_SSHORTDATE = &H1F 'formato de fecha corta

Public Const UFA_ARCHIVO_LOG = ".\Prg_Bug.log"   ' <-- Donde guarda los errores.
Public Const UFA_STOP = False                    ' <-- para q  ufa() haga un STOP (en diseño)  \'o'/  eh!
'

' ************** API **************************************

Public Function SeparadorDecimal() As String
    Dim sBuffer As String, lBufferLen As Long
    lBufferLen = 50
    sBuffer = Space(lBufferLen)
    If (GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDECIMAL, sBuffer, lBufferLen)) Then
        SeparadorDecimal = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
    Else
        SeparadorDecimal = ""
'        MsgBox "No se puedo obtener la información"
    End If
End Function
Public Function FormatoFecha() As String
    Dim sBuffer As String, lBufferLen As Long
    lBufferLen = 50
    sBuffer = Space(lBufferLen)
    If (GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SSHORTDATE, sBuffer, lBufferLen)) Then
        FormatoFecha = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
    Else
        FormatoFecha = ""
'        MsgBox "No se puedo obtener la información"
    End If
End Function

' *********************************************************

' **************  Conversion y tipos   ********************
Public Function esNumero(txt As String) As Boolean
    'OJO aca el vacio es numerico, en vez de error devuelve 0
    esNumero = Len(txt) = 0 Or IsNumeric(Trim(Replace(txt, ".", ",")))
End Function
'
Public Function x2s(ByVal x As Variant) ', Optional decimales As Long = 2)
'    ' Pasa TODO a string INGLES, en vez de error devuelve "", para sql, pone Punto "."
'    On Error Resume Next
'    x2s = ""
''    x = (Format(x, "#.#"))
'
'    If SeparadorDecimal = "," Then
'        x2s = Trim(Replace(CStr(Round(x, 2)), ",", "."))
'    Else
'        x2s = Trim(CStr(Round(x, 2)))
'    End If
    ' Pasa TODO a string, en vez de error devuelve "", para sql, pone Punto "."
    On Error Resume Next
    x2s = ""
    If SeparadorDecimal = "," Then
        x2s = Trim(Replace(CStr(x), ",", "."))
    Else
        x2s = Trim(CStr(x))
    End If
End Function

Public Function n2s(ByVal x As Variant)
    ' util para armar string SQL desde formato usuario
    ' Pasa TODO a string, en vez de error devuelve "", para sql, pone Punto "."
    On Error Resume Next
    n2s = "0"
    If SeparadorDecimal = "," Then
        n2s = Trim(Replace(CStr(x), ",", "."))
    Else
        n2s = Trim(CStr(x))
    End If
End Function
'
Public Function s2n(ByVal s As Variant, Optional Decimales = 2)
    ' Pasa TODO a numero, en vez de error devuelve 0'
    ' Pasa "." (el usuario siempre va a poner ".") a SeparadorDecimal(), usualmente ","
    On Error Resume Next
    'Dim tmp As Variant
    s2n = 0
    'If IsMissing(decimales) Then decimales = 2
    If SeparadorDecimal = "," Then s = Trim(Replace(s, ".", ","))
    If IsNumeric(s) Then s2n = Round(CDbl(s), Decimales)
End Function
'
Public Function n2r(ByVal n As Variant, Optional Decimales = 2) As String
    On Error Resume Next
    'If IsMissing(decimales) Then decimales = 2
    n2r = ""
    If Decimales > 0 Then n = Round(n, Decimales)
    n2r = Format(n, "0." & Left("000000", Decimales)) '
End Function
Public Function b2k(bule As Boolean) ' con valores de checkbox
    b2k = IIf(bule, vbChecked, vbUnchecked)
End Function
Public Function k2b(ChkValue As Long) As Boolean ' con valores de checkbox
    k2b = (ChkValue = vbChecked)
End Function
Public Function k2n(ChkValue As Long) As Long ' con valores de checkbox ' para bit sqlserver
    k2n = IIf(ChkValue = vbChecked, 1, 0)
End Function
Public Function n2k(bBit As Long) As Long ' con valores de checkbox ' para bit sqlserver
    n2k = IIf(bBit = 1, vbChecked, vbUnchecked)
End Function

' **************  Raros de Conversion y tipos   ********************


' **************  Grilla  VSFlexGrid ********************
Public Function grillaBuscoVacio(grilla As VSFlexGrid, Columna As Long) As Long
    'busca el primer vacio, si no hay, lo agrega
    
    Dim i As Long
   
    With grilla
        For i = 1 To .rows - 1
            If .TextMatrix(i, Columna) = "" Then
                grillaBuscoVacio = i
                Exit Function
            End If
        Next i
        .rows = .rows + 1
        grillaBuscoVacio = .rows - 1
    End With
End Function
'
Public Function strComboGrilla(sqlUnCampo) As String
    'arma string para el combo de la VSgrid
    If MODO_ON_ERROR_ABM_ON Then On Error GoTo ERRcarga
    
    Dim ss As String, rs As New ADODB.Recordset
    
    With rs
        .Open sqlUnCampo, DataEnvironment1.AMR, adOpenDynamic, adLockReadOnly
        If Not .EOF And Not .BOF Then
            .MoveFirst
            ss = ""
            While Not .EOF
                 ss = ss & Trim(.Fields(0))
                 .MoveNext
                If Not .EOF Then ss = ss & "|"
            Wend
        End If
    End With
    
    strComboGrilla = ss
   
    GoTo ERR_FIN
ERRcarga:
    ufa "Error al cargar grilla", "strcombogrilla" ', Err
ERR_FIN:
    Set rs = Nothing
End Function
' **************  Grilla  VSFlexGrid ********************


' **************  Manejo de Form ********************
Public Sub frmPintoFoco(f As Form) '*---'DEPRECATED ------------ PintoFocoActivo
    On Error Resume Next
    f.ActiveControl.SelStart = 0
    f.ActiveControl.SelLength = Len(f.ActiveControl)
End Sub
Public Sub FrmBorrarTxt(frm As Form)
    On Error Resume Next
    Dim co As Control
    For Each co In frm.Controls
        If TypeOf co Is TextBox Then
            co.Text = ""
        End If
    Next
End Sub
Public Sub FrmBorrarCbo(frm As Form)
    On Error Resume Next
    Dim co As Control
    For Each co In frm.Controls
        If TypeOf co Is ComboBox Then
            co.ListIndex = -1
            co.ListIndex = 0
        End If
    Next
End Sub

Public Sub iniFrmLoad(frm As Form) ' DEPRECATED --poco practico--
    ' no muy util, no lo aconsejo
    On Error Resume Next
    Dim i As Long
    CentrarMe frm
    frm.KeyPreview = True
'    For i = 0 To frm.Controls.Count
'        frm.Controls(i).BackColor = frm.BackColor
'      '''  if typeof controls(i) is  label ''''
'    Next i
End Sub


Public Sub FrmKeyPress(ByRef k As Integer, enter2tab As Boolean, Optional mayusculas As Boolean, Optional SinComodinesSQL As Boolean, Optional calculadora As Boolean)
    On Error Resume Next
    
    If enter2tab And k = 13 Then
        SendKeys "{tab}"
        k = 0
    ElseIf mayusculas And Chr(k) <> UCase(Chr(k)) Then
            k = Asc(UCase(Chr(k)))
    
   ElseIf k = 39 Then            'If SinComodinesSQL And InStr("'", Chr(k)) Then
        k = 180
    ElseIf k = 12 Then  ' control L
        calculin
    End If
    
End Sub
Public Sub CentrarMe(frmMe As Form) 'DEPRECATED- use propiedad .StartupPosition
    On Error Resume Next
    frmMe.Move (Screen.Width - frmMe.Width) \ 2, (Screen.Height - frmMe.Height) \ 2
End Sub

Public Function encajar(que As Object, donde As Object, Optional mT, Optional mL, Optional mB, Optional mR)
    ' para anclar un control dentro de otro,
    ' se pone en resize() del control padre o del form
    On Error Resume Next
    Dim oH As Object, oP As Object
    Dim oPh As Long, oPw As Long
    
    Set oH = que: Set oP = donde
    oPh = oP.Height
    oPh = oP.ScaleHeight - 100 'err
    oPw = oP.Width
    oPw = oP.ScaleWidth - 100
    
    If IsMissing(mL) And IsMissing(mR) Then
        oH.Left = (oP.Width - oH.Width) / 2
    ElseIf IsMissing(mL) Then
        oH.Left = (oP.Width - oH.Width) - mR
    ElseIf IsMissing(mR) Then
        oH.Left = mL
    Else
        oH.l = mL
        oH.Width = oPw - mR - mL 'oP.Width - mR - mL
    End If
    
    If IsMissing(mT) And IsMissing(mB) Then
        oH.Top = (oPh - oH.Height) / 2
    ElseIf IsMissing(mT) Then
        oH.Top = (oPh - oH.Height) - mB
    ElseIf IsMissing(mB) Then
        oH.Top = mT
    Else
        oH.Top = mT
        oH.Height = oPh - mT - mB
    End If
End Function

'Public Function alinear(que As Object, comoT_L_R_B_CV_CH As String)' SI INSISTO, USAR ENUM !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
' SI INSISTO, USAR ENUM !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'    On Error Resume Next
'    Dim oh As Object, oP As Object, t As long, l As long
'    Dim oPh As long
'
'    Set oh = que: Set oP = oh.Parent
'    oPh = oP.Height
'    oPh = oP.ScaleHeight - 100 'err
'
'    t = oh.Top
'    l = oh.Left
'    Select Case UCase(comoTLRB)' SI INSISTO, USAR ENUM !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'    Case "T"
'    Case "L"
'    Case "R"
'    Case "B"
'    End Select
'End Function
' **************  Manejo de Form ********************

' **************  Manejo de Controles ********************
Public Sub FrmPreviewPintoFocoActivo()
    On Error Resume Next
    Static co As Control
    With Screen.ActiveControl
        If .Name <> co.Name Then
            PintoFocoActivo
            Set co = Screen.ActiveControl
        End If
    End With
End Sub

Public Sub PintoFocoActivo()
    On Error Resume Next
    With Screen.ActiveControl
        .SelStart = 0
        .SelLength = Len(Screen.ActiveControl)
    End With
End Sub



Public Sub GotFocusPinto(que As Control) '*---'DEPRECATED ------- use PintoActivo() y FrmPreviewPintoActivo()
    On Error Resume Next
    que.SelStart = 0
    que.SelLength = Len(que)
End Sub

'Public Sub BorroControles(cole As Controls)
'    Dim co As Control
'
'    For Each co In Controls 'type
'
'        'if typeobject (  co) =
'    Next '
'End Sub
' **************  Manejo de Controles ********************

' **************  Control Combo  ********************
Public Function comboArray(cmb As ComboBox, arrayItems As Variant, Optional arrayItemData As Variant)
Dim que As Variant, i As Long ' no se que pasa si se autoordena OJO propiedad de sorted
    
    If IsMissing(arrayItemData) Then
        For Each que In arrayItems
            cmb.AddItem que
        Next
    Else
        For i = 0 To UBound(arrayItems)
            cmb.AddItem arrayItems(i)  ' ,i  ?pongo i?
            cmb.ItemData(i) = arrayItemData(i)
        Next i
    End If
    cmb.ListIndex = 0
End Function
'
Public Function ComboCodigo(Combo As ComboBox)
    Dim i
    i = Combo.ListIndex
    If i = -1 Then
        ComboCodigo = -1
    Else
        ComboCodigo = Combo.ItemData(Combo.ListIndex)
    End If
End Function
'
Public Function comboSql(cmb As ComboBox, ssql As String)
    Dim rs As New ADODB.Recordset, i As Long
    
    i = 0
    cmb.Clear
    With rs
        .Open ssql, DataEnvironment1.AMR, adOpenForwardOnly, adLockReadOnly
        While Not .EOF
        
            If .Fields.Count = 1 Then
                If Not IsNull(.Fields(0)) Then
                    cmb.AddItem .Fields(0), i
                    i = i + 1
                End If
            Else
                If Not IsNull(.Fields(1)) And Not IsNull(.Fields(0)) Then
                    cmb.AddItem .Fields(0), i
                    cmb.ItemData(i) = .Fields(1)
                    i = i + 1
                End If
            End If

            .MoveNext
        Wend
    End With
    If i > 0 Then cmb.ListIndex = 0
    
    Set rs = Nothing
End Function
' **************  Control Combo  ********************


Public Sub ufa(msgUsuario, msgLog, Optional NO_USAR_PARAMETRO_OBSOLETO)
'   Graba ERR en UFA_ARCHIVO_LOG (const definida en este modulo).
'   Si msgUsuario = "" , graba LOG pero no muestra MSGBOX;
'   Si UFA_stop =  TRUE, se para para debug; UFA_stop (const definida en este modulo).
'     Ergo, si "" y FALSE, graba ERR en LOG y usuario no se aviva. Salvo q no pueda grabar log.

    Dim nE, sE, Fe, ti, fo
    Dim ss, asse As String
    Dim cnx As New ADODB.Connection
    
    nE = Err.Number
    sE = Err.Description
    On Error GoTo ufaErr
    
'    sE = Error(nE)
    Fe = Format(Date, "dd/mm/yy")
    ti = Format(Time, "hh:mm:ss")
    fo = " Frm: " & NombreFormActivo() & " (usr " & UsuarioActual() & ") "
    ss = "Err " & nE & " " & sE
    'log
    
    asse = "1"
    Open UFA_ARCHIVO_LOG For Append Access Write As #1
    'Write #1, "Err " & nE & " " & sE, CStr(Fe), Time, msgUsuario, msgLog
    Write #1, Fe, ti, ss, msgUsuario, msgLog, fo
    Close #1
    
    asse = "2"
    cnx.Open DataEnvironment1.AMR.ConnectionString
    ss = Replace(ss, "'", " ")
    cnx.Execute "insert into bsuLog ( mE, mU, mL, mF ) values ( '" & Fe & " " & ti & " " & ss & "', '" & msgUsuario & "', '" & msgLog & "', '" & fo & "' )"
    cnx.Close
    'si debo aviso amablemente al user
    If msgUsuario > "" Then
        If nE > 0 Then
            che msgUsuario & vbCrLf & vbCrLf & "Err " & nE & vbCrLf & sE
        Else
            che msgUsuario
        End If
    End If
    
    If UFA_STOP Then Stop

fin:
    Exit Sub
ufaErr:
    If asse = "1" Then
        MsgBox "Falla acceso a Log, no se puede grabar ni el log de errores"
    ElseIf asse = "2" Then
        che "log grabado en disco, pero sin acceso a grabar en sql"
        che msgUsuario & vbCrLf & vbCrLf & "Err " & nE & vbCrLf & sE
    End If
    Resume fin
End Sub

Private Function NombreFormActivo() As String
    On Error Resume Next
    NombreFormActivo = Screen.ActiveForm.Name
End Function

' **************  SQL ********************
'
Public Function obtenerDato(Tabla As String, CodigoQueBuscar, campoResultado As String) As Variant
    ' Si CODIGO es string, pasarlo entre comillas
    'On Error GoTo FIN
    Dim rs As New ADODB.Recordset, ssql  As String
    
'    If IsNumeric(CodigoQueBuscar) Then
    ssql = "select " & campoResultado & " from " & Tabla & " where codigo = " & CodigoQueBuscar
'    Else
'        sSql = "select " & campoResultado & " from " & Tabla & " where codigo = '" & CodigoQueBuscar & "'"
'    End If

    rs.Open ssql, DataEnvironment1.AMR, adOpenDynamic, adLockReadOnly
    If Not rs.EOF Then
        rs.MoveLast
        obtenerDato = rs.Fields(0)
    End If
fin:
    Set rs = Nothing
End Function
'
Public Function obtenerDeSQL(SelectCampos As String) As Variant
    ' select de un campo devuelve lo q sea tipo de campo
    ' select de varios campos devuelve array, ubound(resultado) da cant de items

    'NO!!! 'On Error GoTo fin  -  Mejor q pare asi el err lo tratan desde do llaman
    Dim rs As New ADODB.Recordset, ssql  As String, i As Long, rc As Long, resu As Variant
    
    With rs
        'como falta la columna sucursal en la base de amrat uso esto que esta comentado para que arranque
        'SelectCampos = "select  Default_ProductoConSerie, EmiteFacturaConRemito, FormulaEsVirtual, FC_CargaCalculaIVA, Nombre  ,ImprimeCertCalidad, ConSistContable  from  DatosEmpresa where idEmpresa = 6"
        
        .Open SelectCampos, DataEnvironment1.AMR, adOpenDynamic, adLockReadOnly
        rc = .Fields.Count
        
        If rc = 1 Then
            If Not .BOF Then
                .MoveFirst
                obtenerDeSQL = .Fields(0)
            End If
        Else
            If Not .BOF Then
                .MoveFirst
               'obtenerDeSQL = rs.Fields(0)
               ReDim resu(rc)
                For i = 0 To rc - 1
                    resu(i) = .Fields(i)
                Next i
                obtenerDeSQL = resu
            End If

        End If
    End With
    
fin:
    Set rs = Nothing
End Function
'
Public Function obtenerDatoS(Tabla As String, CodigoQueBuscar, campoResultado As String) As Variant
'On Error GoTo FIN
    Dim rs As New ADODB.Recordset, ssql  As String
    
    ssql = "select " & campoResultado & " from " & Tabla & " where codigo = '" & CodigoQueBuscar & "'"

    rs.Open ssql, DataEnvironment1.AMR, adOpenDynamic, adLockReadOnly
    If Not rs.EOF Then
        rs.MoveLast
        obtenerDatoS = rs.Fields(0)
    End If
fin:
    Set rs = Nothing
End Function
'
Public Function obtenerDeS(NombreTabla As String, NombreCampoBuscar, NombreCampoResultado, CodigoQueBuscar) As Variant
'On Error GoTo FIN
    Dim rs As New ADODB.Recordset, ssql  As String
    
    ssql = "select " & NombreCampoResultado & " from " & NombreTabla & " where NombreCampoBuscar = '" & CodigoQueBuscar & "'"

    rs.Open ssql, DataEnvironment1.AMR, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then
        rs.MoveLast
        obtenerDeS = rs.Fields(0)
    End If
fin:
    Set rs = Nothing
End Function
'
'''Public Function nuevoCodigo(Tabla As String, Optional cpocodigo As String, Optional whe As String) As long
''''como ObtenerNuevoCodigo() de ModGral-Pero NO filtra inactivos
'''    Dim rs As New ADODB.Recordset
'''    Dim sSql As String
'''
'''    If cpocodigo = "" Then cpocodigo = "codigo"
'''
'''    sSql = "Select max (" & cpocodigo & ")  as NN From " & Tabla
'''    If whe > "" Then sSql = sSql & " where " & whe
'''
'''    rs.Open sSql, daTaenvironment1.amr, adOpenDynamic, adLockReadOnly
'''    If Not rs.EOF Then
'''        If IsNull(rs.Fields("NN")) Then
'''            nuevoCodigo = 1
'''        Else
'''            nuevoCodigo = rs.Fields("NN") + 1
'''        End If
'''    Else
'''        nuevoCodigo = 1
'''    End If
'''
'''    Set rs = Nothing
'''End Function

' string sql para concatenar en el where
' Para SqlServer
Public Function ssFecha(dFecha As Date) As String
    'METODO MSSQL SERVER seria
    Dim sFecha As String
    'sFecha = Month(dFecha) & "-" & Day(dFecha) & "-" & Year(dFecha)
    sFecha = Format(dFecha, "mm-dd-yy")
    ssFecha = " convert( datetime, '" & sFecha & "', 1) "
    'METODO LITO
'    ssFecha = " " & CLng(dFecha) - 2 & " "     ' mi favorito, algo desprolijo por la resta
End Function

Public Function ssFecha_new(dFecha As Date) As String
    ssFecha_new = " '" & Format(dFecha, "YYYYMMDD") & "' "
End Function

Public Function ssBetween(dDesde As Date, dHasta As Date) As String
    'METODO MSSQL SERVER seria (convierte vb y luego mssql)
    Dim s1 As String, s2 As String
'    s1 = Month(dDesde) & "-" & Day(dDesde) & "-" & Year(dDesde)
'    s2 = Month(dHasta) & "-" & Day(dHasta) & "-" & Year(dHasta)
    s1 = Format(dDesde, "mm-dd-yy")
    s2 = Format(dHasta, "mm-dd-yy")
    ssBetween = " between convert(datetime , '" & s1 & "', 1)  AND convert(datetime , '" & s2 & "', 1) "
   'METODO LITO (convierte solo vb)
'    ssEntreFechas = " between " & CLng(s1) - 2 & " AND " & CLng(s2) - 2 & " "
End Function
' para Access
Public Function asFecha(dFecha As Date) As String
    asFecha = " #" & Format(dFecha, "mm-dd-yyyy") & "# "
End Function
Public Function asBetween(dDesde As Date, dHasta As Date) As String
    asBetween = " between #" & Format(dDesde, "mm-dd-yyyy") & "# AND #" & Format(dHasta, "mm-dd-yyyy") & "# "
End Function
Public Function ssNum(ByVal numero)
    ' Pasa Numero a string, formato SQL, no usuario
    'On Error Resume Next -NO- Si falla, que falle
    ssNum = Trim(CStr(numero))
    If SeparadorDecimal = "," Then ssNum = Replace(ssNum, ",", ".")
End Function
' ************** SQL ********************

' ************** UTILITY ********************
Public Sub sumoEnCuenta(cuentas, Cuenta, monto) ' sorry, habre de mod para
    Dim que, i As Long, largo As Long

    'largo = UBound(cuentas, 0)
    'If monto = 0 Then Exit Sub ' asumo q lo suma en otra categoria?

    'For Each que In cuentas
    For i = 0 To 99
        If cuentas(i, 0) = Cuenta Then
            cuentas(i, 1) = cuentas(i, 1) + monto
            Exit Sub
        End If
    Next i

    'ReDim Preserve cuentas(largo + 1, 1)
    For i = 0 To 99
        If cuentas(i, 0) = 0 Then
            cuentas(i, 0) = Cuenta
            cuentas(i, 1) = monto
            Exit Sub
        End If
    Next i
    ufa "no pude agregar item", "sumoencuenta" ', Err
End Sub
' ************** UTILITY ********************

'**************** daTaenvironment1.amr ********************************** SACAR DE ACA!!!!
Public Function UsuarioActual()
    UsuarioActual = UsuarioSistema!codigo
End Function
Public Sub grabaBitacora(Abm, cod, Tabla)
    DataEnvironment1.dbo_GRABARBITACORA cod, Tabla, UsuarioSistema!codigo, Date, Time, Abm
End Sub
Public Function dEnvTxt()
    Dim i
    Open ".\data_environment.log" For Output Access Write As #1
    For i = 1 To DataEnvironment1.Commands.Count - 1
        Write #1, DataEnvironment1.Commands.Item(i).Name
    Next i
    Close #1
End Function
'**************** daTaenvironment1.amr ********************************** SACAR DE ACA!!!!


Public Sub lMsg(lineas As Variant) 'porquria, mas facil acordate del VbCrLf
    Dim s As String, l
    For Each l In lineas
        s = s & vbCrLf & l
    Next
    MsgBox s
End Sub
Public Sub che(que)
    MsgBox que, vbExclamation, "Atencion"
End Sub

Public Function confirma(sMsg As String) As Boolean
    confirma = (MsgBox(sMsg, vbQuestion + vbYesNo, "Pregunta") = vbYes)
End Function

Public Function nSinNull(que)
    nSinNull = IIf(IsNull(que), 0, que)
End Function
Public Function sSinNull(que)
    sSinNull = IIf(IsNull(que), "", que)
End Function
Public Sub dtpKUP(kcode As Integer)
    Dim co As Control
    If kcode <> 13 Then Exit Sub
    Set co = Screen.ActiveControl
    If TypeOf co Is DTPicker Then
'        SendKeys "{tab}"
    End If
End Sub

Public Function calculin()
    On Error Resume Next
    calculin = frmCalculin.mostrar()
End Function

Public Sub SubimeSi800x600() 'Optional formulario As Form)
    '*** Poner en        form_Activate() : SubimeSi800x600
    '*** NO SIRVE en     form_load()
    ' usar solo si altura queda justa, porque se vuelve a mover cada vez q activa
   
    On Error Resume Next
'    If formulario Is nothing Then Set formulario = Screen.ActiveForm
    If Screen.Height < 9500 Then Screen.ActiveForm.Top = -30
End Sub

Public Sub relojito(Optional poner As Boolean = True)
    On Error Resume Next
    Screen.ActiveForm.MousePointer = IIf(poner, vbHourglass, vbDefault)
End Sub


Public Sub grillaWidth(grilla, arrayAnchos)
    On Error GoTo fin
    Dim i As Long
    With grilla
        For i = 0 To UBound(arrayAnchos): .ColWidth(i) = arrayAnchos(i): Next
    End With
fin:
End Sub
Public Sub grillaTitulos(grilla, arrayTit)
    On Error GoTo fin
    Dim i As Long
    With grilla
        For i = 0 To UBound(arrayTit): .TextMatrix(0, i) = arrayTit(i): Next
    End With
fin:
End Sub

Public Sub grillaSumarizo(grilla, arraycolumnas)
    'solo flexgrid
    On Error GoTo fin
    Dim i As Long
    With grilla
        .SubtotalPosition = flexSTBelow
        For i = 0 To UBound(arraycolumnas):  .Subtotal flexSTSum, -1, arraycolumnas(i), , , , True, , , True: Next
    End With
fin:
End Sub

Public Sub AcomodarArrayEnFrame(fra As Frame, coso, aTit)
    Dim i As Long, n As Long, hh As Long, h As Long
    
    For i = 0 To coso.UBound
        coso(i).Visible = False
    Next i
    
    n = UBound(aTit) 'coso.UBound
    hh = fra.Width
    h = hh / (1 + n) - 40 - (20 * (1 + n))
    
    For i = 0 To n
        With coso(i)
            .Width = h
            .Height = fra.Height - 40
            .Left = i * (h + 40)
            .Top = 20
            .caption = aTit(i)
            .Visible = True
        End With
    Next i
End Sub
Public Function ligriwi(grilla As Object)
    ' la uso para averiguar anchos en la ventana inmediato, pego resultado a grillawidth
    Dim i As Long, s As String
    
    s = "array("
    For i = 0 To grilla.cols - 1: s = s & grilla.ColWidth(i) & ",": Next
    s = Left(s, Len(s) - 1)
    s = s & ")"
    
    Debug.Print s
End Function


Public Function ssStr(que, Optional paraWhere As Boolean = False) As String
    'PARA GRABAR; no para hacer consultas
    'para instrucciones sql, quita nulos y reemplaza comillas simples " ' " por espacio acentuado " ´ "
    If IsNull(que) Then Exit Function
    
    If Not paraWhere Then
        ssStr = Replace(que, "'", "´")
    Else
        ssStr = Replace(que, "'", "''") ' reemplazo UNA comita "   '    " por DOS comitas "   '   '    "
    End If
End Function

Public Function ParametroEmpresa(cual As String, Optional nuevoValor)
    Dim x As Long
    'x =
    'parametroEmpresa = obtenerdesql("select " & cual & " from DatosEmpresa
End Function

'twips a mm y vice
Public Function tw2mm(t As Long) As Double
    tw2mm = (25.4 * t) / 1440
End Function
Public Function mm2tw(m As Double) As Long
    mm2tw = CLng((m * 1440) / 25.4)
End Function


''public function strExtraer(sTexto as String,
'
'Public Sub sumaSeparada(aCuentas, vCuenta, nMonto)
'    Dim que, i As long, largo As long
'
'    largo = UBound(aCuentas, 1)
'    'If monto = 0 Then Exit Sub ' asumo q lo suma en otra categoria?
'
'    'For Each que In cuentas
'    For i = 0 To 99
'        If cuentas(i, 0) = cuenta Then
'            cuentas(i, 1) = cuentas(i, 1) + monto
'            Exit Sub
'        End If
'    Next i
'
'    'ReDim Preserve cuentas(largo + 1, 1)
'    For i = 0 To 99
'        If cuentas(i, 0) = 0 Then
'            cuentas(i, 0) = cuenta
'            cuentas(i, 1) = monto
'            Exit Sub
'        End If
'    Next i
'    ufa "no pude agregar item", "sumoencuenta",err
'End Sub
' ************** UTILITY ********************

'prueba
'Public Function FilterField(rstTemp As ADODB.Recordset, _
'    strField As String, strFilter As String) As ADODB.Recordset
'    ' Set a filter on the specified Recordset object and then
'    ' open a new Recordset object.
'    rstTemp.Filter = strField & " = '" & strFilter & "'"
'    Set FilterField = rstTemp
'End Function


'
'28/7/4
'   comboarray() + array itemdata
'29/7/4
'   comboCodigo
'11/8/4
'   LIMPIEZA se eliminan func no genericas, se reordena un poco
'   ObtenerDeSQL () puede devolver array varios campos, compatible pa tras
'17/8/4
'   pase grabaBitacora y usuarioActual paracá
'18/8/4
'   dEnvTxt()
'19/8/4
'   lMsg(), confirma()
'27/8/4
'   saque obtenerParametro y lo meti en en modulo de la empresa
'30/8/4
'   ssFecha(), ssBetwenn()  String p' SQL
'31/8/4
'   n2r() formato
'1/9/4
'   s2n() q tome numero tambien cdo empieza con "."      .5  => 0,5
'3/9/4
'   centrarme()
'6/9/4
'   separadorDecimal() : new
'   s2n(),x2s() :
'       antes reempl el "." x ","  antes convertir a numerico siempre
'       ahora reempl solo si conf regional decimal = ","
'16/9/4
'   FrmLoad()  = new
'       no sirve de mucho,
'       centra (mejor CentrarMe me)
'       habiilita .KeyPreview
'       facil podria pintar fondo labl  con fondo = frm
'   FrmKeyPress() = new
'       en evento Form_KeyPress(KeyAscii) para transformar teclas
'       a mayusculas, ENTER a TAB
'22/9/4
'   FrmBorrarTxt()
'       borra todos los textbox de un form
'28/9/4
'   x2s() ahora va con trim
'30/9/4
'   combosql  tratamiento null
'1/10/4
'   nSinNull, sSinNull : new
'   frmBorrarCbo(): new
'26/10/4
'   b2n()
'11-11-4
'   s2n() optional decimales, default 2
'   asFecha() asBetween() new, para access
'17/11-4
'   Fix Style = 1 como parametro al convert MSSQL ssFecha ssBetween
'19/11/4
'   fix s2n decimales 0
'22/11/4
'   new: che() pavada, msgbox con atencion y titulo
'25/11/4
'   fix ssbetween faltaba cambiar otro convert al fix 11/11/4
'26/11/4
'   fix where,sw,mwhere
'30/11/4
'   frmPintoFoco : new
'6/12/4
'   fix ufa() POR FIN!! visibilidad de err.number.
'17/1/5
'   fix s2n() labura sobre temp para no modificar parametro !! debi pasar byval, pero es lo mismo
'21/1/5
'   ufa formato mas legible, y mas datos, menos parametros
'1/2/5
'   fix encajar cuando es padre es form - scaleWidth
'8/3/5
'   fix s2n(): USO ByVal, podria haber arreglado temp, pero es mejor asi
'17/3/5
'   prueba  FilterField()
'29/3/5
'   on err en func
'15/4/5
'   Calculin
'   SubimeSiNoQuepo()
'21/4/5
'   ufa graba en nueva conexion para no entrar en las transacciones
'28/4/5
'   fix calculin control
'26/8/5
'   FormatoFecha()
'27/12/5
'   relojito (OJO sacado de tonka, no se cual es la version ya)

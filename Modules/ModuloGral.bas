Attribute VB_Name = "ModuloGral"
Option Explicit

Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Public Const LOCALE_SDECIMAL = &HE 'separador decimal
Public Const LOCALE_USER_DEFAULT = &H400 'presentar información del usuario
Public Const LOCALE_SSHORTDATE = &H1F 'formato de fecha corta
Public Const UFA_ARCHIVO_LOG = ".\Prg_Bug.log"   ' <-- Donde guarda los errores.
Public Const fileLogFacturacion = ".\fileLogFacturacion.log"
Public Const UFA_STOP = False                    ' <-- para q  ufa() haga un STOP (en diseño)  \'o'/  eh!
'

Public UsuarioSistema As New ADODB.Recordset
Public RsParam As New ADODB.Recordset
Public vieneDE As String


Sub AbrirDB() ' BASURA
    DE_abrir
End Sub


Public Function AccesoSistema(nom As String, cla As String) As Boolean
On Error Resume Next
Dim rs As New ADODB.Recordset

rs.Open "select * from Usuarios where usuario = '" & nom & "' AND clave = '" & cla & "'", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
If rs.EOF Then
    AccesoSistema = False
Else
    AccesoSistema = True
    UsuarioSistema.Open "select * from Usuarios where usuario = '" & nom & "' AND clave = '" & cla & "'", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
End If
rs.Close
Set rs = Nothing

End Function

Function BuscoDato(tabla As String, dato As Long) As String
Dim rs As New ADODB.Recordset

Dim sqlstrCC As String
    
    sqlstrCC = "Select descripcion from " + tabla + " where Codigo = " & dato & " and activo = 1"
    rs.Open sqlstrCC, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        BuscoDato = rs!DESCRIPCION
    End If
    rs.Close
    Set rs = Nothing
    
End Function

Function BuscoTipocomp(dato As String, codigo As Long) As String
Dim rs As New ADODB.Recordset

Dim sqlstrCC As String, abuscar As String, x As Long
    
    sqlstrCC = "Select " & dato & " from Prov where Codigo = " & codigo & " and activo = 1"
    rs.Open sqlstrCC, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        If Not IsNull(rs.Fields(dato)) Then BuscoTipocomp = ObtenerDescripcion("Tipocompras", rs.Fields(dato))
    End If
    rs.Close
    Set rs = Nothing
    
End Function

Function BuscoDatoProv(dato As String, codigo As Long) As String
Dim rs As New ADODB.Recordset

Dim sqlstrCC As String, abuscar As String, x As Long
    
    sqlstrCC = "Select " & dato & " from Prov where Codigo = " & codigo & " and activo = 1"
    rs.Open sqlstrCC, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        If Not IsNull(rs.Fields(dato)) Then BuscoDatoProv = rs.Fields(dato)
    End If
    rs.Close
    Set rs = Nothing
    
End Function


Sub CargarHelpLisProductos(tabla As String, nomCampo1 As String, Campo1 As String, Optional agrupo As String)

Dim rs As New ADODB.Recordset

FrmHelp.grillahelp.Row = 1
FrmHelp.grillahelp.FormatString = "^" + nomCampo1 + "        |<"


    If Trim(agrupo) <> "" Then
        rs.Open "select distinct " & Campo1 & " from " + tabla & " where activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
    Else
        rs.Open "select * from " + tabla & " where activo = 1 order by " & Campo1, DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
    End If

    If Not rs.EOF Then
        rs.MoveFirst
        'Agrego la primera linea sin hacer additem
        FrmHelp.grillahelp.Col = 0
        FrmHelp.grillahelp.Text = Trim(rs(Campo1))
        rs.MoveNext
        Do While Not rs.EOF
            FrmHelp.grillahelp.AddItem Trim(rs(Campo1))
            rs.MoveNext
        Loop
    End If
End Sub

'Sub CargarHelpCuentas(Tabla As String, nomCampo1 As String, nomCampo2 As String, Campo1 As String, Campo2 As String, Optional order As String)
'
'Dim rs As New ADODB.Recordset, ingreso As Integer
'
'FrmHelp.grillahelp.row = 1
'FrmHelp.grillahelp.FormatString = "^" + nomCampo1 + "        |<" + nomCampo2 + "                                                    "
'
'
'    If Trim(order) <> "" Then
'        rs.Open "select * from " + Tabla & " where imputable = 1 and activo = 1 order by " & order, DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
'    Else
'        rs.Open "select * from " + Tabla & " where imputable = 1 and activo = 1 order by Codigo", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
'    End If
'
'
'    If Not rs.EOF Then
'        rs.MoveFirst
'        'Agrego la primera linea sin hacer additem
'        ingreso = 0
'        If esimputable(rs.Fields(Campo1)) Then
'            FrmHelp.grillahelp.col = 0
'            FrmHelp.grillahelp.Text = Trim(rs(Campo1))
'            FrmHelp.grillahelp.col = 1
'            FrmHelp.grillahelp.Text = Trim(rs(Campo2))
'            ingreso = 1
'        End If
'        rs.MoveNext
'        Do While Not rs.EOF
'            If esimputable(rs.Fields(Campo1)) Then
'                If ingreso = 1 Then
'                    FrmHelp.grillahelp.AddItem Trim(rs(Campo1)) & Chr(9) & Trim(rs(Campo2))
'                Else
'                    FrmHelp.grillahelp.col = 0
'                    FrmHelp.grillahelp.Text = Trim(rs(Campo1))
'                    FrmHelp.grillahelp.col = 1
'                    FrmHelp.grillahelp.Text = Trim(rs(Campo2))
'                    ingreso = 1
'                End If
'            End If
'            rs.MoveNext
'        Loop
'    End If
'    Set rs = Nothing
'End Sub

Sub CargarHelpFact(tabla As String, nomCampo1 As String, nomCampo2 As String, Campo1 As String, Campo2 As String, Optional order As String)

Dim rs As New ADODB.Recordset

FrmHelp.grillahelp.Row = 1
FrmHelp.grillahelp.FormatString = "^" + nomCampo1 + "        |<" + nomCampo2 + "                                                    "


    If Trim(order) <> "" Then
        rs.Open "select * from " + tabla & " where activo = 1 order by " & order, DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
    Else
        rs.Open "select * from " + tabla & " where activo = 1 order by Fecha", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
    End If


    If Not rs.EOF Then
        rs.MoveFirst
        'Agrego la primera linea sin hacer additem
        FrmHelp.grillahelp.Col = 0
        FrmHelp.grillahelp.Text = Trim(rs(Campo1))
        FrmHelp.grillahelp.Col = 1
        FrmHelp.grillahelp.Text = Trim(rs(Campo2))
        rs.MoveNext
        Do While Not rs.EOF
            FrmHelp.grillahelp.AddItem Trim(rs(Campo1)) & Chr(9) & Trim(rs(Campo2))
            rs.MoveNext
        Loop
    End If
End Sub


Sub CargarHelp(tabla As String, nomCampo1 As String, nomCampo2 As String, Campo1 As String, Campo2 As String, Optional order As String, Optional sWhere As String = "")

Dim rs As New ADODB.Recordset

FrmHelp.grillahelp.Row = 1
FrmHelp.grillahelp.FormatString = "^" + nomCampo1 + "        |<" + nomCampo2 + "                                                    "

    If sWhere > "" Then sWhere = " and " & sWhere & " "
    If Trim(order) <> "" Then
        rs.Open "select * from " + tabla & " where activo = 1 " & sWhere & " order by " & order, DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
    Else
        rs.Open "select * from " + tabla & " where activo = 1 " & sWhere & "  order by Codigo", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
    End If


    If Not rs.EOF Then
        rs.MoveFirst
        'Agrego la primera linea sin hacer additem
        FrmHelp.grillahelp.Col = 0
        FrmHelp.grillahelp.Text = sSinNull(rs(Campo1))
        FrmHelp.grillahelp.Col = 1
        FrmHelp.grillahelp.Text = sSinNull(rs(Campo2))
        rs.MoveNext
        Do While Not rs.EOF
            FrmHelp.grillahelp.AddItem Trim(rs(Campo1)) & Chr(9) & Trim(rs(Campo2))
            rs.MoveNext
        Loop
    End If
End Sub

Sub CargarHelpGtosBanc(tabla As String, nomCampo1 As String, nomCampo2 As String, Campo1 As String, Campo2 As String, Campo3 As String, Optional order As String)

    Dim rs As New ADODB.Recordset
    
    FrmHelp.grillahelp.Row = 1
    FrmHelp.grillahelp.FormatString = "^" + nomCampo1 + "         |<" + nomCampo2 + "             " ' |<" + nomCampo3 + """" ***************** LLLLLLLLLLLLLLLL 5/10/4
    

    If Trim(order) <> "" Then
        rs.Open "select * from " + tabla & " where activo = 1 order by " & order, DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
    Else
        rs.Open "select * from " + tabla & " where activo = 1 order by Codigo", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
    End If
    
    If Not rs.EOF Then
        rs.MoveFirst
        'Agrego la primera linea sin hacer additem
        FrmHelp.grillahelp.ColWidth(2) = 0
        FrmHelp.grillahelp.Col = 0
        FrmHelp.grillahelp.Text = Trim(rs(Campo1))
        FrmHelp.grillahelp.Col = 1
        FrmHelp.grillahelp.Text = Trim(rs(Campo2))
        FrmHelp.grillahelp.Col = 2
        FrmHelp.grillahelp.Text = Trim(rs(Campo3))
        rs.MoveNext
        Do While Not rs.EOF
            FrmHelp.grillahelp.AddItem Trim(rs(Campo1)) & Chr(9) & Trim(rs(Campo2)) & Chr(9) & Trim(rs(Campo3))
            rs.MoveNext
        Loop
    End If
End Sub


Sub CargarHelpCompImput(tabla As String, nomCampo1 As String, nomCampo2 As String, nomCampo3 As String, nomCampo4 As String, Campo1 As String, Campo2 As String, Campo3 As String, Campo4 As String, Campo5 As String, Optional order As String, Optional Tipo As String)

Dim rs As New ADODB.Recordset, ingreso As Long

FrmHelp.grillahelp.Row = 1
FrmHelp.grillahelp.FormatString = "^" + nomCampo1 + "         |<" + nomCampo2 + "        |<" + nomCampo3 + "        |<" + nomCampo4 + "           "

    If Trim(order) <> "" Then
        If Campo5 <> "" Then
            rs.Open "select * from " + tabla & " where activo = 1 and tipodoc = '" & Tipo & "' and codpr = " & val(Campo5) & " order by " & order, DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        Else
            rs.Open "select * from " + tabla & " where activo = 1 and tipodoc = '" & Tipo & "' order by " & order, DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        End If
    Else
        If Campo5 <> "" Then
            rs.Open "select * from " + tabla & " where activo = 1 and tipodoc = '" & Tipo & "' and codpr = " & val(Campo5) & " order by Codigo", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        Else
            rs.Open "select * from " + tabla & " where activo = 1 and tipodoc = '" & Tipo & "' order by Codigo", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        End If
    End If

    If Not rs.EOF Then
        rs.MoveFirst
        'Agrego la primera linea sin hacer additem
        ingreso = 0
        If esimputable(rs.Fields(Campo1)) Then
            FrmHelp.grillahelp.Col = 0
            FrmHelp.grillahelp.Text = Trim(rs(Campo1))
            FrmHelp.grillahelp.Col = 1
            FrmHelp.grillahelp.Text = Trim(rs(Campo2))
            FrmHelp.grillahelp.Col = 2
            FrmHelp.grillahelp.Text = Trim(rs(Campo3))
            FrmHelp.grillahelp.Col = 3
            FrmHelp.grillahelp.Text = Trim(rs(Campo4))
            ingreso = 1
        End If
        rs.MoveNext
        Do While Not rs.EOF
            If esimputable(rs.Fields(Campo1)) Then
                If ingreso = 1 Then
                    FrmHelp.grillahelp.AddItem Trim(rs(Campo1)) & Chr(9) & Trim(rs(Campo2)) & Chr(9) & Trim(rs(Campo3)) & Chr(9) & Trim(rs(Campo4))
                Else
                    FrmHelp.grillahelp.Col = 0
                    FrmHelp.grillahelp.Text = Trim(rs(Campo1))
                    FrmHelp.grillahelp.Col = 1
                    FrmHelp.grillahelp.Text = Trim(rs(Campo2))
                    FrmHelp.grillahelp.Col = 2
                    FrmHelp.grillahelp.Text = Trim(rs(Campo3))
                    FrmHelp.grillahelp.Col = 3
                    FrmHelp.grillahelp.Text = Trim(rs(Campo4))
                    ingreso = 1
                End If
            End If
            rs.MoveNext
        Loop
    End If
End Sub


Sub CargarHelpComp(tabla As String, nomCampo1 As String, nomCampo2 As String, nomCampo3 As String, nomCampo4 As String, Campo1 As String, Campo2 As String, Campo3 As String, Campo4 As String, Campo5 As String, Optional order As String, Optional Tipo As String)

Dim rs As New ADODB.Recordset

FrmHelp.grillahelp.Row = 1
FrmHelp.grillahelp.FormatString = "^" + nomCampo1 + "         |<" + nomCampo2 + "        |<" + nomCampo3 + "        |<" + nomCampo4 + "           "

    If Trim(order) <> "" Then
        If Campo5 <> "" Then
            rs.Open "select * from " + tabla & " where activo = 1 and tipodoc = '" & Tipo & "' and codpr = " & val(Campo5) & " order by " & order, DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        Else
            rs.Open "select * from " + tabla & " where activo = 1 and tipodoc = '" & Tipo & "' order by " & order, DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        End If
    Else
        If Campo5 <> "" Then
            rs.Open "select * from " + tabla & " where activo = 1 and tipodoc = '" & Tipo & "' and codpr = " & val(Campo5) & " order by Codigo", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        Else
            rs.Open "select * from " + tabla & " where activo = 1 and tipodoc = '" & Tipo & "' order by Codigo", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        End If
    End If

    If Not rs.EOF Then
        rs.MoveFirst
        'Agrego la primera linea sin hacer additem
        FrmHelp.grillahelp.Col = 0
        FrmHelp.grillahelp.Text = Trim(rs(Campo1))
        FrmHelp.grillahelp.Col = 1
        FrmHelp.grillahelp.Text = Trim(rs(Campo2))
        FrmHelp.grillahelp.Col = 2
        FrmHelp.grillahelp.Text = Trim(rs(Campo3))
        FrmHelp.grillahelp.Col = 3
        FrmHelp.grillahelp.Text = Trim(rs(Campo4))
        rs.MoveNext
        Do While Not rs.EOF
            FrmHelp.grillahelp.AddItem Trim(rs(Campo1)) & Chr(9) & Trim(rs(Campo2)) & Chr(9) & Trim(rs(Campo3)) & Chr(9) & Trim(rs(Campo4))
            rs.MoveNext
        Loop
    End If
End Sub


Sub CargarHelpOPago(tabla As String, nomCampo1 As String, nomCampo2 As String, nomCampo3 As String, nomCampo4 As String, Campo1 As String, Campo2 As String, Campo3 As String, Campo4 As String, Campo5 As String, Optional order As String)

Dim rs As New ADODB.Recordset

FrmHelp.grillahelp.Row = 1
FrmHelp.grillahelp.FormatString = "^" + nomCampo1 + "         |<" + nomCampo2 + "        |<" + nomCampo3 + "        |<" + nomCampo4 + "         "


'    If Trim(order) <> "" Then
        rs.Open "select * from " + tabla & " where activo = 1 and codpr = " & val(Campo3) & " order by " & Campo1, DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
'    Else
'        rs.Open "select * from " + Tabla & " where activo = 1 and codpr = " & Val(Campo3) & " order by Codigo", daTaenvironment1.Sistema, adOpenStatic, adLockOptimistic
'    End If
    
    If Not rs.EOF Then
        rs.MoveFirst
        'Agrego la primera linea sin hacer additem
        FrmHelp.grillahelp.Col = 0
        FrmHelp.grillahelp.Text = Trim(rs(Campo1))
        FrmHelp.grillahelp.Col = 1
        FrmHelp.grillahelp.Text = Trim(rs(Campo2))
        FrmHelp.grillahelp.Col = 2
        FrmHelp.grillahelp.Text = Trim(rs(Campo3))
        FrmHelp.grillahelp.Col = 3
        FrmHelp.grillahelp.Text = Trim(rs(Campo4))
        rs.MoveNext
        Do While Not rs.EOF
            FrmHelp.grillahelp.AddItem Trim(rs(Campo1)) & Chr(9) & Trim(rs(Campo2)) & Chr(9) & Trim(rs(Campo3)) & Chr(9) & Trim(rs(Campo4))
            rs.MoveNext
        Loop
    End If
End Sub



Sub CargarHelpChequesTerceros(tabla As String, nomCampo1 As String, nomCampo2 As String, nomCampo3 As String, Campo1 As String, Campo2 As String, Campo3 As String, Campo4 As String, Optional order As String)

Dim rs As New ADODB.Recordset
Dim descmoneda As String

FrmHelp.grillahelp.Row = 1
FrmHelp.grillahelp.FormatString = "^" + nomCampo1 + "         |<" + nomCampo2 + "        |<" + nomCampo3 + "                                                     """


    If Trim(order) <> "" Then
        'rs.Open "select * from " + tabla & " where activo = 1 and estado = 'C' order by " & order, DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        rs.Open "select * from " + tabla & " where activo = 1 order by " & order, DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
    Else
        'rs.Open "select * from " + tabla & " where activo = 1 and estado = 'C' order by Codigo", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        rs.Open "select * from " + tabla & " where activo = 1  order by Codigo", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
    End If


    If Not rs.EOF Then
        rs.MoveFirst
        'Agrego la primera linea sin hacer additem
        FrmHelp.grillahelp.Col = 0
        FrmHelp.grillahelp.Text = Trim(rs(Campo1))
        FrmHelp.grillahelp.Col = 1
        FrmHelp.grillahelp.Text = Trim(rs(Campo2))
        FrmHelp.grillahelp.Col = 2
        FrmHelp.grillahelp.Text = ObtenerDescripcion("BancosGrales", rs(Campo3)) & " - " & Trim(rs(Campo4))
        rs.MoveNext
        Do While Not rs.EOF
            FrmHelp.grillahelp.AddItem Trim(rs(Campo1)) & Chr(9) & Trim(rs(Campo2)) & Chr(9) & ObtenerDescripcion("BancosGrales", rs(Campo3)) & " - " & Trim(rs(Campo4))
            rs.MoveNext
        Loop
    End If
End Sub

Sub CargarHelpCheques(tabla As String, nomCampo1 As String, nomCampo2 As String, Campo1 As String, Campo2 As String, Campo3 As String, Optional order As String)

Dim rs As New ADODB.Recordset
Dim descmoneda As String

FrmHelp.grillahelp.Row = 1
FrmHelp.grillahelp.FormatString = "^" + nomCampo1 + "         |<" + nomCampo2 + "                                                    "


    If Trim(order) <> "" Then
        rs.Open "select * from " + tabla & " where activo = 1 and estado = 'C' order by " & order, DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
    Else
        rs.Open "select * from " + tabla & " where activo = 1 and estado = 'C' order by Codigo", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
    End If


    If Not rs.EOF Then
        rs.MoveFirst
        'Agrego la primera linea sin hacer additem
        FrmHelp.grillahelp.Col = 0
        FrmHelp.grillahelp.Text = Trim(rs(Campo1))
        FrmHelp.grillahelp.Col = 1
        FrmHelp.grillahelp.Text = ObtenerDescripcion("BancosGrales", rs(Campo2)) & " - " & Trim(rs(Campo3))
        rs.MoveNext
        Do While Not rs.EOF
            FrmHelp.grillahelp.AddItem Trim(rs(Campo1)) & Chr(9) & ObtenerDescripcion("BancosGrales", rs(Campo2)) & " - " & Trim(rs(Campo3))
            rs.MoveNext
        Loop
    End If
End Sub

Sub CargarHelpMovibanc(tabla As String, nomCampo1 As String, nomCampo2 As String, Campo1 As String, Campo2 As String, Campo3 As String, Campo4 As String, Optional order As String)

Dim rs As New ADODB.Recordset
Dim descmoneda As String

FrmHelp.grillahelp.Row = 1
FrmHelp.grillahelp.FormatString = "^" + nomCampo1 + "         |<" + nomCampo2 + "                                                                                  "


    If Trim(order) <> "" Then
        rs.Open "select * from " + tabla & " where activo = 1 and operacion = 'S' order by " & order, DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
    Else
        rs.Open "select * from " + tabla & " where activo = 1 and operacion = 'S' order by Codigo", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
    End If


    If Not rs.EOF Then
        rs.MoveFirst
        'Agrego la primera linea sin hacer additem
        FrmHelp.grillahelp.Col = 0
        FrmHelp.grillahelp.Text = Trim(rs(Campo1))
        FrmHelp.grillahelp.Col = 1
        FrmHelp.grillahelp.Text = Trim(rs(Campo2)) & " - " & Trim(rs(Campo3))
        rs.MoveNext
        Do While Not rs.EOF
            FrmHelp.grillahelp.AddItem Trim(rs(Campo1)) & Chr(9) & Trim(rs(Campo2)) & " - " & Trim(rs(Campo3))
            rs.MoveNext
        Loop
    End If
End Sub

Sub CargarHelpCtasBanc(tabla As String, nomCampo1 As String, nomCampo2 As String, Campo1 As String, Campo2 As String, Campo3 As String, Campo4 As String, Optional order As String)

Dim rs As New ADODB.Recordset
Dim descmoneda As String

FrmHelp.grillahelp.Row = 1
FrmHelp.grillahelp.FormatString = "^" + nomCampo1 + "         |<" + nomCampo2 + "                                                                                  "


    If Trim(order) <> "" Then
        rs.Open "select * from " + tabla & " where activo = 1 order by " & order, DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
    Else
        rs.Open "select * from " + tabla & " where activo = 1 order by Codigo", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
    End If


    If Not rs.EOF Then
        rs.MoveFirst
        'Agrego la primera linea sin hacer additem
        FrmHelp.grillahelp.Col = 0
        FrmHelp.grillahelp.Text = Trim(rs(Campo1))
        FrmHelp.grillahelp.Col = 1
        FrmHelp.grillahelp.Text = ObtenerDescripcion("BancosGrales", rs(Campo2)) & " - " & Trim(rs(Campo3)) & " - " & ObtenerDescripcion("Monedas", rs(Campo4))
        rs.MoveNext
        Do While Not rs.EOF
            FrmHelp.grillahelp.AddItem Trim(rs(Campo1)) & Chr(9) & ObtenerDescripcion("BancosGrales", rs(Campo2)) & " - " & Trim(rs(Campo3)) & " - " & ObtenerDescripcion("Monedas", rs(Campo4))
            rs.MoveNext
        Loop
    End If
End Sub

Sub CargarHelpCompras(tabla As String, nomCampo1 As String, nomCampo2 As String, Campo1 As String, Campo2 As String, Optional order As String)

Dim rs As New ADODB.Recordset
Dim descmoneda As String

FrmHelp.grillahelp.Row = 1
FrmHelp.grillahelp.FormatString = "^" + nomCampo1 + "         |<" + nomCampo2 + "                                                                                  "


    If Trim(order) <> "" Then
        rs.Open "select * from " + tabla & " where activo = 1 order by " & order, DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
    Else
        rs.Open "select * from " + tabla & " where activo = 1 order by Codigo", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
    End If
    
    If Not rs.EOF Then
        rs.MoveFirst
        'Agrego la primera linea sin hacer additem
        FrmHelp.grillahelp.Col = 0
        FrmHelp.grillahelp.Text = Trim(rs(Campo1))
        FrmHelp.grillahelp.Col = 1
        FrmHelp.grillahelp.Text = Trim(rs(Campo2)) & " - " & ObtenerDescripcion("Prov", rs(Campo2))
        rs.MoveNext
        Do While Not rs.EOF
            FrmHelp.grillahelp.AddItem Trim(rs(Campo1)) & Chr(9) & Trim(rs(Campo2)) & " - " & ObtenerDescripcion("Prov", rs(Campo2))
            rs.MoveNext
        Loop
    End If
End Sub


Sub CargarHelpLibracheques(tabla As String, nomCampo1 As String, nomCampo2 As String, Campo1 As String, Campo2 As String, Campo3 As String, Optional filtro As String, Optional Orden As String)

Dim rs As New ADODB.Recordset

FrmHelp.grillahelp.Row = 1
FrmHelp.grillahelp.FormatString = "^" + nomCampo1 + "        |<" + nomCampo2 + "                                                    "


    If Trim(filtro) <> "" Then
        If Trim(Orden) <> "" Then
            rs.Open "select * from " + tabla & " where activo = 1 and " & filtro & " order by " & Orden, DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        Else
            rs.Open "select * from " + tabla & " where activo = 1 and " & filtro, DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        End If
    Else
        If Trim(Orden) <> "" Then
            rs.Open "select * from " + tabla & " where activo = 1 order by " & Orden, DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        Else
            rs.Open "select * from " + tabla & " where activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        End If
    End If


    If Not rs.EOF Then
        rs.MoveFirst
        'Agrego la primera linea sin hacer additem
        FrmHelp.grillahelp.Col = 0
        FrmHelp.grillahelp.Text = Trim(rs(Campo1))
        FrmHelp.grillahelp.Col = 1
        FrmHelp.grillahelp.Text = rs(Campo2) & " - " & str(Trim(rs(Campo3)))
        rs.MoveNext
        Do While Not rs.EOF
            FrmHelp.grillahelp.AddItem Trim(rs(Campo1)) & Chr(9) & Trim(rs(Campo2)) & " - " & Trim(rs(Campo3))
            rs.MoveNext
        Loop
    End If
End Sub

Public Function ObtenerIvaProv(tabla As String, COD As Long) As Long

Dim rs As New ADODB.Recordset

    rs.Open "select tipoiva from " & tabla & " where codigo = " & Trim(COD), DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly

    If Not rs.EOF Then
        ObtenerIvaProv = rs!tipoiva
    Else
        ObtenerIvaProv = 0
    End If

    rs.Close
    Set rs = Nothing

End Function
Public Function ObtenerDescripcion(tabla As String, COD As Long) As String

Dim rs As New ADODB.Recordset


rs.Open "select * from " & tabla & " where CODIGO = '" & Trim(COD) & "' ", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly

If Not rs.EOF Then
    ObtenerDescripcion = rs!DESCRIPCION
Else
    ObtenerDescripcion = ""
End If

rs.Close
Set rs = Nothing

End Function

Public Function ObtenerMoneda(tabla As String, COD As Long) As Long

Dim rs As New ADODB.Recordset


rs.Open "select * from " & tabla & " where CODIGO = " & Trim(COD), DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly

If Not rs.EOF Then
    If Not IsNull(rs!moneda) Then
        ObtenerMoneda = rs!moneda
    Else
        ObtenerMoneda = 0
    End If
Else
    ObtenerMoneda = 0
End If

rs.Close
Set rs = Nothing

End Function


Public Function ObtenerCotizacion(tabla As String, COD As Long) As Double

Dim rs As New ADODB.Recordset


rs.Open "select * from " & tabla & " where moneda = " & Trim(COD) & "", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly

If Not rs.EOF Then
    ObtenerCotizacion = rs!cotizacion
Else
    ObtenerCotizacion = ""
End If

rs.Close
Set rs = Nothing

End Function

Public Function ObtenerDescripcionCuentas(tabla As String, COD As Long) As String

Dim rs As New ADODB.Recordset
Dim NroCue As String
Dim DescCue As String


rs.Open "select * from " & tabla & " where codigo = " & COD, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly

If Not rs.EOF Then
    NroCue = Trim(rs!numero)
    DescCue = obtenerDeSQL("select descripcion from bancosgrales where codigo=" & rs!Banco)
    ObtenerDescripcionCuentas = DescCue & " - " & NroCue
Else
    ObtenerDescripcionCuentas = ""
End If

rs.Close
Set rs = Nothing

End Function

Public Function ObtenerImporteCuentas(tabla As String, COD As Long) As Double

Dim rs As New ADODB.Recordset


rs.Open "select * from " & tabla & " where Cuenta = '" & Trim(COD) & "'", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly

If Not rs.EOF Then
    ObtenerImporteCuentas = rs!Importe
Else
    ObtenerImporteCuentas = 0
End If

rs.Close
Set rs = Nothing

End Function

Public Function ObtenerFechaCuentas(tabla As String, COD As Long) As String
Dim rs As New ADODB.Recordset


rs.Open "select * from " & tabla & " where Cuenta = '" & Trim(COD) & "'", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly

If Not rs.EOF Then
    ObtenerFechaCuentas = rs!Fecha
Else
    ObtenerFechaCuentas = ""
End If

rs.Close
Set rs = Nothing

End Function


Public Function ObtenerDescripcionCajas(tabla As String, COD As Long) As String

Dim rs As New ADODB.Recordset


rs.Open "select * from " & tabla & " where CODIGO = '" & Trim(COD) & "'", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly

If Not rs.EOF Then
    ObtenerDescripcionCajas = rs!responsable
Else
    ObtenerDescripcionCajas = ""
End If

rs.Close
Set rs = Nothing

End Function
Public Function ObtenerDescripcionBancos(tabla As String, COD As Long) As String

Dim rs As New ADODB.Recordset


rs.Open "select * from " & tabla & " where CODIGO = " & Trim(COD) & "", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly

If Not rs.EOF Then
    ObtenerDescripcionBancos = rs!DESCRIPCION
Else
    ObtenerDescripcionBancos = ""
End If

rs.Close
Set rs = Nothing

End Function


Public Function ObtenerCodigo(tabla As String, DESCRIPCION As String) As Long

Dim rs As New ADODB.Recordset

rs.Open "select * from " & tabla & " where activo=1 and DESCRIPCION = '" & DESCRIPCION & "'", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly

If Not rs.EOF Then
    ObtenerCodigo = rs!codigo
Else
    ObtenerCodigo = 0
End If

rs.Close
Set rs = Nothing

End Function

Public Function ObtenerCodigo2(tabla As String, DESCRIPCION As String) As Long

Dim rs As New ADODB.Recordset

rs.Open "select * from " & tabla & " where DESCRIPCIONusuario = '" & DESCRIPCION & "'", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly

If Not rs.EOF Then
    ObtenerCodigo2 = rs!codigo
Else
    ObtenerCodigo2 = 0
End If

rs.Close
Set rs = Nothing

End Function

Public Function NroEnLetras(ByVal curNumero As Double, Optional blnO_Final As Boolean = True) As String
'Devuelve un número expresado en letras.
'El parámetro blnO_Final se utiliza en la recursión para saber si se debe colocar
'la "O" final cuando la palabra es UN(O)
    Dim dblCentavos As Double
    Dim lngContDec As Long
    Dim lngContCent As Long
    Dim lngContMil As Long
    Dim lngContMillon As Long
    Dim strNumLetras As String
    Dim strNumero As Variant
    Dim strDecenas As Variant
    Dim strCentenas As Variant
    Dim blnNegativo As Boolean
    Dim blnPlural As Boolean

    If Int(curNumero) = 0# Then
        strNumLetras = "CERO"
    End If

    strNumero = Array(vbNullString, "UN", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", _
                   "OCHO", "NUEVE", "DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", _
                   "QUINCE", "DIECISEIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE", _
                   "VEINTE")

    strDecenas = Array(vbNullString, vbNullString, "VEINTI", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", _
                    "SETENTA", "OCHENTA", "NOVENTA", "CIEN")

    strCentenas = Array(vbNullString, "CIENTO", "DOSCIENTOS", "TRESCIENTOS", _
                     "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", _
                     "OCHOCIENTOS", "NOVECIENTOS")

    If curNumero < 0# Then
        blnNegativo = True
        curNumero = Abs(curNumero)
    End If

    If Int(curNumero) <> curNumero Then
        dblCentavos = Abs(curNumero - Int(curNumero))
        curNumero = Int(curNumero)
    End If

    Do While curNumero >= 1000000#
        lngContMillon = lngContMillon + 1
        curNumero = curNumero - 1000000#
    Loop

    Do While curNumero >= 1000#
        lngContMil = lngContMil + 1
        curNumero = curNumero - 1000#
    Loop

    Do While curNumero >= 100#
        lngContCent = lngContCent + 1
        curNumero = curNumero - 100#
    Loop

    If Not (curNumero > 10# And curNumero <= 20#) Then
        Do While curNumero >= 10#
            lngContDec = lngContDec + 1
            curNumero = curNumero - 10#
        Loop
    End If

    If lngContMillon > 0 Then
        If lngContMillon >= 1 Then   'si el número es >1000000 usa recursividad
            strNumLetras = NroEnLetras(lngContMillon, False)
            If Not blnPlural Then blnPlural = (lngContMillon > 1)
            lngContMillon = 0
        End If
        strNumLetras = Trim(strNumLetras) & strNumero(lngContMillon) & " MILLON" & _
                                                                    IIf(blnPlural, "ES ", " ")
    End If

    If lngContMil > 0 Then
        If lngContMil >= 1 Then   'si el número es >100000 usa recursividad
            strNumLetras = strNumLetras & NroEnLetras(lngContMil, False)
            lngContMil = 0
        End If
        strNumLetras = Trim(strNumLetras) & strNumero(lngContMil) & " MIL "
    End If

    If lngContCent > 0 Then
        If lngContCent = 1 And lngContDec = 0 And curNumero = 0# Then
            strNumLetras = strNumLetras & "CIEN"
        Else
            strNumLetras = strNumLetras & strCentenas(lngContCent) & " "
        End If
    End If

    If lngContDec >= 1 Then
        If lngContDec = 1 Then
            strNumLetras = strNumLetras & strNumero(10)
        Else
            strNumLetras = strNumLetras & strDecenas(lngContDec)
        End If

        If lngContDec >= 3 And curNumero > 0# Then
            strNumLetras = strNumLetras & " Y "
        End If
    Else
        If curNumero >= 0# And curNumero <= 20# Then
            strNumLetras = strNumLetras & strNumero(curNumero)
            If curNumero = 1# And blnO_Final Then
                strNumLetras = strNumLetras & "O"
            End If
            If dblCentavos > 0# Then
                strNumLetras = Trim(strNumLetras) & " CON " & Format$(CInt(dblCentavos * 100#), "00") & "/100"
            End If
            NroEnLetras = strNumLetras
            Exit Function
        End If
    End If

    If curNumero > 0# Then
        strNumLetras = strNumLetras & strNumero(curNumero)
        If curNumero = 1# And blnO_Final Then
            strNumLetras = strNumLetras & "O"
        End If
    End If

    If dblCentavos > 0# Then
        strNumLetras = strNumLetras & " CON " + Format$(CInt(dblCentavos * 100#), "00") & "/100"
    End If

    NroEnLetras = IIf(blnNegativo, "(" & strNumLetras & ")", strNumLetras)
End Function


Function noestaenlagrilla(codigo As String, GRILLA As MSHFlexGrid) As Boolean
Dim x As Long
    noestaenlagrilla = False
    For x = 1 To GRILLA.rows - 1
        If GRILLA.TextMatrix(x, 0) = codigo And codigo <> "" Then
            noestaenlagrilla = True
        End If
    Next

End Function
Function noestaenlagrillaVS(codigo As String, GRILLA As VSFlexGrid) As Boolean
    Dim x As Long
    noestaenlagrillaVS = False
    For x = 1 To GRILLA.rows - 1
        If GRILLA.TextMatrix(x, 0) = codigo And codigo <> "" Then
            noestaenlagrillaVS = True
        End If
    Next

End Function

Function esimputable(codigo As String) As Boolean
Dim rs As New ADODB.Recordset

    If Len(codigo) <= 9 Then
        'rs.Open "select imputable from Cuentas where _codigo = " & Val(codigo) & " and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        rs.Open "select imputable from Cuentas where cuenta = '" & val(codigo) & "' and activo = 1", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        If Not rs.EOF Then
            If rs!IMPUTABLE = True Then
                esimputable = True
            Else
                esimputable = False
            End If
        End If
        rs.Close
        Set rs = Nothing
    Else
        esimputable = False
    End If

End Function

Public Function ObtenerCodigoS(tabla As String, DESCRIPCION As String) As String

Dim rs As New ADODB.Recordset


rs.Open "select * from " & tabla & " where DESCRIPCION = '" & DESCRIPCION & "'", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly

If Not rs.EOF Then
    ObtenerCodigoS = rs!codigo
Else
    ObtenerCodigoS = ""
End If

rs.Close
Set rs = Nothing

End Function

Public Function ObtenerCodigoCue(tabla As String, DESCRIPCION As String) As String

Dim rs As New ADODB.Recordset


rs.Open "select * from " & tabla & " where DESCRIPCION = '" & DESCRIPCION & "'", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly

If Not rs.EOF Then
    ObtenerCodigoCue = rs!Cuenta
Else
    ObtenerCodigoCue = ""
End If

rs.Close
Set rs = Nothing

End Function

Public Function ObtenerDescripcionS(tabla As String, COD As String) As String

Dim rs As New ADODB.Recordset


rs.Open "select * from " & tabla & " where CODIGO = '" & Trim(COD) & "'", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly

If Not rs.EOF Then
    ObtenerDescripcionS = sSinNull(rs!DESCRIPCION)
Else
    ObtenerDescripcionS = ""
End If

rs.Close
Set rs = Nothing

End Function

Public Function MovimientoConcepto(COD As Long) As String 'fix 17/8/4 movimiento = 'N'
    Dim rs As New ADODB.Recordset

rs.Open "select * from conceptos where codigo = " & COD, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly

If Not rs.EOF Then
    MovimientoConcepto = rs!movimiento
Else
    MovimientoConcepto = "N"
End If

rs.Close
Set rs = Nothing

End Function

Public Function ComprobanteConcepto(COD As Long) As Long
Dim rs As New ADODB.Recordset

rs.Open "select * from conceptos where codigo = " & COD, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
If Not rs.EOF Then
    ComprobanteConcepto = nSinNull(rs!COMPROBANTE)
Else
    ComprobanteConcepto = 0
End If

rs.Close
Set rs = Nothing
End Function

''' ------ No son func grales, se usasn en los 3 frm de cheques------------
''Sub CargoGrillaterceros(grilla As MSHFlexGrid)
''    On Error GoTo ufaErr
''    Dim x As Integer
''    Dim rs As New ADODB.Recordset
'''    rs.Open "Select * from Cheques where estado = 'C' and activo = 1", daTaenvironment1.Sistema, adOpenDynamic, adLockOptimistic
''    rs.Open "Select NroInt, BancosGrales.Descripcion as Banco, Nro, Importe, Fecha from Cheques inner join BancosGrales on Banco_Nro = BancosGrales.codigo where estado = 'C' and cheques.activo = 1", daTaenvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
''    With grilla
''        .Redraw = False
''        .row = 1
''        .Col = 0
''        While Not rs.EOF
''            If Trim(.Text) = "" Then
''                .row = 1
''                .Col = 0
''                .Text = rs!nroInt
''                .Col = 1
''                .Text = rs!Banco  'ObtenerDescripcion("BancosGrales", rs!banco_nro)
''                .Col = 2
''                .Text = rs!Nro
''                .Col = 3
''                .Text = rs!importe
''                .Col = 4
''                .Text = rs!fecha
''            Else
''                .AddItem rs!nroInt & Chr(9) & rs!Banco & Chr(9) & rs!Nro & Chr(9) & rs!importe & Chr(9) & rs!fecha
''            End If
''            rs.MoveNext
''        Wend
''    End With
''FIN:
''    grilla.Redraw = True
''    rs.Close
''    Set rs = Nothing
''    Exit Sub
''ufaErr:
''    Resume FIN
''End Sub
''Sub CargoGrillaPropios(grilla As MSHFlexGrid)
''    On Error GoTo ufaErr
''    Dim x As Integer
''    Dim rs As New ADODB.Recordset
''    'rs.Open "Select * from Chq_comp where estado = 'C' and activo = 1", daTaenvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
''    rs.Open "Select chq_comp.Codigo, Nro, Importe, fecha_cheque,  BancosGrales.Descripcion as Banco from Chq_comp inner join bancosGrales on BancosGrales.codigo = banco  where estado = 'C' and Chq_comp.activo = 1", daTaenvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
''    With grilla
''        .Redraw = False
''        .row = 1
''        .Col = 0
''        While Not rs.EOF
''            If Trim(.Text) = "" Then
''                .row = 1
''                .Col = 0
''                .Text = rs!Codigo
''                .Col = 1
''                .Text = rs!Banco 'ObtenerDescripcion("BancosGrales", rs!Banco)
''                .Col = 2
''                .Text = rs!Nro
''                .Col = 3
''                .Text = rs!importe
''                .Col = 4
''                .Text = rs!fecha_cheque
''            Else
''                .AddItem rs!Codigo & Chr(9) & rs!Banco & Chr(9) & rs!Nro & Chr(9) & rs!importe & Chr(9) & rs!fecha_cheque
''            End If
''            rs.MoveNext
''        Wend
''    End With
''FIN:
''    grilla.Redraw = True
''    rs.Close
''    Set rs = Nothing
''    Exit Sub
''ufaErr:
''    Resume FIN
''End Sub
''' ------ No son func grales, se usasn en los 3 frm de cheques------------

'05/10/4'
'agregue func grilla cheq  para amir
'6/1/4  manejo error al abrir daTaenvironment1.Sistema
'24/1/5 comprobanteconcepto null

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
Public Function x2s(ByVal x As Variant)
    ' Pasa TODO a string, en vez de error devuelve "", para sql, pone Punto "."
    On Error Resume Next
    x2s = ""
    If x = "" Then x = "0"
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

Public Function s2n(ByVal s As Variant, Optional Decimales = 2, Optional AlaFuerza As Boolean = False)
'If siError Then On Error Resume Next
Dim ff As String, i As Long
    
    s2n = 0
    If IsMissing(s) Or IsNull(s) Then s = "0"
    If Trim(s) = "" Then s = "0"
    If IsLetra(s) Then s = "0"
    If s = "," Or s = "." Then s = 0
    
    If SeparadorDecimal = "," Then
        If sPuntos(s) >= 1 Then s = CDbl(s)
        s = Trim(Replace(s, ".", ","))
    End If
    s = CDbl(s)
    If IsNumeric(s) Then s2n = Round(CDbl(s), Decimales)  's2n=G_Redondeo(CDbl(s), CLng(Decimales))
    If AlaFuerza Then ff = "0." & Left("000000000", Decimales): s2n = Format(s2n, ff)
    
End Function

Public Function s2t(s As Recordset, dato As String)
    If IsMissing(s(dato)) Or IsNull(s(dato)) Or (s.EOF = True And s.BOF = True) Then
        s2t = ""
    Else
        s2t = s(dato)
    End If
End Function

Public Function IsLetra(s As Variant) As Boolean
Dim Letras As String, L As Long, Palabra As String
Dim m As String
Palabra = UCase(s)
Letras = "QWRTYUIOP[]ASDFGHJKLÑ{}<>ZXCVBNM;:_@!¿¡?*"
IsLetra = False
    For L = 0 To Len(Palabra) - 1
        If InStr(Letras, Mid(Palabra, L + 1, 1)) Then
            IsLetra = True
            Exit For
        End If
    Next
End Function

Public Function sPuntos(ss As Variant) As Long
Dim cadena As String, cantidad As Long, i As Long
Dim tiene_coma As Boolean
cadena = CStr(ss)
cantidad = 0
tiene_coma = False
For i = 0 To Len(cadena)
    If Mid(cadena, i + 1, 1) = "." Then
        cantidad = cantidad + 1
    End If
    If Mid(cadena, i + 1, 1) = "," Then
        tiene_coma = True
    End If
Next
If Not tiene_coma Then
    cantidad = cantidad - 1
End If

sPuntos = cantidad
End Function

Public Function s2n_(ByVal s As Variant, Optional Decimales = 2)
'On Error Resume Next
'    s2n = 0
'    If SeparadorDecimal = "," Then s = Trim(Replace(s, ".", ","))
'    If IsNumeric(s) Then s2n = Round(CDbl(s), Decimales)
End Function

Public Function s2nt(s As Variant)
    ' Pasa TODO a numero, en vez de error devuelve 0
    On Error Resume Next
    s2nt = 0
    
    If SeparadorDecimal() = "," Then
        s = Trim(Replace(s, ".", ","))
    End If
    
    If IsNumeric(s) Then
        s2nt = Round(CDbl(s), 2)
'    Else
'        MsgBox "Esta operación podría causar un error dado que ha ingresado un dato que no es numérico por lo que sera convertido a '0'"
    End If
    
End Function
Public Function d2n(ByVal s As Variant, Optional Decimales = 2) 'para errores de numeros miles con punto
On Error Resume Next
Dim dValor As Double, v As Long, vDetras As String, vEntero As String, vEncontro As Boolean
vEncontro = False
vEntero = ""
vDetras = ""
For v = Len(s) To 1 Step -1
    If vEncontro Then
        If Mid(s, v, 1) = "." Or Mid(s, v, 1) = "," Then
        Else
            vEntero = Mid(s, v, 1) & vEntero
        End If
    Else
        If Mid(s, v, 1) = "." Or Mid(s, v, 1) = "," Then
            vEncontro = True
            If Len(vDetras) >= 3 And Mid(s, v, 1) = "." Then
                vEncontro = False
            End If
        Else
            vDetras = Mid(s, v, 1) & vDetras
        End If
    End If
Next
If vEntero = "" And vEncontro = False Then
    dValor = CDbl(vDetras)
Else
    dValor = CDbl(vEntero & "," & vDetras)
End If
s = dValor

    d2n = 0
    If SeparadorDecimal = "," Then s = Trim(Replace(s, ".", ","))
    If IsNumeric(s) Then d2n = Round(CDbl(s), Decimales)
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
Public Function k2b(ChkValue As Integer) As Boolean ' con valores de checkbox
    k2b = (ChkValue = vbChecked)
End Function
Public Function k2n(ChkValue As Integer) As Integer ' con valores de checkbox ' para bit sqlserver
    k2n = IIf(ChkValue = vbChecked, 1, 0)
End Function
Public Function n2k(bBit As Integer) As Integer ' con valores de checkbox ' para bit sqlserver
    n2k = IIf(Abs(bBit) = 1, vbChecked, vbUnchecked)
End Function

' **************  Raros de Conversion y tipos   ********************


' **************  Grilla  VSFlexGrid ********************
Public Function grillaBuscoVacio(GRILLA As VSFlexGrid, Columna As Long) As Long
    'busca el primer vacio, si no hay, lo agrega
    
    Dim i As Long
   
    With GRILLA
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
    On Error GoTo ERRcarga
    
    Dim ss As String, rs As New ADODB.Recordset
    
    With rs
        .Open sqlUnCampo, DataEnvironment1.Sistema, adOpenDynamic, adLockReadOnly
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
Public Sub FrmBorrarNum(frm As Form)
    On Error Resume Next
    Dim co As Control
    For Each co In frm.Controls
        If TypeOf co Is uNum Then
            co.num = 0
        End If
    Next
End Sub

Public Sub FrmKeyPress(ByRef k As Integer, enter2tab As Boolean, Optional mayusculas As Boolean, Optional SinComodinesSQL As Boolean, Optional calculadora As Boolean)
    On Error Resume Next
    
    If enter2tab And k = 13 Then
        SendKeys "{tab}"
        k = 0
    ElseIf mayusculas And Chr(k) <> UCase(Chr(k)) Then
            k = Asc(UCase(Chr(k)))
    ElseIf SinComodinesSQL And InStr("'*?", Chr(k)) Then
        k = 0
    ElseIf k = 12 Then  ' control L
        calculin
    ElseIf k = 39 Then  '   comilla simple por acento D'Elia a D´Elia
        k = 180
    End If
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
        oH.L = mL
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
Public Function ComboCodigo(Combo As ComboBox) ', Optional IndexSET As Long = -99)
    Dim i As Long
    
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
    cmb.clear
    With rs
        .Open ssql, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
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

    Dim nE, sE, fe, tI, fo
    Dim ss, asse As String
    Dim cnx As New ADODB.Connection
    
    nE = Err.Number
    sE = Err.Description
    
'    Que falle si falla? que aborte ?
    On Error GoTo ufaErr
    
    
'    sE = Error(nE)
    fe = Format(Date, "dd/mm/yy")
    tI = Format(Time, "hh:mm:ss")
    fo = " Frm: " & NombreFormActivo()
    ss = "Err " & nE & " " & sE
    'log
    
    asse = "1"
    
    Open UFA_ARCHIVO_LOG For Append Access Write As #1
    Write #1, fe, tI, ss, msgUsuario, msgLog, fo
    Close #1
    
    asse = "2"
    cnx.Open DataEnvironment1.Sistema.ConnectionString
    ss = Replace(ss, "'", " ")
    cnx.Execute "insert into bsuLog ( mE, mU, mL, mF ) values ( '" & fe & " " & tI & " " & ss & "', '" & msgUsuario & "', '" & msgLog & "', '" & fo & "' )"
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
    Set cnx = Nothing
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

Public Sub logFacturacion(process As String, subProcess As String, annexed As String)
    Dim dateProcess   As String, timeProcess As String
    
    dateProcess = Format(Date, "dd/mm/yy")
    timeProcess = Format(Time, "hh:mm:ss")
    
    Open fileLogFacturacion For Append Access Write As #1
    Write #1, dateProcess, timeProcess, process, subProcess, annexed
    Close #1
    
    DataEnvironment1.Sistema.Execute "insert into [BSLogFacturacion] ( [time],[process],[subprocess],[annexed]) " & _
    " values ( '" & dateProcess & " " & timeProcess & "', '" & process & "', '" & subProcess & "', '" & annexed & "' )"
End Sub

Private Function NombreFormActivo() As String
    On Error Resume Next
    NombreFormActivo = Screen.ActiveForm.Name
End Function

' **************  SQL ********************
'
Public Function obtenerDato(tabla As String, CodigoQueBuscar, campoResultado As String) As Variant
    '  Si codigo es string, pasar entre comitas  " 'codigo ' ",

'    On Error GoTo fin MANEJARLO DO SE LLAMA
    Dim rs As New ADODB.Recordset, ssql  As String
    
    ssql = "select " & campoResultado & " from " & tabla & " where codigo = " & CodigoQueBuscar

    rs.Open ssql, DataEnvironment1.Sistema, adOpenDynamic, adLockReadOnly
    If Not rs.EOF Then
        rs.MoveLast
        obtenerDato = rs.Fields(0)
    End If
fin:
    Set rs = Nothing
End Function
'
Public Function obtenerDeSQL(SelectCampos As String) As Variant
    Dim rs As New ADODB.Recordset, ssql  As String, i As Long, rc As Long, resu As Variant
    
    With rs
        
        .Open SelectCampos, DataEnvironment1.Sistema, adOpenDynamic, adLockReadOnly
        
        rc = .Fields.Count
        
        If rc = 1 Then
            If Not .BOF Then
                .MoveFirst
                obtenerDeSQL = .Fields(0)
            End If
        Else
            If Not .BOF Then
                .MoveFirst
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
Public Function obtenerDatoS(tabla As String, CodigoQueBuscar, campoResultado As String) As Variant
'On Error GoTo FIN
    Dim rs As New ADODB.Recordset, ssql  As String
    
    ssql = "select " & campoResultado & " from " & tabla & " where codigo = '" & CodigoQueBuscar & "'"

    rs.Open ssql, DataEnvironment1.Sistema, adOpenDynamic, adLockReadOnly
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

    rs.Open ssql, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then
        rs.MoveLast
        obtenerDeS = rs.Fields(0)
    End If
fin:
    Set rs = Nothing
End Function
' ss funciones SQL
' string sql para concatenar en el where
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

Public Function ssFecha(dfecha As Date) As String
    ssFecha = " '" & Format(dfecha, "YYYYMMDD") & "' "
End Function

Public Function afipFecha(dfecha As Date) As String
    afipFecha = "" & Format(dfecha, "YYYYMMDD") & ""
End Function

Public Function aFecha(dfecha As String) As Date
Dim a As Long, m As Long, d As Long
a = CORTO(dfecha, 0, 4)
m = CORTO(dfecha, 4, 2)
d = CORTO(dfecha, 6, 0)
    aFecha = CDate(d & "/" & m & "/" & a)
End Function

Public Function ssBetween(dDesde As Date, dHasta As Date) As String
    Dim s1 As String, s2 As String
    s1 = Format(dDesde, "yyyymmdd")
    s2 = Format(dHasta, "yyyymmdd")
    ssBetween = "  between  '" & s1 & "' and '" & s2 & "' "
End Function
Public Function ssNum(ByVal numero) As String
    ' Pasa Numero a string, formato SQL, no usuario
    On Error Resume Next
    ssNum = Trim(CStr(numero))
    If SeparadorDecimal = "," Then ssNum = Replace(ssNum, ",", ".")
End Function

' ************** SQL ********************

' ************** UTILITY ********************
Public Function nSinNull(que)
    nSinNull = IIf(IsNull(que), 0, que)
    If nSinNull = "Null" Then nSinNull = 0
End Function
Public Function sSinNull(que)
    sSinNull = IIf(IsNull(que), "", que)
End Function

Public Function fSinNull(que)
    fSinNull = IIf(IsNull(que), Date, que)
End Function

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

'**************** daTaenvironment1.Sistema ********************************** SACAR DE ACA!!!!
Public Function UsuarioActual()
On Error GoTo mal2
    UsuarioActual = UsuarioSistema!codigo
Exit Function
mal2:
    UsuarioActual = 100
End Function
Public Sub grabaBitacora(Abm, COD, tabla)
    DataEnvironment1.dbo_GRABARBITACORA COD, tabla, UsuarioSistema!codigo, Date, Time, Abm
End Sub
Public Function dEnvTxt()
    Dim i
    Open ".\data_environment.log" For Output Access Write As #1
    For i = 1 To DataEnvironment1.Commands.Count - 1
        Write #1, DataEnvironment1.Commands.item(i).Name
    Next i
    Close #1
End Function
'**************** daTaenvironment1.Sistema ********************************** SACAR DE ACA!!!!


Public Sub lMsg(lineas As Variant) 'porquria, mas facil acordate del VbCrLf
    Dim s As String, L
    For Each L In lineas
        s = s & vbCrLf & L
    Next
    MsgBox s
End Sub
Public Sub che(que)
    MsgBox que, vbExclamation, "Atencion"
End Sub

Public Function confirma(sMsg As String) As Boolean
    confirma = (MsgBox(sMsg, vbQuestion + vbYesNo, "Pregunta") = vbYes)
End Function

Public Function calculin()
    On Error Resume Next
    calculin = frmCalculin.mostrar()
End Function

Public Sub SubimeSi800x600(Optional formulario As Form)
    '*** Poner en        form_Activate() : SubimeSi800x600
    '*** NO SIRVE en     form_load()
    ' usar solo si altura queda justa, porque se vuelve a mover cada vez q activa
   
    On Error Resume Next
'    If formulario Is nothing Then Set formulario = Screen.ActiveForm
    If Screen.Height < 9500 And Screen.ActiveForm.Top > 0 Then Screen.ActiveForm.Top = -30
End Sub

Public Sub relojito(Optional poner As Boolean = True)
    On Error Resume Next
    Screen.ActiveForm.MousePointer = IIf(poner, vbHourglass, vbDefault)
End Sub


Public Function AS_Base_2_Arch(campo As Field, Archivo As String)
    ' ADO Stream
    ' AS_Base_2_Arch(rs!logo, "c:\temp\loguito.jpg")
    
    Dim mstream As New ADODB.Stream
    
    mstream.Type = adTypeBinary
    mstream.Open
    mstream.Write campo.Value
    mstream.SaveToFile Archivo, adSaveCreateOverWrite
    
    Set mstream = Nothing
End Function
Public Function AS_Arch_2_Base(campo As Field, Archivo As String) ' rs as adodb.recordset )
    Dim mstream As New ADODB.Stream
    
    mstream.Type = adTypeBinary
    mstream.Open
    mstream.LoadFromFile Archivo
    campo.Value = mstream.Read
    'rs.Update
    
    Set mstream = Nothing
End Function

Public Sub grillaWidth(GRILLA, arrayAnchos)
    On Error GoTo fin
    Dim i As Long
    With GRILLA
        For i = 0 To UBound(arrayAnchos): .ColWidth(i) = arrayAnchos(i): Next
    End With
fin:
End Sub
Public Sub grillaTitulos(GRILLA, arrayTit)
    On Error GoTo fin
    Dim i As Long
    With GRILLA
        For i = 0 To UBound(arrayTit): .TextMatrix(0, i) = arrayTit(i): Next
    End With
fin:
End Sub

Public Sub grillaSumarizo(GRILLA, arraycolumnas)
    'solo flexgrid
    On Error GoTo fin
    Dim i As Long
    With GRILLA
        .SubtotalPosition = flexSTBelow
        For i = 0 To UBound(arraycolumnas)
            .subtotal flexSTSum, -1, arraycolumnas(i), , , , True, , , True
        Next
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
Public Function ligriwi(GRILLA As Object)
    ' la uso para averiguar anchos en la ventana inmediato, pego resultado a grillawidth
    Dim i As Long, s As String
    
    s = "array("
    For i = 0 To GRILLA.cols - 1: s = s & GRILLA.ColWidth(i) & ",": Next
    s = Left(s, Len(s) - 1)
    s = s & ")"
    
    Debug.Print s
End Function
Public Function liRsCampos(queRS As ADODB.Recordset)
    ' la uso para averiguar campos de un RS
    Dim i As Long
    With queRS
        For i = 0 To .Fields.Count - 1
             Debug.Print .Fields(i).Name, .Fields(i).Type
        Next i
    End With
End Function

Sub CargarHelpCuentas(tabla As String, nomCampo1 As String, nomCampo2 As String, Campo1 As String, Campo2 As String, Optional order As String)

Dim rs As New ADODB.Recordset, ingreso As Long

FrmHelp.grillahelp.Row = 1
FrmHelp.grillahelp.FormatString = "^" + nomCampo1 + "        |<" + nomCampo2 + "                                                    "


    If Trim(order) <> "" Then
        rs.Open "select * from " + tabla & " where imputable = 1 and activo = 1 order by " & order, DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
    Else
        'rs.Open "select * from " + tabla & " where imputable = 1 and activo = 1 order by _Codigo", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
        rs.Open "select * from " + tabla & " where imputable = 1 and activo = 1 order by cuenta", DataEnvironment1.Sistema, adOpenStatic, adLockOptimistic
    End If


    If Not rs.EOF Then
        rs.MoveFirst
        'Agrego la primera linea sin hacer additem
        ingreso = 0
        If esimputable(rs.Fields(Campo1)) Then
            FrmHelp.grillahelp.Col = 0
            FrmHelp.grillahelp.Text = Trim(rs(Campo1))
            FrmHelp.grillahelp.Col = 1
            FrmHelp.grillahelp.Text = Trim(rs(Campo2))
            ingreso = 1
        End If
        rs.MoveNext
        Do While Not rs.EOF
            If esimputable(rs.Fields(Campo1)) Then
                If ingreso = 1 Then
                    FrmHelp.grillahelp.AddItem Trim(rs(Campo1)) & Chr(9) & Trim(rs(Campo2))
                Else
                    FrmHelp.grillahelp.Col = 0
                    FrmHelp.grillahelp.Text = Trim(rs(Campo1))
                    FrmHelp.grillahelp.Col = 1
                    FrmHelp.grillahelp.Text = Trim(rs(Campo2))
                    ingreso = 1
                End If
            End If
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing
End Sub


Public Function G_Redondeo(numero As Double, Optional Decima As Long) As Double
    Dim NumC As Double
    Dim num  As String
    Dim Deci As String
    
    Dim entero As Long
    Dim positivo As Double
    entero = Fix(numero)
    positivo = Abs(numero)
    entero = Len(Right(positivo, Len(CStr(positivo)) - 1 - Len(CStr(entero))))
    
    NumC = numero - Fix(numero)
    'If Len(CStr(NumC)) - 2 <= Decima Then
    If entero <= Decima Then
        G_Redondeo = Round(numero, Decima)
        Exit Function
    Else
        'el 3 es el 0, de Left(CStr(NumC) (=2) + 1, que es el ultimo decimal, el q redondeamos
        num = CStr(Fix(numero) + CDbl(Left(CStr(NumC), Decima + 3)))
    End If
    
    num = CStr(numero)
    Deci = Right$(num, 1)
    ''Si es 5 hacemos el redondeo hacia arriba, por eso el 6
    If Deci = "5" Then
        Deci = "6"
    End If
    num = Left$(num, Len(num) - 1) & Deci
    If IsMissing(Decima) Then
        G_Redondeo = Round(CDbl(num))
    Else
        G_Redondeo = Round(CDbl(num), Decima)
    End If
    
End Function


'Public Sub kk_re()
'    Dim rs As New ADODB.Recordset
'
'    rs.Open "select tempo from _dbf_prov where cod_pr = 1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
'        AS_Arch_2_Base rs!tempo, "c:\recibos.xls"
'        rs.Update
'    rs.Close
'
'
'    rs.Open "select tempo from _dbf_prov where cod_pr = 1", DataEnvironment1.Sistema, adOpenForwardOnly, adLockOptimistic
'        AS_Base_2_Arch rs!tempo, "c:\recibosNew.xls"
'    rs.Close
'
'    Set rs = Nothing
'End Sub



''public function strExtraer(sTexto as String,
'
'Public Sub sumaSeparada(aCuentas, vCuenta, nMonto)
'    Dim que, i As Integer, largo As Integer
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
'2/9/5
'   ssfecha ssbetwenn formato YYYYMMDD
'6/12/5
'   relojito()
'10/3/6 ssStr() le puse opcion para busquedas, para campos string ya grabados con '
'11/5/6
'   grilla
'31/5/6
'   o no se cuando, AS_  AdoStream, grabar/leer archivos al SQL
'27/8/6 (domingo) frmBorrarNum para borar mi uc numerico
'


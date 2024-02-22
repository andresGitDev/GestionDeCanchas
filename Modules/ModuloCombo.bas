Attribute VB_Name = "ModuloCombo"
Option Explicit '16/9/4


Sub CargaCombo(Combo As Object, tabla As String, campo As String, Bound As String, wer As String)

Dim rsCargacombo As New ADODB.Recordset
Dim sqlstrCC As String
Dim i As Long
    If Bound <> "" Then
        sqlstrCC = "Select " + campo + " as NN," + Bound + " from " + tabla + " where activo=1"
    Else
        sqlstrCC = "Select " + campo + " as NN" + Bound + " from " + tabla + " where activo=1"
    End If
    If wer <> "" Then
        sqlstrCC = sqlstrCC + " and " + wer
    End If
    sqlstrCC = sqlstrCC + " order by " + campo
    rsCargacombo.Open sqlstrCC, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
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
    Set rsCargacombo = Nothing
End Sub

Sub CargaCombo2(Combo As Object, tabla As String, campo As String, Bound As String, wer As String)

Dim rsCargacombo As New ADODB.Recordset
Dim sqlstrCC As String
Dim i As Long
    
    If Bound <> "" Then
        sqlstrCC = "Select " + campo + " as NN," + Bound + " from " + tabla
    Else
        sqlstrCC = "Select " + campo + " as NN" + Bound + " from " + tabla
    End If
    If wer <> "" Then
        sqlstrCC = sqlstrCC + " where " + wer + " and activo=1"
    Else
        sqlstrCC = sqlstrCC + " where activo=1"
    End If
    sqlstrCC = sqlstrCC + " order by " + campo
    rsCargacombo.Open sqlstrCC, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    Combo.clear
    If Not rsCargacombo.EOF And Not rsCargacombo.BOF Then
        rsCargacombo.MoveFirst
        i = 0
        While Not rsCargacombo.EOF
            Combo.AddItem Trim(rsCargacombo.Fields("NN"))
            Combo.ItemData(i) = val(rsCargacombo!codigo)
            i = i + 1
            rsCargacombo.MoveNext
        Wend
    End If
    rsCargacombo.Close
    Set rsCargacombo = Nothing
End Sub

Sub CargaCombo3(Combo As Object, tabla As String, campo As String, Bound As String, wer As String)

Dim rsCargacombo As New ADODB.Recordset
Dim sqlstrCC As String
Dim i As Long
    
    If Bound <> "" Then
        sqlstrCC = "Select " + campo + " as NN," + Bound + " from " + tabla
    Else
        sqlstrCC = "Select " + campo + " as NN" + Bound + " from " + tabla
    End If
    If wer <> "" Then
        sqlstrCC = sqlstrCC + " where " + wer + " and activo=1"
    Else
        sqlstrCC = sqlstrCC + " where activo=1"
    End If
    'sqlstrCC = sqlstrCC + " order by " + Campo
    rsCargacombo.Open sqlstrCC, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    Combo.clear
    If Not rsCargacombo.EOF And Not rsCargacombo.BOF Then
        rsCargacombo.MoveFirst
        i = 0
        While Not rsCargacombo.EOF
            Combo.AddItem Trim(rsCargacombo.Fields("NN"))
            Combo.ItemData(i) = val(rsCargacombo!codigo)
            i = i + 1
            rsCargacombo.MoveNext
        Wend
    End If
    rsCargacombo.Close
    Set rsCargacombo = Nothing
End Sub


Function BuscarEnCombo(Combo As Object, Valor As Variant, Optional EnList As Boolean = False) As Long
    Dim i As Long
    
    i = 0
    If VarType(Valor) = vbNull Then
        Valor = -1
    End If
    If val(Valor) > -1 Then
        If Not EnList Then
            Do While i < Combo.ListCount
                If CInt(Combo.ItemData(i)) = CInt(Valor) Then Exit Do
                i = i + 1
            Loop
        Else
            Do While i < Combo.ListCount
                If Combo.List(i) = Valor Then Exit Do
                i = i + 1
            Loop
        End If
        BuscarEnCombo = IIf(i = Combo.ListCount, -1, i)
    Else
        BuscarEnCombo = -1
    End If
    
End Function


'agregado from pablo
Function BuscarenComboS(Combo As ComboBox, Texto As String) As Long
    Dim i As Long

    BuscarenComboS = -1
    With Combo
        i = 0
        While i < .ListCount And BuscarenComboS = -1
            If UCase(.List(i)) = UCase(Texto) Then
                .ListIndex = i
                BuscarenComboS = i
            End If
            i = i + 1
        Wend
    End With
    
End Function

Public Function BuscoCodigoProvincia(sDato As String) As String
Dim codProvincia As String
codProvincia = sSinNull(obtenerDeSQL("select codigo from provincias where descripcion like '%" & Trim(sDato) & "%'"))
If Trim(codProvincia) = "" Then
    If sDato = "CIUDAD AUTONOMA BUENOS AIRES" Then
        codProvincia = "*"
    Else
        codProvincia = "B"
    End If
End If
BuscoCodigoProvincia = codProvincia
End Function




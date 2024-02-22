Attribute VB_Name = "ModuloSebastian"
Option Explicit

Public Const Operacion_ALTA = "A"
Public Const Operacion_BAJA = "B"
Public Const Operacion_MODIFICACION = "M"

Public Const Estado_PENDIENTE = "PENDIENTE"
Public Const Estado_FACTURADO = "FACTURADO"

Public Const FACTURA_A = 1

Public Function ObtenerDatoDB(Tabla As String, ColumnaABuscar As String, DatoABuscar, ColumnaADevolver As String) As Variant
'Busca DATOABUSCAR en la COLUMNAABUSCAR en TABLA me devuelve el contenido de la COLUMNAADEVOLVER

Dim Consulta As String
Dim rsAux As New ADODB.Recordset

    If Tabla <> "" And ColumnaABuscar <> "" And ColumnaADevolver <> "" And DatoABuscar <> "" Then
        If IsNumeric(DatoABuscar) Then
            Consulta = "Select " & ColumnaADevolver & " From " & Tabla & " Where " & ColumnaABuscar & " = " & DatoABuscar
        Else
            Consulta = "Select " & ColumnaADevolver & " From " & Tabla & " Where " & ColumnaABuscar & " = '" & DatoABuscar & "'"
        End If
        rsAux.Open Consulta, DataEnvironment1.AMR, adOpenDynamic, adLockOptimistic
        If Not rsAux.EOF Then ObtenerDatoDB = rsAux.Fields(0)
        rsAux.Close
        Set rsAux = Nothing
        
    End If
End Function

Public Sub LimpiarGrilla(grilla As Control, Optional Filas As Long = 2, Optional Columnas As Long = 2)
    grilla.Clear
    grilla.rows = Filas
    grilla.cols = Columnas
End Sub

Public Function LlenarGrilla(grilla As Control, ConsultaSQL As String, AjustarAnchos As Boolean) As Boolean
Dim rsAux As New ADODB.Recordset
Dim C As Long
Dim Encabezado As String

    If ConsultaSQL <> "" Then
        grilla.Clear
        rsAux.Open ConsultaSQL, DataEnvironment1.AMR, adOpenForwardOnly, adLockReadOnly
        If Not rsAux.EOF Then
            With rsAux
            
                'hago el encabezado de la grilla
                grilla.FixedCols = 0
                grilla.Row = 0
                For C = 0 To .Fields.Count - 1
                    If C > 0 Then Encabezado = Encabezado & "|"
                    Encabezado = Encabezado & .Fields(C).Name
                Next C
                grilla.FormatString = Encabezado
                
                 'modifico los anchos de las columnas
                If AjustarAnchos Then
                    grilla.Row = 0
                    For C = 0 To .Fields.Count - 1
                        Select Case .Fields(C).Type
                            Case adVarChar '200
                                grilla.ColWidth(C) = 3000
                            Case adInteger
                                grilla.ColWidth(C) = 1000
                            Case adDouble
                                grilla.ColWidth(C) = 1000
                            Case adDate
                                grilla.ColWidth(C) = 1000
                            Case adBoolean
                                grilla.ColWidth(C) = 200
                            Case Else
                                grilla.ColWidth(C) = 1000
                            
                        End Select
                    Next C
                End If
                
                'lleno la grilla con los datos de la consulta
                grilla.cols = .Fields.Count
                grilla.rows = 1
                While Not .EOF
                    grilla.rows = grilla.rows + 1
                    For C = 0 To .Fields.Count - 1
                        grilla.TextMatrix(grilla.rows - 1, C) = IIf(IsNull(.Fields(C)), "", .Fields(C))
                    Next C
                    .MoveNext
                Wend
            End With
            LlenarGrilla = True
        Else
'            MsgBox "No hay datos para mostrar", 48, "Atencion"
            LlenarGrilla = False
        End If
        rsAux.Close
        Set rsAux = Nothing
    End If
End Function

Public Function LlenarGrilla2(grilla As Control, ConsultaSQL As String, AjustarAnchos As Boolean) As Boolean
Dim rsAux As New ADODB.Recordset
Dim C As Long
Dim Encabezado As String

    If ConsultaSQL <> "" Then
        grilla.Clear
        rsAux.Open ConsultaSQL, DataEnvironment1.AMR, adOpenForwardOnly, adLockReadOnly
        If Not rsAux.EOF Then
            With rsAux
            
                'hago el encabezado de la grilla
                grilla.FixedCols = 0
                grilla.Row = 0
                For C = 0 To .Fields.Count - 1
                    If C > 0 Then Encabezado = Encabezado & "|"
                    Encabezado = Encabezado & .Fields(C).Name
                Next C
                grilla.FormatString = Encabezado
                
                 'modifico los anchos de las columnas
                If AjustarAnchos Then
                    grilla.Row = 0
                    For C = 0 To .Fields.Count - 1
                        Select Case .Fields(C).Type
                            Case adVarChar '200
                                grilla.ColWidth(C) = 3000
                            Case adInteger
                                grilla.ColWidth(C) = 1000
                            Case adDouble
                                grilla.ColWidth(C) = 1000
                            Case adDate
                                grilla.ColWidth(C) = 1000
                            Case adBoolean
                                grilla.ColWidth(C) = 200
                            Case Else
                                grilla.ColWidth(C) = 1000
                            
                        End Select
                    Next C
                End If
                
                'lleno la grilla con los datos de la consulta
                grilla.cols = .Fields.Count
                grilla.rows = 1
                While Not .EOF
                    grilla.rows = grilla.rows + 1
                    For C = 0 To .Fields.Count - 1
                        If C = 0 Then
                            If .Fields(5) = "" Or IsNull(.Fields(5)) Then
                                grilla.TextMatrix(grilla.rows - 1, C) = IIf(IsNull(.Fields(C)), "", .Fields(C))
                            Else
                                grilla.TextMatrix(grilla.rows - 1, C) = IIf(IsNull(.Fields(C)), "", 1)
                            End If
                        Else
                            grilla.TextMatrix(grilla.rows - 1, C) = IIf(IsNull(.Fields(C)), "", .Fields(C))
                        End If
                    Next C
                    .MoveNext
                Wend
            End With
            LlenarGrilla2 = True
        Else
'            MsgBox "No hay datos para mostrar", 48, "Atencion"
            LlenarGrilla2 = False
        End If
        rsAux.Close
        Set rsAux = Nothing
    End If
End Function


Public Function ExisteDato(Tabla As String, Columna As String, DatoABuscar As Variant) As Boolean
Dim rsAux As New ADODB.Recordset
Dim Consulta As String

    If IsNumeric(DatoABuscar) Then
        Consulta = "Select * From " & Tabla & " Where " & Columna & " = " & DatoABuscar
    Else
        Consulta = "Select * From " & Tabla & " Where " & Columna & " = '" & DatoABuscar & "'"
    End If
    rsAux.Open Consulta, DataEnvironment1.AMR, adOpenDynamic, adLockOptimistic
    ExisteDato = Not rsAux.EOF
    rsAux.Close
    Set rsAux = Nothing
    
End Function


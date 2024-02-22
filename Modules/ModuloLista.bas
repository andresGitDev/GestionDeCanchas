Attribute VB_Name = "ModuloLista"
Option Explicit '16/9/4



Sub CargaLista(Lista As Object, Tabla As String, orden As String, Bound As String)

Dim rsCarga As New ADODB.Recordset
Dim i As Long
Dim PrimerDato, SegundoDato As String

    PrimerDato = "CODIGO"
    SegundoDato = "DESCRIPCION"
        
    rsCarga.Open "Select * " + "," + Bound + " from " + Tabla + " order by " + orden, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    Lista.clear
    If Not rsCarga.EOF And Not rsCarga.BOF Then
        rsCarga.MoveFirst
        i = 0
        If orden = "DESCRIPCION" Then
            PrimerDato = "DESCRIPCION"
            SegundoDato = "CODIGO"
        End If
        While Not rsCarga.EOF
            If (IsNull(Trim(rsCarga.Fields("DESCRIPCION"))) _
            Or (Trim(rsCarga.Fields("DESCRIPCION")) = "")) Then 'pregunta si el campo DESCRIPCION es nulo o esta vacio
                
                If orden = "DESCRIPCION" Then ' pregunta el orden para saber adonde completar el campo nulo, en 1er o 2do lugar
                    Lista.AddItem "Sin Descripcion - " + Trim(rsCarga.Fields(SegundoDato))
                Else
                    Lista.AddItem (Trim(rsCarga.Fields(PrimerDato)) + " - Sin Descripcion")
                End If
            
            Else
                    
                Lista.AddItem (Trim(rsCarga.Fields(PrimerDato)) + " - " + Trim(rsCarga.Fields(SegundoDato)))
            
            End If
            
            Lista.ItemData(i) = i
            i = i + 1
            rsCarga.MoveNext
        Wend
    End If
    rsCarga.Close
    Set rsCarga = Nothing

End Sub


Public Function SerieStockRepetida(producto As String) As String
'    On Error GoTo UfaBuscaSer
    Dim ss As String, tmpTablaSeries As String
    Dim rs As New ADODB.Recordset
    
    'Dim kk As long, kk2 '*******************************
    'kk2 = Timer
    
    tmpTablaSeries = TablaTempCrear(tt_SeriesEnStockTemp)

If producto = "" Then
    ss = "INSERT INTO " & tmpTablaSeries & " ( Producto, Serie ) " _
    & " SELECT producto, serie From Series " _
    & " Where activo = 1  "

    DataEnvironment1.Sistema.Execute ss

    ss = "SELECT t.Serie as [ Serie               ], s.Producto as [ Producto                  ], t.Descripcion as [ Descripcion                                              ], s.comprobante as [c], t.codigo  as [i]" _
    & " FROM " & tmpTablaSeries & " AS t INNER JOIN Series AS s ON t.codigo = s.codigo left join conceptos as c on c.codigo = s.concepto " _
    & " WHERE s.comprobante = 6 or s.comprobante = 3 or s.comprobante = 4 or (s.comprobante = 7 and c.movimiento <> 'R' )  "

    SerieStockRepetida = frmBuscar.MostrarSql(ss)
    Exit Function
End If

ss = "SELECT DISTINCT producto,Serie FROM Series " _
     & " Where Series.activo = 1 and producto = '" & producto & "'" _
     & " ORDER BY Series.serie "
    
'ss = "SELECT DISTINCT producto,Serie FROM Series " _
'     & " Where Series.activo = 1 and producto = '" & producto & "' and fecha > '20000101'" _
'     & " ORDER BY Series.serie "
'
    
rs.Open (ss), DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly

    'kk = 0

Do While rs.EOF = False
    'kk = kk + 1
     If SerieEnStock((rs!Serie), producto) Then
           
      ss = "INSERT INTO " & tmpTablaSeries & " (Producto, Serie) " _
       & " VALUES ('" & producto & "','" & sSinNull(rs!Serie) & "')"
    
       DataEnvironment1.Sistema.Execute ss
         
     End If
     rs.MoveNext
Loop
        

ss = "SELECT Serie as [ Serie               ], Producto as [ Producto                  ]" _
    & " FROM " & tmpTablaSeries & " "
        
'        MsgBox kk & "   " & Timer - kk2

SerieStockRepetida = frmBuscar.MostrarSql(ss)
    
fin:
    Exit Function
UfaBuscaSer:
    ufa "err: buscando series", "Prod: " & producto, Err.Description
    Resume fin
End Function



Attribute VB_Name = "ModuloMarce"
Option Explicit

'Verifica q si el producto y serie ESTA O NO ESTA en el STOCK
Public Function SerieEnStock(ByRef cualSerie As String, cualProducto As String) As Boolean
Dim CantEntrada As Long
Dim CantSalida As Long


'Dim tttt
'
'Static kk As long
'kk = kk + 1
CantEntrada = s2n(obtenerDeSQL("SELECT count(EsSalida)as Entrada FROM Series " & _
    " where serie = '" & cualSerie & "' and producto = '" & cualProducto & "' " & _
    " and EsSalida = 0 and activo = 1"))

CantSalida = s2n(obtenerDeSQL("select count(esSalida) as Salida from series" & _
    " where serie = '" & cualSerie & "' and producto ='" & cualProducto & "' " & _
    " and EsSalida = 1 and activo = 1"))
    
'
'tttt = (obtenerDeSQL("select sum(esSalida) as kk2 from series" & _
'    " where serie = '" & cualSerie & "' and producto ='" & cualProducto & "' "))
  
  
'MsgBox tttt & "        " & CantSalida - CantEntrada
  
'If CantSalida >= CantEntrada Then
''If tttt < 0 Then
'  SerieEnStock = False              'producto salio o no esta
'   Else
' SerieEnStock = True                'producto dentro
'End If

''MsgBox Timer - kk

    SerieEnStock = (CantSalida < CantEntrada)

End Function

'Funcion para saber q un producto y serie ESTUVIERON ALGUNA VEZ en stock
Public Function NuncaEstuvo(ByRef cualSerie As String, cualProducto As String) As Boolean
Dim sql As String
Dim rs As New ADODB.Recordset

sql = "SELECT producto FROM Series " & _
" where serie = '" & cualSerie & "' and producto = '" & cualProducto & "' "

rs.Open (sql), DataEnvironment1.AMR, adOpenDynamic, adLockOptimistic

If rs.EOF = True Then
   NuncaEstuvo = True      'no se encuentra el producto
  Else
   NuncaEstuvo = False
End If
rs.Close
End Function

'Funcion para q NO se registre dos veces a un producto
Public Function ProdSerieRepetida(ByRef cualSerie As String, cualProducto As String) As Boolean
Dim sql As String
Dim CantEntrada As Long
Dim rs As New ADODB.Recordset

CantEntrada = sSinNull(obtenerDeSQL("SELECT count(EsSalida)as Entrada FROM Series " & _
    " where serie = '" & cualSerie & "' and producto = '" & cualProducto & "' " & _
    " and EsSalida = 0"))

If CInt(CantEntrada) > 1 Then                                  'para las q estan en stock y pasaron varias veces
   ProdSerieRepetida = False
  Else
    If CInt(CantEntrada) = 1 Or CInt(CantEntrada) = 0 Then         'si entro UNA SOLA VEZ
      ProdSerieRepetida = False                                    'o NO entro NUNCA es un
     Else                                                          'producto y serie NUEVA
      ProdSerieRepetida = True
    End If
End If
End Function
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

    DataEnvironment1.AMR.Execute ss

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
    
rs.Open (ss), DataEnvironment1.AMR, adOpenForwardOnly, adLockReadOnly

    'kk = 0

Do While rs.EOF = False
    'kk = kk + 1
     If SerieEnStock((rs!Serie), producto) Then
           
      ss = "INSERT INTO " & tmpTablaSeries & " (Producto, Serie) " _
       & " VALUES ('" & producto & "','" & sSinNull(rs!Serie) & "')"
    
       DataEnvironment1.AMR.Execute ss
         
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
      
      

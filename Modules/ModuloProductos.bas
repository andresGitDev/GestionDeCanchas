Attribute VB_Name = "ModuloProductos"
Option Explicit '16/9/4

Public MODO_ON_ERROR_ABM_ON As Boolean

Public Function HayStock(cod As String) As Long

Dim rs As New ADODB.Recordset


rs.Open "select existencia from Producto where CODIGO = '" & Trim(cod) & "'", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly

If Not rs.EOF Then
    If Not IsNull(rs!existencia) Then
        HayStock = rs!existencia
    Else
        HayStock = 0
    End If
Else
    HayStock = 0
End If

rs.Close
Set rs = Nothing

End Function
Public Function ObtenerDisponibleProducto(prod As String) As Variant
    Dim rsped As New ADODB.Recordset
    Dim pendiente As Variant

    rsped.Open "Select saldo from itempedidocliente where producto='" & prod & "' and saldo<>0", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
    pendiente = 0
    If Not rsped.EOF Then
        Do While Not rsped.EOF
            pendiente = s2n(pendiente) + s2n(rsped!saldo)
            rsped.MoveNext
        Loop
    End If
    rsped.Close
    Set rsped = Nothing
    rsped.Open "select existencia from producto where codigo='" & prod & "'", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    If Not rsped.EOF Then
        ObtenerDisponibleProducto = s2n(rsped!existencia - s2n(pendiente))
    Else
        ObtenerDisponibleProducto = 0
    End If
End Function
Public Function ObtenerPrecioProducto(prod As String, cli As Long) As Double
Dim rs As New ADODB.Recordset
    If cli <> 0 Then
        rs.Open "Select * from relacion_producto_cliente where producto='" & prod & "' and cliente=" & cli, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
        If Not rs.EOF Then
            ObtenerPrecioProducto = rs!precio
        Else
            ObtenerPrecioProducto = 0
        End If
        rs.Close
        Set rs = Nothing
    End If

End Function

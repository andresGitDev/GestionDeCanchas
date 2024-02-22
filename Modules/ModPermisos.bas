Attribute VB_Name = "ModPermisos"
'los true me dice que tengo permiso

Public Function permiteBuscar() As Boolean
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
    
    permiteBuscar = False 'hasta que demuestre lo contrario
        
    rs2.Open "select tipousuario from usuarios where codigo=" & UsuarioActual() & " and activo=1", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    
    'administrador
    If rs2!TIPOUSUARIO = 1 Then
        permiteBuscar = True
        Exit Function
    End If
    
    rs.Open "select permisosxusuario.*,permisosespeciales.descripcion from permisosxusuario inner join permisosespeciales on permisosxusuario.permiso=permisosespeciales.codigo where grupo=" & rs2!TIPOUSUARIO & " and activo=1 ", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    Do While Not rs.EOF
        If rs!permiso > 999 Then 'esto es boton
            If rs!descripcion = "Buscar" Then
                permiteBuscar = True
            End If
        End If
        rs.MoveNext
    Loop
    
    Set rs = Nothing
    Set rs2 = Nothing
    
End Function
Public Function permiteAceptar() As Boolean
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
    
    permiteAceptar = False 'hasta que demuestre lo contrario
        
    rs2.Open "select tipousuario from usuarios where codigo=" & UsuarioActual() & " and activo=1", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    
    'administrador
    If rs2!TIPOUSUARIO = 1 Then
        permiteAceptar = True
        Exit Function
    End If
    
    rs.Open "select permisosxusuario.*,permisosespeciales.descripcion from permisosxusuario inner join permisosespeciales on permisosxusuario.permiso=permisosespeciales.codigo where grupo=" & rs2!TIPOUSUARIO & " and activo=1 ", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    Do While Not rs.EOF
        If rs!permiso > 999 Then 'esto es boton
            If rs!descripcion = "Aceptar" Then
                permiteAceptar = True
            End If
        End If
        rs.MoveNext
    Loop
    
    Set rs = Nothing
    Set rs2 = Nothing
    
End Function
Public Function permiteEliminar() As Boolean
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
    
    permiteEliminar = False 'hasta que demuestre lo contrario
        
    rs2.Open "select tipousuario from usuarios where codigo=" & UsuarioActual() & " and activo=1", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    
    'administrador
    If rs2!TIPOUSUARIO = 1 Then
        permiteEliminar = True
        Exit Function
    End If
    
    rs.Open "select permisosxusuario.*,permisosespeciales.descripcion from permisosxusuario inner join permisosespeciales on permisosxusuario.permiso=permisosespeciales.codigo where grupo=" & rs2!TIPOUSUARIO & " and activo=1 ", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    Do While Not rs.EOF
        If rs!permiso > 999 Then 'esto es boton
            If rs!descripcion = "Eliminar" Then
                permiteEliminar = True
            End If
        End If
        rs.MoveNext
    Loop
    
    Set rs = Nothing
    Set rs2 = Nothing
    
End Function
Public Function permiteModificar() As Boolean
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
    
    permiteModificar = False 'hasta que demuestre lo contrario
        
    rs2.Open "select tipousuario from usuarios where codigo=" & UsuarioActual() & " and activo=1", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    
    'administrador
    If rs2!TIPOUSUARIO = 1 Then
        permiteModificar = True
        Exit Function
    End If
    
    rs.Open "select permisosxusuario.*,permisosespeciales.descripcion from permisosxusuario inner join permisosespeciales on permisosxusuario.permiso=permisosespeciales.codigo where grupo=" & rs2!TIPOUSUARIO & " and activo=1 ", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    Do While Not rs.EOF
        If rs!permiso > 999 Then 'esto es boton
            If rs!descripcion = "Modificar" Then
                permiteModificar = True
            End If
        End If
        rs.MoveNext
    Loop
    
    Set rs = Nothing
    Set rs2 = Nothing
    
End Function


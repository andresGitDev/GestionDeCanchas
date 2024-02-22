Attribute VB_Name = "ModuloSeguridad"
Option Explicit

' extensible, Vinculado a tabla USUARIOAREA.AreaSeguridad
Public Enum AreaSeguridad
    AreaContable = 2
    AreaStock = 4
    AreaPersonal = 3
End Enum


'
' 0) Asumo recordset usuarioActivo abierto
'


'    *** Programacion general ***
'
' 1) Definir Enum:  AreaSeguridad
'


'    *** tareas que en un futuro se puede hacer por prg) ***
'
' 2) poner a cada usuario su tipo
'
' 3) Definir UsuarioArea,
'       para cada usuario
'       agregar un registro de AreaSeguridad
'       asignarle permisos qiue sea la suma del enum AreaPermisos
'



'    *** Programacion particular ***
'
' 4) para saber si tengo permiso para una tarea, usar function
'
'     if puedo(Altas, AreaStock) then
'
'nota:
'       puedo(Altas, AreaStock) or puedo(Modis, AreaStock)
'
'   equivale a
'
'       puedo(Altas + Modis, AreaStock)
'
' 5) puedo poner al form
'
'    Public Function Area() As AreaSeguridad
'        Area = AreaStock
'    End Function
'
' asi se simplifico el llamado a puedo() dentro del form
' si no paso area, la busca del form activo OJO
'


'*************************************************
' NO MODIFICAR
Public Enum AreaPermisos
    'PuedoNADA = 0
    Altas = 1
    Bajas = 2
    Ver = 4
    Modis = 8
    'PuedoTodo = 15
End Enum

Public Function Puedo(queCosa As AreaPermisos, Optional quearea As AreaSeguridad = -1) As Boolean
    'Dim a As AreaSeguridad
    Dim usuario As Long, tipo As Long, permisos As Long
    Dim tempo
    
    'caso especial puedo todo
    tipo = UsuarioSistema!TIPOUSUARIO
    If tipo = 1 Then
        Puedo = True
        Exit Function
    End If
      
    'establezco el area de seguridad, si es que la encuentro
    If quearea = -1 Then
        quearea = areaForm()
    End If
    If quearea < 1 Then
        Puedo = False
        Exit Function
    End If
    
    'usuario = UsuarioActual()
    permisos = s2n(obtenerDeSQL("select permisos from UsuarioArea where TipoUsuario = " & tipo & " and AreaSeguridad = " & quearea))
    
    'Mascara logica, es AND binaria
    'Puedo = (permisos And queCosa)
    If (permisos And queCosa) Then Puedo = True
    
    'MsgBox x
End Function


Private Function areaForm() As AreaSeguridad
    On Error GoTo ufa
    areaForm = Screen.ActiveForm.Area()
fin:
    Exit Function
ufa:
    che "prg: area() form no definida"
    areaForm = -1
    Resume fin
End Function

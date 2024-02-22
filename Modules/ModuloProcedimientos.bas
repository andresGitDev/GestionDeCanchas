Attribute VB_Name = "ModuloProcedimientos"
Public cmd As New ADODB.Command
Public RStraer As New ADODB.Recordset
Private El_titulo As String

'Declaraciónes apis

 ' Lista las ventanas
 Declare Function EnumWindows Lib "user32" ( _
                  ByVal wndenmprc As Long, _
                  ByVal lParam As Long) As Long

 'Recupera el texto de la misma
 Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" ( _
                  ByVal hWnd As Long, _
                  ByVal lpString As String, _
                  ByVal cch As Long) As Long

 'Para finalizar dicha ventana
 Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
                  ByVal hWnd As Long, _
                  ByVal wMsg As Long, _
                  ByVal wParam As Long, _
                  lParam As Any) As Long



Public Function ExisteArch(sArchivo As String) As Integer
    ExisteArch = Len(Dir$(sArchivo))
End Function

' Recibe el título parcial o igual de las ventanas a cerrar
Public Sub Cerrar_ventana(El_Caption As String)
    El_titulo = El_Caption
    Call EnumWindows(AddressOf EnumCallback, 0)
End Sub

' Función para recorrer las ventanas abiertas
Public Function EnumCallback(ByVal A_hwnd As Long, _
                ByVal PARAM As Long) As Long

Dim buffer As String * 256
Dim Titulo_Win As String
Dim Size_buffer As Long

'Retorna la cantidad de caracteres del título de la ventana actual
Size_buffer = GetWindowText(A_hwnd, buffer, Len(buffer))
'Elimina los espacios nulos de la cadena
Titulo_Win = Left$(buffer, Size_buffer)

    'si se encuentra la cadena en el caption de la ventana se cierra
If InStr(Titulo_Win, El_titulo) <> 0 Then
    ' Finaliza la ventana
    SendMessage A_hwnd, WM_SYSCOMMAND, SC_CLOSE, ByVal 0&
End If
    ' Esto continúa enumerando las siguientes ventanas de windows
    EnumCallback = 1
End Function


Public Sub Traer(ByVal store As String, ByVal Orden As String, ByVal grupD As String, ByVal grupH As String, ByVal codD As String, ByVal codH As String)
    Set cmd = New ADODB.Command
    Set RStraer = New ADODB.Recordset
    cmd.ActiveConnection = DataEnvironment1.Sistema
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = store
    cmd.Parameters.Refresh
    cmd.Parameters(1).Value = Orden
    cmd.Parameters(2).Value = Trim(grupD)
    cmd.Parameters(3).Value = Trim(grupH)
    cmd.Parameters(4).Value = codD
    cmd.Parameters(5).Value = codH
    RStraer.CursorLocation = adUseClient
    RStraer.Open cmd
End Sub
Public Sub TraerProveedores(ByVal store As String, ByVal Orden As String, ByVal desde As String, ByVal hasta As String)
    Set cmd = New ADODB.Command
    Set RStraer = New ADODB.Recordset
    cmd.ActiveConnection = DataEnvironment1.Sistema
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = store
    cmd.Parameters.Refresh
    cmd.Parameters(1).Value = Orden
    cmd.Parameters(2).Value = desde
    cmd.Parameters(3).Value = hasta
    RStraer.CursorLocation = adUseClient
    RStraer.Open cmd
End Sub

Public Sub CheqTer()
    Set cmd = New ADODB.Command
    Set RStraer = New ADODB.Recordset
    cmd.ActiveConnection = DataEnvironment1.Sistema
    cmd.CommandType = adCmdText
    cmd.CommandText = "select * from GastoBankTemp order by Codigo"
    
    RStraer.CursorLocation = adUseClient
    RStraer.Open cmd
End Sub


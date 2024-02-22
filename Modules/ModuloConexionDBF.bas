Attribute VB_Name = "ModuloConexionDBF"
Option Explicit





Public cnxDBF As New ADODB.Connection


Sub cnxAbrirDBF(carpeta As String)
    'fuerza la reapertura,
    cnxCerrarDBF

    If carpeta > "" Then
        cnxDBF.ConnectionString = " DSN=dBASE Files;DBQ= " & carpeta & "; " & _
            " DefaultDir=" & carpeta & ";DriverId=533;MaxBufferSize=2048;PageTimeout=5;"
        cnxDBF.Open
    Else
        che "No encuentro la carpeta: " & carpeta
    End If
End Sub


Public Sub cnxCerrarDBF()
    If cnxDBF.State = adStateOpen Then cnxDBF.Close
End Sub

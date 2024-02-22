Attribute VB_Name = "MuduloBackupSql"
Option Explicit


Public Function BackupSql(NombreBase As String, carpeta As String, Optional Zipeado As Boolean = False, Optional AgregaDiaAlNombre As Boolean = False)
    'hace backup  de la base (total NO incremental, y minimo log)
    'NO 'permite zipearla
    'NO 'permite agregarle al nombre la fecha NOMBREXXX_2006-01-13
    
'    if on_error_habilitado then  ufache
    Dim Archivo As String
    
    'faltaalgo?
    If NombreBase = "" Or carpeta = "" Then
        che "err: sin definiciones para backup"
    End If
    
    
    'chequear carpeta / crearla (si no sql no lo hace)
    ' if   'que pasa con el \ final ?
    
    
    'Nombre Archivo, agregar fecha al nombre
    Archivo = NombreBase & ".Sqk"
    
    
    'chequear si ya existe y preguntar si sobreescribo
        'guardar datos de archivo para confirmar sobreescritura
    
    
    
    ' graba
    If Right(carpeta, 1) <> "\" Then carpeta = carpeta & "\"
    DataEnvironment1.Sistema.Execute "BACKUP DATABASE [" & NombreBase & "] TO DISK = '" & carpeta & Archivo & "' WITH  NOINIT ,  NOUNLOAD ,  NAME = '" & NombreBase & "',  NOSKIP ,  STATS = 10,  NOFORMAT"
    
    ' verifica que exista, (y la fecha? )
    
    '
    MsgBox "Backup en " & vbCrLf & carpeta & Archivo
    
    'error
    
End Function


Attribute VB_Name = "ModuloOdioStored"
Option Explicit

Public Enum parametroABM
    abmAlta = 1
    abmbaja = 2
    abmmodi = 3
End Enum
Private Function ssSetBajaWhere() As String
    ssSetBajaWhere = " set fecha_baja = " & ssFecha(Date) & ", usuario_baja = " & UsuarioActual() & ", activo = 0  where "
End Function
Private Function ssValuesALTA() As String
    ssValuesALTA = "  " & ssFecha(Date) & ",  " & UsuarioActual() & ",  1 "
End Function
Private Function ssInsertALTA() As String
    ssInsertALTA = "  fecha_alta, usuario_alta, activo "
End Function



Public Function abmTransfMoviCaja(Ope As parametroABM, caja As Long, movimiento As Long, tipo As String, Ing_egr As String, importe As Double, concepto As String, fecha As Date, cuenta As String, MovBanco As Long, iddoc As Long)
    Dim s As String
    
    If iddoc = 0 Then Err.Raise 55001, , " identificador no correcto iddoc"
    
    Select Case Ope
    Case abmAlta
        s = " INSERT INTO MOVICAJA ( " & _
                        " CAJA, MOVIMIENTO, TIPO, ING_EGR, IMPORTE, CONCEPTO, " & _
                        "  FECHA, CUENTA, MOVBANCO, iddoc, " & _
                        ssInsertALTA() & _
                        " ) VALUES ( " & _
                        caja & ", " & movimiento & ", '" & tipo & "',  '" & Ing_egr & "', " & ssNum(importe) & ",  '" & concepto & "' , " & _
                        ssFecha(fecha) & ", '" & cuenta & "',  " & MovBanco & ",  " & iddoc & ", " & _
                        ssValuesALTA() & " ) "
    'Case abmmodi
    Case abmbaja
        s = "update movicaja " & ssSetBajaWhere & " iddoc = " & iddoc
    Case Else
        Err.Raise 55000, , "parametro no definido en transfMoviCaja"
    End Select
    
    DataEnvironment1.Sistema.Execute s
    
        'Alter PROCEDURE dbo.TRANSFMOVICAJA @OPE VARCHAR(1), @CAJA INT, @MOVIMIENTO INT, @TIPO VARCHAR(1), @ING_EGR VARCHAR(1)
        '               , @IMPORTE FLOAT, @CONCEPTO VARCHAR(30), @FECHA DATETIME, @CUENTA VARCHAR(9), @MOVBANCO INT,
        '               @FECHA_ALTA DATETIME, @USUARIO_ALTA INT, @FECHA_BAJA DATETIME, @USUARIO_BAJA INT
        'AS
        'IF @OPE = 'A'
        '    BEGIN
        '        INSERT INTO MOVICAJA (CAJA, MOVIMIENTO, TIPO, ING_EGR, IMPORTE, CONCEPTO, FECHA, CUENTA, MOVBANCO, FECHA_ALTA, USUARIO_ALTA, ACTIVO) VALUES (@CAJA, @MOVIMIENTO, @TIPO, @ING_EGR, @IMPORTE, @CONCEPTO, @FECHA, @CUENTA, @MOVBANCO, @FECHA_ALTA, @USUARIO_ALTA, 1)
        '    End
        'IF @OPE = 'A'
        '    BEGIN
        '        UPDATE MOVICAJA SET FECHA_BAJA = @FECHA_BAJA, USUARIO_BAJA = @USUARIO_BAJA, ACTIVO = 0 WHERE MOVIMIENTO = @MOVIMIENTO
        '    End


End Function

Public Function abmTransfMoviBanc(Ope As parametroABM, cuentabanc As Long, operacion As String, descripcion As String, fecha As Date, documento As String, importe As Double, MovBanco As Long, iddoc As Long)
    Dim s As String
    
    If iddoc = 0 Then Err.Raise 55001, , " identificador no correcto iddoc"
    
    Select Case Ope
    Case abmAlta
    s = " INSERT INTO MOVIBANC (  " & _
                    " CUENTA, OPERACION, DESCRIPCION, FECHA, DOCUMENTO, IMPORTE, MOVBANCO, iddoc, " & _
                    ssInsertALTA() & _
                    " ) VALUES ( " & _
                     cuentabanc & ",  '" & operacion & "', '" & descripcion & "' , " & ssFecha(fecha) & ", '" & documento & "' , " & ssNum(importe) & ",  " & MovBanco & ",  " & iddoc & ", " & _
                    ssValuesALTA() & " ) "
    'Case abmmodi
    Case abmbaja
        s = "update MOVIBANC " & ssSetBajaWhere & " iddoc = " & iddoc ' set activo = 0, fecha_baja = " & ssFecha(Date)
    Case Else
        Err.Raise 55000, , "parametro no definido en transfMoviCaja"
    End Select

    DataEnvironment1.Sistema.Execute s
                
        'Alter PROCEDURE dbo.TRANSFMOVIBANC @OPE VARCHAR(1), @CUENTA INT, @OPERACION VARCHAR(1), @DESCRIPCION VARCHAR(30), @FECHA DATETIME, @DOCUMENTO VARCHAR(1), @IMPORTE FLOAT, @MOVBANCO INT, @FECHA_ALTA DATETIME, @USUARIO_ALTA INT,  @FECHA_BAJA DATETIME, @USUARIO_BAJA INT
        'AS
        'IF @OPE = 'A'
        '    BEGIN
        '        INSERT INTO MOVIBANC (CUENTA, OPERACION, DESCRIPCION, FECHA, DOCUMENTO, IMPORTE, MOVBANCO, FECHA_ALTA, USUARIO_ALTA, ACTIVO) VALUES (@CUENTA, @OPERACION, @DESCRIPCION, @FECHA, @DOCUMENTO, @IMPORTE, @MOVBANCO, @FECHA_ALTA, @USUARIO_ALTA, 1)
        '    End
        'IF @OPE = 'B'
        '    BEGIN
        '        UPDATE MOVIBANC SET FECHA_BAJA = @FECHA_BAJA, USUARIO_BAJA = @USUARIO_BAJA,  ACTIVO = 0 WHERE MOVBANCO = @MOVBANCO
        '    End
End Function

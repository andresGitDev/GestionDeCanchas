Attribute VB_Name = "ModuloImpresoras"
Option Explicit

Public Function SetImpresora(cual As String) As Boolean
    Dim x As Printer
    For Each x In Printers
        If x.DeviceName = cual Then
            Debug.Print x.Port, x.DeviceName
            Set Printer = x
            SetImpresora = True
            Exit For
        End If
    Next
End Function


Public Function setLpt1() As Boolean
    Dim x As Printer
    For Each x In Printers
        If LCase(x.Port) = "lpt1:" Then
            Set Printer = x
            setLpt1 = True
            Exit For
        End If
    Next
End Function

Public Function setLpt2() As Boolean
    Dim x As Printer
    For Each x In Printers
        If LCase(x.Port) = "lpt2:" Then
            Set Printer = x
            setLpt2 = True
            Exit For
        End If
    Next
End Function

Public Function setPrintPort(cual)
    Dim x As Printer
    For Each x In Printers
        If LCase(x.Port) = cual Then
            Set Printer = x
            setPrintPort = True
            Exit For
        End If
    Next
End Function

Public Function verImprNombre(Optional CambiarPor As String) As String
    Dim x As Printer
    verImprNombre = Printer.DeviceName
    
    If CambiarPor = "" Then Exit Function
    
    For Each x In Printers
        If LCase(x.DeviceName) = Trim(LCase(CambiarPor)) Then
            Set Printer = x
            'setPrintPort = True
            Exit For
        End If
    Next
    Set x = Nothing
End Function

Public Function GetPrinterEnPort(cual As String) As Printer
    Dim x As Printer
   
    For Each x In Printers
        If LCase(x.Port) = Trim(LCase(cual)) Then
            Set GetPrinterEnPort = x
            Exit For
        End If
    Next
    Set x = Nothing
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Asiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Item"
Option Explicit '1/2/5
' Pensado para grabar asientos a tablas Asiento y Mayor

'Se llena cabecera con .Nuevo(), e items mayor con .AgregarItem()

' se lo llena con datos sin verificar, puede estar vacio,
' y cuando se le ordena grabar:
' si no tiene items o si estan todos vacios NO GRABA NADA, y mantiene silencio
' si no cierra NO GRABA NADA, y da aviso al usuario, sin dar error

' pendiente:
' si se le manda dato negativo, q lo pase positivo a la otra columna


'TEMPORARIO -- SACAR ----             ----  ***********************  ----
'   para debug, permite agregar items vacios, pero puede evitarlo
Private Const DESBICHO = False
'  NO DEBE preguntar si graba,
'    si diferencia,  no graba y punto.
'    el msgbox en una transaccion caga el sistema

'Private mIdDoc As Long
Private mvarid As Long
Private mvarFecha As Date
Private mvarEjercicio As Long
Private mvarConcepto As String
Private mItem() As ItemAsiento
Private mCantItems As Long
Private mvarOrigen As String
Private mvarCantitems As Long
Private mvarItem As Object
'Private mIdEjercicio As Integer

Public Enum ErrAsiento
    ErrAsientoGrabSinItems = 50000
    ErrAsientoGrabDiferencia = 50001
    ErrAsientoGrabUnknown = 50003
    ErrAsientoItemSinCuenta = 50010
End Enum


Private Type ItemAsiento
    Cuenta As String
    Fecha As Date
    concepto As String
    Debe As Double
    haber As Double
    DocOrigen As String
End Type

'Public Property Get idDoc() As Long
'    idDoc = mIdDoc
'End Property

Public Property Get ItemCuenta(Index As Long) As String
    ItemCuenta = mItem(Index).Cuenta
End Property
Public Property Get ItemFecha(Index As Long) As Date
    ItemFecha = mItem(Index).Fecha
End Property
Public Property Get ItemConcepto(Index As Long) As String
    ItemConcepto = mItem(Index).concepto
End Property
Public Property Get ItemDebe(Index As Long) As Double
    ItemDebe = s2n(mItem(Index).Debe)
End Property
Public Property Get ItemHaber(Index As Long) As Double
    ItemHaber = s2n(mItem(Index).haber)
End Property
Public Property Get ItemDocOrigen(Index As Long) As String
    ItemDocOrigen = mItem(Index).DocOrigen
End Property


Public Property Get CantItems() As Long
    CantItems = mvarCantitems
End Property

Public Property Let Origen(ByVal vData As String)
    mvarOrigen = vData
End Property
Public Property Get Origen() As String
    Origen = mvarOrigen
End Property

Public Property Let Fecha(ByVal vData As Date)
    mvarFecha = vData
End Property
Public Property Get Fecha() As Date
    Fecha = mvarFecha
End Property

Public Property Get Diferencia() As Double
    Diferencia = s2n(Me.TotalDebe - Me.TotalHaber)
End Property

Public Property Get TotalDebe() As Double
    Dim i As Long, d As Double
    d = 0
    For i = 0 To mvarCantitems - 1
        d = d + mItem(i).Debe
    Next i
    TotalDebe = s2n(d)
End Property

Public Property Get TotalHaber() As Double
    Dim i As Long, h As Double
    h = 0
    For i = 0 To mvarCantitems - 1
        h = h + mItem(i).haber
    Next i
    TotalHaber = s2n(h)
End Property

'Public Function Nuevo(ByVal Ejercicio As Integer, ByVal concepto As String, ByVal Fecha As Date, ByVal Origen As String) As Long
Public Function nuevo(ByVal concepto As String, ByVal Fecha As Date, ByVal Origen As String) As Long
    mvarCantitems = 0
    mvarEjercicio = leerEjercicioId() 'Ejercicio
    mvarConcepto = concepto
    mvarFecha = Fecha
    mvarOrigen = Origen
    ReDim mItem(0)
'    mIdDoc = -1
End Function

Public Property Let concepto(ByVal vData As String)
    mvarConcepto = vData
End Property
Public Property Get concepto() As String
    concepto = mvarConcepto
End Property

Public Function Grabar(iddoc As Long, Optional verDiferecia As Boolean = True, Optional EjerM As Long) As Long
    Dim ID As Long, i As Long
    
    Grabar = 0
    If trIMPUTACION = 1 Then
        'esto significa que es imputacion y por lo tanto no trae la forma de pago
    Else
        If mvarCantitems = 0 Then
            che "No hay items a grabar."
            Err.Raise ErrAsientoGrabSinItems
            Exit Function
        End If
    End If

    If Me.Diferencia <> 0 Then
        If verDiferecia Then
            If Not DESBICHO Then         '''''''''''''''''DEBUG ''''''''''''''''''''
                MsgBox "No se puede generar el asiento, diferencia = " & Diferencia, vbCritical
                'Err.Raise ErrAsientoGrabDiferencia
                Exit Function
            Else
                If Not confirma("No se puede grabar, diferencia " & Diferencia & vbCrLf & " Grabo asiento igualmente ?" & vbCrLf & "ESTE MSG NO DEBE SALIR EN VERSION FINAL") Then
                    Err.Raise ErrAsientoGrabDiferencia
                    Exit Function
                End If
            End If
        Else
            If MsgBox("EL ASIENTO SE GRABARA CON DIFERENCIA DE : " & Me.Diferencia, vbInformation + vbYesNo) = vbNo Then
                    Exit Function
            End If
        End If
    End If
    
    ordenar
    
    If EjerM > 0 Then
        ID = GrabarAsiento(iddoc, EjerM)
    Else
        ID = GrabarAsiento(iddoc)
    End If
    If ID = 0 Then
        'Err.Raise ErrAsientoGrabUnknown
        Exit Function
    End If
'    'grabo los debe
'    For i = 0 To mvarCantitems - 1
'        If mItem(i).Debe > 0 Then
'            GrabarMayor CLng(ID), CStr(mItem(i).Cuenta), mItem(i).Debe, mItem(i).haber, mItem(i).DocOrigen  'deberia ser 0, pero lo pongo para q salte cualquier err q hubiere
'        End If
'    Next i
'    'grabo los haber
'    For i = 0 To mvarCantitems - 1
'        If mItem(i).haber > 0 Then
'            GrabarMayor CLng(ID), CStr(mItem(i).Cuenta), mItem(i).Debe, mItem(i).haber, mItem(i).DocOrigen
'        End If
'    Next i

    'grabo los debe
    For i = 0 To mvarCantitems - 1
        If mItem(i).Debe > 0 Then
            If mItem(i).haber > 0 Then
                MsgBox "Error en la cuenta " & mItem(i).Cuenta & ", Debe/Haber - No se permite el ingreso de datos en ambos campos para la misma cuenta.", vbCritical
                GoTo ErrAsiento
            ElseIf mItem(i).haber = 0 Then
                mItem(i).Debe = mItem(i).Debe
                mItem(i).haber = mItem(i).haber
            End If
        
            GrabarMayor CLng(ID), CStr(mItem(i).Cuenta), mItem(i).Debe, mItem(i).haber, mItem(i).DocOrigen   'deberia ser 0, pero lo pongo para q salte cualquier err q hubiere
EsHaber:
        End If
    Next i
    'grabo los haber
    For i = 0 To mvarCantitems - 1
        If mItem(i).haber > 0 Then
            If mItem(i).Debe > 0 Then
                MsgBox "Error en la cuenta " & mItem(i).Cuenta & ", Debe/Haber - No se permite el ingreso de datos en ambos campos para la misma cuenta.", vbCritical
                GoTo ErrAsiento
            ElseIf mItem(i).Debe = 0 Then
                mItem(i).Debe = mItem(i).Debe
                mItem(i).haber = mItem(i).haber
            End If
            GrabarMayor CLng(ID), CStr(mItem(i).Cuenta), mItem(i).Debe, mItem(i).haber, mItem(i).DocOrigen
        End If
    Next i


    
    Grabar = ID
Exit Function
ErrAsiento:
End Function


Public Function AgregarItem(ByVal Cuenta As String, ByVal Debe As Double, ByVal haber As Double, Optional docori As String) As Long
    
'    If IsNull(cuenta) Then ' or trim(cuenta) = "" then
'        Err.Raise 222 2 2
    If Trim(Cuenta) = "" Then
        Err.Raise ErrAsientoItemSinCuenta   ' NO largo error, asi doy flexibilidad, igual saltara error si al grabar no cierra o esta vacio.
        Exit Function
    End If
    
    If NotExisteCuenta(Cuenta) Then
        'Err.Raise ErrAsientoItemSinCuenta
        MsgBox "No existe cuenta contable Nº " & Cuenta, vbCritical
        Exit Function
    End If
    

    'No agrego si vacio
    If Not DESBICHO Then
        If haber = 0 And Debe = 0 Then Exit Function      ' HABILITAR en vers final
    End If
    'Cambio de columna  si es negativo
    If Debe < 0 Then haber = haber - Debe: Debe = 0
    If haber < 0 Then Debe = Debe - haber: haber = 0
    
    mvarCantitems = mvarCantitems + 1
    ReDim Preserve mItem(mvarCantitems)
    
    With mItem(mvarCantitems - 1)
        .Cuenta = Cuenta
        .Debe = Debe
        .haber = haber
        .Fecha = Fecha
        .concepto = Left("________________", IIf(Debe > 0, 3, 10)) & CuentaDescripcion(Cuenta)  'sSinNull(obtenerDeSQL("select descripcion from cuentas where codigo = " & cuenta))
        .DocOrigen = docori
    End With
    AgregarItem = mvarCantitems - 1
End Function

Public Function AcumularItem(ByVal Cuenta As String, ByVal Debe As Double, ByVal haber As Double, Optional docori As String) As Long
'DocOri solo sirve la primera vez, si ya existe cuenta, lo ignora
    Dim i As Long, t As Double
    
    'If IsNull(cuenta) Then Exit Function
    If Trim(Cuenta) = "" Then
        Err.Raise ErrAsientoItemSinCuenta
    End If
    
    If NotExisteCuenta(Cuenta) Then
        MsgBox "No existe cuenta contable", vbCritical
        'Err.Raise ErrAsientoItemSinCuenta
        Exit Function
    End If
    
    
    'No agrego si vacio
    If Not DESBICHO Then
        If haber = 0 And Debe = 0 Then Exit Function      ' HABILITAR en vers final
    End If
    
'    'Cambio de columna  si es negativo
'    If debe < 0 Then haber = haber - debe: debe = 0
'    If haber < 0 Then debe = debe - haber: haber = 0
    
    For i = 0 To mvarCantitems - 1
        With mItem(i)
            If Cuenta = .Cuenta Then
                .Debe = .Debe + Debe
                .haber = .haber + haber
                If .concepto = "" Then .concepto = sSinNull(docori)
                
                ' ------- No puedo creer q haya tenido q poener este codigo aca... ----------
                t = .Debe - .haber
                If t > 0 Then
                    .Debe = t
                    .haber = 0
                Else
                    .Debe = 0
                    .haber = -t
                End If
                ' ---------------------------------------------------------------------------

                Exit Function
            End If
        End With
    Next i
    AcumularItem = AgregarItem(Cuenta, Debe, haber, docori)
End Function

Private Function NotExisteCuenta(cual As String) As Boolean
    Dim x As Long
    x = s2n(obtenerDeSQL("select id from cuentas where cuenta = '" & cual & "' "))
    NotExisteCuenta = (x = 0)
End Function

Public Property Get ID() As Long
Attribute ID.VB_UserMemId = 0
    ID = mvarid
End Property

Private Function GrabarAsiento(iddoc As Long, Optional DesdeM As Long) As Long

    Dim rs As New ADODB.Recordset, NuevoNroAsiento As Long, idEjer As Long
    GrabarAsiento = 0
    
    If DesdeM > 0 Then
        idEjer = DesdeM
    Else
        idEjer = BuscoEjercicio()
    End If
       

    If idEjer = 0 Then
        che "No se pudo grabar asiento"
        Exit Function
    End If
    NuevoNroAsiento = nuevoCodigo("Asientos", "NroAsiento", "Ejercicio = '" & idEjer & "' and activo = 1 ")
    With rs
        .Open "select top 1 * from asientos", DataEnvironment1.Sistema, adOpenKeyset, adLockOptimistic
            .AddNew
                !iddoc = iddoc
                !ejercicio = idEjer                '!Ejercicio = mvarEjercicio
                !NroAsiento = NuevoNroAsiento
                !Fecha = mvarFecha
                !Origen = Origen
                !concepto = concepto
            .Update
            GrabarAsiento = !idAsiento
        .Close
    End With
fin:
    Set rs = Nothing
End Function

Private Function GrabarMayor(idAsiento As Long, Cuenta As String, Debe As Double, haber As Double, docori As String)
    If Debe = 0 And haber = 0 Then Exit Function
    'DataEnvironment1.Sistema.Execute "insert into mayor  (idAsiento, Cuenta, Debe, Haber) values (" & idAsiento & ", '" & cuenta & "' , " & x2s(debe) & ", " & x2s(haber) & ") "
    DataEnvironment1.Sistema.Execute "insert into mayor  (idAsiento, Cuenta, Debe, Haber, comprobante) values (" & idAsiento & ", '" & Cuenta & "' , " & x2s(Debe) & ", " & x2s(haber) & ", '" & docori & "' ) "
End Function

Private Function BuscoEjercicio() As Long
    Dim tempo
    tempo = obtenerDeSQL("Select idEjercicio, fechainicio, fechafin, cerrado from ejercicio where activo = 1")
    If IsEmpty(tempo) Then 'activo
        ufa "No se encontro ejercicio activo", "grabaAsiento"
        Exit Function
    ElseIf tempo(3) = True Then ' cerrado
        ufa "ejercicio activo figura como cerrado", "grabaAsiento"
        Exit Function
    ElseIf CDate(mvarFecha) < CDate(tempo(1)) Or CDate(mvarFecha) > CDate(tempo(2)) Then
        ufa "fecha asiento fuera de rango de la fecha de ejercicio", "grabaAsiento"
        Exit Function
    End If
    BuscoEjercicio = s2n(tempo(0), 0)
End Function

'Public Sub ordenar()
'    Dim ii As Long, i As Long, j As Long
'    Dim ci As Long, cj As Long
'
'    For i = 0 To mvarCantitems - 1
'        ii = i
'        ci = mItem(ii).Cuenta
'        If mItem(ii).haber > 0 Then ci = ci + 900000000
'
'        For j = i To mvarCantitems - 1
'            cj = mItem(j).Cuenta
'            If mItem(j).haber > 0 Then cj = cj + 900000000
'            If cj < ci Then
'                ii = j
'                ci = cj
'            End If
'        Next j
'        swapeo ii, i
'    Next i
'End Sub

Public Sub ordenar()
    Dim ii As Long, i As Long, j As Long
    Dim ci As String, cj As String
    
    For i = 0 To mvarCantitems - 1
        ii = i
        ci = mItem(ii).Cuenta
        If mItem(ii).haber > 0 Then ci = "Z" & ci
        
        For j = i To mvarCantitems - 1
            cj = mItem(j).Cuenta
            If mItem(j).haber > 0 Then cj = "Z" & cj
            If cj < ci Then
                ii = j
                ci = cj
            End If
        Next j
        swapeo ii, i
    Next i
End Sub

Private Sub swapeo(i1, i2)
    Dim tempo As ItemAsiento
    tempo = mItem(i1)
    mItem(i1) = mItem(i2)
    mItem(i2) = tempo
End Sub

'25/4/4     cuenta null sale
'6/5/5      ordena x cuenta y por d/h
'23/1/6     nuevo numeracion por ejercicio

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AjusteInflacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public Function buscar(bMes As Integer, bAnio As Integer, Optional bAnual As Boolean = False) As Double
On Error GoTo err_buscar
Dim bDia As Integer, bPeriodo As Date
Dim bIndice As Double, bCadena As String

If bAnual Then
    bDia = 31
    bMes = 12
Else
    bDia = 1
End If

bPeriodo = CDate(bDia & "/" & bMes & "/" & bAnio)

bIndice = nSinNull(obtenerDeSQL("select indice from ajuste_indice where periodo=" & ssFecha(bPeriodo)))

If bIndice > 0 Then
    buscar = bIndice
Else
    buscar = 0
End If

Exit Function
err_buscar:
MsgBox Err.Description & " " & Err.Number
End Function

Public Function Guardar(gMes As Integer, gAnio As Integer, gIndice As Double, Optional gAnual As Boolean = False) As Boolean
On Error GoTo err_guardar
Dim gDia As Integer, gPeriodo As Date
Dim gID As Long, gCadena As String

If gAnual Then
    gDia = 31
    gMes = 12
Else
    gDia = 1
End If

gPeriodo = CDate(gDia & "/" & gMes & "/" & gAnio)

gID = nSinNull(obtenerDeSQL("select id from ajuste_indice where periodo=" & ssFecha(gPeriodo)))

If gID > 0 Then
    gCadena = "update ajuste_indice set indice=" & x2s(gIndice) & " where id=" & gID
Else
    gCadena = "insert into ajuste_indice (periodo,indice) values (" & ssFecha(gPeriodo) & "," & x2s(gIndice) & ")"
End If

DataEnvironment1.Sistema.Execute gCadena
Guardar = True
Exit Function
err_guardar:
Guardar = False
MsgBox Err.Description & " " & Err.Number
End Function


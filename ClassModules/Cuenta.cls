VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type item
    Cuenta As Long
    monto As Double
End Type

Private mvarid As Long
Private mItem() As item
Private mvarCantitems As Long
'public function Nuevo("
'

Public Property Get ItemCuenta(Index As Long) As Long
    ItemCuenta = mItem(Index).Cuenta
End Property
Public Property Get ItemMonto(Index As Long) As Double
    ItemMonto = s2n(mItem(Index).monto)
End Property

Public Property Get CantItems() As Long
    CantItems = mvarCantitems
End Property

Public Property Get TotalMonto() As Double
    Dim i As Long, d As Double
    For i = 0 To mvarCantitems - 1
        d = d + mItem(i).monto
    Next i
    TotalMonto = s2n(d)
End Property

Public Function nuevo() As Long
    mvarCantitems = 0
    ReDim mItem(0)
End Function

Private Function AgregarItem(ByVal Cuenta As Long, ByVal monto As Double) As Long
    mvarCantitems = mvarCantitems + 1
    ReDim Preserve mItem(mvarCantitems)

    With mItem(mvarCantitems - 1)
        .Cuenta = Cuenta
        .monto = monto
    End With
    AgregarItem = mvarCantitems - 1
End Function

Public Function AcumularItem(ByVal Cuenta As Long, ByVal monto As Double) As Long
    Dim i As Long
    'No agrego si vacio
    If monto = 0 Then Exit Function       ' HABILITAR en vers final
    
    For i = 0 To mvarCantitems - 1
        With mItem(i)
            If Cuenta = .Cuenta Then
                .monto = .monto + monto
                Exit Function
            End If
        End With
    Next i
    AcumularItem = AgregarItem(Cuenta, monto)
End Function

'
'Public Property Get id() As Long
'    id = mvarid
'End Property
'

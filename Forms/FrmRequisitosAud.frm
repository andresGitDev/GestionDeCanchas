VERSION 5.00
Begin VB.Form FrmRequisitosAud 
   Caption         =   "Requisistos Auditoria"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10095
   Icon            =   "FrmRequisitosAud.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   10095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdEjecutar 
      Caption         =   "Ejecutar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   7665
      TabIndex        =   2
      Top             =   345
      Width           =   1215
   End
   Begin VB.ListBox LstAgregados 
      Height          =   2205
      Left            =   105
      TabIndex        =   4
      Top             =   975
      Width           =   9885
   End
   Begin VB.ListBox LstSeleccion 
      Height          =   2205
      Left            =   105
      TabIndex        =   3
      Top             =   3225
      Width           =   9885
   End
   Begin Gestion.ucBotonera UBotonera 
      Height          =   1515
      Left            =   1155
      TabIndex        =   1
      Top             =   5520
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   2672
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   ""
   End
   Begin Gestion.ucCoDe UProd 
      Height          =   315
      Left            =   105
      TabIndex        =   0
      Top             =   360
      Width           =   7515
      _ExtentX        =   13123
      _ExtentY        =   556
      CodigoWidth     =   1000
   End
End
Attribute VB_Name = "FrmRequisitosAud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql, TipoOp As String
Dim CuantosCargados As Long
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset

Private Sub CmdEjecutar_Click()
Select Case TipoOp
Dim Confir As String
Case "Nuevo"
   If uProd.codigo <> "" Then
      sql = "SELECT * FROM ReqAudit ORDER BY codigo ASC"
      rs.Open sql, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
      Do While Not rs.EOF
         LstSeleccion.AddItem rs!DESCRIPCION
         LstSeleccion.ItemData(LstSeleccion.NewIndex) = rs!codigo
      rs.MoveNext
      Loop
      rs.Close
   Else
      MsgBox "Debe seleccionar primero un producto", vbExclamation, "Aviso"
      Exit Sub
   End If
Case "Buscar"
   CuantosCargados = 0
   Confir = "N"
   sql = "SELECT * FROM ReqAuditprod WHERE codigoprod = '" & uProd.codigo & "' ORDER BY codigo ASC"
      rs.Open sql, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
      Do While Not rs.EOF
         sql = "SELECT * FROM ReqAudit WHERE codigo = " & rs!codigo & ""
         rs2.Open sql, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
         LstAgregados.AddItem rs2!DESCRIPCION
         LstAgregados.ItemData(LstAgregados.NewIndex) = rs!codigo
         CuantosCargados = CuantosCargados + 1
         rs2.Close
         Confir = "S"
      rs.MoveNext
      Loop
      rs.Close
      If Confir = "S" Then
         CargoDiferencia
      End If
End Select
End Sub
Private Sub CargoDiferencia()
Dim X, zz, a As Long
Dim oP As String
  sql = "SELECT * FROM ReqAudit"
  rs.Open sql, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
  Do While Not rs.EOF
   For X = 0 To LstAgregados.ListCount - 1
    If LstAgregados.ItemData(X) = CDbl(rs!codigo) Then
       oP = "S"
       rs.MoveNext
    Else
       oP = "N"
    End If
   Next
   If oP = "N" Then
      LstSeleccion.AddItem rs!DESCRIPCION
      LstSeleccion.ItemData(LstSeleccion.NewIndex) = rs!codigo
      X = 0
      oP = ""
      rs.MoveNext
   End If

  Loop
  rs.Close
End Sub

Private Sub Form_Load()
UBotonera.init True, True, True, False, True
uProd.enabled = False
uProd.ini "Select Descripcion from producto where activo = 1 and codigo = '###' ", "Select codigo as [ Producto ], descripcion as [ Nombre                                     ] from producto where activo = 1 order by codigo", True
UBotonera.MsgConfirmaEliminar = "Esta seguro de eliminar todos los Items"
End Sub

Private Sub LstAgregados_DblClick()
If TipoOp = "Modif" Then
   If LstAgregados.ListIndex < CuantosCargados Then
      If MsgBox("Desea Eliminar un Item", vbQuestion + vbYesNo, "Confirma operacion") = vbYes Then
         DataEnvironment1.dbo_INGRESOAUDITPROD "B", LstAgregados.ItemData(LstAgregados.ListIndex), uProd.codigo
         LstSeleccion.AddItem LstAgregados.Text
         LstSeleccion.ItemData(LstSeleccion.NewIndex) = LstAgregados.ItemData(LstAgregados.ListIndex)
         LstAgregados.RemoveItem (LstAgregados.ListIndex)
      End If
   Else:
         LstSeleccion.AddItem LstAgregados.Text
         LstSeleccion.ItemData(LstSeleccion.NewIndex) = LstAgregados.ItemData(LstAgregados.ListIndex)
         LstAgregados.RemoveItem (LstAgregados.ListIndex)
   End If
Else
   LstSeleccion.AddItem LstAgregados.Text
   LstSeleccion.ItemData(LstSeleccion.NewIndex) = LstAgregados.ItemData(LstAgregados.ListIndex)
   LstAgregados.RemoveItem (LstAgregados.ListIndex)
End If
End Sub

Private Sub LstSeleccion_DblClick()
   LstAgregados.AddItem LstSeleccion.Text
   LstAgregados.ItemData(LstAgregados.NewIndex) = LstSeleccion.ItemData(LstSeleccion.ListIndex)
   LstSeleccion.RemoveItem (LstSeleccion.ListIndex)
End Sub

Private Sub UBotonera_Aceptar()
Dim X, cc As Long
Select Case TipoOp
Case "Nuevo"
   If LstAgregados.ListCount > 0 Then
      For X = 0 To LstAgregados.ListCount - 1
         DataEnvironment1.dbo_INGRESOAUDITPROD "A", LstAgregados.ItemData(X), uProd.codigo
      Next
      LimpiarTodo
   Else
      MsgBox "Primero debe agregar Conceptos a la lista superior", vbInformation, "Aviso"
      Exit Sub
   End If
Case "Modif"
      Controlar
End Select
End Sub

Private Sub Controlar()
Dim X As Long
   For X = 0 To LstAgregados.ListCount - 1
      sql = "SELECT codigo,codigoprod FROM reqauditprod WHERE codigo=" & LstAgregados.ItemData(X) & " AND codigoprod = " & uProd.codigo & ""
      rs.Open sql, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
      If rs.EOF Then
         DataEnvironment1.dbo_INGRESOAUDITPROD "A", LstAgregados.ItemData(X), uProd.codigo
      End If
      rs.Close
   Next
End Sub
Private Sub UBotonera_Buscar()
Dim Resul As String
LimpiarTodo
sql = "SELECT codigo as [ Codigo Producto ],descripcion as [              Descripcion              ] FROM producto ORDER BY codigo ASC"
frmBuscar.MostrarSql sql
Resul = frmBuscar.resultado(1)
If Resul = "" Then
 Exit Sub
Else
   uProd.codigo = Resul
   UBotonera.BuscarOK
End If
TipoOp = "Buscar"
CmdEjecutar_Click
End Sub

Private Sub UBotonera_Cancelar()
   LimpiarTodo
End Sub

Private Sub UBotonera_eliminar()
Dim X As Long
For X = 0 To LstAgregados.ListCount - 1
   DataEnvironment1.dbo_INGRESOAUDITPROD "B", LstAgregados.ItemData(X), uProd.codigo
Next
UBotonera.EliminarOK
End Sub

Private Sub UBotonera_HabilitarEdicion(sino As Boolean)
   LstAgregados.enabled = sino
   LstSeleccion.enabled = sino
End Sub

Private Sub UBotonera_Modificar()
   TipoOp = "Modif"
End Sub

Private Sub UBotonera_Nuevo()
uProd.enabled = True
TipoOp = "Nuevo"
CmdEjecutar.enabled = True
End Sub

Private Sub UBotonera_SALIR()
Unload Me
End Sub

Private Sub LimpiarTodo()
   LstAgregados.clear
   LstSeleccion.clear
   uProd.codigo = ""
   uProd.DESCRIPCION = ""
   CmdEjecutar.enabled = False
End Sub


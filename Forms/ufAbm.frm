VERSION 5.00
Begin VB.Form ufAbm 
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7305
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   Begin GestionTonka.ucBotonera uMenu 
      Align           =   2  'Align Bottom
      Height          =   705
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   1244
      MsgConfirmaSalir=   ""
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   ""
      CaptionEliminar =   "&Eliminar"
   End
End
Attribute VB_Name = "ufAbm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mRs As ADODB.Recordset

Public Function abm(tabla As String, cpoCodigo As String, arrayCampos, frmCaption As String, Optional bCodigoEsString As Boolean = False, Optional bConActivo As Boolean = False)
    Set mRs = New ADODB.Recordset
    
    
End Function


'-------------------------------------------
Private Sub uMenu_AceptarAlta()
'
End Sub
Private Sub uMenu_AceptarModi()
'
End Sub
Private Sub uMenu_BorrarControles()
'
End Sub
Private Sub uMenu_eliminar()
'
End Sub
Private Sub uMenu_SALIR()
    Set mRs = Nothing
    Unload Me
End Sub
Private Sub uMenu_SeMovio()
'
End Sub

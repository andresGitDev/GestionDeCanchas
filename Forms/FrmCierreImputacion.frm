VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmCierreImputacion 
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox UltimaFecha 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   2895
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1005
      Width           =   1815
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   390
      Left            =   3600
      TabIndex        =   1
      Top             =   3615
      Width           =   990
   End
   Begin VB.CommandButton CmdConfirmar 
      Caption         =   "Confirmar"
      Height          =   390
      Left            =   2415
      TabIndex        =   0
      Top             =   3615
      Width           =   990
   End
   Begin MSComCtl2.MonthView MView 
      Height          =   2370
      Left            =   105
      TabIndex        =   2
      Top             =   480
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   1
      StartOfWeek     =   81854465
      CurrentDate     =   39381
      MinDate         =   36526
   End
   Begin VB.Label Label3 
      Caption         =   "Periodo Cerrado :"
      Height          =   225
      Left            =   2955
      TabIndex        =   5
      Top             =   735
      Width           =   1725
   End
   Begin VB.Label Label2 
      Caption         =   "Seleccione el mes que quiere hacer el cierre de Posiscion de IVA"
      Height          =   270
      Left            =   75
      TabIndex        =   4
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "Al selecionar un mes determinado no contempla la fecha seleccionada en el calendario"
      Height          =   420
      Left            =   60
      TabIndex        =   3
      Top             =   2955
      Width           =   4590
   End
End
Attribute VB_Name = "FrmCierreImputacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum tipoiva
    IVAComprasx = 0
    IVAVentasx = 1
End Enum
Dim APorDonde As Long
Dim xIVA As String
Dim ssql As String


Private Sub CmdConfirmar_Click()
Dim AFecha As String
    If ControlPrevio = False Then Exit Sub
    If MsgBox("Confirma el Cierre del " & xIVA & Chr(13) & "para el periodo : " & MView.Month & "/" & MView.Year, vbQuestion + vbYesNo, "Confirmacion") = vbYes Then
        If MView.Month = mvwDecember Then
           MView.Month = mvwJanuary
           MView.Year = MView.Year + 1
        Else
          MView.Month = MView.Month + 1
        End If
        AFecha = "1/" & MView.Month & "/" & MView.Year
        If APorDonde = 0 Then
            ssql = "Update DatosEmpresa SET fechaimputc= '" & AFecha & "'  WHERE idempresa=" & gEMPR_idEmpresa
            DataEnvironment1.AMR.Execute ssql
        Else
            ssql = "Update DatosEmpresa SET fechaimputv= '" & AFecha & "'  WHERE idempresa=" & gEMPR_idEmpresa
            DataEnvironment1.AMR.Execute ssql
        End If
    End If
End Sub
Function ControlPrevio() As Boolean
    Dim rs As New ADODB.Recordset
    If MView.Value < CDate(UltimaFecha) Xor MView.Value > DateAdd("m", 1, UltimaFecha) Then
        MsgBox "No Puede hacerse el Cierre en la fecha Indicada", vbInformation, "Aviso"
             ControlPrevio = False
    Else: ControlPrevio = True
    End If
End Function
Private Sub cmdsalir_Click()
Unload Me
End Sub

Sub PorDonde(tipoiva As tipoiva)
Dim rsb As New ADODB.Recordset
APorDonde = tipoiva
Select Case tipoiva
    Case 0
        Me.caption = "Cierre IVA Compras"
        ssql = "SELECT fechaimputc FROM DatosEmpresa WHERE idempresa=" & gEMPR_idEmpresa & ""
        xIVA = "IVA Compras"
    Case 1
        Me.caption = "Cierre IVA Ventas"
        ssql = "SELECT fechaimputv FROM DatosEmpresa WHERE idempresa=" & gEMPR_idEmpresa & ""
        xIVA = "IVA Ventas"
End Select
rsb.Open ssql, DataEnvironment1.AMR, adOpenKeyset, adLockReadOnly
If Not rsb.EOF Then
    UltimaFecha = Format$(rsb.Fields(0), "mmmm yyyy")
End If
rsb.Close
Set rsb = Nothing
End Sub


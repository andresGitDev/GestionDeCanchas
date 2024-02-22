VERSION 5.00
Begin VB.Form frmRemitoVentaCancela 
   Caption         =   "Cancelacion de Remitos"
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6510
   Icon            =   "frmRemitoVentaCancela.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   6510
   StartUpPosition =   2  'CenterScreen
   Begin Gestion.ucEntreFechas uFechas 
      Height          =   315
      Left            =   3300
      TabIndex        =   7
      Top             =   960
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
   End
   Begin VB.TextBox txtMotivo 
      Height          =   300
      Left            =   300
      TabIndex        =   3
      Top             =   2040
      Width           =   6075
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   5235
      TabIndex        =   2
      Top             =   2430
      Width           =   975
   End
   Begin VB.CommandButton cmdAnular 
      Caption         =   "Cancelacion Remito"
      Height          =   375
      Left            =   3465
      TabIndex        =   1
      Top             =   2430
      Width           =   1755
   End
   Begin Gestion.ucCoDe ucRemitoCliente 
      Height          =   315
      Left            =   300
      TabIndex        =   0
      Top             =   480
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   556
      CodigoWidth     =   1000
   End
   Begin VB.Label Label1 
      Caption         =   "Buscar Entre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   1800
      TabIndex        =   8
      Top             =   1020
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "Motivo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   300
      TabIndex        =   6
      Top             =   1680
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Remito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   300
      TabIndex        =   5
      Top             =   120
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   2220
      TabIndex        =   4
      Top             =   120
      Width           =   795
   End
End
Attribute VB_Name = "frmRemitoVentaCancela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdAnular_Click()
    If ON_ERROR_HABILITADO Then On Error GoTo UfaCANCEL
    Dim tmp, num As Long, tmpFac
    
    num = ucRemitoCliente.codigo
    
    
    'Controlo----------------
    'sin remito
    If num = 0 Then
        che "no se especifico remito"
        Exit Sub
    End If
    'sin motivo
    If Trim(txtMotivo) = "" Then
        che "falta motivo"
        Exit Sub
    End If
    'ya anulado
    If obtenerDeSQL("select cancelado from RemitoVenta where numero = " & num) = True Then
        che "Remito ya figuraba Cancelado"
        Exit Sub
    End If
    'ya cancelado
    If obtenerDeSQL("select anulado from RemitoVenta where numero = " & num) = True Then
        che "Remito ya figuraba como Anulado"
        Exit Sub
    End If
    'ya facturado
    tmp = obtenerDeSQL("select sum(cantidad-facturar) as saldo from RemitoVentaDetalle where numero = " & num)
    If tmp > 0 Then
          che "No puedo anular, remito con factura " '& vbCrLf & tmpFac(0) & " " & tmpFac(1)
          Exit Sub
    End If
'    tmp = obtenerDeSQL("select factura from RemitoVenta where numero = " & num)
    
 '   If tmp > 0 Then
  '      tmpFac = obtenerDeSQL("select TipoDoc, NroFactura from FacturaVenta where activo = 1 and codigo = " & tmp)
  '
  '      If Not IsEmpty(tmpFac) Then
  '          che "No puedo anular, remito con factura " & vbCrLf & tmpFac(0) & " " & tmpFac(1)
  '          Exit Sub
  '      End If
  '  End If
    'Controlo----------------
   
    If confirma("Cancela remito " & num) Then
        'daTaenvironment1.dbo_abmRemitoVenta "B",  num, 0, 0, 0, 0, 0, 0, "", "", "", txtMotivo
        DataEnvironment1.Sistema.Execute "Update RemitoVenta set obs4 = '" & Trim(txtMotivo) & "', cancelado = 1  where  numero = " & num
        MsgBox "Cancelado"
    End If
    
fin:
    Exit Sub
UfaCANCEL:
    ufa "Error en cancelar remito", Me.Name & ucRemitoCliente.codigo ', Err
    Resume fin
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    uFechas.ini CDate("1/1/" & Year(Date)), Date
    ucRemitoCliente.ini "select descripcion from RemitoVenta inner join clientes on RemitoVenta.cliente = clientes.codigo where Numero = ###", sBuscar()
End Sub

Private Sub ucRemitoCliente_Buscar()
    ucRemitoCliente.strSqlBuscar = sBuscar()
End Sub

Private Function sBuscar()
    sBuscar = "select distinct Numero, Fecha as [ Fecha               ], descripcion as [ Cliente                                     ] from RemitoVenta inner join Clientes on clientes.codigo = RemitoVenta.Cliente where fecha " & uFechas.ssBetween & " order by numero desc"
End Function


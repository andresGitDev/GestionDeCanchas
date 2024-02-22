VERSION 5.00
Begin VB.Form frmFacturaVentaAnulacion 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Anular Facturas"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   9360
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optTipoDoc 
      BackColor       =   &H00C0C0C0&
      Caption         =   "NDE"
      Height          =   375
      Index           =   7
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4650
      Width           =   1575
   End
   Begin VB.OptionButton optTipoDoc 
      BackColor       =   &H00C0C0C0&
      Caption         =   "NCE"
      Height          =   375
      Index           =   6
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3570
      Width           =   1575
   End
   Begin VB.OptionButton optTipoDoc 
      BackColor       =   &H00C0C0C0&
      Caption         =   "FAE"
      Height          =   375
      Index           =   5
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2025
      Width           =   1575
   End
   Begin VB.OptionButton optTipoDoc 
      BackColor       =   &H00C0C0C0&
      Caption         =   "NDA"
      Height          =   375
      Index           =   4
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox txtCodigo 
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1260
      Width           =   1035
   End
   Begin VB.OptionButton optTipoDoc 
      BackColor       =   &H00C0C0C0&
      Caption         =   "NCB"
      Height          =   375
      Index           =   3
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3105
      Width           =   1575
   End
   Begin VB.OptionButton optTipoDoc 
      BackColor       =   &H00C0C0C0&
      Caption         =   "NCA"
      Height          =   375
      Index           =   2
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2670
      Width           =   1575
   End
   Begin VB.OptionButton optTipoDoc 
      BackColor       =   &H00C0C0C0&
      Caption         =   "FAB"
      Height          =   375
      Index           =   1
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1560
      Width           =   1575
   End
   Begin VB.OptionButton optTipoDoc 
      BackColor       =   &H00C0C0C0&
      Caption         =   "FAA"
      Height          =   375
      Index           =   0
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1080
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmdAnular 
      Caption         =   "Anular Comprobante"
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Top             =   3120
      Width           =   1755
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   7800
      TabIndex        =   1
      Top             =   3120
      Width           =   975
   End
   Begin Gestion.ucEntreFechas uFechas 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
   End
   Begin Gestion.ucCoDe ucFacturaCliente 
      Height          =   315
      Left            =   3180
      TabIndex        =   3
      Top             =   600
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   556
      CodigoWidth     =   1000
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "cod Interno "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   2
      Left            =   4680
      TabIndex        =   12
      Top             =   960
      Width           =   795
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   0
      Left            =   4740
      TabIndex        =   6
      Top             =   240
      Width           =   795
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Factura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   1
      Left            =   3180
      TabIndex        =   5
      Top             =   240
      Width           =   795
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   180
      Width           =   1395
   End
End
Attribute VB_Name = "frmFacturaVentaAnulacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '15/12/4 NEW

'anula FAA, FAB, NCA, NCB, NDA, NDB

Private Sub cmdAnular_Click()
    Dim iddoc As Long
    If s2n(txtCodigo) = 0 Then Exit Sub
    iddoc = s2n(obtenerDeSQL("select iddoc from FacturaVenta where codigo = '" & s2n(txtCodigo) & "' "))
    
    
    If AnularFacturaVenta(s2n(txtCodigo), iddoc) Then
        che " Documento  " & sTipo() & " " & ucFacturaCliente.codigo & " Anulado"
    End If
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    uFechas.ini CDate("1/1/" & Year(Date)), Date
    ucFacturaCliente.ini "select descripcion from FacturaVenta inner join clientes on FacturaVenta.cliente = clientes.codigo where  tipoDoc = '" & sTipo() & "' and NroFactura = ###", " " ', sBuscar()
End Sub

Private Function sBuscar() As String
    sBuscar = "select  NroFactura, Fecha as [ Fecha               ], descripcion as [ Cliente                                     ] from FacturaVenta inner join Clientes on clientes.codigo = FacturaVenta.Cliente where fecha " & uFechas.ssBetween & " and tipoDoc = '" & sTipo() & "' order by nroFactura desc"
End Function

Private Sub optTipoDoc_Click(Index As Integer)
    ucFacturaCliente.ini "select descripcion from FacturaVenta inner join clientes on FacturaVenta.cliente = clientes.codigo where  tipoDoc = '" & sTipo() & "' and NroFactura = ###", " "
    txtCodigo = ""
End Sub

Private Sub ucFacturaCliente_Buscar()
    ucFacturaCliente.strSqlBuscar = sBuscar()
End Sub

Private Sub ucFacturaCliente_cambio(codigo As Variant)
    txtCodigo = ""
    If ucFacturaCliente.codigo > 0 Then
        txtCodigo = obtenerDeSQL("select codigo from FacturaVenta where TipoDoc = '" & sTipo() & "' and NroFactura = " & ucFacturaCliente.codigo)
    End If
End Sub

Private Function sTipo() As String
    Dim i As Long
    With optTipoDoc
        For i = 0 To .UBound
            If optTipoDoc.Item(i).Value = True Then
                sTipo = .Item(i).caption
            End If
        Next
    End With
End Function


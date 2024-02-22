VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmCotizaciones 
   Caption         =   "Carga de Cotizaciones"
   ClientHeight    =   2130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4320
   Icon            =   "FrmCotizaciones.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2130
   ScaleWidth      =   4320
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker dtfecha 
      Height          =   255
      Left            =   2520
      TabIndex        =   0
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      Format          =   68812801
      CurrentDate     =   38113
   End
   Begin VB.ComboBox cmbmoneda 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdaceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdcancelar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txtcotizacion 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3240
      TabIndex        =   2
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha :"
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
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   240
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   1455
      Left            =   120
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   "Cotizacion :"
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
      Left            =   360
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Moneda :"
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
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "FrmCotizaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '16/9/4

Private Sub cmbmoneda_Click()
Dim rs As New ADODB.Recordset

    rs.Open "select * from cotizaciones where moneda=" & ObtenerCodigo("Monedas", Trim(cmbmoneda.Text)) & " order by fecha desc", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    If Not rs.EOF Then
        rs.MoveFirst
        dtFecha.Value = rs!fecha
        txtcotizacion = Replace(rs!cotizacion, ",", ".")
    End If
    rs.Close
    Set rs = Nothing
End Sub

Private Sub cmdAceptar_Click()

Dim rs As New ADODB.Recordset
Dim fecha As Variant
Dim mensaje As String


rs.Open "select * from cotizaciones where moneda=" & ObtenerCodigo("Monedas", Trim(cmbmoneda.Text)) & " and fecha=" & ssFecha(dtFecha) & "", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
If Not rs.EOF Then
    mensaje = MsgBox("Ya hay una Cotizacion Cargada para esta fecha,desea cambiarla", vbYesNo, "Atencion")
    If mensaje = 6 Then
        
        DataEnvironment1.dbo_COTIZACION "M", dtFecha.Value, ObtenerCodigo("Monedas", Trim(cmbmoneda.Text)), CDbl(Replace(txtcotizacion, ".", ","))
        MsgBox "La Operacion se ha realizado con exito", 48, "Atencion"
    End If
Else
    DataEnvironment1.dbo_COTIZACION "A", dtFecha.Value, ObtenerCodigo("Monedas", Trim(cmbmoneda.Text)), CDbl(Replace(txtcotizacion, ".", ","))
    MsgBox "La Operacion se ha realizado con exito", 48, "Atencion"
End If
rs.Close
Set rs = Nothing
    
End Sub

Private Sub cmdCancelar_Click()
    LimpioTxt
End Sub
Sub LimpioTxt()

    cmbmoneda.Text = ""
    txtcotizacion = "0.00"
    dtFecha.Value = Date
    
End Sub
Private Sub cmdsalir_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    CargaCombo cmbmoneda, "monedas", "descripcion", "codigo", ""
    LimpioTxt
End Sub

Private Sub txtcotizacion_GotFocus()
    txtcotizacion.SelStart = 0
    txtcotizacion.SelLength = Len(txtcotizacion.Text)
End Sub

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmVerCotizaciones 
   Caption         =   "Cotizaciones cargadas"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4815
   Icon            =   "frmVerCotizaciones.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   8
      Top             =   4200
      Width           =   975
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4200
      Width           =   975
   End
   Begin VB.ComboBox cmbmoneda 
      Height          =   315
      Left            =   1800
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   3480
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker dtfechadesde 
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   2520
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Format          =   67305473
      CurrentDate     =   38117
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grillacotizacion 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   3836
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   14737632
      ForeColorFixed  =   8388608
      BackColorBkg    =   14737632
      GridColorFixed  =   14737632
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComCtl2.DTPicker dtfechahasta 
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   3000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Format          =   67305473
      CurrentDate     =   38117
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      Height          =   1575
      Left            =   120
      Top             =   2400
      Width           =   4575
   End
   Begin VB.Label Label3 
      Caption         =   "Moneda"
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
      Left            =   240
      TabIndex        =   5
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Hasta Fecha"
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
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Desde Fecha"
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
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   1815
   End
End
Attribute VB_Name = "frmVerCotizaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '16/9/4


Dim rs As New ADODB.Recordset
Private Sub cmdAceptar_Click()

Dim FechaDesde As Variant
Dim FechaHasta As Variant

    If cmbmoneda.Text = "" Then
        MsgBox "Debe seleccionar una moneda", vbExclamation + vbOKOnly, "Error"
        Exit Sub
    End If
    
    
    FechaDesde = Month(dtfechadesde.Value) & "/" & Day(dtfechadesde.Value) & "/" & Year(dtfechadesde.Value)
    FechaHasta = Month(dtfechahasta.Value) & "/" & Day(dtfechahasta.Value) & "/" & Year(dtfechahasta.Value)
    
    rs.Open "Select moneda, fecha, cotizacion from Cotizaciones where fecha between '" & FechaDesde & "' and '" & FechaHasta & "' and Moneda = " & ObtenerCodigo("Monedas", cmbmoneda.Text), DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    Call CargarGrilla
    
    If rs.State = 1 Then
        rs.Close
        Set rs = Nothing
    End If
End Sub
Private Sub CargarGrilla()
  
    If rs.EOF Then
        MsgBox "No existen cotizaciones para la moneda seleccionada en el rango defechas seleccionados", vbOKOnly + vbExclamation, "No hay datos para mostrar"
        Exit Sub
    End If

    rs.MoveFirst
    If Not rs.EOF Then
        BorrarGrilla
        Do While Not rs.EOF
            If grillacotizacion.rows = 2 Then
                grillacotizacion.Row = 1
                grillacotizacion.Col = 0
                If grillacotizacion.Text = "" Then
                    grillacotizacion.Col = 0
                    grillacotizacion.Text = rs!fecha
                    grillacotizacion.Col = 1
                    grillacotizacion.Text = ObtenerDescripcion("Monedas", rs!moneda)
                    grillacotizacion.Col = 2
                    grillacotizacion.Text = Trim(rs!cotizacion)
                Else
                    grillacotizacion.AddItem rs!fecha & Chr(9) & ObtenerDescripcion("Monedas", rs!moneda) & Chr(9) & Trim(rs!cotizacion)
                End If
            Else
                grillacotizacion.AddItem rs!fecha & Chr(9) & ObtenerDescripcion("Monedas", rs!moneda) & Chr(9) & Trim(rs!cotizacion)
            End If
            rs.MoveNext
        Loop
    End If

End Sub

Private Sub BorrarGrilla()

Dim i As Long

    i = grillacotizacion.rows
    While i <> 1
        If grillacotizacion.rows = 2 Then
                grillacotizacion.Row = 1
                grillacotizacion.Col = 0
                If grillacotizacion.Text <> "" Then
                    grillacotizacion.Col = 0
                    grillacotizacion.Text = ""
                    grillacotizacion.Col = 1
                    grillacotizacion.Text = ""
                    grillacotizacion.Col = 2
                    grillacotizacion.Text = ""
                Else
                    Exit Sub
                End If
        Else
            grillacotizacion.RemoveItem (i - 1)
            i = i - 1
        End If
    Wend
    

End Sub
Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    FrmKeyPress KeyAscii, True, True

    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    CargaCombo cmbmoneda, "monedas", "descripcion", "codigo", "activo=1"
    LimpioTxt
    InicioGrilla
    
    cmbmoneda.ListIndex = BuscarenComboS(cmbmoneda, Const_PESOS)
    
End Sub
Sub LimpioTxt()
    dtfechadesde.Value = Date
    dtfechahasta.Value = Date
    cmbmoneda.Text = ""
End Sub
Sub InicioGrilla()
    grillacotizacion.FormatString = "^Fecha               |<Moneda                              |>Cotizacion      "
    grillacotizacion.rows = 2
    grillacotizacion.cols = 3
    grillacotizacion.Row = 1
    grillacotizacion.Col = 0
    If grillacotizacion.Text <> "" Then
        grillacotizacion.Col = 0
        grillacotizacion.Text = ""
        grillacotizacion.Col = 1
        grillacotizacion.Text = ""
        grillacotizacion.Col = 2
        grillacotizacion.Text = ""
    End If
End Sub

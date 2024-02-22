VERSION 5.00
Begin VB.Form FrmAbmCentrodeCostos 
   Caption         =   "Carga de Centro de Costos"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8385
   Icon            =   "FrmAbmCentrodeCostos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5055
   ScaleWidth      =   8385
   StartUpPosition =   2  'CenterScreen
   Begin Gestion.ucBotonera ucMenu 
      Height          =   1575
      Left            =   120
      TabIndex        =   14
      Top             =   3480
      Width           =   8175
      _extentx        =   14420
      _extenty        =   2778
      msgconfirmasalir=   ""
      msgconfirmaeliminar=   ""
      msgconfirmacancelar=   ""
      captioneliminar =   "&Eliminar"
   End
   Begin VB.TextBox txtNumPresu 
      Height          =   300
      Left            =   6450
      TabIndex        =   13
      Top             =   1260
      Width           =   1470
   End
   Begin VB.TextBox txtObser 
      Height          =   480
      Left            =   2040
      TabIndex        =   12
      Top             =   2685
      Width           =   5940
   End
   Begin VB.TextBox txtPresu 
      Height          =   285
      Left            =   2040
      TabIndex        =   11
      Top             =   2190
      Width           =   2790
   End
   Begin VB.TextBox txtOC 
      Height          =   285
      Left            =   2040
      TabIndex        =   10
      Top             =   1725
      Width           =   2775
   End
   Begin VB.TextBox txtCliente 
      Height          =   285
      Left            =   2040
      TabIndex        =   9
      Top             =   1275
      Width           =   2805
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   780
      Width           =   4680
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1905
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   285
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Nº Presupuesto :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4965
      TabIndex        =   8
      Top             =   1260
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Observacion :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   2670
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Presupuesto $ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   2205
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Orden de compra :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1740
      Width           =   1665
   End
   Begin VB.Label Label3 
      Caption         =   "Cliente :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1260
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Descripcion :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   345
      TabIndex        =   3
      Top             =   780
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   375
      TabIndex        =   2
      Top             =   300
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   3300
      Left            =   60
      Top             =   60
      Width           =   8145
   End
End
Attribute VB_Name = "FrmAbmCentrodeCostos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'19/11/4 new

'Form
Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True, True
End Sub
Private Sub Form_Load()
    CentrarMe Me
    ucMenu.init True, True, True, False, True, "select * from centrodecostos where activo=1 order by codigo", DataEnvironment1.Sistema
    ucMenu.MsgConfirmaEliminar = "Elimina este concepto ? "
End Sub

'Controls
Private Sub txtcodigo_GotFocus()
    GotFocusPinto txtCodigo
End Sub
Private Sub txtDescripcion_GotFocus()
    GotFocusPinto txtDescripcion
End Sub

'****************************************************************************
Private Sub ucMenu_AceptarAlta()
    On Error GoTo UFAalta
    If txtDescripcion.Text = "" Then
        MsgBox "Debe introducir una descripcion"
        Exit Sub
    End If
    
    If Not txtPresu.Text = "" Then
        If Not IsNumeric(txtPresu.Text) Then
            MsgBox "Debe ingresar solo numeros."
            Exit Sub
        End If
    Else
        txtPresu.Text = 0
    End If
    
    DataEnvironment1.dbo_CENTRODECOSTO "A", s2n(txtCodigo), Trim(txtDescripcion), txtCliente, txtOC, txtPresu, txtNumPresu, txtObser, Date, UsuarioSistema!codigo, 0, 0
    MsgBox "La Operación se ha realizado con éxito", 48, "Atencion"
    ucMenu.AceptarOk
fin:
    Exit Sub
UFAalta:
    ufa "err al dar el alta", Me.Name ', Err
    Resume fin
End Sub

Private Sub ucMenu_AceptarModi()
    On Error GoTo ufamodi
    If txtDescripcion.Text = "" Then
        MsgBox "Debe introducir una descripcion"
        Exit Sub
    End If
    
    If Not txtPresu.Text = "" Then
        If Not IsNumeric(txtPresu.Text) Then
            MsgBox "Debe ingresar solo numeros."
            Exit Sub
        End If
    Else
        txtPresu.Text = 0
    End If
    
    DataEnvironment1.dbo_CENTRODECOSTO "M", s2n(txtCodigo), Trim(txtDescripcion), txtCliente, txtOC, txtPresu, txtNumPresu, txtObser, 0, 0, 0, 0
    grabaBitacora "M", s2n(txtCodigo), "CentrodeCostos"
    MsgBox "La Operación se ha realizado con éxito", 48, "Atencion"
    ucMenu.AceptarOk
fin:
    Exit Sub
ufamodi:
    ufa "err al modificar", Me.Name ', Err
    Resume fin
End Sub
Private Sub ucMenu_BorrarControles()
    FrmBorrarTxt Me
End Sub
Private Sub ucMenu_Buscar()
    Dim resu
    Dim rs2 As New ADODB.Recordset
    
    With frmBuscar
        resu = .MostrarCodigoDescripcionActivo("centroDeCostos")
        If resu > "" Then
            txtCodigo = resu
            txtDescripcion = .resultado(2)
            rs2.Open "select * from centrodecostos where codigo='" & resu & "' and descripcion='" & .resultado(2) & "'", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            If rs2.EOF = True And rs2.BOF = True Then
                txtCliente.Text = ""
                txtOC.Text = ""
                txtPresu = ""
                txtNumPresu = ""
                txtObser = ""
            Else
                txtCliente.Text = IIf(rs2!cliente > "", rs2!cliente, "")
                txtOC.Text = IIf(rs2!orden_compra > "", rs2!orden_compra, "")
                txtPresu = IIf(rs2!PresupuestoPESOS > "", rs2!PresupuestoPESOS, "")
                txtNumPresu = IIf(rs2!num_presupuesto > "", rs2!num_presupuesto, "")
                txtObser = IIf(rs2!observacion > "", rs2!observacion, "")
            End If
            Set rs2 = Nothing
            ucMenu.BuscarOK "codigo = " & resu
        End If
    End With
End Sub
Private Sub ucMenu_eliminar()
    On Error GoTo UFAelimina
    DataEnvironment1.dbo_CENTRODECOSTO "B", Trim(txtCodigo), "", "", "", 0, "", "", 0, 0, UsuarioSistema!codigo, Date
    grabaBitacora "B", s2n(txtCodigo), "centrodecostos"
    ucMenu.EliminarOK
fin:
    Exit Sub
UFAelimina:
    ufa "err al eliminar", Me.Name ', Err
    Resume fin
End Sub
Private Sub ucMenu_HabilitarEdicion(sino As Boolean)
    txtDescripcion.enabled = sino
    txtCliente.enabled = sino
    txtOC.enabled = sino
    txtPresu.enabled = sino
    txtNumPresu.enabled = sino
    txtObser.enabled = sino
End Sub
Private Sub ucMenu_Nuevo()
    txtCodigo = nuevoCodigo("CentroDeCostos")
End Sub
Private Sub ucMenu_SALIR()
    Unload Me
End Sub
Private Sub ucMenu_SeMovio()
    Dim rs1 As New ADODB.Recordset
    
    On Error Resume Next
    txtCodigo = s2n(ucMenu.rs!codigo, 0)
    txtDescripcion = ucMenu.rs!DESCRIPCION
    
    rs1.Open "select * from centrodecostos where codigo='" & txtCodigo.Text & "' and descripcion='" & txtDescripcion.Text & "'", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If rs1.EOF = True And rs1.BOF = True Then
        txtCliente.Text = ""
        txtOC.Text = ""
        txtPresu = ""
        txtNumPresu = ""
        txtObser = ""
    Else
        txtCliente.Text = IIf(rs1!cliente > "", rs1!cliente, "")
        txtOC.Text = IIf(rs1!orden_compra > "", rs1!orden_compra, "")
        txtPresu = IIf(rs1!PresupuestoPESOS > "", rs1!PresupuestoPESOS, "")
        txtNumPresu = IIf(rs1!num_presupuesto > "", rs1!num_presupuesto, "")
        txtObser = IIf(rs1!observacion > "", rs1!observacion, "")
    End If
    Set rs1 = Nothing
End Sub



VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmABMCliente 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ABM Clientes"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Categoria de Cliente "
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
      Height          =   1065
      Left            =   135
      TabIndex        =   35
      Top             =   4710
      Width           =   11040
      Begin VB.CheckBox ChkMayor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mayorista"
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
         Height          =   285
         Left            =   8655
         TabIndex        =   17
         Top             =   488
         Width           =   1605
      End
      Begin VB.TextBox Puntos 
         Height          =   360
         Left            =   6960
         TabIndex        =   16
         Top             =   450
         Width           =   1170
      End
      Begin VB.ComboBox CmbTipoCli 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1950
         TabIndex        =   15
         Text            =   "CmbTipoCli"
         Top             =   450
         Width           =   3975
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Puntos"
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
         Height          =   270
         Left            =   6195
         TabIndex        =   37
         Top             =   495
         Width           =   735
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Cliente"
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
         Height          =   300
         Left            =   180
         TabIndex        =   36
         Top             =   480
         Width           =   1485
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Datos Personales "
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
      Height          =   3000
      Left            =   120
      TabIndex        =   26
      Top             =   360
      Width           =   11040
      Begin VB.TextBox Observ 
         Height          =   690
         Left            =   1620
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   2220
         Width           =   8235
      End
      Begin VB.ComboBox CmbTIva 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7155
         TabIndex        =   6
         Top             =   1252
         Width           =   1590
      End
      Begin VB.TextBox Pais 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5265
         TabIndex        =   9
         Top             =   1770
         Width           =   4560
      End
      Begin VB.TextBox mail 
         Height          =   360
         Left            =   1125
         TabIndex        =   8
         Top             =   1770
         Width           =   3330
      End
      Begin VB.CheckBox ChkMailing 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mailing"
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
         Height          =   315
         Left            =   8790
         TabIndex        =   7
         Top             =   1275
         Width           =   1020
      End
      Begin VB.TextBox Telefono 
         Height          =   360
         Left            =   1125
         TabIndex        =   4
         Top             =   1275
         Width           =   3330
      End
      Begin VB.TextBox Empresa 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6285
         TabIndex        =   3
         Top             =   750
         Width           =   4590
      End
      Begin VB.TextBox Apellido 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1125
         TabIndex        =   0
         Top             =   315
         Width           =   4335
      End
      Begin VB.TextBox Nombre 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6300
         TabIndex        =   1
         Top             =   300
         Width           =   4560
      End
      Begin GestionMamy.ucCuit ucCuit 
         Height          =   360
         Left            =   5235
         TabIndex        =   5
         Top             =   1245
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   529
      End
      Begin MSMask.MaskEdBox FechaCumpl 
         Height          =   360
         Left            =   3015
         TabIndex        =   2
         Top             =   750
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones"
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
         Height          =   285
         Left            =   150
         TabIndex        =   39
         Top             =   2190
         Width           =   1485
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo  IVA"
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
         Height          =   450
         Left            =   6675
         TabIndex        =   38
         Top             =   1215
         Width           =   510
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "País :"
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
         Left            =   4590
         TabIndex        =   34
         Top             =   1830
         Width           =   555
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail"
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
         Left            =   180
         TabIndex        =   33
         Top             =   1830
         Width           =   690
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono"
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
         Height          =   240
         Left            =   180
         TabIndex        =   32
         Top             =   1335
         Width           =   1005
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Cumpleaños"
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
         Height          =   270
         Left            =   180
         TabIndex        =   31
         Top             =   795
         Width           =   2685
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Trabaja en :"
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
         Height          =   270
         Left            =   5055
         TabIndex        =   30
         Top             =   795
         Width           =   1185
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "CUIT :"
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
         Height          =   315
         Left            =   4575
         TabIndex        =   29
         Top             =   1290
         Width           =   660
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido"
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
         Height          =   285
         Left            =   180
         TabIndex        =   28
         Top             =   345
         Width           =   885
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
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
         Height          =   285
         Left            =   5520
         TabIndex        =   27
         Top             =   345
         Width           =   930
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Direccion "
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
      Height          =   1350
      Left            =   135
      TabIndex        =   21
      Top             =   3360
      Width           =   11040
      Begin VB.TextBox CP 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6930
         TabIndex        =   14
         Top             =   795
         Width           =   1290
      End
      Begin VB.ComboBox CmbProvincia 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1155
         TabIndex        =   13
         Top             =   795
         Width           =   4590
      End
      Begin VB.TextBox Localidad 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6930
         TabIndex        =   12
         Top             =   300
         Width           =   4020
      End
      Begin VB.TextBox Calle 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1170
         TabIndex        =   11
         Top             =   292
         Width           =   4590
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "C.Postal"
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
         Height          =   315
         Left            =   5880
         TabIndex        =   25
         Top             =   825
         Width           =   1065
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Localidad"
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
         Height          =   240
         Left            =   5880
         TabIndex        =   24
         Top             =   345
         Width           =   1050
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Provincia"
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
         Height          =   300
         Left            =   165
         TabIndex        =   23
         Top             =   840
         Width           =   930
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Calle"
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
         Height          =   285
         Left            =   165
         TabIndex        =   22
         Top             =   330
         Width           =   570
      End
   End
   Begin VB.TextBox NCliente 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   9570
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   22
      Width           =   1560
   End
   Begin GestionMamy.ucBotonera ucMenu 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   18
      Top             =   5820
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   1323
      MsgConfirmaSalir=   ""
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   ""
      CaptionEliminar =   "&Eliminar"
      CaptionImprimir =   "&Imprimir"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente Nº"
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
      Height          =   225
      Left            =   8595
      TabIndex        =   20
      Top             =   90
      Width           =   960
   End
End
Attribute VB_Name = "FrmABMCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Apellido_KeyPress(KeyAscii As Integer)
KeyAscii = Mayuscula(KeyAscii)
End Sub

Private Sub Calle_KeyPress(KeyAscii As Integer)
KeyAscii = Mayuscula(KeyAscii)
End Sub

Private Sub Empresa_KeyPress(KeyAscii As Integer)
KeyAscii = Mayuscula(KeyAscii)
End Sub

Private Sub Form_Load()
    Top = Centrar(0, Height)
    Left = Centrar(Width, 0)
    ucMenu.init True, True, True, False, True
    CargoTipoIva
    CargaProvincia
    Categoria
End Sub
Sub CargoTipoIva()
Dim rs As New ADODB.Recordset
Sql = "SELECT * FROM ivas "
rs.Open Sql, DataEnvironment1.AMR, adOpenKeyset, adLockReadOnly
Do While Not rs.EOF
    CmbTIva.AddItem rs!Descripcion
    CmbTIva.ItemData(CmbTIva.NewIndex) = rs!Codigo
    rs.MoveNext
Loop
rs.Close
End Sub

Sub BuscoCliente()
Dim rsBusq As New ADODB.Recordset
Dim Aux As Double
    Sql = "SELECT numero FROM clientes WHERE numero like '0004%' order by numero desc"
    rsBusq.Open Sql, DataEnvironment1.AMR, adOpenKeyset, adLockReadOnly
    If Not rsBusq.EOF Then
    Aux = CDbl(Mid(rsBusq!numero, 6, Len(rsBusq!numero)))
    Aux = Aux + 1
        NCliente = "0004-" & Format(Aux, "000000")
    End If
End Sub

Private Sub Localidad_KeyPress(KeyAscii As Integer)
KeyAscii = Mayuscula(KeyAscii)
End Sub

Private Sub mail_KeyPress(KeyAscii As Integer)
KeyAscii = Minuscula(KeyAscii)
End Sub

Private Sub Nombre_KeyPress(KeyAscii As Integer)
KeyAscii = Mayuscula(KeyAscii)
End Sub

Private Sub Pais_KeyPress(KeyAscii As Integer)
KeyAscii = Mayuscula(KeyAscii)
End Sub

Private Sub Puntos_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNum(KeyAscii)
End Sub
Public Function SoloNum(ByVal KeyAscii As Integer)
If KeyAscii = 46 Or KeyAscii = 44 Then
    KeyAscii = 44
ElseIf Not (KeyAscii > 47 And KeyAscii < 58 Or KeyAscii = 8) Then
                KeyAscii = 0
End If
SoloNum = KeyAscii
End Function


Private Sub ucMenu_AceptarAlta()
Dim rsCli As New ADODB.Recordset
Dim AuxDesc As String
On Error GoTo Fallo
DE_BeginTrans
    If fechcumpl = "" Then
       FechaCumpl = #1/1/1900#
    End If
    AuxDesc = Apellido.Text & " " & Nombre.Text
    Sql = " insert into clientes (descripcion,direccion,codpostal,localidad,provincia,telefono, " & _
    "fechcumpl,empresa,tipoiva,cuit,obs,mailing,categoria,puntos,numero,sucursal,mail,pais,mayorista,activo) " & _
    "values ('" & AuxDesc & "','" & Calle & "','" & CP & "','" & Localidad & "','" & CmbProvincia.Text & "'" & _
    ",'" & Telefono & "','" & FechaCumpl & "','" & Empresa & "'," & CmbTIva.ItemData(CmbTIva.ListIndex) & ",'" & ucCuit.Text & "'," & _
    "'" & Observ & "'," & ChkMailing.Value & ",'" & CmbTipoCli.ItemData(CmbTipoCli.ListIndex) & "'," & Puntos & ",'" & NCliente & "',0,'" & mail & "','" & Pais & "'," & ChkMayor.Value & ",1)"
    rsCli.Open Sql, DataEnvironment1.AMR, adOpenKeyset, adLockOptimistic
DE_CommitTrans
ucMenu.AceptarOk

Exit Sub
Fallo:
    MsgBox "Fallo la grabacion", vbExclamation, "Error"
    DE_RollbackTrans
End Sub

Private Sub ucMenu_AceptarModi()
Dim rsCli As New ADODB.Recordset
Dim XA As Long, XB As Long, AuxDesc As String

'On Error GoTo Fallo
DE_BeginTrans
    XA = sSinNull(obtenerDeSQL("SELECT codigo FROM ivas WHERE descripcion= '" & CmbTIva.Text & "'"))
    XB = sSinNull(obtenerDeSQL("SELECT id FROM categclie WHERE descripcion= '" & CmbTipoCli.Text & "'"))
    AuxDesc = Apellido & " " & Nombre
    Sql = " update clientes set descripcion='" & AuxDesc & "',direccion='" & Calle & "',codpostal='" & CP & "'" & _
    ",localidad='" & Localidad & "',provincia='" & CmbProvincia.Text & "',telefono='" & Telefono & "'," & _
    "fechcumpl='" & FechaCumpl & "',empresa='" & Empresa & "',tipoiva=" & XA & "" & _
    ",cuit='" & ucCuit.Text & "',obs='" & Observ & "',mailing=" & ChkMailing.Value & ",categoria='" & XB & "'" & _
    ",puntos=" & Puntos & ",numero='" & NCliente & "',sucursal=0,mail='" & mail & "',pais='" & Pais & "',mayorista=" & ChkMayor.Value & ",activo=1 WHERE numero='" & NCliente & "'"
    rsCli.Open Sql, DataEnvironment1.AMR, adOpenKeyset, adLockOptimistic
DE_CommitTrans
ucMenu.AceptarOk

Exit Sub
Fallo:
    MsgBox "Fallo la grabacion", vbExclamation, "Error"
    DE_RollbackTrans
End Sub

Private Sub ucMenu_BorrarControles()
    Limpio
End Sub
Sub Limpio()
    Apellido.Text = ""
    Nombre.Text = ""
    FechaCumpl.Text = "__/__/____"
    Empresa.Text = ""
    Telefono.Text = ""
    ucCuit.Text = ""
    ChkMailing.Value = 0
    mail.Text = ""
    Pais.Text = ""
    Calle.Text = ""
    Localidad.Text = ""
    CmbProvincia.Text = ""
    CP.Text = ""
    CmbTipoCli.Text = ""
    Puntos.Text = ""
    CmbTIva.Text = ""
    Observ.Text = ""
    
End Sub

Private Sub ucMenu_Buscar()
Dim resu As String, Sql As String
Dim rsCli As New ADODB.Recordset
Dim TempX
Sql = "SELECT numero as [Nº de Cliente  ],descripcion as [Cliente                                ],codigo as [_H_codigo] FROM clientes WHERE activo=1"
resu = frmBuscar.MostrarSql(Sql)

If resu > "" Then
   Sql = "SELECT * FROM CLIENTES WHERE numero = '" & frmBuscar.resultado(1) & "'"
   rsCli.Open Sql, DataEnvironment1.AMR, adOpenKeyset, adLockReadOnly
   With rsCli
        NCliente = sSinNull(!numero)
        Apellido = sSinNull(SacarValor(!Descripcion, True, False))
        Nombre = sSinNull(SacarValor(!Descripcion, False, True))
        Calle = sSinNull(!Direccion)
        Localidad = sSinNull(!Localidad)
        Empresa = sSinNull(!Empresa)
        ucCuit.Text = sSinNull(!Cuit)
        
        If !mailing = True Then
            ChkMailing.Value = 1
        Else: ChkMailing.Value = 0
        End If
        
        mail = sSinNull(!mail)
        Pais = sSinNull(!Pais)
        If Not IsNull(!fechcumpl) Then FechaCumpl = !fechcumpl
        CmbProvincia.Text = sSinNull(!provincia)
        CP = sSinNull(!codpostal)
        TempX = sSinNull(obtenerDeSQL("SELECT descripcion FROM ivas WHERE codigo= " & CInt(!tipoiva) & ""))
        CmbTIva.Text = TempX
        Puntos = nSinNull(!Puntos)
        TempX = sSinNull(obtenerDeSQL("SELECT descripcion FROM categclie WHERE id= " & CInt(!Categoria) & ""))
        CmbTipoCli.Text = TempX
   End With
    ucMenu.BuscarOK
End If
End Sub

Private Sub ucMenu_eliminar()
On Error GoTo Fallo
    If MsgBox("Esta seguro de eliminar este Cliente", vbQuestion + vbYesNo, "Pedido de Confirmacion") = vbYes Then
        DE_BeginTrans
        DataEnvironment1.AMR.Execute "UPDATE clientes SET activo=0 WHERE numero='" & NCliente & "'"
        DE_CommitTrans
        ucMenu.EliminarOK
    End If
Exit Sub
Fallo:
    MsgBox "Fallo la Eliminacion", vbInformation, "Error"
    DE_RollbackTrans
End Sub

Private Sub ucMenu_HabilitarEdicion(sino As Boolean)
    Apellido.Enabled = sino
    Nombre.Enabled = sino
    FechaCumpl.Enabled = sino
    Empresa.Enabled = sino
    Telefono.Enabled = sino
    ucCuit.Enabled = sino
    ChkMailing.Enabled = sino
    mail.Enabled = sino
    Pais.Enabled = sino
    Calle.Enabled = sino
    Localidad.Enabled = sino
    CmbProvincia.Enabled = sino
    CP.Enabled = sino
    CmbTipoCli.Enabled = sino
    Puntos.Enabled = sino
    CmbTIva.Enabled = sino
    Observ.Enabled = sino
    
End Sub

Private Sub ucMenu_Nuevo()
    BuscoCliente
'    CargaProvincia
    Categoria
    Apellido.SetFocus
End Sub

Private Sub ucMenu_SALIR()
 Unload Me
End Sub
Sub CargaProvincia()
    CmbProvincia.AddItem "BUENOS AIRES"
    CmbProvincia.AddItem "CAPITAL FEDERAL"
    CmbProvincia.AddItem "CATAMARCA"
    CmbProvincia.AddItem "CHACO"
    CmbProvincia.AddItem "CHUBUT"
    CmbProvincia.AddItem "CORDOBA"
    CmbProvincia.AddItem "CORRIENTES"
    CmbProvincia.AddItem "ENTRE RIOS"
    CmbProvincia.AddItem "FORMOSA"
    CmbProvincia.AddItem "JUJUY"
    CmbProvincia.AddItem "LA PAMPA"
    CmbProvincia.AddItem "LA RIOJA"
    CmbProvincia.AddItem "MENDOZA"
    CmbProvincia.AddItem "MISIONES"
    CmbProvincia.AddItem "NEUQUEN"
    CmbProvincia.AddItem "RIO NEGRO"
    CmbProvincia.AddItem "SALTA"
    CmbProvincia.AddItem "SAN JUAN"
    CmbProvincia.AddItem "SAN LUIS"
    CmbProvincia.AddItem "SANTA CRUZ"
    CmbProvincia.AddItem "SANTA FE"
    CmbProvincia.AddItem "SANTIAGO DEL ESTERO"
    CmbProvincia.AddItem "TIERRA DEL FUEGO"
    CmbProvincia.AddItem "TUCUMAN"
End Sub
Sub Categoria()
Dim rs As New ADODB.Recordset
    Sql = "SELECT * FROM categclie"
    rs.Open Sql, DataEnvironment1.AMR, adOpenKeyset, adLockReadOnly
    Do While Not rs.EOF
            CmbTipoCli.AddItem rs!Descripcion
            CmbTipoCli.ItemData(CmbTipoCli.NewIndex) = rs!Id
    rs.MoveNext
    Loop
    rs.Close
End Sub

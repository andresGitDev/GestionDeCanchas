VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmABMClientes2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clientes"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11220
   Icon            =   "FrmABMClientes2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   11220
   StartUpPosition =   2  'CenterScreen
   Begin Gestion.ucBotonera ucMenu 
      Height          =   7890
      Left            =   1530
      TabIndex        =   35
      Top             =   6285
      Width           =   11220
      _extentx        =   19791
      _extenty        =   13917
      msgconfirmaeliminar=   ""
      msgconfirmacancelar=   ""
   End
   Begin VB.Frame Frame3 
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
      ForeColor       =   &H00000000&
      Height          =   1485
      Left            =   60
      TabIndex        =   30
      Top             =   4740
      Width           =   11040
      Begin VB.CommandButton cmdCuenta 
         Caption         =   "Cuenta"
         Enabled         =   0   'False
         Height          =   300
         Left            =   180
         TabIndex        =   41
         Top             =   1005
         Width           =   975
      End
      Begin VB.TextBox txtCuenta_Descrip 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2490
         TabIndex        =   40
         Top             =   1005
         Width           =   5220
      End
      Begin VB.CheckBox chkTiene_Cuenta 
         Caption         =   "Usar Cuenta"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7755
         TabIndex        =   39
         Top             =   975
         Width           =   1695
      End
      Begin VB.TextBox txtCuenta 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1170
         TabIndex        =   38
         Text            =   "0"
         Top             =   1005
         Width           =   1305
      End
      Begin VB.CheckBox ChkMayor 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   8655
         TabIndex        =   14
         Top             =   488
         Width           =   1605
      End
      Begin VB.TextBox Puntos 
         Height          =   360
         Left            =   6960
         TabIndex        =   13
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
         TabIndex        =   12
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
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   6195
         TabIndex        =   32
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   180
         TabIndex        =   31
         Top             =   480
         Width           =   1485
      End
   End
   Begin VB.Frame Frame2 
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
      ForeColor       =   &H00000000&
      Height          =   3000
      Left            =   45
      TabIndex        =   22
      Top             =   390
      Width           =   11040
      Begin Gestion.ucCuit ucCuit 
         Height          =   315
         Left            =   5160
         TabIndex        =   37
         Top             =   1260
         Width           =   1425
         _extentx        =   2514
         _extenty        =   556
      End
      Begin VB.TextBox Observ 
         Height          =   690
         Left            =   1620
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
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
         TabIndex        =   3
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
         TabIndex        =   6
         Top             =   1770
         Width           =   4560
      End
      Begin VB.TextBox mail 
         Height          =   360
         Left            =   1125
         TabIndex        =   5
         Top             =   1770
         Width           =   3330
      End
      Begin VB.CheckBox ChkMailing 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8790
         TabIndex        =   4
         Top             =   1275
         Width           =   1020
      End
      Begin VB.TextBox Telefono 
         Height          =   360
         Left            =   1125
         TabIndex        =   2
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
         TabIndex        =   1
         Top             =   750
         Width           =   4590
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
         Left            =   1260
         TabIndex        =   0
         Top             =   315
         Width           =   9225
      End
      Begin MSMask.MaskEdBox FechaCumpl 
         Height          =   345
         Left            =   2610
         TabIndex        =   36
         Top             =   765
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   609
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   150
         TabIndex        =   34
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
         ForeColor       =   &H00000000&
         Height          =   450
         Left            =   6675
         TabIndex        =   33
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4590
         TabIndex        =   29
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   180
         TabIndex        =   28
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   180
         TabIndex        =   27
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
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   180
         TabIndex        =   26
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
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   5055
         TabIndex        =   25
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4575
         TabIndex        =   24
         Top             =   1290
         Width           =   660
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   930
      End
   End
   Begin VB.Frame Frame1 
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
      ForeColor       =   &H00000000&
      Height          =   1350
      Left            =   60
      TabIndex        =   17
      Top             =   3390
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5880
         TabIndex        =   21
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5880
         TabIndex        =   20
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   165
         TabIndex        =   19
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   165
         TabIndex        =   18
         Top             =   330
         Width           =   570
      End
   End
   Begin VB.TextBox NCliente 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   9570
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   22
      Width           =   1560
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
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   8520
      TabIndex        =   16
      Top             =   75
      Width           =   960
   End
End
Attribute VB_Name = "FrmABMClientes2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Apellido_KeyPress(KeyAscii As Integer)
    KeyAscii = UCase(KeyAscii)
End Sub
Private Sub Calle_KeyPress(KeyAscii As Integer)
    KeyAscii = UCase(KeyAscii)
End Sub
Private Sub Empresa_KeyPress(KeyAscii As Integer)
    KeyAscii = UCase(KeyAscii)
End Sub
Private Sub Form_Load()
    ucMenu.init True, True, True, False, True
'    CargoTipoIva
    comboSql CmbTIva, "select descripcion, codigo from ivas where activo = 1"
    CargaProvincia
'    Categoria
    comboSql CmbTipoCli, "select descripcion, id from categclie "
End Sub

Sub BuscoCliente()
    Dim rsBusq As New ADODB.Recordset
    Dim Aux As Double, sql As String
    
    sql = "SELECT numero FROM clientes WHERE numero like '0004%' order by numero desc"
    rsBusq.Open sql, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
    If Not rsBusq.EOF Then
        Aux = CDbl(Mid(rsBusq!numero, 6, Len(rsBusq!numero)))
        Aux = Aux + 1
        NCliente = "0004-" & Format(Aux, "000000")
    End If
End Sub

Private Sub Localidad_KeyPress(KeyAscii As Integer)
    KeyAscii = UCase(KeyAscii)
End Sub
Private Sub mail_KeyPress(KeyAscii As Integer)
    KeyAscii = LCase(KeyAscii)
End Sub
Private Sub Nombre_KeyPress(KeyAscii As Integer)
    KeyAscii = UCase(KeyAscii)
End Sub

Private Sub Pais_KeyPress(KeyAscii As Integer)
    KeyAscii = UCase(KeyAscii)
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
    Dim AuxDesc As String
    
    If ON_ERROR_HABILITADO Then On Error GoTo Fallo
    Dim sql As String
    Dim iiva As Long
    Dim itipo As Long
    Dim dCumple As Variant

'    DE_BeginTrans
        If FechaCumpl = "" Then FechaCumpl = #1/1/1900#
        If (CmbTIva.ListIndex) >= 0 Then iiva = CmbTIva.ItemData(CmbTIva.ListIndex)
        If (CmbTipoCli.ListIndex) >= 0 Then itipo = CmbTIva.ItemData(CmbTipoCli.ListIndex)
        dCumple = datecumple()
        'AuxDesc = Apellido.Text & " " & Nombre.Text
        AuxDesc = Nombre.Text
        sql = " insert into clientes (descripcion,direccion,codigopostal,localidad,provincia,telefono, " & _
            " FechaCumpleAnio, empresa, iva, cuit, obs, mailing, categoria, puntos, codigo, mail, pais, mayorista, activo) " & _
            " values ('" & AuxDesc & "','" & Calle & "','" & CP & "','" & Localidad & "','" & CmbProvincia.Text & "'" & _
            ",'" & Telefono & "'," & dCumple & ", '" & Empresa & "'," & iiva & ",'" & ucCuit.Text & "'," & _
            " '" & Observ & "'," & ChkMailing.Value & ",'" & itipo & "'," & s2n(Puntos, 0) & ",'" & NCliente & "','" & mail & "','" & Pais & "'," & ChkMayor.Value & ",1)"
        DataEnvironment1.Sistema.Execute sql
        
'    DE_CommitTrans
    ucMenu.AceptarOk

Exit Sub
Fallo:
    MsgBox "Fallo la grabacion", vbExclamation, "Error"
'    DE_RollbackTrans
End Sub

Private Function datecumple() As String
    On Error Resume Next
    datecumple = "NULL"
    datecumple = ssFecha(CDate(FechaCumpl))
End Function

Private Sub ucMenu_AceptarModi()
    Dim rsCli As New ADODB.Recordset
    Dim XA As Long, XB As Long, AuxDesc As String
    Dim sql As String

'On Error GoTo Fallo
'DE_BeginTrans
    XA = sSinNull(obtenerDeSQL("SELECT codigo FROM ivas WHERE descripcion= '" & CmbTIva.Text & "'"))
    XB = sSinNull(obtenerDeSQL("SELECT id FROM categclie WHERE descripcion= '" & CmbTipoCli.Text & "'"))
    'AuxDesc = Apellido & " " & Nombre
    AuxDesc = Nombre
    
    sql = " update clientes set descripcion='" & AuxDesc & "',direccion='" & Calle & "',codigopostal='" & CP & "'" & _
        ", localidad='" & Localidad & "',provincia='" & CmbProvincia.Text & "',telefono='" & Telefono & "'," & _
        " FechaCumpleAnio = " & datecumple() & ",empresa='" & Empresa & "', iva=" & XA & "" & _
        ", cuit='" & ucCuit.Text & "',obs='" & Observ & "',mailing=" & ChkMailing.Value & ",categoria='" & XB & "'" & _
        ", puntos=" & Puntos & ", codigo ='" & NCliente & "', mail='" & mail & "',pais='" & Pais & "',mayorista=" & ChkMayor.Value & ",activo=1 WHERE codigo ='" & NCliente & "'"
'DE_CommitTrans
    DataEnvironment1.Sistema.Execute sql
    ucMenu.AceptarOk

Exit Sub
Fallo:
    MsgBox "Fallo la grabacion", vbExclamation, "Error"
'    DE_RollbackTrans
End Sub

Private Sub ucMenu_BorrarControles()
    limpio
End Sub
Sub limpio()
    FechaCumpl.Text = "__/__/____"
    ucCuit.Text = ""
    ChkMailing.Value = 0
    FrmBorrarTxt Me
    FrmBorrarCbo Me
End Sub

Private Sub ucMenu_Buscar()
    Dim resu As String, sql As String
    Dim rsCli As New ADODB.Recordset
    Dim TempX
    resu = BuscarCliente()

    If resu > "" Then
        sql = "SELECT * FROM CLIENTES WHERE codigo = '" & frmBuscar.resultado(1) & "' "
       
        rsCli.Open sql, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
        With rsCli
            NCliente = sSinNull(!codigo)
            Nombre = sSinNull(!DESCRIPCION) '  sSinNull(SacarValor(!descripcion, False, True))
            Calle = sSinNull(!direccion)
            Localidad = sSinNull(!Localidad)
            Empresa = sSinNull(!Empresa)
            ucCuit.Text = sSinNull(!CUIT)
            
            If !mailing = True Then
                ChkMailing.Value = 1
            Else: ChkMailing.Value = 0
            End If
            
            mail = sSinNull(!mail)
            Pais = sSinNull(!Pais)
            If Not IsNull(!FechaCumpleAnio) Then FechaCumpl = CStr(!FechaCumpleAnio) Else FechaCumpl = "__/__/____"
            CmbProvincia.Text = sSinNull(!Provincia)
            CP = sSinNull(!codigopostal)
            TempX = sSinNull(obtenerDeSQL("SELECT descripcion FROM ivas WHERE codigo= " & CInt(!Iva) & ""))
            CmbTIva.Text = TempX
            Puntos = nSinNull(!Puntos)
            TempX = sSinNull(obtenerDeSQL("SELECT descripcion FROM categclie WHERE id= " & CInt(!Categoria) & ""))
            CmbTipoCli.Text = TempX
            Observ.Text = sSinNull(!obs)
            Telefono.Text = !Telefono
            ChkMayor.Value = IIf(IsNull(!mayorista), 0, Abs(!mayorista))
        End With
        ucMenu.BuscarOK
    End If
End Sub

Private Sub ucMenu_eliminar()
    If ON_ERROR_HABILITADO Then On Error GoTo Fallo
    
    If s2n(NCliente) = 0 Then Exit Sub
    
    If MsgBox("Esta seguro de eliminar este Cliente", vbQuestion + vbYesNo, "Pedido de Confirmacion") = vbYes Then
        DataEnvironment1.Sistema.Execute "UPDATE clientes SET activo=0 WHERE codigo='" & NCliente & "'"
        ucMenu.EliminarOK
    End If
    Exit Sub
Fallo:
    MsgBox "Fallo la Eliminacion", vbInformation, "Error"
End Sub

Private Sub ucMenu_HabilitarEdicion(sino As Boolean)
    Nombre.enabled = sino
    FechaCumpl.enabled = sino
    Empresa.enabled = sino
    Telefono.enabled = sino
    ucCuit.enabled = sino
    ChkMailing.enabled = sino
    mail.enabled = sino
    Pais.enabled = sino
    Calle.enabled = sino
    Localidad.enabled = sino
    CmbProvincia.enabled = sino
    CP.enabled = sino
    CmbTipoCli.enabled = sino
    Puntos.enabled = sino
    CmbTIva.enabled = sino
    Observ.enabled = sino
End Sub

Private Sub ucMenu_Nuevo()
    NCliente = nuevoCodigo("clientes", "codigo", "codigo < 97000") ' perdon
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


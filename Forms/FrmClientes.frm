VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmAbmClientes1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "CLIENTES"
   ClientHeight    =   7740
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11655
   Icon            =   "FrmClientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkConPercIIBB 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Aplicar Perc IIBB"
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
      Left            =   6810
      TabIndex        =   31
      Top             =   3990
      Width           =   1950
   End
   Begin GestionTonka.ucCuit uCuit 
      Height          =   315
      Left            =   6660
      TabIndex        =   3
      Top             =   660
      Width           =   1215
      _extentx        =   2143
      _extenty        =   556
   End
   Begin GestionTonka.ucBotonera ucMenu 
      Align           =   2  'Align Bottom
      Cancel          =   -1  'True
      CausesValidation=   0   'False
      Height          =   795
      Left            =   0
      TabIndex        =   80
      Top             =   6945
      Width           =   11655
      _extentx        =   20558
      _extenty        =   1402
      msgconfirmaeliminar=   ""
      msgconfirmacancelar=   ""
   End
   Begin VB.ComboBox cmbprovincias 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Tag             =   "5"
      Top             =   1500
      Width           =   1695
   End
   Begin VB.CheckBox chkconsig 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Consignatario"
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
      Left            =   4410
      TabIndex        =   29
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CheckBox chkmay 
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
      Height          =   255
      Left            =   4635
      TabIndex        =   30
      Top             =   4275
      Width           =   1575
   End
   Begin VB.TextBox txtlocalidad 
      Height          =   285
      Left            =   7920
      TabIndex        =   7
      Top             =   1080
      Width           =   3405
   End
   Begin VB.TextBox txtcodprov 
      Height          =   285
      Left            =   10200
      TabIndex        =   21
      Top             =   3000
      Width           =   930
   End
   Begin VB.TextBox txtlimite 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7560
      TabIndex        =   20
      Top             =   3000
      Width           =   1050
   End
   Begin VB.CheckBox chketiqueta 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Etiquetas"
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
      Left            =   2460
      TabIndex        =   28
      Top             =   4290
      Width           =   1575
   End
   Begin VB.TextBox txtweb 
      Height          =   285
      Left            =   9000
      TabIndex        =   24
      Top             =   3510
      Width           =   2355
   End
   Begin VB.TextBox txtmail 
      Height          =   285
      Left            =   5310
      TabIndex        =   23
      Top             =   3480
      Width           =   3090
   End
   Begin VB.TextBox txtdescuento2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10200
      MaxLength       =   3
      TabIndex        =   18
      Top             =   2500
      Width           =   1215
   End
   Begin VB.TextBox txtdescuento1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7590
      MaxLength       =   3
      TabIndex        =   17
      Top             =   2500
      Width           =   1005
   End
   Begin VB.TextBox txtdireccioncom 
      Height          =   285
      Left            =   1455
      TabIndex        =   32
      Top             =   5160
      Width           =   3360
   End
   Begin VB.TextBox txtfaxcom 
      Height          =   285
      Left            =   5400
      TabIndex        =   39
      Top             =   5970
      Width           =   2415
   End
   Begin VB.TextBox txttelcom 
      Height          =   285
      Left            =   1440
      TabIndex        =   38
      Top             =   5970
      Width           =   2415
   End
   Begin VB.TextBox txtbarriocom 
      Height          =   285
      Left            =   5400
      TabIndex        =   36
      Top             =   5565
      Width           =   2520
   End
   Begin VB.TextBox txtlocalidadcom 
      Height          =   285
      Left            =   9135
      TabIndex        =   34
      Top             =   5205
      Width           =   2220
   End
   Begin VB.TextBox txtcontactocom 
      Height          =   285
      Left            =   1440
      TabIndex        =   40
      Top             =   6360
      Width           =   4695
   End
   Begin VB.ComboBox cmbzonascom 
      Height          =   315
      Left            =   9120
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   5595
      Width           =   2265
   End
   Begin VB.ComboBox cmbprovinciacom 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   5565
      Width           =   1785
   End
   Begin VB.ComboBox cmblista 
      Height          =   315
      Left            =   4095
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   2500
      Width           =   2130
   End
   Begin VB.TextBox txthorario 
      Height          =   285
      Left            =   9030
      TabIndex        =   41
      Top             =   6315
      Width           =   2355
   End
   Begin VB.CheckBox chkcertificado 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Certificado"
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
      Left            =   480
      TabIndex        =   25
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CheckBox chkhabilitado 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Habilitado"
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
      Left            =   480
      TabIndex        =   26
      Top             =   4305
      Width           =   1575
   End
   Begin VB.CheckBox chkcorreo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "Correo"
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
      Left            =   2685
      TabIndex        =   27
      Top             =   3960
      Width           =   1335
   End
   Begin VB.ComboBox cmbtransporte 
      Height          =   315
      Left            =   9075
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   2000
      Width           =   2310
   End
   Begin VB.ComboBox cmbcategoria 
      Height          =   315
      Left            =   9360
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1500
      Width           =   2025
   End
   Begin VB.TextBox txtdireccion 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   1080
      Width           =   3465
   End
   Begin VB.TextBox txtnombre 
      Height          =   285
      Left            =   3900
      TabIndex        =   1
      Top             =   195
      Width           =   5055
   End
   Begin MSComCtl2.DTPicker dtfecha 
      Height          =   255
      Left            =   9960
      TabIndex        =   42
      Top             =   195
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   51970049
      CurrentDate     =   38052
   End
   Begin VB.ComboBox cmbivas 
      Height          =   315
      ItemData        =   "FrmClientes.frx":08CA
      Left            =   9105
      List            =   "FrmClientes.frx":08D1
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   600
      Width           =   2265
   End
   Begin VB.ComboBox cmbzonas 
      Height          =   315
      Left            =   6480
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1500
      Width           =   1815
   End
   Begin VB.ComboBox cmbvendedores 
      Height          =   315
      Left            =   1425
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2500
      Width           =   1680
   End
   Begin VB.ComboBox cmbformaspagos 
      Height          =   315
      Left            =   2025
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   2940
      Width           =   3615
   End
   Begin VB.TextBox txtcontacto 
      Height          =   285
      Left            =   1440
      TabIndex        =   22
      Top             =   3450
      Width           =   3015
   End
   Begin VB.TextBox txtfantasia 
      Height          =   285
      Left            =   2475
      TabIndex        =   2
      Top             =   640
      Width           =   3525
   End
   Begin VB.TextBox txtbarrio 
      Height          =   285
      Left            =   3840
      TabIndex        =   9
      Top             =   1500
      Width           =   1965
   End
   Begin VB.TextBox txttel 
      Height          =   285
      Left            =   1440
      TabIndex        =   12
      Top             =   2000
      Width           =   2175
   End
   Begin VB.TextBox txtfax 
      Height          =   285
      Left            =   4800
      TabIndex        =   13
      Top             =   2000
      Width           =   1980
   End
   Begin VB.TextBox txtcodigo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1395
      TabIndex        =   0
      Top             =   210
      Width           =   1095
   End
   Begin MSMask.MaskEdBox txtcodpostal 
      Height          =   300
      Left            =   5400
      TabIndex        =   6
      Top             =   1080
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "c"
      Mask            =   "?9999???"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtcodpostalcom 
      Height          =   300
      Left            =   5415
      TabIndex        =   33
      Top             =   5160
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   8
      Format          =   "c"
      Mask            =   "?9999???"
      PromptChar      =   "_"
   End
   Begin VB.Label Label37 
      BackColor       =   &H00C0C0C0&
      Caption         =   "CP:"
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
      Left            =   5055
      TabIndex        =   79
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label Label25 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Localidad:"
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
      Left            =   6840
      TabIndex        =   78
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label36 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Localidad:"
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
      Left            =   8055
      TabIndex        =   77
      Top             =   5205
      Width           =   975
   End
   Begin VB.Label Label35 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cod.Proveedor:"
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
      Left            =   8760
      TabIndex        =   76
      Top             =   3000
      Width           =   1530
   End
   Begin VB.Label Label34 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Limite de Credito:"
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
      Left            =   5880
      TabIndex        =   75
      Top             =   3000
      Width           =   1650
   End
   Begin VB.Label Label33 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Web:"
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
      Left            =   8505
      TabIndex        =   74
      Top             =   3510
      Width           =   615
   End
   Begin VB.Label Label32 
      BackColor       =   &H00C0C0C0&
      Caption         =   "E-Mail :"
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
      Left            =   4635
      TabIndex        =   73
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label31 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Descuento2:"
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
      Left            =   8760
      TabIndex        =   72
      Top             =   2505
      Width           =   1335
   End
   Begin VB.Label Label30 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Descuento1:"
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
      Left            =   6420
      TabIndex        =   71
      Top             =   2505
      Width           =   1335
   End
   Begin VB.Label Label29 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Telefono/s:"
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
      Left            =   390
      TabIndex        =   70
      Top             =   5970
      Width           =   1215
   End
   Begin VB.Label Label28 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Barrio :"
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
      Left            =   4560
      TabIndex        =   69
      Top             =   5565
      Width           =   735
   End
   Begin VB.Label Label27 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Provincia :"
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
      Left            =   480
      TabIndex        =   68
      Top             =   5565
      Width           =   1050
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Direccion :"
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
      Left            =   450
      TabIndex        =   67
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Zona:"
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
      Left            =   8040
      TabIndex        =   66
      Top             =   5595
      Width           =   615
   End
   Begin VB.Label Label23 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fax:"
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
      Left            =   4560
      TabIndex        =   65
      Top             =   5955
      Width           =   495
   End
   Begin VB.Label Label22 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Contacto :"
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
      Left            =   435
      TabIndex        =   64
      Top             =   6405
      Width           =   975
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Lista :"
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
      Left            =   3510
      TabIndex        =   63
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Horario :"
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
      Left            =   8040
      TabIndex        =   62
      Top             =   6315
      Width           =   855
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Datos Comerciales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   480
      TabIndex        =   61
      Top             =   4815
      Width           =   2145
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   1800
      Left            =   90
      Top             =   4935
      Width           =   11460
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Transporte :"
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
      Left            =   7995
      TabIndex        =   60
      Top             =   2000
      Width           =   1095
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Categoria :"
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
      Left            =   8400
      TabIndex        =   59
      Top             =   1500
      Width           =   1095
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Contacto :"
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
      Left            =   435
      TabIndex        =   58
      Top             =   3450
      Width           =   975
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fax:"
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
      Left            =   4365
      TabIndex        =   57
      Top             =   2000
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   4590
      Left            =   120
      Top             =   120
      Width           =   11460
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tipo Iva :"
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
      Left            =   8160
      TabIndex        =   56
      Top             =   640
      Width           =   975
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Zona:"
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
      Left            =   5880
      TabIndex        =   55
      Top             =   1500
      Width           =   645
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0C0C0&
      Caption         =   "CP:"
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
      Left            =   5040
      TabIndex        =   54
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Codigo: "
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
      Left            =   555
      TabIndex        =   53
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Direccion:"
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
      Left            =   435
      TabIndex        =   52
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Razon Social:"
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
      Left            =   2580
      TabIndex        =   51
      Top             =   195
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cuit : "
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
      Left            =   6075
      TabIndex        =   50
      Top             =   660
      Width           =   480
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Forma de Pago:"
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
      Left            =   435
      TabIndex        =   49
      Top             =   2940
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Provincia :"
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
      Left            =   315
      TabIndex        =   48
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha"
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
      Left            =   9240
      TabIndex        =   47
      Top             =   195
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Barrio :"
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
      Left            =   3240
      TabIndex        =   46
      Top             =   1560
      Width           =   645
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Telefono/s:"
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
      Left            =   315
      TabIndex        =   45
      Top             =   2000
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Vendedor :"
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
      Left            =   420
      TabIndex        =   44
      Top             =   2505
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nombre de Fantasia: "
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
      Left            =   435
      TabIndex        =   43
      Top             =   640
      Width           =   2055
   End
End
Attribute VB_Name = "FrmAbmClientes1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' mod 26/10/4 grosso-grosso
' mod 18/10/4
' mod 12/8/4


Private Sub cmbprovincias_Click()
    cmbprovinciacom.ListIndex = cmbprovincias.ListIndex
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True, True
End Sub

Private Sub Form_Load()
   
    CargaCombo cmbzonas, "Zonas", "descripcion", "codigo", ""
    CargaCombo cmbprovincias, "provincias", "descripcion", "codigo", ""
    CargaCombo cmbprovinciacom, "provincias", "descripcion", "codigo", ""
    CargaCombo cmbzonascom, "Zonas", "descripcion", "codigo", ""
    CargaCombo cmbTransporte, "Transportes", "descripcion", "codigo", ""
    CargaCombo cmbvendedores, "usuarios", "descripcion", "codigo", ""
    CargaCombo cmbivas, "Ivas", "descripcion", "codigo", ""
    CargaCombo cmbcategoria, "categorias", "descripcion", "codigo", ""
    CargaCombo cmbformaspagos, "formaspago", "descripcion", "codigo", ""
    CargaCombo cmblista, "Listas", "descripcion", "codigo", ""
    dtfecha = Date
    
    ucMenu.init True, True, True, False, True, "select * from Clientes where activo = 1 order by codigo", DataEnvironment1.AMR
    ucMenu.MsgConfirmaEliminar = "Elimina Cliente ? "
    ucMenu.MsgConfirmaSalir = "Cerrar formulario ? "
       
End Sub


Private Sub IngresoCuit1_CuitInvalido(Nro As String)
    MsgBox "Cuit invalido"
End Sub

Private Sub txtcodigo_LostFocus()
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    Dim rs As New ADODB.Recordset
    

    rs.Open "Select * from clientes where codigo=" & Val(Trim(txtCodigo)), DataEnvironment1.AMR, adOpenStatic, adLockReadOnly
    If Not rs.EOF Then
        MsgBox "El codigo ya existe,verifiquelo", 48, "Atencion"
'        txtCodigo.SetFocus
    End If

fin:
    Set rs = Nothing
    Exit Sub
ufaErr:
    'ufa
    Resume fin
End Sub

Private Sub txtcodpostal_GotFocus()
    frmPintoFoco Me
End Sub

Private Sub txtcodpostalcom_GotFocus()
    frmPintoFoco Me
End Sub

Private Sub txtcodprov_GotFocus()
    frmPintoFoco Me
End Sub

Private Sub txtdescuento1_LostFocus()
    If Val(txtdescuento1) > 100 Then
        MsgBox "El descuento no puede ser superior al 100%", 48, "Atencion"
        txtdescuento1.SetFocus
    End If
End Sub

Private Sub txtdescuento2_LostFocus()
    If Val(txtdescuento2) > 100 Then
        MsgBox "El descuento no puede ser superior al 100%", 48, "Atencion"
        txtdescuento2.SetFocus
    End If
End Sub

Private Sub txtDireccion_GotFocus()
    frmPintoFoco Me
End Sub

Private Sub txtdireccioncom_GotFocus()
    frmPintoFoco Me
End Sub

Private Sub txtfantasia_GotFocus()
    frmPintoFoco Me
End Sub

Private Sub txtfax_GotFocus()
    frmPintoFoco Me
End Sub

Private Sub txtfaxcom_GotFocus()
'    txtfaxcom.SelStart = 0
'    txtfaxcom.SelLength = Len(txtfaxcom.text)
    frmPintoFoco Me
End Sub

Private Sub txthorario_GotFocus()
'    txthorario.SelStart = 0
'    txthorario.SelLength = Len(txthorario.text)
    frmPintoFoco Me
End Sub


Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    Else
        If KeyAscii < 47 Or KeyAscii > 57 Then
            KeyAscii = 0
        End If
'    End If
End Sub

Private Sub txtlimite_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    Else
        If KeyAscii < 47 Or KeyAscii > 57 Then
            KeyAscii = 0
        End If
    End If
End Sub

Sub HabilitoTxt(habilito As Boolean)

    txtCodigo.Locked = habilito
    txtbarrio.Locked = habilito
    txtbarriocom.Locked = habilito
    Txtcontacto.Locked = habilito
    txtlimite.Locked = habilito
    txtcontactocom.Locked = habilito
    txtdescuento1.Locked = habilito
    txtdescuento2.Locked = habilito
    txtDireccion.Locked = habilito
    txtdireccioncom.Locked = habilito
    txtfantasia.Locked = habilito
    txtfax.Locked = habilito
    txtfaxcom.Locked = habilito
    txthorario.Locked = habilito
    txtLocalidad.Locked = habilito
    txtlocalidadcom.Locked = habilito
    txtmail.Locked = habilito
    txtnombre.Locked = habilito
    txttel.Locked = habilito
    txttelcom.Locked = habilito
    txtweb.Locked = habilito
    cmbcategoria.Locked = habilito
    cmbformaspagos.Locked = habilito
    cmbivas.Locked = habilito
    cmblista.Locked = habilito
    cmbprovincias.Locked = habilito
    cmbprovinciacom.Locked = habilito
    cmbvendedores.Locked = habilito
    cmbzonas.Locked = habilito
    cmbzonascom.Locked = habilito
    cmbcategoria.Locked = habilito
    cmbTransporte.Locked = habilito
    chkcertificado.Enabled = Not habilito
    chkhabilitado.Enabled = Not habilito
    chkcorreo.Enabled = Not habilito
    chketiqueta.Enabled = Not habilito
    txtcodpostal.Enabled = Not habilito
    txtcodpostalcom.Enabled = Not habilito
    'MaskCuit.Enabled = Not habilito
    uCuit.Enabled = Not habilito
End Sub
Sub LimpioTxt()
    On Error Resume Next

    txtCodigo = ""
    txtbarrio = ""
    txtbarriocom = ""
    Txtcontacto = ""
    txtcontactocom = ""
    txtdescuento1 = "0.00"
    txtdescuento2 = "0.00"
    txtDireccion = ""
    txtdireccioncom = ""
    txtfantasia = ""
    txtfax = ""
    txtfaxcom = ""
    txthorario = ""
    txtLocalidad = ""
    txtlocalidadcom = ""
    txtmail = ""
    txtnombre = ""
    txtlimite = "0.00"
    txttel = ""
    txttelcom = ""
    txtweb = ""
    cmbcategoria.ListIndex = -1
    cmbformaspagos.ListIndex = 0
    cmbivas.ListIndex = 0
    cmblista.ListIndex = 0
    cmbprovincias.ListIndex = 1
    cmbprovinciacom.ListIndex = 1
    cmbvendedores.ListIndex = 0
    cmbzonas.ListIndex = -1
    cmbzonascom.ListIndex = -1
    cmbTransporte.ListIndex = 0
    chkcertificado.Value = 0
    chkhabilitado.Value = 1
    chkcorreo.Value = 0
    chketiqueta.Value = 0
'    MaskCuit.Mask = "  -       - "
    uCuit.Text = ""
    txtcodpostal.Mask = "       "
    txtcodpostalcom.Mask = "       "
'    MaskCuit.Mask = "99-99999999-9"
    txtcodpostal.Mask = "?9999???"
    txtcodpostalcom.Mask = "?9999???"
    
End Sub

Private Sub txtcodigo_GotFocus()
'    txtcodigo.SelStart = 0
'    txtcodigo.SelLength = Len(txtcodigo.text)
    frmPintoFoco Me
End Sub
Private Sub txtlocalidad_GotFocus()
'    txtlocalidad.SelStart = 0
'    txtlocalidad.SelLength = Len(txtlocalidad.text)
    frmPintoFoco Me
End Sub
Private Sub txtlocalidadcom_GotFocus()
'    txtlocalidadcom.SelStart = 0
'    txtlocalidadcom.SelLength = Len(txtlocalidadcom.text)
    frmPintoFoco Me
End Sub

Private Sub txtlimite_GotFocus()
'    txtlimite.SelStart = 0
'    txtlimite.SelLength = Len(txtlimite.text)
    frmPintoFoco Me
End Sub
'
'Private Sub txtmail_Change()
'Dim i As Integer
'    txtmail.text = UCase(txtmail.text)
'    i = Len(txtmail.text)
'    txtmail.SelStart = i
'End Sub
'
'Private Sub txtnombre_Change()
'Dim i As Integer
'    txtnombre.text = UCase(txtnombre.text)
'    i = Len(txtnombre.text)
'    txtnombre.SelStart = i
'End Sub
'''
'''Private Sub txtnombre_KeyPress(KeyAscii As Integer)
'''    If KeyAscii = 13 Then
'''        SendKeys "{tab}"
'''        KeyAscii = 0
'''    End If
'''End Sub
'''Private Sub Txtmail_KeyPress(KeyAscii As Integer)
'''    If KeyAscii = 13 Then
'''        SendKeys "{tab}"
'''        KeyAscii = 0
'''    End If
'''End Sub

Private Sub txtnombre_LostFocus()
Dim rs As New ADODB.Recordset

    rs.Open "Select * from clientes where descripcion='" & Trim(txtnombre) & "'", DataEnvironment1.AMR, adOpenStatic, adLockReadOnly
    If Not rs.EOF Then
        MsgBox "La razon Social ya existe,es del cliente nro: " & rs!codigo
    End If
End Sub

'Private Sub txtweb_Change()
'Dim i As Integer
'    txtweb.text = UCase(txtweb.text)
'    i = Len(txtweb.text)
'    txtweb.SelStart = i
'End Sub
'''
'''Private Sub Txtweb_KeyPress(KeyAscii As Integer)
'''    If KeyAscii = 13 Then
'''        SendKeys "{tab}"
'''        KeyAscii = 0
'''    End If
'''End Sub
Private Sub txtnombre_GotFocus()
'    txtnombre.SelStart = 0
'    txtnombre.SelLength = Len(txtnombre.text)
    frmPintoFoco Me
End Sub
Private Sub Txtweb_GotFocus()
'    txtweb.SelStart = 0
'    txtweb.SelLength = Len(txtweb.text)
    frmPintoFoco Me
End Sub
Private Sub Txtmail_GotFocus()
'    txtmail.SelStart = 0
'    txtmail.SelLength = Len(txtmail.text)
    frmPintoFoco Me
End Sub
'''Private Sub txttel_KeyPress(KeyAscii As Integer)
'''    If KeyAscii = 13 Then
'''        SendKeys "{tab}"
'''        KeyAscii = 0
'''    End If
'''End Sub
Private Sub txttel_GotFocus()
'    txttel.SelStart = 0
'    txttel.SelLength = Len(txttel.text)
    frmPintoFoco Me
End Sub
'''Private Sub txttelcom_KeyPress(KeyAscii As Integer)
'''    If KeyAscii = 13 Then
'''        SendKeys "{tab}"
'''        KeyAscii = 0
'''    End If
'''End Sub
Private Sub txttelcom_GotFocus()
'    txttelcom.SelStart = 0
'    txttelcom.SelLength = Len(txttelcom.text)
    frmPintoFoco Me
End Sub
'''Private Sub txtcontacto_KeyPress(KeyAscii As Integer)
'''    If KeyAscii = 13 Then
'''        SendKeys "{tab}"
'''        KeyAscii = 0
'''    End If
'''End Sub
Private Sub txtcontacto_GotFocus()
'    txtcontacto.SelStart = 0
'    txtcontacto.SelLength = Len(txtcontacto.text)
    frmPintoFoco Me
End Sub
'''Private Sub txtcontactocom_KeyPress(KeyAscii As Integer)
'''    If KeyAscii = 13 Then
'''        SendKeys "{tab}"
'''        KeyAscii = 0
'''    End If
'''End Sub
Private Sub txtcontactocom_GotFocus()
'    txtcontactocom.SelStart = 0
'    txtcontactocom.SelLength = Len(txtcontactocom.text)
    frmPintoFoco Me
End Sub
''Private Sub txtbarrio_KeyPress(KeyAscii As Integer)
''    If KeyAscii = 13 Then
''        SendKeys "{tab}"
''        KeyAscii = 0
''    End If
''End Sub
Private Sub txtbarrio_GotFocus()
'    txtbarrio.SelStart = 0
'    txtbarrio.SelLength = Len(txtbarrio.text)
    frmPintoFoco Me
End Sub
'Private Sub txtbarriocom_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    End If
'End Sub
Private Sub txtbarriocom_GotFocus()
'    txtbarriocom.SelStart = 0
'    txtbarriocom.SelLength = Len(txtbarriocom.text)
    frmPintoFoco Me
End Sub
'Private Sub txtdescuento1_KeyPress(KeyAscii As Integer)
''    If KeyAscii = 13 Then
''        SendKeys "{tab}"
''        KeyAscii = 0
''    Else
'        If KeyAscii < 47 Or KeyAscii > 57 Then
'            KeyAscii = 0
'        End If
''    End If
'End Sub
Private Sub txtDescuento1_GotFocus()
'    txtdescuento1.SelStart = 0
'    txtdescuento1.SelLength = Len(txtdescuento1.text)
    frmPintoFoco Me
End Sub
'Private Sub txtdescuento2_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        SendKeys "{tab}"
'        KeyAscii = 0
'    Else
'        If KeyAscii < 47 Or KeyAscii > 57 Then
'            KeyAscii = 0
'        End If
'    End If
'End Sub
Private Sub txtDescuento2_GotFocus()
'    txtdescuento2.SelStart = 0
'    txtdescuento2.SelLength = Len(txtdescuento2.text)
    frmPintoFoco Me
End Sub

''Private Sub txtDireccion_KeyPress(KeyAscii As Integer)
''    If KeyAscii = 13 Then
''        SendKeys "{tab}"
''        KeyAscii = 0
''    End If
''End Sub

''Private Sub txtdireccioncom_KeyPress(KeyAscii As Integer)
''    If KeyAscii = 13 Then
''        SendKeys "{tab}"
''        KeyAscii = 0
''    End If
''End Sub
''Private Sub txtfantasia_KeyPress(KeyAscii As Integer)
''    If KeyAscii = 13 Then
''        SendKeys "{tab}"
''        KeyAscii = 0
''    End If
''End Sub
''Private Sub txtfax_KeyPress(KeyAscii As Integer)
''    If KeyAscii = 13 Then
''        SendKeys "{tab}"
''        KeyAscii = 0
''    End If
''End Sub
''Private Sub txtfaxcom_KeyPress(KeyAscii As Integer)
''    If KeyAscii = 13 Then
''        SendKeys "{tab}"
''        KeyAscii = 0
''    End If
''End Sub
''Private Sub txthorario_KeyPress(KeyAscii As Integer)
''    If KeyAscii = 13 Then
''        SendKeys "{tab}"
''        KeyAscii = 0
''    End If
''End Sub
''Sub CargoRegistro()
''
''    txtcodigo = rsCli!Codigo
''    If Not IsNull(rsCli!barrio) Then
''        txtbarrio = rsCli!barrio
''    Else
''        txtbarrio = ""
''    End If
''    If Not IsNull(rsCli!barrio_comercial) Then
''        txtbarriocom = rsCli!barrio_comercial
''    Else
''        txtbarriocom = ""
''    End If
''    If Not IsNull(rsCli!contacto) Then
''        txtcontacto = rsCli!contacto
''    Else
''        txtcontacto = ""
''    End If
''    If Not IsNull(rsCli!contacto_comercial) Then
''        txtcontactocom = rsCli!contacto_comercial
''    Else
''        txtcontactocom = ""
''    End If
''    If Not IsNull(rsCli!descuento2) Then
''        txtdescuento2 = rsCli!descuento2
''    Else
''        txtdescuento2 = "0"
''    End If
''    If Not IsNull(rsCli!direccion) Then
''        txtdireccion = rsCli!direccion
''    Else
''        txtdireccion = ""
''    End If
''    If Not IsNull(rsCli!direccion_comercial) Then
''        txtdireccioncom = rsCli!direccion_comercial
''    Else
''        txtdireccioncom = ""
''    End If
''    If Not IsNull(rsCli!nombrefantasia) Then
''        txtfantasia = rsCli!nombrefantasia
''    Else
''        txtfantasia = ""
''    End If
''    If Not IsNull(rsCli!fax) Then
''        txtfax = rsCli!fax
''    Else
''        txtfax = ""
''    End If
''    If Not IsNull(rsCli!descuento1) Then
''        txtdescuento1 = rsCli!descuento1
''    Else
''        txtdescuento1 = "0"
''    End If
''    If Not IsNull(rsCli!fax_comercial) Then
''        txtfaxcom = rsCli!fax_comercial
''    Else
''        txtfaxcom = ""
''    End If
''    If Not IsNull(rsCli!horario) Then
''        txthorario = rsCli!horario
''    Else
''        txthorario = ""
''    End If
''    If Not IsNull(rsCli!localidad) Then
''        txtlocalidad = rsCli!localidad
''    Else
''        txtlocalidad = ""
''    End If
''    If Not IsNull(rsCli!localidad_comercial) Then
''        txtlocalidadcom = rsCli!localidad_comercial
''    Else
''        txtlocalidadcom = ""
''    End If
''    If Not IsNull(rsCli!mail) Then
''        txtmail = rsCli!mail
''    Else
''        txtmail = ""
''    End If
''    If Not IsNull(rsCli!descripcion) Then
''        txtnombre = rsCli!descripcion
''    Else
''        txtnombre = ""
''    End If
''    If Not IsNull(rsCli!limitecredito) Then
''        txtlimite = rsCli!limitecredito
''    Else
''        txtlimite = "0"
''    End If
''    If Not IsNull(rsCli!telefono) Then
''        txttel = rsCli!telefono
''    Else
''        txttel = ""
''    End If
''    If Not IsNull(rsCli!proveedor) Then
''        txtcodprov = rsCli!proveedor
''    Else
''        txtcodprov = ""
''    End If
''    If Not IsNull(rsCli!telefono_comercial) Then
''        txttelcom = rsCli!telefono_comercial
''    Else
''        txttelcom = ""
''    End If
''    If Not IsNull(rsCli!web) Then
''        txtweb = rsCli!web
''    Else
''        txtweb = ""
''    End If
''    If Not IsNull(rsCli!categoria) Then
''        cmbcategoria.ListIndex = BuscarenComboS(cmbcategoria, ObtenerDescripcion("categorias", rsCli!categoria))
''    Else
''        cmbcategoria.ListIndex = -1
''    End If
''    If Not IsNull(rsCli!formapago) Then
''        cmbformaspagos.ListIndex = BuscarenComboS(cmbformaspagos, ObtenerDescripcion("formaspago", rsCli!formapago))
''    Else
''        cmbformaspagos.ListIndex = -1
''    End If
''    If Not IsNull(rsCli!iva) Then
''        cmbivas.ListIndex = BuscarenComboS(cmbivas, ObtenerDescripcion("ivas", rsCli!iva))
''    Else
''        cmbivas.ListIndex = -1
''    End If
''    If Not IsNull(rsCli!provincia) Then
''        cmbprovincias.ListIndex = BuscarenComboS(cmbprovincias, ObtenerDescripcionS("provincias", rsCli!provincia))
''    Else
''        cmbprovincias.ListIndex = -1
''    End If
''    If Not IsNull(rsCli!provincia_comercial) Then
''        cmbprovinciacom.ListIndex = BuscarenComboS(cmbprovinciacom, ObtenerDescripcionS("provincias", rsCli!provincia_comercial))
''    Else
''        cmbprovinciacom.ListIndex = -1
''    End If
''    If Not IsNull(rsCli!zona) Then
''        cmbzonas.ListIndex = BuscarenComboS(cmbzonas, ObtenerDescripcion("zonas", rsCli!zona))
''    Else
''        cmbzonas.ListIndex = -1
''    End If
''    If Not IsNull(rsCli!transporte) Then
''        cmbtransporte.ListIndex = BuscarenComboS(cmbtransporte, ObtenerDescripcion("transportes", rsCli!transporte))
''    Else
''        cmbtransporte.ListIndex = -1
''    End If
''    If Not IsNull(rsCli!vendedor) Then
''        cmbvendedores.ListIndex = BuscarenComboS(cmbvendedores, ObtenerDescripcion("usuarios", rsCli!vendedor))
''    Else
''        cmbvendedores.ListIndex = -1
''    End If
''    If Not IsNull(rsCli!Lista) Then
''        cmblista.ListIndex = BuscarenComboS(cmblista, ObtenerDescripcion("listas", rsCli!Lista))
''    Else
''        cmblista.ListIndex = -1
''    End If
''    If Not IsNull(rsCli!zonacomercial) Then
''        cmbzonascom.ListIndex = BuscarenComboS(cmbzonascom, ObtenerDescripcion("zonas", rsCli!zonacomercial))
''    Else
''        cmbzonascom.ListIndex = -1
''    End If
''    If Not IsNull(rsCli!codigopostal) Then
''        txtcodpostal = rsCli!codigopostal
''    Else
''        txtcodpostal.Mask = "        "
''        txtcodpostal.Mask = "?9999???"
''    End If
''    If Not IsNull(rsCli!codigopostal_comercial) Then
''        txtcodpostalcom = rsCli!codigopostal_comercial
''    Else
''        txtcodpostalcom.Mask = "        "
''        txtcodpostalcom.Mask = "?9999???"
''    End If
''    If Not IsNull(rsCli!cuit) Then
''        MaskCuit = rsCli!cuit
''    Else
''        MaskCuit.Mask = "        "
''        MaskCuit.Mask = "?9999???"
''    End If
''    If rsCli!Certificado = True Then
''        chkcertificado.Value = 1
''    Else
''        chkcertificado.Value = 0
''    End If
''    If rsCli!consignatario = True Then
''        chkconsig.Value = 1
''    Else
''        chkconsig.Value = 0
''    End If
''    If rsCli!mayorista = True Then
''        chkmay.Value = 1
''    Else
''        chkmay.Value = 0
''    End If
''    If rsCli!Correo = True Then
''        chkcorreo.Value = 1
''    Else
''        chkcorreo.Value = 0
''    End If
''    If rsCli!puedofacturar = True Then
''        chkhabilitado.Value = 1
''    Else
''        chkhabilitado.Value = 0
''    End If
''End Sub
''Public Sub CargarDatos()
''
''    If rsCli.State = 1 Then
''        rsCli.Close
''        Set rsCli = Nothing
''    End If
''    'Codigo = Val(Trim(txtcodigo))
''    rsCli.Open "select * from Clientes where activo = 1", daTaenvironment1.amr, adOpenStatic, adLockOptimistic
''    If Not rsCli.EOF Then
''        rsCli.MoveFirst
''        rsCli.Find "Codigo= " & str(Trim(txtcodigo))
''        CargoRegistro
''        Call HabilitoBotonesMoverse(True, True, True, True)
''        Call HabilitoControles(True, False, True, False, True, False)
''    End If
''End Sub

Sub CargoRegistro()
    On Error Resume Next
    
    LimpioTxt
    With ucMenu.rs
        txtCodigo = !codigo
        txtbarrio = !barrio
        txtbarriocom = !barrio_comercial
        Txtcontacto = !contacto
        txtcontactocom = !contacto_comercial
        txtDireccion = !direccion
        txtdireccioncom = !direccion_comercial
        txtfantasia = !nombrefantasia
        txtfax = !fax
        txtdescuento1 = s2n(!descuento1)
        txtdescuento2 = s2n(!descuento2)
        txtlimite = s2n(!limitecredito)
        txtfaxcom = !fax_comercial
        txthorario = !horario
        txtLocalidad = !localidad
        txtlocalidadcom = !localidad_comercial
        txtmail = !mail
        txtnombre = !descripcion
        txttel = !telefono
        TxtCodProv = !Proveedor
        txttelcom = !telefono_comercial
        txtweb = !web
        cmbcategoria.ListIndex = BuscarenComboS(cmbcategoria, ObtenerDescripcion("categorias", !categoria))
        cmbformaspagos.ListIndex = BuscarenComboS(cmbformaspagos, ObtenerDescripcion("formaspago", !formapago))
        cmbivas.ListIndex = BuscarenComboS(cmbivas, ObtenerDescripcion("ivas", !iva))
        cmbprovincias.ListIndex = BuscarenComboS(cmbprovincias, ObtenerDescripcionS("provincias", !provincia))
        cmbprovinciacom.ListIndex = BuscarenComboS(cmbprovinciacom, ObtenerDescripcionS("provincias", !provincia_comercial))
        cmbzonas.ListIndex = BuscarenComboS(cmbzonas, ObtenerDescripcion("zonas", !zona))
        cmbTransporte.ListIndex = BuscarenComboS(cmbTransporte, ObtenerDescripcion("transportes", !Transporte))
        cmbvendedores.ListIndex = BuscarenComboS(cmbvendedores, ObtenerDescripcion("usuarios", !Vendedor))
        cmblista.ListIndex = BuscarenComboS(cmblista, ObtenerDescripcion("listas", !Lista))
        cmbzonascom.ListIndex = BuscarenComboS(cmbzonascom, ObtenerDescripcion("zonas", !zonacomercial))
        txtcodpostal = !codigopostal
        txtcodpostalcom = !codigopostal_comercial
'        MaskCuit = !cuit
        uCuit.Text = !Cuit
        chkcertificado.Value = b2k(!Certificado)
        chkconsig.Value = b2k(!consignatario)
        chkmay.Value = b2k(!mayorista)
        chkcorreo.Value = b2k(!Correo)
        chkhabilitado.Value = b2k(!puedofacturar)
        chkConPercIIBB.Value = b2k(!ConPercIIBB)
        chketiqueta.Value = b2k(!etiqueta)
    End With
End Sub

Private Function FaltanCosas() As Boolean
    Dim tmp, s As String, iv As Long

    'nombre, codigo
    If s2n(txtCodigo) = 0 Or Trim(txtnombre) = "" Then
        che "falta cargar codigo y nombre cliente"
        
        FaltanCosas = True
        Exit Function
    End If
    
    
    iv = cmbivas.ListIndex
    If iv <> 0 And iv <> 1 And iv <> 2 Then ' 0 = CONSUMIDOR FINAL, 1 y 2= EXENTOS
        If uCuit.Text = "" Then
            che "CUIT Incorrecto"
            FaltanCosas = True
            Exit Function
        End If
    
        s = "select codigo from clientes where activo = 1 and cuit = '" & uCuit.Text & "' and codigo <> " & s2n(txtCodigo)
        tmp = obtenerDeSQL(s)
        If s2n(tmp) > 0 Then
            che "El cuit ya existe en el cliente nro: " & tmp
            FaltanCosas = True
            Exit Function
        End If
    End If

End Function

Private Sub GrabarCliente(Ope As String)
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr

    Dim Correo As Long
    Dim Puedo As Long
    Dim Certificado As Long
    Dim prov As Long
    Dim consig As Long
    Dim mayor As Long
    Dim etiqueta As Long
    Dim limite As Double
    Dim ConPercIIBB As Long
            
    If FaltanCosas Then Exit Sub
            
    Correo = k2b(chkcorreo.Value)
    Puedo = k2b(chkhabilitado.Value)
    Certificado = k2b(chkcertificado.Value)
    consig = k2b(chkconsig.Value)
    mayor = k2b(chkmay.Value)
    etiqueta = k2b(chketiqueta.Value)
    limite = s2n(txtlimite)
    ConPercIIBB = k2b(chkConPercIIBB.Value)

    If Ope = "A" Then
        'DataEnvironment1.dbo_CLIENTE "A", Val(Trim(txtCodigo)), Trim(txtnombre), _
            txtcodpostal, txtcodpostalcom, Correo, Puedo, Trim(txtfantasia), Trim(txtDireccion), Trim(txtLocalidad), _
            Trim(txtbarrio), ObtenerCodigoS("provincias", Trim(cmbprovincias.Text)), _
            uCuit.Text, Val(Trim(txtcodprov)), Trim(txttel), Trim(txtfax), Trim(Txtcontacto), ObtenerCodigo("usuarios", Trim(cmbvendedores.Text)), _
            ObtenerCodigo("ivas", Trim(cmbivas.Text)), ObtenerCodigo("formaspago", Trim(cmbformaspagos.Text)), _
            ObtenerCodigo("zonas", Trim(cmbzonas.Text)), Trim(txtdireccioncom), _
            Trim(txtlocalidadcom), Trim(txtbarriocom), ObtenerCodigoS("provincias", Trim(cmbprovinciacom.Text)), _
            ObtenerCodigo("zonas", Trim(cmbzonascom.Text)), CDbl(Replace(txtdescuento1, ".", ",")), _
            CDbl(Replace(txtdescuento2, ".", ",")), ObtenerCodigo("Listas", Trim(cmblista.Text)), _
            Trim(txthorario), Trim(txtfaxcom), Trim(txttelcom), _
            Trim(txtcontactocom), ObtenerCodigo("categorias", Trim(cmbcategoria.Text)), _
            Certificado, ObtenerCodigo("transportes", Trim(cmbTransporte.Text)), _
            limite, Trim(txtweb), Trim(txtmail), consig, mayor, Date, UsuarioSistema!codigo, 0, 0
        DataEnvironment1.dbo_CLIENTE "A", s2n(txtCodigo), Trim(txtnombre), _
            txtcodpostal, txtcodpostalcom, Correo, Puedo, Trim(txtfantasia), Trim(txtDireccion), Trim(txtLocalidad), _
            Trim(txtbarrio), ObtenerCodigoS("provincias", Trim(cmbprovincias.Text)), _
            uCuit.Text, Val(Trim(TxtCodProv)), Trim(txttel), Trim(txtfax), Trim(Txtcontacto), ObtenerCodigo("usuarios", Trim(cmbvendedores.Text)), _
            ObtenerCodigo("ivas", Trim(cmbivas.Text)), ObtenerCodigo("formaspago", Trim(cmbformaspagos.Text)), _
            ObtenerCodigo("zonas", Trim(cmbzonas.Text)), Trim(txtdireccioncom), _
            Trim(txtlocalidadcom), Trim(txtbarriocom), ObtenerCodigoS("provincias", Trim(cmbprovinciacom.Text)), _
            ObtenerCodigo("zonas", Trim(cmbzonascom.Text)), s2n(txtdescuento1), _
            s2n(txtdescuento2), ObtenerCodigo("Listas", Trim(cmblista.Text)), _
            Trim(txthorario), Trim(txtfaxcom), Trim(txttelcom), _
            Trim(txtcontactocom), ObtenerCodigo("categorias", Trim(cmbcategoria.Text)), _
            Certificado, ObtenerCodigo("transportes", Trim(cmbTransporte.Text)), _
            limite, Trim(txtweb), Trim(txtmail), consig, mayor, _
            ConPercIIBB, etiqueta, _
            Date, UsuarioSistema!codigo
        
    ElseIf Ope = "M" Then
        'DataEnvironment1.dbo_CLIENTE "M", Val(Trim(txtCodigo)), Trim(txtnombre), _
            txtcodpostal, txtcodpostalcom, Correo, Puedo, Trim(txtfantasia), Trim(txtDireccion), Trim(txtLocalidad), _
            Trim(txtbarrio), ObtenerCodigoS("provincias", Trim(cmbprovincias.Text)), _
            uCuit.Text, Val(Trim(txtcodprov)), Trim(txttel), Trim(txtfax), Trim(Txtcontacto), ObtenerCodigo("usuarios", Trim(cmbvendedores.Text)), _
            ObtenerCodigo("ivas", Trim(cmbivas.Text)), ObtenerCodigo("formaspago", Trim(cmbformaspagos.Text)), _
            ObtenerCodigo("zonas", Trim(cmbzonas.Text)), Trim(txtdireccioncom), _
            Trim(txtlocalidadcom), Trim(txtbarriocom), ObtenerCodigoS("provincias", Trim(cmbprovinciacom.Text)), _
            ObtenerCodigo("zonas", Trim(cmbzonascom.Text)), CDbl(Replace(txtdescuento1, ".", ",")), _
            CDbl(Replace(txtdescuento2, ".", ",")), ObtenerCodigo("Listas", Trim(cmblista.Text)), _
            Trim(txthorario), Trim(txtfaxcom), Trim(txttelcom), _
            Trim(txtcontactocom), ObtenerCodigo("categorias", Trim(cmbcategoria.Text)), _
            Certificado, ObtenerCodigo("transportes", Trim(cmbTransporte.Text)), _
            limite, Trim(txtweb), Trim(txtmail), consig, mayor, Date, UsuarioSistema!codigo, 0, 0
        DataEnvironment1.dbo_CLIENTE "M", s2n(txtCodigo), Trim(txtnombre), _
            txtcodpostal, txtcodpostalcom, Correo, Puedo, Trim(txtfantasia), Trim(txtDireccion), Trim(txtLocalidad), _
            Trim(txtbarrio), ObtenerCodigoS("provincias", Trim(cmbprovincias.Text)), _
            uCuit.Text, s2n(TxtCodProv), Trim(txttel), Trim(txtfax), Trim(Txtcontacto), ObtenerCodigo("usuarios", Trim(cmbvendedores.Text)), _
            ObtenerCodigo("ivas", Trim(cmbivas.Text)), ObtenerCodigo("formaspago", Trim(cmbformaspagos.Text)), _
            ObtenerCodigo("zonas", Trim(cmbzonas.Text)), Trim(txtdireccioncom), _
            Trim(txtlocalidadcom), Trim(txtbarriocom), ObtenerCodigoS("provincias", Trim(cmbprovinciacom.Text)), _
            ObtenerCodigo("zonas", Trim(cmbzonascom.Text)), s2n(txtdescuento1), _
            s2n(txtdescuento2), ObtenerCodigo("Listas", Trim(cmblista.Text)), _
            Trim(txthorario), Trim(txtfaxcom), Trim(txttelcom), _
            Trim(txtcontactocom), ObtenerCodigo("categorias", Trim(cmbcategoria.Text)), _
            Certificado, ObtenerCodigo("transportes", Trim(cmbTransporte.Text)), _
            limite, Trim(txtweb), Trim(txtmail), consig, mayor, _
            ConPercIIBB, etiqueta, _
            Date, UsuarioSistema!codigo
        grabaBitacora "M", s2n(txtCodigo), "Clientes"
    End If
    MsgBox "La Operacion se ha realizado con éxito", 48, "Atencion"
    ucMenu.AceptarOk "codigo = " & txtCodigo

fin:
    Exit Sub
ufaErr:
    ufa "Err al grabar", Me.Name & " " & txtCodigo ', Err
    Resume fin
End Sub


'*----------------------------- MENU ---------------------------------
Private Sub ucMenu_AceptarAlta()
    GrabarCliente "A"
End Sub
Private Sub ucMenu_AceptarModi()
    GrabarCliente "M"
End Sub
Private Sub ucMenu_BorrarControles()
    LimpioTxt
End Sub
Private Sub ucMenu_Buscar()
    Dim resu As String
    resu = frmBuscar.MostrarCodigoDescripcionActivo("Clientes")
    If resu > "" Then
        txtCodigo = resu
        ucMenu.BuscarOK "codigo = " & txtCodigo
        CargoRegistro
    End If
End Sub
Private Sub ucMenu_eliminar()
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    DataEnvironment1.dbo_CLIENTE "B", s2n(txtCodigo), "", "", "", 0, 0, "", "", "", "", "", "", 0, "", "", "", 0, 0, 0, 0, "", "", "", "", 0, 0, 0, 0, "", "", "", "", 0, 0, 0, 0, "", "", 0, 0, 0, 0, Val(UsuarioSistema!codigo), Date
    grabaBitacora "B", s2n(txtCodigo), "Clientes"
    MsgBox "La Operacion se ha realizado con exito", 48, "Atencion"
    ucMenu.EliminarOK
fin:
    Exit Sub
ufaErr:
    ufa "Err al eliminar", Me.Name & " " & s2n(txtCodigo) ', Err
    Resume fin
End Sub
Private Sub ucMenu_HabilitarEdicion(sino As Boolean)
    HabilitoTxt Not sino ' ta'reves
End Sub

Private Sub ucMenu_Modificar()
    txtCodigo.Enabled = False
End Sub

Private Sub ucMenu_nuevo()
    On Error Resume Next
    txtCodigo = nuevoCodigo("Clientes")
    txtCodigo.SetFocus
End Sub
Private Sub ucMenu_SALIR()
    Unload Me
End Sub
Private Sub ucMenu_SeMovio()
    CargoRegistro
End Sub
'*----------------------------- MENU ---------------------------------


' 12/8/4 Lito
'   busq  frmHelp  pasa a frmBUSCAR
'   inhibo mov al cargar
' 17/8/4
'   codigo en CargarDatos() ? lo saque
' 18-10-4 from lorena-default date
' 30/11/4 los change x mayusculas, key press indiv x reempl enter x tab
'       se hacen ahora en 1 sola linea frmKeyPress
'       los gotfocus se reempl x instr con 1 solo parametro, el form


Private Sub uCuit_GotFocus()
    PintoFocoActivo
End Sub

'Private Sub uCuit_LostFocus()
'    Dim tmp, s
'
'    s = "select codigo from clientes where activo = 1 and cuit = '" & uCuit.Text & "' "
'    tmp = obtenerDeSQL(s)
'
'    If Not IsEmpty(tmp) And tmp <> s2n(txtcodigo) Then
'        che "El cuit ya existe,es del cliente nro: " & tmp
'    End If
'
''    Dim rs As New ADODB.Recordset
''
''    rs.Open "Select * from clientes where cuit='" & MaskCuit & "'", daTaenvironment1.amr, adOpenStatic, adLockReadOnly
''    If Not rs.EOF Then
''        MsgBox "El cuit ya existe,es del cliente nro: " & rs!Codigo
''        MaskCuit.SetFocus
''    End If
'
'End Sub

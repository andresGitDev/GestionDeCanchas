VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmFactDeta 
   Caption         =   "Actualizacion de detalle"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11505
   Icon            =   "frmFactDeta.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   11505
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCabecera 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11415
      Begin VB.TextBox TxtNroFactura 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1980
         TabIndex        =   22
         Top             =   465
         Width           =   1305
      End
      Begin VB.ComboBox cmbCliente 
         Height          =   315
         Left            =   2775
         TabIndex        =   21
         Top             =   810
         Width           =   3735
      End
      Begin VB.TextBox txtCodCliente 
         Alignment       =   1  'Right Justify
         Height          =   320
         Left            =   1380
         TabIndex        =   20
         Top             =   810
         Width           =   975
      End
      Begin VB.TextBox txtLocalidad 
         Alignment       =   1  'Right Justify
         Height          =   320
         Left            =   1380
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1560
         Width           =   2235
      End
      Begin VB.TextBox txtDireccion 
         Alignment       =   1  'Right Justify
         Height          =   320
         Left            =   1380
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1200
         Width           =   5115
      End
      Begin VB.ComboBox cmbProvincia 
         Height          =   315
         Left            =   1380
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1920
         Width           =   2235
      End
      Begin VB.TextBox txtTipoDoc 
         Height          =   315
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   465
         Width           =   495
      End
      Begin VB.ComboBox cmbFormaPago 
         Height          =   315
         Left            =   7410
         TabIndex        =   15
         Top             =   450
         Width           =   2655
      End
      Begin VB.ComboBox cmbTipoIva 
         Height          =   315
         Left            =   7425
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   810
         Width           =   2610
      End
      Begin VB.ComboBox cmbVendedor 
         Height          =   315
         Left            =   7440
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1185
         Width           =   2625
      End
      Begin VB.CheckBox chkPropio 
         Alignment       =   1  'Right Justify
         Caption         =   "Codigo Propio"
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
         Left            =   3840
         TabIndex        =   12
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1635
      End
      Begin VB.ComboBox cmbDeposito 
         Height          =   315
         ItemData        =   "frmFactDeta.frx":08CA
         Left            =   6480
         List            =   "frmFactDeta.frx":08CC
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1890
         Width           =   1275
      End
      Begin VB.TextBox txtTipoDocRef 
         Height          =   315
         Left            =   4740
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   465
         Width           =   495
      End
      Begin VB.TextBox txtNroFacturaRef 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5280
         TabIndex        =   9
         Top             =   465
         Width           =   1245
      End
      Begin VB.CommandButton cmdContado 
         Caption         =   "Ingr Contado"
         Height          =   330
         Left            =   10110
         TabIndex        =   8
         Top             =   435
         Width           =   1215
      End
      Begin VB.CommandButton cmdCliente 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2355
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   810
         Width           =   375
      End
      Begin VB.CommandButton cmdRemitosPendientes 
         Caption         =   "Elegir Remito Pend"
         Height          =   315
         Left            =   3720
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1575
      End
      Begin VB.OptionButton optCuentaContado 
         Caption         =   "Cta Cte"
         Height          =   300
         Index           =   0
         Left            =   1530
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   60
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton optCuentaContado 
         Caption         =   "Contado"
         Height          =   300
         Index           =   1
         Left            =   2460
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   40
         Width           =   840
      End
      Begin VB.CommandButton cmdPedidosPendientes 
         Caption         =   "Elegir Pedido Pend"
         Height          =   315
         Left            =   3840
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1920
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         Left            =   9630
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox txtCotizacion 
         Alignment       =   1  'Right Justify
         Height          =   320
         Left            =   9645
         TabIndex        =   1
         Top             =   1905
         Width           =   1680
      End
      Begin Gestion.ucCuit ucCuit 
         Height          =   315
         Left            =   10095
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   825
         Width           =   1230
         _ExtentX        =   2355
         _ExtentY        =   556
      End
      Begin MSComCtl2.DTPicker dtFecha 
         Height          =   315
         Left            =   5220
         TabIndex        =   24
         Top             =   45
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   73334785
         CurrentDate     =   38126
      End
      Begin MSComCtl2.DTPicker dtVencimiento 
         Height          =   315
         Left            =   7440
         TabIndex        =   25
         Top             =   60
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Format          =   73334785
         CurrentDate     =   38229
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
         ForeColor       =   &H00400000&
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   42
         Top             =   840
         Width           =   915
      End
      Begin VB.Label Label3 
         Caption         =   "Comprobante:"
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
         Left            =   120
         TabIndex        =   41
         Top             =   495
         Width           =   1215
      End
      Begin VB.Label Label4 
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
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   4500
         TabIndex        =   40
         Top             =   75
         Width           =   735
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   360
         TabIndex        =   39
         Top             =   1260
         Width           =   915
      End
      Begin VB.Label Label5 
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
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   360
         TabIndex        =   38
         Top             =   1590
         Width           =   915
      End
      Begin VB.Label Label6 
         Caption         =   "Provincia:"
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
         Left            =   360
         TabIndex        =   37
         Top             =   1950
         Width           =   915
      End
      Begin VB.Label Label7 
         Caption         =   "iva/cuit:"
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
         Left            =   6675
         TabIndex        =   36
         Top             =   885
         Width           =   765
      End
      Begin VB.Label Label9 
         Caption         =   "Codigo:"
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
         Left            =   9660
         TabIndex        =   35
         Top             =   75
         Width           =   915
      End
      Begin VB.Label Label13 
         Caption         =   "Venc:"
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
         Left            =   6690
         TabIndex        =   34
         Top             =   90
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "Pago:"
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
         Left            =   6675
         TabIndex        =   33
         Top             =   495
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "Vendor:"
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
         Left            =   6705
         TabIndex        =   32
         Top             =   1215
         Width           =   975
      End
      Begin VB.Label lblDepot 
         Caption         =   "Deposito:"
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
         Left            =   5580
         TabIndex        =   31
         Top             =   1950
         Width           =   1035
      End
      Begin VB.Label lblRef 
         Caption         =   "SobreFactura:"
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
         Left            =   3360
         TabIndex        =   30
         Top             =   495
         Width           =   1335
      End
      Begin VB.Label txtCodigo 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   10440
         TabIndex        =   29
         Top             =   45
         Width           =   795
      End
      Begin VB.Label lblExterior 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "EXTERIOR"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   15
         TabIndex        =   28
         Top             =   15
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label Label21 
         Caption         =   "Moneda:"
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
         Left            =   8715
         TabIndex        =   27
         Top             =   1635
         Width           =   900
      End
      Begin VB.Label Label8 
         Caption         =   "Cotizacion:"
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
         Left            =   8550
         TabIndex        =   26
         Top             =   1935
         Width           =   975
      End
   End
   Begin Gestion.ucBotonera ucBoton 
      Align           =   2  'Align Bottom
      Cancel          =   -1  'True
      Height          =   1635
      Left            =   0
      TabIndex        =   43
      Top             =   7245
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   2884
      MsgConfirmaEliminar=   ""
      MsgConfirmaCancelar=   ""
      Begin VB.Frame fraBuscar 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   195
         TabIndex        =   44
         Top             =   30
         Width           =   6015
         Begin VB.OptionButton optBuscarTipo 
            Caption         =   "Fact E"
            Height          =   255
            Index           =   2
            Left            =   2040
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   30
            Width           =   705
         End
         Begin VB.OptionButton optBuscarTipo 
            Caption         =   "Fact A"
            Height          =   255
            Index           =   0
            Left            =   525
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   30
            Width           =   735
         End
         Begin VB.OptionButton optBuscarTipo 
            Caption         =   "Fact B"
            Height          =   255
            Index           =   1
            Left            =   1305
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   30
            Width           =   705
         End
         Begin Gestion.ucEntreFechas ucFechas 
            Height          =   360
            Left            =   3360
            TabIndex        =   48
            Top             =   -15
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   635
         End
         Begin VB.Label Label12 
            Caption         =   "Entre:"
            Height          =   195
            Index           =   0
            Left            =   2910
            TabIndex        =   50
            Top             =   60
            Width           =   495
         End
         Begin VB.Label Label12 
            Caption         =   "Buscar:"
            Height          =   195
            Index           =   1
            Left            =   -30
            TabIndex        =   49
            Top             =   45
            Width           =   555
         End
      End
      Begin VB.Label lblFacturaB 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Factura B"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000017&
         Height          =   555
         Left            =   8610
         TabIndex        =   51
         Top             =   135
         Visible         =   0   'False
         Width           =   2700
      End
   End
   Begin TabDlg.SSTab TabDetalle 
      Height          =   4845
      Left            =   0
      TabIndex        =   52
      Top             =   2295
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   8546
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Detalle"
      TabPicture(0)   =   "frmFactDeta.frx":08CE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label16"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label17"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label18(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label19"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtNeto"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtTotal"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtIva"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtDescuento"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblIIBB"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label18(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblSubtotal"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label22"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "grilla"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "fraEditDetalle"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtPdescuento"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtPIVA"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "fraOptStock"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Origen"
      TabPicture(1)   =   "frmFactDeta.frx":08EA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grillaOrigen"
      Tab(1).Control(1)=   "cmdOrigen"
      Tab(1).ControlCount=   2
      Begin VB.Frame fraOptStock 
         Caption         =   " Actualiza Stock "
         Height          =   1005
         Left            =   9420
         TabIndex        =   66
         Top             =   1500
         Width           =   1755
         Begin VB.TextBox TxtRemitoNumero 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   180
            TabIndex        =   70
            Top             =   1140
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.OptionButton optStock 
            Caption         =   " "
            Height          =   255
            Index           =   0
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   69
            Top             =   330
            Value           =   -1  'True
            Width           =   315
         End
         Begin VB.OptionButton optStock 
            Caption         =   " "
            Height          =   255
            Index           =   1
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   68
            Top             =   630
            Width           =   315
         End
         Begin VB.OptionButton optStock 
            Caption         =   " "
            Height          =   255
            Index           =   2
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   67
            Top             =   1005
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label Label1 
            Caption         =   "No"
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   73
            Top             =   330
            Width           =   675
         End
         Begin VB.Label Label1 
            Caption         =   "Si"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   72
            Top             =   630
            Width           =   795
         End
         Begin VB.Label Label1 
            Caption         =   "Genera Remito"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   71
            Top             =   1005
            Visible         =   0   'False
            Width           =   1155
         End
      End
      Begin VB.TextBox txtPIVA 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   10665
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   4125
         Width           =   675
      End
      Begin VB.TextBox txtPdescuento 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   10665
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   3225
         Width           =   675
      End
      Begin VB.Frame fraEditDetalle 
         Height          =   975
         Left            =   105
         TabIndex        =   54
         Top             =   330
         Width           =   10215
         Begin VB.TextBox txtIvaProducto 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   320
            Left            =   7620
            TabIndex        =   59
            Top             =   540
            Width           =   1095
         End
         Begin VB.TextBox txtPrecio 
            Alignment       =   1  'Right Justify
            Height          =   320
            Left            =   5100
            TabIndex        =   58
            Top             =   540
            Width           =   1155
         End
         Begin VB.CommandButton cmdBorrarItem 
            BackColor       =   &H00C0C0C0&
            Height          =   495
            Left            =   9480
            MaskColor       =   &H00E0E0E0&
            Picture         =   "frmFactDeta.frx":0906
            Style           =   1  'Graphical
            TabIndex        =   57
            TabStop         =   0   'False
            ToolTipText     =   "Borrar Item"
            Top             =   375
            Width           =   495
         End
         Begin VB.CommandButton cmdAgregarItem 
            BackColor       =   &H00E0E0E0&
            Height          =   495
            Left            =   8895
            MaskColor       =   &H00E0E0E0&
            Picture         =   "frmFactDeta.frx":0C10
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   375
            Width           =   495
         End
         Begin VB.TextBox txtCantidad 
            Alignment       =   1  'Right Justify
            Height          =   320
            Left            =   3120
            TabIndex        =   55
            Top             =   540
            Width           =   1095
         End
         Begin Gestion.ucCoDe uProd 
            Height          =   315
            Left            =   240
            TabIndex        =   60
            Top             =   180
            Width           =   8475
            _ExtentX        =   14949
            _ExtentY        =   556
            CodigoWidth     =   1000
         End
         Begin VB.Label lblIvaProducto 
            Caption         =   "IVA Producto:"
            Height          =   315
            Left            =   6540
            TabIndex        =   63
            Top             =   540
            Width           =   1035
         End
         Begin VB.Label Label11 
            Caption         =   "Precio:"
            Height          =   315
            Left            =   4500
            TabIndex        =   62
            Top             =   540
            Width           =   555
         End
         Begin VB.Label Label10 
            Caption         =   "Cantidad :"
            Height          =   255
            Left            =   2340
            TabIndex        =   61
            Top             =   570
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdOrigen 
         Caption         =   "Traer Items Pendientes"
         Height          =   315
         Left            =   -74640
         TabIndex        =   53
         Top             =   480
         Width           =   1875
      End
      Begin VSFlex7LCtl.VSFlexGrid grillaOrigen 
         Height          =   3360
         Left            =   -74760
         TabIndex        =   74
         Top             =   960
         Width           =   9135
         _cx             =   16113
         _cy             =   5927
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VSFlex7LCtl.VSFlexGrid grilla 
         Height          =   3225
         Left            =   180
         TabIndex        =   75
         Top             =   1380
         Width           =   8715
         _cx             =   15372
         _cy             =   5689
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VB.Label Label22 
         Caption         =   "Subtot:"
         Height          =   255
         Left            =   8910
         TabIndex        =   87
         Top             =   2925
         Width           =   495
      End
      Begin VB.Label lblSubtotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   9480
         TabIndex        =   86
         Top             =   2925
         Width           =   1155
      End
      Begin VB.Label Label18 
         Caption         =   "iibb:"
         Height          =   255
         Index           =   1
         Left            =   9120
         TabIndex        =   85
         Top             =   3870
         Width           =   240
      End
      Begin VB.Label lblIIBB 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   9480
         TabIndex        =   84
         Top             =   3825
         Width           =   1155
      End
      Begin VB.Label txtDescuento 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   9465
         TabIndex        =   83
         Top             =   3225
         Width           =   1155
      End
      Begin VB.Label txtIva 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   9465
         TabIndex        =   82
         Top             =   4125
         Width           =   1155
      End
      Begin VB.Label txtTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   9465
         TabIndex        =   81
         Top             =   4440
         Width           =   1155
      End
      Begin VB.Label txtNeto 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   9465
         TabIndex        =   80
         Top             =   3510
         Width           =   1155
      End
      Begin VB.Label Label19 
         Caption         =   "Desc:"
         Height          =   255
         Left            =   8985
         TabIndex        =   79
         Top             =   3225
         Width           =   375
      End
      Begin VB.Label Label18 
         Caption         =   "Neto:"
         Height          =   255
         Index           =   0
         Left            =   9045
         TabIndex        =   78
         Top             =   3510
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "Iva:"
         Height          =   255
         Left            =   9105
         TabIndex        =   77
         Top             =   4125
         Width           =   255
      End
      Begin VB.Label Label16 
         Caption         =   "Total:"
         Height          =   255
         Left            =   8985
         TabIndex        =   76
         Top             =   4440
         Width           =   435
      End
   End
End
Attribute VB_Name = "frmFactDeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const BS_PED_A_FAC_PREGUNTA_CANT = False
Private Const BS_RESPETAR_PRECIO_OC = True
Private midDoc As Long

Private mClienteConIva As Boolean
Private mPropio As String
Private mValoresOK As Boolean
Private mNuevo As Boolean

'tipo factura, llamo con .mostrar() desde menu
Private mFac As fv_FacVentaSobre2
Private mFAE As Boolean
'
Public Enum fv_FacVentaSobre2
    FacturaVenta_Remito2 ' no se usa
    FacturaVenta_Pedido2
    FacturaVenta_Libre2
    FacturaVenta_NCreditoDevolucion2
End Enum

Private Enum ActuStock
    ActuStock_NO = 0
    ActuStock_SI = 1
    ActuStock_RE = 2
End Enum

'Tabla facturas ventas , pero para busq, porq uso el SP
Private Const mTablaFV = "facturaventa"
Private Const StringCONTADO = "CONTADO"

Private WithEvents g As LiGrilla, WithEvents gO As LiGrilla ', gS As LiGrilla
Attribute g.VB_VarHelpID = -1
Attribute gO.VB_VarHelpID = -1
Private WithEvents cliente As LiCodigo ' cliente
Attribute cliente.VB_VarHelpID = -1

'Detalle
Private gCANT   As Long
'Private gALIA   As Long
Private gprod   As Long
Private gDESC   As Long
Private gPUNI   As Long ' precio uni
Private gPTOT   As Long ' precio tot
Private gNPED   As Long ' pedido
Private gNREM   As Long ' remito
Private gFORM   As Long ' formula
Private gITEM   As Long ' item pedido o remito detalle
Private gIVA    As Long
Private gNPCL   As Long ' Nro Pedido clie

'Origen Pedido-Remito
Private gO_PROD As Long
Private gO_DESC As Long
Private gO_CANT As Long
Private gO_NPED As Long ' nro pedido
Private gO_NREM As Long ' nro remito
Private gO_PREC As Long ' precio
Private gO_PEND As Long
Private gO_PROP As Long
Private gO_ITEM As Long ' item remito o pedido detalle
Private g0_NPCL As Long ' Nro Pedido clie

Public Sub mostrar(FacturarSobre As fv_FacVentaSobre2, Optional FacturaExterior As Boolean = False)
    mFac = FacturarSobre
    mFAE = FacturaExterior
    lblExterior.Visible = mFAE
    
    Select Case mFac
    Case FacturaVenta_Libre2
        verCampos False, True, True, 1500, 3000, True, "Factura Venta", True, IIf(gEMPR_EmiteFacturaConRemito, ActuStock_RE, ActuStock_SI), False, False, False
        
        If gEMPR_EmiteFacturaConRemito Then
            optStock(ActuStock_RE).Value = True ' hace remito
        Else
            optStock(ActuStock_SI).Value = True ' mod stock
        End If
        fraOptStock.Visible = False
    
    Case FacturaVenta_Pedido2
        verCampos True, False, True, 1500, 3000, True, "Factura Venta - SOBRE PEDIDO", True, IIf(gEMPR_EmiteFacturaConRemito, ActuStock_RE, ActuStock_SI), False, False, True
        
        If gEMPR_EmiteFacturaConRemito Then
            optStock(ActuStock_RE).Value = True ' hace remito
        Else
            optStock(ActuStock_NO).Value = True ' mod stock
        End If
    
    Case FacturaVenta_Remito2
        verCampos True, True, True, 1500, 3000, False, "Factura Venta - SOBRE REMITO", False, ActuStock_NO, False, True, False

        optStock(ActuStock_NO).Value = True ' Sin modif stock
        fraOptStock.Visible = False
        
    Case FacturaVenta_NCreditoDevolucion2
        verCampos False, True, True, 1500, 3000, True, "Nota de Credito POR DEVOLUCION", True, ActuStock_SI, True, False, False
        optCuentaContado.Item(0).Value = True
        cmdContado.enabled = False
        optCuentaContado.Item(0).enabled = False
        optCuentaContado.Item(1).enabled = False
        cmdContado.Visible = False
        
        optStock(ActuStock_SI).Value = True
        
        optStock(ActuStock_RE).Visible = False
        TxtRemitoNumero.Visible = False
        Label1(ActuStock_RE).Visible = False
    
    Case Else                 'assert
        ufa "Prg", ".mostrar() " & Me.Name ', Err
    End Select
    Me.Show
End Sub

Private Sub chkPropio_LostFocus()
    set_uProd
End Sub

Private Sub cliente_cambio(codigo) ' As Integer)
    If ON_ERROR_HABILITADO Then On Error GoTo epa_ERR
    
    Dim ac As Variant, a As String, n As Long
    If codigo = 0 Then
        FrmBorrarTxt Me
    Else
        ac = obtenerDeSQL("select codigo, direccion, localidad, provincia, cuit, iva, FormaPago, Vendedor, Descuento1 from clientes where codigo = " & codigo)
        txtdireccion = sSinNull(ac(1))
        txtLocalidad = sSinNull(ac(2))
        a = sSinNull(ac(3))
        CmbProvincia = ObtenerDescripcionS("provincias", a)
        ucCuit.Text = sSinNull(ac(4))
 
        cmbformapago.ListIndex = BuscarEnCombo(cmbformapago, ac(6))
        CalcularVencimiento
        
        cmbvendedor.ListIndex = BuscarEnCombo(cmbvendedor, ac(7))
        
        n = ac(5)
        cmbTipoIva.ListIndex = BuscarEnCombo(cmbTipoIva, n)
        lblFacturaB.Visible = "B" = sSinNull(obtenerDeSQL("Select letra from ivas where codigo = " & n))
        
        mClienteConIva = Not mFAE And (0 < s2n(obtenerDeSQL("select porcentaje from PorcentajesIva where activo = 1 and iva = " & n)))
        txtPdescuento = ac(8) ' * 100)
        
        If ucBoton.estado = ucbEditando Then
            TxtNroFactura = ""
        End If
        set_uProd
    End If
    
    g.Borrar
    gO.Borrar
    mPropio = ""
fin:
    Exit Sub
epa_ERR:
    ufa "", "error leyendo codigo traido por c_cambio Fac Venta " & codigo ', Err
    Resume fin
End Sub

Private Sub cmbFormaPago_LostFocus()
    CalcularVencimiento
End Sub
Private Sub cmbFormaPago_Validate(cancel As Boolean)
    CalcularVencimiento
End Sub

Private Sub CalcularVencimiento()
    On Error GoTo ufa
    If Not dtVencimiento.enabled Then Exit Sub
    
    Dim sql As String
    Dim rsFormaP As New ADODB.Recordset
    
    sql = "Select dias from FormasPago WHERE codigo =" & cmbformapago.ItemData(cmbformapago.ListIndex)
    rsFormaP.Open sql, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
    dtVencimiento.Value = dtFecha + rsFormaP!Dias
ufa:
End Sub

Private Sub cmdOrigen_Click()
    CargaOrigen
End Sub

Private Sub cmdPedidosPendientes_Click()
    If ON_ERROR_HABILITADO Then On Error GoTo UFAcmdPend
    
    Dim s As String, i As Long, rs As New ADODB.Recordset, resu As String
    Dim prod, desc, cant, prec, pedi, remi, Item, prop, r As Long

    s = " Select distinct " _
      & " numero  as Numero, pedido_cli as NroPedidoCliente, cliente " _
      & " from Pedidos_Clientes inner join ItemPedidoCliente " _
      & " on Pedidos_clientes.numero = ItemPedidoCliente.Pedido " _
      & " where facturar > 0 and Pedidos_Clientes.activo = 1"

    If cliente.codigo > 0 Then s = s & " and cliente = " & cliente.codigo
   
    resu = frmBuscar.MostrarSql(s)
    If resu = "" Then Exit Sub
    
    cliente.codigo = s2n(frmBuscar.resultado(3))
    g.Borrar
    TabDetalle.Tab = 0

     s = " Select " _
       & " ItemPedidoCliente.codigo as cod, numero as nPedi, producto, cantidad, precio, facturar, CodigoPropio as codPropio, 0 as nRemi" _
       & " from Pedidos_Clientes inner join ItemPedidoCliente " _
       & " on Pedidos_clientes.numero = ItemPedidoCliente.Pedido " _
       & " where numero = " & resu
    
    rs.Open s, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    While Not rs.EOF
        If rs!cantidad = 0 Or rs!facturar > 0 Then
            prod = VerProductoCliente(rs!producto, rs!codPropio, cliente.codigo)
            desc = ObtenerDescripcionS("producto", rs!producto)

            cant = rs!facturar
            
            prec = rs!precio
            
            pedi = rs!nPedi
            Item = rs!cod
            chkPropio.Value = IIf(rs!codPropio, vbChecked, vbUnchecked)
            set_uProd
            MetoEnGrilla prod, desc, cant, prec, pedi, 0, Item, , False
            
            chkPropio.Value = IIf(rs!codPropio, vbChecked, vbUnchecked)
        End If
        rs.MoveNext
    Wend
fin:
    Set rs = Nothing
    relojito False
    Exit Sub
UFAcmdPend:

    ufa "Err leyendo pendientes", "cmdPeridoPend"
    Resume fin
End Sub

Private Sub cmdRemitosPendientes_Click()
    Dim s ', re

    If cliente.codigo = 0 Then
      
        s = " Select distinct " _
         & " RemitoVenta.numero as Remito, cliente as [ Cliente ], clientes.descripcion as [ Nombre                                  ]  " _
         & " from RemitoVenta inner join RemitoVentaDetalle " _
         & " on RemitoVenta.Numero = RemitoVentaDetalle.Numero " _
         & " inner join clientes " _
         & " on cliente = clientes.codigo" _
         & " where facturar > 0 " _
         & " and RemitoVenta.Anulado = 0 " _
         & " and RemitoVenta.Cancelado = 0 " _
         & " and RemitoVentaDetalle.Cancelado = 0 "
        If gEMPR_FormulaEsVirtual Then s = s & " and (formula = '' or formula = 'V')"
        
    Else
      s = " Select distinct " _
        & " RemitoVenta.numero as Remito, cliente as [ Cliente ], clientes.descripcion as [ Nombre                                    ]  " _
        & " from RemitoVenta inner join RemitoVentaDetalle " _
        & " on RemitoVenta.Numero = RemitoVentaDetalle.Numero " _
        & " inner join clientes " _
        & " on cliente = clientes.codigo" _
        & " where facturar > 0 " _
        & " and RemitoVenta.Anulado = 0 " _
        & " and RemitoVenta.Cancelado = 0 " _
        & " and RemitoVentaDetalle.Cancelado = 0 " _
        & " and cliente = " & cliente.codigo
        If gEMPR_FormulaEsVirtual Then s = s & " and (formula = '' or formula = 'V')"
    End If

    With frmBuscar
        If frmBuscar.MostrarSql(s) > "" Then
            If cliente.codigo = 0 Then cliente.codigo = s2n(.resultado(2))
            CargaDiscriminada s2n(.resultado(2)), s2n(.resultado(1))
        End If
    End With
    
    TabDetalle.Tab = 0
End Sub

Private Sub CargaDiscriminada(clie As Long, remi As Long)
    If ON_ERROR_HABILITADO Then On Error GoTo ufaDISCRI
    Dim rs As New ADODB.Recordset, s As String, i As Long
    
    s = " Select " _
        & " RemitoVentaDetalle.codigo as cod , RemitoVenta.numero as nRemi, producto, cantidad, precio, facturar, codPropio, 0 as nPedi" _
        & " from RemitoVenta inner join RemitoVentaDetalle " _
        & " on RemitoVenta.Numero = RemitoVentaDetalle.Numero " _
        & " where facturar > 0 " _
        & " and RemitoVenta.Anulado = 0 " _
        & " and RemitoVenta.Cancelado = 0 " _
        & " and RemitoVentaDetalle.Cancelado = 0 " _
        & " and cliente = " & clie
    
    rs.Open s, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly

    With rs
        If .EOF Then
            ufa "err al cargar remitos", "CargaDiscriminada no trajo items " & clie & " " & remi & Me.Name ', 0
        Else
            While Not .EOF
                ' si es = remi q eligio en el frmBuscar , lo meto en la grilla 1
                ' si no, lo meto en la segunda por las dudas
                                
                chkPropio.Value = IIf(!codPropio, vbChecked, vbUnchecked)
                                
                If !nRemi = remi Then
                    MetoEnGrilla VerProductoCliente(!producto, !codPropio, CLng(clie)), ObtenerDescripcionS("producto", !producto), !facturar, !precio, !nPedi, !nRemi, !cod, , False
                Else
                    i = gO.addRow
                    gO.tx i, gO_ITEM, !cod
                    gO.tx i, gO_NPED, !nPedi
                    gO.tx i, gO_NREM, !nRemi
                    gO.tx i, gO_PROD, VerProductoCliente(!producto, !codPropio, CLng(clie))
                    gO.tx i, gO_DESC, ObtenerDescripcionS("producto", !producto)
                    gO.tx i, gO_CANT, !cantidad
                    gO.tx i, gO_PEND, !facturar
                    gO.tx i, gO_PREC, !precio
                    gO.tx i, gO_PROP, !codPropio
                End If
                .MoveNext
            Wend
        End If
    End With
fin:
    Set rs = Nothing
    Exit Sub
ufaDISCRI:
    ufa "err cargando remito", " CargaDiscri  " & clie & " " & remi & Me.Name ', Err
    Resume fin
End Sub

Private Sub Form_Activate()
    SubimeSi800x600
End Sub

Private Sub Form_Load()
    Dim rsEjercicio As New ADODB.Recordset
    Dim rsFormaP As New ADODB.Recordset
    Dim sql As String
    
    comboSql CmbProvincia, "select descripcion from provincias where activo = 1"
    
    sql = "select descripcion,codigo,dias from formasPago where activo = 1 order by dias, codigo"
    rsFormaP.Open sql, DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
    Do While Not rsFormaP.EOF
      cmbformapago.AddItem rsFormaP!DESCRIPCION
      cmbformapago.ItemData(cmbformapago.NewIndex) = rsFormaP!codigo
      rsFormaP.MoveNext
    Loop
    rsFormaP.Close
    comboSql cboMoneda, "select descripcion, codigo from monedas order by codigo"
    comboSql cmbTipoIva, "select descripcion, codigo from ivas where activo = 1"
    comboSql cmbvendedor, "select descripcion, codigo  from usuarios where activo = 1 order by descripcion"
    comboArray cmbDeposito, Array("Deposito Central", "Deposito 1", "Deposito 2", "Deposito 3", "Deposito4"), Array(0, 1, 2, 3, 4)

    dtFecha = Date
    
    dtVencimiento = Date
    
    inigrilla 'grillas
    iniCliente 'clientes
    iniBotonera 'botonera

    rsEjercicio.Open "SELECT * From Ejercicio WHERE activo =1", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
   
    ucFechas.ini rsEjercicio!FechaInicio, rsEjercicio!FechaFin, ucefHorizontal, ucefFormatoSqlServer
    rsEjercicio.Close
    set_uProd
    
    optBuscarTipo.Item(0).Value = "1" ' True
    
    If gEMPR_EmiteFacturaConRemito Then
        optStock.Item(1).Value = "1" 'True
    Else
        optStock.Item(0).Value = "1" 'True
    End If
    
    TabDetalle.Tab = 0

    If mFAE Then
        cboMoneda.ListIndex = 1
    End If
    grilla.Editable = flexEDKbdMouse
    
End Sub

Private Function FaltaGrilla() As Boolean
    FaltaGrilla = (g.rows = 1 Or g.suma(gCANT) = 0)
    
    If FaltaGrilla Then
        TabDetalle.Tab = 0
        grilla.SetFocus
    End If
End Function

Private Sub BorrarCampos()
    On Error Resume Next
    TxtNroFactura = ""
    cliente.codigo = 0
    txtCodigo = ""
    ucCuit.Text = ""
    dtFecha = Date
    
    FrmBorrarTxt Me
    txtneto = ""
    txtIva = ""
    txttotal = ""
    txtDescuento = ""
    lblSubTotal = ""
    g.Borrar
    gO.Borrar
    txtCotizacion = ""

End Sub
Private Sub HabilitarEdicion(habilitar As Boolean)
    fraCabecera.enabled = False
    TabDetalle.enabled = habilitar
    fraEditDetalle.enabled = False
End Sub
Private Sub Resetear()
    mValoresOK = False
    optCuentaContado.Item(0).Value = True
    BorrarCampos
    HabilitarEdicion False
    midDoc = 0
    mPropio = ""

End Sub

Private Sub inigrilla()
    Set g = New LiGrilla
    Set gO = New LiGrilla
    
    g.init grilla, 4
    gO.init grillaOrigen
    
    gCANT = g.AddCol(" Cantidad    ", "N", 4)
    gprod = g.AddCol(" Producto            ")
    gDESC = g.AddCol(" Descripcion                                ")  ', "S")
    gPUNI = g.AddCol(" P.Unitario    ", "N", 4)
    gPTOT = g.AddCol(" P.Total       ", "9", 2)
    gNPED = g.AddCol(" Pedido      ") ', IIf(mFac = FacturaVenta_Pedido2, "-", "H"))
    gNPCL = g.AddCol(" Pedido Clie ") ', IIf(mFac = FacturaVenta_Pedido2, "-", "H"))
    gNREM = g.AddCol(" Remito        ", IIf(mFac = FacturaVenta_Remito2, "-", "H"))
    gFORM = g.AddCol(" Formula              ", "H")
    gITEM = g.AddCol(" itemPedidoRemito", "H")
    
    gO_NPED = gO.AddCol(" Pedido     ", IIf(mFac = FacturaVenta_Pedido2, "-", "H")) ' oculto si no es pedido
    g0_NPCL = gO.AddCol(" Pedido Clie ", IIf(mFac = FacturaVenta_Pedido2, "-", "H")) ' oculto si no es pedido
    gO_NREM = gO.AddCol(" Remito      ", IIf(mFac = FacturaVenta_Remito2, "-", "H")) ' oculto si no es   remito
    gO_ITEM = gO.AddCol("it pe/re ", "H")
    gO_CANT = gO.AddCol(" Cantidad    ")
    gO_PEND = gO.AddCol(" Pendiente ")
    gO_PROD = gO.AddCol(" Producto             ")
    gO_DESC = gO.AddCol(" Descripcion                              ")
    gO_PROP = gO.AddCol(" propio ", "H")
        
    If REMITO_CON_PRECIO Then
        gO_PREC = gO.AddCol(" Precio         ")
    Else
        gO_PREC = gO.AddCol(" Precio         ", "H")
    End If
    
End Sub

Private Sub iniCliente()
    Set cliente = New LiCodigo
    cliente.init cmbCliente, txtCodCliente, "Clientes", , , cmdCliente, "activo = 1", True
    cliente.EditaDescripcion = True
End Sub
Private Sub iniBotonera()
    ucBoton.init True, False, True, False, False
End Sub

Private Sub grilla_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    If g.Row > 0 Then g.delRow (g.Row)
End If
If KeyCode = 45 Then
    g.Row = g.addRow()
End If
End Sub

Private Sub optStock_Click(Index As Integer)
    TxtRemitoNumero.Visible = optStock(2).Value
End Sub

Private Sub CargaDatos()
'        cmbDeposito.ListIndex = s2n(!deposito)
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    Dim rs As New ADODB.Recordset, i As Long, z As Double
    
    TabDetalle.Tab = 0
    With rs
        .Open "select Codigo, TipoDoc, NroFactura, Cliente, Fecha, " _
            & " neto, iva, PorcentajeIva, Total, Descuento, Moneda, cotizacion, ActualizaStock, Deposito, iddoc, iibb " _
            & " from FacturaVenta where codigo = " & txtCodigo _
            , DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        
        z = !cotizacion
        If z = 0 Then z = 1
        txtCotizacion = z
        
        txtIva = s2n(!Iva) / z
        txtPdescuento.Text = (s2n(!Descuento, 4) * 100) / z
        txtneto = s2n(!Neto) / z
        txttotal = s2n(!Total) / z
        lblIIBB = s2n(!IIBB) / z
        txtPIVA = s2n(s2n(!PorcentajeIva, 4) * 100)
        midDoc = !iddoc
        lblSubTotal = !Neto + !Descuento
        cliente.codigo = !cliente
        
        cboMoneda.ListIndex = BuscarEnCombo(cboMoneda, !moneda)
        
        If !actualizaStock Then
            If gEMPR_EmiteFacturaConRemito Then
                optStock.Item(ActuStock_RE) = True
            Else
                optStock.Item(ActuStock_SI) = True
            End If
        Else
            optStock.Item(ActuStock_NO) = True
        End If
        
        .Close
        .Open "select cantidad, codpropio, producto, Descripcion, formula, precioUnitario, PrecioTotal, NroRemito from FacturaVentaDetalle where  codigoFactura = " & txtCodigo
        g.Borrar
        While Not .EOF
            i = g.addRow()
            chkPropio.Value = IIf(!codPropio, vbChecked, vbUnchecked)
            g.tx i, gCANT, s2n(!cantidad)
            g.tx i, gprod, VerProductoCliente(sSinNull(!producto), !codPropio, cliente.codigo)
            g.tx i, gDESC, sSinNull(!DESCRIPCION)
            g.tx i, gPUNI, s2n(!PrecioUnitario, 4) / z
            g.tx i, gPTOT, s2n(!PrecioTotal, 4) / z
            g.tx i, gFORM, sSinNull(!formula)
            g.tx i, gNREM, s2n(!NroRemito)
            
            'agregado trucho? ' cargo la descripcion del producto para las q grabamos sin descripcion
            If sSinNull(!DESCRIPCION) = "" And sSinNull(!producto) > "" Then g.tx i, gDESC, ObtenerDescripcionS("producto", !producto)
            
            .MoveNext
        Wend
    End With
    GoTo fin
ufaErr:
    ufa "err leyendo datos", Me.Name & txtCodigo ', Err
fin:
    Set rs = Nothing
End Sub


Private Sub CargaOrigen()
    Dim s As String, i As Long, rs As New ADODB.Recordset
    
    gO.Borrar
    Select Case mFac ' alias nPedi nRemi trae num remito o pedido
    Case FacturaVenta_Pedido2
    
    s = " Select " _
        & " ItemPedidoCliente.codigo as cod, numero as nPedi, pedido_cli as nPediCli, producto, cantidad, precio, facturar, CodigoPropio as codPropio, 0 as nRemi" _
        & " from Pedidos_Clientes inner join ItemPedidoCliente " _
        & " on Pedidos_clientes.numero = ItemPedidoCliente.Pedido " _
        & " where facturar > 0 " _
        & " and formula  = '' " _
        & " and cliente = " & cliente.codigo _
        & " and Pedidos_clientes.activo = 1 " _
        & " order by npedi, cod"
        
    Case FacturaVenta_Remito2
      s = " Select " _
        & " RemitoVentaDetalle.codigo as cod , RemitoVenta.numero as nRemi, producto, cantidad, precio, facturar, codPropio, 0 as nPedi, 0 as nPediCli" _
        & " from RemitoVenta inner join RemitoVentaDetalle " _
        & " on RemitoVenta.Numero = RemitoVentaDetalle.Numero " _
        & " where facturar > 0 " _
        & " and formula = '' " _
        & " and RemitoVenta.Anulado = 0 " _
        & " and RemitoVenta.Cancelado = 0 " _
        & " and RemitoVentaDetalle.Cancelado = 0 " _
        & " and cliente = " & cliente.codigo
        ' And activo = 1
    
    Case FacturaVenta_Remito2
      s = " Select " _
        & " RemitoVentaDetalle.codigo as cod , RemitoVenta.numero as nRemi,  producto, cantidad, precio, facturar, codPropio, 0 as nPedi, 0 as nPediCli " _
        & " from RemitoVenta inner join RemitoVentaDetalle " _
        & " on RemitoVenta.Numero = RemitoVentaDetalle.Numero " _
        & " where facturar > 0 " _
        & " and RemitoVenta.Anulado = 0 " _
        & " and RemitoVenta.Cancelado = 0 " _
        & " and RemitoVentaDetalle.Cancelado = 0 " _
        & " and cliente = " & cliente.codigo
     
    Case Else
        Exit Sub
    End Select
    
    rs.Open s, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    While Not rs.EOF
        i = gO.addRow
        gO.tx i, gO_ITEM, rs!cod
        gO.tx i, gO_NPED, rs!nPedi
        gO.tx i, g0_NPCL, rs!nPedicli
        gO.tx i, gO_NREM, rs!nRemi
        gO.tx i, gO_PROD, VerProductoCliente(rs!producto, rs!codPropio, cliente.codigo)
        gO.tx i, gO_DESC, ObtenerDescripcionS("producto", rs!producto)
        gO.tx i, gO_CANT, rs!cantidad
        gO.tx i, gO_PEND, rs!facturar
        gO.tx i, gO_PREC, rs!precio
        gO.tx i, gO_PROP, rs!codPropio
                
        rs.MoveNext
    Wend
    
    Set rs = Nothing
End Sub

Private Function Propio()
    Propio = (chkPropio.Value = vbChecked)
End Function

Private Sub MetoEnGrilla(prod, desc, cant, prec, pedi, remi, Item, Optional ivaprod, Optional conDesglose As Boolean)  ', cons)
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    
    Dim i As Long, rs As New ADODB.Recordset, ssql As String, codigomio As String, hay As Double, suma As Double ', ivaprod As Double
   
    codigomio = VerProductoMio(prod, Propio())
    
    If cant <> 0 Then
        If mClienteConIva Then
            If IsMissing(ivaprod) Then
                ivaprod = obtenerDeSQL("select iva from producto where codigo = '" & codigomio & "'") * 100
                
            End If
            
            suma = g.suma(gPUNI)
            If suma = 0 Then
                txtPIVA = ivaprod
            Else
                
                If s2n(ivaprod) <> s2n(txtPIVA) Then
                    If UCase(txtTipoDoc) = "FAB" Then
                    Else
                        che "Producto con distinto IVA al cargado"
                        Exit Sub
                    End If
                End If
            End If
        Else
            ivaprod = 0
            txtPIVA = "0"
        End If
    End If

    
    i = g.addRow()
    g.tx i, gprod, prod
    g.tx i, gDESC, desc
    g.tx i, gCANT, cant
    g.tx i, gPUNI, prec
    g.tx i, gPTOT, prec * cant
    g.tx i, gNPED, pedi
    g.tx i, gNPCL, sSinNull(obtenerDeSQL("select pedido_cli from Pedidos_Clientes where numero = " & pedi))
    g.tx i, gNREM, remi
    g.tx i, gITEM, Item
    
fin:
    RevisarTotales
    Exit Sub
ufaErr:
    ufa "err al poner en grilla", Me.Name ', Err
    Resume fin
End Sub

Private Sub MsgFalta(CodProd, canti)
    Dim hay As Double
    
    
    If CodProd = "" Then Exit Sub
    
    If mFac = FacturaVenta_Remito2 Or mFac = FacturaVenta_NCreditoDevolucion2 Then Exit Sub
    hay = HayProducto(VerProductoMio(CodProd, Propio()), cmbDeposito.ItemData(cmbDeposito.ListIndex))
    If hay < canti Then
        MsgBox " Stock para " & VerProductoCliente(CStr(CodProd), Propio(), cliente.codigo) & ", " & cmbDeposito.Text & "  : " & hay & vbCrLf & vbCrLf & " requeridos : " & canti
    End If
End Sub

Private Sub chkPropioEnabled(que As Boolean)
    chkPropio.enabled = grilla.rows < 2 And que
End Sub


Private Sub g_cambio(ByVal Row As Long, ByVal Col As Long, txt As String)
    Dim rs As New ADODB.Recordset
   
    With g
        .tx Row, gPTOT, s2n(.tx(Row, gCANT), 4) * s2n(.tx(Row, gPUNI), 4)
    End With
    
    If Trim(txtTipoDoc.Text) = "NCA" Or Trim(txtTipoDoc.Text) = "NDA" Then
        rs.Open "select * from facturaventadetalle where codigofactura=" & txtCodigo, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        If rs.EOF = True And rs.BOF = True Then
        Else
            If rs!PrecioTotal = 0 Then
            Else
                RevisarTotales
            End If
        End If
        Set rs = Nothing
    Else
        RevisarTotales
    End If
End Sub

Private Sub gO_DblClick()
    'Static sPropio As String
    Dim prod, desc, cant, prec, pedi, remi, Item, prop, r As Long
    Dim cantParaPasar As Double
    
    
    With gO
        r = .Row
        If r > 0 Then
            prod = .tx(r, gO_PROD)
            desc = .tx(r, gO_DESC)
            cant = .tx(r, gO_PEND)
            prec = IIf(BS_RESPETAR_PRECIO_OC, .tx(r, gO_PREC), 0)
            pedi = .tx(r, gO_NPED)
            remi = .tx(r, gO_NREM)
            Item = .tx(r, gO_ITEM)
            prop = .tx(r, gO_PROP)

            If mPropio = "" Then
                mPropio = CStr(prop)
                chkPropio.Value = IIf((prop = "1"), vbChecked, vbUnchecked)
            Else
                If mPropio <> CStr(prop) Then
                    If prop = "True" Then
                        prod = VerProductoCliente(CStr(prod), False, cliente.codigo)
                    Else
                        prod = VerProductoMio(prod, True)
                    End If
                End If
            End If
                        
            If s2n(prec) = 0 Then prec = precioProducto(CStr(prod), Propio(), cliente.codigo)
            If cant > 1 And BS_PED_A_FAC_PREGUNTA_CANT Then
                cant = s2n(InputBox("Cantidad : ", prod & " A facturar", cant))
            End If
            
            MetoEnGrilla prod, desc, cant, prec, pedi, remi, Item, , False
            
            .delRow r
        End If
    End With
End Sub

Private Sub txtcantidad_GotFocus()
    PintoFocoActivo
End Sub

Private Sub txtIvaProducto_GotFocus()
    PintoFocoActivo
End Sub

Private Sub txtPIVA_LostFocus()
    txtPIVA = n2r(s2n(txtPIVA))
    RevisarTotales
End Sub

Private Sub txtprecio_GotFocus()
    PintoFocoActivo
End Sub

Private Sub txtPdescuento_LostFocus()
    txtPdescuento = n2r(s2n(txtPdescuento))
    RevisarTotales
End Sub

Private Function FaltaAlgo() As Boolean
    FaltaAlgo = True
        
    If ComboCodigo(cboMoneda) > 1 And s2n(txtCotizacion) = 0 Then
        che "Falta cotizacion"
        Exit Function
    End If
    If FaltaGrilla() Then
        che "Faltan datos en la grilla"
        Exit Function
    End If
    If Trim$(txtPIVA) = "" Then
        txtPIVA = InputBox("Falta Iva de la factura" & vbCrLf & "si no tiene ingrese 0")
        If Trim$(txtPIVA) > "" Then txtPIVA = s2n(txtPIVA)
        Exit Function
    End If
    
    FaltaAlgo = False
End Function


Private Function EmitirRemito() As Long
    Dim tmp, formula As String, num As Long, sucursal As Long, depot As Long, i As Long, mjstk As Long, produ As String
    
    sucursal = s2n(obtenerDeSQL("select sucursal from datos"))
    depot = cmbDeposito.ItemData(cmbDeposito.ListIndex)
    num = s2n(TxtRemitoNumero)
    
    tmp = obtenerDeSQL("select cliente from RemitoVenta where Numero = " & s2n(TxtRemitoNumero))
    If Not IsEmpty(tmp) Then
        che "Numero Remito ya grabado, para cliente " & tmp
        Exit Function
    End If
    
    '*******************************************************************
    ' quiero Transaccion !  , y/o quiero hacer tabla temp y 1 solo stored

    frmRemitoVenta.ABMRemitoVenta "A", num, cliente.codigo, (dtFecha), 0, 0, depot, Propio(), "", "", "", "", "", 0
    '
    For i = 1 To g.rows - 1 'items
        formula = g.tx(i, gFORM)
        produ = VerProductoMio(g.tx(i, gprod), Propio())
        mjstk = ManejaStock(produ)
        frmRemitoVenta.ABMRVDetalle "A", num, produ _
            , s2n(g.tx(i, gCANT)), s2n(g.tx(i, gPUNI)), s2n(g.tx(i, gNPED)) _
            , depot, 0, formula, mjstk, 0 'cantConsign, formula
    Next i
    
    ' quiero Transaccion, y/o quiero hacer tabla temp y 1 solo stored
    '*******************************************************************

'    MsgBox "Remito " & num & " grabado"
    EmitirRemito = s2n(TxtRemitoNumero)
End Function


' **********************************************
Private Sub ucBoton_AceptarAlta()
    
   If ON_ERROR_HABILITADO Then On Error GoTo UFAalta
    
    Dim NroRemito As Long
    Dim tmp, tmpfec As Date
    Dim QuieroLeyenda As Boolean
    
    If FaltaAlgo() Then Exit Sub
    
    If confirma("Factura: " & vbCrLf & vbCrLf & "Tipo :  " & txtTipoDoc & vbCrLf & "Nro :   " & TxtNroFactura & vbCrLf & "Confirma actualizacion ?") Then
    
'       transaccion va dentro de grabaFactura
''        DE_BeginTrans
        If GrabaFactura() Then          ' graba tamb valores contado
            If VaConRemito Then NroRemito = EmitirRemito()
''            DE_CommitTrans
            TabDetalle.Tab = 0
            
            ucBoton.AceptarOk
        End If
    End If
    
fin:
    Exit Sub
UFAalta:
''    DE_RollbackTrans
    ufa "Fallo el alta", ""
    Resume fin
UFAinprime:
    ufa "falla en impresion", ""
    Resume fin
End Sub

Private Sub ucBoton_AceptarModi()
    
    If ON_ERROR_HABILITADO Then On Error GoTo UFAalta
    
    Dim NroRemito As Long
    Dim tmp, tmpfec As Date
    Dim QuieroLeyenda As Boolean
    
    If FaltaAlgo() Then Exit Sub
    
    If confirma("Factura: " & vbCrLf & vbCrLf & "Tipo :  " & txtTipoDoc & vbCrLf & "Nro :   " & TxtNroFactura & vbCrLf & "Confirma actualizacion ?") Then
    
'       transaccion va dentro de grabaFactura
''        DE_BeginTrans
        If GrabaFactura() Then          ' graba tamb valores contado
            If VaConRemito Then NroRemito = EmitirRemito()
''            DE_CommitTrans
            TabDetalle.Tab = 0
            
            ucBoton.AceptarOk
        End If
    End If
    
fin:
    Exit Sub
UFAalta:
''    DE_RollbackTrans
    ufa "Fallo el alta", ""
    Resume fin
UFAinprime:
    ufa "falla en impresion", ""
    Resume fin
End Sub

Private Sub ucBoton_BorrarControles()
    Resetear
End Sub

Private Sub ucBoton_Buscar()
    If ON_ERROR_HABILITADO Then On Error GoTo fin
    Dim re As Variant, WhereTipo As String, WhereFecha As String
    
'    WhereTipo = IIf(optBuscarTipo.Item(0).Value, " (TipoDoc = 'FAA' or TipoDoc = 'NCA' or TipoDoc = 'NDA') ", "(TipoDoc = 'FAB')")
    If optBuscarTipo.Item(0).Value Then
        WhereTipo = " (TipoDoc = 'FAA' or TipoDoc = 'NCA' or TipoDoc = 'NDA') "
    ElseIf optBuscarTipo.Item(1).Value Then
        WhereTipo = "(TipoDoc = 'FAB' or TipoDoc = 'NCB' or TipoDoc = 'NDB') " ' ahora lo levanta el frm de ND x chq rechaz
    Else
        WhereTipo = " (TipoDoc = 'FAE' or TipoDoc = 'NCE' or TipoDoc = 'NDE') "
    End If
    
    WhereFecha = "fecha " & ucFechas.ssBetween()
    
    With frmBuscar

        re = .MostrarSql("select f.Codigo as Codigo, TipoDoc, NroFactura, Cliente, c.descripcion as [ Nombre                        ], Fecha as [Fecha ], f.activo as Anulada, Remito  from " & mTablaFV & " as f left join clientes as c on c.codigo = f.cliente where " & WhereTipo & " and " & WhereFecha & " order by NroFactura desc ", , , , "", "Anulada", False)
        
        If re = "" Then Exit Sub
        txtCodigo = .resultado(1)
        txtTipoDoc = .resultado(2)
        TxtNroFactura = .resultado(3)
        cliente.codigo = .resultado(4)
        dtFecha = .resultado(6)
        TxtRemitoNumero = .resultado(8)
        
        mFAE = (Trim(.resultado(2)) = "FAE")
        lblExterior.Visible = mFAE
        
    End With
    CargaDatos
    
    gO.Borrar
    ucBoton.BuscarOK
fin:
End Sub

Private Sub ucBoton_HabilitarEdicion(sino As Boolean)
    HabilitarEdicion sino
End Sub

Private Sub ucBoton_Modificar()
    mNuevo = False
End Sub

Private Sub ucBoton_Salir()
    Unload Me
End Sub
' **********************************************

Private Sub Form_KeyPress(KeyAscii As Integer) ' con Frm.KeyPreView = true
    FrmKeyPress KeyAscii, True, True, True
End Sub

Private Sub Form_Terminate()
    Set g = Nothing
    Set gO = Nothing
    Set cliente = Nothing
End Sub

Private Function GrabaFactura() As Boolean
    If ON_ERROR_HABILITADO Then On Error GoTo UfaGraba
    
    Dim i As Long, bol As Boolean, j As Long
    Dim asse As String ' assert
    Dim cant As Double, prod As String, formu As String, puni, plis As Double, pedi As Long, ptot, remi As Long, Item As Long, depot, desc As String, RemitoJunto As Long, alia As String
    Dim codTipoDoc, Serie, intBajaStock As Long
    Dim Total As Double, saldo As Double, bContado As Long, bNCxDevol As Long
    Dim iddoc As Long, asieVenta As New Asiento, cuco As String
    Dim TextoAsientoComprobante As String
    Dim z As Double ' COTIZACION
    Dim intereses As Double
    Dim NroFac As Integer
    Dim Aux As New ADODB.Recordset
    
    z = s2n(txtCotizacion, 4)
    If z = 0 Then z = 1

    GrabaFactura = False
    intBajaStock = 0
    
    depot = cmbDeposito.ItemData(cmbDeposito.ListIndex)
    
    If VaConRemito() Then RemitoJunto = s2n(TxtRemitoNumero)
    
'*** una transaccion aqui ..... *********************
   
DE_BeginTrans
    
    Total = s2n(txttotal, 2)
    saldo = IIf(EsContado(), 0, Total)
    bContado = optCuentaContado.Item(1).Value
    bNCxDevol = (mFac = FacturaVenta_NCreditoDevolucion2)
    
    iddoc = obtenerDeSQL("select iddoc from facturaventa where nrofactura=" & TxtNroFactura & " and tipodoc='" & txtTipoDoc & "'") 'NuevoDocumento(txtTipoDoc, TxtNroFactura, 0, 0)
    
    NroFac = TxtNroFactura.Text
    
    asse = "GrabaDetalle"
    For i = 1 To g.rows - 1
        If g.tx(i, gDESC) = "" Then i = i + 1
        asse = "GrabaDet: calc grilla: prod,form,puni,ptot"
        cant = s2n(g.tx(i, gCANT))

        prod = VerProductoMio(g.tx(i, gprod), Propio())
        If prod = "" Then prod = "-"
        formu = "" ' VerProductoMio(g.tx(i, gFORM), Propio())
        puni = s2n(g.tx(i, gPUNI))

        ptot = s2n(g.tx(i, gPTOT))
        
        asse = "GrabaDet: calc grilla: p lis"
        If prod = "-" Then
            plis = 0
        Else
            plis = s2n(obtenerDeSQL("select precio from producto where codigo = '" & VerProductoMio(prod, Propio()) & "'"))
        End If
        asse = "GrabaDet: calc grilla: pedi,remi,item,desc"
        pedi = s2n(g.tx(i, gNPED))
        remi = s2n(g.tx(i, gNREM))
        Item = s2n(g.tx(i, gITEM))
        desc = Trim(g.tx(i, gDESC))
        
        asse = "GrabaDet: calc grilla: cuentaproducto"
        
        'MODIFICACION INTELIGENTE: POR PRODUCTO!!! aguante cuco
'        cuco = CuentaProducto(prod)
        
        asse = "Actualiza existencia: Graba SP"
        'asse = "SELECT * From facturaventadetalle d inner join facturaventa f on f.tipodoc=d.tipodoc and f.nrofactura=d.nrofactura WHERE d.tipodoc='" & Trim(txtTipoDoc) & "' and d.nrofactura=" & TxtNroFactura & " and d.producto='" & Trim(prod) & "'"
        Aux.Open "SELECT * From facturaventadetalle d inner join facturaventa f on f.tipodoc=d.tipodoc and f.nrofactura=d.nrofactura WHERE d.tipodoc='" & Trim(txtTipoDoc) & "' and d.nrofactura=" & TxtNroFactura & " and d.producto='" & Trim(prod) & "'", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
        If ABMFVDetalle("B", s2n(txtCodigo), txtTipoDoc, TxtNroFactura, Aux!cantidad, Propio(), prod, desc, formu, Aux!PrecioUnitario, ptot * z, plis, pedi, remi, Aux!item_p_r, iddoc, IIf(Aux!actualizaStock = True, 1, 0)) = False Then GoTo UfaGraba
                
        asse = "GrabaDetalle: Graba SP"
        If ABMFVDetalle("M", s2n(txtCodigo), txtTipoDoc, TxtNroFactura, cant, Propio(), prod, desc, formu, puni * z, ptot * z, plis, pedi, remi, Aux!item_p_r, iddoc, IIf(Aux!actualizaStock = True, 1, 0)) = False Then GoTo UfaGraba
        
        Set Aux = Nothing
    Next i
    
DE_CommitTrans
'*** una transaccion hasta aqui ..... *********************
    
    GrabaFactura = True
    MsgBox "Se ha actualizado con exito.", , "ATENCION"
    
fin:
    Exit Function
    
UfaGraba:
    DE_RollbackTrans
    ufa "Err al grabar ", " grabaFV() - " & asse & " " & prod ', Err
    Resume fin
End Function

Private Sub RevisarTotales()
    Dim subtotal As Double, descu As Double, Neto As Double
    Dim dTipo As String
    dTipo = UCase(txtTipoDoc)
    If InStr(dTipo, "B") Then txtPIVA = 0
    
    If gEMPR_idEmpresa = 2 Then
        subtotal = s2n(g.suma(gPTOT), 4)
    Else
        subtotal = s2n(g.suma(gPTOT), 2)
    End If
    subtotal = s2n(g.suma(gPTOT), 2)
    descu = s2n(subtotal, 4) * (s2n(txtPdescuento, 4) / 100)
    Neto = s2n(subtotal - descu, 4)
    
    lblSubTotal = s2n(subtotal, 4)
    txtDescuento = s2n(descu, 2)
    txtneto = Neto
    
    lblIIBB = CalcPercIIBB(Neto, cliente.codigo)
    
    
    txtIva = s2n(Neto * (s2n(txtPIVA, 4) / 100), 2)
    
    txttotal = s2n(Neto + s2n(txtIva, 4), 2) + s2n(lblIIBB)
    
End Sub


Private Sub verCampos(verOrigen As Boolean, habChk As Boolean, habFraEditDet As Boolean, gTop As Long, gHe As Long, verDep As Boolean, capt As String, verchkBajaStk As Boolean, chkBajaStkValue, verFacturaReferencia As Boolean, verBotonRemito As Boolean, verBotonPedido As Boolean)
    TabDetalle.TabVisible(1) = verOrigen
    
    'chkPropio.Enabled = habChk
    chkPropio.enabled = True  ' creo que falla con migrados
    
    fraEditDetalle.Visible = habFraEditDet
    grilla.Top = gTop
    grilla.Height = gHe
    lblDepot.Visible = verDep
    cmbDeposito.Visible = verDep
    Me.caption = capt
    
    fraOptStock.Visible = verchkBajaStk
    
    optStock.Item(chkBajaStkValue).Value = True
    
    lblref.Visible = verFacturaReferencia
    txtTipoDocRef.Visible = verFacturaReferencia
    txtNroFacturaRef.Visible = verFacturaReferencia
    cmdRemitosPendientes.Visible = verBotonRemito
    cmdPedidosPendientes.Visible = verBotonPedido
End Sub

Private Sub BuscarFacturaReferencia()
    If ON_ERROR_HABILITADO Then On Error GoTo UfaErrFRef
    Dim s As String, re As String, aRe As Variant, i As Long
    Dim rs As New ADODB.Recordset
    Dim asse As String
    Dim tempo

    s = "select FacturaVenta.codigo, TipoDoc, NroFactura, Fecha, descripcion, contado from FacturaVenta inner join Clientes on FacturaVenta.Cliente = Clientes.Codigo where FacturaVenta.activo = 1 and  (tipodoc = '" & TipoDoc_FACTURA_A & "' or tipoDoc = '" & TipoDoc_FACTURA_B & "') and fecha  " & ucFechas.ssBetween() & " order by FacturaVenta.codigo desc "
    
    With frmBuscar
        asse = "por buscar"
        re = .MostrarSql(s, , "Credito Sobre Factura:", , StringCONTADO, "Cta Cte")
        If re <> "" Then
            asse = "tipodoc"
            txtTipoDocRef = .resultado(2)
            asse = "nro ref"
            txtNroFacturaRef = .resultado(3)
            asse = "clie Desc"
            cliente.DESCRIPCION = .resultado(5) ' datos cliente cambian solos
            asse = "optCont"
            optCuentaContado.Item(1).Value = (.resultado(6) = StringCONTADO)
            tempo = obtenerDeSQL("select iva, porcentajeiva, descuento from facturaventa where codigo = " & .resultado(1))
            txtIva = s2n(tempo(0))
            txtPIVA = s2n(tempo(1), 4) * 100
            txtPdescuento = s2n(tempo(2), 4) * 100
            
            With rs
                asse = "leo items"
                .Open "select Producto, Cantidad, codPropio, precioUnitario, descripcion from FacturaVentaDetalle where CodigoFactura = " & re, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                g.Borrar
                While Not .EOF
                    asse = "it propio"
                    chkPropio.Value = IIf(!codPropio, vbChecked, vbUnchecked)

                    asse = " it grilla " & !producto
                    MetoEnGrilla VerProductoCliente(!producto, Propio(), cliente.codigo), !DESCRIPCION, !cantidad, !PrecioUnitario, 0, 0, 0, s2n(txtPIVA), False
                    .MoveNext
                Wend
            End With
        End If
    End With
fin:
    Set rs = Nothing
    Exit Sub
UfaErrFRef:
    ufa "Err al buscar factura referencia", re & " - BuscRef:" & asse
    Resume fin
End Sub

Private Function EsContado()
    EsContado = (optCuentaContado.Item(1).Value) And (Not mFac = FacturaVenta_NCreditoDevolucion2)
End Function

Private Sub set_uProd() ' lo copie de pedido cliente
    Dim sqlbuscar As String, sqldesc As String

    If Propio() Then    'propio
        sqldesc = "select descripcion from producto where codigo = '###' "
        sqlbuscar = "select codigo as [ Codigo                 ],  descripcion as [ Descripcion                                                 ] from producto where activo = 1 order by codigo "
    Else    'relCliente
        sqldesc = "select descripcion from producto  " _
            & " inner join relacion_Producto_Cliente " _
            & " on producto.codigo = relacion_Producto_cliente.producto " _
            & " where cliente = " & cliente.codigo & " and productoCliente = '###'"
        sqlbuscar = "select relacion_producto_cliente.productoCliente, producto.descripcion, producto.codigo, relacion_producto_cliente.precio " _
            & " from producto  " _
            & " inner join relacion_Producto_Cliente " _
            & " on producto.codigo = relacion_Producto_cliente.producto " _
            & " where cliente = " & cliente.codigo _
            & " and producto.activo = 1 and relacion_producto_cliente.activo = 1 " _
            & " order by producto"
    End If
    uProd.ini sqldesc, sqlbuscar, True
    uProd.EditaDescripcion = True
End Sub

Private Sub uProd_cambio(codigo As Variant)
    
    Dim tmpIvaProd As Variant ', tmpIvaClie
     
     tmpIvaProd = obtenerDeSQL("select iva from producto where codigo = '" & VerProductoMio(uProd.codigo, Propio()) & "'") '* 100
    If mClienteConIva Then
        txtIvaProducto.enabled = (IsEmpty(tmpIvaProd))
        txtIvaProducto = (tmpIvaProd * 100)
        
        If (IsEmpty(tmpIvaProd)) Then
            txtIvaProducto.enabled = True
            txtIvaProducto = txtPIVA
        Else
            txtIvaProducto.enabled = False
            txtIvaProducto = (tmpIvaProd * 100)
        End If
    Else
        txtIvaProducto.enabled = False
        txtIvaProducto = "0"
    End If
    
End Sub

Private Function VaConRemito() As Boolean
    VaConRemito = optStock.Item(ActuStock_RE).Value
End Function

Public Function ABMFVDetalle(vOpe As String, vCodFactura As Long, vTipoDoc As String, vNroFactura As Long, vCant As Double, vCodPropio As Long, vProducto As String, vDescripcion As String, vFormula As String, vPrecUni As Double, vPrecTot As Double, vPrecLista As Double, vNroPedido As Long, vNroRemito As Long, vItem As Long, vIdDoc As Long, Optional vBajaStock As Long = 0, Optional vDeposito As Long = 0) As Boolean
On Error GoTo fvdmal
Dim idd As String
Dim f, fFactor As Double, fCargar As Double

Set f = Nothing
f = obtenerDeSQL("select f.factor from producto p inner join umfactor f on p.ufactor=f.ufcodigo where p.codigo=" & ssTexto(vProducto))
If IsNull(f) Or IsEmpty(f) Then
    fFactor = 1
Else
    fFactor = f
End If
fCargar = fFactor * vCant


ABMFVDetalle = True
Select Case vOpe
    Case "M":
        idd = "  update FacturaVentaDetalle set Cantidad=" & x2s(vCant) & " ,PrecioUnitario=" & x2s(vPrecUni) & " " _
        & " where tipodoc=" & ssTexto(vTipoDoc) & " and nrofactura=" & vNroFactura & " and producto=" & ssTexto(vProducto)
        DataEnvironment1.Sistema.Execute idd
        
        If vNroRemito > 0 Then
            idd = " Update RemitoVentaDetalle " _
                & " Set cantidad = " & x2s(vCant) & ",precio=" & x2s(vPrecUni) & " , facturar=" & x2s(vCant) _
                & " Where codigo = " & ssTexto(vItem)
            DataEnvironment1.Sistema.Execute idd
            
            idd = " Update RemitoVentaDetalle " _
                & " Set facturar = facturar - " & x2s(vCant) _
                & " Where codigo = " & ssTexto(vItem)
            DataEnvironment1.Sistema.Execute idd
            
            If vDeposito = 0 Then
                idd = " Update producto " _
                    & " Set existencia = existencia - " & x2s(fCargar) _
                    & " Where codigo = " & ssTexto(vProducto)
                DataEnvironment1.Sistema.Execute idd
            End If
            If vDeposito = 1 Then
                idd = " Update producto " _
                    & " Set dep1 = dep1 - " & x2s(fCargar) _
                    & " Where codigo = " & ssTexto(vProducto)
                DataEnvironment1.Sistema.Execute idd
            End If
            If vDeposito = 2 Then
                idd = " Update producto " _
                    & " Set dep2 = dep2 - " & x2s(fCargar) _
                    & " Where codigo = " & ssTexto(vProducto)
                DataEnvironment1.Sistema.Execute idd
            End If
            If vDeposito = 3 Then
                idd = " Update producto " _
                    & " Set dep3 = dep3 - " & x2s(fCargar) _
                    & " Where codigo = " & ssTexto(vProducto)
                DataEnvironment1.Sistema.Execute idd
            End If
            If vDeposito = 4 Then
                idd = " Update producto " _
                    & " Set dep4 = dep4 - " & x2s(fCargar) _
                    & " Where codigo = " & ssTexto(vProducto)
                DataEnvironment1.Sistema.Execute idd
            End If
        End If
        
        If vNroRemito = 0 And vBajaStock = 1 Then
            If vDeposito = 0 Then
                idd = " Update producto " _
                    & " Set existencia = existencia - " & x2s(fCargar) _
                    & " Where codigo = " & ssTexto(vProducto)
                DataEnvironment1.Sistema.Execute idd
            End If
            If vDeposito = 1 Then
                idd = " Update producto " _
                    & " Set dep1 = dep1 - " & x2s(fCargar) _
                    & " Where codigo = " & ssTexto(vProducto)
                DataEnvironment1.Sistema.Execute idd
            End If
            If vDeposito = 2 Then
                idd = " Update producto " _
                    & " Set dep2 = dep2 - " & x2s(fCargar) _
                    & " Where codigo = " & ssTexto(vProducto)
                DataEnvironment1.Sistema.Execute idd
            End If
            If vDeposito = 3 Then
                idd = " Update producto " _
                    & " Set dep3 = dep3 - " & x2s(fCargar) _
                    & " Where codigo = " & ssTexto(vProducto)
                DataEnvironment1.Sistema.Execute idd
            End If
            If vDeposito = 4 Then
                idd = " Update producto " _
                    & " Set dep4 = dep4 - " & x2s(fCargar) _
                    & " Where codigo = " & ssTexto(vProducto)
                DataEnvironment1.Sistema.Execute idd
            End If
        End If
        

    Case "B":
        If vNroRemito > 0 Then
             idd = " Update RemitoVentaDetalle " _
                & " Set facturar = facturar + " & x2s(vCant) _
                & " Where codigo = " & ssTexto(vItem)
            DataEnvironment1.Sistema.Execute idd
            
            If vDeposito = 0 Then
                idd = " Update producto " _
                    & " Set existencia = existencia + " & x2s(vCant) _
                    & " Where codigo = " & ssTexto(vProducto)
                DataEnvironment1.Sistema.Execute idd
            End If
            If vDeposito = 1 Then
                idd = " Update producto " _
                    & " Set dep1 = dep1 + " & x2s(vCant) _
                    & " Where codigo = " & ssTexto(vProducto)
                DataEnvironment1.Sistema.Execute idd
            End If
            If vDeposito = 2 Then
                idd = " Update producto " _
                    & " Set dep2 = dep2 + " & x2s(vCant) _
                    & " Where codigo = " & ssTexto(vProducto)
                DataEnvironment1.Sistema.Execute idd
            End If
            If vDeposito = 3 Then
                idd = " Update producto " _
                    & " Set dep3 = dep3 + " & x2s(vCant) _
                    & " Where codigo = " & ssTexto(vProducto)
                DataEnvironment1.Sistema.Execute idd
            End If
            If vDeposito = 4 Then
                idd = " Update producto " _
                    & " Set dep4 = dep4 + " & x2s(vCant) _
                    & " Where codigo = " & ssTexto(vProducto)
                DataEnvironment1.Sistema.Execute idd
            End If
        End If
        
        If vNroRemito = 0 And vBajaStock = 1 Then
            If vDeposito = 0 Then
                idd = " Update producto " _
                    & " Set existencia = existencia + " & x2s(vCant) _
                    & " Where codigo = " & ssTexto(vProducto)
                DataEnvironment1.Sistema.Execute idd
            End If
            If vDeposito = 1 Then
                idd = " Update producto " _
                    & " Set dep1 = dep1 + " & x2s(vCant) _
                    & " Where codigo = " & ssTexto(vProducto)
                DataEnvironment1.Sistema.Execute idd
            End If
            If vDeposito = 2 Then
                idd = " Update producto " _
                    & " Set dep2 = dep2 + " & x2s(vCant) _
                    & " Where codigo = " & ssTexto(vProducto)
                DataEnvironment1.Sistema.Execute idd
            End If
            If vDeposito = 3 Then
                idd = " Update producto " _
                    & " Set dep3 = dep3 + " & x2s(vCant) _
                    & " Where codigo = " & ssTexto(vProducto)
                DataEnvironment1.Sistema.Execute idd
            End If
            If vDeposito = 4 Then
                idd = " Update producto " _
                    & " Set dep4 = dep4 + " & x2s(vCant) _
                    & " Where codigo = " & ssTexto(vProducto)
                DataEnvironment1.Sistema.Execute idd
            End If
        End If
End Select
Exit Function
fvdmal:
ABMFVDetalle = False
End Function


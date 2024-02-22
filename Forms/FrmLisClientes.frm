VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form FrmLisClientes 
   Caption         =   "Listado de Clientes"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7035
   Icon            =   "FrmLisClientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin Gestion.ucXls ucXls1 
      Height          =   975
      Left            =   240
      TabIndex        =   31
      Top             =   4320
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1720
   End
   Begin VSFlex7LCtl.VSFlexGrid gClientes 
      Height          =   3615
      Left            =   120
      TabIndex        =   30
      Top             =   5040
      Visible         =   0   'False
      Width           =   6615
      _cx             =   11668
      _cy             =   6376
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
   Begin VB.OptionButton optcateg 
      Caption         =   "Categoría"
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
      Height          =   195
      Left            =   2640
      TabIndex        =   29
      Top             =   1320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.OptionButton optelegir 
      Caption         =   "Elegir"
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
      Height          =   195
      Left            =   2640
      TabIndex        =   28
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.OptionButton opttodos 
      Caption         =   "Todos"
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
      Height          =   195
      Left            =   2640
      TabIndex        =   27
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtdescateg 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3840
      TabIndex        =   25
      Tag             =   "2"
      Top             =   2040
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmbcateg 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Categoría"
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
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtcateg 
      Height          =   285
      Left            =   1200
      TabIndex        =   23
      Tag             =   "1"
      Top             =   2040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtcodvend 
      Height          =   285
      Left            =   1200
      TabIndex        =   21
      Tag             =   "1"
      Top             =   2040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmbvend 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Vendedor"
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
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtdesvend 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3840
      TabIndex        =   19
      Tag             =   "2"
      Top             =   2040
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1380
      Left            =   240
      TabIndex        =   16
      Top             =   240
      Width           =   2055
      Begin VB.OptionButton optvend 
         Caption         =   "Vendedores"
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
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton optcliente 
         Caption         =   "Clientes"
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
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1455
      End
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
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   12
      Tag             =   "0"
      Top             =   4320
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
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4320
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
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox txtdeshasta 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3840
      TabIndex        =   10
      Tag             =   "2"
      Top             =   2760
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmbhasta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cliente"
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
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txthasta 
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Tag             =   "0"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtdesde 
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Tag             =   "1"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmbdesde 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cliente"
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
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtdesdesde 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3840
      TabIndex        =   5
      Tag             =   "2"
      Top             =   1920
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox cargar 
      Height          =   285
      Left            =   5280
      TabIndex        =   4
      Tag             =   "1"
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame frameorden 
      Caption         =   "Orden"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   780
      Left            =   240
      TabIndex        =   1
      Top             =   3240
      Visible         =   0   'False
      Width           =   6495
      Begin VB.OptionButton optDesc 
         Caption         =   "Descripción"
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
         Left            =   3840
         TabIndex        =   3
         Top             =   260
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton optcodigo 
         Caption         =   "Código"
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
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.Frame FrameCli 
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1380
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      Height          =   1695
      Left            =   120
      Top             =   120
      Width           =   6855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      X1              =   120
      X2              =   6960
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label lblcateg 
      Caption         =   "Categoría:"
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
      Left            =   240
      TabIndex        =   26
      Top             =   2040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblvend 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Código"
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
      TabIndex        =   22
      Top             =   2040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblhasta 
      Caption         =   "Hasta:"
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
      TabIndex        =   15
      Top             =   2760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lbldesde 
      Caption         =   "Desde:"
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
      TabIndex        =   13
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "FrmLisClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit ' 16/9/4 ' 11/3/5


Private Sub cmbcateg_Click()
'    FrmHelp.Show
'    CargarHelp "Categorias", "Codigo", "Descripción", "codigo", "descripcion", "codigo"
'    FrmHelp.Tag = Me.Name
'    cargar = "Categ"
    Dim resu As String

    resu = frmBuscar.MostrarSql("select codigo,descripcion from categorias where activo=1")
    If resu > "" Then
        txtcateg = frmBuscar.resultado
        txtdescateg = frmBuscar.resultado(2)
    End If
End Sub

Private Sub cmbdesde_Click()
    Dim resu As String

    resu = frmBuscar.MostrarSql("select codigo,descripcion from clientes where activo=1")
    If resu > "" Then
        txtdesde = frmBuscar.resultado
        txtdesdesde = frmBuscar.resultado(2)
    End If
'    FrmHelp.Show
'    CargarHelp "Clientes", "Código", "Descripción", "codigo", "descripcion", "codigo"
'    FrmHelp.Tag = Me.Name
'    cargar = "CliDesde"
End Sub

Private Sub cmbhasta_Click()
'    FrmHelp.Show
'    CargarHelp "Clientes", "Código", "Descripción", "codigo", "descripcion", "codigo"
'    FrmHelp.Tag = Me.Name
'    cargar = "CliHasta"

    Dim resu As String

    resu = frmBuscar.MostrarSql("select codigo,descripcion from clientes where activo=1")
    If resu > "" Then
        txthasta = frmBuscar.resultado
        txtdeshasta = frmBuscar.resultado(2)
    End If
End Sub

Private Sub cmbvend_Click()
'    FrmHelp.Show
'    CargarHelp "usuarios", "Codigo", "Descripción", "codigo", "descripcion", "codigo"
'    FrmHelp.Tag = Me.Name

    Dim resu As String

    resu = frmBuscar.MostrarSql("select codigo,descripcion from usuarios where activo=1")
    If resu > "" Then
        txtcodvend = frmBuscar.resultado
        txtdesvend = frmBuscar.resultado(2)
    End If
End Sub

Private Sub cmdAceptar_Click()
Dim codigodesde As Long, codigohasta As Long
Dim str As String, str2 As String
        
    If optvend = True And txtcodvend = "" Then
        MsgBox "debe ingresar un vendedor"
        Exit Sub
    End If
    
    If optcodigo = False And optDesc = False Then
        MsgBox "debe ingresar un tipo de orden para realizar el listado"
        Exit Sub
    End If
        
    codigodesde = 1
    codigohasta = 9999
    
    If optelegir = True Then
        codigodesde = val(txtdesde)
        codigohasta = val(txthasta)
    End If
    
    If optcodigo = True Then
        If optCliente = False Then
            
            str = "select clientes.*, usuarios.descripcion as des " _
            & "from clientes inner join usuarios on clientes.vendedor = usuarios.codigo " _
            & "where clientes.vendedor =" & Trim(txtcodvend) & " order by clientes.codigo"
            
            str2 = "select CLIENTES.CODIGO,CLIENTES.DESCRIPCION,CLIENTES.DIRECCION,CLIENTES.TELEFONO,CLIENTES.CUIT " _
            & "from clientes inner join usuarios on clientes.vendedor = usuarios.codigo " _
            & "where clientes.vendedor =" & Trim(txtcodvend) & " order by clientes.codigo"
            
            RptClientes.lblfecha = Date
            RptClientes.lblTitulo = "LISTADO DE CLIENTES X VENDEDOR"
            RptClientes.data1.Connection = DataEnvironment1.Sistema
            RptClientes.data1.Source = str
            RptClientes.Show
            
            
'            daTaenvironment1.LisClientesPorCodigo codigodesde, codigohasta
'            rptClientesElijoCodigo.Show vbModal
'            daTaenvironment1.rsLisClientesPorCodigo.Close
        Else
            If opttodos = True Then
                
                
                str = "select clientes.*,usuarios.descripcion " _
                & "as des from clientes inner join usuarios " _
                & "on clientes.vendedor = usuarios.codigo where " _
                & "clientes.codigo >=" & codigodesde & " and clientes.codigo <=" & codigohasta & " order by clientes.codigo"
                
                str2 = "select  CLIENTES.CODIGO,CLIENTES.DESCRIPCION,CLIENTES.DIRECCION,CLIENTES.TELEFONO,CLIENTES.CUIT  " _
                & " from clientes inner join usuarios " _
                & "on clientes.vendedor = usuarios.codigo where " _
                & "clientes.codigo >=" & codigodesde & " and clientes.codigo <=" & codigohasta & " order by clientes.codigo"
                
                RptClientes.lblfecha = Date
                RptClientes.lblTitulo = "LISTADO DE CLIENTES X CODIGO"
                RptClientes.data1.Connection = DataEnvironment1.Sistema
                RptClientes.data1.Source = str
                RptClientes.Show
                
'                daTaenvironment1.LisClientesPorVendedorPorCodigo txtcodvend
'                rptClientesElijoVendedorPorCodigo.Sections("Encabezado").Controls("lblvendedor").Caption = txtdesvend
'                rptClientesElijoVendedorPorCodigo.Show vbModal
'                daTaenvironment1.rsLisClientesPorVendedorPorCodigo.Close
            Else
                If optcateg = True Then
                    If txtcateg <> "" Then
                        str = " select clientes.*, usuarios.descripcion from clientes " _
                        & "inner join usuarios on clientes.vendedor = usuarios.codigo where " _
                        & "clientes.categoria =" & val(txtcateg) & " order by clientes.codigo"
                        
                        str2 = " select   CLIENTES.CODIGO,CLIENTES.DESCRIPCION,CLIENTES.DIRECCION,CLIENTES.TELEFONO,CLIENTES.CUIT   from clientes " _
                        & "inner join usuarios on clientes.vendedor = usuarios.codigo where " _
                        & "clientes.categoria =" & val(txtcateg) & " order by clientes.codigo"
                        
                        RptClientes.lblfecha = Date
                        RptClientes.lblTitulo = "LISTADO DE CLIENTES X CATEGORIA"
                        RptClientes.data1.Connection = DataEnvironment1.Sistema
                        RptClientes.data1.Source = str
                        RptClientes.Show
                        
    '                    daTaenvironment1.LisClientesCategPorCodigo Val(txtcateg)
    '                    rptClientesCategPorCodigo.Sections("Medio").Controls("lblcateg").Caption = txtdescateg
    '                    rptClientesCategPorCodigo.Show vbModal
    '                    daTaenvironment1.rsLisClientesCategPorCodigo.Close
                    Else
                        MsgBox "Debe ingresar una categoría"
                    End If
                Else
                    str = "select clientes.*,usuarios.descripcion " _
                    & "as des from clientes inner join usuarios " _
                    & "on clientes.vendedor = usuarios.codigo where " _
                    & "clientes.codigo >=" & codigodesde & " and clientes.codigo <=" & codigohasta & " order by clientes.codigo"
                    
                    str2 = "select CLIENTES.CODIGO,CLIENTES.DESCRIPCION,CLIENTES.DIRECCION,CLIENTES.TELEFONO,CLIENTES.CUIT " _
                    & " from clientes inner join usuarios " _
                    & "on clientes.vendedor = usuarios.codigo where " _
                    & "clientes.codigo >=" & codigodesde & " and clientes.codigo <=" & codigohasta & " order by clientes.codigo"
                    
                    RptClientes.lblfecha = Date
                    RptClientes.lblTitulo = "LISTADO DE CLIENTES X CODIGO"
                    RptClientes.data1.Connection = DataEnvironment1.Sistema
                    RptClientes.data1.Source = str
                    RptClientes.Show
                End If
            End If
        End If
    Else
        If optCliente = False Then
            str = "select clientes.*, usuarios.codigo from clientes inner " _
            & "join usuarios on clientes.vendedor = usuarios.codigo " _
            & "where clientes.vendedor =" & Trim(txtcodvend) & " order by clientes.codigo"
            
            str2 = "select  CLIENTES.CODIGO,CLIENTES.DESCRIPCION,CLIENTES.DIRECCION,CLIENTES.TELEFONO,CLIENTES.CUIT from clientes inner " _
            & "join usuarios on clientes.vendedor = usuarios.codigo " _
            & "where clientes.vendedor =" & Trim(txtcodvend) & " order by clientes.codigo"
            
            RptClientes.lblfecha = Date
            RptClientes.lblTitulo = "LISTADO DE CLIENTES X DESCRIPCION X VENDEDOR"
            RptClientes.data1.Connection = DataEnvironment1.Sistema
            RptClientes.data1.Source = str
            RptClientes.Show
                
            
            
'                daTaenvironment1.LisClientesPorDescripcion codigodesde, codigohasta
'                rptClientesElijoDescripcion.Show vbModal
'                daTaenvironment1.rsLisClientesPorDescripcion.Close
        Else
            If opttodos = True Then
                       str = "select clientes.*, usuarios.descripcion from clientes inner join " _
                        & "usuarios on clientes.vendedor = usuarios.codigo where " _
                        & "clientes.codigo >=" & codigodesde & " and clientes.codigo <=" & codigohasta & " order by clientes.descripcion"
                        
                       str2 = "select  CLIENTES.CODIGO,CLIENTES.DESCRIPCION,CLIENTES.DIRECCION,CLIENTES.TELEFONO,CLIENTES.CUIT  from clientes inner join " _
                        & "usuarios on clientes.vendedor = usuarios.codigo where " _
                        & "clientes.codigo >=" & codigodesde & " and clientes.codigo <=" & codigohasta & " order by clientes.descripcion"
                        
                        
                        RptClientes.lblfecha = Date
                        RptClientes.lblTitulo = "LISTADO DE CLIENTES X DESCRIPCION"
                        RptClientes.data1.Connection = DataEnvironment1.Sistema
                        RptClientes.data1.Source = str
                        RptClientes.Show
                
'                    daTaenvironment1.LisClientesPorVendedorPorDescripcion txtcodvend
'                    rptClientesElijoVendedorPorDescripcion.Sections("Encabezado").Controls("lblvendedor").Caption = txtdesvend
'                    rptClientesElijoVendedorPorDescripcion.Show vbModal
'                    daTaenvironment1.rsLisClientesPorVendedorPorDescripcion.Close
            Else
                If optcateg = True Then
                    If txtcateg <> "" Then
                        str = "select clientes.*, usuarios.descripcion from clientes " _
                        & "inner join usuarios on clientes.vendedor = usuarios.codigo " _
                        & "where clientes.categoria =" & val(txtcateg) & " order by clientes.descripcion"
                        
                        str2 = "select  CLIENTES.CODIGO,CLIENTES.DESCRIPCION,CLIENTES.DIRECCION,CLIENTES.TELEFONO,CLIENTES.CUIT  from clientes " _
                        & "inner join usuarios on clientes.vendedor = usuarios.codigo " _
                        & "where clientes.categoria =" & val(txtcateg) & " order by clientes.descripcion"
                        
                        RptClientes.lblfecha = Date
                        RptClientes.lblTitulo = "LISTADO DE CLIENTES X DESCRIPCION X CATEGORIA"
                        RptClientes.data1.Connection = DataEnvironment1.Sistema
                        RptClientes.data1.Source = str
                        RptClientes.Show
                        
    '                        daTaenvironment1.LisClientesCategPorDescripcion Val(txtcateg)
    '                        'rptClientesCategPorDescripcion.Sections("Medio").Controls("lblcateg").Caption = txtdescateg
    '                        rptClientesCategPorDescripcion.Show vbModal
    '                        daTaenvironment1.rsLisClientesCategPorDescripcion.Close
                    Else
                        MsgBox "Debe ingresar una categoría"
                    End If
                Else
                    str = "select clientes.*, usuarios.descripcion from clientes inner join " _
                    & "usuarios on clientes.vendedor = usuarios.codigo where " _
                    & "clientes.codigo >=" & codigodesde & " and clientes.codigo <=" & codigohasta & " order by clientes.descripcion"
                    
                    str2 = "select  CLIENTES.CODIGO,CLIENTES.DESCRIPCION,CLIENTES.DIRECCION,CLIENTES.TELEFONO,CLIENTES.CUIT  from clientes inner join " _
                    & "usuarios on clientes.vendedor = usuarios.codigo where " _
                    & "clientes.codigo >=" & codigodesde & " and clientes.codigo <=" & codigohasta & " order by clientes.descripcion"
                    
                    RptClientes.lblfecha = Date
                    RptClientes.lblTitulo = "LISTADO DE CLIENTES X DESCRIPCION"
                    RptClientes.data1.Connection = DataEnvironment1.Sistema
                    RptClientes.data1.Source = str
                    RptClientes.Show
                End If
            End If
        End If
        
    End If
    
    LlenarGrilla gClientes, str2, False
    ucXls1.Visible = True
End Sub

Private Sub cmdCancelar_Click()
    LimpioControles
    VerCliente (False)
    VerVendedor (False)
    VerTextos (False)
    VerOrden (False)
End Sub

Private Sub LimpioControles()
    txtdesde = ""
    txtdesdesde = ""
    txthasta = ""
    txtdeshasta = ""
    txtcateg = ""
    txtdescateg = ""
    cargar = ""
    optcodigo = False
    optDesc = False
    optelegir = False
    opttodos = False
    'optCliente = False
    'optvend = False
End Sub

Private Sub VerTextos(habilito As Boolean)
    txtdesde.Visible = habilito
    txtdesdesde.Visible = habilito
    txthasta.Visible = habilito
    txtdeshasta.Visible = habilito
    cmbdesde.Visible = habilito
    cmbhasta.Visible = habilito
    lbldesde.Visible = habilito
    lblhasta.Visible = habilito
    lblcateg.Visible = habilito
    txtcateg.Visible = habilito
    cmbcateg.Visible = habilito
    txtdescateg.Visible = habilito
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub


Private Sub Form_Load()
gClientes.rows = 1
ucXls1.ini gClientes, "c:\CLIENTES.XLS", "LISTADO DE CLIENTES"
End Sub

Private Sub optcateg_Click()
    VerVendedor (False)
'    VerCliente (False)
    VerTextos (False)
    VerCateg (True)
End Sub

Private Sub optcateg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub optcliente_Click()
    VerCliente (True)
    VerVendedor (False)
    VerOrden (True)
    VerCateg (False)
    LimpioControles
    optCliente.Value = True
End Sub

Private Sub optcliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub optcodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub


Private Sub optDesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub optelegir_Click()
    VerTextos (True)
    VerCateg (False)
End Sub

Private Sub optelegir_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub opttodos_Click()
    VerTextos (False)
    VerCateg (False)
End Sub

Private Sub opttodos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub optvend_Click()
    VerCateg (False)
    VerVendedor (True)
    VerCliente (False)
    VerOrden (True)
    LimpioControles
    optvend.Value = True
    txtdesde.Visible = False
    txtdesdesde.Visible = False
    txthasta.Visible = False
    txtdeshasta.Visible = False
    cmbdesde.Visible = False
    cmbhasta.Visible = False
    lbldesde.Visible = False
    lblhasta.Visible = False
End Sub

Private Sub optvend_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub txtcateg_GotFocus()
    txtcateg.SelStart = 0
    txtcateg.SelLength = Len(txtcateg.Text)
End Sub

Private Sub txtcateg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub txtdesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtdesde_LostFocus()
    If IsNumeric(txtdesde) Then
        txtdesdesde = ObtenerDescripcion("Prov", val(txtdesde))
        If txtdesdesde = "" Then
            MsgBox "Codigo de Cliente incorrecto"
            txtdesde.SetFocus
        End If
    Else
        If txtdesde <> "" Then
            MsgBox "Codigo de Cliente incorrecto"
            txtdesde = "0"
            txtdesde.SetFocus
        End If
    End If
End Sub

Private Sub txthasta_GotFocus()
    txthasta.SelStart = 0
    txthasta.SelLength = Len(txthasta.Text)
End Sub

Private Sub txthasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        KeyAscii = 0
    Else
        If (KeyAscii < 47 Or KeyAscii > 57) And (KeyAscii <> 46 And KeyAscii <> 44) And (KeyAscii <> 7 And KeyAscii <> 8) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txthasta_LostFocus()
    Dim txthastasde
    If IsNumeric(txthasta) Then
        txthastasde = ObtenerDescripcion("Prov", val(txthasta))
        'txthasta
        If txthastasde = "" Then
            MsgBox "Codigo de Cliente incorrecto"
            txthasta.SetFocus
        End If
    Else
        If txthasta <> "" Then
            MsgBox "Codigo de Cliente incorrecto"
            txthasta = "0"
            txthasta.SetFocus
        End If
    End If
End Sub

Public Sub CargarDatos()
    Dim rs As New ADODB.Recordset
    Dim codigo
    
    codigo = val(Trim(Me.Tag))
    
    If cargar = "ProvDesde" Then
        rs.Open "select * from Prov where codigo = " & codigo & " and activo = 1 order by codigo", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
        If Not rs.EOF Then
            txtdede = rs!codigo
            txtdesdesde = rs!DESCRIPCION
        End If
        
        rs.Close
        Set rs = Nothing
    End If

    If cargar = "ProvHasta" Then
        rs.Open "select * from Prov where codigo = " & codigo & " and activo = 1 order by codigo", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
        If Not rs.EOF Then
            txthasta = rs!codigo
            txtdeshasta = rs!DESCRIPCION
        End If
        
        rs.Close
        Set rs = Nothing
    End If
End Sub

Private Sub VerCliente(habilito As Boolean)
    FrameCli.Visible = habilito
    opttodos.Visible = habilito
    optelegir.Visible = habilito
    optcateg.Visible = habilito
End Sub

Private Sub VerVendedor(habilito As Boolean)
    lblvend.Visible = habilito
    txtcodvend.Visible = habilito
    txtdesvend.Visible = habilito
    cmbvend.Visible = habilito
End Sub

Private Sub VerOrden(habilito As Boolean)
    frameorden.Visible = habilito
    optcodigo.Visible = habilito
    optDesc.Visible = habilito
End Sub

Private Sub VerCateg(habilito As Boolean)
    lblcateg.Visible = habilito
    txtcateg.Visible = habilito
    cmbcateg.Visible = habilito
    txtdescateg.Visible = habilito
End Sub

Private Sub txtcateg_Lostfocus()
    If txtcateg <> "" Then
        txtdescateg = ObtenerDescripcion("Categorias", val(txtcateg))
        If txtdescateg = "" Then
            MsgBox "Codigo de Categoría incorrecta"
            txtcateg.SetFocus
        End If
    End If
End Sub

Private Sub ucXls1_Clic(cancel As Boolean)
ucXls1.ini gClientes, "c:\CLIENTES.XLS", "LISTADO DE CLIENTES"
End Sub

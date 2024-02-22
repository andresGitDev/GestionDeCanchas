VERSION 5.00
Begin VB.Form FrmLisProductos 
   Caption         =   "Listado de Productos"
   ClientHeight    =   6330
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9345
   Icon            =   "FrmLisProductos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   9345
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Excel Producto con Stock"
      Height          =   375
      Left            =   120
      TabIndex        =   42
      Top             =   5880
      Width           =   2535
   End
   Begin VB.Frame FrameGrupoSubgrupo 
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
      Height          =   2820
      Left            =   240
      TabIndex        =   23
      Top             =   1920
      Visible         =   0   'False
      Width           =   1815
      Begin VB.OptionButton opts 
         Caption         =   "SubGrupo"
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
         TabIndex        =   25
         Top             =   2040
         Width           =   1455
      End
      Begin VB.OptionButton optg 
         Caption         =   "Grupo"
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
         TabIndex        =   24
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Frame Framecodigo 
      Caption         =   "Código"
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
      Left            =   2100
      TabIndex        =   18
      Top             =   240
      Visible         =   0   'False
      Width           =   7035
      Begin VB.OptionButton opttodoscodigo 
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
         Left            =   2460
         TabIndex        =   20
         Top             =   180
         Width           =   1455
      End
      Begin VB.OptionButton optelegircodigo 
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
         Height          =   255
         Left            =   3960
         TabIndex        =   19
         Top             =   180
         Width           =   1455
      End
      Begin Gestion.ucCoDe uProductoDesde 
         Height          =   315
         Left            =   840
         TabIndex        =   28
         Top             =   540
         Width           =   6075
         _extentx        =   9975
         _extenty        =   556
         codigowidth     =   1000
      End
      Begin Gestion.ucCoDe uProductoHasta 
         Height          =   315
         Left            =   840
         TabIndex        =   29
         Top             =   960
         Width           =   6075
         _extentx        =   9975
         _extenty        =   556
         codigowidth     =   1000
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
         Left            =   120
         TabIndex        =   22
         Top             =   540
         Width           =   615
      End
      Begin VB.Label lblhasta 
         Caption         =   "Hasta"
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
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   615
      End
   End
   Begin VB.Frame Framesubgrupo 
      Caption         =   "SubGrupo"
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
      Left            =   2100
      TabIndex        =   13
      Top             =   3360
      Visible         =   0   'False
      Width           =   7035
      Begin VB.ComboBox cboSubgrupoHasta 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "FrmLisProductos.frx":08CA
         Left            =   2220
         List            =   "FrmLisProductos.frx":08CC
         Style           =   2  'Dropdown List
         TabIndex        =   41
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox txtSubGrupoHasta 
         Height          =   315
         Left            =   1380
         TabIndex        =   40
         Top             =   840
         Width           =   795
      End
      Begin VB.CommandButton cmdSubgrupoHasta 
         Caption         =   "SG"
         Height          =   315
         Left            =   900
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   840
         Width           =   435
      End
      Begin VB.ComboBox cboSubgrupoDesde 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "FrmLisProductos.frx":08CE
         Left            =   2220
         List            =   "FrmLisProductos.frx":08D0
         Style           =   2  'Dropdown List
         TabIndex        =   35
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   420
         Width           =   2775
      End
      Begin VB.TextBox txtSubGrupoDesde 
         Height          =   315
         Left            =   1380
         TabIndex        =   34
         Top             =   420
         Width           =   795
      End
      Begin VB.CommandButton cmdSubgrupoDesde 
         Caption         =   "SG"
         Height          =   315
         Left            =   900
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   420
         Width           =   435
      End
      Begin VB.OptionButton opttodossubgrupo 
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
         Left            =   2520
         TabIndex        =   15
         Top             =   180
         Width           =   1455
      End
      Begin VB.OptionButton optelegirsubgrupos 
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
         Left            =   4140
         TabIndex        =   14
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label lblsubgruposhasta 
         Caption         =   "Hasta"
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
         Left            =   135
         TabIndex        =   27
         Top             =   870
         Width           =   615
      End
      Begin VB.Label lblsubgruposdesde 
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
         Left            =   120
         TabIndex        =   17
         Top             =   435
         Width           =   615
      End
   End
   Begin VB.Frame Framegrupo 
      Caption         =   "Grupo"
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
      Left            =   2100
      TabIndex        =   10
      Top             =   1920
      Visible         =   0   'False
      Width           =   7035
      Begin VB.ComboBox cboGrupoHasta 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2220
         Style           =   2  'Dropdown List
         TabIndex        =   38
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   900
         Width           =   2775
      End
      Begin VB.TextBox txtGrupoHasta 
         Height          =   315
         Left            =   1380
         TabIndex        =   37
         Top             =   900
         Width           =   795
      End
      Begin VB.CommandButton cmdGrupoHasta 
         Caption         =   "G"
         Height          =   315
         Left            =   900
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   900
         Width           =   435
      End
      Begin VB.ComboBox cboGrupoDesde 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2220
         Style           =   2  'Dropdown List
         TabIndex        =   32
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox txtGrupoDesde 
         Height          =   315
         Left            =   1380
         TabIndex        =   31
         Top             =   480
         Width           =   795
      End
      Begin VB.CommandButton cmdGrupoDesde 
         Caption         =   "G"
         Height          =   315
         Left            =   900
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   480
         Width           =   435
      End
      Begin VB.OptionButton optelegirgrupos 
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
         Left            =   4020
         TabIndex        =   12
         Top             =   180
         Width           =   1455
      End
      Begin VB.OptionButton opttodosgrupo 
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
         Left            =   2520
         TabIndex        =   11
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label lblgruposhasta 
         Caption         =   "Hasta"
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
         Left            =   120
         TabIndex        =   26
         Top             =   900
         Width           =   615
      End
      Begin VB.Label lblgruposdesde 
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
         Left            =   120
         TabIndex        =   16
         Top             =   540
         Width           =   615
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "2"
      Top             =   5865
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "1"
      Top             =   5865
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "3"
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox cargar 
      Height          =   285
      Left            =   6240
      TabIndex        =   9
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
      Height          =   660
      Left            =   240
      TabIndex        =   6
      Top             =   5040
      Width           =   8895
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
         Height          =   195
         Left            =   5400
         TabIndex        =   8
         Top             =   240
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
         Left            =   1320
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame framenota 
      Caption         =   "Producto"
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
      Left            =   240
      TabIndex        =   0
      Tag             =   "0"
      Top             =   240
      Width           =   1815
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
         Left            =   240
         TabIndex        =   5
         Top             =   360
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
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00800000&
      Height          =   855
      Left            =   120
      Top             =   4920
      Width           =   9135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BorderColor     =   &H00800000&
      Height          =   3015
      Left            =   120
      Top             =   1860
      Width           =   9135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00800000&
      Height          =   1575
      Left            =   120
      Top             =   120
      Width           =   9135
   End
End
Attribute VB_Name = "FrmLisProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' 10/01/05

Private grupoDesde As LiCodigo
Private grupoHasta As LiCodigo
Private subGrupoDesde As LiCodigo
Private subGrupoHasta As LiCodigo

'Private Sub cmbdesde_Click()
'    FrmHelp.Show
'    CargarHelpLisProductos "Producto", "Codigo", "codigo" ', "descripcion ", "descripcion"
'    FrmHelp.Tag = Me.Name
'    cargar = "ProdDesde"
'End Sub

'Private Sub cmbgruposdesde_Click()
'    FrmHelp.Show
'    CargarHelpLisProductos "Producto", "Grupo", "grupo", "grupo"
'    FrmHelp.Tag = Me.Name
'    cargar = "ProdDesdeGrupo"
'End Sub
'
'Private Sub cmbgruposhasta_Click()
'    FrmHelp.Show
'    CargarHelpLisProductos "Producto", "Grupo", "grupo", "grupo"
'    FrmHelp.Tag = Me.Name
'    cargar = "ProdHastaGrupo"
'End Sub

'Private Sub cmbhasta_Click()
'    FrmHelp.Show
'    CargarHelpLisProductos "Producto", "Codigo", "codigo", ""
'    FrmHelp.Tag = Me.Name
'    cargar = "ProdHasta"
'End Sub

'Private Sub cmbsubgruposdesde_Click()
'    FrmHelp.Show
'    CargarHelpLisProductos "Producto", "SubGrupo", "subgrupo", "subgrupo"
'    FrmHelp.Tag = Me.Name
'    cargar = "ProdDesdeSubGrupo"
'End Sub
'Private Sub cmbsubgruposhasta_Click()
'    FrmHelp.Show
'    CargarHelpLisProductos "Producto", "SubGrupo", "subgrupo", "subgrupo"
'    FrmHelp.Tag = Me.Name
'    cargar = "ProdHastaSubGrupo"
'End Sub

Private Sub cmdaceptar_Click()
Dim FiltrarStock As Boolean
Dim Consulta As String
Dim Orden As String
Dim SubDividir As Boolean 'variable utilizada para saber si divido el informe en grupos y subgrupos
Dim rs As New ADODB.Recordset
    
    ' 16/9/4 agregue... espero q no sean controles
    Dim codigodesde, codigohasta, gruposDesde, GruposHasta
    
    If uProductoDesde.codigo <> "" Then
        codigodesde = uProductoDesde.codigo
    Else
        rs.Open "select min(distinct(codigo)) as min from producto", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
        codigodesde = rs!min
        Set rs = Nothing
    End If
    If uProductoHasta.codigo <> "" Then
        codigohasta = uProductoHasta.codigo
    Else
        rs.Open "select max(distinct(codigo)) as maxi from producto", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
        codigohasta = rs!maxi
        Set rs = Nothing
    End If
    If txtGrupoDesde.Text <> "" Then
        gruposDesde = txtGrupoDesde.Text
    Else
        rs.Open "select min(distinct(grupo)) as grupo from producto", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
        gruposDesde = Trim(rs!grupo)
        Set rs = Nothing
    End If
    If txtGrupoHasta.Text <> "" Then
        GruposHasta = txtGrupoHasta.Text
    Else
        rs.Open "select max(distinct(grupo)) as grupo2 from producto", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
        GruposHasta = Trim(rs!grupo2)
        Set rs = Nothing
    End If
        
    If optcodigo = False And optDesc = False Then
        MsgBox "debe ingresar un tipo de orden para realizar el listado"
        Exit Sub
    End If
        
    If MsgBox("¿ Desea ver el Stock Actual ?", vbYesNo, "Atencion") = vbYes Then
        FiltrarStock = MsgBox("¿ Desea Filtrar los Productos sin Stock ?", vbYesNo, "Atencion") = vbYes
        If optcodigo.Value Then
            Orden = "GRUPO, SUBGRUPO, Codigo"
            SubDividir = True
        Else
            Orden = "Descripcion"
            SubDividir = False
        End If
        
        With RptProductosConStock
        
        If MsgBox("¿ Desea ver el stock separado por depositos ?", vbYesNo + vbDefaultButton2, "Atencion") = vbYes Then
            Consulta = "Select GRUPO, SUBGRUPO, CODIGO, DESCRIPCION, EXISTENCIA, OBSERVACIONES, DEP1, DEP2, DEP3, DEP4,EXISTENCIACALCULADA,RESERVACALCULADA,iva " & _
                        "From PRODUCTO Where ACTIVO = 1"
        Else
            Consulta = "Select GRUPO, SUBGRUPO, CODIGO, DESCRIPCION, EXISTENCIA,EXISTENCIACALCULaDA,RESERVACALCULADA,iva From PRODUCTO Where ACTIVO = 1"
            
            .lblDep1.Visible = False
            .lblDep2.Visible = False
            .lblDep3.Visible = False
            .lblDep4.Visible = False
            .fieDep1.Visible = False
            .fieDep2.Visible = False
            .fieDep3.Visible = False
            .fieDep4.Visible = False
            .fieTotal1.Visible = False
            .FieTotal2.Visible = False
            .fieTotal3.Visible = False
            .fieTotal4.Visible = False
        End If
        
        If FiltrarStock Then Consulta = Consulta & " AND (EXISTENCIA > 0 or ExistenciaCalculada >0 ) "
        
        If optelegir.Value And optelegircodigo.Value Then
            If uProductoDesde.codigo = "" Or uProductoHasta.codigo = "" Then
                che "Debe Ingresar un codigo de producto donde comenzar y otro donde terminar."
                Exit Sub
            Else
                'aca va el desde y hasta del codigo del producto
            
                Consulta = Consulta & " AND (CODIGO >= '" & Trim(uProductoDesde.codigo) & "' AND CODIGO <= '" & Trim(uProductoHasta.codigo) & "')"
            
            End If
        End If
'            If Trim$(txtdesde) = "" Or Trim$(txthasta) = "" Then
'                MsgBox "Debe Ingresar un codigo de producto donde comenzar y otro donde terminar.", vbOKOnly, "Atencion"
'                Exit Sub
'            Else
'                Consulta = Consulta & " AND (CODIGO >= '" & Trim(txtdesde) & "' AND CODIGO <= '" & Trim(txthasta) & "')"
'            End If
'        End If
        
        If optg.Value And optelegirgrupos.Value Then
'            If txtdesdegrupos <> "" And txthastagrupos <> "" Then
'                MsgBox "Debe Ingresar un codigo de grupo donde comenzar y otro donde terminar.", vbOKOnly, "Atencion"
'                Exit Sub
'            Else
'                Consulta = Consulta & " and (GRUPO >= '" & Trim(txtdesdegrupos) & "' AND GRUPO <= '" & Trim(txthastagrupos) & "')"
'            End If
            If grupoDesde.codigo = "" Or grupoHasta.codigo = "" Then
                che "Debe Ingresar un codigo de grupo donde comenzar y otro donde terminar."
                Exit Sub
            Else
                Consulta = Consulta & " and (GRUPO >= '" & grupoDesde.codigo & "' AND GRUPO <= '" & grupoHasta.codigo & "')"
            End If
        End If
        
        If opts.Value And optelegirsubgrupos.Value Then
'            If txtdesdesubgrupos <> "" And txthastasubgrupos <> "" Then
'                MsgBox "Debe Ingresar un codigo de subgrupo donde comenzar y otro donde terminar.", vbOKOnly, "Atencion"
'                Exit Sub
'            Else
'                Consulta = Consulta & " and (SUBGRUPO >= '" & Trim(txtdesdesubgrupos) & "' AND SUBGRUPO <= '" & Trim(txthastasubgrupos) & "')"
'            End If
            If subGrupoDesde.codigo = "" Or subGrupoHasta.codigo = "" Then
                che "Debe Ingresar un codigo de subgrupo donde comenzar y otro donde terminar."
                Exit Sub
            Else
                Consulta = Consulta & " and (SUBGRUPO >= '" & subGrupoDesde.codigo & "' AND SUBGRUPO <= '" & subGrupoHasta.codigo & "')"
            End If
        End If
        
        Consulta = Consulta & " Order By " & Orden
        
        If SubDividir Then
            .GroupHeader1.DataField = "GRUPO"
            .GroupHeader2.DataField = "SUBGRUPO"
        End If
        
        .Data.Connection = DataEnvironment1.Sistema
        .Data.Source = Consulta
        
        .lblTitulo.caption = "LISTADO DE PRODUCTOS ORDENADO POR " & Orden
        .lblFechaReporte.caption = Date
        .fieGrupo.DataField = "GRUPO"
        .fieSubGrupo.DataField = "SUBGRUPO"
        .fieCodigo.DataField = "CODIGO"
        .FieDescripcion.DataField = "DESCRIPCION"
        .fieDep1.DataField = "DEP1"
        .fieDep2.DataField = "DEP2"
        .fieDep3.DataField = "DEP3"
        .fieDep4.DataField = "DEP4"
        .FiePrincipal.DataField = "EXISTENCIA"
        .fieexistencia.DataField = "EXISTENCIACALCULADA"
        '.fiereservada.DataField = "RESERVACALCULADA"
        .fiereservada.DataField = "iva"
        
        .fieTotal1.DataField = "DEP1"
        .FieTotal2.DataField = "DEP2"
        .fieTotal3.DataField = "DEP3"
        .fieTotal4.DataField = "DEP4"
        .fieTotal.DataField = "EXISTENCIA"
        .fietotexis.DataField = "EXISTENCIACALCULADA"
        '.fietotreser.DataField = "RESERVACALCULADA"
        .Show
        
        End With
    Else
        If optcodigo = True Then
            'ordeno por codigo C
            'DataEnvironment1.LisProductosPorCodigo codigodesde, codigohasta, gruposDesde, GruposHasta
            
            'DataEnvironment1.dbo_LisProductosXcodigo "C", gruposDesde, GruposHasta, codigodesde, codigohasta
            Traer "lisproductosxcodigo", "C", gruposDesde, GruposHasta, codigodesde, codigohasta
            Set rptProductosElijoCodigo1.DataControl1.Recordset = RStraer
            rptProductosElijoCodigo1.Field6.Text = RStraer.RecordCount
            rptProductosElijoCodigo1.Field7.Text = Date
            rptProductosElijoCodigo1.Show vbModal
            Set RStraer = Nothing
            'DataEnvironment1.rsdbo_LisProductosXcodigo.Close
            
            'DataEnvironment1.rsLisProductosPorCodigo.Close
        Else
            'ordeno por descripcion D
            'DataEnvironment1.LisProductosPorDescripcion codigodesde, codigohasta, gruposDesde, GruposHasta
            
            'DataEnvironment1.dbo_LisProductosXcodigo "D", gruposDesde, GruposHasta, codigodesde, codigohasta
            Traer "LisProductosXcodigo", "D", gruposDesde, GruposHasta, codigodesde, codigohasta
            Set rptProductosElijoCodigo1.DataControl1.Recordset = RStraer
            rptProductosElijoCodigo1.Field6.Text = RStraer.RecordCount
            rptProductosElijoCodigo1.Field7.Text = Date
            rptProductosElijoCodigo1.Label2 = "LISTADO DE PROVEEDORES POR DESCRIPCION"
            rptProductosElijoCodigo1.Show vbModal 'estaba usando rptProductosElijoDescripcion
            Set RStraer = Nothing
            'DataEnvironment1.rsdbo_LisProductosXcodigo.Close
            
            'DataEnvironment1.rsLisProductosPorDescripcion.Close
        End If
    End If
End Sub
Private Sub cmdcancelar_Click()
    LimpioControles
    opttodos.Value = True
    
    uProductoDesde.enabled = False
    uProductoHasta.enabled = False
    cmdGrupoDesde.enabled = False
    txtGrupoDesde.enabled = False
    cboGrupoDesde.enabled = False
    cmdGrupoHasta.enabled = False
    txtGrupoHasta.enabled = False
    cboGrupoHasta.enabled = False
    cmdSubgrupoDesde.enabled = False
    txtSubGrupoDesde.enabled = False
    cboSubgrupoDesde.enabled = False
    cmdSubgrupoHasta.enabled = False
    txtSubGrupoHasta.enabled = False
    cboSubgrupoHasta.enabled = False
End Sub
Private Sub LimpioControles()
    'txtdesde = ""
    'txthasta = ""
    uProductoDesde.clear
    uProductoHasta.clear
    grupoDesde.codigo = ""
    grupoHasta.codigo = ""
    subGrupoDesde.codigo = ""
    subGrupoHasta.codigo = ""
'    cargar = ""
    
    opttodosgrupo = False
    optelegirgrupos = False
    opttodossubgrupo = False
    optelegirsubgrupos = False
    
'    txtdesdegrupos = ""
'    txthastagrupos = ""
'    txtdesdesubgrupos = ""
'    txthastasubgrupos = ""
    
    optcodigo = False
    optDesc = False
    optelegir = False
    opttodos = False
End Sub

Private Sub VerGrupos(habilito As Boolean)
    Framegrupo.Visible = habilito
End Sub
Private Sub VerSubGrupos(habilito As Boolean)
    Framesubgrupo.Visible = habilito
End Sub
Private Sub VerOrden(habilito As Boolean)
    frameorden.Visible = habilito
End Sub
Private Sub VerCodigo(habilito As Boolean)
    Framecodigo.Visible = habilito
End Sub
Private Sub VerGrupoSubGrupo(habilito As Boolean)
    FrameGrupoSubgrupo.Visible = habilito
End Sub
Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Dim FiltrarStock As Boolean
    Dim Consulta As String
    Dim Orden As String
    Dim SubDividir As Boolean 'variable utilizada para saber si divido el informe en grupos y subgrupos
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    
    ' 16/9/4 agregue... espero q no sean controles
    Dim codigodesde, codigohasta, gruposDesde, GruposHasta
    
    If uProductoDesde.codigo <> "" Then
        codigodesde = uProductoDesde.codigo
    Else
        rs.Open "select min(distinct(codigo)) as min from producto", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
        codigodesde = rs!min
        Set rs = Nothing
    End If
    If uProductoHasta.codigo <> "" Then
        codigohasta = uProductoHasta.codigo
    Else
        rs.Open "select max(distinct(codigo)) as maxi from producto", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
        codigohasta = rs!maxi
        Set rs = Nothing
    End If
    If txtGrupoDesde.Text <> "" Then
        gruposDesde = txtGrupoDesde.Text
    Else
        rs.Open "select min(distinct(grupo)) as grupo from producto", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
        gruposDesde = Trim(rs!grupo)
        Set rs = Nothing
    End If
    If txtGrupoHasta.Text <> "" Then
        GruposHasta = txtGrupoHasta.Text
    Else
        rs.Open "select max(distinct(grupo)) as grupo2 from producto", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
        GruposHasta = Trim(rs!grupo2)
        Set rs = Nothing
    End If
        
    If optcodigo = False And optDesc = False Then
        MsgBox "debe ingresar un tipo de orden para realizar el listado"
        Exit Sub
    End If
        
    If MsgBox("¿ Desea ver el Stock Actual ?", vbYesNo, "Atencion") = vbYes Then
        FiltrarStock = MsgBox("¿ Desea Filtrar los Productos sin Stock ?", vbYesNo, "Atencion") = vbYes
        If optcodigo.Value Then
            Orden = "GRUPO, SUBGRUPO, Codigo"
            SubDividir = True
        Else
            Orden = "Descripcion"
            SubDividir = False
        End If
        
        Consulta = "Select GRUPO, SUBGRUPO, CODIGO, DESCRIPCION, EXISTENCIA,EXISTENCIACALCULaDA,RESERVACALCULADA,iva From PRODUCTO Where ACTIVO = 1"
        
        If FiltrarStock Then Consulta = Consulta & " AND (EXISTENCIA > 0 or ExistenciaCalculada >0 ) "
        
        If optelegir.Value And optelegircodigo.Value Then
            If uProductoDesde.codigo = "" Or uProductoHasta.codigo = "" Then
                che "Debe Ingresar un codigo de producto donde comenzar y otro donde terminar."
                Exit Sub
            Else
                'aca va el desde y hasta del codigo del producto
                Consulta = Consulta & " AND (CODIGO >= '" & Trim(uProductoDesde.codigo) & "' AND CODIGO <= '" & Trim(uProductoHasta.codigo) & "')"
            
            End If
        End If
        
        If optg.Value And optelegirgrupos.Value Then
            If grupoDesde.codigo = "" Or grupoHasta.codigo = "" Then
                che "Debe Ingresar un codigo de grupo donde comenzar y otro donde terminar."
                Exit Sub
            Else
                Consulta = Consulta & " and (GRUPO >= '" & grupoDesde.codigo & "' AND GRUPO <= '" & grupoHasta.codigo & "')"
            End If
        End If
        
        If opts.Value And optelegirsubgrupos.Value Then
            If subGrupoDesde.codigo = "" Or subGrupoHasta.codigo = "" Then
                che "Debe Ingresar un codigo de subgrupo donde comenzar y otro donde terminar."
                Exit Sub
            Else
                Consulta = Consulta & " and (SUBGRUPO >= '" & subGrupoDesde.codigo & "' AND SUBGRUPO <= '" & subGrupoHasta.codigo & "')"
            End If
        End If
        
        Consulta = Consulta & " Order By " & Orden
        
        rs2.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        VinculoXl "C:\LisProd.xls", "Listado de productos", , , rs2
        Set rs = Nothing
    Else
        If optcodigo = True Then
            Traer "lisproductosxcodigo", "C", gruposDesde, GruposHasta, codigodesde, codigohasta
            VinculoXl "C:\LisProd.xls", "Listado de productos", , , RStraer
            Set RStraer = Nothing
        Else
            Traer "LisProductosXcodigo", "D", gruposDesde, GruposHasta, codigodesde, codigohasta
            VinculoXl "C:\LisProd.xls", "Listado de productos", , , RStraer
            Set RStraer = Nothing
        End If
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub

Private Sub Form_Load()
    Dim sqldesc As String, sqlbuscar As String
    GeneraExistenciaCalculada
    sqldesc = "select descripcion from producto where codigo = '###' "
    sqlbuscar = "select codigo as [ Codigo                 ], descripcion as [ Descripcion                                                 ] from producto where activo = 1 order by codigo "
    uProductoDesde.ini sqldesc, sqlbuscar, True
    uProductoHasta.ini sqldesc, sqlbuscar, True
    
    Set grupoDesde = New LiCodigo
    Set subGrupoDesde = New LiCodigo
    Set grupoHasta = New LiCodigo
    Set subGrupoHasta = New LiCodigo
    
    grupoDesde.init cboGrupoDesde, txtGrupoDesde, "GruposProducto", False, True, cmdGrupoDesde, "activo = 1 "
    subGrupoDesde.init cboSubgrupoDesde, txtSubGrupoDesde, "SubGruposProducto", False, True, cmdSubgrupoDesde, "activo=1"
    grupoHasta.init cboGrupoHasta, txtGrupoHasta, "GruposProducto", False, True, cmdGrupoHasta, "activo = 1 "
    subGrupoHasta.init cboSubgrupoHasta, txtSubGrupoHasta, "SubGruposProducto", False, True, cmdSubgrupoHasta, "activo=1"
    
    txtGrupoDesde.Text = ""
    txtGrupoHasta.Text = ""
    
    uProductoDesde.enabled = False
    uProductoHasta.enabled = False
    cmdGrupoDesde.enabled = False
    txtGrupoDesde.enabled = False
    cboGrupoDesde.enabled = False
    cmdGrupoHasta.enabled = False
    txtGrupoHasta.enabled = False
    cboGrupoHasta.enabled = False
    cmdSubgrupoDesde.enabled = False
    txtSubGrupoDesde.enabled = False
    cboSubgrupoDesde.enabled = False
    cmdSubgrupoHasta.enabled = False
    txtSubGrupoHasta.enabled = False
    cboSubgrupoHasta.enabled = False
End Sub

Private Sub optelegir_Click()
    VerCodigo True
    VerGrupoSubGrupo True
    opttodoscodigo.Value = False
    optelegircodigo.Value = False
    optg.Value = False
    opts.Value = False
End Sub

Private Sub HabilitoControlesCodigo(habilito As Boolean)
'    txtdesde.Enabled = habilito
'    txthasta.Enabled = habilito
'    cmbdesde.Enabled = habilito
'    cmbhasta.Enabled = habilito
    uProductoDesde.enabled = habilito
    uProductoHasta.enabled = habilito
End Sub

Private Sub HabilitoControlesGrupos(habilito As Boolean)
'    txtdesdegrupos.Enabled = habilito
'    txthastagrupos.Enabled = habilito
'    cmbgruposdesde.Enabled = habilito
'    cmbgruposhasta.Enabled = habilito
    grupoDesde.codigo = ""
    grupoHasta.codigo = ""
    grupoDesde.enabled = habilito
    grupoHasta.enabled = habilito
End Sub
Private Sub optelegircodigo_Click()
    HabilitoControlesCodigo (True)
End Sub
Private Sub optelegirgrupos_Click()
    HabilitoControlesGrupos (True)
End Sub
Private Sub optelegirsubgrupos_Click()
    HabilitoControlesSubGrupos (True)
End Sub
Private Sub optg_Click()
    VerGrupos (True)
    VerSubGrupos (False)
End Sub
Private Sub opts_Click()
    VerSubGrupos (True)
    VerGrupos (False)
End Sub
Private Sub opttodos_Click()
    VerCodigo False
    VerGrupoSubGrupo False
    VerGrupos False
    VerSubGrupos False
End Sub
Private Sub HabilitoControlesSubGrupos(habilito As Boolean)
'    txtdesdesubgrupos.Enabled = habilito
'    txthastasubgrupos.Enabled = habilito
'    cmbsubgruposdesde.Enabled = habilito
'    cmbsubgruposhasta.Enabled = habilito
    subGrupoDesde.enabled = habilito
    subGrupoHasta.enabled = habilito
End Sub
Private Sub opttodoscodigo_Click()
'    cmbdesde.Enabled = False
'    cmbhasta.Enabled = False
'    txtdesde = ""
'    txthasta = ""
    uProductoDesde.enabled = False
    uProductoHasta.enabled = False
    uProductoDesde.clear
    uProductoHasta.clear
End Sub
Private Sub opttodosgrupo_Click()
'    cmbgruposdesde.Enabled = False
'    cmbgruposhasta.Enabled = False
'    txtdesdegrupos = ""
'    txthastagrupos = ""
    grupoDesde.codigo = ""
    grupoHasta.codigo = ""
    grupoDesde.enabled = False
    grupoHasta.enabled = False
End Sub
Private Sub opttodossubgrupo_Click()
'    cmbsubgruposdesde.Enabled = False
'    cmbsubgruposhasta.Enabled = False
'    txtdesdesubgrupos = ""
'    txthastasubgrupos = ""
    subGrupoDesde.codigo = ""
    subGrupoHasta.codigo = ""
    subGrupoDesde.enabled = False
    subGrupoHasta.enabled = False
End Sub

''Private Sub txtdesde_LostFocus()
''    On Error Resume Next
''    Dim prod As String
''    prod = Trim$(txtdesde)
''    If prod > "" Then
''        If ObtenerDescripcionS("Producto", prod) = "" Then
''            che "Código de Producto incorrecto"
''            txtdesde.SetFocus
''        End If
''    End If
''End Sub
''
''Private Sub txthasta_LostFocus()
''    On Error Resume Next
''    Dim prod As String
''    prod = Trim$(txthasta)
''    If prod > "" Then
''        If ObtenerDescripcionS("Producto", prod) = "" Then
''            che "Código de Producto incorrecto"
''            txthasta.SetFocus
''        End If
''    End If
''End Sub

''Public Sub CargarDatos()
''Dim rs As New ADODB.Recordset
''Dim Codigo
''
''    Codigo = Trim(Me.Tag)
''
''''    If cargar = "ProdDesde" Then
''''        rs.Open "select * from Producto where codigo = '" & Codigo & "' and activo = 1 order by codigo", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
''''        If Not rs.EOF Then
''''            txtdesde = rs!Codigo
'''''            txtdesdesde = rs!descripcion
''''        End If
''''
''''        rs.Close
''''        Set rs = Nothing
''''    End If
''''
''''    If cargar = "ProdHasta" Then
''''        rs.Open "select * from Producto where codigo = '" & Codigo & "' and activo = 1 order by codigo", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
''''        If Not rs.EOF Then
''''            txthasta = rs!Codigo
'''''            txtdeshasta = rs!descripcion
''''        End If
''''
''''        rs.Close
''''        Set rs = Nothing
''''    End If
''
''    If cargar = "ProdDesdeGrupo" Then
''        rs.Open "select * from Producto where grupo = '" & Codigo & "' and activo = 1 order by grupo", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
''        If Not rs.EOF Then
''            txtdesdegrupos = rs!grupo
''        End If
''
''        rs.Close
''        Set rs = Nothing
''    End If
''
''    If cargar = "ProdHastaGrupo" Then
''        rs.Open "select * from Producto where grupo = '" & Codigo & "' and activo = 1 order by grupo", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
''        If Not rs.EOF Then
''            txthastagrupos = rs!grupo
''        End If
''
''        rs.Close
''        Set rs = Nothing
''    End If
''
''
''    If cargar = "ProdDesdeSubGrupo" Then
''        rs.Open "select * from Producto where subgrupo = '" & Codigo & "' and activo = 1 order by grupo", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
''        If Not rs.EOF Then
''            txtdesdesubgrupos = rs!SubGrupo
''        End If
''
''        rs.Close
''        Set rs = Nothing
''    End If
''
''    If cargar = "ProdHastaSubGrupo" Then
''        rs.Open "select * from Producto where subgrupo = '" & Codigo & "' and activo = 1 order by grupo", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
''        If Not rs.EOF Then
''            txthastasubgrupos = rs!SubGrupo
''        End If
''
''        rs.Close
''        Set rs = Nothing
''    End If
''
''End Sub

'**********************************************************
'2/5/5
'   Pablo: agrego existencia calculada y reserva
'   yo: dataenvironment1
'3/5/5
'   cambie todos los controles, pero no la logica
'

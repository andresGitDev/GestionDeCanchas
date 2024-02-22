VERSION 5.00
Begin VB.Form frmLisFormulas 
   Caption         =   "Listado de formulas"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9360
   Icon            =   "frmLisFormulas.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   9360
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame framenota 
      Caption         =   "Formulas"
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
      Left            =   180
      TabIndex        =   38
      Tag             =   "0"
      Top             =   225
      Width           =   1815
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
         TabIndex        =   40
         Top             =   840
         Width           =   1455
      End
      Begin VB.OptionButton opttodos 
         Caption         =   "Todas"
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
         TabIndex        =   39
         Top             =   360
         Width           =   1455
      End
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
      Left            =   180
      TabIndex        =   35
      Top             =   5025
      Width           =   8895
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
         TabIndex        =   37
         Top             =   240
         Width           =   1455
      End
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
         TabIndex        =   36
         Top             =   240
         Width           =   1455
      End
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
      Left            =   8220
      Style           =   1  'Graphical
      TabIndex        =   34
      Tag             =   "3"
      Top             =   5940
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
      Left            =   6060
      Style           =   1  'Graphical
      TabIndex        =   33
      Tag             =   "1"
      Top             =   5925
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
      Left            =   7140
      Style           =   1  'Graphical
      TabIndex        =   32
      Tag             =   "2"
      Top             =   5940
      Width           =   975
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
      Left            =   2040
      TabIndex        =   21
      Top             =   1905
      Visible         =   0   'False
      Width           =   7035
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
         TabIndex        =   29
         Top             =   180
         Width           =   1455
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
         TabIndex        =   28
         Top             =   180
         Width           =   1455
      End
      Begin VB.CommandButton cmdGrupoDesde 
         Caption         =   "G"
         Height          =   315
         Left            =   900
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   480
         Width           =   435
      End
      Begin VB.TextBox txtGrupoDesde 
         Height          =   315
         Left            =   1380
         TabIndex        =   26
         Top             =   480
         Width           =   795
      End
      Begin VB.ComboBox cboGrupoDesde 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2220
         Style           =   2  'Dropdown List
         TabIndex        =   25
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   480
         Width           =   2775
      End
      Begin VB.CommandButton cmdGrupoHasta 
         Caption         =   "G"
         Height          =   315
         Left            =   900
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   900
         Width           =   435
      End
      Begin VB.TextBox txtGrupoHasta 
         Height          =   315
         Left            =   1380
         TabIndex        =   23
         Top             =   900
         Width           =   795
      End
      Begin VB.ComboBox cboGrupoHasta 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2220
         Style           =   2  'Dropdown List
         TabIndex        =   22
         TabStop         =   0   'False
         Tag             =   "0"
         Top             =   900
         Width           =   2775
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
         TabIndex        =   31
         Top             =   540
         Width           =   615
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
         TabIndex        =   30
         Top             =   900
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
      Left            =   2040
      TabIndex        =   10
      Top             =   3345
      Visible         =   0   'False
      Width           =   7035
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
         TabIndex        =   18
         Top             =   180
         Width           =   1455
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
         TabIndex        =   17
         Top             =   180
         Width           =   1455
      End
      Begin VB.CommandButton cmdSubgrupoDesde 
         Caption         =   "SG"
         Height          =   315
         Left            =   900
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   420
         Width           =   435
      End
      Begin VB.TextBox txtSubGrupoDesde 
         Height          =   315
         Left            =   1380
         TabIndex        =   15
         Top             =   420
         Width           =   795
      End
      Begin VB.ComboBox cboSubgrupoDesde 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmLisFormulas.frx":08CA
         Left            =   2220
         List            =   "frmLisFormulas.frx":08CC
         Style           =   2  'Dropdown List
         TabIndex        =   14
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   420
         Width           =   2775
      End
      Begin VB.CommandButton cmdSubgrupoHasta 
         Caption         =   "SG"
         Height          =   315
         Left            =   900
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   840
         Width           =   435
      End
      Begin VB.TextBox txtSubGrupoHasta 
         Height          =   315
         Left            =   1380
         TabIndex        =   12
         Top             =   840
         Width           =   795
      End
      Begin VB.ComboBox cboSubgrupoHasta 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmLisFormulas.frx":08CE
         Left            =   2220
         List            =   "frmLisFormulas.frx":08D0
         Style           =   2  'Dropdown List
         TabIndex        =   11
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   840
         Width           =   2775
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
         TabIndex        =   20
         Top             =   435
         Width           =   615
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
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   615
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
      Left            =   2040
      TabIndex        =   3
      Top             =   225
      Visible         =   0   'False
      Width           =   7035
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
         TabIndex        =   5
         Top             =   180
         Width           =   1455
      End
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
         TabIndex        =   4
         Top             =   180
         Width           =   1455
      End
      Begin Gestion.ucCoDe uProductoDesde 
         Height          =   315
         Left            =   840
         TabIndex        =   6
         Top             =   540
         Width           =   6075
         _ExtentX        =   9975
         _ExtentY        =   556
         CodigoWidth     =   1000
      End
      Begin Gestion.ucCoDe uProductoHasta 
         Height          =   315
         Left            =   840
         TabIndex        =   7
         Top             =   960
         Width           =   6075
         _ExtentX        =   9975
         _ExtentY        =   556
         CodigoWidth     =   1000
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
         TabIndex        =   9
         Top             =   960
         Width           =   615
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
         TabIndex        =   8
         Top             =   540
         Width           =   615
      End
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
      Left            =   180
      TabIndex        =   0
      Top             =   1905
      Visible         =   0   'False
      Width           =   1815
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
         TabIndex        =   2
         Top             =   720
         Width           =   1455
      End
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
         TabIndex        =   1
         Top             =   2040
         Width           =   1455
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00800000&
      Height          =   1575
      Left            =   60
      Top             =   105
      Width           =   9135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      Height          =   3015
      Left            =   60
      Top             =   1845
      Width           =   9135
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00800000&
      Height          =   855
      Left            =   60
      Top             =   4905
      Width           =   9135
   End
End
Attribute VB_Name = "frmLisFormulas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private grupoDesde As LiCodigo
Private grupoHasta As LiCodigo
Private subGrupoDesde As LiCodigo
Private subGrupoHasta As LiCodigo

Private Sub cmdAceptar_Click()
    Dim FiltrarStock As Boolean
    'Dim Consulta As String
    Dim ORDEN As String
    Dim SubDividir As Boolean 'variable utilizada para saber si divido el informe en grupos y subgrupos
    
    Dim rsIMPRE As New ADODB.Recordset 'para llevar a impresion
    Dim rsVER As New ADODB.Recordset 'para ver componentes
    Dim rscomp As New ADODB.Recordset 'para ver el precio
    Dim consulta2 As String
    Dim impo As Double
            
    If optcodigo = False And optDesc = False Then
        MsgBox "debe ingresar un tipo de orden para realizar el listado"
        Exit Sub
    End If
        
    'FiltrarStock = MsgBox("¿ Desea Filtrar los Productos sin Stock ?", vbYesNo, "Atencion") = vbYes
    If optcodigo.Value Then
        ORDEN = " P.GRUPO, P.SUBGRUPO, F.Codigo"
        SubDividir = True
    Else
        ORDEN = "P.Descripcion"
        SubDividir = False
    End If
        
    With rptLisFormulas
        
        'Consulta = "select f.codigo,p.descripcion,f.componente,pp.descripcion as descripcion2,f.cantidad,p.grupo,p.subgrupo,r.precio,(f.cantidad*r.precio)as importe from formulas f inner join producto p on f.codigo=p.codigo inner join producto pp on f.componente=pp.codigo inner join relacion_producto_proveedor r on f.componente=r.producto where f.activo=1 and r.id=(select min(id) from relacion_producto_proveedor re where producto=f.componente) "
        consulta2 = "select f.codigo,p.descripcion,f.componente,pp.descripcion as descripcion2,f.cantidad,p.grupo,p.subgrupo,p.precio from formulas f inner join producto p on f.codigo=p.codigo inner join producto pp on f.componente=pp.codigo where f.activo=1 "
                           
        'If FiltrarStock Then Consulta = Consulta & " AND (p.EXISTENCIA > 0 or p.ExistenciaCalculada >0 ) "
        'If FiltrarStock Then consulta2 = consulta2 & " AND (p.EXISTENCIA > 0 or p.ExistenciaCalculada >0 ) "
            
        If optelegir.Value And optelegircodigo.Value Then
            If uProductoDesde.codigo = "" Or uProductoHasta.codigo = "" Then
                che "Debe Ingresar un codigo de producto donde comenzar y otro donde terminar."
                Exit Sub
            Else
                'aca va el desde y hasta del codigo del producto
                'Consulta = Consulta & " AND (f.CODIGO >= '" & Trim(uProductoDesde.codigo) & "' AND f.CODIGO <= '" & Trim(uProductoHasta.codigo) & "')"
                consulta2 = consulta2 & " AND (f.CODIGO >= '" & Trim(uProductoDesde.codigo) & "' AND f.CODIGO <= '" & Trim(uProductoHasta.codigo) & "')"
            End If
        End If
    
        If optg.Value And optelegirgrupos.Value Then
            If grupoDesde.codigo = "" Or grupoHasta.codigo = "" Then
                che "Debe Ingresar un codigo de grupo donde comenzar y otro donde terminar."
                Exit Sub
            Else
                'Consulta = Consulta & " and (p.GRUPO >= '" & grupoDesde.codigo & "' AND p.GRUPO <= '" & grupoHasta.codigo & "')"
                consulta2 = consulta2 & " and (p.GRUPO >= '" & grupoDesde.codigo & "' AND p.GRUPO <= '" & grupoHasta.codigo & "')"
            End If
        End If
            
        If opts.Value And optelegirsubgrupos.Value Then
            If subGrupoDesde.codigo = "" Or subGrupoHasta.codigo = "" Then
                che "Debe Ingresar un codigo de subgrupo donde comenzar y otro donde terminar."
                Exit Sub
            Else
                'Consulta = Consulta & " and (p.SUBGRUPO >= '" & subGrupoDesde.codigo & "' AND p.SUBGRUPO <= '" & subGrupoHasta.codigo & "')"
                consulta2 = consulta2 & " and (p.SUBGRUPO >= '" & subGrupoDesde.codigo & "' AND p.SUBGRUPO <= '" & subGrupoHasta.codigo & "')"
            End If
        End If
            
        'Consulta = Consulta & " Order By " & orden
        consulta2 = consulta2 & " Order By " & ORDEN
            
            '*********************************************
            
        Set rsIMPRE = New ADODB.Recordset
    
        rsIMPRE.Fields.Append "CODIGO", adChar, 24, adFldUpdatable
        rsIMPRE.Fields.Append "DESCRIPCION", adChar, 65, adFldUpdatable
        rsIMPRE.Fields.Append "COMPONENTE", adChar, 24, adFldUpdatable
        rsIMPRE.Fields.Append "DESCRIPCION2", adChar, 65, adFldUpdatable
        rsIMPRE.Fields.Append "CANTIDAD", adDouble, 8, adFldUpdatable
        rsIMPRE.Fields.Append "GRUPO", adChar, 3, adFldUpdatable
        rsIMPRE.Fields.Append "SUBGRUPO", adChar, 3, adFldUpdatable
        rsIMPRE.Fields.Append "PRECIO", adDouble, 8, adFldUpdatable
        rsIMPRE.Fields.Append "IMPORTE", adDouble, 10, adFldUpdatable
        ' Utilice el tipo de cursor Keyset para permitir la actualización
        ' de los registros.
        rsIMPRE.CursorType = adOpenKeyset
        rsIMPRE.LockType = adLockOptimistic
        rsIMPRE.Open
                    
        rsVER.Open consulta2, DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
        If rsVER.EOF = True And rsVER.BOF = True Then
            MsgBox "No se han encontrado datos, intente filtrar utilizando otras opciones"
        Else
            rsVER.MoveFirst
            Do While Not rsVER.EOF
                
                
                rscomp.Open "select id,producto,precio from relacion_producto_proveedor where  id=(select min(id) from relacion_producto_proveedor where producto='" & rsVER!componente & "')", DataEnvironment1.Sistema, adOpenDynamic, adLockBatchOptimistic
                If Not IsNull(rscomp!ID) Then
                    If rscomp.EOF = True And rscomp.BOF = True Then
                        rsIMPRE.AddNew
                        rsIMPRE!codigo = rsVER!codigo
                        rsIMPRE!DESCRIPCION = rsVER!DESCRIPCION
                        rsIMPRE!componente = rsVER!componente
                        rsIMPRE!descripcion2 = rsVER!descripcion2
                        rsIMPRE!cantidad = rsVER!cantidad
                        rsIMPRE!grupo = rsVER!grupo
                        rsIMPRE!SubGrupo = rsVER!SubGrupo
                        rsIMPRE!precio = rsVER!precio
                        impo = s2n(rsVER!precio * rsVER!cantidad, 4)
                        rsIMPRE!importe = impo
                        rsIMPRE.Update
                    Else
                        If (Not IsNull(rscomp!precio) Or Not IsEmpty(rscomp!precio)) And rscomp!precio > 0 Then
                            'paso los datos al rs
                            rsIMPRE.AddNew
                            rsIMPRE!codigo = rsVER!codigo
                            rsIMPRE!DESCRIPCION = rsVER!DESCRIPCION
                            rsIMPRE!componente = rsVER!componente
                            rsIMPRE!descripcion2 = rsVER!descripcion2
                            rsIMPRE!cantidad = rsVER!cantidad
                            rsIMPRE!grupo = rsVER!grupo
                            rsIMPRE!SubGrupo = rsVER!SubGrupo
                            rsIMPRE!precio = rscomp!precio
                            impo = s2n(rscomp!precio * rsVER!cantidad, 4)
                            rsIMPRE!importe = impo
                            rsIMPRE.Update
                        Else
                            If rscomp!precio = 0 Then
                                rsIMPRE.AddNew
                                rsIMPRE!codigo = rsVER!codigo
                                rsIMPRE!DESCRIPCION = rsVER!DESCRIPCION
                                rsIMPRE!componente = rsVER!componente
                                rsIMPRE!descripcion2 = rsVER!descripcion2
                                rsIMPRE!cantidad = rsVER!cantidad
                                rsIMPRE!grupo = rsVER!grupo
                                rsIMPRE!SubGrupo = rsVER!SubGrupo
                                rsIMPRE!precio = rsVER!precio
                                impo = s2n(rsVER!precio * rsVER!cantidad, 4)
                                rsIMPRE!importe = impo
                                rsIMPRE.Update
                            End If
                        End If
                    End If
                End If
                Set rscomp = Nothing
                rsVER.MoveNext
            Loop
                
                '*******************************************
                
            If SubDividir Then
                .GroupHeader1.DataField = "GRUPO"
                .GroupHeader2.DataField = "SUBGRUPO"
            End If
                
            .GroupHeader3.DataField = "CODIGO"
            .GroupFooter3.DataField = "CODIGO"
                        
            rsIMPRE.MoveFirst
            Set .Data.Recordset = rsIMPRE
                
            .lblTitulo.caption = "LISTADO DE FORMULAS ORDENADO POR " & ORDEN
            .lblFechaReporte.caption = Date
            .fieGrupo.DataField = "GRUPO"
            .fieSubGrupo.DataField = "SUBGRUPO"
            .fieCod.DataField = "CODIGO"
            .fieDesc.DataField = "DESCRIPCION"
            .fieCodigo.DataField = "COMPONENTE"
            .fieDescripcion.DataField = "DESCRIPCION2"
            .FiePrincipal.DataField = "CANTIDAD"
            .fieexistencia.DataField = "PRECIO"
            .fiereservada.DataField = "IMPORTE"
            .fieTotal.DataField = "CANTIDAD"
            .fietotreser.DataField = "IMPORTE"
            .Show
            
        End If
    End With
    Set rsIMPRE = Nothing
    Set rsVER = Nothing
End Sub
Private Sub cmdCancelar_Click()
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

    uProductoDesde.clear
    uProductoHasta.clear
    grupoDesde.codigo = ""
    grupoHasta.codigo = ""
    subGrupoDesde.codigo = ""
    subGrupoHasta.codigo = ""
    
    opttodosgrupo = False
    optelegirgrupos = False
    opttodossubgrupo = False
    optelegirsubgrupos = False
    
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
Private Sub Form_KeyPress(KeyAscii As Integer)
    FrmKeyPress KeyAscii, True, True
End Sub

Private Sub Form_Load()
    Dim sqldesc As String, sqlbuscar As String
    GeneraExistenciaCalculada
    sqldesc = "select descripcion from producto where codigo = '###' "
    sqlbuscar = "select distinct(f.codigo) as [ Codigo                 ], p.descripcion as [ Descripcion                                                 ] from formulas f inner join producto p on f.codigo=p.codigo where f.activo = 1 order by f.codigo "
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
    uProductoDesde.enabled = habilito
    uProductoHasta.enabled = habilito
End Sub

Private Sub HabilitoControlesGrupos(habilito As Boolean)
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
    subGrupoDesde.enabled = habilito
    subGrupoHasta.enabled = habilito
End Sub
Private Sub opttodoscodigo_Click()
    uProductoDesde.enabled = False
    uProductoHasta.enabled = False
    uProductoDesde.clear
    uProductoHasta.clear
End Sub
Private Sub opttodosgrupo_Click()
    grupoDesde.codigo = ""
    grupoHasta.codigo = ""
    grupoDesde.enabled = False
    grupoHasta.enabled = False
End Sub
Private Sub opttodossubgrupo_Click()
    subGrupoDesde.codigo = ""
    subGrupoHasta.codigo = ""
    subGrupoDesde.enabled = False
    subGrupoHasta.enabled = False
End Sub



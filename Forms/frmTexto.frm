VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmTexto 
   Caption         =   "Textos"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton nuevo 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   1080
      TabIndex        =   13
      Top             =   3240
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos"
      Height          =   2895
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6615
      Begin VB.CheckBox Check3 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2160
         Width           =   375
      End
      Begin VB.CheckBox Check2 
         Caption         =   "K"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1680
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1200
         Width           =   375
      End
      Begin RichTextLib.RichTextBox txtTitu 
         Height          =   375
         Left            =   1200
         TabIndex        =   9
         Top             =   600
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         _Version        =   393217
         Enabled         =   -1  'True
         HideSelection   =   0   'False
         TextRTF         =   $"frmTexto.frx":0000
      End
      Begin RichTextLib.RichTextBox txtDesc 
         Height          =   1335
         Left            =   1200
         TabIndex        =   10
         Top             =   1200
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   2355
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmTexto.frx":0082
      End
      Begin VB.Label Label4 
         Caption         =   "Solo actua para la busqueda."
         Height          =   375
         Left            =   5280
         TabIndex        =   17
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblid 
         Height          =   255
         Left            =   4920
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "ID:"
         Height          =   255
         Left            =   4560
         TabIndex        =   11
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Descripcion :"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Titulo :"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.CommandButton salir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   5880
      TabIndex        =   5
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton aceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton modificar 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton eliminar 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton buscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   855
   End
End
Attribute VB_Name = "frmTexto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private midDoc As Long
Dim estado As Integer

Private Sub limpio()
    txtTitu.Text = ""
    txtDesc.Text = ""
    Check1.Value = 0
    Check2.Value = 0
    Check3.Value = 0
    lblid.caption = nSinNull(obtenerDeSQL("select max(id) as mas from texto where activo=1")) + 1
End Sub

Private Function habilito(a As Boolean)
    Frame1.enabled = a
        
    buscar.enabled = Not a
    eliminar.enabled = a
    modificar.enabled = a
    aceptar.enabled = a
    cancelar.enabled = Not a
    nuevo.enabled = Not a
End Function

Private Sub aceptar_Click()
If ON_ERROR_HABILITADO Then On Error GoTo Errror
'    If txtTitu.Text = "" Then
'        MsgBox "Debe ingresar el titulo.", , "ATENCION"
'        Exit Sub
'    End If
    If txtDesc.Text = "" Then
        MsgBox "Debe ingresar una descripcion", , "ATENCION"
        Exit Sub
    End If
    
    DE_BeginTrans
    
    If estado = 0 Then 'alta de item
        DataEnvironment1.Sistema.Execute "insert into texto (titulo,descripcion,activo,grupo) values('" & Trim(txtTitu.TextRTF) & "','" & Trim(txtDesc.TextRTF) & "',1,2)"
        MsgBox "El alta se realizo con exito.", , "ATENCION"
    Else 'modifico item
        DataEnvironment1.Sistema.Execute "update texto set titulo='" & Trim(txtTitu.TextRTF) & "',descripcion= '" & Trim(txtDesc.TextRTF) & "' where id=" & lblid.caption
        MsgBox "La modificacion se realizo con exito.", , "ATENCION"
    End If
    
    DE_CommitTrans
        
    limpio
    habilito False
    Exit Sub
Errror:
    DE_RollbackTrans
    MsgBox "Error al grabar texto", vbCritical
End Sub

Private Sub buscar_Click()
    Dim resu
    
    limpio
    
    resu = frmBuscar.MostrarSql("select id,dbo.rtf2txt(titulo) [ Titulo                ],dbo.rtf2txt(descripcion) [ Descripcion                                          ] from texto where activo=1 and grupo<>1")
    If resu > "" Then
        lblid.caption = resu
        txtTitu.TextRTF = obtenerDeSQL("select titulo from texto where id=" & resu)
        txtDesc.TextRTF = obtenerDeSQL("select descripcion from texto where id=" & resu)
        habilito False
        buscar.enabled = True
        aceptar.enabled = False
        nuevo.enabled = True
        modificar.enabled = True
    End If
    
End Sub

Private Sub cancelar_Click()
    limpio
    habilito False
End Sub

Private Sub Check1_Click()
    fuente
End Sub

Private Sub fuente()
    If Check1.Value = 1 Then
        txtTitu.SelBold = True
        txtDesc.SelBold = True
    Else
        txtTitu.SelBold = False
        txtDesc.SelBold = False
    End If
    If Check2.Value = 1 Then
        txtTitu.SelItalic = True
        txtDesc.SelItalic = True
    Else
        txtTitu.SelItalic = False
        txtDesc.SelItalic = False
    End If
    If Check3.Value = 1 Then
        txtTitu.SelUnderline = True
        txtDesc.SelUnderline = True
    Else
        txtTitu.SelUnderline = False
        txtDesc.SelUnderline = False
    End If
End Sub

Private Sub Check2_Click()
    fuente
End Sub

Private Sub Check3_Click()
    fuente
End Sub

Private Sub eliminar_Click()
If ON_ERROR_HABILITADO Then On Error GoTo Errror
        
    DE_BeginTrans
    
    DataEnvironment1.Sistema.Execute "update texto set activo=0 where id=" & lblid.caption
    
    DE_CommitTrans
    
    MsgBox "La eliminacion se realizo con exito.", , "ATENCION"
    limpio
    habilito False
Errror:
    midDoc = 0
    DE_RollbackTrans
    'uCheques.resetNroIntPropios
    MsgBox "Error al grabar factura", vbCritical
End Sub

Private Sub Form_Load()
    estado = 0
    limpio
    habilito False
End Sub

Private Sub modificar_Click()
    estado = 1
    habilito True
    nuevo.enabled = False
    eliminar.enabled = False
    cancelar.enabled = True
End Sub

Private Sub nuevo_Click()
    limpio
    habilito True
    cancelar.enabled = True
    eliminar.enabled = False
    modificar.enabled = False
    estado = 0
End Sub

Private Sub salir_Click()
    Unload Me
End Sub

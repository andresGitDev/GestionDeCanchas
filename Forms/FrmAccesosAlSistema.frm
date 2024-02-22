VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAccesosAlSistema 
   Caption         =   "Accesos al Sistema"
   ClientHeight    =   10200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10485
   Icon            =   "FrmAccesosAlSistema.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10200
   ScaleWidth      =   10485
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPunto 
      Caption         =   "Puntos de venta por defecto"
      Height          =   615
      Left            =   5760
      TabIndex        =   11
      Top             =   9480
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<<"
      Height          =   495
      Left            =   4920
      TabIndex        =   10
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">>"
      Height          =   495
      Left            =   4920
      TabIndex        =   9
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grillaus 
      Height          =   2520
      Left            =   75
      TabIndex        =   2
      Top             =   7395
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4445
      _Version        =   393216
      ForeColor       =   8388608
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColorFixed  =   16777215
      BackColorSel    =   16777215
      BackColorBkg    =   -2147483639
      GridColor       =   16777215
      GridColorFixed  =   16777215
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir"
      Height          =   465
      Left            =   9045
      TabIndex        =   1
      Top             =   9600
      Width           =   1335
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   4950
      Top             =   6075
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAccesosAlSistema.frx":08CA
            Key             =   "Grupo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAccesosAlSistema.frx":0D1C
            Key             =   "candado"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvPermisos 
      Height          =   6540
      Left            =   60
      TabIndex        =   0
      Top             =   345
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   11536
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList"
      Appearance      =   1
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grillapermisos 
      Height          =   6630
      Left            =   5790
      TabIndex        =   3
      Top             =   345
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   11695
      _Version        =   393216
      ForeColor       =   128
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColorFixed  =   16777215
      BackColorSel    =   16777215
      BackColorBkg    =   -2147483639
      GridColor       =   16777215
      GridColorFixed  =   16777215
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid especial 
      Height          =   1830
      Left            =   5760
      TabIndex        =   7
      Top             =   7440
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   3228
      _Version        =   393216
      ForeColor       =   128
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColorFixed  =   16777215
      BackColorSel    =   16777215
      BackColorBkg    =   -2147483639
      GridColor       =   16777215
      GridColorFixed  =   16777215
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Permisos Especiales"
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
      Left            =   7200
      TabIndex        =   8
      Top             =   7200
      Width           =   1995
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Permisos"
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
      Left            =   7110
      TabIndex        =   6
      Top             =   90
      Width           =   1995
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Usuarios"
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
      Left            =   1140
      TabIndex        =   5
      Top             =   7125
      Width           =   1995
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00745134&
      Caption         =   "Grupos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1215
      TabIndex        =   4
      Top             =   3330
      Width           =   1995
   End
End
Attribute VB_Name = "FrmAccesosAlSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Nodo As Node
Dim Imagen As String
Dim k As String
Private Sub CargarTV()
Dim rsGrupo As New ADODB.Recordset
Dim rspermisos As New ADODB.Recordset
Dim Consulta As String
    
    Consulta = "Select * From Tipousuarios Where ACTIVO = 1 and codigo>1"
    rsGrupo.Open Consulta, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not rsGrupo.EOF Then
        rsGrupo.MoveFirst
        
        Do While Not rsGrupo.EOF
            Imagen = "Grupo"
            Set Nodo = tvPermisos.Nodes.Add(, , "K" & Trim(str(rsGrupo!codigo)), rsGrupo!DESCRIPCION, Imagen)
            rspermisos.Open "select permisosxusuario.* ,permisos.descripcion as descrip,permisos.codigo as cod from permisosxusuario inner join permisos on permisosxusuario.permiso=permisos.codigo where permisosxusuario.grupo=" & rsGrupo!codigo, DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            If Not rspermisos.EOF Then
                Imagen = "candado"
                Do While Not rspermisos.EOF
                    Set Nodo = tvPermisos.Nodes.Add("K" & Trim(str(rsGrupo!codigo)), 4, "K" & Trim(str(rsGrupo!codigo)) & "CH" & Trim(str(rspermisos!COD)), rspermisos!descrip, Imagen)
                    rspermisos.MoveNext
                Loop
            End If
            rspermisos.Close
            Set rspermisos = Nothing
            rsGrupo.MoveNext
        Loop
    End If
    
    rsGrupo.Close
    Set rsGrupo = Nothing
End Sub

Private Sub cmdPunto_Click()
    FrmAccesoPunto.Show
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

'Private Sub Command2_Click()
'Dim row, i As Long
'On Error GoTo Err
'    grillapermisos.SetFocus
'    row = grillapermisos.rows - 1
'    For i = 1 To grillaProveedor.rows - 1
'        If Mid(k, 1, 1) = "K" Then
'            grillapermisos.Col = 0
'            If grillapermisos.Text <> "" Then
'                Imagen = "candado"
'                Set Nodo = tvPermisos.Nodes.Add(k, 4, k & "CH" & Trim(grillapermisos.Text), Trim(grillapermisos.TextMatrix(grillapermisos.row, 1)), Imagen)
'                DataEnvironment1.Sistema.Execute "insert into permisosxusuario(grupo,permiso)values(" & Val(Mid(k, 2, Len(k))) & "," & Val(grillapermisos.TextMatrix(grillapermisos.row, 0)) & ")"
'            End If
'        End If
'fin:
'    Exit Sub
'Err:
'    If Err.Number = 35602 Then
'        MsgBox "El Permiso ya existe para este Grupo"
'        Resume fin
'    Else
'        MsgBox "Error en la carga del Permiso"
'        Resume fin
'    End If
'End Sub

Private Sub especial_DblClick()
On Error GoTo Err
    If Mid(k, 1, 1) = "K" Then
        especial.Col = 0
        If especial.Text <> "" Then
            Imagen = "candado"
            Set Nodo = tvPermisos.Nodes.Add(k, 4, k & "CH" & Trim(especial.Text), Trim(especial.TextMatrix(especial.Row, 1)), Imagen)
            DataEnvironment1.Sistema.Execute "insert into permisosxusuario(grupo,permiso)values(" & val(Mid(k, 2, Len(k))) & "," & val(especial.TextMatrix(especial.Row, 0)) & ")"
        End If
    End If
fin:
    Exit Sub
Err:
    If Err.Number = 35602 Then
        MsgBox "El Permiso ya existe para este Grupo"
        Resume fin
    Else
        MsgBox "Error en la carga del Permiso"
        Resume fin
    End If
End Sub

Private Sub Form_Load()
    CargarTV
    CargarPermisos
    CargarPermisosEspeciales
End Sub
Sub CargarPermisosEspeciales()
Dim rs As New ADODB.Recordset
    rs.Open "Select * from permisosespeciales where activo=1   order by codigo", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not rs.EOF Then
        especial.cols = 2
        especial.ColWidth(0) = 0
        especial.ColWidth(1) = 5000
        Do While Not rs.EOF
            especial.AddItem rs!codigo & Chr(9) & rs!DESCRIPCION
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Sub
Sub CargarPermisos()
Dim rsper As New ADODB.Recordset
    rsper.Open "Select * from permisos where activo=1 order by codigo", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not rsper.EOF Then
        grillapermisos.cols = 2
        grillapermisos.ColWidth(0) = 0
        grillapermisos.ColWidth(1) = 5000
        Do While Not rsper.EOF
            grillapermisos.AddItem rsper!codigo & Chr(9) & rsper!DESCRIPCION
            rsper.MoveNext
        Loop
    End If
    rsper.Close
    Set rsper = Nothing
End Sub

Private Sub grillapermisos_DblClick()
On Error GoTo Err
    If Mid(k, 1, 1) = "K" Then
        grillapermisos.Col = 0
        If grillapermisos.Text <> "" Then
            Imagen = "candado"
            Set Nodo = tvPermisos.Nodes.Add(k, 4, k & "CH" & Trim(grillapermisos.Text), Trim(grillapermisos.TextMatrix(grillapermisos.Row, 1)), Imagen)
            DataEnvironment1.Sistema.Execute "insert into permisosxusuario(grupo,permiso)values(" & val(Mid(k, 2, Len(k))) & "," & val(grillapermisos.TextMatrix(grillapermisos.Row, 0)) & ")"
        End If
    End If
fin:
    Exit Sub
Err:
    If Err.Number = 35602 Then
        MsgBox "El Permiso ya existe para este Grupo"
        Resume fin
    Else
        MsgBox "Error en la carga del Permiso"
        Resume fin
    End If
End Sub

Private Sub tvPermisos_Click()
    k = tvPermisos.SelectedItem.Key
End Sub
Private Sub tvPermisos_DblClick()
On Error GoTo Err
Dim pos As Integer
Dim mensaje As String

    If Len(tvPermisos.SelectedItem.Key) > 3 Then
        
        pos = InStr(1, k, "C", vbTextCompare)
        mensaje = MsgBox("Esta seguro de quitar este Permiso", vbYesNo, "Atencion")
        If mensaje = 6 Then
            DataEnvironment1.Sistema.Execute "delete from permisosxusuario where grupo=" & val(Mid(k, 2, pos - 1)) & " and permiso=" & val(Mid(k, pos + 2, Len(k)))
            tvPermisos.Nodes.Remove (tvPermisos.SelectedItem.Key)
            k = Mid(k, 1, pos - 1)
        End If
    End If
fin:
    Exit Sub
Err:
    MsgBox "Error al eliminar"
    Resume fin

End Sub

Private Sub tvPermisos_NodeClick(ByVal Node As MSComctlLib.Node)
Dim rsUs As New ADODB.Recordset
    
    grillaus.rows = 0
    rsUs.Open "select usuarios.* from usuarios inner join tipousuarios on usuarios.tipousuario=tipousuarios.codigo where tipousuarios.descripcion='" & Node & "' and usuarios.activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
    If Not rsUs.EOF Then
        grillaus.cols = 1
        grillaus.ColWidth(0) = 3000
        Do While Not rsUs.EOF
             
            grillaus.AddItem rsUs!DESCRIPCION
            rsUs.MoveNext
        Loop
    End If
    rsUs.Close
    Set rsUs = Nothing
End Sub


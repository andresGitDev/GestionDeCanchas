VERSION 5.00
Begin VB.Form frmPorcenList 
   Caption         =   "Porcentajes de listas"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6990
   Icon            =   "frmPorcenList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   6990
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   1320
      TabIndex        =   13
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3240
      TabIndex        =   12
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5880
      TabIndex        =   10
      Top             =   2640
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2760
         TabIndex        =   8
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2760
         TabIndex        =   7
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2760
         TabIndex        =   6
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2760
         TabIndex        =   5
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6240
         TabIndex        =   9
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Porcentaje descuento de lista 4 :"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Porcentaje descuento de lista 3 :"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Porcentaje descuento de lista 2 :"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Porcentaje descuento de lista 1 :"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmPorcenList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function limpio()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
End Function

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    limpio
    habilito
End Sub

Private Sub Command3_Click()
    Dim rs As New ADODB.Recordset
        
    If IsNumeric(Text1.Text) Then
        If Text1.Text < 0 Then
            MsgBox "Los porcentajes no pueden ser negativos.", , "ATENCION"
            Text1.Text = 0
            Text1.SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Los porcentajes deben ser numericos.", , "ATENCION"
        Text1.Text = 0
        Text1.SetFocus
        Exit Sub
    End If
    If IsNumeric(Text2.Text) Then
        If Text2.Text < 0 Then
            MsgBox "Los porcentajes no pueden ser negativos.", , "ATENCION"
            Text2.Text = 0
            Text2.SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Los porcentajes deben ser numericos.", , "ATENCION"
        Text2.Text = 0
        Text2.SetFocus
        Exit Sub
    End If
    If IsNumeric(Text3.Text) Then
        If Text3.Text < 0 Then
            MsgBox "Los porcentajes no pueden ser negativos.", , "ATENCION"
            Text3.Text = 0
            Text3.SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Los porcentajes deben ser numericos.", , "ATENCION"
        Text3.Text = 0
        Text3.SetFocus
        Exit Sub
    End If
    If IsNumeric(Text4.Text) Then
        If Text4.Text < 0 Then
            MsgBox "Los porcentajes no pueden ser negativos.", , "ATENCION"
            Text4.Text = 0
            Text4.SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Los porcentajes deben ser numericos.", , "ATENCION"
        Text4.Text = 0
        Text4.SetFocus
        Exit Sub
    End If
    
    rs.Open "Select * from porcentajelistas where activo=1", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    If (rs.EOF = True And rs.BOF = True) Or IsNull(rs!ID) Or IsEmpty(rs!ID) Then
    Else
        DataEnvironment1.Sistema.Execute "update porcentajelistas set activo=0,fecha_baja=" & ssFecha(Date) & ",usuario_baja=" & UsuarioActual & " where id=" & Label5.caption
    End If
    
    DataEnvironment1.Sistema.Execute "insert into porcentajelistas (lista1,lista2,lista3,lista4,fecha_alta,usuario_alta,activo) values (" & x2s(Text1.Text) & "," & x2s(Text2.Text) & "," & x2s(Text3.Text) & "," & x2s(Text4.Text) & "," & ssFecha(Date) & "," & UsuarioActual & ",1)"
    
    MsgBox "La operacion se realizo con exito.", , "ATENCION"
    limpio
    habilito
End Sub

Private Sub Command4_Click()
    Frame1.enabled = True
    botones True, True, True, False
End Sub

Private Sub Form_Load()
    limpio
    Frame1.enabled = False
    botones True, False, False, False
    habilito
End Sub

Private Function habilito()
    Dim rs As New ADODB.Recordset
    
    rs.Open "Select * from porcentajelistas where activo=1", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
    Frame1.enabled = False
    If (rs.EOF = True And rs.BOF = True) Or IsNull(rs!activo) Or IsEmpty(rs!activo) Then
        Label5.caption = 1
        Frame1.enabled = True
        botones False, True, True, False
    Else
        Text1.Text = nSinNull(rs!lista1)
        Text2.Text = nSinNull(rs!lista2)
        Text3.Text = nSinNull(rs!lista3)
        Text4.Text = nSinNull(rs!lista4)
        Label5.caption = rs!ID
        botones True, False, False, True
    End If
    Set rs = Nothing
End Function

Private Function botones(SALIR As Boolean, Cancelar As Boolean, Aceptar As Boolean, Modificar As Boolean)
    Command1.enabled = SALIR
    Command2.enabled = Cancelar
    Command3.enabled = Aceptar
    Command4.enabled = Modificar
End Function


VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmImprime 
   Caption         =   "Form1"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8385
   Icon            =   "frmImprime.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   8385
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1770
      Left            =   165
      TabIndex        =   30
      Top             =   4425
      Visible         =   0   'False
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   3122
      _Version        =   393216
      Enabled         =   0   'False
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   2730
      Top             =   4335
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   315
      Left            =   6690
      TabIndex        =   26
      Top             =   4035
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   5160
      TabIndex        =   25
      Top             =   4035
      Width           =   1305
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3630
      TabIndex        =   24
      Top             =   4035
      Width           =   1305
   End
   Begin VB.Frame Frame2 
      Caption         =   "Alineacion"
      Height          =   2535
      Left            =   225
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   7935
      Begin VB.CommandButton Command1 
         Caption         =   "Borrar"
         Height          =   300
         Left            =   6510
         TabIndex        =   29
         Top             =   1065
         Width           =   615
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   7215
         TabIndex        =   28
         Top             =   1065
         Width           =   570
      End
      Begin VB.CommandButton cmdSelec 
         Caption         =   "Seleccionar"
         Height          =   300
         Left            =   5535
         TabIndex        =   27
         Top             =   1065
         Width           =   990
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         Left            =   5535
         TabIndex        =   22
         Top             =   1500
         Width           =   1800
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   5535
         TabIndex        =   19
         Top             =   660
         Width           =   1800
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   5535
         TabIndex        =   17
         Top             =   255
         Width           =   1800
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2160
         TabIndex        =   16
         Top             =   1410
         Width           =   1395
      End
      Begin VB.OptionButton Option4 
         Caption         =   "No visible"
         Height          =   195
         Left            =   3375
         TabIndex        =   14
         Top             =   2085
         Width           =   1050
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Visible"
         Height          =   210
         Left            =   2280
         TabIndex        =   13
         Top             =   2070
         Width           =   945
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2160
         TabIndex        =   12
         Top             =   1005
         Width           =   1395
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2175
         TabIndex        =   10
         Top             =   615
         Width           =   1395
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2190
         TabIndex        =   8
         Top             =   255
         Width           =   1395
      End
      Begin VB.Label Label8 
         Caption         =   "Estilo de fondo:"
         Height          =   195
         Left            =   3840
         TabIndex        =   23
         Top             =   1560
         Width           =   1275
      End
      Begin VB.Label Label7 
         Caption         =   "Color para resaltar:"
         Height          =   225
         Left            =   3840
         TabIndex        =   21
         Top             =   1110
         Width           =   1395
      End
      Begin VB.Label Label6 
         Caption         =   "Tamaño de letra:"
         Height          =   210
         Left            =   3810
         TabIndex        =   20
         Top             =   675
         Width           =   1275
      End
      Begin VB.Label Label5 
         Caption         =   "Alineacion horizontal:"
         Height          =   195
         Left            =   3795
         TabIndex        =   18
         Top             =   285
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Ingrese ancho max.:"
         Height          =   210
         Left            =   150
         TabIndex        =   15
         Top             =   1455
         Width           =   1470
      End
      Begin VB.Label Label3 
         Caption         =   "Ingrese largo max.:"
         Height          =   210
         Left            =   150
         TabIndex        =   11
         Top             =   1020
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Ingrese posicion vertical:"
         Height          =   210
         Left            =   150
         TabIndex        =   9
         Top             =   630
         Width           =   1830
      End
      Begin VB.Label Label1 
         Caption         =   "Ingrese posicion Horizontal:"
         Height          =   210
         Left            =   150
         TabIndex        =   7
         Top             =   285
         Width           =   1995
      End
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   3360
      TabIndex        =   5
      Top             =   705
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmImprime.frx":08CA
      Left            =   3360
      List            =   "frmImprime.frx":08CC
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   255
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de impresion"
      Height          =   945
      Left            =   255
      TabIndex        =   1
      Top             =   120
      Width           =   2595
      Begin VB.OptionButton Option2 
         Caption         =   "REMITO"
         Height          =   195
         Left            =   1470
         TabIndex        =   3
         Top             =   435
         Width           =   960
      End
      Begin VB.OptionButton Option1 
         Caption         =   "FACTURA"
         Height          =   270
         Left            =   210
         TabIndex        =   2
         Top             =   390
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Imprimir pagina de prueba"
      Enabled         =   0   'False
      Height          =   360
      Left            =   195
      TabIndex        =   0
      Top             =   4020
      Width           =   2085
   End
End
Attribute VB_Name = "frmImprime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset

Private Sub cmdaceptar_Click()
    Dim variable As Integer
    Dim color As Long
    Dim tipo As String
    Dim Nombre As String
    Dim letra As Integer
    Dim rs2 As New ADODB.Recordset
    If ON_ERROR_HABILITADO Then On Error GoTo UfaGraba

    If Option1.Value = True And Not Combo1.Text = "" Then
        If Not Text1.Text = "" And Not Text2.Text = "" And (Option3.Value = True Or Option4.Value = True) Then
            If Option3.Value = True Then
                variable = "1"
            Else
                variable = "0"
            End If
            color = Text5.BackColor
            tipo = "FACTURA"
            Nombre = CorregirCampo(Trim(Combo1.Text))
            If Text3.Text = "" Then Text3.Text = 0
            If Text4.Text = "" Then Text4.Text = 0
            If Combo4.Text = "" Then
                letra = 0
            Else
                letra = Combo4.Text
            End If
            rs.Open "Select codigo from posicionar where nombre='" & Nombre & "' and imprecionde='" & tipo & "'", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
            DataEnvironment1.Sistema.Execute "UPDATE Posicionar SET [Nombre]='" & Nombre & "',[ImprecionDe]='" & tipo & "',[PosX]= " & Text1.Text & ",[PosY]= " & Text2.Text & ",[Visible]= " & variable & ",[Largo]= " & Text3.Text & "," _
                    & "[Ancho]= " & Text4.Text & ",[AlineaHorizontal]= '" & Combo3.Text & "',[SizeLetra]= " & letra & ",[Backcolor]= '" & color & "',[BackStyle]= '" & Combo6.Text & "' Where ( [Nombre]= '" & Nombre & "' AND" _
                    & "[ImprecionDe]= '" & tipo & "' and [codigo]=" & rs!codigo & ")"
            'DE_CommitTrans
            If Text3.Text = 0 Then Text3.Text = ""
            If Text4.Text = 0 Then Text4.Text = ""
        End If
    ElseIf Option2.Value = True And Not Combo2.Text = "" Then
        If Not Text1.Text = "" And Not Text2.Text = "" And (Option3.Value = True Or Option4.Value = True) Then
            If Option3.Value = True Then
                variable = 1
            Else
                variable = 0
            End If
            color = Text5.BackColor
            tipo = "REMITO"
            Nombre = CorregirCampo(Trim(Combo2.Text))
            If Text3.Text = "" Then Text3.Text = 0
            If Text4.Text = "" Then Text4.Text = 0
            If Combo4.Text = "" Then
                letra = 0
            Else
                letra = Combo4.Text
            End If
            rs.Open "Select codigo from posicionar where nombre='" & Nombre & "' and imprecionde='" & tipo & "'", DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
            'DataEnvironment1.dbo_ACTUALIZAPOSICION Nombre, Trim(tipo), rs!codigo, Nombre, Trim(tipo), Trim(Text1.Text), Trim(Text2.Text), _
            '        Trim(variable), Trim(Text3.Text), Trim(Text4.Text), Trim(Combo3.Text), Trim(letra), Trim(color), Trim(Combo6.Text)
            DataEnvironment1.Sistema.Execute "UPDATE Posicionar SET [Nombre]='" & Nombre & "',[ImprecionDe]='" & tipo & "',[PosX]= " & Text1.Text & ",[PosY]= " & Text2.Text & ",[Visible]= " & variable & ",[Largo]= " & Text3.Text & "," _
                    & "[Ancho]= " & Text4.Text & ",[AlineaHorizontal]= '" & Combo3.Text & "',[SizeLetra]= " & letra & ",[Backcolor]= '" & color & "',[BackStyle]= '" & Combo6.Text & "' Where ( [Nombre]= '" & Nombre & "' AND" _
                    & "[ImprecionDe]= '" & tipo & "' and [codigo]=" & rs!codigo & ")"
            If Text3.Text = 0 Then Text3.Text = ""
            If Text4.Text = 0 Then Text4.Text = ""
        End If
    End If
    'Set rs = Nothing
fin:
    Set rs = Nothing
    Exit Sub
UfaGraba:
    'DE_RollbackTrans
    ufa "err al grabar ", "aceptar"
    Resume fin

End Sub

Private Sub cmdcancelar_Click()
    Limpiar
    Combo1.ListIndex = 0
    Combo2.ListIndex = 0
End Sub

Private Sub cmdImprimir_Click()
    If Option1.Value = True Then 'para factura
        ImprimirComprobantePrueba
    ElseIf Option2.Value = True Then 'para remitos
        ImprimirRemitoPrueba
    End If
End Sub

Private Sub cmdsalir_Click()
    Limpiar
    Frame2.Visible = False
    Unload Me
End Sub

Private Sub cmdSelec_Click()
    dlgColor.ShowColor
    Text5.BackColor = dlgColor.color
    MsgBox "Recuerde que para que se aplique el color debe tener el estilo normal."
End Sub

Private Sub Combo1_Click()
    Dim strSql As String
    Dim Nombre As String
    Limpiar
    If Combo1.Text = "Domicilio" Or Combo1.Text = "CUIT" Or Combo1.Text = "IVA" Or Combo1.Text = "Debe" Or Combo1.Text = "Producto" Or Combo1.Text = "Dia" Or Combo1.Text = "Mes" Or Combo1.Text = "Localidad" Or Combo1.Text = "Subtotal" Or Combo1.Text = "Impuesto" Or Combo1.Text = "Total" Or Combo1.Text = "Provincia" Or Combo1.Text = "Pais" Then
    'Or Combo1.Text = "Cantidad" Or Combo1.Text = "Articulo" Or Combo1.Text = "Descripcion"
        strSql = "select * from posicionar where imprecionde='FACTURA' and nombre='" & Combo1.Text & "'"
        rs.Open strSql, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
        If Not (rs.EOF = True And rs.BOF = True) Then
            Frame2.Visible = True
            If Not IsNull(rs.Fields(3).Value) Then Text1.Text = rs.Fields(3).Value
            If Not IsNull(rs.Fields(4).Value) Then Text2.Text = rs.Fields(4).Value
            If rs.Fields(5).Value = True Then
                Option3.Value = True
            Else
                Option4.Value = True
            End If
            If Not IsNull(rs.Fields(6).Value) And rs.Fields(6).Value > 0 Then Text3.Text = rs.Fields(6).Value
            If Not IsNull(rs.Fields(7).Value) And rs.Fields(7).Value > 0 Then Text4.Text = rs.Fields(7).Value
            If Not IsNull(rs.Fields(8).Value) Or Not rs.Fields(8).Value = "" Then
                Combo3.SelText = rs.Fields(8).Value
            Else
                Combo3.ListIndex = 0
            End If
            If Not IsNull(rs.Fields(9).Value) And rs.Fields(9).Value > 0 Then Combo4.SelText = rs.Fields(9).Value
            If Not IsNull(rs.Fields(10).Value) Then
                If Not rs.Fields(10).Value = "" Then Text5.BackColor = rs.Fields(10).Value
            End If
            If Not IsNull(rs.Fields(11).Value) Then Combo6.SelText = rs.Fields(11).Value
        Else
            MsgBox "No seha encontrado datos referidos a este campo."
        End If
    Else 'aca cargo los nombres no coincidentes
                
        If Combo1.Text = "" Then
            Exit Sub
        Else
            Nombre = CorregirCampo(Trim(Combo1.Text))
        End If
        strSql = "select * from posicionar where imprecionde='FACTURA' and nombre='" & Nombre & "'"
        
        rs.Open strSql, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
        If Not (rs.EOF = True And rs.BOF = True) Then
            Frame2.Visible = True
            If Not IsNull(rs.Fields(3).Value) Then Text1.Text = rs.Fields(3).Value
            If Not IsNull(rs.Fields(4).Value) Then Text2.Text = rs.Fields(4).Value
            If rs.Fields(5).Value = True Then
                Option3.Value = True
            Else
                Option4.Value = True
            End If
            If Not IsNull(rs.Fields(6).Value) And rs.Fields(6).Value > 0 Then Text3.Text = rs.Fields(6).Value
            If Not IsNull(rs.Fields(7).Value) And rs.Fields(7).Value > 0 Then Text4.Text = rs.Fields(7).Value
            If Not IsNull(rs.Fields(8).Value) Or Not rs.Fields(8).Value = "" Then
                Combo3.SelText = rs.Fields(8).Value
            Else
                Combo3.ListIndex = 0
            End If
            If Not IsNull(rs.Fields(9).Value) Then Combo4.SelText = rs.Fields(9).Value
            If Not IsNull(rs.Fields(10).Value) Then
                If Not rs.Fields(10).Value = "" Then Text5.BackColor = rs.Fields(10).Value
            End If
            If Not IsNull(rs.Fields(11).Value) Then Combo6.SelText = rs.Fields(11).Value
        Else
            MsgBox "No seha encontrado datos referidos a este campo."
        End If
    End If
    'rs.Close
    CmdImprimir.enabled = True
    cmdaceptar.enabled = True
    Set rs = Nothing
End Sub
Private Sub Combo2_Click()
    Dim strSql As String
    Dim Nombre As String
    Limpiar
    If Combo2.Text = "Domicilio" Or Combo2.Text = "CUIT" Or Combo2.Text = "Dia" Or Combo2.Text = "Mes" Then
        strSql = "select * from posicionar where imprecionde='REMITO' and nombre='" & Combo2.Text & "'"
        rs.Open strSql, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
        If Not (rs.EOF = True And rs.BOF = True) Then
            Frame2.Visible = True
            If Not IsNull(rs.Fields(3).Value) Then Text1.Text = rs.Fields(3).Value
            If Not IsNull(rs.Fields(4).Value) Then Text2.Text = rs.Fields(4).Value
            If rs.Fields(5).Value = True Then
                Option3.Value = True
            Else
                Option4.Value = True
            End If
            If Not IsNull(rs.Fields(6).Value) And rs.Fields(6).Value > 0 Then Text3.Text = rs.Fields(6).Value
            If Not IsNull(rs.Fields(7).Value) And rs.Fields(7).Value > 0 Then Text4.Text = rs.Fields(7).Value
            If Not IsNull(rs.Fields(8).Value) Or Not rs.Fields(8).Value = "" Then
                Combo3.SelText = rs.Fields(8).Value
            Else
                Combo3.ListIndex = 0
            End If
            If Not IsNull(rs.Fields(9).Value) And rs.Fields(9).Value > 0 Then Combo4.SelText = rs.Fields(9).Value
            If Not IsNull(rs.Fields(10).Value) Then
                If Not rs.Fields(10).Value = "" Then Text5.BackColor = rs.Fields(10).Value
            End If
            If Not IsNull(rs.Fields(11).Value) Then Combo6.SelText = rs.Fields(11).Value
        Else
            MsgBox "No se ha encontrado datos referidos a este campo."
        End If
    Else 'aca cargo los nombres no coincidentes
        
        If Combo2.Text = "" Then
            Exit Sub
        Else
            Nombre = CorregirCampo(Trim(Combo2.Text))
        End If
        strSql = "select * from posicionar where imprecionde='REMITO' and nombre='" & Nombre & "'"
        
        rs.Open strSql, DataEnvironment1.Sistema, adOpenStatic, adLockReadOnly
        If Not (rs.EOF = True And rs.BOF = True) Then
            Frame2.Visible = True
            If Not IsNull(rs.Fields(3).Value) Then Text1.Text = rs.Fields(3).Value
            If Not IsNull(rs.Fields(4).Value) Then Text2.Text = rs.Fields(4).Value
            If rs.Fields(5).Value = True Then
                Option3.Value = True
            Else
                Option4.Value = True
            End If
            If Not IsNull(rs.Fields(6).Value) And rs.Fields(6).Value > 0 Then Text3.Text = rs.Fields(6).Value
            If Not IsNull(rs.Fields(7).Value) And rs.Fields(7).Value > 0 Then Text4.Text = rs.Fields(7).Value
            If Not IsNull(rs.Fields(8).Value) Or Not rs.Fields(8).Value = "" Then
                Combo3.SelText = rs.Fields(8).Value
            Else
                Combo3.ListIndex = 0
            End If
            If Not IsNull(rs.Fields(9).Value) Then Combo4.SelText = rs.Fields(9).Value
            If Not IsNull(rs.Fields(10).Value) Then
                If Not rs.Fields(10).Value = "" Then Text5.BackColor = rs.Fields(10).Value
            End If
            If Not IsNull(rs.Fields(11).Value) Then Combo6.SelText = rs.Fields(11).Value
        Else
            MsgBox "No se ha encontrado datos referidos a este campo."
        End If
    End If
    'rs.Close
    CmdImprimir.enabled = True
    cmdaceptar.enabled = True
    Set rs = Nothing
End Sub

Private Sub Command1_Click()
    Text5.BackColor = &H80000005
    dlgColor.color = &H80000005
End Sub

Private Sub Form_Load()
    cargoCombo1
    cargoCombo2
    cargoCombo3
    cargoCombo4
    cargocombo6
End Sub

Private Sub Option1_Click()
    validarFACT True, False
    Limpiar
    Combo2.Text = ""
End Sub
Public Function validarFACT(valor1 As String, valor2 As String)
    Combo1.Visible = valor1
    Combo2.Visible = valor2
    Frame2.Visible = False
    CmdImprimir.enabled = False
    cmdaceptar.enabled = False
End Function
Public Sub Limpiar()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.BackColor = &H80000005
    'Combo1.SelText = ""
    'Combo2.Text = ""
    Combo3.Text = ""
    Combo4.Text = ""
    Combo6.Text = ""
    Option3.Value = False
    Option4.Value = False
End Sub

Private Sub Option2_Click()
    validarFACT False, True
    Limpiar
    'Combo1.Text = ""
End Sub
Public Sub cargoCombo1()
    Combo1.AddItem "", 0
    Combo1.AddItem "Señor", 1
    Combo1.AddItem "Domicilio", 2
    Combo1.AddItem "CUIT", 3
    Combo1.AddItem "IVA", 4
    Combo1.AddItem "Condicion de venta", 5
    Combo1.AddItem "Ingresos brutos", 6
    Combo1.AddItem "Nro de cliente", 7
    Combo1.AddItem "Presupuesto", 8
    Combo1.AddItem "Orden de compra", 9
    Combo1.AddItem "Debe", 10
    Combo1.AddItem "Producto", 11
    Combo1.AddItem "Nro de remito", 12
    Combo1.AddItem "Dia", 13
    Combo1.AddItem "Mes", 14
    Combo1.AddItem "Año", 15
    Combo1.AddItem "Nombre del mes", 16
    Combo1.AddItem "Responsable inscripto", 17
    Combo1.AddItem "Responsable no inscripto", 18
    
    Combo1.AddItem "Localidad", 19
    Combo1.AddItem "Factura", 20
    Combo1.AddItem "Provincia", 21
    Combo1.AddItem "Pais", 22
    Combo1.AddItem "Fecha", 23
    Combo1.AddItem "Codigo postal", 24
    Combo1.AddItem "Descuento", 25
    Combo1.AddItem "Porcentaje de descuento", 26
    Combo1.AddItem "Iva inscripto", 27
    Combo1.AddItem "Porcentaje de iva inscripto", 28
    Combo1.AddItem "IIBB", 29
    Combo1.AddItem "Porcentaje de IIBB", 30
    
    Combo1.AddItem "Subtotal", 31
    Combo1.AddItem "Impuesto", 32
    Combo1.AddItem "Segundo subtotal", 33
    Combo1.AddItem "IVA inscripto", 34
    Combo1.AddItem "IVA no inscripto", 35
    Combo1.AddItem "Total", 36
    
    Combo1.AddItem "Cantidad", 37
    Combo1.AddItem "Articulo", 38
    Combo1.AddItem "Descripcion", 39
    Combo1.AddItem "Precio Unitario", 40
    Combo1.AddItem "Precio Total", 41
End Sub
Public Sub cargoCombo3()
    With Combo3
        .AddItem "", 0
        .AddItem "Izquierda", 1
        .AddItem "Derecha", 2
        .AddItem "Centro", 3
    End With
End Sub
Public Sub cargoCombo2() 'este combo es para remitos
    Combo2.AddItem "", 0
    Combo2.AddItem "Señor", 1
    Combo2.AddItem "Domicilio", 2
    Combo2.AddItem "Localidad", 3
    Combo2.AddItem "Nro de Provincia", 4
    Combo2.AddItem "CUIT", 5
    Combo2.AddItem "Transportista", 6
    Combo2.AddItem "Presupuesto", 7
    Combo2.AddItem "Orden de compra", 8
    Combo2.AddItem "Dia", 9
    Combo2.AddItem "Mes", 10
    Combo2.AddItem "Año", 11
    Combo2.AddItem "Nombre del mes", 12
    Combo2.AddItem "Fecha", 13
    Combo2.AddItem "Nro de referencia", 14
    Combo2.AddItem "Referencia", 15
    Combo2.AddItem "Tachar", 16
    Combo2.AddItem "Comprobante", 17
    Combo2.AddItem "Factura", 18
    Combo2.AddItem "Atencion", 19
    Combo2.AddItem "Condicion de IVA", 20
    
    Combo2.AddItem "Cantidad", 21
    Combo2.AddItem "Articulo", 22
    Combo2.AddItem "Descripcion", 23
    Combo2.AddItem "Precio Unitario", 24
    Combo2.AddItem "Precio Total", 25
End Sub
Public Sub cargoCombo4()
    With Combo4
        .AddItem "", 0
        .AddItem "8", 1
        .AddItem "9", 2
        .AddItem "10", 3
        .AddItem "11", 4
        .AddItem "12", 5
        .AddItem "14", 6
        .AddItem "16", 7
        .AddItem "18", 8
        .AddItem "20", 9
        .AddItem "22", 10
        .AddItem "24", 11
        .AddItem "26", 12
        .AddItem "28", 13
        .AddItem "36", 14
        .AddItem "48", 15
    End With
End Sub
Public Sub cargocombo6()
    With Combo6
        .AddItem "", 0
        .AddItem "Transparente", 1
        .AddItem "Normal", 2
    End With
End Sub
Public Function CorregirCampo(dato As String) As String
    If Not dato = "Domicilio" Or Not dato = "CUIT" Or Not dato = "IVA" Or Not dato = "Debe" Or Not dato = "Producto" Or Not dato = "Dia" Or Not dato = "Mes" Or Not dato = "Subtotal" Or Not dato = "Impuesto" Or Not dato = "Total" Then
        'aca cargo los nombres no coincidentes
        If dato = "Señor" Then
            CorregirCampo = "senor"
        ElseIf dato = "Condicion de venta" Then
            CorregirCampo = "condiventa"
        ElseIf dato = "Ingresos brutos" Then
            CorregirCampo = "ingbruto"
        ElseIf dato = "Nro de cliente" Then
            CorregirCampo = "nrocli"
        ElseIf dato = "Presupuesto" Then
            CorregirCampo = "presupu"
        ElseIf dato = "Orden de compra" Then
            CorregirCampo = "ordencomp"
        ElseIf dato = "Nro de remito" Then
            CorregirCampo = "nroremito"
        ElseIf dato = "Año" Then
            CorregirCampo = "ano"
        ElseIf dato = "Nombre del mes" Then
            CorregirCampo = "meses"
        ElseIf dato = "Responsable inscripto" Then
            CorregirCampo = "lblrespinsc"
        ElseIf dato = "Responsable no inscripto" Then
            CorregirCampo = "lblrespnoinsc"
        ElseIf dato = "Segundo subtotal" Then
            CorregirCampo = "subtotal2"
        ElseIf dato = "IVA inscripto" Then
            CorregirCampo = "ivainsc"
        ElseIf dato = "IVA no inscripto" Then
            CorregirCampo = "ivanoinsc"
        ElseIf dato = "Transportista" Then
            CorregirCampo = "transpo"
        ElseIf dato = "Precio Unitario" Then
            CorregirCampo = "lblPrecUnitario"
        ElseIf dato = "Precio Total" Then
            CorregirCampo = "lblPrecTotal"
        ElseIf dato = "Factura" Then
            CorregirCampo = "lblfactura"
        ElseIf dato = "Fecha" Then
            CorregirCampo = "lblfecha"
        ElseIf dato = "Codigo postal" Then
            CorregirCampo = "codpos"
        ElseIf dato = "Descuento" Then
            CorregirCampo = "txtdcto"
        ElseIf dato = "Porcentaje de descuento" Then
            CorregirCampo = "txtdctop"
        ElseIf dato = "Iva inscripto" Then
            CorregirCampo = "txtivains"
        ElseIf dato = "Porcentaje de iva inscripto" Then
            CorregirCampo = "txtivap"
        ElseIf dato = "IIBB" Then
            CorregirCampo = "txtiibb"
        ElseIf dato = "Porcentaje de IIBB" Then  'falta agregar remito
            CorregirCampo = "txtiibbp"
        ElseIf dato = "Cantidad" Then
            CorregirCampo = "lblCantidad"
        ElseIf dato = "Articulo" Then
            CorregirCampo = "lblArticulo"
        ElseIf dato = "Descripcion" Then
            CorregirCampo = "lblDescripcion"
        ElseIf dato = "Precio Unitario" Then
            CorregirCampo = "lblPrecUnitario"
        ElseIf dato = "Precio Total" Then
            CorregirCampo = "lblPrecTotal"
        ElseIf dato = "Atencion" Then
            CorregirCampo = "lblAtencion"
        ElseIf dato = "Condicion de IVA" Then
            CorregirCampo = "lbliva"
        ElseIf dato = "Nro de Provincia" Then
            CorregirCampo = "lblnroprov"
        ElseIf dato = "Localidad" And Option2.Value = True Then
            CorregirCampo = "lbllocalidad"
        ElseIf dato = "Nro de referencia" Then
            CorregirCampo = "lblnroref"
        ElseIf dato = "Referencia" Then
            CorregirCampo = "lblref"
        ElseIf dato = "Tachar" Then
            CorregirCampo = "lblTachar"
        ElseIf dato = "Comprobante" Then
            CorregirCampo = "lblComp"
        Else
            CorregirCampo = dato
        End If
    End If
End Function


'*************************************************************
'*************************************************************
Private Function ImprimirRemitoPrueba() As Boolean
'el codigo es de remitoventa y creo que no es necesario
    If ON_ERROR_HABILITADO Then On Error GoTo ErrImpresora
    
    Dim cod As Long
    Dim Propio As Boolean
    
    Propio = 1
    RptImpresionRemitoVenta.lblcliente = "Hoja de prueba"
    RptImpresionRemitoVenta.lblcuit = "25-12345678-1"
    cod = 856
    RptImpresionRemitoVenta.lblcomp = "Remito"
    
               
    RptImpresionRemitoVenta.lblcliente = "Perez Juan"
    RptImpresionRemitoVenta.lbldomicilio = "Marbella 123"
    RptImpresionRemitoVenta.lblfactura = "0001-" & Format(856, "00000000")
    RptImpresionRemitoVenta.lblfecha = "22/04/2006"
    RptImpresionRemitoVenta.lbllocalidad = "Capital Federal"
    RptImpresionRemitoVenta.Transportista = "Diego Torres"
    RptImpresionRemitoVenta.lblAtencion = "1 Lopez Damian"
    RptImpresionRemitoVenta.OrdenComp = 12
    RptImpresionRemitoVenta.lbliva = "INSCRIPTO"
        
    RptImpresionRemitoVenta.Dia = "22"
    RptImpresionRemitoVenta.Mes = "04"
    RptImpresionRemitoVenta.Ano = "2006"
    RptImpresionRemitoVenta.Meses = "Abril"
        
    'ver que el rs tiene un addnew y hay que usar update
    Set rs2 = New ADODB.Recordset

    With rs2
        ' Establece IdCliente como la clave principal.
        .Fields.Append "Id", adChar, 5, adFldRowID
        .Fields.Append "Cantidad", adInteger, 4, adFldUpdatable
        .Fields.Append "Codigo", adChar, 24, adFldUpdatable
        .Fields.Append "Descrip", adChar, 65, adFldUpdatable
        ' Utilice el tipo de cursor Keyset para permitir la actualización
        ' de los registros.
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With
        
    addrs 1, 3, "202", "PANTALON LIVIANO"
    addrs 2, 7, "225", "BERMUDA"
    addrs 3, 12, "4046", "BLAZER  CHANEL"
    addrs 4, 9, "190", "CAMISA 3/4"
    addrs 5, 5, "4017", "CAMPERA PANA"
        
        
    rs2.MoveFirst
    Set RptImpresionRemitoVenta.DataControl1.Recordset = rs2 'str
    Set MSHFlexGrid1.DataSource = rs2
        
    
    Set rs2 = Nothing
    RptImpresionRemitoVenta.Printer.Copies = 1
    
    If PREVIEW_IMPRESIONES Then
        RptImpresionRemitoVenta.Show
        
        'esto es para acomodar las posiciones del remito
        Posicionar (False) 'false para remito y true para factura
    Else
        RptImpresionRemitoVenta.PrintReport True
    End If
    RptImpresionRemitoVenta.Restart
    
fin:
    Exit Function
ErrImpresora:
    ufa "error de impresión Remito Venta", ""
    Resume fin
End Function
Private Function addrs(ID As Integer, cant As Integer, art As String, desc As String) As Boolean
    With rs2
        .AddNew
        !ID = ID
        !cantidad = cant
        !codigo = art
        !descrip = desc
        .Update
        '.Bookmark = .LastModified
    End With
End Function
Private Function addrs2(ID As Integer, cant As Integer, art As String, desc As String, unit As Double, tot As Double) As Boolean
    With rs2
        .AddNew
        !ID = ID
        !cantidad = cant
        !producto = art
        !DESCRIPCION = desc
        !punit = s2n(unit, 2)
        !ptot = s2n(tot, 2)
        .Update
        '.Bookmark = .LastModified
    End With
End Function


Public Function ImprimirComprobantePrueba() As Boolean
'el codigo es de facturaVenta
    If ON_ERROR_HABILITADO Then On Error GoTo ErrImpresion
    
    Dim cod As Long
    Dim PORCENTAJE As String
    Dim tdoc As String, z As Double, mone As Long
    Dim valor, Neto As Double
    Dim mImpresoraDefecto As String
    Dim Propio As Boolean
    Dim strDetalle As String
    
    tdoc = "FAA"
    z = 1
    mone = 1
    
    RptImpresionFacturaVenta.CodPos = "1678"
    RptImpresionFacturaVenta.lblcliente = "GUERRERO S.A."
    RptImpresionFacturaVenta.lblcuit = "25-12345678-9"
    cod = 958
    RptImpresionFacturaVenta.txtIvaP = 21
    RptImpresionFacturaVenta.OrdenComp = 0
            
    RptImpresionFacturaVenta.NroCli = 644
        
    RptImpresionFacturaVenta.lblNroProv = ""
    Propio = 1
        
    If Left(tdoc, 2) = "FA" Then
        RptImpresionFacturaVenta.lblcomp = "Factura"
        RptImpresionFacturaVenta.Remitos = Format(0, "00000000")
        If tdoc = "FAB" Then
        Else ' tdoc = "FAA"
            Neto = Format$(s2n("3885,96", 2), "Standard")
            PORCENTAJE = Format$(s2n(21 / 100, 2), "Standard")
            RptImpresionFacturaVenta.IvaNoInsc = Format$(s2n(Neto * PORCENTAJE, 2), "Standard")
            RptImpresionFacturaVenta.txtivains = Format$(s2n(Neto * PORCENTAJE, 2), "Standard")
            RptImpresionFacturaVenta.txtneto = Format$(s2n(Neto, 2), "Standard")
            RptImpresionFacturaVenta.txtsub = Format$(s2n(Neto, 2), "Standard")
            RptImpresionFacturaVenta.txtIibbP = Format$(s2n((((40 * 100) / Neto) / 100), 2), "Standard")
            RptImpresionFacturaVenta.txtIIBB = Format$(s2n(40, 2), "Standard")
            valor = RptImpresionFacturaVenta.txtIIBB
            
            RptImpresionFacturaVenta.Subtotal2 = s2n((RptImpresionFacturaVenta.txtsub - valor), 2)
            
            RptImpresionFacturaVenta.txtDctoP = 5
            Dim t_sub, t_neto, t_coef
                    
            t_neto = s2n(Neto, 2)
            t_coef = s2n(5 / 100, 2)
            t_sub = s2n(t_neto / (1 - t_coef), 2)
            
            RptImpresionFacturaVenta.txtsub = Format$(s2n(t_sub, 2), "Standard")
            RptImpresionFacturaVenta.txtDcto = Format$(s2n(t_neto - t_sub, 2), "standard")
            
        End If
        
    End If
            
    RptImpresionFacturaVenta.lbldomicilio = "MARBELLA 123"
    RptImpresionFacturaVenta.Provincia = "BUENOS AIRES"
    RptImpresionFacturaVenta.lblfactura = "0001-" & Format(644, "00000000")
    RptImpresionFacturaVenta.lblfecha = "22/04/2006"

    RptImpresionFacturaVenta.Dia = "22"
    RptImpresionFacturaVenta.Mes = "04"
    RptImpresionFacturaVenta.Ano = "2006"
    RptImpresionFacturaVenta.Meses = "ABRIL"
        
    RptImpresionFacturaVenta.lbliva = "INSCRIPTO" 's2n(neto * (21 / 100), 2)
    
    RptImpresionFacturaVenta.lbllocalidad = "LANUS"
    RptImpresionFacturaVenta.lblref = "Remito"
    RptImpresionFacturaVenta.lblnroref = "0001-" & Format(22, "00000000")
        
    RptImpresionFacturaVenta.lblpago = "CONTADO"
    
    RptImpresionFacturaVenta.lblimp = "Son " & ObtenerDescripcion("Monedas", mone) & ": " & enletras(s2n(4702.01 / z))
    RptImpresionFacturaVenta.txttotalfinal = Format$(s2n(4702.01 / z), "standard")
        
    Set rs2 = New ADODB.Recordset

    With rs2
        ' Establece IdCliente como la clave principal.
        .Fields.Append "Id", adChar, 5, adFldRowID
        .Fields.Append "Cantidad", adInteger, 4, adFldUpdatable
        .Fields.Append "Producto", adChar, 24, adFldUpdatable
        .Fields.Append "Descripcion", adChar, 65, adFldUpdatable
        .Fields.Append "punit", adDouble, 10, adFldUpdatable
        .Fields.Append "ptot", adDouble, 10, adFldUpdatable
        ' Utilice el tipo de cursor Keyset para permitir la actualización
        ' de los registros.
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With
    
    addrs2 1, 3, "202", "PANTALON LIVIANO", "61,57", "184,71"
    addrs2 2, 7, "225", "BERMUDA", "53,31", "373,17"
    addrs2 3, 12, "4046", "BLAZER  CHANEL", "177,27", "2127,24"
    addrs2 4, 9, "190", "CAMISA 3/4", "53,31", "479,79"
    addrs2 5, 5, "4017", "CAMPERA PANA", "144,21", "721,05"
        
    rs2.MoveFirst
    Set RptImpresionFacturaVenta.DataControl1.Recordset = rs2
    Set MSHFlexGrid1.DataSource = rs2
    
    RptImpresionFacturaVenta.Printer.Copies = 1
    RptImpresionFacturaVenta.Show
    
    Posicionar (True) 'acomodo factura, false para remito
    
    Set rs2 = Nothing

fin:
    Set rs = Nothing
    Exit Function
ErrImpresion:
    ufa "Error de impresión en factura venta", ""
    Resume fin
End Function


'en la base el valor 0(cero) significa que tomara el valor por
'defecto que tiene el objeto en tiempo de diseño,por lo tanto el 0
'no se toma en cuenta, al igual que el vacio.

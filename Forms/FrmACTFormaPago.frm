VERSION 5.00
Begin VB.Form FrmACTFormaPago 
   Caption         =   "Actualizacion Unica"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   8910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Productos Cortos"
      Height          =   495
      Left            =   165
      TabIndex        =   17
      Top             =   810
      Width           =   1935
   End
   Begin VB.CommandButton cmdComprasRet 
      Caption         =   "Genera tabla Compras Retencion"
      Height          =   555
      Left            =   180
      TabIndex        =   16
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox txtcomp 
      Height          =   375
      Index           =   5
      Left            =   6240
      TabIndex        =   15
      Text            =   "Text3"
      Top             =   4560
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.CommandButton Command3 
      Caption         =   "OJO: Pasar a cta cte"
      Height          =   495
      Left            =   180
      TabIndex        =   14
      Top             =   5040
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.TextBox txtcomp 
      Height          =   375
      Index           =   4
      Left            =   4800
      TabIndex        =   13
      Text            =   "Text3"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtcomp 
      Height          =   375
      Index           =   3
      Left            =   3600
      TabIndex        =   12
      Text            =   "Text3"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.TextBox txtcomp 
      Height          =   375
      Index           =   2
      Left            =   2340
      TabIndex        =   11
      Text            =   "Text3"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.TextBox txtcomp 
      Height          =   375
      Index           =   1
      Left            =   1260
      TabIndex        =   10
      Text            =   "Text3"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtcomp 
      Height          =   375
      Index           =   0
      Left            =   180
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton Command2 
      Caption         =   "BuscaComprasContado"
      Height          =   495
      Left            =   180
      TabIndex        =   8
      Top             =   3960
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdUnicaVez 
      Caption         =   "Actualizacion"
      Enabled         =   0   'False
      Height          =   495
      Left            =   180
      TabIndex        =   7
      Top             =   1380
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   900
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2820
      Width           =   5475
   End
   Begin VB.CommandButton Command1 
      Caption         =   "buscar documento anulado"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2220
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2820
      Width           =   675
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   5700
      Width           =   1215
   End
   Begin VB.CommandButton CmdEjecutar 
      Caption         =   "OJO; Borrar factura"
      Height          =   495
      Left            =   4140
      TabIndex        =   1
      Top             =   2220
      Width           =   1695
   End
   Begin VB.Line Line1 
      Index           =   1
      Visible         =   0   'False
      X1              =   180
      X2              =   6660
      Y1              =   1980
      Y2              =   1980
   End
   Begin VB.Line Line1 
      Index           =   0
      Visible         =   0   'False
      X1              =   120
      X2              =   6600
      Y1              =   3540
      Y2              =   3540
   End
   Begin VB.Label Label2 
      Caption         =   "Factura venta / NC "
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   2220
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Este Formulario debe ser ejecutado una Sola Vez"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2700
      TabIndex        =   0
      Top             =   120
      Width           =   4395
   End
End
Attribute VB_Name = "FrmACTFormaPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdComprasRet_Click()
    Dim ss As String
    ss = "CREATE TABLE dbo.ComprasRetenciones ( [idComprasRet] [numeric](18, 0) IDENTITY (1, 1) NOT NULL , [iddoc] [numeric](18, 0) NOT NULL ,    [idCuentasParam] [numeric](18, 0) NOT NULL ,    [Numero] [numeric](18, 0) NULL ,    [Fecha] [datetime] NULL ,    [Importe] [float] NOT NULL ,    [NroFactura] [numeric](18, 0) NULL ,    [Cuenta] [char] (7) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL) ON [PRIMARY]"
    DataEnvironment1.Sistema.Execute ss
    ss = "ALTER TABLE dbo.ComprasRetenciones WITH NOCHECK ADD CONSTRAINT [PK_ComprasRetenciones] PRIMARY KEY  CLUSTERED ([idComprasRet] )  ON [PRIMARY]"
    DataEnvironment1.Sistema.Execute ss
    che "ya tá"
End Sub

Private Sub CmdEjecutar_Click()
Dim sql As String
Dim n As Long, tempo

n = s2n(Text1)
If n = 0 Then
    che " numero?"
    Exit Sub
End If

''**************************************************************************
'' PROCESO PARA borrar factura / nc anulada
   sql = "delete from facturaventadetalle where codigofactura = '" & x2s(n) & "' "
   DataEnvironment1.Sistema.Execute sql

   sql = "delete from facturaventa where codigo = '" & x2s(n) & "' "
   DataEnvironment1.Sistema.Execute sql

''**************************************************************************
'' PROCESO PARA AGREGAR AL CAMPO IVA DE FACTURA VENTA EL 21%     01/08/2005
'   Sql = "UPDATE facturaventa SET porcentajeiva=0.21 WHERE porcentajeiva=0"
'   DataEnvironment1.Sistema.Execute Sql
''**************************************************************************
'' PROCESO PARA CAMBIAR LA DENOMINACION DEL EJERCICIO     01/08/2005
'   Sql = "Update ejercicio SET denominacion=year(fechainicio)"
'   DataEnvironment1.Sistema.Execute Sql
''**************************************************************************

MsgBox "Proceso Terminado", vbExclamation, "Aviso"
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub cmdUnicaVez_Click()
Dim rsProd As New ADODB.Recordset
Dim rsfact As New ADODB.Recordset
    'Este proceso por la nueva migracion del 19/08/2005 Gaston
    DataEnvironment1.Sistema.Execute "update FacturaVenta set formapago = 1"
     
rsfact.Open "SELECT id,producto,descripcion from facturaventadetalle WHERE descripcion is null", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
Do While Not rsfact.EOF
   rsProd.Open "SELECT descripcion,codigo from producto where codigo='" & Trim(rsfact!Producto) & "'", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
   If Not rsProd.EOF Then
    DataEnvironment1.Sistema.Execute "UPDATE facturaventadetalle SET descripcion= '" & Trim(rsProd!DESCRIPCION) & "' WHERE id=" & rsfact!ID & ""
   End If
   rsProd.Close
rsfact.MoveNext
Loop

che "listo"
End Sub

Private Sub Command1_Click()
Dim resu As String
    resu = frmBuscar.MostrarSql("select codigo, TipoDoc, NroFactura, fecha from FacturaVenta where activo = 0 ")
    If resu = "" Then
        Text1 = ""
        Text2 = ""
    Else
        With frmBuscar
        Text1 = .resultado(1)
        Text2 = .resultado(2) & " " & .resultado(3) & " " & .resultado(4)
        End With
    End If
End Sub

Private Sub Command2_Click()
    Dim resu As String, i As Long
    resu = frmBuscar.MostrarSql("select NroDoc, TipoDoc, CodPr, total, fecha, razonsocialprov from compras order by fecha desc")
    If resu = "" Then
        FrmBorrarTxt Me
    Else
        For i = 0 To 5: txtcomp(i) = frmBuscar.resultado(i + 1): Next i
    End If
End Sub

Private Sub Command3_Click()
    If Trim(txtcomp(0)) = "" Then Exit Sub
    DataEnvironment1.dbo_DECOMPRASATRANSCOM s2n(txtcomp(2)), Trim(txtcomp(1)), s2n(txtcomp(0))
    DataEnvironment1.dbo_MODIFICOSALDOTRANS s2n(txtcomp(2)), Trim(txtcomp(1)), s2n(txtcomp(0)), s2n(txtcomp(3))
    
    MsgBox "ya ta"
    
End Sub

Private Sub Form_Load()
    FrmBorrarTxt Me
End Sub

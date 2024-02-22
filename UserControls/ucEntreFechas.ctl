VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ucEntreFechas 
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2790
   ScaleHeight     =   330
   ScaleWidth      =   2790
   Begin MSComCtl2.DTPicker dtDesde 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   21299201
      CurrentDate     =   38229
   End
   Begin MSComCtl2.DTPicker dtHasta 
      Height          =   315
      Left            =   1380
      TabIndex        =   1
      Top             =   0
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   393216
      Format          =   21299201
      CurrentDate     =   38229
   End
End
Attribute VB_Name = "ucEntreFechas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit ' mod 17/11/4
'Lito Explicit 26/9/4


Private mOrientacion As ucefOrientacion
Private mFormatoSql As ucefFormatoSql

Public Enum ucefOrientacion
    ucefHorizontal
    ucefVertical
End Enum

'El formato string para SQL siempre es MM-DD-YY
Public Enum ucefFormatoSql
   ucefFormatoSqlServer     '   WHERE FECHA = CONVERT(datetime, '12-31-04', 1)
   ucefFormatoSqlAccess     '   WHERE FECHA = #12-31-04#
   'ucefFormato
End Enum

Public Property Get Formatosql() As ucefFormatoSql
    Formatosql = mFormatoSql
End Property
Public Property Let Formatosql(que As ucefFormatoSql)
    mFormatoSql = que
End Property
Public Property Get orientacion() As ucefOrientacion
    orientacion = mOrientacion
End Property
Public Property Let orientacion(que As ucefOrientacion)
    mOrientacion = que
End Property
Public Property Get desde() As Date
    desde = dtDesde
End Property
Public Property Let desde(que As Date)
    dtDesde = que
End Property
Public Property Get hasta() As Date
    hasta = dtHasta
End Property
Public Property Let hasta(que As Date)
    dtHasta = que
End Property

Public Sub ini(Optional fDesde, Optional fHasta, Optional orientacionVH As ucefOrientacion = ucefHorizontal, Optional uformatosql As ucefFormatoSql = ucefFormatoSqlServer)
    If Not IsMissing(fDesde) Then dtDesde = fDesde
    If Not IsMissing(fHasta) Then dtHasta = fHasta
    mOrientacion = orientacionVH
    mFormatoSql = uformatosql
End Sub

Public Function ssDesde() As String
    d2s (dtDesde)
End Function
Public Function ssHasta() As String
    ssHasta = d2s(dtHasta)
End Function
Public Function ssBetween() As String
    ssBetween = " between " & d2s(dtDesde) & "  AND " & d2s(dtHasta)
End Function

Private Function d2s(dFecha As Date) As String
    Dim sFe As String
    sFe = Format(dFecha, "mm-dd-yy")
    If mFormatoSql = ucefFormatoSqlServer Then
        d2s = " convert (datetime, '" & sFe & "', 1) "
    Else
        d2s = " #" & sFe & "# "
    End If
End Function


' -------- privado --------------------
Private Sub UserControl_Initialize()
    mOrientacion = ucefHorizontal
    mFormatoSql = ucefFormatoSqlServer
    
    dtDesde = "1/1/" & Year(Date)
    dtHasta = CDate("1/1/" & (Year(Date) + 1)) - 1
End Sub

Private Sub UserControl_Resize()
    If mOrientacion = ucefHorizontal Then
        dtDesde.Top = 0
        dtDesde.Left = 0
        dtDesde.Height = UserControl.Height
        dtDesde.Width = UserControl.Width / 2

        dtHasta.Top = 0
        dtHasta.Left = UserControl.Width / 2
        dtHasta.Height = UserControl.Height
        dtHasta.Width = UserControl.Width / 2
    Else
        dtDesde.Top = 0
        dtDesde.Left = 0
        dtDesde.Height = UserControl.Height
        dtDesde.Width = UserControl.Width / 2

        dtHasta.Top = 0
        dtHasta.Left = UserControl.Width / 2
        dtHasta.Height = UserControl.Height
        dtHasta.Width = UserControl.Width / 2
    End If
End Sub

'2/11/4
'   Fix default en orientacion
'   Fix formato access
'17/11/4
'   Fix Stype = 1 como parametro al convert MSSQL
'14/2/5
'   ssXXXX() as string
'

VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmParteProduccion 
   Caption         =   "Generar de Parte de Produccion"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9405
   Icon            =   "frmParteProduccion.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   9405
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtParte 
      Enabled         =   0   'False
      Height          =   285
      Left            =   615
      TabIndex        =   22
      Text            =   "0"
      Top             =   60
      Width           =   840
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   345
      TabIndex        =   1
      Top             =   900
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      Format          =   58523649
      CurrentDate     =   39681
   End
   Begin VB.TextBox txtObs 
      BackColor       =   &H8000000F&
      Height          =   330
      Left            =   1650
      MaxLength       =   30
      TabIndex        =   20
      Top             =   5625
      Width           =   7455
   End
   Begin VB.Frame Frame2 
      Height          =   1200
      Left            =   4320
      TabIndex        =   18
      Top             =   6045
      Width           =   4905
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar"
         Height          =   1035
         Left            =   915
         Picture         =   "frmParteProduccion.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   915
      End
      Begin VB.CommandButton cmBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar"
         DisabledPicture =   "frmParteProduccion.frx":1194
         Height          =   1035
         Left            =   30
         Picture         =   "frmParteProduccion.frx":2E8E
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   870
      End
      Begin VB.CommandButton cmdAceptar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Aceptar"
         DisabledPicture =   "frmParteProduccion.frx":4B88
         Enabled         =   0   'False
         Height          =   1035
         Left            =   2145
         Picture         =   "frmParteProduccion.frx":5452
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   870
      End
      Begin VB.CommandButton cmdCancelar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cancelar"
         Height          =   1035
         Left            =   3030
         Picture         =   "frmParteProduccion.frx":5D1C
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   915
      End
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Salir"
         Height          =   1035
         Left            =   3960
         Picture         =   "frmParteProduccion.frx":65E6
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   915
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1290
      Left            =   240
      TabIndex        =   17
      Top             =   5985
      Width           =   2550
      Begin VB.CommandButton cmdAgregar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Agregar componente"
         DisabledPicture =   "frmParteProduccion.frx":6EB0
         Enabled         =   0   'False
         Height          =   1125
         Left            =   1305
         Picture         =   "frmParteProduccion.frx":777A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1200
      End
      Begin VB.CommandButton cmdModFormu 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Modificar Formulas"
         DisabledPicture =   "frmParteProduccion.frx":8044
         Enabled         =   0   'False
         Height          =   1125
         Left            =   45
         Picture         =   "frmParteProduccion.frx":890E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1245
      End
   End
   Begin VB.Frame fraEditDetalle 
      Height          =   1635
      Left            =   120
      TabIndex        =   15
      Top             =   315
      Width           =   9120
      Begin VB.TextBox txtFParte 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   320
         Left            =   3300
         TabIndex        =   4
         Top             =   1215
         Width           =   1095
      End
      Begin VB.TextBox txtFSuma 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   320
         Left            =   3300
         TabIndex        =   3
         Top             =   870
         Width           =   1095
      End
      Begin VB.TextBox txtCantidad 
         Alignment       =   1  'Right Justify
         Height          =   320
         Left            =   3300
         TabIndex        =   2
         Top             =   525
         Width           =   1095
      End
      Begin VB.CommandButton cmdAgregarItem 
         Height          =   600
         Left            =   7230
         MaskColor       =   &H00E0E0E0&
         Picture         =   "frmParteProduccion.frx":91D8
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   540
         Width           =   630
      End
      Begin VB.CommandButton cmdBorrarItem 
         Height          =   600
         Left            =   7890
         MaskColor       =   &H00E0E0E0&
         Picture         =   "frmParteProduccion.frx":9EA2
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Borrar Item"
         Top             =   540
         Width           =   660
      End
      Begin Gestion.ucCoDe uProd 
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Top             =   180
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   556
         CodigoWidth     =   1000
      End
      Begin VB.Label Label4 
         Caption         =   "Cantidad p/Parte"
         Height          =   255
         Left            =   1845
         TabIndex        =   24
         Top             =   1245
         Width           =   1755
      End
      Begin VB.Label Label3 
         Caption         =   "Cantidad a Sumar:"
         Height          =   255
         Left            =   1845
         TabIndex        =   23
         Top             =   915
         Width           =   1755
      End
      Begin VB.Label Label10 
         Caption         =   "Cantidad a producir:"
         Height          =   255
         Left            =   1845
         TabIndex        =   16
         Top             =   570
         Width           =   1755
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Height          =   3225
      Left            =   165
      TabIndex        =   14
      Top             =   2385
      Width           =   8985
      _cx             =   15849
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
   Begin VB.Label Label5 
      Caption         =   "Haciendo doble clic sobre el renglon cambia su valor al del factor"
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   150
      TabIndex        =   25
      Top             =   2025
      Width           =   9015
   End
   Begin VB.Label Label2 
      Caption         =   "Parte"
      Height          =   255
      Left            =   150
      TabIndex        =   21
      Top             =   90
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Observaciones :"
      Height          =   210
      Left            =   465
      TabIndex        =   19
      Top             =   5670
      Width           =   1215
   End
End
Attribute VB_Name = "frmParteProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public modificar As Long
Private WithEvents g As LiGrilla
Attribute g.VB_VarHelpID = -1
Private gCANT   As Long
Private gUMC As Long
Private gUMD As Long
Private gULT As Long

Private gprod   As Long
Private gDESC   As Long
Private gPUNI   As Long
Private gPTOT   As Long
Private gNPED   As Long

Private busco As Boolean

Private Sub cmBuscar_Click()
Dim p, rsC As New ADODB.Recordset, dat1
Dim C, cFactor As Double, cCargar As Double
Set C = Nothing
    p = frmBuscar.MostrarSql("Select nro as Parte,Fecha,Producido as [Producto Generado],Cantidad,Activo from partesproduccion order by nro", , "Partes de Produccion", "N/D", "Si", "No")
    If p > "" Then
        txtParte = p
        Set dat1 = Nothing
        dat1 = obtenerDeSQL("select fecha,producido,cantidad from partesproduccion where nro=" & p)
        uProd.codigo = sSinNull(dat1(1))
            C = obtenerDeSQL("select f.factor from producto p inner join umfactor f on p.ufactor=f.ufcodigo where p.codigo=" & ssTexto(uProd.codigo))
            If IsNull(C) Or IsEmpty(C) Then
                cFactor = 1
            Else
                cFactor = C
            End If
        txtCantidad = nSinNull(dat1(2)) / cFactor
        DTPicker1.Value = fSinNull(dat1(0))
        LlenarGrilla grilla, "select " & txtCantidad & " as Cantidad,d.codigo_articulo as Producido, " & ssTexto(uProd.DESCRIPCION) & " as Descripcion, d.Componente, p.descripcion as [Descripcion del componente],d.Cantidad as [Cantidad utilizada],0 as CdM,u.abreviatura as Medida from formulasdetalle d inner join (producto p inner join unidadesmedida u on p.umedida=u.umcodigo)on d.componente=p.codigo where d.activo=1 and d.codigo_parte=" & p, False
        estPart True
        If frmBuscar.resultado(5) = False Then
            cmdeliminar.enabled = False
        Else
            cmdeliminar.enabled = True
        End If
    End If
End Sub

Private Sub cmdAceptar_Click()
On Error GoTo ppmal
If busco Then Exit Sub
    Dim Nro As Long
    Dim i As Long
    Dim neww
    
    For i = 1 To g.rows - 1
        grilla.Row = i
        If grilla.TextMatrix(grilla.RowSel, 3) = "" Then
            MsgBox "Debe seleccionar un componente para el producto."
            grilla.Col = 3
            grilla.ColSel = 3
            grilla.RowSel = grilla.Row
            Exit Sub
        Else
            If grilla.TextMatrix(grilla.RowSel, 4) = "" Then
                MsgBox "Debe seleccionar un componente para el producto."
                grilla.Col = 4
                grilla.ColSel = 4
                grilla.RowSel = grilla.Row
                Exit Sub
            Else
                If grilla.TextMatrix(grilla.RowSel, 5) = 0 Then
                    MsgBox "Debe completar la cantidad del componente utilizado."
                    grilla.Col = 5
                    grilla.ColSel = 5
                    grilla.RowSel = grilla.Row
                    Exit Sub
                End If
            End If
        End If
    Next i
    neww = obtenerDeSQL("select max(nro) as num from partesproduccion")
    If IsNull(neww) Or IsEmpty(neww) Then
        Nro = 1
    Else
        Nro = neww + 1
    End If
    DE_BeginTrans
    If ABMParteProduccion("A", Nro, DTPicker1.Value, DTPicker1.Value, txtobs, uProd.codigo, s2n(txtCantidad), UsuarioActual()) = False Then GoTo ppmal
    For i = 1 To g.rows - 1
        grilla.Row = i
        If grilla.TextMatrix(grilla.RowSel, 5) > 0 And grilla.TextMatrix(grilla.RowSel, 3) > "" And grilla.TextMatrix(grilla.RowSel, 4) > "" Then
            If ABMPPItem("A", Nro, uProd.codigo, grilla.TextMatrix(grilla.RowSel, 3), s2n(grilla.TextMatrix(grilla.RowSel, 5)), UsuarioActual()) = False Then GoTo ppmal
        End If
    Next i
   
    DE_CommitTrans
    MsgBox "Parte de produccion generada.", vbInformation

    
    cmdCancelar_Click
    grilla.Editable = flexEDNone
    modificar = 0
Exit Sub
ppmal:
DE_RollbackTrans
    MsgBox "Error al guardar parte de produccion", vbCritical
End Sub

Private Sub cmdAgregar_Click()
    Dim producto As String
    Dim descrip As String
    Dim canti As Long
    Dim i As Long
    Dim vvalor1(0 To 7) As Variant
    Dim vvalor2(0 To 7) As Variant
    Dim ultimo As Integer
    Dim reglonUNO As Boolean
    
    ultimo = 1
    
    If grilla.ColSel = -1 And grilla.RowSel = 0 Then
        MsgBox "Debe seleccionar algun producto de la grilla.", vbInformation
    Else
        i = g.addRow()
        
        reglonUNO = True
                        
        Do While ultimo = 1
            If grilla.ValueMatrix(grilla.RowSel, 0) > 0 Or ultimo = 1 Then
                If reglonUNO = True Then

                    producto = grilla.TextMatrix(grilla.RowSel, 1)
                    descrip = grilla.TextMatrix(grilla.RowSel, 2)
                    canti = grilla.ValueMatrix(grilla.RowSel, 0)
                    
                    grilla.RowSel = grilla.RowSel + 1
                    i = grilla.RowSel
                    If Not grilla.ValueMatrix(grilla.RowSel, 0) > 0 Then ultimo = 0

                    vvalor1(0) = grilla.ValueMatrix(grilla.RowSel, 0)
                    vvalor1(1) = grilla.TextMatrix(grilla.RowSel, 1)
                    vvalor1(2) = grilla.TextMatrix(grilla.RowSel, 2)
                    vvalor1(3) = grilla.TextMatrix(grilla.RowSel, 3)
                    vvalor1(4) = grilla.TextMatrix(grilla.RowSel, 4)
                    vvalor1(5) = grilla.ValueMatrix(grilla.RowSel, 5)
                                        
                    g.tx i, gprod, producto
                    g.tx i, gDESC, descrip
                    g.tx i, gCANT, canti
                    
                    g.tx i, gPUNI, ""
                    g.tx i, gPTOT, ""
                    g.tx i, gNPED, "0"
                    reglonUNO = False
                Else
                    vvalor2(0) = grilla.TextMatrix(grilla.RowSel, 0)
                    vvalor2(1) = grilla.TextMatrix(grilla.RowSel, 1)
                    vvalor2(2) = grilla.TextMatrix(grilla.RowSel, 2)
                    vvalor2(3) = grilla.TextMatrix(grilla.RowSel, 3)
                    vvalor2(4) = grilla.TextMatrix(grilla.RowSel, 4)
                    vvalor2(5) = grilla.ValueMatrix(grilla.RowSel, 5)
                    
                    i = grilla.RowSel
                    If Not grilla.ValueMatrix(grilla.RowSel, 0) > 0 Then ultimo = 0
                    g.tx i, gprod, vvalor1(1)
                    g.tx i, gDESC, vvalor1(2)
                    g.tx i, gCANT, vvalor1(0)
                    
                    g.tx i, gPUNI, vvalor1(3)
                    g.tx i, gPTOT, vvalor1(4)
                    g.tx i, gNPED, vvalor1(5)
                    
                    vvalor1(0) = vvalor2(0)
                    vvalor1(1) = vvalor2(1)
                    vvalor1(2) = vvalor2(2)
                    vvalor1(3) = vvalor2(3)
                    vvalor1(4) = vvalor2(4)
                    vvalor1(5) = vvalor2(5)
                End If
            End If
            If ultimo = 1 Then grilla.RowSel = grilla.RowSel + 1
        Loop
        grilla.Col = -1
    End If
End Sub

Private Sub cmdAgregarItem_Click()
If busco Then
    grilla.rows = 1
    inigrilla
    estPart False
End If
    Dim r As Long, pco, pde
    pco = uProd.codigo
    pde = Trim(uProd.DESCRIPCION)

    If pco = "" And pde = "" Then
        Exit Sub
    End If
    If s2n(txtCantidad, 4) = 0 Or s2n(txtCantidad, 4) = "" Then
        MsgBox "Falta especificar cantidad.", vbExclamation
        Exit Sub
    End If
    
    
    MetoEnGrilla pco, pde, s2n(txtCantidad, 4)
    'txtCantidad = ""
    'uProd.codigo = ""
    uProd.SetFocus
    cmdAceptar.enabled = True
    cmdModFormu.enabled = True

End Sub

Private Sub cmdBorrarItem_Click()
    If g.Row > 0 Then g.delRow (g.Row)
End Sub

Private Sub cmdCancelar_Click()
    uProd.codigo = 0
    txtCantidad = ""
    txtFSuma = ""
    txtFParte = ""
    txtParte = nuevoparte
    DTPicker1 = Date
    g.Borrar
    grilla.Editable = flexEDNone
    modificar = 0
    cmdAceptar.enabled = False
    cmdModFormu.enabled = False
    cmdAgregar.enabled = False
    estPart False
End Sub

Private Sub cmdeliminar_Click()
On Error GoTo emal
Dim i As Long
    If busco Then
        DE_BeginTrans
        If ABMParteProduccion("B", s2n(txtParte), DTPicker1.Value, DTPicker1.Value, txtobs, uProd.codigo, s2n(txtCantidad), UsuarioActual()) = False Then GoTo emal
        For i = 1 To g.rows - 1
            grilla.Row = i
            If grilla.TextMatrix(grilla.RowSel, 6) > 0 And grilla.TextMatrix(grilla.RowSel, 3) > "" And grilla.TextMatrix(grilla.RowSel, 4) > "" Then
                If ABMPPItem("B", s2n(txtParte), uProd.codigo, grilla.TextMatrix(grilla.RowSel, 3), s2n(grilla.TextMatrix(grilla.RowSel, 5)), UsuarioActual()) = False Then GoTo emal
            End If
        Next i
        DE_CommitTrans
        MsgBox "Parte de produccion eliminada.", vbInformation
        cmdCancelar_Click
    End If
Exit Sub
emal:
DE_RollbackTrans
    MsgBox "Error al eliminar parte de produccion", vbCritical
End Sub

Private Sub cmdModFormu_Click()
    If grilla.ColSel = -1 And grilla.RowSel = 0 Then

    Else
        grilla.Editable = flexEDKbdMouse
        modificar = 1
        cmdAgregar.enabled = True
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rsEjercicio As New ADODB.Recordset
    Dim rsFormaP As New ADODB.Recordset
    Dim sql As String

    estPart False
    modificar = 0
    
    inigrilla
    set_uProd
 
   g.Borrar
   grilla.Editable = flexEDNone
   DTPicker1 = Date
   txtParte = nuevoparte
   
End Sub

Private Function nuevoparte() As Long
Dim neww1
    neww1 = obtenerDeSQL("select max(nro) as num from partesproduccion")
    If IsNull(neww1) Or IsEmpty(neww1) Then
        nuevoparte = 1
    Else
        nuevoparte = neww1 + 1
    End If
End Function


Private Function estPart(ep As Boolean)
    busco = ep
    cmdeliminar.enabled = ep
    cmdAceptar.enabled = Not ep
End Function

Private Sub set_uProd()
    Dim sqlbuscar As String, sqldesc As String

        sqldesc = "select descripcion from producto where codigo = '###' "
        sqlbuscar = "select p.codigo as [ Codigo                 ],  p.descripcion as [ Descripcion                                                 ],m.abreviatura as [  Medida  ] from producto as p inner join unidadesmedida as m on p.umedida=m.umcodigo where p.activo = 1 and formula=1 order by p.codigo "

    uProd.ini sqldesc, sqlbuscar, True
End Sub

Private Sub MetoEnGrilla(prod, desc, cant)
    If ON_ERROR_HABILITADO Then On Error GoTo ufaErr
    Dim e, tFactor As Double, tCargar As Double, x As Long
    Dim i As Long, rs As New ADODB.Recordset, ssql As String, codigomio As String, hay As Double, suma As Double
    
    ssql = "SELECT f.codigo,p.descripcion,f.componente,pp.descripcion as descripcion2,u.abreviatura,f.cantidad from formulas f inner join producto p on f.codigo=p.codigo inner join (producto pp inner join unidadesmedida u on p.umedida=u.umcodigo)on f.componente=pp.codigo WHERE f.activo =1 AND (f.CODIGO = '" & prod & "') Order By  F.Codigo"
    rs.Open "SELECT f.codigo,p.descripcion,f.componente,pp.descripcion as descripcion2,f.cantidad,u.umCodigo as UMC,u.Abreviatura  from formulas f inner join producto p on f.codigo=p.codigo inner join (producto pp inner join unidadesmedida u on pp.umedida=u.umcodigo)on f.componente=pp.codigo WHERE f.activo =1 AND (f.CODIGO = '" & prod & "') Order By  F.Codigo", DataEnvironment1.Sistema, adOpenKeyset, adLockReadOnly
    

    If rs.EOF = True And rs.BOF = True Then
        MsgBox "No se ha encontrado formula para este producto.", vbExclamation
    Else
        rs.MoveFirst
        With rs
        g.rows = 1
            For x = 0 To .RecordCount - 1
                
                i = g.addRow()
                g.tx i, gprod, prod
                g.tx i, gDESC, desc
                
                'OBTENEMOS EL FACTOR DE CONVERSION DEL COMPONENTE
                'Y  LO MULTIPLICAMOS POR LA CANTIDAD
                'SI NO TIENE FACTOR SE MULTIPLICA POR 0
                Set e = Nothing
                e = obtenerDeSQL("select f.factor from producto p inner join umfactor f on p.ufactor=f.ufcodigo where p.codigo=" & ssTexto(rs!Componente))
                If IsNull(e) Or IsEmpty(e) Then
                    tFactor = 1
                Else
                    tFactor = e
                End If
                tCargar = tFactor * cant * s2n(rs!cantidad)
                
                g.tx i, gCANT, cant
                g.tx i, gPUNI, rs!Componente
                g.tx i, gPTOT, rs!descripcion2
                g.tx i, gUMC, rs!umc
                g.tx i, gUMD, rs!abreviatura
                g.tx i, gNPED, tCargar
                g.tx i, gULT, ""
                
                rs.MoveNext
            Next
        End With
    End If
Exit Sub
ufaErr:
    MsgBox "Error en la carga de formula", vbCritical
End Sub

Private Sub inigrilla()
    Set g = New LiGrilla
    
    g.init grilla, 4
    gCANT = g.AddCol(" Cantidad    ", 4)
    gprod = g.AddCol(" Producir            ")
    gDESC = g.AddCol(" Descripcion                                ")
    gPUNI = g.AddCol(" Componente    ")
    gPTOT = g.AddCol(" Descripcion de componente                  ")
    gNPED = g.AddCol(" Cantidad ulitizada   ", "N", 4)
    gUMC = g.AddCol("  UMCodigo ", "H")
    gUMD = g.AddCol(" Medida   ")
    gULT = g.AddCol("  Valor ", "H")

End Sub

Private Sub g_DblClick()
    Dim re As String
    If g.Row = 0 Or g.Col <> gPUNI Or modificar = 0 Then Exit Sub
    
    re = frmBuscar.MostrarCodigoDescripcionActivo("producto")
    If re = "" Then Exit Sub
    
    g.tx g.Row, gPUNI, re
    g.tx g.Row, gPTOT, frmBuscar.resultado(2)
End Sub

Public Function ABMParteProduccion(pOPE As String, pNRO As Long, pFecha As Date, pConfir As Date, pObs As String, pProd As String, pCantidad As Double, pUsu As Long) As Boolean
On Error GoTo pmal
ABMParteProduccion = True
Dim iudp As String, e, eFactor As Double, eCargar As Double
Dim Alma As Integer

Set e = Nothing
e = obtenerDeSQL("select f.factor from producto p inner join umfactor f on p.ufactor=f.ufcodigo where p.codigo=" & ssTexto(pProd))
If IsNull(e) Or IsEmpty(e) Then
    eFactor = 1
Else
    eFactor = e
End If
eCargar = eFactor * pCantidad

Alma = s2n(obtenerDeSQL("select almacen from producto where codigo='" & Trim(pProd) & "'"))
'EL DETALLE DE LA PARTE ESTA EN FORMULASDETALLE
Select Case pOPE
    Case "A":
        iudp = "INSERT INTO PARTESPRODUCCION (NRO,FECHA,CONFIRMACION,OBSERVACIONES,PRODUCIDO,CANTIDAD,FECHA_ALTA,USUARIO_ALTA,ACTIVO) " _
            & " VALUES( " & pNRO & "," & ssFecha(pFecha) & ", " & ssFecha(pConfir) & ", " & ssTexto(pObs) & "," & ssTexto(pProd) & "," & x2s(eCargar) & "," & ssFecha(Date) & "," & pUsu & ",1)"
        DataEnvironment1.Sistema.Execute iudp
        iudp = "UPDATE PRODUCTO " _
            & " SET EXISTENCIA= EXISTENCIA + (" & x2s(eCargar) _
            & ") WHERE CODIGO=" & ssTexto(pProd)
        DataEnvironment1.Sistema.Execute iudp
        
        If Alma <> 0 Then DataEnvironment1.dbo_SumaStock pProd, eCargar, Alma
    Case "M": 'NO HAY MODIFICACION
    Case "B":
        iudp = "UPDATE PARTESPRODUCCION " _
            & " SET ACTIVO=0,FECHA_BAJA=" & ssFecha(Date) & ", USUARIO_BAJA=" & pUsu _
            & " WHERE NRO=" & pNRO
        DataEnvironment1.Sistema.Execute iudp
        iudp = "UPDATE PRODUCTO " _
            & " SET EXISTENCIA= EXISTENCIA - (" & x2s(eCargar) _
            & ") WHERE CODIGO=" & ssTexto(pProd)
        DataEnvironment1.Sistema.Execute iudp
        
        If Alma <> 0 Then DataEnvironment1.dbo_SumaStock pProd, -eCargar, Alma
End Select
Exit Function
pmal:
ABMParteProduccion = False
End Function

Public Function ABMPPItem(tOpe As String, tParte As Long, tProducido As String, tComponente As String, tCantidad As Double, tUsu As Long) As Boolean
On Error GoTo ppimal
ABMPPItem = True
Dim iudt As String
Dim Alma As Integer

Alma = s2n(obtenerDeSQL("select almacen from producto where codigo='" & Trim(tComponente) & "'"))
Select Case tOpe
    Case "A":
        'ALTA EN DETALLE DE FORMULA QUE SIRVE PARA TERNER EL DETALLE DE LA PARTE Y ADEMAS UN HISTORIAL DE LA FORMULA
        iudt = "INSERT INTO FORMULASDETALLE (CODIGO_PARTE,CODIGO_ARTICULO,COMPONENTE,CANTIDAD,FECHA_ALTA,USUARIO_ALTA,ACTIVO) " _
            & " VALUES( " & tParte & "," & ssTexto(tProducido) & "," & ssTexto(tComponente) & ", " & x2s(tCantidad) & " , " & ssFecha(Date) & ", " & tUsu & ",1)"
        DataEnvironment1.Sistema.Execute iudt
                    
        'ACTUALIZAMOS LA EXISTENCIA DEL COMPONENTE, LA CANTIDAD TIENE QUE VENIR CONVERTIDO A LA UNIDAD DE DICHO COMPONENTE
        iudt = "UPDATE PRODUCTO" _
            & " SET EXISTENCIA= EXISTENCIA - (" & x2s(tCantidad) & ") " _
            & " WHERE CODIGO=" & ssTexto(tComponente)
        DataEnvironment1.Sistema.Execute iudt
        
        If Alma <> 0 Then DataEnvironment1.dbo_SumaStock tComponente, -tCantidad, Alma
    Case "B":
        'ACTUALIZO EL DETALLE DE LA FORMULA PARA FUTUROS REPORTES
        iudt = "UPDATE FORMULASDETALLE " _
            & " SET  FECHA_BAJA=" & ssFecha(Date) & ",ACTIVO=0, USUARIO_BAJA=" & tUsu _
            & " WHERE CODIGO_PARTE=" & tParte & " AND COMPONENTE=" & ssTexto(tComponente)
        DataEnvironment1.Sistema.Execute iudt
                    
        'ACTUALIZAMOS LA EXISTENCIA DEL COMPONENTE, LA CANTIDAD TIENE QUE VENIR CONVERTIDO A LA UNIDAD DE DICHO COMPONENTE
        iudt = "UPDATE PRODUCTO" _
            & " SET EXISTENCIA= EXISTENCIA + (" & x2s(tCantidad) & ") " _
            & " WHERE CODIGO=" & ssTexto(tComponente)
        DataEnvironment1.Sistema.Execute iudt
        
        If Alma <> 0 Then DataEnvironment1.dbo_SumaStock tComponente, tCantidad, Alma
End Select
Exit Function
ppimal:
ABMPPItem = False
End Function


Private Sub grilla_DblClick()
    If grilla.TextMatrix(grilla.Row, gULT) = "" Then
        grilla.TextMatrix(grilla.Row, gULT) = grilla.TextMatrix(grilla.Row, gNPED)
        grilla.TextMatrix(grilla.Row, gNPED) = txtFParte
    Else
        grilla.TextMatrix(grilla.Row, gNPED) = grilla.TextMatrix(grilla.Row, gULT)
        grilla.TextMatrix(grilla.Row, gULT) = ""
    End If
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAgregarItem.SetFocus
    End If
End Sub

Private Sub txtCantidad_LostFocus()
Dim e, tFactor
    If s2n(txtCantidad) < 0 Then
        txtCantidad = 0
    ElseIf s2n(txtCantidad) > 0 Then
        Set e = Nothing
        e = obtenerDeSQL("select f.factor from producto p inner join umfactor f on p.ufactor=f.ufcodigo where p.codigo=" & ssTexto(uProd.codigo))
        If IsNull(e) Or IsEmpty(e) Then
            txtFSuma = 1 * s2n(txtCantidad)
        Else
            txtFSuma = e * s2n(txtCantidad)
        End If
        Set e = Nothing
        
        e = obtenerDeSQL("select f.factor from producto p inner join umfactor f on p.uparte=f.ufcodigo where p.codigo=" & ssTexto(uProd.codigo))
        If IsNull(e) Or IsEmpty(e) Then
            txtFParte = 1 * s2n(txtCantidad)
        Else
            txtFParte = e * s2n(txtCantidad)
        End If

    Else
        txtFSuma = 0
        txtFParte = 0
    End If
    
End Sub

Private Sub uProd_cambio(codigo As Variant)
    grilla.rows = 1
    inigrilla
End Sub

    'ANTES
    'cadena = "insert into partesproduccion" _
            & " (nro,fecha,confirmacion,observaciones,fecha_alta,fecha_baja,usuario_alta,usuario_baja,activo) " _
            & " values( " & Nro & "," & ssFecha(DTPicker1.Value) & ", " & ssFecha(DTPicker1.Value) & ", '" & txtObs.Text & "'," & ssFecha(Date) & ",01/01/1900," & UsuarioActual() & "," _
            & " 0, 1)"
    'DataEnvironment1.Sistema.Execute cadena
    
    
    
                'cadena = "insert into formulasdetalle" _
                '& " (codigo_parte,codigo_articulo,componente,cantidad,fecha_alta,usuario_alta,activo) " _
                '& " values( " & Nro & ",'" & grilla.TextMatrix(grilla.RowSel, 1) & "', '" & grilla.TextMatrix(grilla.RowSel, 3) & "', " & x2s(grilla.TextMatrix(grilla.RowSel, 5)) & "" _
                '& " , " & ssFecha(Date) & ", " & UsuarioActual() & ",1)"
            'DataEnvironment1.Sistema.Execute cadena
            
            'Set parte = Nothing
            'pExistencia = obtenerDeSQL("select existencia from producto where codigo=" & sstexto(grilla.TextMatrix(grilla.RowSel, 3)))
            
            'cadena = "update producto" _
                & " set existencia= " & x2s(pExistencia - grilla.TextMatrix(grilla.RowSel, 5)) & " " _
                & " where codigo='" & grilla.TextMatrix(grilla.RowSel, 3) & "'"
            'DataEnvironment1.Sistema.Execute cadena
            
            'If grilla.TextMatrix(grilla.RowSel, 1) <> Texto Then
            '    Texto = grilla.TextMatrix(grilla.RowSel, 1)
            '    cadena = "insert into itempartesproduccion" _
            '        & " (parte,producto,cantidad) " _
            '        & " values( " & Nro & ",'" & Texto & "', '" & x2s(grilla.TextMatrix(grilla.RowSel, 0)) & "')"
            '    DataEnvironment1.Sistema.Execute cadena
            '    pExistencia = obtenerDeSQL("select existencia from producto where codigo=" & sstexto(grilla.TextMatrix(grilla.RowSel, 1)))

                
             '   cadena = "update producto" _
             '       & " set existencia= " & x2s(pExistencia + grilla.TextMatrix(grilla.RowSel, 0)) & " " _
             '       & " where codigo='" & grilla.TextMatrix(grilla.RowSel, 1) & "'"
             '   DataEnvironment1.Sistema.Execute cadena


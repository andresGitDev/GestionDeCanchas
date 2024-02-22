VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDespacho 
   Caption         =   "Carga de Depacho"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8565
   Icon            =   "frmDespacho.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   8565
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1800
      TabIndex        =   18
      Top             =   915
      Width           =   6645
   End
   Begin MSComCtl2.DTPicker dtFecha 
      Height          =   315
      Left            =   6705
      TabIndex        =   16
      Top             =   90
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      Format          =   80936961
      CurrentDate     =   39611
   End
   Begin VB.CommandButton cmdAfuera 
      Caption         =   "Sacar"
      Height          =   570
      Left            =   7440
      Picture         =   "frmDespacho.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2865
      Width           =   675
   End
   Begin VB.TextBox txtObs 
      Height          =   315
      Left            =   1785
      TabIndex        =   13
      Top             =   510
      Width           =   6645
   End
   Begin VB.Frame fraDoc 
      Caption         =   "Elija el tipo de documento a Cargar"
      Height          =   1335
      Left            =   60
      TabIndex        =   6
      Top             =   1350
      Width           =   8430
      Begin VB.CommandButton cmdAdentro 
         Caption         =   "Ingresar"
         Height          =   510
         Left            =   4560
         Picture         =   "frmDespacho.frx":0C54
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   690
         Width           =   735
      End
      Begin VB.OptionButton Op3 
         Caption         =   "Facturas de Venta"
         Height          =   330
         Left            =   135
         TabIndex        =   9
         Top             =   840
         Width           =   1635
      End
      Begin VB.OptionButton Op2 
         Caption         =   "Remitos de Venta"
         Height          =   240
         Left            =   135
         TabIndex        =   8
         Top             =   570
         Width           =   2340
      End
      Begin VB.OptionButton Op1 
         Caption         =   "Pedido de Clientes"
         Height          =   240
         Left            =   135
         TabIndex        =   7
         Top             =   270
         Value           =   -1  'True
         Width           =   2220
      End
      Begin Gestion.ucCoDe ucDocumento 
         Height          =   315
         Left            =   2745
         TabIndex        =   10
         Top             =   270
         Width           =   5250
         _extentx        =   9260
         _extenty        =   556
         codigoinvalido  =   0
         codigowidth     =   1455
      End
   End
   Begin VB.TextBox txtNumero 
      Height          =   330
      Left            =   4395
      TabIndex        =   4
      Text            =   "0"
      Top             =   90
      Width           =   1545
   End
   Begin VB.TextBox txtCodigo 
      Enabled         =   0   'False
      Height          =   330
      Left            =   1785
      TabIndex        =   2
      Text            =   "0"
      Top             =   90
      Width           =   930
   End
   Begin VSFlex7LCtl.VSFlexGrid gDetalle 
      Height          =   4320
      Left            =   30
      TabIndex        =   1
      Top             =   2820
      Width           =   7320
      _cx             =   12912
      _cy             =   7620
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
   Begin Gestion.ucBotonera ucBotonera 
      Align           =   2  'Align Bottom
      Height          =   1500
      Left            =   0
      TabIndex        =   0
      Top             =   6675
      Width           =   8565
      _extentx        =   15108
      _extenty        =   2646
      msgconfirmasalir=   ""
      msgconfirmaeliminar=   ""
      msgconfirmacancelar=   ""
      captioneliminar =   "&Eliminar"
      Begin Gestion.ucXls ucXls1 
         Height          =   765
         Left            =   7395
         TabIndex        =   15
         Top             =   645
         Width           =   855
         _extentx        =   1508
         _extenty        =   1349
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Observacion 2"
      Height          =   315
      Left            =   105
      TabIndex        =   19
      Top             =   915
      Width           =   1860
   End
   Begin VB.Label Label4 
      Caption         =   "Fecha"
      Height          =   345
      Left            =   6165
      TabIndex        =   17
      Top             =   135
      Width           =   1065
   End
   Begin VB.Label Label3 
      Caption         =   "Observacion 1"
      Height          =   315
      Left            =   90
      TabIndex        =   12
      Top             =   510
      Width           =   1860
   End
   Begin VB.Label Label2 
      Caption         =   "Numero de Despacho"
      Height          =   285
      Left            =   2775
      TabIndex        =   5
      Top             =   120
      Width           =   2190
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo del Despacho"
      Height          =   315
      Left            =   90
      TabIndex        =   3
      Top             =   135
      Width           =   1755
   End
End
Attribute VB_Name = "frmDespacho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private g As New LiGrilla

Private Sub cmdAdentro_Click()
Dim fDoc, qEss As String, i As Long
Dim e
    If ucDocumento.codigo = 0 Then
        MsgBox "Indique documento.", vbExclamation
    Else
        If qES = "Factura Venta" Then
            fDoc = Trim(frmBuscar.resultado(2))
        Else
            fDoc = ""
        End If
        qEss = qES & " " & fDoc
        
        e = obtenerDeSQL("select numerodespacho from despachodetalle where documento='" & Trim(qEss) & "' and numero='" & Trim(ucDocumento.codigo) & "'")
        If IsNull(e) Or IsEmpty(e) Then
        Else
            If MsgBox("Este documento ya existe en el despacho : " & e & Chr(13) & "Desea continuar?", vbExclamation + vbYesNo) = vbNo Then
                Exit Sub
            End If
        End If
        For i = 1 To gDetalle.rows - 1
            If Trim(gDetalle.TextMatrix(i, 2)) = Trim(qEss) And Trim(gDetalle.TextMatrix(i, 3)) = Trim(ucDocumento.codigo) Then
                MsgBox "Este documento ya esta en la grilla.", vbExclamation
                Exit Sub
            End If
        Next
        
        txtObs.Text = txtObs & qEss & " " & ucDocumento.codigo & ", "
        gDetalle.AddItem 0 & Chr(9) & " " & Chr(9) & qEss & Chr(9) & ucDocumento.codigo
    End If
End Sub

Private Function qES() As String
If Op1.Value = True Then
    qES = "Pedido de Clientes"
ElseIf Op2.Value = True Then
    qES = "Remito Venta"
ElseIf Op3.Value = True Then
    qES = "Factura Venta"
End If
End Function

Private Sub cmdAfuera_Click()
    Dim i As Long
    If gDetalle.Row = 0 Then Exit Sub
    gDetalle.RemoveItem gDetalle.Row
    i = 1
    txtObs = ""
    While i < gDetalle.rows
        txtObs.Text = txtObs & gDetalle.TextMatrix(i, 2) & " " & gDetalle.TextMatrix(i, 3) & ", "
        i = i + 1
    Wend
End Sub

Private Sub Form_Load()
qCargo
cleard
txtCodigo = dNuevo
ucXls1.ini gDetalle, "C:\Despacho.xls"
ucBotonera.init True, True, True, False, True
End Sub

Private Sub cleard()
    txtCodigo = 0
    txtNumero = ""
    txtObs = ""
    Text1.Text = ""
    ucDocumento.codigo = 0
    dtfecha = Date
    gDetalle.rows = 1
    inigrilla
End Sub

Private Sub inigrilla()
    Set g = New LiGrilla
    g.init gDetalle
    g.rows = 1
    g.cols = 0
    g.AddCol "CodigoDetalle "
    g.AddCol "NumeroDespacho"
    g.AddCol "Documento                                                    "
    g.AddCol "Numero                                                       "
    gDetalle.ColHidden(0) = True
    gDetalle.ColHidden(1) = True
End Sub

Private Function qCargo()
If Op1.Value = True Then
    'ucDocumento.ini "Select c.descripcion from pedidos_clientes p inner join clientes c on p.cliente=c.codigo where p.activo=1 and p.numero=###", "Select p.Numero,c.Descripcion from pedidos_clientes p inner join clientes c on p.cliente=c.codigo where p.activo=1", False
    ucDocumento.ini "Select c.descripcion from pedidos_clientes p inner join clientes c on p.cliente=c.codigo where p.activo=1 and p.numero=###", "Select p.Numero,c.Descripcion,'' as despacho from pedidos_clientes p inner join clientes c on p.cliente=c.codigo where p.numero not in (Select p.Numero " & _
        " from pedidos_clientes p inner join clientes c on p.cliente=c.codigo " & _
        " inner join despachodetalle d on d.numero=p.numero and d.documento='Pedido de Clientes'" & _
        " where p.activo=1)" & _
        " Union " & _
        " Select p.Numero,c.Descripcion,numerodespacho as despacho " & _
        " from pedidos_clientes p " & _
        " inner join clientes c on p.cliente=c.codigo " & _
        " inner join despachodetalle d on d.numero=p.numero and d.documento='Pedido de Clientes' " & _
        " Where p.activo = 1" & _
        " order by p.numero desc", False
ElseIf Op2.Value = True Then
    'ucDocumento.ini "Select c.descripcion from remitoventa r inner join clientes c on r.cliente=c.codigo where r.anulado=0 and r.numero=###", "Select r.Numero,c.Descripcion from remitoventa r inner join clientes c on r.cliente=c.codigo where r.anulado=0", False
    ucDocumento.ini "Select c.descripcion from remitoventa r inner join clientes c on r.cliente=c.codigo where r.anulado=0 and r.numero=###", "Select r.Numero,c.Descripcion,'' as despacho " & _
        " from remitoventa r " & _
        " inner join clientes c on r.cliente=c.codigo " & _
        " where r.numero not in (Select r.Numero from remitoventa r " & _
        " inner join clientes c on r.cliente=c.codigo " & _
        " inner join despachodetalle d on d.numero=r.numero and d.documento='Remito Venta' " & _
        " where r.anulado=0) " & _
        " Union " & _
        " Select r.Numero,c.Descripcion,numerodespacho as despacho " & _
        " from remitoventa r " & _
        " inner join clientes c on r.cliente=c.codigo " & _
        " inner join despachodetalle d on d.numero=r.numero and d.documento='Remito Venta' " & _
        " Where r.anulado = 0 " & _
        " order by r.numero desc", False
    
ElseIf Op3.Value = True Then
    'ucDocumento.ini "Select c.descripcion from facturaventa f inner join clientes c on f.cliente=c.codigo where f.activo=1 and f.tipodoc like 'FA%' and f.nrofactura=###", "Select f.NroFactura,f.TipoDoc as Doc,c.Descripcion as [Nombre de Cliente] from facturaventa f inner join clientes c on f.cliente=c.codigo where f.activo=1 and f.tipodoc like 'FA%'", False
    ucDocumento.ini "Select c.descripcion from facturaventa f inner join clientes c on f.cliente=c.codigo where f.activo=1 and f.tipodoc like 'FA%' and f.nrofactura=###", "Select f.NroFactura,f.TipoDoc as Doc,c.Descripcion as [Nombre de Cliente],'' as despacho " & _
        " from facturaventa f " & _
        " inner join clientes c on f.cliente=c.codigo " & _
        " where f.activo=1 and f.tipodoc like 'FA%' and f.nrofactura not in (Select f.NroFactura from facturaventa f " & _
        " inner join clientes c on f.cliente=c.codigo " & _
        " inner join despachodetalle d on d.numero=f.nrofactura and d.documento like 'Factura Venta%' " & _
        " where f.activo=1 and f.tipodoc like 'FA%') " & _
        " Union " & _
        " Select f.NroFactura,f.TipoDoc as Doc,c.Descripcion as [Nombre de Cliente],numerodespacho as despacho " & _
        " from facturaventa f " & _
        " inner join clientes c on f.cliente=c.codigo " & _
        " inner join despachodetalle d on d.numero=f.nrofactura and d.documento like 'Factura Venta%' " & _
        " where f.activo=1 and f.tipodoc like 'FA%'" & _
        " order by f.nrofactura desc", False
    
End If
ucDocumento.codigo = 0
End Function

Private Function dNuevo() As Long
    dNuevo = nSinNull(obtenerDeSQL("select max(codigodespacho) from despacho")) + 1
End Function
Private Function dNuevod() As Long
    dNuevod = nSinNull(obtenerDeSQL("select max(codigodetalle) from despachodetalle")) + 1
End Function

Private Sub Op1_Click()
    qCargo
End Sub

Private Sub Op2_Click()
    qCargo
End Sub

Private Sub Op3_Click()
    qCargo
End Sub

Private Function ABMDespacho(dOpe As String, dCodigo As Long, dNumero As String, dobs As String, dfecha As Date, dobs2 As String) As Boolean
On Error GoTo dmal:
Dim iudd As String
ABMDespacho = True
Select Case dOpe
    Case "A":
        iudd = "INSERT INTO DESPACHO (CODIGODESPACHO,NUMERODESPACHO,OBSERVACION1,observacion2,FECHA,ACTIVO) " _
             & " VALUES (" & dCodigo & "," & sstexto(dNumero) & "," & sstexto(dobs) & "," & sstexto(dobs2) & "," & ssFecha(dfecha) & ",1)"
        DataEnvironment1.Sistema.Execute iudd
    Case "M":
        iudd = "UPDATE DESPACHO SET " _
             & " NUMERODESPACHO=" & sstexto(dNumero) & ",OBSERVACION1=" & sstexto(dobs) & ",observacion2=" & sstexto(dobs2) & ",Fecha=" & ssFecha(dfecha) _
             & " WHERE CODIGODESPACHO=" & dCodigo
        DataEnvironment1.Sistema.Execute iudd
    Case "B":
        iudd = "UPDATE DESPACHO SET ACTIVO=0, " _
            & " FECHA=" & ssFecha(dfecha) _
            & " WHERE CODIGODESPACHO=" & dCodigo
        DataEnvironment1.Sistema.Execute iudd
End Select
Exit Function
dmal:
ABMDespacho = False
End Function

Private Function ABMDDetalle(dOpe As String, dCodigo As Long, dNumeroDes As String, dDocumento As String, dNumero As String) As Boolean
On Error GoTo ddmal:
Dim iudd As String
ABMDDetalle = True
Select Case dOpe
    Case "A":
        iudd = "INSERT INTO DESPACHODETALLE (CODIGODETALLE,NUMERODESPACHO,DOCUMENTO,NUMERO) " _
             & " VALUES (" & dCodigo & "," & sstexto(dNumeroDes) & "," & sstexto(dDocumento) & "," & sstexto(dNumero) & ")"
        DataEnvironment1.Sistema.Execute iudd
    Case "M": 'NO HAY
    Case "B":
        iudd = "DELETE FROM DESPACHODETALLE " _
            & " WHERE NUMERODESPACHO=" & sstexto(dNumeroDes)
        DataEnvironment1.Sistema.Execute iudd
End Select
Exit Function
ddmal:
ABMDDetalle = False
End Function

Private Function dHay() As Boolean
If gDetalle.rows > 1 Then
    dHay = True
Else
    dHay = False
End If
End Function

Private Sub txtNumero_LostFocus()
Dim e
    e = obtenerDeSQL("select fecha from despacho where numerodespacho='" & txtNumero & "' and activo=1")
    If IsNull(e) Or IsEmpty(e) Then
    Else
        MsgBox "Este numero de despacho ya existe con fecha : " & e, vbExclamation
        txtNumero.SetFocus
    End If
End Sub

Private Sub ucBotonera_AceptarAlta()
If dHay = False Then Exit Sub
DE_BeginTrans
    If GrabaDespacho("A") = True Then
        DE_CommitTrans
        MsgBox "Despacho guardado.", vbInformation
        cleard
        ucBotonera.AceptarOk
    Else
        DE_RollbackTrans
    End If
End Sub

Private Sub ucBotonera_AceptarModi()
If dHay = False Then Exit Sub
DE_BeginTrans
    If GrabaDespacho("M") = True Then
        DE_CommitTrans
        MsgBox "Despacho modificado.", vbInformation
        cleard
        ucBotonera.AceptarOk
    Else
        DE_RollbackTrans
    End If
End Sub

Private Function GrabaDespacho(gOPE As String) As Boolean
GrabaDespacho = True
If txtNumero = "" Then
    MsgBox "ingrese Numero de Despacho", vbCritical
    txtNumero.SetFocus
    GrabaDespacho = False
    Exit Function
End If
Select Case gOPE
    Case "A":
        If ABMDespacho(gOPE, s2n(txtCodigo), txtNumero, txtObs, dtfecha, Text1.Text) = False Then GoTo gdmal
    Case "M":
        If ABMDespacho(gOPE, s2n(txtCodigo), txtNumero, txtObs, dtfecha, Text1.Text) = False Then GoTo gdmal
End Select
If GrabaDetalle = False Then GoTo gdmal
Exit Function
gdmal:
GrabaDespacho = False
End Function

Private Function GrabaDetalle() As Boolean
Dim i As Long
GrabaDetalle = True
    If ABMDDetalle("B", 0, txtNumero, "", "") = False Then GoTo demal
    For i = 1 To gDetalle.rows - 1
        If ABMDDetalle("A", dNuevod, txtNumero, gDetalle.TextMatrix(i, 2), gDetalle.TextMatrix(i, 3)) = False Then GoTo demal
    Next
Exit Function
demal:
GrabaDetalle = False
End Function


Private Sub ucBotonera_Buscar()
Dim r
    r = frmBuscar.MostrarSql("Select codigodespacho as Int, numerodespacho as Despacho,Fecha, Observacion1,observacion2 from despacho where activo=1 order by codigodespacho desc")
    If r > "" Then
        txtCodigo = r
        txtNumero = frmBuscar.resultado(2)
        txtObs = frmBuscar.resultado(4)
        Text1.Text = frmBuscar.resultado(5)
        LlenarGrilla gDetalle, "Select codigodetalle, numerodespacho,Documento,Numero from despachodetalle where numerodespacho=" & sstexto(txtNumero) & " order by codigodetalle", True
        ucBotonera.BuscarOK
    End If
End Sub

Private Sub ucBotonera_Cancelar()
    cleard
End Sub

Private Sub ucBotonera_Eliminar()
    If MsgBox("¿Desea eliminar el Despacho?", vbYesNo) = vbYes Then
        If ABMDespacho("B", s2n(txtCodigo), "", "", Date, "") = False Then GoTo errdel
        If ABMDDetalle("B", 0, txtNumero, "", "") = False Then GoTo errdel
        MsgBox "Despacho eliminado.", vbInformation
        cleard
        ucBotonera.EliminarOK
    End If
Exit Sub
errdel:
    MsgBox "Despacho no eliminado.", vbCritical
End Sub

Private Sub ucBotonera_Nuevo()
    cleard
    txtCodigo = dNuevo
End Sub

Private Sub ucBotonera_SALIR()
    Unload Me
End Sub

Private Sub ucXls1_Clic(Cancel As Boolean)
    ucXls1.ini gDetalle, "C:\Despacho_" & txtNumero & ".xls"
End Sub

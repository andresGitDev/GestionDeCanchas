VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmLisComprasDetalle 
   Caption         =   "Listado de Detalle de Comprobantes de Proveedores"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6435
   Icon            =   "FrmLisComprasDetalle.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   6435
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
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
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
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
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   975
   End
   Begin VB.OptionButton optmes 
      Alignment       =   1  'Right Justify
      Caption         =   "Por Mes de Imputacion"
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
      Height          =   240
      Left            =   480
      TabIndex        =   12
      Top             =   240
      Width           =   2415
   End
   Begin VB.OptionButton optfecha 
      Alignment       =   1  'Right Justify
      Caption         =   "Entre Fechas"
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
      Left            =   600
      TabIndex        =   11
      Top             =   1440
      Width           =   1695
   End
   Begin VB.ComboBox cmbmes 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "FrmLisComprasDetalle.frx":08CA
      Left            =   1320
      List            =   "FrmLisComprasDetalle.frx":08F2
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   600
      Width           =   1935
   End
   Begin VB.ComboBox cmbaño 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "FrmLisComprasDetalle.frx":095B
      Left            =   4560
      List            =   "FrmLisComprasDetalle.frx":0974
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   600
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker dtfechad 
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   81068033
      CurrentDate     =   38252
   End
   Begin MSComCtl2.DTPicker dtfechah 
      Height          =   255
      Left            =   4560
      TabIndex        =   6
      Top             =   1800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   81068033
      CurrentDate     =   38252
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400000&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   2445
      Left            =   120
      Top             =   120
      Width           =   6240
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00400000&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   765
      Left            =   360
      Top             =   360
      Width           =   5760
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00400000&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Height          =   765
      Left            =   360
      Top             =   1560
      Width           =   5760
   End
   Begin VB.Label Label1 
      Caption         =   "Mes"
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
      Left            =   720
      TabIndex        =   10
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha Desde"
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
      Left            =   480
      TabIndex        =   9
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha Hasta"
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
      Left            =   3360
      TabIndex        =   8
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Año"
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
      Left            =   4080
      TabIndex        =   7
      Top             =   600
      Width           =   495
   End
End
Attribute VB_Name = "FrmLisComprasDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '5/10/4

Private msFCtmp As String



Private Sub cmdaceptar_Click()

    Dim STR As String
    Dim rs As New ADODB.Recordset
    Dim rstrans As New ADODB.Recordset
    Dim Total As Variant
    
    Dim fvIva21 As Variant
    Dim fvIva27 As Variant
    Dim fvIva10 As Variant
    Dim fvretgan As Variant
    Dim fvImp As Variant
    Dim fvex As Variant
    Dim fvper As Variant
    Dim fvtot As Variant
    Dim fvneto As Variant
    Dim fvrg3431 As Variant
    Dim fvretiva As Variant
    Dim fvperib As Variant
    
    Dim cvIva21 As Variant
    Dim cvIva27 As Variant
    Dim cvIva10 As Variant
    Dim cvretgan As Variant
    Dim cvImp As Variant
    Dim cvex As Variant
    Dim cvper As Variant
    Dim cvtot As Variant
    Dim cvneto As Variant
    Dim cvrg3431 As Variant
    Dim cvretiva As Variant
    Dim cvperib As Variant
    
    Dim dvIva21 As Variant
    Dim dvIva27 As Variant
    Dim dvIva10 As Variant
    Dim dvretgan As Variant
    Dim dvImp As Variant
    Dim dvex As Variant
    Dim dvper As Variant
    Dim dvtot As Variant
    Dim dvneto As Variant
    Dim dvrg3431 As Variant
    Dim dvretiva As Variant
    Dim dvperib As Variant
    
    fvIva21 = 0
    fvIva27 = 0
    fvIva10 = 0
    fvretgan = 0
    fvImp = 0
    fvex = 0
    fvper = 0
    fvtot = 0
    fvneto = 0
    fvrg3431 = 0
    fvretiva = 0
    fvperib = 0
    dvIva21 = 0
    dvIva27 = 0
    dvIva10 = 0
    dvretgan = 0
    dvImp = 0
    dvex = 0
    dvper = 0
    dvtot = 0
    dvneto = 0
    dvrg3431 = 0
    dvretiva = 0
    dvperib = 0
    cvIva21 = 0
    cvIva27 = 0
    cvIva10 = 0
    cvretgan = 0
    cvImp = 0
    cvex = 0
    cvper = 0
    cvtot = 0
    cvneto = 0
    cvrg3431 = 0
    cvretiva = 0
    cvperib = 0
    
    'tabla temp
    'If msFCtmp = "" Then
    msFCtmp = TablaTempCrear(tt_FacturaCompra_temp)
    


    'daTaenvironment1.Sistema.Execute "delete from " & msFCtmp 'Totalfacturascomprastemp"
    If optFecha.Value = True Then
        
        With rstrans
            .Open "select tipodoc,Sum(neto) as netos, Sum(iva_21) as iva, sum(percepc) as per, sum(iva_27) as iva27, sum(iva_10) as iva10," _
                        & "sum(Iva_9) as retiva,sum(imp_int) as imp,sum(total) as tot,sum(retganpago) as RetGanPago,sum(der_est)as rg3431,sum(exento)as ex" _
                        & ",sum(ibcapital) as perib from transcom  where fecha " & ssBetween(dtfechad, dtfechah) & " and (tipodoc='FAC' or tipodoc='N/D' or tipodoc='N/C') and activo=1 group by tipodoc", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
                        
            
            If Not .EOF Then
                Do While Not .EOF
                    Select Case !TIPODOC
                        Case "FAC"
                            fvIva21 = !Iva
                            fvIva27 = !iva27
                            fvIva10 = !iva10
                            fvretgan = !retganpago
                            fvImp = !imp
                            fvex = !ex
                            fvper = !per
                            fvtot = !tot
                            fvneto = !netos
                            fvrg3431 = !rg3431
                            fvretiva = !retIva
                            fvperib = !perIb
    '
                        Case "N/D"
                            dvIva21 = !Iva
                            dvIva27 = !iva27
                            dvIva10 = !iva10
                            dvretgan = !retganpago
                            dvImp = !imp
                            dvex = !ex
                            dvper = !per
                            dvtot = !tot
                            dvneto = !netos
                            dvrg3431 = !rg3431
                            dvretiva = !retIva
                            dvperib = !perIb
    
                        Case "N/C"
                            cvIva21 = -!Iva
                            cvIva27 = -!iva27
                            cvIva10 = -!iva10
                            cvretgan = -!retganpago
                            cvImp = -!imp
                            cvex = -!ex
                            cvper = -!per
                            cvtot = -!tot
                            cvneto = -!netos
                            cvrg3431 = -!rg3431
                            cvretiva = -!retIva
                            cvperib = -!perIb
                    End Select
                    .MoveNext
                Loop
            End If
        End With
        
        
        With rs
            .Open "select tipodoc,sum(neto) as netos,sum(iva_21) as iva21,sum(percepc) as per,sum(iva_27) as iva27,sum(iva_10) as iva10," _
                        & "sum(Iva_9) as retiva,sum(imp_int) as imp,sum(RetGanPago) as RetGanPago,sum(total) as tot,sum(der_est)as rg3431,sum(exento)as ex" _
                        & ",sum(ibcapital) as perib from compras  where fecha " & ssBetween(dtfechad, dtfechah) & " and (tipodoc='FAC' or tipodoc='N/D' or tipodoc='N/C') and activo=1 group by tipodoc", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
                        
            If Not rs.EOF Then
                Do While Not rs.EOF
                    Select Case rs!TIPODOC
                        Case "FAC"
                            fvIva21 = fvIva21 + !iva21
                            fvIva27 = fvIva27 + !iva27
                            fvIva10 = fvIva10 + !iva10
                            fvretgan = fvretgan + !retganpago
                            fvImp = fvImp + !imp
                            fvex = fvex + !ex
                            fvper = fvper + !per
                            fvtot = fvtot + !tot
                            fvneto = fvneto + !netos
                            fvrg3431 = fvrg3431 + !rg3431
                            fvretiva = fvretiva + !retIva
                            fvperib = fvperib + !perIb
    
                        Case "N/D"
                            dvIva21 = dvIva21 + !iva21
                            dvIva27 = dvIva27 + !iva27
                            dvIva10 = dvIva10 + !iva10
                            dvretgan = dvretgan + !retganpago
                            dvImp = dvImp + !imp
                            dvex = dvex + !ex
                            dvper = dvper + !per
                            dvtot = dvtot + !tot
                            dvneto = dvneto + !netos
                            dvrg3431 = dvrg3431 + !rg3431
                            dvretiva = dvretiva + !retIva
                            dvperib = dvperib + !perIb
    
                        Case "N/C"
                            cvIva21 = cvIva21 - !iva21
                            cvIva27 = cvIva27 - !iva27
                            cvIva10 = cvIva10 - !iva10
                            cvretgan = cvretgan - !retganpago
                            cvImp = cvImp - !imp
                            cvex = cvex - !ex
                            cvper = cvper - !per
                            cvtot = cvtot - !tot
                            cvneto = cvneto - !netos
                            cvrg3431 = cvrg3431 - !rg3431
                            cvretiva = cvretiva - !retIva
                            cvperib = cvperib - !perIb
    
                    End Select
                    .MoveNext
                Loop
            End If
            .Close
        End With

        
        Set rstrans = Nothing
        Set rs = Nothing
        
        DataEnvironment1.Sistema.Execute "Insert into " & msFCtmp & " (descripcion,total,neto,iva21,rg3337,iva27,iva10,impint,RetGanPago,retiva,rg3431,perib,exento)" _
        & "values('Facturas'," & Replace(fvtot, ",", ".") & "," & Replace(fvneto, ",", ".") & "," & Replace(fvIva21, ",", ".") & "," & Replace(fvper, ",", ".") _
        & "," & Replace(fvIva27, ",", ".") & "," & Replace(fvIva10, ",", ".") & "," & Replace(fvImp, ",", ".") _
         & "," & Replace(fvretgan, ",", ".") & "," & Replace(fvretiva, ",", ".") & "," & Replace(fvrg3431, ",", ".") & "," & Replace(fvperib, ",", ".") & "," & Replace(fvex, ",", ".") & ")"
                 
        DataEnvironment1.Sistema.Execute "Insert into " & msFCtmp & " (descripcion,total,neto,iva21,rg3337,iva27,iva10,impint,RetGanPago,retiva,rg3431,perib,exento)" _
        & "values('Notas de Debito'," & Replace(dvtot, ",", ".") & "," & Replace(dvneto, ",", ".") & "," & Replace(dvIva21, ",", ".") & "," & Replace(dvper, ",", ".") _
        & "," & Replace(dvIva27, ",", ".") & "," & Replace(dvIva10, ",", ".") & "," & Replace(dvImp, ",", ".") _
         & "," & Replace(dvretgan, ",", ".") & "," & Replace(dvretiva, ",", ".") & "," & Replace(dvrg3431, ",", ".") & "," & Replace(dvperib, ",", ".") & "," & Replace(dvex, ",", ".") & ")"
        
        DataEnvironment1.Sistema.Execute "Insert into " & msFCtmp & " (descripcion,total,neto,iva21,rg3337,iva27,iva10,impint,RetGanPago,retiva,rg3431,perib,exento)" _
        & "values('Notas de Credito'," & Replace(cvtot, ",", ".") & "," & Replace(cvneto, ",", ".") & "," & Replace(cvIva21, ",", ".") & "," & Replace(cvper, ",", ".") _
        & "," & Replace(cvIva27, ",", ".") & "," & Replace(cvIva10, ",", ".") & "," & Replace(cvImp, ",", ".") _
         & "," & Replace(cvretgan, ",", ".") & "," & Replace(cvretiva, ",", ".") & "," & Replace(cvrg3431, ",", ".") & "," & Replace(cvperib, ",", ".") & "," & Replace(cvex, ",", ".") & ")"
        
        
        RptLisDetalleComprobantesProveedores.DataControl1.Connection = DataEnvironment1.Sistema
        STR = "Select * from " & msFCtmp & " "
        RptLisDetalleComprobantesProveedores.lbltotiva.caption = Format(s2n(fvIva10) + s2n(cvIva10) + s2n(dvIva10) + s2n(fvIva27) + s2n(cvIva27) + s2n(dvIva27) + s2n(fvIva21) + s2n(cvIva21) + s2n(dvIva21), "#,###,##0.###0")
        RptLisDetalleComprobantesProveedores.DataControl1.Source = STR
        
        ' Sacar las compras contado del exterior
        rs.Open "select sum(total) as tot from compras" _
        & " where fecha " & ssBetween(dtfechad, dtfechah) & " and (tipodoc='FAC' or tipodoc='N/D')  and (exterior = 1) and (contado=1) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        Total = 0
        If Not rs.EOF Then
            Total = rs!tot
        End If
        rs.Close
        Set rs = Nothing
        rs.Open "select sum(total) as tot from compras" _
        & " where fecha " & ssBetween(dtfechad, dtfechah) & " and (tipodoc='N/C')  and (exterior= 1) and (contado=1) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            Total = Total - rs!tot
        End If
        rs.Close
        Set rs = Nothing
        RptLisDetalleComprobantesProveedores.lblcontext.caption = Format(Total, "#,###,##0.###0")
        
        ' Sacar las compras contado locales
        rs.Open "select sum(total) as tot from compras" _
        & " where fecha " & ssBetween(dtfechad, dtfechah) & " and (tipodoc='FAC' or tipodoc='N/D')  and (exterior = 0) and (contado=1) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        Total = 0
        If Not rs.EOF Then
            Total = rs!tot
        End If
        rs.Close
        Set rs = Nothing
        rs.Open "select sum(total) as tot from compras" _
        & " where fecha " & ssBetween(dtfechad, dtfechah) & " and (tipodoc='N/C')  and (exterior=0) and (contado=1) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            Total = Total - rs!tot
        End If
        rs.Close
        Set rs = Nothing
        RptLisDetalleComprobantesProveedores.lblcontloc.caption = Format(Total, "#,###,##0.###0")
        
        ' Sacar las compras cta cte del exterior
        rs.Open "select sum(total) as tot from compras" _
        & " where fecha " & ssBetween(dtfechad, dtfechah) & " and (tipodoc='FAC' or tipodoc='N/D')  and (exterior= 1) and (contado=0) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        Total = 0
        If Not rs.EOF Then
            Total = rs!tot
        End If
        rs.Close
        Set rs = Nothing
        rs.Open "select sum(total) as tot from compras" _
        & " where fecha " & ssBetween(dtfechad, dtfechah) & " and (tipodoc='N/C')  and (exterior = 1) and (contado=0) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            Total = Total - rs!tot
        End If
        rs.Close
        Set rs = Nothing
        '**********************
        rs.Open "select sum(total) as tot from transcom" _
        & " where fecha " & ssBetween(dtfechad, dtfechah) & " and (tipodoc='FAC' or tipodoc='N/D')  and (exterior= 1) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            Total = Total + rs!tot
        End If
        rs.Close
        Set rs = Nothing
        rs.Open "select sum(total) as tot from transcom" _
        & " where fecha " & ssBetween(dtfechad, dtfechah) & " and (tipodoc='N/C')  and (exterior = 1) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            Total = Total - rs!tot
        End If
        rs.Close
        Set rs = Nothing
        RptLisDetalleComprobantesProveedores.lblctaext.caption = Format(Total, "#,###,##0.###0")
        
        ' Sacar las compras cta cte locales
        rs.Open "select sum(total) as tot from compras" _
        & " where fecha " & ssBetween(dtfechad, dtfechah) & " and (tipodoc='FAC' or tipodoc='N/D')  and (exterior= 0) and (contado=0) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        Total = 0
        If Not rs.EOF Then
            Total = rs!tot
        End If
        rs.Close
        Set rs = Nothing
        rs.Open "select sum(total) as tot from compras" _
        & " where fecha " & ssBetween(dtfechad, dtfechah) & " and (tipodoc='N/C')  and (exterior = 0) and (contado=0) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            Total = Total - rs!tot
        End If
        rs.Close
        Set rs = Nothing
        '**********************
        rs.Open "select sum(total) as tot from transcom" _
        & " where fecha " & ssBetween(dtfechad, dtfechah) & " and (tipodoc='FAC' or tipodoc='N/D')  and (exterior= 0) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            Total = Total + rs!tot
        End If
        rs.Close
        Set rs = Nothing
        rs.Open "select sum(total) as tot from transcom" _
        & " where fecha " & ssBetween(dtfechad, dtfechah) & " and (tipodoc='N/C')  and (exterior = 0) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            Total = Total - rs!tot
        End If
        rs.Close
        Set rs = Nothing
        RptLisDetalleComprobantesProveedores.lblctaloc.caption = Format(Total, "#,###,##0.###0")
        
        ' Sacar las compras contado
        rs.Open "select sum(total) as tot from compras" _
        & " where fecha " & ssBetween(dtfechad, dtfechah) & " and (tipodoc='FAC' or tipodoc='N/D') and (contado=1) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        Total = 0
        If Not rs.EOF Then
            Total = rs!tot
        End If
        rs.Close
        Set rs = Nothing
        rs.Open "select sum(total) as tot from compras" _
        & " where fecha " & ssBetween(dtfechad, dtfechah) & " and (tipodoc='N/C') and (contado=1) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            Total = Total - rs!tot
        End If
        rs.Close
        Set rs = Nothing
        RptLisDetalleComprobantesProveedores.lblcontado.caption = Format(Total, "#,###,##0.###0")
        
        ' Sacar las Compras Cta Cte
        rs.Open "select sum(total) as tot from compras" _
        & " where fecha " & ssBetween(dtfechad, dtfechah) & " and (tipodoc='FAC' or tipodoc='N/D') and (contado=0) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        Total = 0
        If Not rs.EOF Then
            Total = rs!tot
        End If
        rs.Close
        Set rs = Nothing
        rs.Open "select sum(total) as tot from compras" _
        & " where fecha " & ssBetween(dtfechad, dtfechah) & " and (tipodoc='N/C') and (contado=0) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            Total = Total - rs!tot
        End If
        rs.Close
        Set rs = Nothing
        '**********************
        rs.Open "select sum(total) as tot from transcom" _
        & " where fecha " & ssBetween(dtfechad, dtfechah) & " and (tipodoc='FAC' or tipodoc='N/D')  and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            Total = Total + rs!tot
        End If
        rs.Close
        Set rs = Nothing
        rs.Open "select sum(total) as tot from transcom" _
        & " where fecha " & ssBetween(dtfechad, dtfechah) & " and (tipodoc='N/C') and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            Total = Total - rs!tot
        End If
        rs.Close
        Set rs = Nothing
        RptLisDetalleComprobantesProveedores.lblcta.caption = Format(Total, "#,###,##0.###0")
        
        ' Sacar las Compras Locales
        rs.Open "select sum(total) as tot from compras" _
        & " where fecha " & ssBetween(dtfechad, dtfechah) & " and (tipodoc='FAC' or tipodoc='N/D')  and (exterior= 0) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        Total = 0
        If Not rs.EOF Then
            Total = rs!tot
        End If
        rs.Close
        Set rs = Nothing
        rs.Open "select sum(total) as tot from compras" _
        & " where fecha " & ssBetween(dtfechad, dtfechah) & " and (tipodoc='N/C')  and (exterior = 0) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            Total = Total - rs!tot
        End If
        rs.Close
        Set rs = Nothing
        '**********************
        rs.Open "select sum(total) as tot from transcom" _
        & " where fecha " & ssBetween(dtfechad, dtfechah) & " and (tipodoc='FAC' or tipodoc='N/D') and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            Total = Total + rs!tot
        End If
        rs.Close
        Set rs = Nothing
        rs.Open "select sum(total) as tot from transcom" _
        & " where fecha " & ssBetween(dtfechad, dtfechah) & " and (tipodoc='N/C') and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            Total = Total - rs!tot
        End If
        rs.Close
        Set rs = Nothing
        RptLisDetalleComprobantesProveedores.lblloc.caption = Format(Total, "#,###,##0.###0")
        
        'Sacar las Compras Exterior
        rs.Open "select sum(total) as tot from compras" _
        & " where fecha " & ssBetween(dtfechad, dtfechah) & " and (tipodoc='FAC' or tipodoc='N/D')  and (exterior= 1) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        Total = 0
        If Not rs.EOF Then
            Total = rs!tot
        End If
        rs.Close
        Set rs = Nothing
        rs.Open "select sum(total) as tot from compras" _
        & " where fecha " & ssBetween(dtfechad, dtfechah) & " and (tipodoc='N/C')  and (exterior = 1) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            Total = Total - rs!tot
        End If
        rs.Close
        Set rs = Nothing
        '**********************
        rs.Open "select sum(total) as tot from transcom" _
        & " where fecha " & ssBetween(dtfechad, dtfechah) & " and (tipodoc='FAC' or tipodoc='N/D')  and (exterior= 1) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            Total = Total + rs!tot
        End If
        rs.Close
        Set rs = Nothing
        rs.Open "select sum(total) as tot from transcom" _
        & " where fecha " & ssBetween(dtfechad, dtfechah) & " and (tipodoc='N/C')  and (exterior = 1) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            Total = Total - rs!tot
        End If
        rs.Close
        Set rs = Nothing
        RptLisDetalleComprobantesProveedores.lblext.caption = Format(Total, "#,###,##0.###0")
        
        'Sacar Compras
        rs.Open "select sum(total) as tot from compras" _
        & " where fecha " & ssBetween(dtfechad, dtfechah) & " and (tipodoc='FAC' or tipodoc='N/D') and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        Total = 0
        If Not rs.EOF Then
            Total = rs!tot
        End If
        rs.Close
        Set rs = Nothing
        rs.Open "select sum(total) as tot from compras" _
        & " where fecha " & ssBetween(dtfechad, dtfechah) & " and (tipodoc='N/C') and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            Total = Total - rs!tot
        End If
        rs.Close
        Set rs = Nothing
        '**********************
        rs.Open "select sum(total) as tot from transcom" _
        & " where fecha " & ssBetween(dtfechad, dtfechah) & " and (tipodoc='FAC' or tipodoc='N/D') and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            Total = Total + rs!tot
        End If
        rs.Close
        Set rs = Nothing
        rs.Open "select sum(total) as tot from transcom" _
        & " where fecha " & ssBetween(dtfechad, dtfechah) & " and (tipodoc='N/C') and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            Total = Total - rs!tot
        End If
        rs.Close
        Set rs = Nothing
        RptLisDetalleComprobantesProveedores.lblcomp.caption = Format(Total, "#,###,##0.###0")
        RptLisDetalleComprobantesProveedores.lblfecha = Date
        RptLisDetalleComprobantesProveedores.lblTitulo.caption = "Listado de Deatalles de comprobantes Proveedores desde el " & CStr(dtfechad) & " al " & CStr(dtfechah)
        RptLisDetalleComprobantesProveedores.PageSettings.Orientation = ddOLandscape
        RptLisDetalleComprobantesProveedores.Show
    Else
        If optmes.Value = True Then
            rstrans.Open "select tipodoc,Sum(neto) as netos, Sum(iva_21) as iva, sum(percepc) as per, sum(iva_27) as iva27, sum(iva_10) as iva10," _
            & "sum(Iva_9) as retiva,sum(imp_int) as imp,sum(total) as tot,sum(RetGanPago) as RetGanPago,sum(der_est)as rg3431,sum(exento)as ex" _
            & ",sum(ibcapital) as perib from transcom  where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='FAC' or tipodoc='N/D' or tipodoc='N/C') and activo=1 group by tipodoc", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            If Not rstrans.EOF Then
                Do While Not rstrans.EOF
                    Select Case rstrans!TIPODOC
                        Case "FAC"
                            fvIva21 = rstrans!Iva
                            fvIva27 = rstrans!iva27
                            fvIva10 = rstrans!iva10
                            fvretgan = rstrans!retganpago
                            fvImp = rstrans!imp
                            fvex = rstrans!ex
                            fvper = rstrans!per
                            fvtot = rstrans!tot
                            fvneto = rstrans!netos
                            fvrg3431 = rstrans!rg3431
                            fvretiva = rstrans!retIva
                            fvperib = rstrans!perIb
    '
                        Case "N/D"
                            dvIva21 = rstrans!Iva
                            dvIva27 = rstrans!iva27
                            dvIva10 = rstrans!iva10
                            dvretgan = rstrans!retganpago
                            dvImp = rstrans!imp
                            dvex = rstrans!ex
                            dvper = rstrans!per
                            dvtot = rstrans!tot
                            dvneto = rstrans!netos
                            dvrg3431 = rstrans!rg3431
                            dvretiva = rstrans!retIva
                            dvperib = rstrans!perIb
    
                        Case "N/C"
                            cvIva21 = -rstrans!Iva
                            cvIva27 = -rstrans!iva27
                            cvIva10 = -rstrans!iva10
                            cvretgan = -rstrans!retganpago
                            cvImp = -rstrans!imp
                            cvex = -rstrans!ex
                            cvper = -rstrans!per
                            cvtot = -rstrans!tot
                            cvneto = -rstrans!netos
                            cvrg3431 = -rstrans!rg3431
                            cvretiva = -rstrans!retIva
                            cvperib = -rstrans!perIb
                    End Select
                    rstrans.MoveNext
                Loop
            End If
            rs.Open "select tipodoc,sum(neto) as netos,sum(iva_21) as iva21,sum(percepc) as per,sum(iva_27) as iva27,sum(iva_10) as iva10," _
            & "sum(Iva_9) as retiva,sum(imp_int) as imp,sum(RetGanPago) as RetGanPago,sum(total) as tot,sum(der_est)as rg3431,sum(exento)as ex" _
            & ",sum(ibcapital) as perib from compras  where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='FAC' or tipodoc='N/D' or tipodoc='N/C') and activo=1 group by tipodoc", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                Do While Not rs.EOF
                    Select Case rs!TIPODOC
                        Case "FAC"
                            fvIva21 = fvIva21 + rs!iva21
                            fvIva27 = fvIva27 + rs!iva27
                            fvIva10 = fvIva10 + rs!iva10
                            fvretgan = fvretgan + rs!retganpago
                            fvImp = fvImp + rs!imp
                            fvex = fvex + rs!ex
                            fvper = fvper + rs!per
                            fvtot = fvtot + rs!tot
                            fvneto = fvneto + rs!netos
                            fvrg3431 = fvrg3431 + rs!rg3431
                            fvretiva = fvretiva + rs!retIva
                            fvperib = fvperib + rs!perIb
    
                        Case "N/D"
                            dvIva21 = dvIva21 + rs!iva21
                            dvIva27 = dvIva27 + rs!iva27
                            dvIva10 = dvIva10 + rs!iva10
                            dvretgan = dvretgan + rs!retganpago
                            dvImp = dvImp + rs!imp
                            dvex = dvex + rs!ex
                            dvper = dvper + rs!per
                            dvtot = dvtot + rs!tot
                            dvneto = dvneto + rs!netos
                            dvrg3431 = dvrg3431 + rs!rg3431
                            dvretiva = dvretiva + rs!retIva
                            dvperib = dvperib + rs!perIb
    
                        Case "N/C"
                            cvIva21 = cvIva21 - rs!iva21
                            cvIva27 = cvIva27 - rs!iva27
                            cvIva10 = cvIva10 - rs!iva10
                            cvretgan = cvretgan - rs!retganpago
                            cvImp = cvImp - rs!imp
                            cvex = cvex - rs!ex
                            cvper = cvper - rs!per
                            cvtot = cvtot - rs!tot
                            cvneto = cvneto - rs!netos
                            cvrg3431 = cvrg3431 - rs!rg3431
                            cvretiva = cvretiva - rs!retIva
                            cvperib = cvperib - rs!perIb
    
                    End Select
                    rs.MoveNext
                Loop
            End If
            rs.Close
            Set rs = Nothing
            rstrans.Close
            Set rstrans = Nothing
            
            DataEnvironment1.Sistema.Execute "Insert into " & msFCtmp & " (descripcion,total,neto,iva21,rg3337,iva27,iva10,impint,RetGanPago,retiva,rg3431,perib,exento)" _
            & "values('Facturas'," & Replace(fvtot, ",", ".") & "," & Replace(fvneto, ",", ".") & "," & Replace(fvIva21, ",", ".") & "," & Replace(fvper, ",", ".") _
            & "," & Replace(fvIva27, ",", ".") & "," & Replace(fvIva10, ",", ".") & "," & Replace(fvImp, ",", ".") _
             & "," & Replace(nSinNull(fvretgan), ",", ".") & "," & Replace(fvretiva, ",", ".") & "," & Replace(fvrg3431, ",", ".") & "," & Replace(fvperib, ",", ".") & "," & Replace(fvex, ",", ".") & ")"
                     
            DataEnvironment1.Sistema.Execute "Insert into " & msFCtmp & " (descripcion,total,neto,iva21,rg3337,iva27,iva10,impint,RetGanPago,retiva,rg3431,perib,exento)" _
            & "values('Notas de Debito'," & Replace(dvtot, ",", ".") & "," & Replace(dvneto, ",", ".") & "," & Replace(dvIva21, ",", ".") & "," & Replace(dvper, ",", ".") _
            & "," & Replace(dvIva27, ",", ".") & "," & Replace(dvIva10, ",", ".") & "," & Replace(dvImp, ",", ".") _
             & "," & Replace(nSinNull(dvretgan), ",", ".") & "," & Replace(dvretiva, ",", ".") & "," & Replace(dvrg3431, ",", ".") & "," & Replace(dvperib, ",", ".") & "," & Replace(dvex, ",", ".") & ")"
            
            DataEnvironment1.Sistema.Execute "Insert into " & msFCtmp & " (descripcion,total,neto,iva21,rg3337,iva27,iva10,impint,RetGanPago,retiva,rg3431,perib,exento)" _
            & "values('Notas de Credito'," & Replace(cvtot, ",", ".") & "," & Replace(cvneto, ",", ".") & "," & Replace(cvIva21, ",", ".") & "," & Replace(cvper, ",", ".") _
            & "," & Replace(cvIva27, ",", ".") & "," & Replace(cvIva10, ",", ".") & "," & Replace(cvImp, ",", ".") _
             & "," & Replace(nSinNull(cvretgan), ",", ".") & "," & Replace(cvretiva, ",", ".") & "," & Replace(cvrg3431, ",", ".") & "," & Replace(cvperib, ",", ".") & "," & Replace(cvex, ",", ".") & ")"
            
            
            RptLisDetalleComprobantesProveedores.DataControl1.Connection = DataEnvironment1.Sistema
            STR = "Select * from " & msFCtmp & " "
            RptLisDetalleComprobantesProveedores.lbltotiva.caption = Format(s2n(fvIva10) + s2n(cvIva10) + s2n(dvIva10) + s2n(fvIva27) + s2n(cvIva27) + s2n(dvIva27) + s2n(fvIva21) + s2n(cvIva21) + s2n(dvIva21), "#,###,##0.###0")
            RptLisDetalleComprobantesProveedores.DataControl1.Source = STR
            
            ' Sacar las compras contado del exterior
            rs.Open "select sum(total) as tot from compras" _
            & " where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='FAC' or tipodoc='N/D')  and (exterior = 1) and (contado=1) and activo=1 ", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            Total = 0
            If Not rs.EOF Then
                If Not IsNull(rs!tot) Then
                    Total = rs!tot
                End If
            End If
            rs.Close
            Set rs = Nothing
            rs.Open "select sum(total) as tot from compras" _
            & " where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='N/C')  and (exterior= 1) and (contado=1) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                If Not IsNull(rs!tot) Then
                    Total = Total - rs!tot
                End If
            End If
            rs.Close
            Set rs = Nothing
            RptLisDetalleComprobantesProveedores.lblcontext.caption = Format(Total, "#,###,##0.###0")
            
            ' Sacar las compras contado locales
            rs.Open "select sum(total) as tot from compras" _
            & " where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='FAC' or tipodoc='N/D')  and (exterior = 0) and (contado=1) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            Total = 0
            If Not rs.EOF Then
                If Not IsNull(rs!tot) Then
                    Total = rs!tot
                End If
            End If
            rs.Close
            Set rs = Nothing
            rs.Open "select sum(total) as tot from compras" _
            & " where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='N/C')  and (exterior=0) and (contado=1) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                If Not IsNull(rs!tot) Then
                    Total = Total - rs!tot
                End If
            End If
            rs.Close
            Set rs = Nothing
            RptLisDetalleComprobantesProveedores.lblcontloc.caption = Format(Total, "#,###,##0.###0")
            
            ' Sacar las compras cta cte del exterior
            rs.Open "select sum(total) as tot from compras" _
            & " where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='FAC' or tipodoc='N/D')  and (exterior= 1) and (contado=0) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            Total = 0
            If Not rs.EOF Then
                If Not IsNull(rs!tot) Then
                    Total = rs!tot
                End If
            End If
            rs.Close
            Set rs = Nothing
            rs.Open "select sum(total) as tot from compras" _
            & " where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='N/C')  and (exterior = 1) and (contado=0) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                If Not IsNull(rs!tot) Then
                    Total = Total - rs!tot
                End If
            End If
            rs.Close
            Set rs = Nothing
            '**********************
            rs.Open "select sum(total) as tot from transcom" _
            & " where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='FAC' or tipodoc='N/D')  and (exterior= 1) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                If Not IsNull(rs!tot) Then
                    Total = Total + rs!tot
                End If
            End If
            rs.Close
            Set rs = Nothing
            rs.Open "select sum(total) as tot from transcom" _
            & " where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='N/C')  and (exterior = 1) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                If Not IsNull(rs!tot) Then
                    Total = Total - rs!tot
                End If
            End If
            rs.Close
            Set rs = Nothing
            RptLisDetalleComprobantesProveedores.lblctaext.caption = Format(Total, "#,###,##0.###0")
            
            ' Sacar las compras cta cte locales
            rs.Open "select sum(total) as tot from compras" _
            & " where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='FAC' or tipodoc='N/D')  and (exterior= 0) and (contado=0) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            Total = 0
            If Not rs.EOF Then
                If Not IsNull(rs!tot) Then
                    Total = rs!tot
                End If
            End If
            rs.Close
            Set rs = Nothing
            rs.Open "select sum(total) as tot from compras" _
            & " where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='N/C')  and (exterior = 0) and (contado=0) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                If Not IsNull(rs!tot) Then
                    Total = Total - rs!tot
                End If
            End If
            rs.Close
            Set rs = Nothing
            '**********************
            rs.Open "select sum(total) as tot from transcom" _
            & " where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='FAC' or tipodoc='N/D')  and (exterior= 0) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                If Not IsNull(rs!tot) Then
                    Total = Total + rs!tot
                End If
            End If
            rs.Close
            Set rs = Nothing
            rs.Open "select sum(total) as tot from transcom" _
            & " where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='N/C')  and (exterior = 0) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                If Not IsNull(rs!tot) Then
                    Total = Total - rs!tot
                End If
            End If
            rs.Close
            Set rs = Nothing
            RptLisDetalleComprobantesProveedores.lblctaloc.caption = Format(Total, "#,###,##0.###0")
            
            ' Sacar las compras contado
            rs.Open "select sum(total) as tot from compras" _
            & " where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='FAC' or tipodoc='N/D') and (contado=1) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            Total = 0
            If Not rs.EOF Then
                If Not IsNull(rs!tot) Then
                    Total = rs!tot
                End If
            End If
            rs.Close
            Set rs = Nothing
            rs.Open "select sum(total) as tot from compras" _
            & " where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='N/C') and (contado=1) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                If Not IsNull(rs!tot) Then
                    Total = Total - rs!tot
                End If
            End If
            rs.Close
            Set rs = Nothing
            RptLisDetalleComprobantesProveedores.lblcontado.caption = Format(Total, "#,###,##0.###0")
            
            ' Sacar las Compras Cta Cte
            rs.Open "select sum(total) as tot from compras" _
            & " where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='FAC' or tipodoc='N/D') and (contado=0) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            Total = 0
            If Not rs.EOF Then
                If Not IsNull(rs!tot) Then
                    Total = rs!tot
                End If
            End If
            rs.Close
            Set rs = Nothing
            rs.Open "select sum(total) as tot from compras" _
            & " where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='N/C') and (contado=0) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                If Not IsNull(rs!tot) Then
                    Total = Total - rs!tot
                End If
            End If
            rs.Close
            Set rs = Nothing
            '**********************
            rs.Open "select sum(total) as tot from transcom" _
            & " where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='FAC' or tipodoc='N/D')  and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                If Not IsNull(rs!tot) Then
                    Total = Total + rs!tot
                End If
            End If
            rs.Close
            Set rs = Nothing
            rs.Open "select sum(total) as tot from transcom" _
            & " where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='N/C') and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                If Not IsNull(rs!tot) Then
                    Total = Total - rs!tot
                End If
            End If
            rs.Close
            Set rs = Nothing
            RptLisDetalleComprobantesProveedores.lblcta.caption = Format(Total, "#,###,##0.###0")
            
            ' Sacar las Compras Locales
            rs.Open "select sum(total) as tot from compras" _
            & " where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='FAC' or tipodoc='N/D')  and (exterior= 0) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            Total = 0
            If Not rs.EOF Then
                If Not IsNull(rs!tot) Then
                    Total = Total + rs!tot
                End If
            End If
            rs.Close
            Set rs = Nothing
            rs.Open "select sum(total) as tot from compras" _
            & " where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='N/C')  and (exterior = 0) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                If Not IsNull(rs!tot) Then
                    Total = Total - rs!tot
                End If
            End If
            rs.Close
            Set rs = Nothing
            '**********************
            rs.Open "select sum(total) as tot from transcom" _
            & " where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='FAC' or tipodoc='N/D') and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                If Not IsNull(rs!tot) Then
                    Total = Total + rs!tot
                End If
            End If
            rs.Close
            Set rs = Nothing
            rs.Open "select sum(total) as tot from transcom" _
            & " where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='N/C') and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                If Not IsNull(rs!tot) Then
                    Total = Total - rs!tot
                End If
            End If
            rs.Close
            Set rs = Nothing
            RptLisDetalleComprobantesProveedores.lblloc.caption = Format(Total, "#,###,##0.###0")
            
            'Sacar las Compras Exterior
            rs.Open "select sum(total) as tot from compras" _
            & " where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='FAC' or tipodoc='N/D')  and (exterior= 1) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            Total = 0
            If Not rs.EOF Then
                If Not IsNull(rs!tot) Then
                    Total = Total + rs!tot
                End If
            End If
            rs.Close
            Set rs = Nothing
            rs.Open "select sum(total) as tot from compras" _
            & " where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='N/C')  and (exterior = 1) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                If Not IsNull(rs!tot) Then
                    Total = Total - rs!tot
                End If
            End If
            rs.Close
            Set rs = Nothing
            '**********************
            rs.Open "select sum(total) as tot from transcom" _
            & " where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='FAC' or tipodoc='N/D')  and (exterior= 1) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                If Not IsNull(rs!tot) Then
                    Total = Total + rs!tot
                End If
            End If
            rs.Close
            Set rs = Nothing
            rs.Open "select sum(total) as tot from transcom" _
            & " where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='N/C')  and (exterior = 1) and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                If Not IsNull(rs!tot) Then
                    Total = Total - rs!tot
                End If
            End If
            rs.Close
            Set rs = Nothing
            RptLisDetalleComprobantesProveedores.lblext.caption = Format(Total, "#,###,##0.###0")
            
            'Sacar Compras
            rs.Open "select sum(total) as tot from compras" _
            & " where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='FAC' or tipodoc='N/D') and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            Total = 0
            If Not rs.EOF Then
                If Not IsNull(rs!tot) Then
                    Total = Total + rs!tot
                End If
            End If
            rs.Close
            Set rs = Nothing
            rs.Open "select sum(total) as tot from compras" _
            & " where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='N/C') and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                If Not IsNull(rs!tot) Then
                    Total = Total - rs!tot
                End If
            End If
            rs.Close
            Set rs = Nothing
            '**********************
            rs.Open "select sum(total) as tot from transcom" _
            & " where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='FAC' or tipodoc='N/D') and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                If Not IsNull(rs!tot) Then
                    Total = Total + rs!tot
                End If
            End If
            rs.Close
            Set rs = Nothing
            rs.Open "select sum(total) as tot from transcom" _
            & " where mesimp=" & (cmbmes.ListIndex + 1) & " and anoimp=" & Val(Trim(cmbaño.Text)) & " and (tipodoc='N/C') and activo=1", DataEnvironment1.Sistema, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                If Not IsNull(rs!tot) Then
                    Total = Total - rs!tot
                End If
            End If
            rs.Close
            Set rs = Nothing
            RptLisDetalleComprobantesProveedores.lblcomp.caption = Format(Total, "#,###,##0.###0")
            RptLisDetalleComprobantesProveedores.lblfecha = Date
            RptLisDetalleComprobantesProveedores.lblTitulo.caption = "Listado de Deatalles de comprobantes Proveedores del  mes de " & Trim(cmbmes.Text) & " - Año " & Trim(cmbaño.Text)
            RptLisDetalleComprobantesProveedores.PageSettings.Orientation = ddOLandscape
            RptLisDetalleComprobantesProveedores.Show
            
        Else
            MsgBox "Debe seleccionar alguna de las opciones", 48, "Atencion"
        End If
    End If
End Sub

Private Sub cmdcancelar_Click()
    dtfechad = Date
    dtfechah = Date
    cmbmes.Text = ""
    cmbaño.Text = ""
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim min As Long, max As Long, i As Long

    dtfechad = Date
    dtfechah = Date
    cmbmes.Text = ""
    cmbaño.Text = ""
max = Year(Date)
min = max - 6
cmbaño.clear
For i = min To max
    cmbaño.AddItem i
Next

    
End Sub

''Private Sub Form_Unload(cancel As Integer)
''    If msFCtmp > "" Then
''        TablaTempBorrar msFCtmp
''        msFCtmp = ""
''    End If
''End Sub


Private Sub optfecha_Click()
    If optmes.Value = True Then
        cmbmes.enabled = True
        cmbaño.enabled = True
        dtfechad.enabled = False
        dtfechah.enabled = False
    Else
        If optFecha.Value = True Then
            cmbmes.enabled = False
            cmbaño.enabled = False
            dtfechad.enabled = True
            dtfechah.enabled = True
        End If
    End If
End Sub

Private Sub optmes_Click()
    If optmes.Value = True Then
        cmbmes.enabled = True
        cmbaño.enabled = True
        dtfechad.enabled = False
        dtfechah.enabled = False
    Else
        If optFecha.Value = True Then
            cmbmes.enabled = False
            cmbaño.enabled = False
            dtfechad.enabled = True
            dtfechah.enabled = True
        End If
    End If
    
End Sub


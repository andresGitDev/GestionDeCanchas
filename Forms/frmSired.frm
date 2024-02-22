VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSired 
   Caption         =   "SIRED"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   2115
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "CITI COMPRAS"
      Height          =   375
      Left            =   3840
      TabIndex        =   15
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton Command7 
      Caption         =   "CITI VENTAS"
      Height          =   375
      Left            =   3840
      TabIndex        =   13
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Salir"
      Height          =   375
      Left            =   4080
      TabIndex        =   12
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Otras Percepciones"
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   5160
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Detalle de Facturas"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   4440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cabecera de Facturas"
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   3720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Libro COMPRAS"
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   3000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      OLEDropMode     =   1
      CustomFormat    =   "MM/yyyy"
      Format          =   99287043
      CurrentDate     =   39361
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Libro VENTAS"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblCitiCompra 
      Caption         =   "C:\CITI\"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label lblCitiVenta 
      Caption         =   "C:\CITI\"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label Label6 
      Caption         =   "Periodo de informe :"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "C:\gestion\SIRED\"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   5280
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label4 
      Caption         =   "C:\gestion\SIRED\"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   4560
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "C:\gestion\SIRED\"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3720
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "C:\gestion\SIRED\"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3000
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "C:\gestion\SIRED\"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   3375
   End
End
Attribute VB_Name = "frmSired"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()  'LIBRO DE VENTAS
    Dim ARCH As String
    Dim documento As Integer
    Dim nom As String
    Dim tI As Double
    Dim impLIQ As String
    Dim impRNI As String
    Dim impEXE As String
    Dim totimpLIQ As Double
    Dim totimpRNI As Double
    Dim totimpEXE As Double
    Dim IIBB As Double
    Dim fec As String
    Dim pri As Date
    Dim seg As Date
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim rs3 As New ADODB.Recordset
    Dim i As Long
    Dim j As Long
    Dim tot As Double
    Dim netgra As Double
    Dim netNOgra As Double
    Dim rni As Double
    Dim relleno As String
    Dim relleno2 As String
    Dim a
    
    
    pri = "01/" & Month(DTPicker1.Value) & "/" & Year(DTPicker1.Value)
    seg = ultimoDiaDelMes(DTPicker1.Value)
    a = "select * from facturaventa where fecha>=" & ssFecha(pri) & " and fecha<=" & ssFecha(seg) & " and (tipodoc='FAA' or tipodoc='FAB' or tipodoc='NCA' or tipodoc='NCB') order by fecha,tipodoc,nrofactura"
    rs3.Open "select * from facturaventa where fecha>=" & ssFecha(pri) & " and fecha<=" & ssFecha(seg) & " and (tipodoc='FAA' or tipodoc='FAB' or tipodoc='NCA' or tipodoc='NCB') order by fecha,tipodoc,nrofactura", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    
'********************************** ARCHIVO

    ARCH = "VENTAS_" & Year(DTPicker1.Value) & Format(Month(DTPicker1.Value), "00")
    rs.Open "select * from datosempresa where idempresa=15", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    
    If (rs3.EOF = True And rs3.BOF = True) Or IsNull(rs3!codigo) Then
        MsgBox "No se ha encontrado datos para este periodo.", , "ATENCION"
        Exit Sub
    End If
    
    rs3.MoveFirst
    If Not IsNull(rs3) Or Not IsEmpty(rs3) Or Not (rs3.EOF = True And rs3.BOF = True) Then
        Open "C:\gestion\SIRED\" & ARCH & ".txt" For Output As #1
    End If
    i = 0
    While Not rs3.EOF
    
        rs2.Open "select i.*,C.CUIT from clientes c inner join ivas i on i.codigo=c.iva where c.codigo=" & rs3!cliente, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        If rs3!tipoiva = 1 Or rs3!tipoiva = 4 Or rs3!tipoiva = 6 Then
        'If ComboCodigo(cmbTipoIva) = 1 Or ComboCodigo(cmbTipoIva) = 4 Or ComboCodigo(cmbTipoIva) = 6 Then
            documento = 80
        Else
            documento = 86
        End If
        nom = Trim(rs3!RAZONSOCIAL)
        While Len(nom) < 30
            nom = Chr(32) & nom
        Wend
        
        fec = Year(rs3!Fecha) & Format(Month(rs3!Fecha), "00") & Format(Day(rs3!Fecha), "00")
        'ver si tengo que multiplicar * 100 los valores para sacarle la coma!!!!!!
        tI = Trim(rs3!tipoiva)
        If tI = 5 Or tI = 4 Or tI = 6 Or tI = 10 Then 'facturas B
            impLIQ = "000000000000000"
            impRNI = "000000000000000"
            impEXE = "000000000000000" 'Format(TxtIVA * 100, "000000000000000")
        ElseIf tI = 1 Or tI = 2 Then  'facturas A
            If tI = 1 Then
                impLIQ = Format(rs3!Iva * 100, "000000000000000")
                impRNI = "000000000000000"
                impEXE = "000000000000000"
            ElseIf tI = 2 Then
                impLIQ = "000000000000000"
                impRNI = Format(rs3!Iva * 100, "000000000000000")
                impEXE = "000000000000000"
            ElseIf tI = 8 Then
                impLIQ = "000000000000000"
                impRNI = "000000000000000"
                impEXE = Format(rs3!Iva * 100, "000000000000000")
            End If
        End If
        
        If Trim(rs3!TIPODOC) = "FAA" Or Trim(rs3!TIPODOC) = "FAB" Or Trim(rs3!TIPODOC) = "NDA" Or Trim(rs3!TIPODOC) = "NDB" Then
            tot = tot + CDbl(rs3!Total)
            totimpLIQ = totimpLIQ + CDbl(impLIQ)
            totimpRNI = totimpRNI + CDbl(impRNI)
            totimpEXE = totimpEXE + CDbl(impEXE)
            IIBB = IIBB + CDbl(rs3!IIBB)
        ElseIf Trim(rs3!TIPODOC) = "NCA" Or Trim(rs3!TIPODOC) = "NCB" Then
            tot = tot - CDbl(rs3!Total)
            totimpLIQ = totimpLIQ - CDbl(impLIQ)
            totimpRNI = totimpRNI - CDbl(impRNI)
            totimpEXE = totimpEXE - CDbl(impEXE)
            IIBB = IIBB - CDbl(rs3!IIBB)
        End If
        
        
        If Trim(rs3!TIPODOC) = "FAA" Then  'en las tablas de la afip el 01 es factura A
            'Open "C:\gestion\SIRED\" & arch & ".txt" For Output As #1
            'Write #1, cmbCliente.Text & cmbCliente.Text & Chr(32) & cmbCliente.Text            ' Text1.Text el chr(9)=tab  vbCrLf=enter   chr(32)=un espacio
            relleno2 = ""
            j = 1
            While j < 61
                relleno2 = relleno2 & Chr(32)
                j = j + 1
            Wend
            Print #1, "1" & fec & "01" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & "2100" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & Chr(32) & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
            "Sin comentarios" & relleno2
            'Close #1
            netgra = netgra + CDbl(rs3!Neto)
            i = i + 1
        ElseIf Trim(rs3!TIPODOC) = "FAB" Then
            'If CDbl(txtTotal.caption) >= CDbl("10000") Then  'lo quieren hacer uno por uno
                'Open "C:\gestion\SIRED\" & arch & ".txt" For Output As #1
                'Write #1, cmbCliente.Text & cmbCliente.Text & Chr(32) & cmbCliente.Text            ' Text1.Text el chr(9)=tab  vbCrLf=enter
                relleno2 = ""
                j = 1
                While j < 61
                    relleno2 = relleno2 & Chr(32)
                    j = j + 1
                Wend
                Print #1, "1" & fec & "06" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & "E" & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                "Sin comentarios" & relleno2 '& vbCrLf
                'txtneto
                netNOgra = netNOgra + CDbl(rs3!Total)
                'Close #1
            'End If
            i = i + 1
        ElseIf Trim(rs3!TIPODOC) = "NCA" Then
            relleno2 = ""
            j = 1
            While j < 61
                relleno2 = relleno2 & Chr(32)
                j = j + 1
            Wend
            Print #1, "1" & fec & "03" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & "2100" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & Chr(32) & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
            "Sin comentarios" & relleno2
            netgra = netgra - CDbl(rs3!Neto)
            i = i + 1
        ElseIf Trim(rs3!TIPODOC) = "NDA" Then
            relleno2 = ""
            j = 1
            While j < 61
                relleno2 = relleno2 & Chr(32)
                j = j + 1
            Wend
            Print #1, "1" & fec & "02" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & "2100" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & Chr(32) & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
            "Sin comentarios" & relleno2
            netgra = netgra + CDbl(rs3!Neto)
            i = i + 1
        ElseIf Trim(rs3!TIPODOC) = "NCB" Then
            relleno2 = ""
            j = 1
            While j < 61
                relleno2 = relleno2 & Chr(32)
                j = j + 1
            Wend
            Print #1, "1" & fec & "08" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & "E" & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                "Sin comentarios" & relleno2 '& vbCrLf
            netNOgra = netNOgra - CDbl(rs3!Total)
            i = i + 1
        ElseIf Trim(rs3!TIPODOC) = "NDB" Then
            relleno2 = ""
            j = 1
            While j < 61
                relleno2 = relleno2 & Chr(32)
                j = j + 1
            Wend
            Print #1, "1" & fec & "07" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & "E" & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                "Sin comentarios" & relleno2 '& vbCrLf
            netNOgra = netNOgra + CDbl(rs3!Total)
            i = i + 1
        End If
        
        
        rs3.MoveNext
        Set rs2 = Nothing
    Wend
    'registro de tipo 2
    j = 1
    While j < 123
        relleno = relleno & Chr(32)
        j = j + 1
    Wend
    relleno2 = ""
    j = 1
    While j < 30
        relleno2 = relleno2 & Chr(32)
        j = j + 1
    Wend
    
    Print #1, "2" & Year(DTPicker1.Value) & Format(Month(DTPicker1.Value), "00") & relleno2 & Format(i, "000000000000") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Trim(Mid(rs!CuitEmpresa, 1, 2)) & Trim(Mid(rs!CuitEmpresa, 4, 8)) & Trim(Mid(rs!CuitEmpresa, 13, 1)) & relleno2 & Chr(32) & _
     Format(tot * 100, "000000000000000") & Format(netNOgra * 100, "000000000000000") & Format(netgra * 100, "000000000000000") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Format(totimpLIQ, "000000000000000") & Format(totimpRNI, "000000000000000") & Format(totimpEXE, "000000000000000") & "000000000000000" & Format(IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & relleno
    
    
    If Not IsNull(rs3) Or Not IsEmpty(rs3) Or Not (rs3.EOF = True And rs3.BOF = True) Then
        Close #1
    End If
    Set rs3 = Nothing
    Set rs = Nothing
    MsgBox "Archivo generado correctamente."
                
            
    '******************************************************************************************

End Sub

Private Sub Command2_Click() 'LIBRO DE COMPRAS
    Dim ARCH As String
    Dim documento As Integer
    Dim nom As String
    Dim tI As Double
    Dim impLIQ As String
    Dim impRNI As String
    Dim impEXE As String
    Dim totimpLIQ As Double
    Dim totimpRNI As Double
    Dim totimpEXE As Double
    Dim totPerAgr As Double
    Dim totPerOtro As Double
    Dim totImpInt As Double
    Dim IIBB As Double
    Dim fec As String
    Dim pri As Date
    Dim seg As Date
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim rs3 As New ADODB.Recordset
    Dim i As Long
    Dim j As Long
    Dim x As Long
    Dim p As Long
    Dim tot As Double
    Dim netgra As Double
    Dim netNOgra As Double
    Dim rni As Double
    Dim relleno As String
    Dim relleno2 As String
    Dim val As Double
    Dim val2 As Double
    Dim val3 As Double
    Dim val4 As Double
    Dim val5 As Double
    Dim val6 As Double
    Dim val7 As Double
    Dim val8 As Double
    Dim val9 As Double
    Dim a As Integer
    Dim fecCAI As String
    Dim cui As String
    
    
    pri = "01/" & Month(DTPicker1.Value) & "/" & Year(DTPicker1.Value)
    seg = ultimoDiaDelMes(DTPicker1.Value)
    rs3.Open "select fecha,codpr,tipoiva,tipodoc,nrodoc,total,neto,exento,iva_21,iva_27,iva_9,iva_10,imp_int,percepc,der_est,ibcapital,ibprovincia,retgan,retganpago,formadepago,razonsocialprov,cuitprov,controlador,suc,anoimp,aduana,destinacion,verifidespacho,despacho,cai,vencecai from TRANSCOM where fecha>=" & ssFecha(pri) & " and fecha<=" & ssFecha(seg) & " and tipodoc<>'RAC'" & _
        "union select fecha,codpr,tipoiva,tipodoc,nrodoc,total,neto,exento,iva_21,iva_27,iva_9,iva_10,imp_int,percepc,der_est,ibcapital,ibprovincia,retgan,retganpago,formadepago,razonsocialprov,cuitprov,controlador,suc,anoimp,aduana,destinacion,verifidespacho,despacho,cai,vencecai from compras where fecha>=" & ssFecha(pri) & " and fecha<=" & ssFecha(seg) & " and tipodoc<>'RAC' order by fecha,tipodoc,nrodoc", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    
    'MsgBox "" & rs3.RecordCount
'********************************** ARCHIVO
                ARCH = "COMPRAS_" & Year(DTPicker1.Value) & Format(Month(DTPicker1.Value), "00")
                rs.Open "select * from datosempresa where idempresa=15", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                
                If (rs3.EOF = True And rs3.BOF = True) Or IsNull(rs3) Then
                    MsgBox "No se ha encontrado datos para este periodo.", , "ATENCION"
                    Exit Sub
                End If
                
                rs3.MoveFirst
                If Not IsNull(rs3) Or Not IsEmpty(rs3) Or Not (rs3.EOF = True And rs3.BOF = True) Then
                    Open "C:\gestion\SIRED\" & ARCH & ".txt" For Output As #1
                End If
                i = 0
                While Not rs3.EOF
                
                    rs2.Open "select i.*,C.CUIT from prov c inner join ivas i on i.codigo=c.tipoiva where c.codigo=" & rs3!CODPR, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                    a = 0
                    If rs3!tipoiva = 1 Or rs3!tipoiva = 4 Or rs3!tipoiva = 6 Then
                    'If ComboCodigo(cmbTipoIva) = 1 Or ComboCodigo(cmbTipoIva) = 4 Or ComboCodigo(cmbTipoIva) = 6 Then
                        documento = 80
                    Else
                        documento = 86
                    End If
                    
                    If rs3!cuitprov = "" Then
                        cui = Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32)
                        documento = 99
                    Else
                        cui = Trim(Mid(rs3!cuitprov, 1, 2)) & Trim(Mid(rs3!cuitprov, 4, 8)) & Trim(Mid(rs3!cuitprov, 13, 1))
                    End If
                    
                    nom = Trim(rs3!razonsocialprov)
                    If Trim(rs3!razonsocialprov) = "Nextel" Then
                        nom = ""
                    End If
                    While Len(nom) < 30
                        nom = Chr(32) & nom
                    Wend
                    
                    fec = Year(rs3!Fecha) & Format(Month(rs3!Fecha), "00") & Format(Day(rs3!Fecha), "00")
                    'ver si tengo que multiplicar * 100 los valores para sacarle la coma!!!!!!
                    tI = Trim(rs3!tipoiva)
                    If tI = 5 Or tI = 4 Or tI = 6 Or tI = 10 Then 'facturas B
                        If tI = 5 Or tI = 6 Then
                            impLIQ = "000000000000000"
                            impRNI = "000000000000000"
                            impEXE = "000000000000000" 'Format(TxtIVA * 100, "000000000000000")
                        ElseIf tI = 4 Or tI = 10 Then
                            impLIQ = "000000000000000"
                            impRNI = "000000000000000"
                            impEXE = Format(rs3!EXENTO * 100, "000000000000000")
                        End If
                    ElseIf tI = 1 Or tI = 2 Then  'facturas A
                        If tI = 1 Then
                            impLIQ = Format(rs3!IVA_21 * 100, "000000000000000")
                            impRNI = "000000000000000"
                            impEXE = "000000000000000"
                        ElseIf tI = 2 Then
                            impLIQ = "000000000000000"
                            impRNI = Format(rs3!IVA_21 * 100, "000000000000000")
                            impEXE = "000000000000000"
                        
                        End If
                    End If
                    
                    tot = tot + CDbl(rs3!Total)
                    totimpLIQ = totimpLIQ + CDbl((rs3!IVA_21 + rs3!iva_10 + rs3!IVA_27) * 100)
                    totimpRNI = totimpRNI + CDbl(impRNI)
                    totimpEXE = totimpEXE + CDbl(rs3!EXENTO * 100)
                    totPerAgr = totPerAgr + CDbl(rs3!percepc * 100)
                    totPerOtro = totPerOtro + CDbl(rs3!der_est * 100)
                    IIBB = IIBB + CDbl((rs3!ibcapital + rs3!ibprovincia) * 100)
                    totImpInt = totImpInt + CDbl(rs3!imp_int * 100)
                    
                    x = 0
                    If rs3!IVA_21 > 0 Then x = x + 1
                    If rs3!IVA_27 > 0 Then x = x + 1
                    If rs3!iva_10 > 0 Then x = x + 1
                    
                    j = 1
                    relleno = "Sin comentarios"
                    While j < 61
                        relleno = relleno & Chr(32)
                        j = j + 1
                    Wend
                    
                    fecCAI = IIf(Trim(rs3!vencecai) = "01/01/1900", "00000000", Year(rs3!vencecai) & Format(Month(rs3!vencecai), "00") & Format(Day(rs3!vencecai), "00"))
                    
                    'If tI = 1 Or tI = 2 Then 'Trim(rs3!TIPODOC) = "FAA" Then 'en las tablas de la afip el 01 es factura A
                        'Open "C:\gestion\SIRED\" & arch & ".txt" For Output As #1
                        'Write #1, cmbCliente.Text & cmbCliente.Text & Chr(32) & cmbCliente.Text            ' Text1.Text el chr(9)=tab  vbCrLf=enter   chr(32)=un espacio
                        If x = 0 Then
                            Print #1, "1" & fec & "01" & IIf(rs3!controlador = 1, "C", Chr(32)) & Format(rs3!suc, "0000") & Format(rs3!NroDoc, "00000000000000000000") & fec & Format(IIf(Trim(rs3!aduana) = "", 0, Trim(rs3!aduana)), "000") & IIf(IsNull(rs3!destinacion) Or Trim(rs3!destinacion) = "", "    ", Trim(rs3!destinacion)) & Format(Trim(rs3!despacho), "000000") & IIf(Trim(rs3!verifidespacho) = "", Chr(32), Trim(rs3!verifidespacho)) & documento & cui & nom & _
                                Format(rs3!Total * 100, "000000000000000") & Format(rs3!EXENTO * 100, "000000000000000") & Format(rs3!Neto * 100, "000000000000000") & "0000" & impLIQ & Format(rs3!EXENTO * 100, "000000000000000") & Format(rs3!percepc * 100, "000000000000000") & Format(rs3!der_est * 100, "000000000000000") & Format((rs3!ibcapital + rs3!ibprovincia) * 100, "000000000000000") & "000000000000000" & Format(rs3!imp_int * 100, "000000000000000") & _
                                Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & "E" & Format(rs3!cai, "00000000000000") & fecCAI & relleno
                            i = i + 1
                        Else
                            For p = 1 To x
                                If x = p Then
                                    val = rs3!Total * 100
                                    val2 = rs3!EXENTO * 100  '(rs3!EXENTO + rs3!imp_int + rs3!IVA_21 + rs3!IVA_27 + rs3!iva_10) * 100 '
                                    val3 = rs3!Neto * 100
                                    val4 = rs3!imp_int * 100
                                    val5 = rs3!EXENTO * 100
                                    'val6 = (rs3!IVA_21 + rs3!IVA_27 + rs3!iva_10) * 100
                                    val7 = rs3!percepc
                                    val8 = rs3!der_est
                                    val9 = (rs3!ibcapital + rs3!ibprovincia) * 100
                                Else
                                    val = 0 ' SE MUESTRA SOLO EN UN REGISTRO EL TOTAL QUE ES EL ULTIMO EN EL CASO DE TENER VARIAS ALICUOTAS
                                    val2 = 0
                                    val3 = 0
                                    val4 = 0
                                    val5 = 0
                                    'val6 = 0
                                    val7 = 0
                                    val8 = 0
                                    val9 = 0
                                End If
                                If x > 0 Then
                                    If rs3!IVA_21 > 0 And a = 0 Then
                                        val6 = rs3!IVA_21 * 100
                                        'val2 = rs3!IVA_21 * 100
                                    ElseIf rs3!IVA_27 > 0 And (a = 1 Or a = 0) Then
                                        val6 = rs3!IVA_27 * 100
                                        'val2 = rs3!IVA_27 * 100
                                    ElseIf rs3!iva_10 And (a = 2 Or a = 1 Or a = 0) Then
                                        val6 = rs3!iva_10 * 100
                                        'val2 = rs3!iva_10 * 100
                                    End If
                                Else
                                    val6 = 0
                                    'val2 = (CDbl(rs3!EXENTO) + CDbl(rs3!imp_int) + CDbl(rs3!IVA_21) + CDbl(rs3!IVA_27) + CDbl(rs3!iva_10)) * 100
                                End If
                                
                                If rs3!IVA_21 > 0 And a = 0 Then
                                    'If val3 = 0 Then val3 = rs3!IVA_21
                                    Print #1, "1" & fec & "01" & IIf(rs3!controlador = 1, "C", Chr(32)) & Format(rs3!suc, "0000") & Format(rs3!NroDoc, "00000000000000000000") & fec & Format(IIf(Trim(rs3!aduana) = "", 0, Trim(rs3!aduana)), "000") & IIf(IsNull(rs3!destinacion) Or Trim(rs3!destinacion) = "", "    ", Trim(rs3!destinacion)) & Format(Trim(rs3!despacho), "000000") & IIf(Trim(rs3!verifidespacho) = "", Chr(32), Trim(rs3!verifidespacho)) & documento & cui & nom & _
                                        Format(val, "000000000000000") & Format(val2, "000000000000000") & Format(val3, "000000000000000") & "2100" & Format(val6, "000000000000000") & Format(val5, "000000000000000") & Format(val7 * 100, "000000000000000") & Format(val8 * 100, "000000000000000") & Format(val9, "000000000000000") & "000000000000000" & Format(val4, "000000000000000") & _
                                        Format(rs3!tipoiva, "00") & "PES" & "0001000000" & Format(x, "0") & Chr(32) & Format(rs3!cai, "00000000000000") & fecCAI & relleno
                                    a = a + 1
                                    i = i + 1
                                Else
                                    If rs3!IVA_27 > 0 And (a = 1 Or a = 0) Then
                                        'If val3 = 0 Then val3 = rs3!IVA_27
                                        Print #1, "1" & fec & "01" & IIf(rs3!controlador = 1, "C", Chr(32)) & Format(rs3!suc, "0000") & Format(rs3!NroDoc, "00000000000000000000") & fec & Format(IIf(Trim(rs3!aduana) = "", 0, Trim(rs3!aduana)), "000") & IIf(IsNull(rs3!destinacion) Or Trim(rs3!destinacion) = "", "    ", Trim(rs3!destinacion)) & Format(Trim(rs3!despacho), "000000") & IIf(Trim(rs3!verifidespacho) = "", Chr(32), Trim(rs3!verifidespacho)) & documento & cui & nom & _
                                            Format(val, "000000000000000") & Format(val2, "000000000000000") & Format(val3, "000000000000000") & "2700" & Format(val6, "000000000000000") & Format(val5, "000000000000000") & Format(val7 * 100, "000000000000000") & Format(val8 * 100, "000000000000000") & Format(val9, "000000000000000") & "000000000000000" & Format(val4, "000000000000000") & _
                                            Format(rs3!tipoiva, "00") & "PES" & "0001000000" & Format(x, "0") & Chr(32) & Format(rs3!cai, "00000000000000") & fecCAI & relleno
                                        a = a + 1
                                        i = i + 1
                                    Else
                                        If rs3!iva_10 > 0 And (a = 2 Or a = 1 Or a = 0) Then
                                            'If val3 = 0 Then val3 = rs3!iva_10
                                            Print #1, "1" & fec & "01" & IIf(rs3!controlador = 1, "C", Chr(32)) & Format(rs3!suc, "0000") & Format(rs3!NroDoc, "00000000000000000000") & fec & Format(IIf(Trim(rs3!aduana) = "", 0, Trim(rs3!aduana)), "000") & IIf(IsNull(rs3!destinacion) Or Trim(rs3!destinacion) = "", "    ", Trim(rs3!destinacion)) & Format(Trim(rs3!despacho), "000000") & IIf(Trim(rs3!verifidespacho) = "", Chr(32), Trim(rs3!verifidespacho)) & documento & cui & nom & _
                                                Format(val, "000000000000000") & Format(val2, "000000000000000") & Format(val3, "000000000000000") & "1050" & Format(val6, "000000000000000") & Format(val5, "000000000000000") & Format(val7 * 100, "000000000000000") & Format(val8 * 100, "000000000000000") & Format(val9, "000000000000000") & "000000000000000" & Format(val4, "000000000000000") & _
                                                Format(rs3!tipoiva, "00") & "PES" & "0001000000" & Format(x, "0") & Chr(32) & Format(rs3!cai, "00000000000000") & fecCAI & relleno
                                            i = i + 1
                                        End If
                                    End If
                                End If
                            Next p
                        End If
                        'Close #1
                        If tI = 1 Or tI = 2 Then
                            netgra = netgra + CDbl(rs3!Neto)
                            netNOgra = netNOgra + CDbl(rs3!EXENTO) '+ CDbl(rs3!imp_int) + CDbl(rs3!IVA_21) + CDbl(rs3!IVA_27) + CDbl(rs3!iva_10)
                        ElseIf tI = 4 Or tI = 6 Or tI = 10 Then    'Or tI = 5
                            netNOgra = netNOgra + CDbl(rs3!Neto) '+ CDbl(rs3!imp_int) + CDbl(rs3!IVA_21) + CDbl(rs3!IVA_27) + CDbl(rs3!iva_10)
                        ElseIf tI = 5 Then
                            netNOgra = netNOgra + CDbl(rs3!EXENTO) '+ CDbl(rs3!imp_int) + CDbl(rs3!IVA_21) + CDbl(rs3!IVA_27) + CDbl(rs3!iva_10)
                            netgra = netgra + CDbl(rs3!Neto)
                        End If
                        
                    'ElseIf tI = 4 Or tI = 6 Or tI = 10 Then 'Trim(rs3!TIPODOC) = "FAB" Then
                        'If CDbl(txtTotal.caption) >= CDbl("10000") Then  'lo quieren hacer uno por uno
                            'Open "C:\gestion\SIRED\" & arch & ".txt" For Output As #1
                            'Write #1, cmbCliente.Text & cmbCliente.Text & Chr(32) & cmbCliente.Text            ' Text1.Text el chr(9)=tab  vbCrLf=enter
                    '        Print #1, "1" & fec & "06" & IIf(rs3!controlador = 1, "C", Chr(32)) & Format(rs3!suc, "0000") & Format(rs3!NroDoc, "00000000000000000000") & fec & Format(Trim(rs3!aduana), "000") & IIf(IsNull(rs3!destinacion) Or Trim(rs3!destinacion) = "", "    ", Trim(rs3!destinacion)) & Format(Trim(rs3!despacho), "000000") & IIf(Trim(rs3!verifidespacho) = "", Chr(32), Trim(rs3!verifidespacho)) & documento & Trim(Mid(rs3!cuitprov, 1, 2)) & Trim(Mid(rs3!cuitprov, 4, 8)) & Trim(Mid(rs3!cuitprov, 13, 1)) & nom & _
                    '            Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!iibb * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & "E" & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & "00000000" & _
                    '            Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(13)
                            'txtneto
                    '        netNOgra = netNOgra + CDbl(rs3!Total)
                            'Close #1
                        'End If
                    '    i = i + 1
                    'ElseIf tI = 5 Then 'consumidor final
                    '    Print #1, "1" & fec & "06" & IIf(rs3!controlador = 1, "C", Chr(32)) & Format(rs3!suc, "0000") & Format(rs3!NroDoc, "00000000000000000000") & fec & Format(Trim(rs3!aduana), "000") & IIf(IsNull(rs3!destinacion) Or Trim(rs3!destinacion) = "", "    ", Trim(rs3!destinacion)) & Format(Trim(rs3!despacho), "000000") & IIf(Trim(rs3!verifidespacho) = "", Chr(32), Trim(rs3!verifidespacho)) & documento & Trim(Mid(rs3!cuitprov, 1, 2)) & Trim(Mid(rs3!cuitprov, 4, 8)) & Trim(Mid(rs3!cuitprov, 13, 1)) & nom & _
                    '        Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!iibb * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & "E" & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & "00000000" & _
                    '        Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(13)
                    'End If
                    rs3.MoveNext
                    Set rs2 = Nothing
                Wend
                'registro de tipo 2
                relleno = ""
                j = 1
                While j < 115
                    relleno = relleno & Chr(32)
                    j = j + 1
                Wend
                
                relleno2 = ""
                j = 1
                While j < 11
                    relleno2 = relleno2 & Chr(32)
                    j = j + 1
                Wend
                
                Print #1, "2" & Year(DTPicker1.Value) & Format(Month(DTPicker1.Value), "00") & relleno2 & Format(i, "000000000000") & relleno2 & relleno2 & relleno2 & Chr(32) & Trim(Mid(rs!CuitEmpresa, 1, 2)) & Trim(Mid(rs!CuitEmpresa, 4, 8)) & Trim(Mid(rs!CuitEmpresa, 13, 1)) & relleno2 & relleno2 & relleno2 & _
                 Format(tot * 100, "000000000000000") & Format(netNOgra * 100, "000000000000000") & Format(netgra * 100, "000000000000000") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Format(totimpLIQ, "000000000000000") & Format(totimpEXE, "000000000000000") & Format(totPerAgr, "000000000000000") & Format(totPerOtro, "000000000000000") & Format(IIBB, "000000000000000") & "000000000000000" & Format(totImpInt, "000000000000000") & relleno

                
                If Not IsNull(rs3) Or Not IsEmpty(rs3) Or Not (rs3.EOF = True And rs3.BOF = True) Then
                    Close #1
                End If
                Set rs3 = Nothing
                Set rs = Nothing
                MsgBox "Archivo generado correctamente."
                
            
    '******************************************************************************************

End Sub

Private Sub Command3_Click()   'CABECERA DE FACTURAS
    Dim ARCH As String
    Dim documento As Integer
    Dim nom As String
    Dim tI As Double
    Dim impLIQ As String
    Dim impRNI As String
    Dim impEXE As String
    Dim totimpLIQ As Double
    Dim totimpRNI As Double
    Dim totimpEXE As Double
    Dim IIBB As Double
    Dim fec As String
    Dim pri As Date
    Dim seg As Date
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim rs3 As New ADODB.Recordset
    Dim i As Long
    Dim j As Long
    Dim tot As Double
    Dim netgra As Double
    Dim netNOgra As Double
    Dim rni As Double
    Dim relleno As String
    
    
    pri = "01/" & Month(DTPicker1.Value) & "/" & Year(DTPicker1.Value)
    seg = ultimoDiaDelMes(DTPicker1.Value)
    rs3.Open "select * from facturaventa where fecha>=" & ssFecha(pri) & " and fecha<=" & ssFecha(seg) & " and (tipodoc='FAA' or tipodoc='FAB' or tipodoc='NCA' or tipodoc='NCB') order by fecha,tipodoc,nrofactura", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'********************************** ARCHIVO
                ARCH = "CABECERA_" & Year(DTPicker1.Value) & Format(Month(DTPicker1.Value), "00")
                rs.Open "select * from datosempresa where idempresa=15", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                
                If (rs3.EOF = True And rs3.BOF = True) Or IsNull(rs3) Then
                    MsgBox "No se ha encontrado datos para este periodo.", , "ATENCION"
                    Exit Sub
                End If
                
                rs3.MoveFirst
                If Not IsNull(rs3) Or Not IsEmpty(rs3) Or Not (rs3.EOF = True And rs3.BOF = True) Then
                    Open "C:\gestion\SIRED\" & ARCH & ".txt" For Output As #1
                End If
                i = 0
                While Not rs3.EOF
                
                    rs2.Open "select i.*,C.CUIT from clientes c inner join ivas i on i.codigo=c.iva where c.codigo=" & rs3!cliente, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                    If rs3!tipoiva = 1 Or rs3!tipoiva = 4 Or rs3!tipoiva = 6 Then
                    'If ComboCodigo(cmbTipoIva) = 1 Or ComboCodigo(cmbTipoIva) = 4 Or ComboCodigo(cmbTipoIva) = 6 Then
                        documento = 80
                    Else
                        documento = 86
                    End If
                    nom = Trim(rs3!RAZONSOCIAL)
                    While Len(nom) < 30
                        nom = Chr(32) & nom
                    Wend
                    
                    fec = Year(rs3!Fecha) & Format(Month(rs3!Fecha), "00") & Format(Day(rs3!Fecha), "00")
                    'ver si tengo que multiplicar * 100 los valores para sacarle la coma!!!!!!
                    tI = Trim(rs3!tipoiva)
                    If tI = 5 Or tI = 4 Or tI = 6 Or tI = 10 Then 'facturas B
                        impLIQ = "000000000000000"
                        impRNI = "000000000000000"
                        impEXE = "000000000000000" 'Format(TxtIVA * 100, "000000000000000")
                    ElseIf tI = 1 Or tI = 2 Then  'facturas A
                        If tI = 1 Then
                            impLIQ = Format(rs3!Iva * 100, "000000000000000")
                            impRNI = "000000000000000"
                            impEXE = "000000000000000"
                        ElseIf tI = 2 Then
                            impLIQ = "000000000000000"
                            impRNI = Format(rs3!Iva * 100, "000000000000000")
                            impEXE = "000000000000000"
                        ElseIf tI = 8 Then
                            impLIQ = "000000000000000"
                            impRNI = "000000000000000"
                            impEXE = Format(rs3!Iva * 100, "000000000000000")
                        End If
                    End If
                    
                    If Trim(rs3!TIPODOC) = "FAA" Or Trim(rs3!TIPODOC) = "FAB" Or Trim(rs3!TIPODOC) = "NDA" Or Trim(rs3!TIPODOC) = "NDB" Then
                        tot = tot + CDbl(rs3!Total)
                        totimpLIQ = totimpLIQ + CDbl(impLIQ)
                        totimpRNI = totimpRNI + CDbl(impRNI)
                        totimpEXE = totimpEXE + CDbl(impEXE)
                        IIBB = IIBB + CDbl(rs3!IIBB)
                    ElseIf Trim(rs3!TIPODOC) = "NCA" Or Trim(rs3!TIPODOC) = "NCB" Then
                        tot = tot - CDbl(rs3!Total)
                        totimpLIQ = totimpLIQ - CDbl(impLIQ)
                        totimpRNI = totimpRNI - CDbl(impRNI)
                        totimpEXE = totimpEXE - CDbl(impEXE)
                        IIBB = IIBB - CDbl(rs3!IIBB)
                    End If
                                        
                    
                    If Trim(rs3!TIPODOC) = "FAA" Then 'en las tablas de la afip el 01 es factura A
                        'Open "C:\gestion\SIRED\" & arch & ".txt" For Output As #1
                        'Write #1, cmbCliente.Text & cmbCliente.Text & Chr(32) & cmbCliente.Text            ' Text1.Text el chr(9)=tab  vbCrLf=enter   chr(32)=un espacio
                        Print #1, "1" & fec & "01" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000") & Format(rs3!NroFactura, "00000000") & "001" & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & Chr(32) & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) '& Chr(13)
                        'Close #1
                        netgra = netgra + CDbl(rs3!Neto)
                        i = i + 1
                    ElseIf Trim(rs3!TIPODOC) = "FAB" Then
                        'If CDbl(txtTotal.caption) >= CDbl("10000") Then  'lo quieren hacer uno por uno
                            'Open "C:\gestion\SIRED\" & arch & ".txt" For Output As #1
                            'Write #1, cmbCliente.Text & cmbCliente.Text & Chr(32) & cmbCliente.Text            ' Text1.Text el chr(9)=tab  vbCrLf=enter
                            Print #1, "1" & fec & "06" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000") & Format(rs3!NroFactura, "00000000") & "001" & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & "E" & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) '& Chr(13)
                            'txtneto
                            netNOgra = netNOgra + CDbl(rs3!Total)
                            'Close #1
                        'End If
                        i = i + 1
                    ElseIf Trim(rs3!TIPODOC) = "NCA" Then
                        
                        Print #1, "1" & fec & "03" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000") & Format(rs3!NroFactura, "00000000") & "001" & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & Chr(32) & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) '& Chr(13)
                        netgra = netgra - CDbl(rs3!Neto)
                        i = i + 1
                    ElseIf Trim(rs3!TIPODOC) = "NDA" Then
                        
                        Print #1, "1" & fec & "02" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000") & Format(rs3!NroFactura, "00000000") & "001" & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & Chr(32) & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) '& Chr(13)
                        netgra = netgra + CDbl(rs3!Neto)
                        i = i + 1
                    ElseIf Trim(rs3!TIPODOC) = "NCB" Then
                        
                        Print #1, "1" & fec & "08" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000") & Format(rs3!NroFactura, "00000000") & "001" & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & "E" & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) '& Chr(13)
                        netNOgra = netNOgra - CDbl(rs3!Total)
                        i = i + 1
                    ElseIf Trim(rs3!TIPODOC) = "NDB" Then
                        
                        Print #1, "1" & fec & "07" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000") & Format(rs3!NroFactura, "00000000") & "001" & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & "E" & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) '& Chr(13)
                        netNOgra = netNOgra + CDbl(rs3!Total)
                        i = i + 1
                        
                    End If
                    rs3.MoveNext
                    Set rs2 = Nothing
                Wend
                'registro de tipo 2
                j = 1
                While j < 63
                    relleno = relleno & Chr(32)
                    j = j + 1
                Wend
                
                Print #1, "2" & Year(DTPicker1.Value) & Format(Month(DTPicker1.Value), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Format(i, "00000000") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Trim(Mid(rs!CuitEmpresa, 1, 2)) & Trim(Mid(rs!CuitEmpresa, 4, 8)) & Trim(Mid(rs!CuitEmpresa, 13, 1)) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                 Format(tot * 100, "000000000000000") & Format(netNOgra * 100, "000000000000000") & Format(netgra * 100, "000000000000000") & Format(totimpLIQ, "000000000000000") & Format(totimpRNI, "000000000000000") & Format(totimpEXE, "000000000000000") & "000000000000000" & Format(IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & relleno

                
                If Not IsNull(rs3) Or Not IsEmpty(rs3) Or Not (rs3.EOF = True And rs3.BOF = True) Then
                    Close #1
                End If
                Set rs3 = Nothing
                Set rs = Nothing
                MsgBox "Archivo generado correctamente."
                
            
    '******************************************************************************************

End Sub

Private Sub Command4_Click() 'DETALLE DE FACTURAS
    Dim ARCH As String
    Dim documento As Integer
    Dim nom As String
    Dim prod As String
    Dim tI As Double
    Dim impLIQ As String
    Dim impRNI As String
    Dim impEXE As String
    Dim totimpLIQ As Double
    Dim totimpRNI As Double
    Dim totimpEXE As Double
    Dim IIBB As Double
    Dim fec As String
    Dim pri As Date
    Dim seg As Date
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim rs3 As New ADODB.Recordset
    Dim rs4 As New ADODB.Recordset
    Dim i As Long
    Dim j As Long
    Dim z As Long
    Dim tot As Double
    Dim netgra As Double
    Dim netNOgra As Double
    Dim rni As Double
    Dim relleno As String
    Dim cant As Double
    Dim PU As Double
    Dim Item As Long
    Dim PT As Double
    
    
    pri = "01/" & Month(DTPicker1.Value) & "/" & Year(DTPicker1.Value)
    seg = ultimoDiaDelMes(DTPicker1.Value)
    rs3.Open "select * from facturaventa where fecha>=" & ssFecha(pri) & " and fecha<=" & ssFecha(seg) & " order by fecha,tipodoc,nrofactura", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'********************************** ARCHIVO
                ARCH = "DETALLE_" & Year(DTPicker1.Value) & Format(Month(DTPicker1.Value), "00")
                rs.Open "select * from datosempresa where idempresa=15", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                
                If (rs3.EOF = True And rs3.BOF = True) Or IsNull(rs3) Then
                    MsgBox "No se ha encontrado datos para este periodo.", , "ATENCION"
                    Exit Sub
                End If
                
                rs3.MoveFirst
                If Not IsNull(rs3) Or Not IsEmpty(rs3) Or Not (rs3.EOF = True And rs3.BOF = True) Then
                    Open "C:\gestion\SIRED\" & ARCH & ".txt" For Output As #1
                End If
                i = 0
                
                'j = 1
                'While j < 201
                '    relleno = relleno & Chr(32)
                '    j = j + 1
                'Wend
                
                While Not rs3.EOF
                
                    rs2.Open "select i.*,C.CUIT from clientes c inner join ivas i on i.codigo=c.iva where c.codigo=" & rs3!cliente, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                    If rs3!tipoiva = 1 Or rs3!tipoiva = 4 Or rs3!tipoiva = 6 Then
                    'If ComboCodigo(cmbTipoIva) = 1 Or ComboCodigo(cmbTipoIva) = 4 Or ComboCodigo(cmbTipoIva) = 6 Then
                        documento = 80
                    Else
                        documento = 86
                    End If
                    nom = Trim(rs3!RAZONSOCIAL)
                    While Len(nom) < 30
                        nom = Chr(32) & nom
                    Wend
                    
                    fec = Year(rs3!Fecha) & Format(Month(rs3!Fecha), "00") & Format(Day(rs3!Fecha), "00")
                    'ver si tengo que multiplicar * 100 los valores para sacarle la coma!!!!!!
                    'tI = Trim(rs3!tipoiva)
                    'If tI = 5 Or tI = 4 Or tI = 6 Or tI = 10 Then 'facturas B
                    '    impLIQ = "000000000000000"
                    '    impRNI = "000000000000000"
                    '    impEXE = "000000000000000" 'Format(TxtIVA * 100, "000000000000000")
                    'ElseIf tI = 1 Or tI = 2 Then  'facturas A
                    '    If tI = 1 Then
                    '        impLIQ = Format(rs3!iva * 100, "000000000000000")
                    '        impRNI = "000000000000000"
                    '        impEXE = "000000000000000"
                    '    ElseIf tI = 2 Then
                    '        impLIQ = "000000000000000"
                    '        impRNI = Format(rs3!iva * 100, "000000000000000")
                    '        impEXE = "000000000000000"
                    '    ElseIf tI = 8 Then
                    '        impLIQ = "000000000000000"
                    '        impRNI = "000000000000000"
                    '        impEXE = Format(rs3!iva * 100, "000000000000000")
                    '    End If
                    'End If
                    
                    'tot = tot + CDbl(rs3!Total)
                    'totimpLIQ = totimpLIQ + CDbl(impLIQ)
                    'totimpRNI = totimpRNI + CDbl(impRNI)
                    'totimpEXE = totimpEXE + CDbl(impEXE)
                    'iibb = iibb + CDbl(rs3!iibb)
                    
                    Item = obtenerDeSQL("select count(distinct(producto)) from facturaventadetalle where tipodoc='" & rs3!TIPODOC & "' and nrofactura=" & rs3!NroFactura)
                    
                    rs4.Open "select distinct(producto) as prod,* from facturaventadetalle where tipodoc='" & rs3!TIPODOC & "' and nrofactura=" & rs3!NroFactura, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                    
                    If Trim(rs3!TIPODOC) = "FAA" Then 'en las tablas de la afip el 01 es factura A
                        'Open "C:\gestion\SIRED\" & arch & ".txt" For Output As #1
                        'Write #1, cmbCliente.Text & cmbCliente.Text & Chr(32) & cmbCliente.Text            ' Text1.Text el chr(9)=tab  vbCrLf=enter   chr(32)=un espacio
                        z = 0
                        While z < Item
                            cant = obtenerDeSQL("select sum(cantidad) from facturaventadetalle where tipodoc='" & Trim(rs3!TIPODOC) & "' and nrofactura=" & rs3!NroFactura & " and producto='" & Trim(rs4!prod) & "'")
                            PU = obtenerDeSQL("select preciounitario from facturaventadetalle where tipodoc='" & Trim(rs3!TIPODOC) & "' and nrofactura=" & rs3!NroFactura & " and producto='" & Trim(rs4!prod) & "'")
                            PT = obtenerDeSQL("select sum(preciototal) from facturaventadetalle where tipodoc='" & Trim(rs3!TIPODOC) & "' and nrofactura=" & rs3!NroFactura & " and producto='" & Trim(rs4!prod) & "'")
                            prod = Trim(rs4!DESCRIPCION)
                            While Len(prod) < 75
                                prod = prod & Chr(32)
                            Wend
                            
                            Print #1, "01" & Chr(32) & fec & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000") & Format(rs3!NroFactura, "00000000") & Format(cant * 100000, "000000000000") & "98" & Format(PU * 1000, "0000000000000000") & "000000000000000" & "0000000000000000" & Format(PT * 1000, "0000000000000000") & "2100" & "G" & Chr(32) & prod '& Chr(13)
                            z = z + 1
                            rs4.MoveNext
                        Wend
                        'Close #1
                        'netgra = netgra + CDbl(rs3!neto)
                        'i = i + 1
                    ElseIf Trim(rs3!TIPODOC) = "FAB" Then
                        'If CDbl(txtTotal.caption) >= CDbl("10000") Then  'lo quieren hacer uno por uno
                            'Open "C:\gestion\SIRED\" & arch & ".txt" For Output As #1
                            'Write #1, cmbCliente.Text & cmbCliente.Text & Chr(32) & cmbCliente.Text            ' Text1.Text el chr(9)=tab  vbCrLf=enter
                        z = 0
                        While z < Item
                            cant = obtenerDeSQL("select sum(cantidad) from facturaventadetalle where tipodoc='" & Trim(rs3!TIPODOC) & "' and nrofactura=" & rs3!NroFactura & " and producto='" & Trim(rs4!prod) & "'")
                            PU = obtenerDeSQL("select preciounitario from facturaventadetalle where tipodoc='" & Trim(rs3!TIPODOC) & "' and nrofactura=" & rs3!NroFactura & " and producto='" & Trim(rs4!prod) & "'")
                            PT = obtenerDeSQL("select preciototal from facturaventadetalle where tipodoc='" & Trim(rs3!TIPODOC) & "' and nrofactura=" & rs3!NroFactura & " and producto='" & Trim(rs4!prod) & "'")
                            prod = Trim(rs4!DESCRIPCION)
                            While Len(prod) < 75
                                prod = prod & Chr(32)
                            Wend
                            
                            Print #1, "06" & Chr(32) & fec & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000") & Format(rs3!NroFactura, "00000000") & Format(cant * 100000, "000000000000") & "98" & Format(PU * 1000, "0000000000000000") & "000000000000000" & "0000000000000000" & Format(PT * 1000, "0000000000000000") & "0000" & "E" & Chr(32) & prod '& Chr(13)
                            z = z + 1
                            rs4.MoveNext
                        Wend
                            'txtneto
                            'netNOgra = netNOgra + CDbl(rs3!Total)
                            'Close #1
                        'End If
                        'i = i + 1
                    End If
                    rs3.MoveNext
                    Set rs2 = Nothing
                    Set rs4 = Nothing
                Wend
                'registro de tipo 2
                
                
                'Print #1, "2" & Year(DTPicker1.Value) & Format(Month(DTPicker1.Value), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Format(i, "000000000000") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Trim(Mid(rs!cuitempresa, 1, 2)) & Trim(Mid(rs!cuitempresa, 4, 8)) & Trim(Mid(rs!cuitempresa, 13, 1)) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                ' Format(tot * 100, "000000000000000") & Format(netNOgra * 100, "000000000000000") & Format(netgra * 100, "000000000000000") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Format(totimpLIQ, "000000000000000") & Format(totimpRNI, "000000000000000") & Format(totimpEXE, "000000000000000") & "000000000000000" & Format(iibb * 100, "000000000000000") & "000000000000000" & "000000000000000" & relleno & "j"

                
                If Not IsNull(rs3) Or Not IsEmpty(rs3) Or Not (rs3.EOF = True And rs3.BOF = True) Then
                    Close #1
                End If
                Set rs3 = Nothing
                Set rs = Nothing
                MsgBox "Archivo generado correctamente."
                
            
    '******************************************************************************************

End Sub

Private Sub Command5_Click() 'OTRAS PERCEPCIONES
    Dim ARCH As String
    Dim documento As Integer
    Dim nom As String
    Dim tI As Double
    Dim impLIQ As String
    Dim impRNI As String
    Dim impEXE As String
    Dim totimpLIQ As Double
    Dim totimpRNI As Double
    Dim totimpEXE As Double
    Dim IIBB As Double
    Dim fec As String
    Dim pri As Date
    Dim seg As Date
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim rs3 As New ADODB.Recordset
    Dim i As Long
    Dim j As Long
    Dim tot As Double
    Dim netgra As Double
    Dim netNOgra As Double
    Dim rni As Double
    Dim relleno As String
    
    
    pri = "01/" & Month(DTPicker1.Value) & "/" & Year(DTPicker1.Value)
    seg = ultimoDiaDelMes(DTPicker1.Value)
    rs3.Open "select * from facturaventa where fecha>=" & ssFecha(pri) & " and fecha<=" & ssFecha(seg) & " and iibb<>0 order by fecha,tipodoc,nrofactura", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'********************************** ARCHIVO
                ARCH = "OTRAS_PERCEP_" & Year(DTPicker1.Value) & Format(Month(DTPicker1.Value), "00")
                rs.Open "select * from datosempresa where idempresa=15", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                
                If (rs3.EOF = True And rs3.BOF = True) Or IsNull(rs3) Then
                    MsgBox "No se ha encontrado datos para este periodo.", , "ATENCION"
                    Exit Sub
                End If
                
                rs3.MoveFirst
                If Not IsNull(rs3) Or Not IsEmpty(rs3) Or Not (rs3.EOF = True And rs3.BOF = True) Then
                    Open "C:\gestion\SIRED\" & ARCH & ".txt" For Output As #1
                End If
                i = 0
                
                j = 1
                While j < 41
                    relleno = relleno & Chr(32)
                    j = j + 1
                Wend
                
                While Not rs3.EOF
                
                    rs2.Open "select i.*,C.CUIT from clientes c inner join ivas i on i.codigo=c.iva where c.codigo=" & rs3!cliente, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                    If rs3!tipoiva = 1 Or rs3!tipoiva = 4 Or rs3!tipoiva = 6 Then
                    'If ComboCodigo(cmbTipoIva) = 1 Or ComboCodigo(cmbTipoIva) = 4 Or ComboCodigo(cmbTipoIva) = 6 Then
                        documento = 80
                    Else
                        documento = 86
                    End If
                    'nom = Trim(rs3!RAZONSOCIAL)
                    'While Len(nom) < 30
                    '    nom = Chr(32) & nom
                    'Wend
                    
                    fec = Year(rs3!Fecha) & Format(Month(rs3!Fecha), "00") & Format(Day(rs3!Fecha), "00")
                    'ver si tengo que multiplicar * 100 los valores para sacarle la coma!!!!!!
                    'tI = Trim(rs3!tipoiva)
                    'If tI = 5 Or tI = 4 Or tI = 6 Or tI = 10 Then 'facturas B
                    '    impLIQ = "000000000000000"
                    '    impRNI = "000000000000000"
                    '    impEXE = "000000000000000" 'Format(TxtIVA * 100, "000000000000000")
                    'ElseIf tI = 1 Or tI = 2 Then  'facturas A
                    '    If tI = 1 Then
                    '        impLIQ = Format(rs3!iva * 100, "000000000000000")
                    '        impRNI = "000000000000000"
                    '        impEXE = "000000000000000"
                    '    ElseIf tI = 2 Then
                    '        impLIQ = "000000000000000"
                    '        impRNI = Format(rs3!iva * 100, "000000000000000")
                    '        impEXE = "000000000000000"
                    '    ElseIf tI = 8 Then
                    '        impLIQ = "000000000000000"
                    '        impRNI = "000000000000000"
                    '        impEXE = Format(rs3!iva * 100, "000000000000000")
                    '    End If
                    'End If
                    
                    'tot = tot + CDbl(rs3!Total)
                    'totimpLIQ = totimpLIQ + CDbl(impLIQ)
                    'totimpRNI = totimpRNI + CDbl(impRNI)
                    'totimpEXE = totimpEXE + CDbl(impEXE)
                    'iibb = iibb + CDbl(rs3!iibb)
                    
                    
                    If Trim(rs3!TIPODOC) = "FAA" Then 'en las tablas de la afip el 01 es factura A
                        'Open "C:\gestion\SIRED\" & arch & ".txt" For Output As #1
                        'Write #1, cmbCliente.Text & cmbCliente.Text & Chr(32) & cmbCliente.Text            ' Text1.Text el chr(9)=tab  vbCrLf=enter   chr(32)=un espacio
                        Print #1, fec & "01" & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000") & provi(rs3!Provincia) & Format(rs3!IIBB * 100, "000000000000000") & relleno & "000000000000000" '& Chr(13)
                        'Close #1
                        netgra = netgra + CDbl(rs3!Neto)
                        i = i + 1
                    ElseIf Trim(rs3!TIPODOC) = "FAB" Then
                        'If CDbl(txtTotal.caption) >= CDbl("10000") Then  'lo quieren hacer uno por uno
                            'Open "C:\gestion\SIRED\" & arch & ".txt" For Output As #1
                            'Write #1, cmbCliente.Text & cmbCliente.Text & Chr(32) & cmbCliente.Text            ' Text1.Text el chr(9)=tab  vbCrLf=enter
                            Print #1, fec & "06" & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000") & provi(rs3!Provincia) & Format(rs3!IIBB * 100, "000000000000000") & relleno & "000000000000000" '& Chr(13)
                            'txtneto
                            netNOgra = netNOgra + CDbl(rs3!Total)
                            'Close #1
                        'End If
                        i = i + 1
                    End If
                    rs3.MoveNext
                    Set rs2 = Nothing
                Wend
                'registro de tipo 2
                
                
                'Print #1, "2" & Year(DTPicker1.Value) & Format(Month(DTPicker1.Value), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Format(i, "00000000") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Trim(Mid(rs!cuitempresa, 1, 2)) & Trim(Mid(rs!cuitempresa, 4, 8)) & Trim(Mid(rs!cuitempresa, 13, 1)) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                ' Format(tot * 100, "000000000000000") & Format(netNOgra * 100, "000000000000000") & Format(netgra * 100, "000000000000000") & Format(totimpLIQ, "000000000000000") & Format(totimpRNI, "000000000000000") & Format(totimpEXE, "000000000000000") & "000000000000000" & Format(iibb * 100, "000000000000000") & "000000000000000" & "000000000000000" & relleno

                
                If Not IsNull(rs3) Or Not IsEmpty(rs3) Or Not (rs3.EOF = True And rs3.BOF = True) Then
                    Close #1
                End If
                Set rs3 = Nothing
                Set rs = Nothing
                MsgBox "Archivo generado correctamente."
                
            
    '******************************************************************************************

End Sub

Private Sub Command6_Click()
    Unload Me
End Sub

Public Function ultimoDiaDelMes(Fecha As Date) As Date
ultimoDiaDelMes = DateAdd("m", 1, Fecha)
ultimoDiaDelMes = DateSerial(Year(ultimoDiaDelMes), Month(ultimoDiaDelMes), 1)
ultimoDiaDelMes = DateAdd("d", -1, ultimoDiaDelMes)
End Function

Private Function provi(p As String) As String
    Select Case p
        Case "*":
            provi = "00"
        Case "B":
            provi = "01"
        Case "S":
            provi = "12"
        Case "Z":
            provi = "23"
        Case "K":
            provi = "02"
        Case "H":
            provi = "16"
        Case "U":
            provi = "17"
        Case "X":
            provi = "03"
        Case "W":
            provi = "04"
        Case "E":
            provi = "05"
        Case "P":
            provi = "18"
        Case "Y":
            provi = "06"
        Case "L":
            provi = "21"
        Case "F":
            provi = "08"
        Case "M":
            provi = "07"
        Case "N":
            provi = "19"
        Case "Q":
            provi = "20"
        Case "R":
            provi = "22"
        Case "A":
            provi = "09"
        Case "J":
            provi = "10"
        Case "D":
            provi = "11"
        Case "G":
            provi = "13"
        Case "V":
            provi = "24"
        Case "T":
            provi = "14"
    End Select
End Function

Private Sub Command7_Click()
    Dim ARCH As String
    Dim documento As Integer
    Dim nom As String
    Dim tI As Double
    Dim impLIQ As String
    Dim impLIQ2 As Double
    Dim impRNI As String
    Dim impEXE As String
    Dim totimpLIQ As Double
    Dim totimpRNI As Double
    Dim totimpEXE As Double
    Dim IIBB As Double
    Dim fec As String
    Dim pri As Date
    Dim seg As Date
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim rs3 As New ADODB.Recordset
    Dim rs4 As New ADODB.Recordset
    Dim i As Long
    Dim j As Long
    Dim tot As Double
    Dim netgra As Double
    Dim netNOgra As Double
    Dim rni As Double
    Dim relleno As String
    Dim relleno2 As String
    Dim str As String
    Dim cantIVA As Long
    Dim Iva As String ' Double
    Dim Neto As Double
    Dim tipo As String
    
    
    pri = "01/" & Month(DTPicker1.Value) & "/" & Year(DTPicker1.Value)
    seg = ultimoDiaDelMes(DTPicker1.Value)
    str = "select * from facturaventa where fecha>=" & ssFecha(pri) & " and fecha<=" & ssFecha(seg) & " and (tipodoc='FAA' or tipodoc='FAB' or tipodoc='NCA' or tipodoc='NCB') order by fecha,tipodoc,nrofactura"
    rs3.Open str, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    
'********************************** ARCHIVO

    ARCH = "VENTAS_" & Year(DTPicker1.Value) & Format(Month(DTPicker1.Value), "00")
    rs.Open "select * from datosempresa where idempresa=4", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    
    If (rs3.EOF = True And rs3.BOF = True) Or IsNull(rs3!codigo) Then
        MsgBox "No se ha encontrado datos para este periodo.", , "ATENCION"
        Exit Sub
    End If
    
    rs3.MoveFirst
    If Not IsNull(rs3) Or Not IsEmpty(rs3) Or Not (rs3.EOF = True And rs3.BOF = True) Then
        Call Verificar_Existe("C:\CITI\") '& ARCH & ".txt"
        
        Open "C:\CITI\" & ARCH & ".txt" For Output As #1
    End If
    i = 0
    While Not rs3.EOF
        If rs3!activo = False Then 'si esta anulado va todo en cero
            fec = Year(rs3!Fecha) & Format(Month(rs3!Fecha), "00") & Format(Day(rs3!Fecha), "00")
            If Trim(rs3!TIPODOC) = "FAA" Then
                tipo = "01"
            ElseIf Trim(rs3!TIPODOC) = "FAB" Then
                tipo = "06"
            ElseIf Trim(rs3!TIPODOC) = "NCA" Then
                tipo = "03"
            ElseIf Trim(rs3!TIPODOC) = "NCB" Then
                tipo = "08"
            ElseIf Trim(rs3!TIPODOC) = "NDA" Then
                tipo = "02"
            ElseIf Trim(rs3!TIPODOC) = "NDB" Then
                tipo = "07"
            End If
            If rs3!tipoiva = 2 Or rs3!tipoiva = 4 Or rs3!tipoiva = 7 Then 'revisar
                documento = 80 'cuit
            ElseIf rs3!tipoiva = 1 Then
                documento = IIf(Mid(sSinNull(rs3!CUIT), 1, 4) = "00-0", 96, 86) '96=dni,86=cuil
            Else
                documento = 86
            End If
            nom = Mid(Trim(rs3!RAZONSOCIAL), 1, 30)
            While Len(nom) < 30
                nom = Chr(32) & nom
            Wend
            impLIQ = "000000000000000"
            impRNI = "000000000000000"
            impEXE = "000000000000000"
            cantIVA = 1
            relleno2 = ""
            j = 1
            While j <= 75
                relleno2 = relleno2 & Chr(32)
                j = j + 1
            Wend
                        
            Print #1, "1" & fec & tipo & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim("00000000000") & nom & Format(0, "000000000000000") & "000000000000000" & Format(0, "000000000000000") & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & "000000000000000" & "000000000000000" & "000000000000000" & "00" & "PES" & "0000000000" & cantIVA & Chr(32) & "00000000000000" & "00000000" & "00000000" & _
                            relleno2 & "00000000" & "000000000000000"
        Else
            cantIVA = obtenerDeSQL("select count(distinct(_iva)) from facturaventadetalle where codigofactura=" & rs3!codigo)
            If cantIVA > 1 Then
                
                rs4.Open "select *,_iva as iva from facturaventadetalle where codigofactura=" & rs3!codigo & " order by _iva", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        
                rs2.Open "select i.*,C.CUIT from clientes c inner join ivas i on i.codigo=c.iva where c.codigo=" & rs3!cliente, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                If rs3!tipoiva = 2 Or rs3!tipoiva = 4 Or rs3!tipoiva = 7 Then 'revisar
                    documento = 80 'cuit
                ElseIf rs3!tipoiva = 1 Then
                    documento = IIf(Mid(sSinNull(rs3!CUIT), 1, 4) = "00-0", 96, 86) '96=dni,86=cuil
                Else
                    documento = 86
                End If
                nom = Mid(Trim(rs3!RAZONSOCIAL), 1, 30)
                While Len(nom) < 30
                    nom = Chr(32) & nom
                Wend
                
                fec = Year(rs3!Fecha) & Format(Month(rs3!Fecha), "00") & Format(Day(rs3!Fecha), "00")
                tI = Trim(rs3!tipoiva)
                                        
                If tI = 1 Or tI = 4 Or tI = 7 Or tI = 8 Then 'facturas B
                    impLIQ = "000000000000000"
                    impRNI = "000000000000000"
                    impEXE = "000000000000000"
                ElseIf tI = 2 Or tI = 3 Then  'facturas A
                    If tI = 2 Then
                        impLIQ = Format(rs3!Iva * 100, "000000000000000")
                        impRNI = "000000000000000"
                        impEXE = "000000000000000"
                    ElseIf tI = 3 Then
                        impLIQ = "000000000000000"
                        impRNI = Format(rs3!Iva * 100, "000000000000000")
                        impEXE = "000000000000000"
                    ElseIf tI = 8 Then ' a este no entra nunca, pero asi estaba en tavi...
                        impLIQ = "000000000000000"
                        impRNI = "000000000000000"
                        impEXE = Format(rs3!Iva * 100, "000000000000000")
                    End If
                End If
                
                If Trim(rs3!TIPODOC) = "FAA" Or Trim(rs3!TIPODOC) = "FAB" Or Trim(rs3!TIPODOC) = "NDA" Or Trim(rs3!TIPODOC) = "NDB" Then
                    tot = tot + CDbl(rs3!Total)
                    totimpLIQ = totimpLIQ + CDbl(impLIQ)
                    totimpRNI = totimpRNI + CDbl(impRNI)
                    totimpEXE = totimpEXE + CDbl(impEXE)
                    IIBB = IIBB + CDbl(rs3!IIBB)
                ElseIf Trim(rs3!TIPODOC) = "NCA" Or Trim(rs3!TIPODOC) = "NCB" Then
                    tot = tot - CDbl(rs3!Total)
                    totimpLIQ = totimpLIQ - CDbl(impLIQ)
                    totimpRNI = totimpRNI - CDbl(impRNI)
                    totimpEXE = totimpEXE - CDbl(impEXE)
                    IIBB = IIBB - CDbl(rs3!IIBB)
                End If
                
                '***************************
                If Trim(rs3!TIPODOC) = "FAA" Then
                    tipo = "01"
                ElseIf Trim(rs3!TIPODOC) = "FAB" Then
                    tipo = "06"
                ElseIf Trim(rs3!TIPODOC) = "NCA" Then
                    tipo = "03"
                ElseIf Trim(rs3!TIPODOC) = "NCB" Then
                    tipo = "08"
                ElseIf Trim(rs3!TIPODOC) = "NDA" Then
                    tipo = "02"
                ElseIf Trim(rs3!TIPODOC) = "NDB" Then
                    tipo = "07"
                End If
                Do While Not rs4.EOF
                    If tI = 1 Or tI = 4 Or tI = 7 Or tI = 8 Then 'facturas B
                        'impLIQ = Format(s2n(rs4!PrecioTotal * 1 + (rs4!Iva / 100)) * 100, "000000000000000")
                        impLIQ = Format(s2n(rs4!PrecioTotal - (rs4!PrecioTotal / (1 + (rs4!Iva / 100)))) * 100, "000000000000000")
                        Iva = Format(rs4!Iva * 100, "0000")
                        Neto = rs4!PrecioTotal - s2n(rs4!PrecioTotal - (rs4!PrecioTotal / (1 + (rs4!Iva / 100))))
                    ElseIf tI = 2 Then
                        impLIQ = Format(s2n(rs4!PrecioTotal * (rs4!Iva / 100)) * 100, "000000000000000")
                        Iva = Format(rs4!Iva * 100, "0000")
                        Neto = rs4!PrecioTotal
                    End If
                    If rs4.AbsolutePosition <> rs4.RecordCount Then
                        relleno2 = ""
                        j = 1
                        While j <= 75
                            relleno2 = relleno2 & Chr(32)
                            j = j + 1
                        Wend
                        'Print #1, "1" & fec & tipo & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(0, "000000000000000") & "000000000000000" & Format(rs4!PrecioTotal * 100, "000000000000000") & (rs4!Iva * 100) & impLIQ & impRNI & impEXE & "000000000000000" & "000000000000000" & "000000000000000" & "000000000000000" & "00" & Chr(32) & Chr(32) & Chr(32) & "0000000000" & cantIVA & Chr(32) & "00000000000000" & "00000000" & "00000000" & _
                            relleno2 & "00000000" & "000000000000000"
                        Print #1, "1" & fec & tipo & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(0, "000000000000000") & "000000000000000" & Format(Neto * 100, "000000000000000") & Format((rs4!Iva * 100), "0000") & impLIQ & impRNI & impEXE & "000000000000000" & "000000000000000" & "000000000000000" & "000000000000000" & "00" & Chr(32) & Chr(32) & Chr(32) & "0000000000" & cantIVA & Chr(32) & "00000000000000" & "00000000" & "00000000" & _
                            relleno2 & "00000000" & "000000000000000"
                        rs4.MoveNext
                    Else
                        Exit Do
                    End If
                Loop
                Set rs4 = Nothing
                '*********************************************
                
                If Trim(rs3!TIPODOC) = "FAA" Then  'en las tablas de la afip el 01 es factura A
                    relleno2 = ""
                    j = 1
                    While j <= 75
                        relleno2 = relleno2 & Chr(32)
                        j = j + 1
                    Wend
        ''            Print #1, "1" & fec & "01" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & "2100" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & Chr(32) & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                    "Sin comentarios" & relleno2
    '''                Print #1, "1" & fec & "01" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & "2100" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0000000000" & "1" & Chr(32) & Format(nSinNull(rs3!CAE), "00000000000000") & Year(rs3!CAEV) & Format(Month(rs3!CAEV), "00") & Format(Day(rs3!CAEV), "00") & "00000000" & _
                    relleno2 & "00000000" & "000000000000000"
                    Print #1, "1" & fec & "01" & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(Neto * 100, "000000000000000") & Format(Iva, "0000") & impLIQ & impRNI & impEXE & "000000000000000" & "000000000000000" & "000000000000000" & "000000000000000" & "00" & Chr(32) & Chr(32) & Chr(32) & "0000000000" & cantIVA & Chr(32) & "00000000000000" & "00000000" & "00000000" & _
                    relleno2 & "00000000" & "000000000000000"
                    'Close #1
                    netgra = netgra + CDbl(rs3!Neto)
                    i = i + 1
                ElseIf Trim(rs3!TIPODOC) = "FAB" Then
                    'If CDbl(txtTotal.caption) >= CDbl("10000") Then  'lo quieren hacer uno por uno
                        'Open "C:\gestion\SIRED\" & arch & ".txt" For Output As #1
                        'Write #1, cmbCliente.Text & cmbCliente.Text & Chr(32) & cmbCliente.Text            ' Text1.Text el chr(9)=tab  vbCrLf=enter
                        relleno2 = ""
                        j = 1
                        While j <= 75
                            relleno2 = relleno2 & Chr(32)
                            j = j + 1
                        Wend
        ''                Print #1, "1" & fec & "06" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & "E" & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                        "Sin comentarios" & relleno2 '& vbCrLf
    '''                    Print #1, "1" & fec & "06" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0000000000" & "1" & Chr(32) & Format(nSinNull(rs3!CAE), "00000000000000") & Year(rs3!CAEV) & Format(Month(rs3!CAEV), "00") & Format(Day(rs3!CAEV), "00") & "00000000" & _
                        relleno2 & "00000000" & "000000000000000"
                        Print #1, "1" & fec & "06" & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(0, "000000000000000") & Format(Neto * 100, "000000000000000") & Format(Iva, "0000") & impLIQ & impRNI & impEXE & "000000000000000" & "000000000000000" & "000000000000000" & "000000000000000" & "00" & Chr(32) & Chr(32) & Chr(32) & "0000000000" & cantIVA & Chr(32) & "00000000000000" & "00000000" & "00000000" & _
                        relleno2 & "00000000" & "000000000000000"
                        'txtneto
                        netNOgra = netNOgra + CDbl(rs3!Total)
                        'Close #1
                    'End If
                    i = i + 1
                    Set rs4 = Nothing
                ElseIf Trim(rs3!TIPODOC) = "NCA" Then
                    relleno2 = ""
                    j = 1
                    While j <= 75
                        relleno2 = relleno2 & Chr(32)
                        j = j + 1
                    Wend
        ''            Print #1, "1" & fec & "03" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & "2100" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & Chr(32) & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                    "Sin comentarios" & relleno2
    '''                Print #1, "1" & fec & "03" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & "2100" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0000000000" & "1" & Chr(32) & Format(nSinNull(rs3!CAE), "00000000000000") & Year(rs3!CAEV) & Format(Month(rs3!CAEV), "00") & Format(Day(rs3!CAEV), "00") & "00000000" & _
                    relleno2 & "00000000" & "000000000000000"
                    Print #1, "1" & fec & "03" & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(Neto * 100, "000000000000000") & Format(Iva, "0000") & impLIQ & impRNI & impEXE & "000000000000000" & "000000000000000" & "000000000000000" & "000000000000000" & "00" & Chr(32) & Chr(32) & Chr(32) & "0000000000" & cantIVA & Chr(32) & "00000000000000" & "00000000" & "00000000" & _
                    relleno2 & "00000000" & "000000000000000"
                    netgra = netgra - CDbl(rs3!Neto)
                    i = i + 1
                ElseIf Trim(rs3!TIPODOC) = "NDA" Then
                    relleno2 = ""
                    j = 1
                    While j <= 75
                        relleno2 = relleno2 & Chr(32)
                        j = j + 1
                    Wend
        ''            Print #1, "1" & fec & "02" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & "2100" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & Chr(32) & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                    "Sin comentarios" & relleno2
    '''                Print #1, "1" & fec & "02" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & "2100" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0000000000" & "1" & Chr(32) & Format(nSinNull(rs3!CAE), "00000000000000") & Year(rs3!CAEV) & Format(Month(rs3!CAEV), "00") & Format(Day(rs3!CAEV), "00") & "00000000" & _
                    relleno2 & "00000000" & "000000000000000"
                    Print #1, "1" & fec & "02" & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(Neto * 100, "000000000000000") & Format(Iva, "0000") & impLIQ & impRNI & impEXE & "000000000000000" & "000000000000000" & "000000000000000" & "000000000000000" & "00" & Chr(32) & Chr(32) & Chr(32) & "0000000000" & cantIVA & Chr(32) & "00000000000000" & "00000000" & "00000000" & _
                    relleno2 & "00000000" & "000000000000000"
                    netgra = netgra + CDbl(rs3!Neto)
                    i = i + 1
                ElseIf Trim(rs3!TIPODOC) = "NCB" Then
                    relleno2 = ""
                    j = 1
                    While j <= 75
                        relleno2 = relleno2 & Chr(32)
                        j = j + 1
                    Wend
        ''            Print #1, "1" & fec & "08" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & "E" & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                        "Sin comentarios" & relleno2 '& vbCrLf
    '''                Print #1, "1" & fec & "08" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0000000000" & "1" & Chr(32) & Format(nSinNull(rs3!CAE), "00000000000000") & Year(rs3!CAEV) & Format(Month(rs3!CAEV), "00") & Format(Day(rs3!CAEV), "00") & "00000000" & _
                        relleno2 & "00000000" & "000000000000000"
                    Print #1, "1" & fec & "08" & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(Neto * 100, "000000000000000") & Format(Iva, "0000") & impLIQ & impRNI & impEXE & "000000000000000" & "000000000000000" & "000000000000000" & "000000000000000" & "00" & Chr(32) & Chr(32) & Chr(32) & "0000000000" & cantIVA & Chr(32) & "00000000000000" & "00000000" & "00000000" & _
                        relleno2 & "00000000" & "000000000000000"
                    netNOgra = netNOgra - CDbl(rs3!Total)
                    i = i + 1
                ElseIf Trim(rs3!TIPODOC) = "NDB" Then
                    relleno2 = ""
                    j = 1
                    While j <= 75
                        relleno2 = relleno2 & Chr(32)
                        j = j + 1
                    Wend
        ''            Print #1, "1" & fec & "07" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & "E" & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                        "Sin comentarios" & relleno2 '& vbCrLf
    '''                Print #1, "1" & fec & "07" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0000000000" & "1" & Chr(32) & Format(nSinNull(rs3!CAE), "00000000000000") & Year(rs3!CAEV) & Format(Month(rs3!CAEV), "00") & Format(Day(rs3!CAEV), "00") & "00000000" & _
                        relleno2 & "00000000" & "000000000000000"
                    Print #1, "1" & fec & "07" & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(Neto * 100, "000000000000000") & Format(Iva, "0000") & impLIQ & impRNI & impEXE & "000000000000000" & "000000000000000" & "000000000000000" & "000000000000000" & "00" & Chr(32) & Chr(32) & Chr(32) & "0000000000" & cantIVA & Chr(32) & "00000000000000" & "00000000" & "00000000" & _
                        relleno2 & "00000000" & "000000000000000"
                    netNOgra = netNOgra + CDbl(rs3!Total)
                    i = i + 1
                End If
                k = k + 1
            Else
                
                rs4.Open "select *,_iva as iva from facturaventadetalle where codigofactura=" & rs3!codigo & " order by _iva", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                
                rs2.Open "select i.*,C.CUIT from clientes c inner join ivas i on i.codigo=c.iva where c.codigo=" & rs3!cliente, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                If rs3!tipoiva = 2 Or rs3!tipoiva = 4 Or rs3!tipoiva = 7 Then 'revisar
                    documento = 80
                ElseIf rs3!tipoiva = 1 Then
                    documento = IIf(Mid(sSinNull(rs3!CUIT), 1, 4) = "00-0", 96, 86) '96=dni,86=cuil
                Else
                    documento = 86
                End If
                nom = Mid(Trim(rs3!RAZONSOCIAL), 1, 30)
                While Len(nom) < 30
                    nom = Chr(32) & nom
                Wend
                
                fec = Year(rs3!Fecha) & Format(Month(rs3!Fecha), "00") & Format(Day(rs3!Fecha), "00")
                tI = Trim(rs3!tipoiva)
                If tI = 1 Or tI = 4 Or tI = 7 Or tI = 8 Then 'facturas B
                    impLIQ = "000000000000000"
                    impRNI = "000000000000000"
                    impEXE = "000000000000000"
                ElseIf tI = 2 Or tI = 3 Then  'facturas A
                    If tI = 2 Then
                        impLIQ = Format(rs3!Iva * 100, "000000000000000")
                        impRNI = "000000000000000"
                        impEXE = "000000000000000"
                    ElseIf tI = 3 Then
                        impLIQ = "000000000000000"
                        impRNI = Format(rs3!Iva * 100, "000000000000000")
                        impEXE = "000000000000000"
                    ElseIf tI = 8 Then ' a este no entra nunca, pero asi estaba en tavi...
                        impLIQ = "000000000000000"
                        impRNI = "000000000000000"
                        impEXE = Format(rs3!Iva * 100, "000000000000000")
                    End If
                End If
                
                If Trim(rs3!TIPODOC) = "FAA" Or Trim(rs3!TIPODOC) = "FAB" Or Trim(rs3!TIPODOC) = "NDA" Or Trim(rs3!TIPODOC) = "NDB" Then
                    tot = tot + CDbl(rs3!Total)
                    totimpLIQ = totimpLIQ + CDbl(impLIQ)
                    totimpRNI = totimpRNI + CDbl(impRNI)
                    totimpEXE = totimpEXE + CDbl(impEXE)
                    IIBB = IIBB + CDbl(rs3!IIBB)
                ElseIf Trim(rs3!TIPODOC) = "NCA" Or Trim(rs3!TIPODOC) = "NCB" Then
                    tot = tot - CDbl(rs3!Total)
                    totimpLIQ = totimpLIQ - CDbl(impLIQ)
                    totimpRNI = totimpRNI - CDbl(impRNI)
                    totimpEXE = totimpEXE - CDbl(impEXE)
                    IIBB = IIBB - CDbl(rs3!IIBB)
                End If
                '***************************
                If Trim(rs3!TIPODOC) = "FAA" Then
                    tipo = "01"
                ElseIf Trim(rs3!TIPODOC) = "FAB" Then
                    tipo = "06"
                ElseIf Trim(rs3!TIPODOC) = "NCA" Then
                    tipo = "03"
                ElseIf Trim(rs3!TIPODOC) = "NCB" Then
                    tipo = "08"
                ElseIf Trim(rs3!TIPODOC) = "NDA" Then
                    tipo = "02"
                ElseIf Trim(rs3!TIPODOC) = "NDB" Then
                    tipo = "07"
                End If
                impLIQ2 = 0
                Neto = 0
                Do While Not rs4.EOF
                    If tI = 1 Or tI = 4 Or tI = 7 Or tI = 8 Then 'facturas B
                        impLIQ2 = impLIQ2 + s2n(rs4!PrecioTotal - (rs4!PrecioTotal / (1 + (rs4!Iva / 100)))) * 100
                        impLIQ = Format(s2n(impLIQ2), "000000000000000")
                        'impLIQ = Format(s2n(rs4!PrecioTotal - (rs4!PrecioTotal / (1 + (rs4!Iva / 100)))) * 100, "000000000000000")
                        Iva = Format(rs4!Iva * 100, "0000")
                        Neto = Neto + rs4!PrecioTotal - s2n(rs4!PrecioTotal - (rs4!PrecioTotal / (1 + (rs4!Iva / 100))))
                    ElseIf tI = 2 Then
                        impLIQ2 = impLIQ2 + s2n(rs4!PrecioTotal * (rs4!Iva / 100)) * 100
                        impLIQ = Format(s2n(impLIQ), "000000000000000")
                        'impLIQ = Format(s2n(rs4!PrecioTotal * (rs4!Iva / 100)) * 100, "000000000000000")
                        Iva = Format(rs4!Iva * 100, "0000")
                        Neto = Neto + rs4!PrecioTotal
                    End If
    '                If rs4.AbsolutePosition <> rs4.RecordCount Then
    '                    relleno2 = ""
    '                    j = 1
    '                    While j <= 75
    '                        relleno2 = relleno2 & Chr(32)
    '                        j = j + 1
    '                    Wend
    '                    'Print #1, "1" & fec & tipo & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(0, "000000000000000") & "000000000000000" & Format(rs4!PrecioTotal * 100, "000000000000000") & (rs4!Iva * 100) & impLIQ & impRNI & impEXE & "000000000000000" & "000000000000000" & "000000000000000" & "000000000000000" & "00" & Chr(32) & Chr(32) & Chr(32) & "0000000000" & cantIVA & Chr(32) & "00000000000000" & "00000000" & "00000000" & _
    '                        relleno2 & "00000000" & "000000000000000"
    '                    Print #1, "1" & fec & tipo & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(0, "000000000000000") & "000000000000000" & Format(Neto * 100, "000000000000000") & (rs4!Iva * 100) & impLIQ & impRNI & impEXE & "000000000000000" & "000000000000000" & "000000000000000" & "000000000000000" & "00" & Chr(32) & Chr(32) & Chr(32) & "0000000000" & cantIVA & Chr(32) & "00000000000000" & "00000000" & "00000000" & _
    '                        relleno2 & "00000000" & "000000000000000"
                        rs4.MoveNext
    '                Else
    '                    Exit Do
    '                End If
                Loop
                Set rs4 = Nothing
                '*********************************************
                
                If Trim(rs3!TIPODOC) = "FAA" Then  'en las tablas de la afip el 01 es factura A
                    relleno2 = ""
                    j = 1
                    While j <= 75
                        relleno2 = relleno2 & Chr(32)
                        j = j + 1
                    Wend
        ''            Print #1, "1" & fec & "01" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & "2100" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & Chr(32) & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                    "Sin comentarios" & relleno2
                    Print #1, "1" & fec & "01" & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & "2100" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0000000000" & "1" & Chr(32) & Format(nSinNull(rs3!CAE), "00000000000000") & Year(nSinNull(rs3!CAEV)) & Format(Month(nSinNull(rs3!CAEV)), "00") & Format(Day(nSinNull(rs3!CAEV)), "00") & "00000000" & _
                    relleno2 & "00000000" & "000000000000000"
                    'Close #1
                    netgra = netgra + CDbl(rs3!Neto)
                    i = i + 1
                ElseIf Trim(rs3!TIPODOC) = "FAB" Then
                    'If CDbl(txtTotal.caption) >= CDbl("10000") Then  'lo quieren hacer uno por uno
                        'Open "C:\gestion\SIRED\" & arch & ".txt" For Output As #1
                        'Write #1, cmbCliente.Text & cmbCliente.Text & Chr(32) & cmbCliente.Text            ' Text1.Text el chr(9)=tab  vbCrLf=enter
                        relleno2 = ""
                        j = 1
                        While j <= 75
                            relleno2 = relleno2 & Chr(32)
                            j = j + 1
                        Wend
        ''                Print #1, "1" & fec & "06" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & "E" & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                        "Sin comentarios" & relleno2 '& vbCrLf
                        'Print #1, "1" & fec & "06" & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0000000000" & "1" & Chr(32) & Format(nSinNull(rs3!CAE), "00000000000000") & Year(nSinNull(rs3!CAEV)) & Format(Month(nSinNull(rs3!CAEV)), "00") & Format(Day(nSinNull(rs3!CAEV)), "00") & "00000000" & _
                            relleno2 & "00000000" & "000000000000000"
                        Print #1, "1" & fec & "06" & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(0, "000000000000000") & Format(Neto * 100, "000000000000000") & Format(Iva, "0000") & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & "00" & Chr(32) & Chr(32) & Chr(32) & "0000000000" & cantIVA & Chr(32) & "00000000000000" & "00000000" & "00000000" & _
                            relleno2 & "00000000" & "000000000000000"
                        'txtneto
                        netNOgra = netNOgra + CDbl(rs3!Total)
                        'Close #1
                    'End If
                    i = i + 1
                ElseIf Trim(rs3!TIPODOC) = "NCA" Then
                    relleno2 = ""
                    j = 1
                    While j <= 75
                        relleno2 = relleno2 & Chr(32)
                        j = j + 1
                    Wend
        ''            Print #1, "1" & fec & "03" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & "2100" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & Chr(32) & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                    "Sin comentarios" & relleno2
                    Print #1, "1" & fec & "03" & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & "2100" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0000000000" & "1" & Chr(32) & Format(nSinNull(rs3!CAE), "00000000000000") & Year(nSinNull(rs3!CAEV)) & Format(Month(nSinNull(rs3!CAEV)), "00") & Format(Day(nSinNull(rs3!CAEV)), "00") & "00000000" & _
                    relleno2 & "00000000" & "000000000000000"
                    netgra = netgra - CDbl(rs3!Neto)
                    i = i + 1
                ElseIf Trim(rs3!TIPODOC) = "NDA" Then
                    relleno2 = ""
                    j = 1
                    While j <= 75
                        relleno2 = relleno2 & Chr(32)
                        j = j + 1
                    Wend
        ''            Print #1, "1" & fec & "02" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & "2100" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & Chr(32) & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                    "Sin comentarios" & relleno2
                    Print #1, "1" & fec & "02" & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & Format(rs3!Neto * 100, "000000000000000") & "2100" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0000000000" & "1" & Chr(32) & Format(nSinNull(rs3!CAE), "00000000000000") & Year(nSinNull(rs3!CAEV)) & Format(Month(nSinNull(rs3!CAEV)), "00") & Format(Day(nSinNull(rs3!CAEV)), "00") & "00000000" & _
                    relleno2 & "00000000" & "000000000000000"
                    netgra = netgra + CDbl(rs3!Neto)
                    i = i + 1
                ElseIf Trim(rs3!TIPODOC) = "NCB" Then
                    relleno2 = ""
                    j = 1
                    While j <= 75
                        relleno2 = relleno2 & Chr(32)
                        j = j + 1
                    Wend
        ''            Print #1, "1" & fec & "08" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & "E" & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                        "Sin comentarios" & relleno2 '& vbCrLf
                    Print #1, "1" & fec & "08" & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0000000000" & "1" & Chr(32) & Format(nSinNull(rs3!CAE), "00000000000000") & Year(nSinNull(rs3!CAEV)) & Format(Month(nSinNull(rs3!CAEV)), "00") & Format(Day(nSinNull(rs3!CAEV)), "00") & "00000000" & _
                        relleno2 & "00000000" & "000000000000000"
                    netNOgra = netNOgra - CDbl(rs3!Total)
                    i = i + 1
                ElseIf Trim(rs3!TIPODOC) = "NDB" Then
                    relleno2 = ""
                    j = 1
                    While j <= 75
                        relleno2 = relleno2 & Chr(32)
                        j = j + 1
                    Wend
        ''            Print #1, "1" & fec & "07" & Chr(32) & Format(rs!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0001000000" & "1" & "E" & Format(rs3!nrocae, "00000000000000") & Year(rs3!vencecae) & Format(Month(rs3!vencecae), "00") & Format(Day(rs3!vencecae), "00") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & _
                        "Sin comentarios" & relleno2 '& vbCrLf
                    Print #1, "1" & fec & "07" & Chr(32) & Format(rs3!PuntoVenta, "0000") & Format(rs3!NroFactura, "00000000000000000000") & Format(rs3!NroFactura, "00000000000000000000") & documento & Trim(Mid(rs3!CUIT, 1, 2)) & Trim(Mid(rs3!CUIT, 4, 8)) & Trim(Mid(rs3!CUIT, 13, 1)) & nom & Format(rs3!Total * 100, "000000000000000") & Format(rs3!Total * 100, "000000000000000") & "000000000000000" & "0000" & impLIQ & impRNI & impEXE & "000000000000000" & Format(rs3!IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & Format(rs3!tipoiva, "00") & "PES" & "0000000000" & "1" & Chr(32) & Format(nSinNull(rs3!CAE), "00000000000000") & Year(nSinNull(rs3!CAEV)) & Format(Month(nSinNull(rs3!CAEV)), "00") & Format(Day(nSinNull(rs3!CAEV)), "00") & "00000000" & _
                        relleno2 & "00000000" & "000000000000000"
                    netNOgra = netNOgra + CDbl(rs3!Total)
                    i = i + 1
                End If
            End If
        End If
        
        rs3.MoveNext
        Set rs2 = Nothing
    Wend
    'registro de tipo 2
    j = 1
    While j < 123
        relleno = relleno & Chr(32)
        j = j + 1
    Wend
    relleno2 = ""
    j = 1
    While j < 30
        relleno2 = relleno2 & Chr(32)
        j = j + 1
    Wend
    
''    Print #1, "2" & Year(DTPicker1.Value) & Format(Month(DTPicker1.Value), "00") & relleno2 & Format(i, "000000000000") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Trim(Mid(rs!CuitEmpresa, 1, 2)) & Trim(Mid(rs!CuitEmpresa, 4, 8)) & Trim(Mid(rs!CuitEmpresa, 13, 1)) & relleno2 & Chr(32) & _
     Format(tot * 100, "000000000000000") & Format(netNOgra * 100, "000000000000000") & Format(netgra * 100, "000000000000000") & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Format(totimpLIQ, "000000000000000") & Format(totimpRNI, "000000000000000") & Format(totimpEXE, "000000000000000") & "000000000000000" & Format(IIBB * 100, "000000000000000") & "000000000000000" & "000000000000000" & relleno
    
    
    If Not IsNull(rs3) Or Not IsEmpty(rs3) Or Not (rs3.EOF = True And rs3.BOF = True) Then
        Close #1
    End If
    Set rs3 = Nothing
    Set rs = Nothing
    MsgBox "Archivo generado correctamente."
                
            
    '******************************************************************************************


End Sub

Private Sub Command8_Click()
    Dim ARCH As String
    Dim nom As String
    Dim impLIQ As String
    Dim fec As String
    Dim pri As Date
    Dim seg As Date
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim rs3 As New ADODB.Recordset
    Dim i As Long
    Dim j As Long
    Dim relleno As String
    Dim cui As String
    Dim tipo As String
        
    pri = "01/" & Month(DTPicker1.Value) & "/" & Year(DTPicker1.Value)
    seg = ultimoDiaDelMes(DTPicker1.Value)
    'rs3.Open "select fecha,codpr,tipoiva,tipodoc,nrodoc,total,neto,exento,iva_21,iva_27,iva_9,iva_10,imp_int,percepc,der_est,ibcapital,ibprovincia,retgan,retganpago,formadepago,razonsocialprov,cuitprov,suc,anoimp,tipocompro from TRANSCOM where fecha>=" & ssFecha(pri) & " and fecha<=" & ssFecha(seg) & " and tipodoc<>'RAC'" & _
        "union select fecha,codpr,tipoiva,tipodoc,nrodoc,total,neto,exento,iva_21,iva_27,iva_9,iva_10,imp_int,percepc,der_est,ibcapital,ibprovincia,retgan,retganpago,formadepago,razonsocialprov,cuitprov,suc,anoimp,tipocompro from compras where fecha>=" & ssFecha(pri) & " and fecha<=" & ssFecha(seg) & " and tipodoc<>'RAC' order by fecha,tipodoc,nrodoc", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
    rs3.Open "select fecha,codpr,nrodoc,iva_21,iva_27,iva_9,iva_10,imp_int,razonsocialprov,cuitprov,suc,anoimp,tipocompro,(iva_21+iva_27+iva_10) as iva " & _
                " from TRANSCOM " & _
                " where fecha>=" & ssFecha(pri) & " and fecha<=" & ssFecha(seg) & " and tipodoc<>'RAC'" & _
                " group by fecha,codpr,nrodoc,iva_21,iva_27,iva_9,iva_10,imp_int,razonsocialprov,cuitprov,suc,anoimp,tipocompro " & _
                " Having (IVA_21 + IVA_27 + iva_10) >= 500 " & _
        "union " & _
            " select fecha,codpr,nrodoc,iva_21,iva_27,iva_9,iva_10,imp_int,razonsocialprov,cuitprov,suc,anoimp,tipocompro,(iva_21+iva_27+iva_10) as iva " & _
                " from compras " & _
                " where fecha>=" & ssFecha(pri) & " and fecha<=" & ssFecha(seg) & " and tipodoc<>'RAC' " & _
                " group by fecha,codpr,nrodoc,iva_21,iva_27,iva_9,iva_10,imp_int,razonsocialprov,cuitprov,suc,anoimp,tipocompro " & _
                " Having (IVA_21 + IVA_27 + iva_10) >= 500 " & _
                " order by fecha,nrodoc", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
'********************************** ARCHIVO
        ARCH = "COMPRAS_" & Year(DTPicker1.Value) & Format(Month(DTPicker1.Value), "00")
        rs.Open "select * from datosempresa where idempresa=4", DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
        
        If (rs3.EOF = True And rs3.BOF = True) Or IsNull(rs3) Then
            MsgBox "No se ha encontrado datos para este periodo.", , "ATENCION"
            Exit Sub
        End If
        
        rs3.MoveFirst
        If Not IsNull(rs3) Or Not IsEmpty(rs3) Or Not (rs3.EOF = True And rs3.BOF = True) Then
            Call Verificar_Existe("C:\CITI\") '& ARCH & ".txt"

            Open "C:\CITI\" & ARCH & ".txt" For Output As #1
        End If
        i = 0
        While Not rs3.EOF
            
                rs2.Open "select i.*,C.CUIT from prov c inner join ivas i on i.codigo=c.tipoiva where c.codigo=" & rs3!CODPR, DataEnvironment1.Sistema, adOpenForwardOnly, adLockReadOnly
                tipo = Format(rs3!tipocompro, "00")
                If rs3!cuitprov = "" Then
                    cui = Format(0, "00000000000")
                Else
                    cui = Trim(Mid(rs3!cuitprov, 1, 2)) & Trim(Mid(rs3!cuitprov, 4, 8)) & Trim(Mid(rs3!cuitprov, 13, 1))
                End If
                nom = Trim(rs3!razonsocialprov)
                While Len(nom) < 25
                    nom = Chr(32) & nom
                Wend
                
                fec = Year(rs3!Fecha) & Format(Month(rs3!Fecha), "00") & Format(Day(rs3!Fecha), "00")
                
                impLIQ = Format((rs3!IVA_21 + rs3!IVA_27 + rs3!iva_10) * 100, "000000000000")
                
                j = 1
                relleno = " "
                While j < 25
                    relleno = relleno & Chr(32)
                    j = j + 1
                Wend
                
                
                Print #1, tipo & Format(rs3!suc, "0000") & Format(rs3!NroDoc, "00000000000000000000") & fec & cui & nom & _
                    impLIQ & "00000000000" & relleno & "000000000000"
    
                rs3.MoveNext
                Set rs2 = Nothing
        Wend
        
        If Not IsNull(rs3) Or Not IsEmpty(rs3) Or Not (rs3.EOF = True And rs3.BOF = True) Then
            Close #1
        End If
        Set rs3 = Nothing
        Set rs = Nothing
        MsgBox "Archivo generado correctamente."
                
            
    '******************************************************************************************

End Sub

Private Sub DTPicker1_Change()
    Label1.caption = "C:\gestion\SIRED\VENTAS_" & Year(DTPicker1.Value) & Format(Month(DTPicker1.Value), "00") & ".txt"
    Label2.caption = "C:\gestion\SIRED\COMPRAS_" & Year(DTPicker1.Value) & Format(Month(DTPicker1.Value), "00") & ".txt"
    Label3.caption = "C:\gestion\SIRED\CABECERA_" & Year(DTPicker1.Value) & Format(Month(DTPicker1.Value), "00") & ".txt"
    Label4.caption = "C:\gestion\SIRED\DETALLE_" & Year(DTPicker1.Value) & Format(Month(DTPicker1.Value), "00") & ".txt"
    Label5.caption = "C:\gestion\SIRED\Otras_Percep_" & Year(DTPicker1.Value) & Format(Month(DTPicker1.Value), "00") & ".txt"
    lblCitiVenta.caption = "C:\CITI\VENTAS_" & Year(DTPicker1.Value) & Format(Month(DTPicker1.Value), "00") & ".txt"
    lblCitiCompra.caption = "C:\CITI\COMPRAS_" & Year(DTPicker1.Value) & Format(Month(DTPicker1.Value), "00") & ".txt"
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Date
    Label1.caption = "C:\gestion\SIRED\VENTAS_" & Year(DTPicker1.Value) & Format(Month(DTPicker1.Value), "00") & ".txt"
    Label2.caption = "C:\gestion\SIRED\COMPRAS_" & Year(DTPicker1.Value) & Format(Month(DTPicker1.Value), "00") & ".txt"
    Label3.caption = "C:\gestion\SIRED\CABECERA_" & Year(DTPicker1.Value) & Format(Month(DTPicker1.Value), "00") & ".txt"
    Label4.caption = "C:\gestion\SIRED\DETALLE_" & Year(DTPicker1.Value) & Format(Month(DTPicker1.Value), "00") & ".txt"
    Label5.caption = "C:\gestion\SIRED\Otras_Percep_" & Year(DTPicker1.Value) & Format(Month(DTPicker1.Value), "00") & ".txt"
    lblCitiVenta.caption = "C:\CITI\VENTAS_" & Year(DTPicker1.Value) & Format(Month(DTPicker1.Value), "00") & ".txt"
    lblCitiCompra.caption = "C:\CITI\COMPRAS_" & Year(DTPicker1.Value) & Format(Month(DTPicker1.Value), "00") & ".txt"
End Sub




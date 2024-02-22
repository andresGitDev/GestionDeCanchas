VERSION 5.00
Begin VB.UserControl ucRetCompras 
   ClientHeight    =   660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10455
   ScaleHeight     =   660
   ScaleWidth      =   10455
   Begin VB.ComboBox cboTipoIB 
      Height          =   315
      Left            =   2055
      TabIndex        =   7
      Top             =   345
      Width           =   2355
   End
   Begin VB.ComboBox cboTipoRetGan 
      Height          =   315
      Left            =   2055
      TabIndex        =   4
      Top             =   30
      Width           =   2355
   End
   Begin VB.TextBox txtRetIB 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   650
      TabIndex        =   3
      Text            =   "0"
      Top             =   330
      Width           =   1335
   End
   Begin VB.TextBox txtRetGan 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   650
      TabIndex        =   2
      Text            =   "0"
      Top             =   30
      Width           =   1350
   End
   Begin VB.Label lblIB 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "----"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4485
      TabIndex        =   8
      Top             =   345
      Width           =   5925
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "----"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4485
      TabIndex        =   6
      Top             =   60
      Width           =   5925
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblRG 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5040
      TabIndex        =   5
      Top             =   45
      Width           =   3555
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ret IIBB"
      Height          =   195
      Index           =   1
      Left            =   1
      TabIndex        =   1
      Top             =   330
      Width           =   780
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ret Gan"
      Height          =   195
      Index           =   0
      Left            =   1
      TabIndex        =   0
      Top             =   60
      Width           =   780
   End
End
Attribute VB_Name = "ucRetCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mRG_MinimoNoImponible     As Double
Private mRG_Coeficiente                  As Double
Private mRG_PagosAnterioresMes  As Double
Private mRG_PagosTotalMes            As Double
Private mRG_PagosRetAnteriores    As Double
Private mRG_TxtFormula            As String
Private mIB_Coef                        As Double
Private mIB_Base                            As Double
Private mIB_TxtFormula                  As String

Private mUltTipoIB As Long
Private mUltTipoGan As Long


Private MinPagoGan As Double
Private Const MINIMOPAGO = 20
Private Const TIPOGAN_H = 1 ' codigos metidos en tabla, 1 DEBE ser Honorarios, los demas tienen base y coef

Private Enum Tipo_IB
    tipoIB_B = 1
    tipoIB_L = 2
End Enum

Private mRetIB As Double
Private mRetGan As Double
Private mRetGanBase As Double
Private mRetGanCoef As Double

Private mCalculado As Boolean
Private mCodProveedor As Long
Private mMontoNeto As Double
Private mMontoNeto2 As Double
Private mFechaPeriodo As Date
Public tieneIIBB As Boolean
Public tieneGAN As Boolean
Private CodProv As Long

Public Event cambio(Total As Double)
'
'*********************************************************

Public Property Let enabled(como As Boolean)
    txtRetGan.enabled = como
    txtRetIB.enabled = como
End Property
Public Property Let retIB(Importe As Double)
    mRetIB = Importe
    txtRetIB = Importe
End Property
Public Property Let retgan(Importe As Double)
    mRetGan = Importe
    txtRetGan = Importe
End Property
Public Property Get retIB() As Double
    retIB = mRetIB
End Property
Public Property Get retgan() As Double
    retgan = mRetGan
End Property
Public Property Get CuentaIB() As String
    CuentaIB = CuentaParam(ID_Cuenta_P_RET_IB_3ros)
End Property
Public Property Get CuentaGan() As String
    CuentaGan = CuentaParam(ID_Cuenta_P_RET_GAN_3ros)
End Property

Public Property Get RG_MinimoNoImponible()
RG_MinimoNoImponible = mRG_MinimoNoImponible
End Property
Public Property Get RG_Coeficiente()
RG_Coeficiente = mRG_Coeficiente
End Property
Public Property Get RG_PagosAnterioresMes()
RG_PagosAnterioresMes = mRG_PagosAnterioresMes
End Property
Public Property Get RG_PagosTotalMes()
RG_PagosTotalMes = mRG_PagosTotalMes
End Property
Public Property Get RG_PagosRetAnteriores()
RG_PagosRetAnteriores = mRG_PagosRetAnteriores
End Property
Public Property Get RG_TxtFormula()
RG_TxtFormula = mRG_TxtFormula
End Property
Public Property Get IB_Coef()
IB_Coef = mIB_Coef
End Property
Public Property Get IB_base()
IB_base = mIB_Base
End Property
Public Property Get IB_TxtFormula()
IB_TxtFormula = mIB_TxtFormula
End Property
Public Property Get IB_Tipo() As Variant
IB_Tipo = cboTipoIB.Text
End Property
Public Property Get IG_Tipo() As Variant
IG_Tipo = cboTipoRetGan.Text
End Property
Public Property Get IG_CodTipo() As Long
    IG_CodTipo = ComboCodigo(cboTipoRetGan)
End Property
Public Property Get IB_CodTipo() As Long
    IB_CodTipo = ComboCodigo(cboTipoIB)
End Property


'Public Function borrar()
'    mCalculado = False
'
'
'End Function

Public Function Calcular(CodProveedor As Long, montoNetoIB As Double, montoNetoG As Double, fechaperiodo As Date, Optional pTipoGan, Optional pTipoIB)

    If Not txtRetGan.enabled Then Exit Function
    
    'recalculo si cambia algo
    If mCodProveedor <> CodProveedor _
                    Or mMontoNeto <> montoNetoG _
                    Or mMontoNeto2 <> montoNetoIB _
                    Or mFechaPeriodo <> fechaperiodo _
                    Or mUltTipoGan <> cboTipoRetGan.ListIndex _
                    Or mUltTipoIB <> cboTipoIB.ListIndex _
                    Or Not IsMissing(pTipoGan) _
                    Or Not IsMissing(pTipoIB) _
                    Then
    
        'reseteo por las dudas
        mRG_MinimoNoImponible = 0
        mRG_Coeficiente = 0
        mRG_PagosAnterioresMes = 0
        mRG_PagosTotalMes = 0
        mRG_PagosRetAnteriores = 0
        mRG_TxtFormula = 0
        mIB_Coef = 0
        mIB_Base = 0
        mIB_TxtFormula = 0
'            mCalculado = False
   
        mCodProveedor = CodProveedor
        mMontoNeto = montoNetoG
        mMontoNeto2 = montoNetoIB
        mFechaPeriodo = fechaperiodo
        
        comboSql cboTipoRetGan, "select descripcion, codigo from ProvTipoRetGan where activo = 1"
        If IsMissing(pTipoGan) Then
            cboTipoRetGan.ListIndex = BuscarEnCombo(cboTipoRetGan, s2n(obtenerDeSQL(" select TipoRetGan from prov where codigo = " & CodProveedor)))
        Else
            cboTipoRetGan.ListIndex = pTipoGan
        End If
        
        If IsMissing(pTipoIB) Then
            comboSql cboTipoIB, "select descripcion, codigo from ProvTipoRetIB where activo = 1"
            cboTipoIB.ListIndex = BuscarEnCombo(cboTipoIB, s2n(obtenerDeSQL("select TipoRetIIBB   from prov where codigo = " & CodProveedor)))
        
            'si no corresponde...
            If Not obtenerDeSQL("select reteneriibb from prov  where codigo = " & CodProveedor) Then
                If cboTipoIB.ListCount = 0 Then
                Else
                    cboTipoIB.ListIndex = 0
                End If
            End If
        Else
            cboTipoIB.ListIndex = pTipoIB
        End If
        
        mUltTipoGan = cboTipoRetGan.ListIndex
        mUltTipoIB = cboTipoIB.ListIndex
    
'        verRG_IB
        
        mRetGan = Round(CalculaRetGan(CodProveedor, montoNetoG, fechaperiodo), 2)
        txtRetGan = mRetGan
        mRetIB = Round(CalculaRetIb(CodProveedor, montoNetoIB), 2)
        txtRetIB = mRetIB
        verRG_IB
    End If

    Calcular = TotalRet()
    mCalculado = True
End Function
Public Property Get TotalRet() As Double
    TotalRet = mRetIB + mRetGan
End Property

Private Sub cboTipoIB_Validate(cancel As Boolean)
    'RaiseEvent cambio
    If mCalculado And mUltTipoIB <> cboTipoIB.ListIndex Then
        Calcular mCodProveedor, mMontoNeto2, mMontoNeto, mFechaPeriodo, , cboTipoIB.ListIndex
        If cboTipoIB.enabled Then RaiseEvent cambio(TotalRet)
    End If
End Sub
Private Sub cboTipoRetGan_Validate(cancel As Boolean)
    'RaiseEvent cambio
    If mCalculado And mUltTipoGan <> cboTipoRetGan.ListIndex Then
        Calcular mCodProveedor, mMontoNeto2, mMontoNeto, mFechaPeriodo, cboTipoRetGan.ListIndex
        If cboTipoRetGan.enabled Then RaiseEvent cambio(TotalRet)
    End If
End Sub

'*****************************************
Private Sub txtRetGan_GotFocus()
    txtRetGan.SelStart = 0
    txtRetGan.SelLength = Len(txtRetGan.Text)
End Sub
Private Sub txtRetGan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{tab}"
End Sub
Private Sub txtRetGan_LostFocus()
    mRetGan = s2n(txtRetGan)
    txtRetGan = mRetGan
End Sub
Private Sub txtRetIB_GotFocus()
    txtRetIB.SelStart = 0
    txtRetIB.SelLength = Len(txtRetIB.Text)
End Sub
Private Sub txtRetIB_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: SendKeys "{tab}"
End Sub
Private Sub txtRetIB_LostFocus()
    mRetIB = s2n(txtRetIB)
    txtRetIB = mRetIB
End Sub
'*********************************************************

Private Function CalculaRetGan(prov As Long, monto As Double, fechaperiodo As Date) As Double
    ' Busca pagos del mes, retenciones hechas y hace el calculo de proxima retencion
    
    Dim rangoFecha As String, sWhere As String
    Dim Anio, Mes, tempo
    Dim baseImp As Double, coef As Double, tipoGan As Long 'Tipo_Gan
    Dim SumPago As Double, SumRetGan3ro As Double
    
    ' tipoprov
'    tempo = obtenerDeSQL("select TipoRetGan, Baseimponible, coeficiente from prov inner join ProvTipoRetGan on prov.tipoRetGan = ProvTipoRetGan.codigo where prov.codigo = '" & prov & "' and prov.activo = 1 and ProvTipoRetGan.activo = 1 ")
    If tieneGAN Then
        tempo = obtenerDeSQL("select codigo, Baseimponible, coeficiente from ProvTipoRetGan_per where prov=" & prov & " and codigo = " & ComboCodigo(cboTipoRetGan))
    Else
        tempo = obtenerDeSQL("select codigo, Baseimponible, coeficiente from ProvTipoRetGan where codigo = " & ComboCodigo(cboTipoRetGan))
    End If
    If IsEmpty(tempo) Then
        che "No se encontraron datos de ProvTipoRetGan para " & prov
        CalculaRetGan = 0
        Exit Function
    Else
        tipoGan = tempo(0)
        baseImp = s2n(tempo(1))
        coef = s2n(tempo(2), 6)
    End If


    If prov <> 0 Then
    
        'where consulta
        Anio = Year(fechaperiodo)
        Mes = Month(fechaperiodo)
        rangoFecha = " fecha between '" & Format(DateSerial(Anio, Mes, 1), "yyyymmdd") & "' and '" & Format(DateSerial(Anio, Mes + 1, 0), "yyyymmdd") & "' "
        sWhere = " " & rangoFecha & " and activo = 1 and codpr = '" & prov & "' "
    
    
        'compras FAC contado
        tempo = obtenerDeSQL(" select sum(neto + exento) as SumTotal, sum(RetGanpago) as SumRetGan " & _
            " from compras where tipodoc = 'FAC' and contado = 1 and " & sWhere)
    
        SumPago = SumPago + s2n(tempo(0))
        SumRetGan3ro = SumRetGan3ro + s2n(tempo(1))
    
        'compras RAC
        tempo = obtenerDeSQL("select sum(neto + exento) as SumTotal, sum(RetGanpago) as SumRetGan " & _
            " from compras where tipodoc = 'RAC' and " & sWhere)
    
        SumPago = SumPago + s2n(tempo(0))
        SumRetGan3ro = SumRetGan3ro + s2n(tempo(1))
        
        'transcom RAC
        tempo = obtenerDeSQL("select sum(neto + exento) as SumTotal, sum(RetGanpago) as SumRetGan " & _
            " from transcom where tipodoc = 'RAC' and " & sWhere)
        SumPago = SumPago + s2n(tempo(0))
        SumRetGan3ro = SumRetGan3ro + s2n(tempo(1))
        
        'rec_com
        tempo = obtenerDeSQL("select sum(neto) as SumTotal, sum(RetGanpago) as SumRetGan " & _
            " from rec_comp where " & sWhere)
        SumPago = SumPago + s2n(tempo(0))
        SumRetGan3ro = SumRetGan3ro + s2n(tempo(1))
    End If

    'If tipoGan = TIPOGAN_H Then
    '    CalculaRetGan = CalculaGanH(monto, baseImp, coef, SumPago, SumRetGan3ro) ' fechaperiodo, baseimp, coef)??
    'Else ' tipoGan_LE, tipoGan_LS
    '    CalculaRetGan = calculaGanL(monto, baseImp, coef, SumPago, SumRetGan3ro)
    'End If
    CalculaRetGan = calculoGanGen(monto, baseImp, coef, SumPago, SumRetGan3ro, tipoGan, prov)
    'set propiedad
    mRG_MinimoNoImponible = baseImp
    mRG_Coeficiente = coef
    mRG_PagosAnterioresMes = SumPago
    mRG_PagosTotalMes = SumPago + monto
    mRG_PagosRetAnteriores = SumRetGan3ro
    
End Function

Private Function calculoGanGen(monto, baseImp, coef, SumPago, SumRetGan3ro, Tipo, provB) As Double
Dim totalSujeto  As Double, baseAImponer As Double, mSujeto As Double
Dim montoSobreEscala As Double, RetPorcentaje As Double, RetFijo As Double
Dim t As Variant, Iva, valorIva As Double
Dim tmpRet As Double
Iva = obtenerDeSQL("select tipoiva from prov where codigo=" & provB)
valorIva = obtenerDeSQL("select porcentaje from porcentajesiva where iva=" & Iva)
MinPagoGan = nSinNull(obtenerDeSQL("select minpagogan from bs where id=1"))
    Select Case Tipo
        Case 1:
            If valorIva > 0 Then
                totalSujeto = SumPago + (monto / (1 + valorIva))
                'totalSujeto = (SumPago + monto) / (1 + valorIva)
            Else
                totalSujeto = (SumPago + monto)
            End If
            mSujeto = totalSujeto - baseImp
            'totalSujeto = SumPago + monto - baseImp
            
            
'            If tieneGAN Then
'                t = obtenerDeSQL("select max(porcentaje), max(sobreExcedente) " & _
'                        " from ProvTipoRetGanRangosPorc_per " & _
'                        " where prov= " & provB & " and hasta < " & x2s(totalSujeto))
'                RetPorcentaje = s2n(t(0))
'                montoSobreEscala = s2n(t(1))
'                t = obtenerDeSQL(" select max(fijo) from ProvTipoRetGanRangoFijo_per " & _
'                        " where prov= " & provB & " and desde  < " & x2s(totalSujeto))
'            Else
                t = obtenerDeSQL("select max(porcentaje), max(sobreExcedente) " & _
                        " from ProvTipoRetGanRangosPorc " & _
                        " where hasta < " & x2s(totalSujeto))
                RetPorcentaje = s2n(t(0))
                montoSobreEscala = s2n(t(1))
                t = obtenerDeSQL(" select max(fijo) from ProvTipoRetGanRangoFijo " & _
                        " where desde  < " & x2s(totalSujeto))
'            End If
            
            RetFijo = s2n(t)
            
            baseAImponer = mSujeto - montoSobreEscala
            
            tmpRet = ((baseAImponer * (RetPorcentaje / 100)) + RetFijo)
            
            calculoGanGen = tmpRet - SumRetGan3ro
            
            If calculoGanGen < MinPagoGan Then calculoGanGen = 0
            
            mRG_TxtFormula = " (" & totalSujeto & " - " & montoSobreEscala & ") * " & RetPorcentaje & "% + " & RetFijo
            lblRG = "TS " & totalSujeto & ", SumPago " & SumPago & ", RetCalc " & tmpRet
        Case 2:
            If valorIva > 0 Then
                totalSujeto = SumPago + (monto / (1 + valorIva))
            Else
                totalSujeto = (SumPago + monto)
            End If
            
            mSujeto = totalSujeto - baseImp
            calculoGanGen = mSujeto * coef
            calculoGanGen = calculoGanGen - SumRetGan3ro
            If calculoGanGen < MinPagoGan Then calculoGanGen = 0
            
            lblRG = "sumPago = " & SumPago & " RetCalc = " & tmpRet
            mRG_TxtFormula = " " & s2n((totalSujeto - baseImp)) & " * " & s2n(coef * 100) & "% "
        Case 3:
            If valorIva > 0.21 Then valorIva = valorIva - 0.21
            If valorIva > 0 Then
                'totalSujeto = (SumPago + monto) / (1 + valorIva)
                totalSujeto = SumPago + (monto / (1 + valorIva))
            Else
                totalSujeto = (SumPago + monto)
            End If
            
            mSujeto = totalSujeto - baseImp
            'totalSujeto = SumPago + monto - baseImp
            calculoGanGen = mSujeto * coef
            calculoGanGen = calculoGanGen - SumRetGan3ro
            If calculoGanGen < MinPagoGan Then calculoGanGen = 0
            lblRG = "sumPago = " & SumPago & " RetCalc = " & tmpRet
            mRG_TxtFormula = " " & s2n((totalSujeto - baseImp)) & " * " & s2n(coef * 100) & "% "
    End Select
End Function
    
Private Function calculaGanL(monto, baseImp, coef, SumPago, SumRetGan3ro) As Double 'fuera de servicio
    Dim tmpRet As Double
    
    If SumPago + monto > baseImp Then
        tmpRet = (SumPago + monto - baseImp)
        tmpRet = tmpRet * coef
        'tmpRet = (SumPago + monto - baseimp) * coef
        calculaGanL = s2n(tmpRet - SumRetGan3ro)
        If calculaGanL < MINIMOPAGO Then calculaGanL = 0
    End If
    lblRG = "sumPago = " & SumPago & " RetCalc = " & tmpRet
    
    mRG_TxtFormula = " " & s2n((SumPago + monto - baseImp)) & " * " & s2n(coef * 100) & "% "
End Function

Private Function CalculaGanH(monto, baseImp, coef, SumPago, SumRetGan3ro) As Double 'fuera de servicio
    Dim totalSujeto  As Double, baseAImponer As Double
    Dim montoSobreEscala As Double, RetPorcentaje As Double, RetFijo As Double
    Dim t As Variant
    Dim tmpRet As Double
    

    totalSujeto = SumPago + monto - baseImp
    
    t = obtenerDeSQL("select max(porcentaje), max(sobreExcedente) " & _
            " from ProvTipoRetGanRangosPorc " & _
            " where hasta < '" & ssNum(totalSujeto) & "' ")
    RetPorcentaje = s2n(t(0))
    montoSobreEscala = s2n(t(1))
    t = obtenerDeSQL(" select max(fijo) from ProvTipoRetGanRangoFijo " & _
            " where desde  < '" & ssNum(totalSujeto) & "' ")
    RetFijo = s2n(t)
    
    baseAImponer = totalSujeto - montoSobreEscala
    
    tmpRet = ((baseAImponer * (RetPorcentaje / 100)) + RetFijo)
    
    CalculaGanH = tmpRet - SumRetGan3ro
    
    If CalculaGanH < MINIMOPAGO Then CalculaGanH = 0
    
    mRG_TxtFormula = " (" & totalSujeto & " - " & montoSobreEscala & ") * " & RetPorcentaje & "% + " & RetFijo
    lblRG = "TS " & totalSujeto & ", SumPago " & SumPago & ", RetCalc " & tmpRet
End Function

Private Function CalculaRetIb(CodProveedor As Long, monto As Double) As Double
    Dim tempo
    Dim baseImp As Double, coef As Double
    CodProv = s2n(CodProveedor)
    'agregado para ret iibb personal
    If tieneIIBB Then
        tempo = obtenerDeSQL("select  Baseimponible, coeficiente from ProvTipoRetIB_Per where codigo = " & ComboCodigo(cboTipoIB) & " and codprov =" & s2n(CodProveedor))
    Else
        tempo = obtenerDeSQL("select  Baseimponible, coeficiente from ProvTipoRetIB where codigo = " & ComboCodigo(cboTipoIB))
    End If
    
    If IsEmpty(tempo) Then
    Else
        baseImp = s2n(tempo(0))
        coef = s2n(tempo(1), 6)
    End If
    
    If monto > baseImp Then
        CalculaRetIb = monto * (coef / 100)
    End If
    
    mIB_Coef = coef
    mIB_Base = baseImp
    mIB_TxtFormula = s2n(monto) & " *  %" & (coef)
End Function
'*************************************
Private Sub verRG_IB()
    Dim t
    If tieneGAN Then
        t = obtenerDeSQL("select Baseimponible, coeficiente from ProvTipoRetGan_per where codigo = '" & ComboCodigo(cboTipoRetGan) & "' and prov=" & CodProv)
    Else
        t = obtenerDeSQL("select Baseimponible, coeficiente from ProvTipoRetGan where codigo = '" & ComboCodigo(cboTipoRetGan) & "' ")
    End If
    If ComboCodigo(cboTipoRetGan) <> 1 Then
        lblRG = ""
        If Not IsEmpty(t) Then
            If s2n(t(0)) > 0 Then lblRG = "Base = " & t(0)
            If s2n(t(1)) > 0 Then lblRG = lblRG & " Coef = " & t(1)
        End If
    End If
    If tieneIIBB Then
        t = obtenerDeSQL("select coeficiente from ProvTipoRetIB_per  where codigo = '" & ComboCodigo(cboTipoIB) & "' and codprov=" & CodProv)
    Else
        t = obtenerDeSQL("select coeficiente from ProvTipoRetIB  where codigo = '" & ComboCodigo(cboTipoIB) & "' ")
    End If
    lblIB = ""
    If Not IsEmpty(t) Then
        If s2n(t) > 0 Then lblIB = "Coef " & s2n(t)
    End If
End Sub


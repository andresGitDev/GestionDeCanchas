Attribute VB_Name = "ModuloTablaTemp"
Option Explicit

Public Const tt_OrdenPagoTemp = " ([TIPODOC] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [NRODOC] [numeric](18, 0) NULL , [FECHA] [datetime] NULL , [SALDO] [float] NULL , [PAGADO] [float] NULL , [ULTIMOSALDO] [float] NULL) "
Public Const tt_ChequeOPtmp = "( [nroint] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [banco] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,[cheque] [varchar] (150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [importe] [float] NULL ,  [fecha] [datetime] NULL, [propio] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL )"
Public Const tt_CajasTemp = "([saldoanterior] [float] NULL, [fecha] [datetime] NULL ,   [caja] [smallint] NULL , [movimiento] [numeric](18, 0) NULL , [ingegr] [varchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [importe] [float] NULL , [desctipo] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [interno] [numeric](18, 0) NULL , [concepto] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [total] [float] NULL)"

'Public Const tt_iva_compras_temp = "([fecha] [datetime] NULL , [razonsocial] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [nrocuit] [varchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [tipoynro] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,  [neto] [float] NULL , [ivaresp] [float] NULL , [rg3337] [float] NULL , [iva27] [float] NULL , [iva10] [float] NULL , [imptotal] [float] NULL , [retenciva] [float] NULL , [impint] [float] NULL , [retencgan] [float] NULL , [rg3431] [float] NULL , [exento] [float] NULL , [IB_CAPITAL] [float] NULL , [IB_PROVINCIA] [float] NULL, [letra] [varchar] (2) NULL)"
Public Const tt_FacturaCompra_temp = "([descripcion] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [total] [float] NULL , [iva21] [float] NULL , [iva10] [float] NULL , [iva27] [float] NULL , [rg3337] [float] NULL , [rg3431] [float] NULL , [neto] [float] NULL , [perIb] [float] NULL , [exento] [float] NULL , [RetGanPago] [float] NULL , [impint] [float] NULL , [retiva] [float] NULL)"
Public Const tt_SaldoCliTemp = "([codigo] [numeric](18, 0) NULL , [descripcion] [varchar] (100) COLLATE Modern_Spanish_CI_AS NULL , [saldo] [float] NULL)"
Public Const tt_SaldoProvTemp = "([CODPR] [numeric](18, 0) NULL ,[RSOCIAL] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [FECHA] [datetime] NULL , [TIPODOC] [varchar] (3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [NRODOC] [numeric](18, 0) NULL , [VENCIM] [datetime] NULL , [DEBE] [float] NULL , [HABER] [float] NULL , [SALDO] [float] NULL ,[INTERNO] [numeric](18, 0) NULL)"
Public Const tt_TipoComprasTemp = "([descripcion] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [neto] [float] NULL ,[iva] [float] NULL)"
Public Const tt_SeriesEnStockTemp = "(codigo numeric (18, 0), Producto char (35), Descripcion char (500), Serie varchar (500), TipoDoc char (3) NULL, codComprobante numeric (18,0))"
Public Const tt_IIBB = "([PROVINCIA] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [NETO] [float] (8), [IVA] [float] (8))"
Public Const tt_LibroDiario = "([columna1] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [columna2] [varchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [columna3] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [columna4] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL , [columna5] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL)"


' Mandar sql para crear tabla, sin el "CREATE TABLE XXXXX ", solo la definicion de campos, y si sabes como (yo no) de indices
Public Function TablaTempCrear(SqlCreateTableTMP_ As String) As String
'       "(Nombre TEXT (25) , Apellidos TEXT (50))"
'       "(Nombre TEXT (10), Apellidos TEXT, Fecha_Nacimiento DATETIME) CONSTRAINT IndiceGeneral UNIQUE ([Nombre], [Apellidos])"
'       "(ID INTEGER CONSTRAINT IndicePrimario PRIMARY, Nombre TEXT, Apellidos TEXT, Fecha_Nacimiento DATETIME)"

'    On Error GoTo UfaTTCrear ' q lo maneje el q manda
    Dim s As String, t As String
    
    If SqlCreateTableTMP_ = "" Then
        ufa "prg: Falta especificar sql", "TablaTempCrear"
        Exit Function
    End If
    If TABLA_TEMP_FIJA Then
        t = "__" & NuevoNombre() ' **************************
    Else
        t = "#" & NuevoNombre()
    End If
    s = "create table " & t & " " & SqlCreateTableTMP_
    
    DataEnvironment1.Sistema.Execute s
    TablaTempCrear = t

fin:
    Exit Function
UfaTTCrear:
    TablaTempCrear = ""
    Resume fin:
End Function

Public Function TablaTempCopiar(NombreTablaBase As String, Optional sWhere As String) As String '
'    On Error GoTo UfaTTCopiar
    Dim s As String, t As String
    
    If NombreTablaBase = "" Then
        ufa "prg: Falta especificar TablaBase", "TablaTempCopiar"
        Exit Function
    End If
    
    t = "#" & NuevoNombre()
   'debug
   't = NuevoNombre()
    s = "select * into " & t & " from " & NombreTablaBase
    If sWhere > "" Then s = s & " where " & sWhere
    
    DataEnvironment1.Sistema.Execute s
    TablaTempCopiar = t
   
fin:
    Exit Function
UfaTTCopiar:
    TablaTempCopiar = ""
    Resume fin:
End Function

''               con el CREATE # se volvio obsoleta
'Public Function TablaTempBorrar(ByRef TablaTemp As String) As Boolean  ' deprecated
''    ' devuelve true si borro, y pone nombre tabla = ""
''    ' NO  da err ni msg si es vacio, pero tampoco devuelve true porque no borro nada
''    On Error GoTo UfaTmpBorrar
''
''    If TablaTemp = "" Then Exit Function
''
''    If Left(TablaTemp, 6) <> "TmpTbl" Then
''        ufa "Prg: No se pudo borrar tabla temporal", "NO ES TmpTblXXXXXX, Parametro: " & TablaTemp
''        Exit Function
''    End If
''
''    daTaenvironment1.Sistema.Execute "drop table " & TablaTemp
'    TablaTemp = ""
'    TablaTempBorrar = True
''FIN:
''    Exit Function
''UfaTmpBorrar:
''    ufa "Err Al borrar tabla temp", "Parametro: " & TablaTemp
''    Resume FIN
'End Function


Private Function NuevoNombre()
    Randomize
    NuevoNombre = "TmpTbl" & Format(Now, "MMDDhhmmss") & Format(Int(Rnd() * 1000), "000")
End Function

'La cláusula CONSTRAINT
'Se utiliza la cláusula CONSTRAINT en las instrucciones ALTER TABLE y CREATE TABLE para crear o eliminar índices. Existen dos sintaxis para esta cláusula dependiendo si desea Crear ó Eliminar un índice de un único campo o si se trata de un campo multiíndice. Si se utiliza el motor de datos de Microsoft, sólo podrá utilizar esta cláusula con las bases de datos propias de dicho motor.
'Para los índices de campos únicos:
'CONSTRAINT nombre {PRIMARY KEY | UNIQUE | REFERENCES tabla externa
'[(campo externo1, campo externo2)]}
'Para los índices de campos múltiples:
'CONSTRAINT nombre {PRIMARY KEY (primario1[, primario2 [, ...]]) |
'UNIQUE (único1[, único2 [, ...]]) |
'FOREIGN KEY (ref1[, ref2 [, ...]]) REFERENCES tabla externa [(campo externo1
'[,campo externo2 [, ...]])]}
'
'Parte Descripción
'nombre  Es el nombre del índice que se va a crear.
'primarioN   Es el nombre del campo o de los campos que forman el índice primario.
'únicoN  Es el nombre del campo o de los campos que forman el índice de clave única.
'refN    Es el nombre del campo o de los campos que forman el índice externo (hacen referencia a campos de otra tabla).
'tabla externa   Es el nombre de la tabla que contiene el campo o los campos referenciados en refN
'campos externos Es el nombre del campo o de los campos de la tabla externa especificados por ref1, ref2, ..., refN
'Si se desea crear un índice para un campo cuando se esta utilizando las instrucciones ALTER TABLE o CREATE TABLE la cláusula CONTRAINT debe aparecer inmediatamente después de la especificación del campo indexeado.
'Si se desea crear un índice con múltiples campos cuando se está utilizando las instrucciones ALTER TABLE o CREATE TABLE la cláusula CONSTRAINT debe aparecer fuera de la cláusula de creación de tabla.
'
'Tipo de Indice  Descripción
'UNIQUE  Genera un índece de clave única. Lo que implica que los registros de la tabla no pueden contener el mismo valor en los campos indexados.
'PRIMARY KEY Genera un índice primario el campo o los campos especificados. Todos los campos de la clave principal deben ser únicos y no nulos, cada tabla sólo puede contener una única clave principal.
'FOREIGN KEY Genera un índice externo (toma como valor del índice campos contenidos en otras tablas). Si la clave principal de la tabla externa consta de más de un campo, se debe utilizar una definición de índice de múltiples campos, listando todos los campos de referencia, el nombre de la tabla externa, y los nombres de los campos referenciados en la tabla externa en el mismo orden que los campos de referencia listados. Si los campos referenciados son la clave principal de la tabla externa, no tiene que especificar los campos referenciados, predeterminado por valor, el motor Jet se comporta como si la clave principal de la tabla externa fueran los campos referenciados .
'

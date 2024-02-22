Attribute VB_Name = "ModActCliProv"
'****************************************************************
'Microsoft SQL Server 2000
'Visual Basic file generated for DTS Package
'File Name: C:\Documents and Settings\RaulBeta\Mis documentos\Coeficientes.bas
'Package Name: Coeficientes
'Package Description: DTS Clientes Proveedor
'Generated Date: 10/09/2008
'Generated Time: 10:15:25 a.m.
'****************************************************************

Option Explicit
Public goPackageOld As New DTS.Package
Public goPackage As DTS.Package2
Public MAQ As String
Public BAS As String
Public TABL As String
Public ARCH As String
Public NOMB As String



'Private Sub Main()
'OJO LO UNICO FIJO ES QUE SON 12 COLUMNAS
Public Function ACTCLIPROV(MAQ_O As String, BAS_O As String, TABL_O As String, ARCH_O As String, NOMB_O As String) As Boolean
MAQ = "german"
MAQ = MAQ_O
BAS = BAS_O
TABL = TABL_O
ARCH = ARCH_O
NOMB = NOMB_O

On Error GoTo ErrImport
ACTCLIPROV = True
        Set goPackage = goPackageOld

        goPackage.Name = "Coeficientes"
        goPackage.Description = "DTS Clientes Proveedor"
        goPackage.WriteCompletionStatusToNTEventLog = False
        goPackage.FailOnError = False
        goPackage.PackagePriorityClass = 2
        goPackage.MaxConcurrentSteps = 4
        goPackage.LineageOptions = 0
        goPackage.UseTransaction = True
        goPackage.TransactionIsolationLevel = 4096
        goPackage.AutoCommitTransaction = True
        goPackage.RepositoryMetadataOptions = 0
        goPackage.UseOLEDBServiceComponents = True
        goPackage.LogToSQLServer = False
        goPackage.LogServerFlags = 0
        goPackage.FailPackageOnLogFailure = False
        goPackage.ExplicitGlobalVariables = False
        goPackage.PackageType = 0
        

Dim oConnProperty As DTS.OleDBProperty

'---------------------------------------------------------------------------
' create package connection information
'---------------------------------------------------------------------------

Dim oConnection As DTS.Connection2

'------------- a new connection defined below.
'For security purposes, the password is never scripted

Set oConnection = goPackage.Connections.New("DTSFlatFile")

        oConnection.ConnectionProperties("Data Source") = ARCH '"D:\TONKA\TMPTXTDAT.txt"
        oConnection.ConnectionProperties("Mode") = 1
        oConnection.ConnectionProperties("Row Delimiter") = vbCrLf
        oConnection.ConnectionProperties("File Format") = 1
        oConnection.ConnectionProperties("Column Delimiter") = ";"
        oConnection.ConnectionProperties("File Type") = 1
        oConnection.ConnectionProperties("Skip Rows") = 0
        oConnection.ConnectionProperties("Text Qualifier") = """"
        oConnection.ConnectionProperties("First Row Column Name") = False
        oConnection.ConnectionProperties("Max characters per delimited column") = 8000
        
        oConnection.Name = "Connection 1"
        oConnection.ID = 1
        oConnection.Reusable = True
        oConnection.ConnectImmediate = False
        oConnection.DataSource = ARCH '"D:\TONKA\TMPTXTDAT.txt"
        oConnection.ConnectionTimeout = 60
        oConnection.UseTrustedConnection = False
        oConnection.UseDSL = False
        
        'If you have a password for this connection, please uncomment and add your password below.
        'oConnection.Password = "<put the password here>"

goPackage.Connections.Add oConnection
Set oConnection = Nothing

'------------- a new connection defined below.
'For security purposes, the password is never scripted

Set oConnection = goPackage.Connections.New("SQLOLEDB")

        oConnection.ConnectionProperties("Integrated Security") = "SSPI"
        oConnection.ConnectionProperties("Persist Security Info") = True
        oConnection.ConnectionProperties("Initial Catalog") = BAS '"Tonka"
        oConnection.ConnectionProperties("Data Source") = MAQ '"RAUL"
        oConnection.ConnectionProperties("Application Name") = "DTS  Import/Export Wizard"
        
        oConnection.Name = "Connection 2"
        oConnection.ID = 2
        oConnection.Reusable = True
        oConnection.ConnectImmediate = False
        oConnection.DataSource = MAQ '"RAUL"
        oConnection.ConnectionTimeout = 60
        oConnection.Catalog = BAS '"Tonka"
        oConnection.UseTrustedConnection = True
        oConnection.UseDSL = False
        
        'If you have a password for this connection, please uncomment and add your password below.
        'oConnection.Password = "<put the password here>"

goPackage.Connections.Add oConnection
Set oConnection = Nothing

'---------------------------------------------------------------------------
' create package steps information
'---------------------------------------------------------------------------

Dim oStep As DTS.Step2
Dim oPrecConstraint As DTS.PrecedenceConstraint

'------------- a new step defined below

Set oStep = goPackage.Steps.New

        oStep.Name = "Create Table [" & BAS & "].[dbo].[" & TABL & "] Step"
        oStep.Description = "Create Table [" & BAS & "].[dbo].[" & TABL & "] Step"
        oStep.ExecutionStatus = 1
        oStep.TaskName = "Create Table [" & BAS & "].[dbo].[" & TABL & "] Task"
        oStep.CommitSuccess = False
        oStep.RollbackFailure = False
        oStep.ScriptLanguage = "VBScript"
        oStep.AddGlobalVariables = True
        oStep.RelativePriority = 3
        oStep.CloseConnection = False
        oStep.ExecuteInMainThread = False
        oStep.IsPackageDSORowset = False
        oStep.JoinTransactionIfPresent = False
        oStep.DisableStep = False
        oStep.FailPackageOnError = False
        
goPackage.Steps.Add oStep
Set oStep = Nothing

'------------- a new step defined below

Set oStep = goPackage.Steps.New

        oStep.Name = "Copy Data from " & NOMB & " to [" & BAS & "].[dbo].[" & TABL & "] Step"
        oStep.Description = "Copy Data from " & NOMB & " to [" & BAS & "].[dbo].[" & TABL & "] Step"
        oStep.ExecutionStatus = 1
        oStep.TaskName = "Copy Data from " & NOMB & " to [" & BAS & "].[dbo].[" & TABL & "] Task"
        oStep.CommitSuccess = False
        oStep.RollbackFailure = False
        oStep.ScriptLanguage = "VBScript"
        oStep.AddGlobalVariables = True
        oStep.RelativePriority = 3
        oStep.CloseConnection = False
        oStep.ExecuteInMainThread = False
        oStep.IsPackageDSORowset = False
        oStep.JoinTransactionIfPresent = False
        oStep.DisableStep = False
        oStep.FailPackageOnError = False
        
goPackage.Steps.Add oStep
Set oStep = Nothing

'------------- a precedence constraint for steps defined below

Set oStep = goPackage.Steps("Copy Data from " & NOMB & " to [" & BAS & "].[dbo].[" & TABL & "] Step")
Set oPrecConstraint = oStep.PrecedenceConstraints.New("Create Table [" & BAS & "].[dbo].[" & TABL & "] Step")
        oPrecConstraint.StepName = "Create Table [" & BAS & "].[dbo].[" & TABL & "] Step"
        oPrecConstraint.PrecedenceBasis = 0
        oPrecConstraint.Value = 4
        
oStep.PrecedenceConstraints.Add oPrecConstraint
Set oPrecConstraint = Nothing

'---------------------------------------------------------------------------
' create package tasks information
'---------------------------------------------------------------------------

'------------- call Task_Sub1 for task Create Table [Tonka].[dbo].[ActCliProv] Task (Create Table [Tonka].[dbo].[ActCliProv] Task)
Call Task_Sub1(goPackage)

'------------- call Task_Sub2 for task Copy Data from TMPTXTDAT to [Tonka].[dbo].[ActCliProv] Task (Copy Data from TMPTXTDAT to [Tonka].[dbo].[ActCliProv] Task)
Call Task_Sub2(goPackage)

'---------------------------------------------------------------------------
' Save or execute package
'---------------------------------------------------------------------------

'goPackage.SaveToSQLServer "(local)", "sa", ""
goPackage.Execute
tracePackageError goPackage
goPackage.UnInitialize
'to save a package instead of executing it, comment out the executing package line above and uncomment the saving package line
Set goPackage = Nothing

Set goPackageOld = Nothing


Exit Function
ErrImport:
ACTCLIPROV = False
End Function


'-----------------------------------------------------------------------------
' error reporting using step.GetExecutionErrorInfo after execution
'-----------------------------------------------------------------------------
Public Sub tracePackageError(oPackage As DTS.Package)
Dim ErrorCode As Long
Dim ErrorSource As String
Dim ErrorDescription As String
Dim ErrorHelpFile As String
Dim ErrorHelpContext As Long
Dim ErrorIDofInterfaceWithError As String
Dim i As Integer

        For i = 1 To oPackage.Steps.Count
                If oPackage.Steps(i).ExecutionResult = DTSStepExecResult_Failure Then
                        oPackage.Steps(i).GetExecutionErrorInfo ErrorCode, ErrorSource, ErrorDescription, _
                                        ErrorHelpFile, ErrorHelpContext, ErrorIDofInterfaceWithError
                        'MsgBox oPackage.Steps(i).Name & " failed" & vbCrLf & ErrorSource & vbCrLf & ErrorDescription
                End If
        Next i

End Sub

'------------- define Task_Sub1 for task Create Table [Tonka].[dbo].[ActCliProv] Task (Create Table [Tonka].[dbo].[ActCliProv] Task)
Public Sub Task_Sub1(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask1 As DTS.ExecuteSQLTask2
Set oTask = goPackage.Tasks.New("DTSExecuteSQLTask")
oTask.Name = "Create Table [" & BAS & "].[dbo].[" & TABL & "] Task"
Set oCustomTask1 = oTask.CustomTask

        oCustomTask1.Name = "Create Table [" & BAS & "].[dbo].[" & TABL & "] Task"
        oCustomTask1.Description = "Create Table [" & BAS & "].[dbo].[" & TABL & "] Task"
        oCustomTask1.SQLStatement = "CREATE TABLE [" & BAS & "].[dbo].[" & TABL & "] (" & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[F1] varchar (8000) NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[F2] varchar (8000) NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[F3] varchar (8000) NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[F4] varchar (8000) NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[F5] varchar (8000) NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[F6] varchar (8000) NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[F7] varchar (8000) NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[F8] varchar (8000) NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[F9] varchar (8000) NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[F10] varchar (8000) NULL, " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "[F11] varchar (8000) NULL" & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & ")"
        oCustomTask1.ConnectionID = 2
        oCustomTask1.CommandTimeout = 0
        oCustomTask1.OutputAsRecordset = False
        
goPackage.Tasks.Add oTask
Set oCustomTask1 = Nothing
Set oTask = Nothing

End Sub

'------------- define Task_Sub2 for task Copy Data from TMPTXTDAT to [Tonka].[dbo].[ActCliProv] Task (Copy Data from TMPTXTDAT to [Tonka].[dbo].[ActCliProv] Task)
Public Sub Task_Sub2(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask2 As DTS.DataPumpTask2
Set oTask = goPackage.Tasks.New("DTSDataPumpTask")
oTask.Name = "Copy Data from " & NOMB & " to [" & BAS & "].[dbo].[" & TABL & "] Task"
Set oCustomTask2 = oTask.CustomTask

        oCustomTask2.Name = "Copy Data from " & NOMB & " to [" & BAS & "].[dbo].[" & TABL & "] Task"
        oCustomTask2.Description = "Copy Data from " & NOMB & " to [" & BAS & "].[dbo].[" & TABL & "] Task"
        oCustomTask2.SourceConnectionID = 1
        oCustomTask2.SourceObjectName = ARCH '"D:\TONKA\TMPTXTDAT.txt"
        oCustomTask2.DestinationConnectionID = 2
        oCustomTask2.DestinationObjectName = "[" & BAS & "].[dbo].[" & TABL & "]"
        oCustomTask2.ProgressRowCount = 1000
        oCustomTask2.MaximumErrorCount = 0
        oCustomTask2.FetchBufferSize = 1
        oCustomTask2.UseFastLoad = True
        oCustomTask2.InsertCommitSize = 0
        oCustomTask2.ExceptionFileColumnDelimiter = "|"
        oCustomTask2.ExceptionFileRowDelimiter = vbCrLf
        oCustomTask2.AllowIdentityInserts = False
        oCustomTask2.FirstRow = 0
        oCustomTask2.LastRow = 0
        oCustomTask2.FastLoadOptions = 2
        oCustomTask2.ExceptionFileOptions = 1
        oCustomTask2.DataPumpOptions = 0
        
Call oCustomTask2_Trans_Sub1(oCustomTask2)
                
                
goPackage.Tasks.Add oTask
Set oCustomTask2 = Nothing
Set oTask = Nothing

End Sub

Public Sub oCustomTask2_Trans_Sub1(ByVal oCustomTask2 As Object)

        Dim oTransformation As DTS.Transformation2
        Dim oTransProps As DTS.Properties
        Dim oColumn As DTS.Column
        Set oTransformation = oCustomTask2.Transformations.New("DTS.DataPumpTransformCopy")
                oTransformation.Name = "DirectCopyXform"
                oTransformation.TransformFlags = 63
                oTransformation.ForceSourceBlobsBuffered = 0
                oTransformation.ForceBlobsInMemory = False
                oTransformation.InMemoryBlobSize = 1048576
                oTransformation.TransformPhases = 4
                
                Set oColumn = oTransformation.SourceColumns.New("Col001", 1)
                        oColumn.Name = "Col001"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 32
                        oColumn.Size = 8000
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Col002", 2)
                        oColumn.Name = "Col002"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 32
                        oColumn.Size = 8000
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Col003", 3)
                        oColumn.Name = "Col003"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 32
                        oColumn.Size = 8000
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Col004", 4)
                        oColumn.Name = "Col004"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 32
                        oColumn.Size = 8000
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Col005", 5)
                        oColumn.Name = "Col005"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 32
                        oColumn.Size = 8000
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Col006", 6)
                        oColumn.Name = "Col006"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 32
                        oColumn.Size = 8000
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Col007", 7)
                        oColumn.Name = "Col007"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 32
                        oColumn.Size = 8000
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Col008", 8)
                        oColumn.Name = "Col008"
                        oColumn.Ordinal = 8
                        oColumn.Flags = 32
                        oColumn.Size = 8000
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Col009", 9)
                        oColumn.Name = "Col009"
                        oColumn.Ordinal = 9
                        oColumn.Flags = 32
                        oColumn.Size = 8000
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Col010", 10)
                        oColumn.Name = "Col010"
                        oColumn.Ordinal = 10
                        oColumn.Flags = 32
                        oColumn.Size = 8000
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Col011", 11)
                        oColumn.Name = "Col011"
                        oColumn.Ordinal = 11
                        oColumn.Flags = 32
                        oColumn.Size = 8000
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("F1", 1)
                        oColumn.Name = "F1"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 32
                        oColumn.Size = 8000
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("F2", 2)
                        oColumn.Name = "F2"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 32
                        oColumn.Size = 8000
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("F3", 3)
                        oColumn.Name = "F3"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 32
                        oColumn.Size = 8000
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("F4", 4)
                        oColumn.Name = "F4"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 32
                        oColumn.Size = 8000
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("F5", 5)
                        oColumn.Name = "F5"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 32
                        oColumn.Size = 8000
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("F6", 6)
                        oColumn.Name = "F6"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 32
                        oColumn.Size = 8000
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("F7", 7)
                        oColumn.Name = "F7"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 32
                        oColumn.Size = 8000
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("F8", 8)
                        oColumn.Name = "F8"
                        oColumn.Ordinal = 8
                        oColumn.Flags = 32
                        oColumn.Size = 8000
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("F9", 9)
                        oColumn.Name = "F9"
                        oColumn.Ordinal = 9
                        oColumn.Flags = 32
                        oColumn.Size = 8000
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("F10", 10)
                        oColumn.Name = "F10"
                        oColumn.Ordinal = 10
                        oColumn.Flags = 32
                        oColumn.Size = 8000
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("F11", 11)
                        oColumn.Name = "F11"
                        oColumn.Ordinal = 11
                        oColumn.Flags = 32
                        oColumn.Size = 8000
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

        Set oTransProps = oTransformation.TransformServerProperties

                
        Set oTransProps = Nothing

        oCustomTask2.Transformations.Add oTransformation
        Set oTransformation = Nothing
End Sub


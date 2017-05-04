Imports System.Data.SqlClient
Imports System.IO

''' <summary>
''' This class provides the functionality required for storing program events to SQL database
''' </summary>
''' <remarks>
''' This class is for PERSONAL use
''' </remarks>
''' <author>Mattias Wallner</author>
''' <created>21/01/2014</created>
Public Class SQLLog

#Region "--== Structures ==--"

    ''' <summary>
    ''' This enumeration provides various log filter levels to use
    ''' </summary>
    ''' <remarks>
    ''' Level:
    ''' 1, All: This level will log Errors, Successes and Info. A lot of entries will be generated
    ''' 2, ErrorAndSuccess: This will log errors and success messages, but not information
    ''' 3, ErrorOnly: This is the default once deployed and will only log errors, not successes
    ''' 4, Debug: This level will log the same as 1, but also Debug messages
    ''' </remarks>
    ''' <author>Mattias Wallner</author>
    ''' <created>22/01/2014</created>
    ''' <history>
    ''' 17/06/2014 - MW - Added the Debug level
    ''' </history>
    Public Enum LogFilterEnum
        LogDebug = 1
        LogAll = 2
        LogErrorAndSuccess = 3
        LogErrorOnly = 4
    End Enum

    ''' <summary>
    ''' This enumeration provides verbose log levels to match the filter
    ''' </summary>
    ''' <remarks></remarks>
    ''' <author>Mattias Wallner</author>
    ''' <created>22/01/2014</created>
    ''' <history>
    ''' 17/06/2014 - MW - Added the Debug level
    ''' </history>
    Public Enum LogLevelEnum
        lvlDebug = 1
        lvlInfo = 2
        lvlSuccess = 3
        lvlError = 4
    End Enum

#End Region '--== Structures ==--

#Region "--== Variables and Properties ==--"

    Private _MstrAppName As String   'This is the application name to log

    'This section is provided by Karla and will store and open a connection to LincPort
    Private _LincportConnection As SqlConnection
    Private _MstrLincportConnection As String

    ''' <summary>
    ''' This property holds the current log filter level. It decided if a log entry is passed to the database or not
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    ''' <author>Mattias Wallner</author>
    ''' <created>21/01/2014</created>
    Public Property LogLevel As LogLevelEnum
#End Region '--== Variables and Properties ==--

#Region "--== Constructor(s) ==--"

    ''' <summary>
    ''' The constructor for the SQLLog class 
    ''' </summary>
    ''' <remarks></remarks>
    ''' <author>Mattias Wallner</author>
    ''' <created>21/01/2014</created>
    ''' <history></history>
    Public Sub New(ByVal ConnectionString As String, ByVal ServiceName As String)

        _MstrLincportConnection = ConnectionString
        _LincportConnection = New SqlConnection(_MstrLincportConnection)

        Try
            _LincportConnection.Open()
            _LincportConnection.Close()
        Catch ex As Exception
        End Try

        _MstrAppName = ServiceName
    End Sub

#End Region '--== Constructor(s) ==--

#Region "--== Class Events ==--"

    ''' <summary>
    ''' This event is fired if the database can't be written to and a log file is written to instead.
    ''' </summary>
    ''' <remarks></remarks>
    ''' <author>Mattias Wallner</author>
    ''' <created>21/01/2014</created>
    Public Event DatabaseWriteError()
#End Region '--== Class Events ==--

#Region "--== Methods ==--"

    ''' <summary>
    ''' This routine will return a string representing the passed log level
    ''' </summary>
    ''' <param name="logLevel">The log level enum to convert</param>
    ''' <returns>A string representing the Log level</returns>
    ''' <remarks>This is done as the Enum.ToString() can't be overridden and will only return the exact name of the value</remarks>
    ''' <author>Mattias Wallner</author>
    ''' <created>22/01/2014</created>
    ''' <history>
    ''' 23/01/2014 - MW - Changed to Public
    ''' 17/06/2014 - MW - Added Debug case
    ''' </history>
    Public Function LogLevelText(ByVal logLevel As LogLevelEnum) As String
        Select Case logLevel
            Case LogLevelEnum.lvlDebug
                Return "Debug"
            Case LogLevelEnum.lvlError
                Return "Error"
            Case LogLevelEnum.lvlSuccess
                Return "Success"
            Case LogLevelEnum.lvlInfo
                Return "Info"
            Case Else
                Return "?"  'This will only be encountered if the enum gets a new member and it is not included in the Case statement. This will shut up the warning for no return on all paths
        End Select
    End Function

    ''' <summary>
    ''' This routine will store the log information in the SQL database
    ''' </summary>
    ''' <param name="logLevel">Error/Success/Info (Enum)</param>
    ''' <param name="ErrorRoutine">The name of the Routine where the log entry originated</param>
    ''' <param name="ErrorDesc">The Error description</param>
    ''' <param name="ErrorNo">The Error number</param>
    ''' <remarks></remarks>
    ''' <author>Karla McPhearson</author>
    ''' <created>?</created>
    ''' <history>
    ''' 21/01/2014 - MW - Minor changes to variable names from Original to fit with the class and what is used in the BTMService project
    ''' 22/01/2014 - MW - Changed to take the LogLevel input and to only enter the log levels that are accepted by the log filter setting
    ''' </history>
    Public Sub WriteSQLLog(ByVal logLevel As LogLevelEnum, ByVal ErrorRoutine As String, _
                           ByVal ErrorDesc As String, ByVal ErrorNo As Integer)

        Dim cmdSQL As New SqlClient.SqlCommand
        Dim intProcessStatus As Integer

        'Filter the log messages
        If logLevel >= LogLevel Then

            Try
                If OpenDatabase() = 0 Then  'The database was opened successfully

                    ErrorDesc = Strings.Replace(ErrorDesc, "'", "")

                    With cmdSQL
                        .Connection = _LincportConnection
                        .CommandText = "spPro_P_WriteErrorLog_I"
                        .CommandType = CommandType.StoredProcedure
                        .Parameters.Add(New SqlParameter("@prAppName", SqlDbType.VarChar, 50)).Value = _MstrAppName
                        .Parameters.Add(New SqlParameter("@prErrorType", SqlDbType.VarChar, 50)).Value = LogLevelText(logLevel)
                        .Parameters.Add(New SqlParameter("@prErrorName", SqlDbType.VarChar, 50)).Value = Strings.Left(ErrorRoutine, 50)
                        .Parameters.Add(New SqlParameter("@prErrorDescription", SqlDbType.VarChar, 200)).Value = Strings.Left(ErrorDesc, 200)
                        .Parameters.Add(New SqlParameter("@ProcessStatus", SqlDbType.Int)).Direction = ParameterDirection.Output
                        .ExecuteNonQuery()
                        intProcessStatus = CInt(.Parameters("@ProcessStatus").Value)
                    End With

                    If intProcessStatus <> 1 Then
                        WriteLog("WriteSQLLog", "No records written to database | intProcessStatus: " & intProcessStatus)
                    End If

                    CloseDatabase()

                End If
            Catch ex As Exception
                WriteLog("WriteSQLLog", ex.Message)

            Finally
                cmdSQL.Dispose()
            End Try
        End If
    End Sub
#End Region '--== Methods ==--

#Region "--== Private Event Handlers ==--"
#End Region '--== Private Event Handlers ==--

#Region "--== Private Routines ==--"

    ''' <summary>
    ''' Close the database connection if it is open
    ''' </summary>
    ''' <remarks></remarks>
    ''' <author>Karla McPhearson</author>
    ''' <created>?</created>
    ''' <history>
    ''' 21/01/2014 - MW - Changed to function to return 0 if successful and 1 otherwise
    ''' </history>
    Private Function CloseDatabase() As Integer
        Try
            If _LincportConnection.State = ConnectionState.Open Then
                _LincportConnection.Close()
            End If
            Return 0
        Catch ex As Exception
            WriteLog("CloseDatabase", ex.Message)
            Return 1
        End Try
    End Function

    ''' <summary>
    ''' This function will return the connection string to the LincPort database
    ''' </summary>
    ''' <param name="Server"></param>
    ''' <param name="Database"></param>
    ''' <param name="Encrypt"></param>
    ''' <param name="ServiceName"></param>
    ''' <param name="Security"></param>
    ''' <param name="TrustServerCert"></param>
    ''' <param name="User"></param>
    ''' <param name="Password"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <author>Karla McPherson</author>
    ''' <created>?</created>
    ''' <history></history>
    Public Function GetConnectionString(ByVal Server As String, ByVal Database As String, ByVal Encrypt As Boolean, ByVal ServiceName As String, Optional ByVal Security As Boolean = True, Optional ByVal TrustServerCert As Boolean = True, Optional ByVal User As String = "", Optional ByVal Password As String = "") As String
        GetConnectionString = ""
        Dim strConn As New SqlClient.SqlConnectionStringBuilder

        Try
            With strConn
                .Encrypt = Encrypt
                .ApplicationName = ServiceName
                .DataSource = Server
                .InitialCatalog = Database

                'If Security = False Then 'if security = True does not need a user/password, will use the user/password for the user that is running the service i.e. SQLService
                '    .UserID = User
                '    .Password = PasswordDecrypt(Password)
                'End If

                .IntegratedSecurity = Security
                .TrustServerCertificate = TrustServerCert
            End With

            GetConnectionString = strConn.ConnectionString

        Catch ex As Exception
            System.IO.File.WriteAllText("GetConnectionStringFail.txt", "GetConnectionStrring failed. The error message is: " & ex.Message & _
                                                    vbCrLf & "Time stamp: " & Now())
        End Try
        Return GetConnectionString
    End Function

    ''' <summary>
    ''' Open the database connection if it is closed
    ''' </summary>
    ''' <remarks></remarks>
    ''' <author>Karla McPhearson</author>
    ''' <created>?</created>
    ''' <history>
    ''' 21/01/2014 - MW - Changed to function to return 0 if successful and 1 otherwise
    ''' </history>
    Private Function OpenDatabase() As Integer
        Try
            If _LincportConnection.State = ConnectionState.Closed Then
                _LincportConnection.ConnectionString = _MstrLincportConnection
                _LincportConnection.Open()
            End If
            Return 0

        Catch ex As Exception
            WriteLog("OpenDatabase", ex.Message)
            Return 1
        End Try
    End Function

    ''' <summary>
    ''' decrypts password - same procedure as Nigel uses in the Lincport Service
    ''' </summary>
    ''' <param name="EncryptedPassword">The encrypted password to be decrypted</param>
    ''' <returns>The decrypted password</returns>
    ''' <remarks>
    ''' Password=sdcbp!ps@v1r12 converts to db!svr1@ppcs (12 is the number of chars of password)
    ''' Password=mseyts06 converts to system (06 is the number of chars of password)
    ''' </remarks>
    ''' <author>Karla McPhearson</author>
    ''' <created>?</created>
    ''' <history>
    ''' 20/08/2014 - MW - Set Option Strict, so needed to do some type casting to comply
    ''' </history>
    Private Function PasswordDecrypt(ByVal EncryptedPassword As String) As String

        Dim I As Integer
        Dim J As Integer
        Dim X As String
        Dim Z As String
        Dim Y As String

        Try
            If Len(EncryptedPassword) > 0 Then
                I = CInt(Strings.Right(EncryptedPassword, 2))
                J = 0
                X = Strings.Left(EncryptedPassword, I)
                Z = ""
                Y = ""
                For I = 1 To I
                    If J = 0 Then
                        Z = Strings.Left(X, 1) & Z
                        X = Strings.Right(X, Len(X) - 1)
                        J = 1
                    Else
                        Y = Y & Strings.Left(X, 1)
                        X = Strings.Right(X, Len(X) - 1)
                        J = 0
                    End If
                Next

                Return Y & Z
            Else
                Return "0"
            End If

        Catch ex As Exception
            WriteLog("PasswordDecrypt", "Exception: " & ex.Message)
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' This routine is for writing an error to a log file. To be used when the SQL log fails
    ''' </summary>
    ''' <remarks>
    ''' This is only used locally, and for errors only. One file per day
    ''' </remarks>
    ''' <author>Mattias Wallner</author>
    ''' <created>21/01/2014</created>
    ''' <history></history>
    Private Sub WriteLog(ByVal routine As String, ByVal errorMsg As String)

        Dim fileName As String = "ErrorLog_" & Format(Now(), "dd-MM-yyyy") & ".txt"
        Dim fileLines(0 To 0) As String

        Try
            fileLines(0) = Format(Now(), "HH:mm:ss") & " | Error |" & vbTab & routine & " - Message: " & errorMsg
            File.AppendAllLines(fileName, fileLines)

        Catch ex As Exception
            'The backup plan failed... Only do something in debug mode
#If DEBUG Then
            MsgBox("Error in WriteLog. Error message: " & ex.Message, vbOKOnly)
#End If
        End Try

        RaiseEvent DatabaseWriteError() 'Raise the event to tell the service that the SQL log failed
    End Sub
#End Region '--== Private Routines ==--
End Class

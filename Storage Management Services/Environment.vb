#Region "Imports"

Imports System.Text
Imports FileNet.Api.Core
Imports FileNet.Api.Authentication
Imports FileNet.Api.Admin
Imports Documents.Exceptions
Imports Documents.Utilities

#End Region

<DebuggerDisplay("{DebuggerIdentifier(),nq}")>
Public Class Environment

#Region "Class Variables"

  Private mstrEnvironmentName As String = String.Empty
  Private mstrCEServerName As String = String.Empty
  Private mintPort As Integer
  Private mstrServerName As String = Nothing
  Private mobjDomain As IDomain = Nothing
  Private mstrUserName As String = Nothing
  Private mstrPassword As String = Nothing
  Private mblnIsConnected As Boolean

#End Region

#Region "Public Properties"

  Public Property Name As String
    Get
      Return mstrEnvironmentName
    End Get
    Set(ByVal value As String)
      mstrEnvironmentName = value
    End Set
  End Property

  Public Property ServerName As String
    Get
      Return mstrServerName
    End Get
    Set(ByVal value As String)
      mstrServerName = value
    End Set
  End Property

  Public Property Port As Integer
    Get
      Return mintPort
    End Get
    Set(ByVal value As Integer)
      mintPort = value
    End Set
  End Property

  Public ReadOnly Property URL As String
    Get
      Try
        Return CreateURL()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

  <Xml.Serialization.XmlIgnore()>
  Public Property IsConnected As Boolean
    Get
      Return mblnIsConnected
    End Get
    Private Set(ByVal value As Boolean)
      mblnIsConnected = value
    End Set
  End Property

  Public ReadOnly Property Domain As IDomain
    Get
      Return mobjDomain
    End Get
  End Property

#End Region

#Region "Constructors"

  Public Sub New()

  End Sub

  Public Sub New(ByVal lpName As String, ByVal lpServerName As String, ByVal lpPort As Integer, ByVal lpUserName As String, ByVal lpPassword As String)
    Try
      Name = lpName
      ServerName = lpServerName
      Port = lpPort
      mstrUserName = lpUserName
      mstrPassword = lpPassword
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

#Region "Public Methods"

  Public Function CreateURL() As String
    Try
      Return String.Format("{0}://{1}:{2}/wsi/FNCEWS40MTOM/", "http", ServerName, Port)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function CreateEnvironmentRoot(ByVal lpSharedStorageRoot As String) As String
    Try
      Return String.Format("{0}\{1}", lpSharedStorageRoot, Me.Name)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Overrides Function ToString() As String
    Try
      Return DebuggerIdentifier()
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

#Region "Private Methods"

  Public Function InitializeConnection() As Boolean

    Console.WriteLine("Please wait, attempting to connect to {0} server {1}...", Name, ServerName)

    Try

      If IsConnected = True Then
        Return True
      End If

      Dim lobjCredentials As Credentials = Nothing

      lobjCredentials = New UsernameCredentials(mstrUserName, mstrPassword)
      FileNet.Api.Util.ClientContext.SetProcessCredentials(lobjCredentials)

      mobjDomain = GetDomain()

      Try
        mobjDomain.Refresh()

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        IsConnected = False
        '  Re-throw the exception
        Throw
      End Try

      IsConnected = True

      Return True

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      IsConnected = False
      Throw (New RepositoryNotAvailableException(Me.ServerName, String.Format("Unable to initialize connection to CE server {0}.", Me.ServerName), ex)) ' Exception("Unable to initialize connection", ex))
    End Try

  End Function

  Private Function GetDomain() As IDomain

    Try

      Dim lobjConnection As IConnection = Factory.Connection.GetConnection(URL)
      Dim lobjDomain As IDomain = Factory.Domain.GetInstance(lobjConnection, Nothing)

      Return lobjDomain

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  Friend Function DebuggerIdentifier() As String

    Dim lstrReturnBuilder As New StringBuilder

    Try

      lstrReturnBuilder.AppendFormat("{0}: {1}", Name, ServerName)
      If IsConnected Then
        lstrReturnBuilder.Append(" (Connected)")
      Else
        lstrReturnBuilder.Append(" (Not Connected)")
      End If

      Return lstrReturnBuilder.ToString

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

End Class

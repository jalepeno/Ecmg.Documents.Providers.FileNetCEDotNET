'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------


#Region "Imports"

Imports System
Imports System.IO
Imports FileNet.Api
Imports FileNet.Api.Admin
Imports FileNet.Api.Constants
Imports FileNet.Api.Core
Imports FileNet.Api.Property
Imports FileNet.Api.Collection
Imports FileNet.Api.Security
Imports FileNet.Api.Util
Imports FileNet.Api.Query
Imports System.Text
Imports System.Security.Principal
Imports FileNet.Api.Exception
Imports FileNet.Api.Authentication
Imports FileNet.Api.Action
Imports System.Collections.Specialized
Imports System.Reflection
Imports System.Net
Imports System.Net.Sockets
Imports System.Net.NetworkInformation
Imports System.Net.Security
Imports System.Security.Cryptography.X509Certificates
Imports Documents.Core
Imports Documents.Core.ChoiceLists
Imports Documents.Exceptions
Imports Documents.Arguments
Imports Documents.Providers
Imports Documents.Search
Imports Documents.Security
Imports Documents.Utilities
Imports DCore = Documents.Core
Imports DProviders = Documents.Providers
'Imports P8Security = FileNet.Api.Security
Imports DSecurity = Documents.Security
Imports Microsoft.Extensions.Configuration
Imports SysConfig = System.Configuration.ConfigurationManager
Imports Documents.Providers.FileNetCEDotNET.Exceptions

#End Region

#Const SupportAnnotations = True

Public Class CENetProvider
  Inherits DProviders.CProvider
  Implements IChoiceListImporter

#Region "Class Constants"

  Private Const PROVIDER_NAME As String = "P8 Content Engine .NET Provider"
  Private Const PROVIDER_SYSTEM_TYPE As String = "FileNet P8 4.x and above"
  Private Const PROVIDER_COMPANY_NAME As String = "IBM"
  Private Const PROVIDER_PRODUCT_NAME As String = "FileNet P8 Content Engine"
  Private Const PROVIDER_PRODUCT_VERSION As String = "5.5"

  Public Const DOMAIN_NAME As String = "DomainName"
  Public Const SERVER_NAME As String = "ServerName"
  Public Const OBJECT_STORE_NAME As String = "ObjectStore"
  Public Const PORT_NUMBER As String = "PortNumber"
  Public Const PROTOCOL As String = "Protocol"
  Public Const TRUSTED_CONNECTION As String = "TrustedConnection"
  Public Const PRESERVE_ID_ON_IMPORT As String = "PreserveIdOnImport"
  ' <Added by: Ernie at: 7/6/2021>
  Public Const PING_TO_VERIFY As String = "PingToVerify"
  ' </Added by: Ernie at: 7/6/2021>

  ' <Added by: Ernie at: 1/11/2013-3:43:14 PM on machine: ERNIE-THINK>
  Public Const EXPORT_SYSTEM_OBJECT_VALUED_PROPERTIES As String = "ExportSystemObjectValuedProperties"
  ' </Added by: Ernie at: 1/11/2013-3:43:14 PM on machine: ERNIE-THINK>
  Private Const FOLDER_DELIMITER As String = "/"
  Private Const PROP_CAN_DECLARE As String = "CanDeclare"
  Private Const PROP_CREATOR As String = "Creator"
  Private Const PROP_CREATE_DATE As String = "DateCreated"
  Private Const PROP_MODIFY_USER As String = "LastModifier"
  Private Const PROP_MODIFY_DATE As String = "DateLastModified"
  Private Const PROP_CHECK_IN_DATE As String = "DateCheckedIn"
  Private Const PROP_IS_CURRENT_VERSION As String = "IsCurrentVersion"
  Private Const PROP_MINOR_VERSION_NUMBER As String = "MinorVersionNumber"
  Private Const PROP_MAJOR_VERSION_NUMBER As String = "MajorVersionNumber"
  Private Const PROP_MIME_TYPE As String = "MimeType"
  Private Const PROP_OBJECT_NAME As String = "Name"
  Private Const PROP_OWNER As String = "Owner"
  Private Const PROP_RECORD_INFORMATION As String = "RecordInformation"
  Private Const PROP_VERSION_STATUS As String = "VersionStatus"
  Private Const PROP_VERSION_ID As String = "VersionId"
  Private Const PROP_VERSION_SERIES_ID As String = "VersionSeriesId"

  Public Const TREAT_OBJECT_IDS_AS_NUMBERS As String = "TreatObjectIdsAsNumbers"
  Public Const TREAT_OBJECT_IDS_AS_NUMBERS_DEFAULT_VALUE As String = "False"

  Private Const DEFAULT_EXPORT_PATH As String = "C:\Exports"
  Private Const DEFAULT_IMPORT_PATH As String = "C:\Imports"

  Private Const MAJOR_VERSION As String = "MajorVersion"

  Private Const ACCESS_LEVEL_MODIFY_PROPERTIES As Integer = 134599
  Private Const ACCESS_RIGHT_ALL_BUT_OWNER_CONTROL_DOCUMENT_ROLLUP As Integer = 400895
  Private Const ACCESS_RIGHT_ALL_BUT_OWNER_CONTROL_AND_PUBLISH_DOCUMENT_ROLLUP As Integer = 136703
  Private Const ACCESS_RIGHT_MODIFY_CONTENT_DOCUMENT_ROLLUP As Integer = 138747
  Private Const ACCESS_RIGHT_MODIFY_CONTENT_AND_PROPERTIES_ONLY_DOCUMENT_ROLLUP As Integer = 136699
  Private Const ACCESS_RIGHT_MODIFY_PROPERTIES_AND_VIEW_CONTENT_ONLY_DOCUMENT_ROLLUP As Integer = 136635
  Private Const ACCESS_RIGHT_VIEW_CONTENT_AND_PROPERTIES_ONLY_DOCUMENT_ROLLUP As Integer = 135337

#End Region

#Region "Class Variables"

  ' This is where you declare the system identifiers 
  ' These constants are duplicated in all classes implementing IProvider.
  ' The actual values are stored in the project's app.config file
  Private _
    mobjSystem As _
      New ProviderSystem(PROVIDER_NAME, PROVIDER_SYSTEM_TYPE, PROVIDER_COMPANY_NAME,
                         PROVIDER_PRODUCT_NAME, PROVIDER_PRODUCT_VERSION)

  Private mobjFolder As CFolder = New CENetFolder
  Private mobjSearch As CSearch = New CENetSearch

  'Private mstrDomainName As String = String.Empty
  Private Shared mstrServerName As String = String.Empty
  Private mstrObjectStoreName As String = String.Empty
  Private mstrPortNumber As String = String.Empty
  Private mstrUserName As String = String.Empty
  Private mstrPassword As String = String.Empty
  Private mstrProtocol As String = String.Empty
  Private mblnTrustedConnection As Boolean = False
  Private mblnPreserveIdOnImport As Boolean = True
  Private mblnTreatObjectIdsAsNumbers As Boolean = False
  Private mblnExportSystemObjectValuedProperties As Boolean = False
  Private mblnPingToVerify As Boolean = True

  ' Private mobjPropertyValueStrings As New Dictionary(Of String, String)

  Private mstrContentExportPropertyExclusions As String()

  Private mobjDomain As IDomain = Nothing
  Private mobjObjectStore As IObjectStore = Nothing

  Private mblnHasPrivilegedAccess As Nullable(Of Boolean)
  Private mobjSecurityPolicies As SecurityPolicies = Nothing


#End Region

#Region "Enumerations"

  Public Enum LifecycleChangeFlagEnum
    ClearException = 4194304
    Demote = 2097152
    Promote = 1048576
    Reset = 5242880
    SetException = 3145728
  End Enum

#End Region

#Region "Constructors"

  Public Sub New()

    MyBase.New()

    Try
      AddProperties()
      MyBase.ExportPath = DEFAULT_EXPORT_PATH
      ' <Removed by: Ernie at: 9/29/2014-11:18:17 AM on machine: ERNIE-THINK>
      '       MyBase.ImportPath = My.Settings.DefaultImportPath
      ' </Removed by: Ernie at: 9/29/2014-11:18:17 AM on machine: ERNIE-THINK>
      mobjSearch = CreateSearch()

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub New(ByVal lpConnectionString As String)

    MyBase.New(lpConnectionString)

    Try
      AddProperties()
      ParseConnectionString()
      mobjSearch = CreateSearch()

      'Login(SystemName, ServerName, UserName, Password)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

#Region "Provider Identification"

  Private Sub AddProperties()

    ' Add the properties here that you want to show up in the 'Create Data Source' dialog.

    Try
      '' Add the 'DomainName' property
      'MyBase.ProviderProperties.Add(New ProviderProperty(DOMAIN_NAME, GetType(System.String), True))

      ' Add the 'ServerName' property
      ProviderProperties.Add(New ProviderProperty(SERVER_NAME, GetType(String), True, , 1,
        "The P8 Content Engine server name or address.", , True))

      ' Add the 'PortNumber' property
      ProviderProperties.Add(New ProviderProperty(PORT_NUMBER, GetType(String), True, "9080", 2,
        "The port number used to connect to the P8 Content Engine."))

      ' Add the 'Protocol' property
      ProviderProperties.Add(New ProviderProperty(PROTOCOL, GetType(String), True, "http", 3,
        "The protocol used for the P8 Content Engine connection.", True, False))

      ' Add the 'PingToVerify' property
      ProviderProperties.Add(New ProviderProperty(PING_TO_VERIFY, GetType(String), True, True, 4,
        "The protocol used for the P8 Content Engine connection.", True, False))

      ' Add the 'TrustedConnection' property
      ProviderProperties.Add(New ProviderProperty(TRUSTED_CONNECTION, GetType(Boolean), False, "False", 5,
        "Specifies whether or not to authenticate using single sign on."))

      ' Add the 'UserName' property
      ProviderProperties.Add(New ProviderProperty(USER, GetType(String), False, , 6,
        "The user name to authenticate to P8.", , True))

      ' Add the 'Password' property
      ProviderProperties.Add(New ProviderProperty(PWD, GetType(String), False, , 7,
        "The password for the user login.", , True))

      ' Add the 'ObjectStoreName' property
      ' <Modified by: Ernie at 7/25/2014-6:47:19 AM on machine: ERNIE-THINK>
      'ProviderProperties.Add(New ProviderProperty(OBJECT_STORE_NAME, GetType(String), True, , 7, _
      '  "The object store to connect to.", True, True))
      ProviderProperties.Add(New RepositoryKeyProperty(OBJECT_STORE_NAME, , 8, "The object store to connect to."))
      ' </Modified by: Ernie at 7/25/2014-6:47:19 AM on machine: ERNIE-THINK>

      ' Add the 'PreserveIdOnMigrate' property
      ProviderProperties.Add(New ProviderProperty(PRESERVE_ID_ON_IMPORT, GetType(Boolean), False, "True", 9,
        "Specifies whether or not to attempt to seed the target document identifier with the incoming identifier."))

      ' Add the 'ExportSystemObjectValuedProperties' property
      ProviderProperties.Add(New ProviderProperty(EXPORT_SYSTEM_OBJECT_VALUED_PROPERTIES, GetType(Boolean), False, "False", 10,
        "Specifies whether or not to optionally export system defined object valued properties."))

      ' Add the 'Treat Object Ids As Numbers' property for DB2 based systems
      ' This popped up when we were working with the Army
      ' Ernie Bahr 8/25/2016
      ProviderProperties.Add(New ProviderProperty(TREAT_OBJECT_IDS_AS_NUMBERS, GetType(Boolean), False, TREAT_OBJECT_IDS_AS_NUMBERS_DEFAULT_VALUE, 11,
        "Specifies whether or not to treat object ids as numbers when generating document searches (usually only needed for DB2 based environments)."))


      ' Set the available action properties
      ' Add MajorVersion
      ActionProperties.Add(New ActionProperty("MajorVersion", False, PropertyType.ecmBoolean,
                                              "Determines whether or not a newly created version will be created as a major version or not."))

      ' Add SecurityPolicy
      ActionProperties.Add(New ActionProperty("SecurityPolicy", String.Empty, PropertyType.ecmString,
                                              "If specified, will set the P8 security policy for the newly added document."))



    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

#Region "Public Properties"

  Public ReadOnly Property TrustedConnection As Boolean
    Get
      Return mblnTrustedConnection
    End Get
  End Property

  Public ReadOnly Property ExportSystemObjectValuedProperties As Boolean
    Get
      Return mblnExportSystemObjectValuedProperties
    End Get
  End Property

  Public ReadOnly Property PreserveIdOnImport As Boolean
    Get
      Return mblnPreserveIdOnImport
    End Get
  End Property

  Public ReadOnly Property Domain As IDomain
    Get
      Return mobjDomain
    End Get
  End Property

  'Public ReadOnly Property DomainName() As String
  '  Get
  '    Return mstrDomainName
  '  End Get
  'End Property

  Public ReadOnly Property ServerName() As String
    Get
      Return mstrServerName
    End Get
  End Property

  Public ReadOnly Property ObjectStoreName() As String
    Get
      Return mstrObjectStoreName
    End Get
  End Property

  Public ReadOnly Property PingToVerify() As Boolean
    Get
      Return mblnPingToVerify
    End Get
  End Property

  Public ReadOnly Property PortNumber() As String
    Get
      Return mstrPortNumber
    End Get
  End Property

  Friend ReadOnly Property ObjectStore() As IObjectStore
    Get
      Return mobjObjectStore
    End Get
  End Property

  Public ReadOnly Property ContentExportPropertyExclusions() As String()
    Get

      Try

        If mstrContentExportPropertyExclusions Is Nothing Then
          mstrContentExportPropertyExclusions = GetAllContentExportPropertyExclusions()
        End If

        Return mstrContentExportPropertyExclusions

      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property


  ''' <summary>
  '''   Gets the folder delimiter used by a specific repository.
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Overrides ReadOnly Property FolderDelimiter() As String
    Get
      Return FOLDER_DELIMITER
    End Get
  End Property


  ''' <summary>
  '''   Gets a value specifying whether or
  '''   not the repository expects a leading
  '''   delimiter for all folder paths.
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Overrides ReadOnly Property LeadingFolderDelimiter() As Boolean
    Get
      Return True
    End Get
  End Property

  Public ReadOnly Property URL() As String
    Get

      Try
        Return CreateURL()

      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

  Public ReadOnly Property HasElevatedPrivileges As Nullable(Of Boolean)
    Get

      Try

        If mblnHasPrivilegedAccess.HasValue = False Then
          mblnHasPrivilegedAccess = GetHasPrivelegedAccess()
        End If

        Return mblnHasPrivilegedAccess

      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

  Friend ReadOnly Property SecurityPolicies As SecurityPolicies
    Get
      Try
        If mobjSecurityPolicies Is Nothing Then
          mobjSecurityPolicies = GetAllSecurityPolicies()
        End If
        Return mobjSecurityPolicies
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

#End Region

  '#Region "Private Properties"

  '  Private Property PropertyValueStrings As Dictionary(Of String, String)
  '    Get
  '      Try
  '        Return mobjPropertyValueStrings
  '      Catch ex As Exception
  '        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '        ' Re-throw the exception to the caller
  '        Throw
  '      End Try
  '    End Get
  '    Set(value As Dictionary(Of String, String))
  '      Try
  '        mobjPropertyValueStrings = value
  '      Catch ex As Exception
  '        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '        ' Re-throw the exception to the caller
  '        Throw
  '      End Try
  '    End Set
  '  End Property

  '#End Region

#Region "Public Overrides Methods"

  'Public Overrides ReadOnly Property 

#Region "Connect Methods"

  Public Overrides Sub Connect()

    Try

      ApplicationLogging.LogInformation(String.Format("ContentSourceInfo: {0}", ContentSource.ConnectionString))
      InitializeConnection()

    Catch RepNotAvailEx As RepositoryNotAvailableException
      ApplicationLogging.LogException(RepNotAvailEx, MethodBase.GetCurrentMethod)
      IsConnected = False
      Throw
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      IsConnected = False
      Throw New ApplicationException("Connection Failed: " & Helper.FormatCallStack(ex))
    End Try
  End Sub

  Public Overrides Sub Connect(ByVal ConnectionString As String)

    Try
      MyBase.ConnectionString = ConnectionString
      ParseConnectionString()
      ApplicationLogging.LogInformation(String.Format("ContentSourceInfo From ConnectionString: {0}", ContentSource.ConnectionString))
      InitializeConnection()

    Catch RepNotAvailEx As RepositoryNotAvailableException
      ApplicationLogging.LogException(RepNotAvailEx, MethodBase.GetCurrentMethod)
      IsConnected = False
      Throw
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      IsConnected = False
      Throw New ApplicationException("Connection Failed: " & Helper.FormatCallStack(ex))
    End Try
  End Sub

  Public Overrides Sub Connect(ByVal lpContentSource As ContentSource)

    Try
      ApplicationLogging.LogInformation(String.Format("ContentSourceInfo From ContentSource: {0}", lpContentSource.ConnectionString))
      'Connect(lpContentSource.ConnectionString)
      InitializeProvider(lpContentSource)

      'InitializeProperties()

      'InitializeConnection()
      IsConnected = True
    Catch RepNotAvailEx As RepositoryNotAvailableException
      ApplicationLogging.LogException(RepNotAvailEx, MethodBase.GetCurrentMethod)
      IsConnected = False
      Throw
    Catch RepAuthEx As RepositoryAuthenticationException
      ApplicationLogging.LogException(RepAuthEx, MethodBase.GetCurrentMethod)
      IsConnected = False
      Throw
    Catch RepNotConnectedEx As RepositoryNotConnectedException
      ApplicationLogging.LogException(RepNotConnectedEx, MethodBase.GetCurrentMethod)
      IsConnected = False
      Throw
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      IsConnected = False
      Throw New ApplicationException("Connection Failed: " & Helper.FormatCallStack(ex))
    End Try
  End Sub

#End Region

  Public Overrides Function Login(ByVal lpUserName As String,
                                  ByVal lpPassword As String) As Boolean
  End Function

  Public Overrides Function Logout() As Boolean
  End Function

  'Public Overrides ReadOnly Property Feature As FeatureEnum
  '  Get
  '    Try
  '      Return FeatureEnum.CENetProvider
  '    Catch Ex As Exception
  '      ApplicationLogging.LogException(Ex, MethodBase.GetCurrentMethod)
  '      ' Re-throw the exception to the caller
  '      Throw
  '    End Try
  '  End Get
  'End Property

  Public Overrides ReadOnly Property ProviderSystem() As DProviders.ProviderSystem
    Get
      Return mobjSystem
    End Get
  End Property


  ''' <summary>
  '''   Find and remove invalid properties so user doesn't have to create a transform to delete each one
  ''' </summary>
  ''' <param name="lpDocument"></param>
  ''' <param name="lpPropertyScope"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Overrides Function FindInvalidProperties(ByVal lpDocument As Document,
                                                  ByVal lpPropertyScope As PropertyScope) As DCore.IProperties

    Try

      Dim lobjInvalidProperties As InvalidProperties = MyBase.FindInvalidProperties(lpDocument, lpPropertyScope)
      Dim lobjFoundProperty As ECMProperty = Nothing

      'If Me.ActionProperties(ACTION_SET_SYSTEM_PROPERTIES).Value = True Then
      '  ' Remove the settable system properties from the list
      '  lobjInvalidProperties.Remove(PROP_CREATOR)
      '  lobjInvalidProperties.Remove(PROP_CREATE_DATE)
      '  lobjInvalidProperties.Remove(PROP_MODIFY_USER)
      '  lobjInvalidProperties.Remove(PROP_MODIFY_DATE)
      '  lobjInvalidProperties.Remove(PROP_MIME_TYPE)
      'End If

      lobjInvalidProperties.Remove(PropertyNames.ID)

      lobjInvalidProperties.Remove(PropertyNames.MAJOR_VERSION_NUMBER)
      lobjInvalidProperties.Remove(PropertyNames.MINOR_VERSION_NUMBER)
      lobjInvalidProperties.Remove(PropertyNames.FOLDERS_FILED_IN)

      ' If the current user has elevated privileges we can leave some of the system properties
      If HasElevatedPrivileges Then
        lobjInvalidProperties.Remove(PropertyNames.CREATOR)
        lobjInvalidProperties.Remove(PropertyNames.DATE_CREATED)
        lobjInvalidProperties.Remove(PropertyNames.LAST_MODIFIER)
        lobjInvalidProperties.Remove(PropertyNames.DATE_LAST_MODIFIED)
        lobjInvalidProperties.Remove(PropertyNames.DATE_CHECKED_IN)
      End If

      For Each lobjInvalidProperty As InvalidProperty In lobjInvalidProperties

        Select Case lobjInvalidProperty.Scope

          Case InvalidProperty.InvalidPropertyScope.Document
            lpDocument.DeleteProperty(PropertyScope.DocumentProperty, lobjInvalidProperty.Name)

            If LogInvalidPropertyRemovals Then
              ApplicationLogging.WriteLogEntry(
                String.Format("Removed invalid property '{0}' from document '{1}'", lobjInvalidProperty.Name,
                              lpDocument.Name), TraceEventType.Information, 5251)
            End If

          Case InvalidProperty.InvalidPropertyScope.AllExceptFirstVersion
            ' Skip for now
            Continue For

            For lintVersionCounter As Integer = 1 To lpDocument.Versions.Count - 1
              lpDocument.Versions(lintVersionCounter).Properties.Delete(lobjInvalidProperty.Name)

              If LogInvalidPropertyRemovals Then
                ApplicationLogging.WriteLogEntry(
                  String.Format("Removed invalid property '{0}' from version id {1} of document '{2}'",
                                lobjInvalidProperty.Name, lintVersionCounter, lpDocument.Name),
                  TraceEventType.Information, 5252)
              End If

            Next

            'lpDocument.DeleteProperty(PropertyScope.DocumentProperty, lobjInvalidProperty.Name)
            'ApplicationLogging.WriteLogEntry(String.Format("Removed invalid property '{0}' from document '{1}'", _
            '                                               lobjInvalidProperty.Name, lpDocument.Name), TraceEventType.Information, 4251)

          Case InvalidProperty.InvalidPropertyScope.AllVersions

            For Each lobjVersion As DCore.Version In lpDocument.Versions
              lobjVersion.Properties.Delete(lobjInvalidProperty.Name)

              If LogInvalidPropertyRemovals Then
                ApplicationLogging.WriteLogEntry(
                  String.Format("Removed invalid property '{0}' from version id {1} of document '{2}'",
                                lobjInvalidProperty.Name, lobjVersion.ID, lpDocument.Name), TraceEventType.Information,
                  5253)
              End If

            Next

          Case InvalidProperty.InvalidPropertyScope.FirstVersion
            lpDocument.FirstVersion.Properties.Delete(lobjInvalidProperty.Name)

            If LogInvalidPropertyRemovals Then
              ApplicationLogging.WriteLogEntry(
                String.Format("Removed invalid property '{0}' from version id {1} of document '{2}'",
                              lobjInvalidProperty.Name, lpDocument.FirstVersion.ID, lpDocument.Name),
                TraceEventType.Information, 5254)
            End If

            ApplicationLogging.WriteLogEntry(
              String.Format("Removed invalid property '{0}' from version id {1} of document '{2}'",
                            lobjInvalidProperty.Name, lpDocument.FirstVersion.ID, lpDocument.Name),
              TraceEventType.Information, 5254)
        End Select

        'lpDocument.DeleteProperty(PropertyScope.BothDocumentAndVersionProperties, lobjInvalidProperty.Name)
        'ApplicationLogging.WriteLogEntry(String.Format("Removed invalid property '{0}' from document '{1}'", lobjInvalidProperty.Name, lpDocument.Name), TraceEventType.Information, 4251)
      Next

      'If ContentEngineVersion = CEVersion.ThreeFive Then
      '  RemoveInvalidProperty(PROP_CREATE_DATE, lpDocument, lobjInvalidProperties)
      '  RemoveInvalidProperty(PROP_CREATOR, lpDocument, lobjInvalidProperties)
      '  RemoveInvalidProperty(PROP_MODIFY_USER, lpDocument, lobjInvalidProperties)
      '  RemoveInvalidProperty(PROP_MODIFY_DATE, lpDocument, lobjInvalidProperties)
      'End If

      'Don't need to do this because these properties are read-only in P8.
      'RemoveInvalidProperty(PROP_OBJECT_NAME, lpDocument, lobjInvalidProperties)
      'RemoveInvalidProperty(PROP_OWNER, lpDocument, lobjInvalidProperties)
      'RemoveInvalidProperty(PROP_VERSION_STATUS, lpDocument, lobjInvalidProperties)
      'RemoveInvalidProperty(PROP_VERSION_ID, lpDocument, lobjInvalidProperties)
      'RemoveInvalidProperty(PROP_IS_CURRENT_VERSION, lpDocument, lobjInvalidProperties)
      'RemoveInvalidProperty(PROP_MINOR_VERSION_NUMBER, lpDocument, lobjInvalidProperties)
      'RemoveInvalidProperty(PROP_MAJOR_VERSION_NUMBER, lpDocument, lobjInvalidProperties)
      'RemoveInvalidProperty(PROP_VERSION_SERIES_ID, lpDocument, lobjInvalidProperties)

      Return lobjInvalidProperties

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Overrides Function ReadyToGetAvailableValues(ByVal lpProviderProperty As ProviderProperty) As Boolean
    Try
      Dim lblnReturnValue As Boolean = MyBase.ReadyToGetAvailableValues(lpProviderProperty)

      Select Case lpProviderProperty.PropertyName
        Case OBJECT_STORE_NAME
          lblnReturnValue = True
          If String.IsNullOrEmpty(ServerName) Then
            lblnReturnValue = False
          End If
          If String.IsNullOrEmpty(UserName) Then
            lblnReturnValue = False
          End If
          If String.IsNullOrEmpty(Password) Then
            lblnReturnValue = False
          End If
        Case Else
          lblnReturnValue = False
      End Select

      Return lblnReturnValue

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  ''' <summary>
  '''   Gets the available values to set for the specified provider property.
  ''' </summary>
  ''' <param name="lpProviderProperty"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Overrides Function GetAvailableValues(ByVal lpProviderProperty As ProviderProperty) As IEnumerable(Of String)

    Try

      If lpProviderProperty.SupportsValueList = False Then
        Throw New PropertyDoesNotSupportValueListException(lpProviderProperty)
      End If

      Dim lobjReturnValues As List(Of String) = MyBase.GetAvailableValues(lpProviderProperty)

      Select Case lpProviderProperty.PropertyName

        Case OBJECT_STORE_NAME
          lobjReturnValues.Clear()

          Dim lobjObjectStoreIdentifiers As RepositoryIdentifiers = GetRepositories() ' GetObjectStoreIdentifiers()

          For Each lobjObjectStoreIdentifier As RepositoryIdentifier In lobjObjectStoreIdentifiers
            lobjReturnValues.Add(lobjObjectStoreIdentifier.Name)
          Next

          lobjReturnValues.Sort()
        Case PROTOCOL
          lobjReturnValues.Clear()
          lobjReturnValues.Add("http")
          lobjReturnValues.Add("https")
      End Select

      Return lobjReturnValues

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region


  Friend Function CreateECMProperty(ByVal lpIProperty As [Property].IProperty) As ECMProperty

    Dim lobjECMProperty As ECMProperty

    'Dim lobjIList As IList
    'Dim lobjECMValues As CTS.Values
    Dim lobjPropertyValue As Object = Nothing
    Dim lobjValueType As Type = Nothing
    Dim lstrPropertyValue As String = String.Empty
    Dim lstrPropertyObjectType As String = String.Empty
    Dim lstrPropertyName As String = String.Empty
    'Dim lstrPropertyTypeParts() As String

    Try

      lstrPropertyName = lpIProperty.GetPropertyName

      If lstrPropertyName = "Versions" AndAlso Helper.CallStackContainsMethodName("ExportVersion") Then
        ' Skip this, it is a hog
        Return Nothing
      End If
      'Debug.Print(lstrPropertyName)

      Dim lobjClassificationProperty As ClassificationProperty = ContentProperty(lstrPropertyName)
      If lobjClassificationProperty IsNot Nothing Then
        lobjECMProperty = lobjClassificationProperty.ToECMProperty(False)
      Else
        lobjECMProperty = Nothing
        Return lobjECMProperty
      End If

      'If ContentProperties.Contains(lstrPropertyName) Then
      '  'lobjECMProperty = ContentProperties.ItemByName(lstrPropertyName).ToECMProperty
      '  lobjECMProperty = ContentProperties.ItemByName(lstrPropertyName).ToECMProperty(False)
      'Else
      '  lobjECMProperty = Nothing
      '  Return lobjECMProperty
      'End If

      'Debug.Print(CType(lpIProperty, Object).GetType.ToString)
      'Debug.Print(lpIProperty.GetPropertyName)

      Select Case lobjECMProperty.Cardinality

        Case DCore.Cardinality.ecmSingleValued

          Select Case lobjECMProperty.Type

            Case PropertyType.ecmString
              lobjECMProperty.Value = lpIProperty.GetStringValue

            Case PropertyType.ecmBinary
              lobjECMProperty.Value = lpIProperty.GetBinaryValue

            Case PropertyType.ecmBoolean
              lobjECMProperty.Value = lpIProperty.GetBooleanValue

            Case PropertyType.ecmDate
              lobjECMProperty.Value = lpIProperty.GetDateTimeValue

            Case PropertyType.ecmDouble
              lobjECMProperty.Value = lpIProperty.GetFloat64Value

            Case PropertyType.ecmGuid

              Dim lobjValue As Object = lpIProperty.GetIdValue

              If lobjValue IsNot Nothing Then
                lobjECMProperty.Value = New Guid(lobjValue.ToString)
              End If

            Case PropertyType.ecmLong
              lobjECMProperty.Value = lpIProperty.GetInteger32Value


            Case PropertyType.ecmObject

              If lobjClassificationProperty.IsSystemProperty AndAlso
                 ExportSystemObjectValuedProperties = False Then
                ' We need to skip this property
                Return Nothing
              End If
              Dim lobjObjectValue As Object = lpIProperty.GetObjectValue

              If lobjObjectValue IsNot Nothing Then
                lobjECMProperty.Value = GetObjectDescriptor(lobjObjectValue)
              End If

          End Select

        Case DCore.Cardinality.ecmMultiValued

          Select Case lobjECMProperty.Type

            Case PropertyType.ecmString

              ' TODO: Implement retrieval of multi-valued attributes
              ' Until then this is likely to break
              'lobjECMProperty.Values = lpIProperty.GetStringListValue
              Dim lobjValues As IStringList = lpIProperty.GetStringListValue

              For Each lstrValue As String In lobjValues
                DirectCast(lobjECMProperty, MultiValueStringProperty).Values.AddString(lstrValue)
              Next

            Case PropertyType.ecmDate
              Dim lobjValues As IDateTimeList = lpIProperty.GetDateTimeListValue

              For Each ldatValue As Date In lobjValues
                DirectCast(lobjECMProperty, MultiValueDateTimeProperty).Values.AddDate(ldatValue)
              Next

            Case PropertyType.ecmBoolean
              Dim lobjValues As IBooleanList = lpIProperty.GetBooleanListValue

              For Each lblnValue As Boolean In lobjValues
                DirectCast(lobjECMProperty, MultiValueBooleanProperty).Values.AddBoolean(lblnValue)
              Next

            Case PropertyType.ecmDouble
              Dim lobjValues As IFloat64List = lpIProperty.GetFloat64ListValue

              For Each ldatValue As Double In lobjValues
                DirectCast(lobjECMProperty, MultiValueDoubleProperty).Values.AddDouble(ldatValue)
              Next

            Case PropertyType.ecmGuid
              Dim lobjValues As IIdList = lpIProperty.GetIdListValue

              For Each ldatValue As Id In lobjValues
                DirectCast(lobjECMProperty, MultiValueGuidProperty).Values.AddString(ldatValue.ToString)
              Next

            Case PropertyType.ecmLong
              Dim lobjValues As IInteger32List = lpIProperty.GetInteger32ListValue

              For Each ldatValue As Long In lobjValues
                DirectCast(lobjECMProperty, MultiValueLongProperty).Values.AddLong(ldatValue)
              Next

            Case PropertyType.ecmObject

              If lobjClassificationProperty.IsSystemProperty AndAlso
                 ExportSystemObjectValuedProperties = False Then
                ' We need to skip this property
                Return Nothing
              End If

              Dim lstrP8TypeName As String = lpIProperty.GetType.Name

              Select Case lstrP8TypeName

                Case "PropertyEngineObjectSetImpl"

                  Dim lobjValues As IIndependentObjectSet = lpIProperty.GetIndependentObjectSetValue

                  For Each lstrValue As Object In lobjValues

                    If TypeOf lstrValue IsNot String Then
                      lstrValue = GetObjectDescriptor(lstrValue)
                    End If

                    If Not DirectCast(lobjECMProperty, MultiValueObjectProperty).Values.Contains(lstrValue) Then
                      DirectCast(lobjECMProperty, MultiValueObjectProperty).Values.AddString(lstrValue)
                    End If

                  Next

              End Select

              'Dim lobjValues As IDependentObjectList = lpIProperty.GetDependentObjectListValue
              'For Each lstrValue As String In lobjValues
              '  DirectCast(lobjECMProperty, MultiValueStringProperty).Values.AddString(lstrValue)
              'Next

          End Select

      End Select

      If lobjECMProperty IsNot Nothing Then
        Return lobjECMProperty

      Else
        Return Nothing
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod, 62342)
      Debug.WriteLine("Error with property: " & lstrPropertyName & vbCrLf & vbCrLf & ex.Message)
      Return Nothing
    Finally
      lstrPropertyName = Nothing
    End Try
  End Function

  Friend Function GetObjectDescriptor(ByVal lpObject As Object) As String

    Try

      Dim lobjIdProperty As PropertyInfo = lpObject.GetType.GetProperty("Id")
      Dim lobjNameProperty As PropertyInfo = lpObject.GetType.GetProperty("Name")
      Dim lobjDisplayNameProperty As PropertyInfo = lpObject.GetType.GetProperty("DisplayName")
      Dim lobjClassDescriptionProperty As PropertyInfo = lpObject.GetType.GetProperty("ClassDescription")

      Dim lstrIdValue As Object = Nothing
      Dim lstrNameValue As String = Nothing
      Dim lstrDisplayNameValue As String = Nothing

      Dim lstrStringBuilder As New StringBuilder

      If lobjIdProperty IsNot Nothing Then
        lstrIdValue = lobjIdProperty.GetValue(lpObject, Nothing)

        If lstrIdValue IsNot Nothing Then
          lstrStringBuilder.Append(lstrIdValue.ToString)

          If lobjClassDescriptionProperty IsNot Nothing Then
            Dim lobjClassName As String = lpObject.ClassDescription.SymbolicName
            If Not String.IsNullOrEmpty(lobjClassName) Then
              lstrStringBuilder.AppendFormat(":{0}", lobjClassName)
            End If
          End If


          If lobjDisplayNameProperty IsNot Nothing Then
            lstrDisplayNameValue = lobjDisplayNameProperty.GetValue(lpObject, Nothing)

            If Not String.IsNullOrEmpty(lstrDisplayNameValue) Then
              lstrStringBuilder.AppendFormat(" - {0}", lstrDisplayNameValue)
            End If

          ElseIf lobjNameProperty IsNot Nothing Then
            lstrNameValue = lobjNameProperty.GetValue(lpObject, Nothing)

            If Not String.IsNullOrEmpty(lstrNameValue) Then
              lstrStringBuilder.AppendFormat(" - {0}", lstrNameValue)
            End If

          End If

        End If

      End If

      Return lstrStringBuilder.ToString

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Sub ExportFolderEventCallBack(ByVal sender As Object,
                                       ByVal e As EventArgs)

    Try

      If TypeOf e Is FolderDocumentExportedEventArgs Then

        Dim Args As FolderDocumentExportedEventArgs = CType(e, FolderDocumentExportedEventArgs)
        RaiseEvent FolderDocumentExported(sender, Args)

      ElseIf TypeOf e Is FolderExportedEventArgs Then

        Dim Args As FolderExportedEventArgs = CType(e, FolderExportedEventArgs)
        RaiseEvent FolderExported(sender, Args)
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Overrides Function GetFolder(ByVal lpFolderPath As String,
                                      ByVal lpMaxContentCount As Long) As DProviders.IFolder

    Try
      Return New CENetFolder(lpFolderPath, Me, lpMaxContentCount)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function


  Private Function CreateExcludedPropertiesFilter() As PropertyFilter

    Try

      ' Create the property filter to exclude the properties we do not want
      Dim lobjPropertyFilter As New PropertyFilter
      Dim lstrExclusions() As String = ContentExportPropertyExclusions

      'For Each lstrPropertyName As String In ContentExportPropertyExclusions
      For lintExclusionCounter As Int16 = 0 To lstrExclusions.Length - 1
        lobjPropertyFilter.AddExcludeProperty(lstrExclusions(lintExclusionCounter))
      Next

      Return lobjPropertyFilter

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function


  ''' <summary>
  '''   Retrieves the document object and only includes properties in lpPropertyFilter
  ''' </summary>
  ''' <param name="lpId"></param>
  ''' <param name="lpIncludePropertyFilter"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function GetIDocument(ByVal lpId As String,
                                ByVal lpIncludePropertyFilter As PropertyFilter) As IDocument

    Dim lobjIDocument As IDocument = Nothing

    Try

      If Me.DocumentExists(lpId) = False Then
        Throw New DocumentDoesNotExistException(lpId)
      End If

      lobjIDocument = Factory.Document.FetchInstance(ObjectStore, lpId, lpIncludePropertyFilter)

      If lobjIDocument Is Nothing Then
        Throw New ArgumentException("Document does not exist", lpId)

      Else
        Return lobjIDocument
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      Throw New ArgumentException("Document does not exist", lpId, ex)
    End Try
  End Function

  Private Function GetIAnnotation(ByVal lpId As String) As IAnnotation

    Dim lobjIAnnotation As IAnnotation = Nothing

    Try

      If Me.DocumentExists(lpId, "Annotation") = False Then
        Throw New ItemDoesNotExistException(lpId)
      End If

      lobjIAnnotation = ObjectStore.FetchObject("Annotation", lpId, Nothing)

      Return lobjIAnnotation

    Catch NoItemExistsEx As ItemDoesNotExistException
      ApplicationLogging.LogException(NoItemExistsEx, MethodBase.GetCurrentMethod)
      Throw NoItemExistsEx
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      Throw New DocumentException(ex.Message, lpId, ex)
    End Try
  End Function

  Private Function GetIDocument(ByVal lpId As String) As IDocument

    Dim lobjIDocument As IDocument = Nothing

    Try

      '' Create the property filter to exclude the properties we do not want
      'Dim lobjPropertyFilter As New FileNet.Api.Property.PropertyFilter

      'lobjPropertyFilter = CreateExcludedPropertiesFilter()

      'lobjIDocument = ObjectStore.FetchObject("Document", lpId, lobjPropertyFilter)
      If Me.DocumentExists(lpId) = False Then
        Throw New DocumentDoesNotExistException(lpId)
      End If

      lobjIDocument = ObjectStore.FetchObject("Document", lpId, Nothing)

      Return lobjIDocument

    Catch NoDocExistsEx As DocumentDoesNotExistException
      ApplicationLogging.LogException(NoDocExistsEx, MethodBase.GetCurrentMethod)
      Throw NoDocExistsEx
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      Throw New DocumentException(ex.Message, lpId, ex)
    End Try
  End Function

  Private Function GetIFolder(ByVal lpId As String) As FileNet.Api.Core.IFolder

    Try

      If FolderIDExists(lpId) Then
        Return Factory.Folder.FetchInstance(ObjectStore, lpId, Nothing)
      Else
        Throw New FolderDoesNotExistException(lpId)
      End If

    Catch NoFolderEx As FolderDoesNotExistException
      ApplicationLogging.LogException(NoFolderEx, MethodBase.GetCurrentMethod)
      Throw NoFolderEx
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      Throw New ArgumentException(ex.Message, lpId, ex)
    End Try
  End Function

  Private Function GetIFolder(ByVal lpId As String,
                              ByVal lpIncludePropertyFilter As PropertyFilter) As FileNet.Api.Core.IFolder

    Try

      If FolderIDExists(lpId) Then
        Return Factory.Folder.FetchInstance(ObjectStore, lpId, lpIncludePropertyFilter)
      Else
        Throw New FolderDoesNotExistException(lpId)
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      Throw New ArgumentException("Document does not exist", lpId, ex)
    End Try
  End Function

#If SupportAnnotations Then

  Private Function RetrieveAnnotationContent(ByVal lpAnnotation As IAnnotation) As Stream
    Return Me.RetrieveAnnotationContent(lpAnnotation, 0)
  End Function

  Private Function RetrieveAnnotationContent(ByVal lpAnnotation As IAnnotation, ByVal lpContentIndex As Integer) _
    As Stream
    Dim lobjResult As Stream = Nothing
    Try
      ArgumentNullException.ThrowIfNull(lpAnnotation)

      If lpAnnotation.ContentElements.Count < lpContentIndex + 1 Then
        ApplicationLogging.WriteLogEntry(String.Format("No content element '{0}' for annotation id '{1}'",
                                                       lpContentIndex, lpAnnotation.Id.ToString()))
        Exit Try
      End If

      lobjResult = lpAnnotation.AccessContentStream(lpContentIndex)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try
    Return lobjResult
  End Function

#End If


  Private Function FileDocument(ByRef lpIdocument As IDocument,
                                ByVal lpPath As String,
                                Optional ByVal lpContainmentName As String = "",
                                Optional ByRef lpErrorMessage As String = "") As FileNet.Api.Core.IFolder

    Try

      Dim lobjIFolder As FileNet.Api.Core.IFolder = Me.GetFolderByPath(lpPath)
      Dim lstrContainmentName As String = ""

      If lpContainmentName.Length = 0 Then
        lstrContainmentName = GetDocumentTitle(lpIdocument) ' lpIdocument.Properties("DocumentTitle").ToString

      Else
        lstrContainmentName = lpContainmentName
      End If

      lstrContainmentName = CleanContainmentName(lstrContainmentName)

      'lobjIFolder.File(lpIdocument, Constants.AutoUniqueName.AUTO_UNIQUE, _
      '  lstrContainmentName, _
      '  Constants.DefineSecurityParentage.DO_NOT_DEFINE_SECURITY_PARENTAGE).Save(Constants.RefreshMode.REFRESH)

      ' lobjIFolder.File(lpIdocument, AutoUniqueName.AUTO_UNIQUE, lstrContainmentName, DefineSecurityParentage.DO_NOT_DEFINE_SECURITY_PARENTAGE).Save(RefreshMode.NO_REFRESH)

      Dim lobjNewRCR As IReferentialContainmentRelationship = lobjIFolder.File(lpIdocument, AutoUniqueName.AUTO_UNIQUE,
                                                                               lstrContainmentName,
                                                                               DefineSecurityParentage.
                                                                                DO_NOT_DEFINE_SECURITY_PARENTAGE)

      If lobjNewRCR IsNot Nothing Then
        lobjNewRCR.Save(RefreshMode.NO_REFRESH)
      Else
        lpErrorMessage = String.Format("Failed to file document '{0}' in folder '{1}'", lpIdocument.Name, lpPath)
        ApplicationLogging.WriteLogEntry(lpErrorMessage, MethodBase.GetCurrentMethod, TraceEventType.Warning, 62887)
      End If
      'lobjIFolder.Save(Constants.RefreshMode.REFRESH)

      Return lobjIFolder

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      lpErrorMessage = lpErrorMessage =
                       String.Format("Failed to file document '{0}' in folder '{1}': {2}", lpIdocument.Name, lpPath,
                                     ex.StackTrace)
      ApplicationLogging.WriteLogEntry(lpErrorMessage, MethodBase.GetCurrentMethod, TraceEventType.Warning, 62887)
      Return Nothing
    End Try
  End Function


  ''' <summary>
  '''   Checks to see if the version is marked as a major version
  ''' </summary>
  ''' <param name="lpVersion">CTS Version object</param>
  ''' <returns>True if the version contains a boolean property called MajorVersion that is set to true</returns>
  ''' <remarks></remarks>
  Private Function IsMajorVersion(ByVal lpVersion As DCore.Version,
                                  Optional ByVal lpDeleteProperty As Boolean = True) As Boolean

    Try

      If (lpVersion.Properties.PropertyExists(MAJOR_VERSION)) AndAlso lpVersion.Properties(MAJOR_VERSION).Value = True _
        Then

        If lpDeleteProperty = True Then
          lpVersion.Properties.Delete(MAJOR_VERSION)
        End If

        Return True

      ElseIf _
        (lpVersion.Properties.PropertyExists(PROP_MAJOR_VERSION_NUMBER)) AndAlso
        (lpVersion.Properties.PropertyExists(PROP_MINOR_VERSION_NUMBER)) Then

        Dim lintMajorVersionNumber As Integer = lpVersion.Properties(PROP_MAJOR_VERSION_NUMBER).Value
        Dim lintMinorVersionNumber As Integer = lpVersion.Properties(PROP_MINOR_VERSION_NUMBER).Value

        If lintMajorVersionNumber > 0 AndAlso lintMinorVersionNumber = 0 Then
          Return True

        ElseIf lintMinorVersionNumber > 0 Then
          Return False

        Else
          Return False
        End If

      Else
        Return False
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw

    Finally

    End Try
  End Function

  Private Function GetDocumentTitle(ByRef lpIdocument As IDocument) As String

    Try

      Dim lstrDocumentTitle As String = Nothing
      Dim lstrPropertyName As String = String.Empty

      ' Try to get the document title from the properties collection.
      For Each lobjIProperty As [Property].IProperty In lpIdocument.Properties
        lstrPropertyName = lobjIProperty.GetPropertyName

        If String.Compare(lstrPropertyName, "DocumentTitle", True) = 0 Then
          lstrDocumentTitle = lobjIProperty.GetStringValue
          Exit For
        End If

      Next

      If lstrDocumentTitle Is Nothing Then
        lstrDocumentTitle = String.Empty

        If lpIdocument.Id IsNot Nothing Then
          ApplicationLogging.WriteLogEntry(
            String.Format("Unable to get document title for document '{0}'.", lpIdocument.Id.ToString),
            TraceEventType.Warning, 62341)

        Else
          ApplicationLogging.WriteLogEntry("Unable to get document title for document.", TraceEventType.Warning, 62342)
        End If

      End If

      Return lstrDocumentTitle

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function CleanContainmentName(ByVal lpContainmentName As String) As String

    Try
      Static lstrInvalidPathChars As Char() = Path.GetInvalidFileNameChars

      If lstrInvalidPathChars.Length = 0 Then
        lstrInvalidPathChars = Path.GetInvalidFileNameChars
      End If

      Dim lstrNewValue As String
      lstrNewValue = lpContainmentName

      For lintCharCounter As Int16 = 0 To lstrInvalidPathChars.Length - 1
        lstrNewValue = lstrNewValue.Replace(lstrInvalidPathChars(lintCharCounter), "")
      Next

      Return lstrNewValue

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function IsPendingCreate(ByVal lpP8Document As IDocument) As Boolean

    Try

      Dim lobjPendingActions As IEnumerable(Of PendingAction)
      lobjPendingActions = lpP8Document.GetPendingActions()

      If lobjPendingActions.Any() Then

        If TypeOf lobjPendingActions(0) Is Create Then
          Return True
        End If

      End If

      Return False

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Friend Function AddP8Document(ByVal lpDocument As Document,
                                ByRef lpErrorMessage As String, lpVersionType As VersionTypeEnum) As IDocument
    Try
      Return AddP8Document(lpDocument, True, lpErrorMessage, lpVersionType)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Friend Function AddP8Document(ByVal lpDocument As Document, lpSetPermissions As Boolean,
                                ByRef lpErrorMessage As String, lpVersionType As VersionTypeEnum) As IDocument

    ' BMK AddDocument
    Dim lobjP8Document As IDocument = Nothing

    Try

      Dim lobjFirstVersion As DCore.Version = lpDocument.FirstVersion

      Dim lstrLastModifer As String = Me.UserName
      Dim ldatDateLastModified As Date = Now
      Dim ldatDateCheckedIn As Date = Now

      If HasElevatedPrivileges Then

        ' If we have last modifier and date last modified in the supplied document let's try to use it
        If lobjFirstVersion.Properties.PropertyExists(PROP_MODIFY_USER) Then

          Dim lobjLastModifierProperty As ECMProperty = lobjFirstVersion.Properties(PROP_MODIFY_USER)

          If lobjLastModifierProperty.HasValue Then
            lstrLastModifer = lobjLastModifierProperty.Value
          End If

        End If

        If lobjFirstVersion.Properties.PropertyExists(PROP_MODIFY_DATE) Then

          Dim lobjLastModifiedProperty As ECMProperty = lobjFirstVersion.Properties(PROP_MODIFY_DATE)

          If lobjLastModifiedProperty.HasValue Then
            ldatDateLastModified = lobjLastModifiedProperty.Value
          End If

        End If

        If lobjFirstVersion.Properties.PropertyExists(PROP_CHECK_IN_DATE) Then

          Dim lobjDateCheckedInProperty As ECMProperty = lobjFirstVersion.Properties(PROP_CHECK_IN_DATE)

          If lobjDateCheckedInProperty.HasValue Then
            ldatDateCheckedIn = lobjDateCheckedInProperty.Value
          End If

        End If

      End If

      ' Get the candidate document class
      Dim lobjCandidateClass As DocumentClass = DocumentClass(lpDocument.DocumentClass)

      If lobjCandidateClass Is Nothing Then
        Throw New DocumentClassDoesNotExistException(lpDocument.DocumentClass,
                                                     String.Format(
                                                       "There is no document class named '{0}' defined in the object store '{1}'",
                                                       lpDocument.DocumentClass, ObjectStoreName))
      End If

      Dim lblnIsMajorVersion As Boolean

      Select Case lpVersionType
        Case VersionTypeEnum.Unspecified
          lblnIsMajorVersion = IsMajorVersion(lobjFirstVersion, True)
        Case VersionTypeEnum.Major
          lblnIsMajorVersion = True
        Case VersionTypeEnum.Minor
          lblnIsMajorVersion = False
      End Select

      'Dim lobjCandidateIdGuid As Guid
      'Dim lobjCandidateId As FileNet.Api.Util.Id = Nothing

      'If PreserveIdOnImport AndAlso Helper.IsGuid(lpDocument.ID, lobjCandidateIdGuid) Then
      '  ' Check to see if the guid has already been used in this object store
      '  If DocumentExists(lpDocument.ID) = False Then
      '    lobjCandidateId = New FileNet.Api.Util.Id(lobjCandidateIdGuid)
      '  Else
      '    ApplicationLogging.WriteLogEntry( _
      '      String.Format("Unable to preserve id for incoming document {0}, the identifier is already in use in object store '{1}'.", _
      '                    lpDocument.ID, ObjectStoreName), Reflection.MethodBase.GetCurrentMethod, TraceEventType.Warning, 63401)
      '  End If
      'End If

      Dim lobjCandidateId As Id = GetVersionId(lobjFirstVersion)
      Dim lobjVersionSeriesId As Id = GetVersionSeriesId(lobjFirstVersion)

      ' Check to make sure we have a document class
      If Not String.IsNullOrEmpty(lpDocument.DocumentClass) Then

        If lobjCandidateId IsNot Nothing Then

          Try
            If lobjVersionSeriesId IsNot Nothing Then
              lobjP8Document = Factory.Document.CreateInstance(mobjObjectStore, lobjCandidateClass.Name, lobjCandidateId, lobjVersionSeriesId, ReservationType.OBJECT_STORE_DEFAULT)
            Else
              lobjP8Document = Factory.Document.CreateInstance(mobjObjectStore, lobjCandidateClass.Name, lobjCandidateId)
            End If

          Catch ex As Exception
            ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
            lobjP8Document = Factory.Document.CreateInstance(mobjObjectStore, lobjCandidateClass.Name)
          End Try

        Else
          lobjP8Document = Factory.Document.CreateInstance(mobjObjectStore, lobjCandidateClass.Name)
        End If

      Else
        ' Make a notation in the log and add as the default class
        ApplicationLogging.WriteLogEntry(
          "No document class specified for the new document, using the default document class.",
          MethodBase.GetCurrentMethod, TraceEventType.Warning, 64512)
        lobjP8Document = Factory.Document.CreateInstance(mobjObjectStore, Nothing)
      End If

      ' Set the properties needed on create
      For Each lobjCDFVersionProperty As ECMProperty In lobjFirstVersion.Properties

        If lobjCDFVersionProperty.HasValue Then
          SetPropertyValue(lobjP8Document, lobjCandidateClass, lobjCDFVersionProperty, True)
        End If

      Next

      ' Set the Content Elements
      lobjP8Document.ContentElements = CreateContentElementList(lobjFirstVersion.Contents)

      ' Set the permissions if we have them and we were asked to.
      ' <Modified by: Ernie Bahr at 9/13/2012-10:27:31 on machine: ERNIEBAHR-THINK>
      ' Updated to only set permissions if the SetPermissions parameter is set to true.
      If lpSetPermissions AndAlso lpDocument.Permissions IsNot Nothing AndAlso lpDocument.Permissions.Count > 0 Then
        lobjP8Document.Permissions = CreateP8PermissionList(lpDocument.Permissions)
      End If
      ' </Modified by: Ernie Bahr at 9/13/2012-10:27:31 on machine: ERNIEBAHR-THINK>


      lobjP8Document.Save(RefreshMode.REFRESH)

      If HasElevatedPrivileges Then
        lobjP8Document.Properties.RemoveFromCache(PropertyNames.LAST_MODIFIER)
        lobjP8Document.Properties.RemoveFromCache(PropertyNames.DATE_LAST_MODIFIED)
        lobjP8Document.LastModifier = lstrLastModifer
        lobjP8Document.DateLastModified = ldatDateLastModified
        lobjP8Document.DateCheckedIn = ldatDateCheckedIn
      End If

      ' Check in the document
      If lblnIsMajorVersion = True Then
        lobjP8Document.Checkin(AutoClassify.DO_NOT_AUTO_CLASSIFY, CheckinType.MAJOR_VERSION)

      Else
        lobjP8Document.Checkin(AutoClassify.DO_NOT_AUTO_CLASSIFY, CheckinType.MINOR_VERSION)
      End If

      lobjP8Document.Save(RefreshMode.REFRESH)

      AssignSecurityPolicy(lpDocument.FirstVersion, lobjP8Document)

      ' Go through all the versions of the document
      For lintVersionCounter As Integer = 1 To lpDocument.Versions.Count - 1
        AddVersion(lpDocument.Versions(lintVersionCounter), lpSetPermissions, lobjP8Document, lpVersionType, lobjVersionSeriesId)
      Next

      'Dispose of Streams
      'For Each lobjContentElement As IContentTransfer In lobjP8Document.ContentElements
      '  lobjContentElement.AccessContentStream.Dispose()
      'Next

      If lpDocument.Relationships IsNot Nothing Then
        For Each lobjRelationship As Relationship In lpDocument.Relationships
          AddRelatedDocument(lobjP8Document, lpSetPermissions, lobjRelationship)
        Next
      End If

      Return lobjP8Document.VersionSeries.CurrentVersion

    Catch OutOfMemEx As OutOfMemoryException
      ApplicationLogging.LogException(OutOfMemEx, MethodBase.GetCurrentMethod)
      Throw New ContentTooLargeException(lpDocument.GetLargestFileSize.ToString)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      'lpErrorMessage = "Unable to Add Document" & vbCrLf & Helper.FormatCallStack(ex)
      'lpErrorMessage = String.Format("Unable to Add Document: Exception of type '{0}' thrown in {1} ({2})", _
      '                               ExceptionTracker.LastException.GetType.Name, _
      '                               ExceptionTracker.LastExceptionLocation, _
      '                               ExceptionTracker.LastException.Message)
      lpErrorMessage = String.Format("Unable to add document: {0}", ex.Message)
      Return Nothing
    End Try
  End Function

  Private Function AddRelatedDocument(ByVal lpParentP8Document As IDocument,
                                      ByVal lpSetPermissions As Boolean,
                                      ByVal lpRelationship As Relationship) As IDocument
    Try

      Dim lstrErrorMessage As String = String.Empty
      Dim lobjRelatedP8Document As IDocument = Nothing
      Dim lobjComponentRelationship As IComponentRelationship = Nothing

      ' Make sure we have a related document
      If lpRelationship.RelatedDocument Is Nothing Then
        Throw New DocumentReferenceNotSetException("The related document reference is not set.")
      End If

      ' Flag the parent document as a compound document
      lpParentP8Document.CompoundDocumentState = CompoundDocumentState.COMPOUND_DOCUMENT
      lpParentP8Document.Save(RefreshMode.REFRESH)

      ' Add the related document
      lobjRelatedP8Document = AddP8Document(lpRelationship.RelatedDocument, lpSetPermissions, lstrErrorMessage,
                                            VersionTypeEnum.Unspecified)

      If ((lobjRelatedP8Document Is Nothing) AndAlso
          (Not String.IsNullOrEmpty(lstrErrorMessage)) AndAlso
          (lstrErrorMessage.ToLower.Contains("class")) AndAlso
          (lstrErrorMessage.ToLower.EndsWith("not found."))) Then
        Throw New DocumentClassDoesNotExistException(lpRelationship.RelatedDocument.DocumentClass,
                                                     String.Format("Unable to add related document: {0}",
                                                                   lstrErrorMessage))
      End If

      ' Create the relationship
      lobjComponentRelationship = Factory.ComponentRelationship.CreateInstance(Me.ObjectStore, Nothing)

      With lobjComponentRelationship
        ' Set the parent document
        .ParentComponent = lpParentP8Document
        ' Set the child document
        .ChildComponent = lobjRelatedP8Document

        ' If the order has been specified, set it also
        If lpRelationship.Order.HasValue Then
          .ComponentSortOrder = lpRelationship.Order
        End If

        ' Set the relationship type
        .ComponentRelationshipType = lpRelationship.Persistance

        ' Depending on the relationship type, set the other information as well.
        Select Case lpRelationship.Persistance
          Case RelationshipPersistance.Dynamic
            .VersionBindType = lpRelationship.VersionBindType
          Case RelationshipPersistance.DynamicLabel
            .VersionBindType = lpRelationship.VersionBindType
            .LabelBindValue = lpRelationship.LabelBindValue
          Case RelationshipPersistance.URI
            .URIValue = lpRelationship.RelatedURI
          Case Else
            ' Do nothing
        End Select

        lobjComponentRelationship.ComponentCascadeDelete = lpRelationship.CascadeDeleteAction
        lobjComponentRelationship.ComponentPreventDelete = lpRelationship.PreventDeleteAction

        lobjComponentRelationship.Save(RefreshMode.REFRESH)

      End With

      Return lobjRelatedP8Document

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Sub AssignSecurityPolicy(ByVal lpVersion As DCore.Version, ByVal lpItem As IDocument)
    Try
      If lpVersion.Properties.Contains(PropertyNames.SECURITY_POLICY) Then
        Dim lobjSecurityPolicyProperty As DCore.IProperty = Nothing
        Dim lobjSecurityPolicy As ISecurityPolicy = Nothing

        If lpVersion.Properties.Contains(PropertyNames.SECURITY_POLICY) Then
          lobjSecurityPolicyProperty = lpVersion.Properties(PropertyNames.SECURITY_POLICY)

          Select Case lobjSecurityPolicyProperty.Type
            Case PropertyType.ecmString
              ' Get the name of the policy from the value
              If lobjSecurityPolicyProperty.HasValue Then
                lobjSecurityPolicy = GetSecurityPolicy(lobjSecurityPolicyProperty.Value)
              Else
                ApplicationLogging.WriteLogEntry(
                  String.Format("Unable to set security policy for document '{0}', the incoming value is not set.",
                                lpItem.Id.ToString), TraceEventType.Warning, 62983)
              End If

            Case PropertyType.ecmObject
              ' Get the guid from the value
              If lobjSecurityPolicyProperty.HasValue Then
                'lobjSecurityPolicy = GetObject("SecurityPolicy", lobjSecurityPolicyProperty.Value)
                lobjSecurityPolicy = GetSecurityPolicy(lobjSecurityPolicyProperty.Value)
              Else
                ApplicationLogging.WriteLogEntry(
                  String.Format("Unable to set security policy for document '{0}', the incoming value is not set.",
                                lpItem.Id.ToString), TraceEventType.Warning, 62983)
              End If

            Case Else
              Throw New InvalidPropertyTypeException(lobjSecurityPolicyProperty)
          End Select

          If lobjSecurityPolicy IsNot Nothing Then
            AssignSecurityPolicy(lobjSecurityPolicy, lpItem)
          End If
        End If
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Private Sub AssignSecurityPolicy(ByVal lpPolicyName As String, ByVal lpItem As IDocument)
    Try
      Dim lobjPolicy As ISecurityPolicy = GetSecurityPolicy(lpPolicyName)
      AssignSecurityPolicy(lobjPolicy, lpItem)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Private Sub AssignSecurityPolicy(ByVal lpPolicy As ISecurityPolicy, ByVal lpItem As IDocument)
    Try
      lpItem.SecurityPolicy = lpPolicy
      lpItem.Save(RefreshMode.NO_REFRESH)
      ApplicationLogging.WriteLogEntry(String.Format("Assigned Security Policy '{0}' to item '{1}'.",
                                                     lpPolicy.Name, lpItem.Name),
                                       MethodBase.GetCurrentMethod,
                                       TraceEventType.Information, 61211)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Private Function GetAllSecurityPolicyIds() As List(Of String)
    Try
      Dim lobjResults As IRepositoryRowSet
      Dim lobjTargetId As Id = Nothing
      Dim lobjReturnList As New List(Of String)

      lobjResults = GetP8Objects("SELECT [Id] FROM [SecurityPolicy] OPTIONS(TIMELIMIT 180)")

      If lobjResults IsNot Nothing Then
        For Each lobjRepositoryRow As IRepositoryRow In lobjResults
          ' Get the id of the first row
          lobjReturnList.Add(lobjRepositoryRow.Properties("Id").ToString)
        Next
      End If

      Return lobjReturnList

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Friend Function GetAuditEvents(lpP8Item As IContainable) As IAuditEvents
    Try
      Dim lobjAuditEvents As New AuditEvents
      Dim lobjCreateEvent As New AuditEvent
      Dim lobjLastModifiedEvent As New AuditEvent

      With lobjCreateEvent
        .Name = "Create"
        .Id = String.Format("{0}: {1}", lpP8Item.Id, .Name)
        .User = lpP8Item.Creator
        .EventDate = lpP8Item.DateCreated
        .Status = AuditEventStatusEnum.Success
      End With

      lobjAuditEvents.Add(lobjCreateEvent)

      If lpP8Item.LastModifier IsNot Nothing Then
        With lobjLastModifiedEvent
          .Name = "LastModified"
          .Id = String.Format("{0}: {1}", lpP8Item.Id, .Name)
          .User = lpP8Item.LastModifier
          .EventDate = lpP8Item.DateLastModified
          .Status = AuditEventStatusEnum.Success
        End With
        lobjAuditEvents.Add(lobjLastModifiedEvent)
      End If


      Return lobjAuditEvents

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function GetSecurityPolicy(ByVal lpPolicyReference As String) As ISecurityPolicy
    Try
      'Dim lobjSQLBuilder As New StringBuilder
      'Dim lobjResults As IRepositoryRowSet
      'Dim lobjTargetId As FileNet.Api.Util.Id = Nothing

      'lobjSQLBuilder.AppendFormat("SELECT [Id] FROM [SecurityPolicy] WHERE ([Name] = '{0}') OPTIONS(TIMELIMIT 180)", lpPolicyName)
      'lobjResults = GetP8Objects(lobjSQLBuilder.ToString)

      'If lobjResults Is Nothing Then
      '  Throw New Exceptions.ItemDoesNotExistException(lpPolicyName)
      'Else
      '  For Each lobjRepositoryRow As IRepositoryRow In lobjResults
      '    ' Get the id of the first row
      '    lobjTargetId = lobjRepositoryRow.Properties("Id")
      '    ' If there are more rows we don't care at this point, just exit the loop.
      '    Exit For
      '  Next

      '  Dim lobjObjectStore As IObjectStore = Factory.ObjectStore.GetInstance(Me.Domain, Me.ObjectStoreName)
      '  Factory.SecurityPolicy.FetchInstance(lobjObjectStore, lobjTargetId, Nothing)
      '  lobjObjectStore = Nothing
      '  Return lobjTargetId
      'End If

      Dim lobjFoundPolicy As SecurityPolicy = Nothing
      If SecurityPolicies.Contains(lpPolicyReference, lobjFoundPolicy) Then
        Return lobjFoundPolicy.Policy
      Else
        ApplicationLogging.WriteLogEntry(
          String.Format("No security policy found using reference '{0}'.", lpPolicyReference), TraceEventType.Warning,
          67404)
        Return Nothing
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function GetAllSecurityPolicies() As SecurityPolicies
    Try

      Dim lobjSecurityPolicies As New SecurityPolicies
      Dim lobjSecurityPolicyIds As List(Of String) = GetAllSecurityPolicyIds()
      Dim lobjSecurityPolicy As ISecurityPolicy

      For Each lstrPolicyId As String In lobjSecurityPolicyIds
        lobjSecurityPolicy = Me.ObjectStore.FetchObject("SecurityPolicy", lstrPolicyId, Nothing)
        If lobjSecurityPolicy IsNot Nothing Then
          lobjSecurityPolicies.Add(lobjSecurityPolicy)
        End If
      Next

      Return lobjSecurityPolicies

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function GetP8Objects(ByVal lpP8Sql As String) As IRepositoryRowSet
    Try

      Dim lobjObjectStore As IObjectStore = Factory.ObjectStore.GetInstance(Me.Domain, Me.ObjectStoreName)
      Dim lobjSearchScope As New SearchScope(lobjObjectStore)
      Dim lobjSQLObject As New SearchSQL()
      lobjSQLObject.SetQueryString(lpP8Sql)

      Return lobjSearchScope.FetchRows(lobjSQLObject, Nothing, Nothing, True)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function AddVersion(ByVal lpCTSVersion As DCore.Version,
                              ByVal lpSetPermissions As Boolean,
                              ByRef lpP8Document As IDocument) As IDocument
    Try
      Return AddVersion(lpCTSVersion, lpSetPermissions, lpP8Document, VersionTypeEnum.Unspecified, Nothing)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function AddVersion(ByVal lpCTSVersion As DCore.Version,
                              ByVal lpSetPermissions As Boolean,
                              ByRef lpP8Document As IDocument,
                              ByVal lpVersionType As VersionTypeEnum,
                              ByVal lpVersionSeriesId As Id) As IDocument

    ' BMK AddVersion
    Try

      Dim lobjP8Document As IDocument = Nothing
      Dim lobjId As Id = Id.CreateId
      Dim lstrPreviousVersionLastModifier As String = lpP8Document.LastModifier
      Dim ldatPreviousVersionLastModified As Date = lpP8Document.DateLastModified
      Dim lstrLastModifier As String = lpP8Document.LastModifier
      Dim ldatLastModified As Date = lpP8Document.DateLastModified

      '+ If supplied with the CTS version get the preferred values for LastModifier and LastModifiedDate

      ' Get LastModifier
      If _
        lpCTSVersion.Properties.PropertyExists(PropertyNames.LAST_MODIFIER) AndAlso
        lpCTSVersion.Properties(PropertyNames.LAST_MODIFIER).HasValue Then
        lstrLastModifier = lpCTSVersion.Properties(PropertyNames.LAST_MODIFIER).Value
      End If

      ' Get LastModifiedDate
      If _
        lpCTSVersion.Properties.PropertyExists(PropertyNames.DATE_LAST_MODIFIED) AndAlso
        lpCTSVersion.Properties(PropertyNames.DATE_LAST_MODIFIED).HasValue Then
        ldatLastModified = lpCTSVersion.Properties(PropertyNames.DATE_LAST_MODIFIED).Value
      End If

      ' Get the candidate document class
      Dim lobjCandidateClass As DocumentClass = DocumentClass(lpCTSVersion.Document.DocumentClass)

      ' See if we had an action property to set this as a major version
      ' TODO: Alternatively base the decision on the previous P8 MajorVersionNumber / MinorVersionNumber if available
      Dim lblnIsMajorVersion As Boolean

      ' <Added by: Ernie at: 1/10/2013-9:15:07 PM on machine: ERNIE-THINK>
      ' If the version type was specified then set it
      Select Case lpVersionType
        Case VersionTypeEnum.Major
          lblnIsMajorVersion = True
        Case VersionTypeEnum.Minor
          lblnIsMajorVersion = False
        Case VersionTypeEnum.Unspecified
          lblnIsMajorVersion = IsMajorVersion(lpCTSVersion, True)
      End Select
      ' </Added by: Ernie at: 1/10/2013-9:15:07 PM on machine: ERNIE-THINK>

      ' Do we have a version guid to preserve?
      'Dim lobjVersionIdProperty As ECMProperty = Nothing
      'Dim lobjCandidateId As FileNet.Api.Util.Id = Nothing
      'If lpCTSVersion.Properties.PropertyExists("Id", False, lobjVersionIdProperty) Then
      '  Dim lobjCandidateIdGuid As Guid

      '  Dim lstrVersionId As String = Nothing
      '  If lobjVersionIdProperty.Value IsNot Nothing Then
      '    If TypeOf lobjVersionIdProperty.Value Is Guid Then
      '      lstrVersionId = DirectCast(lobjVersionIdProperty.Value, Guid).ToString
      '    ElseIf TypeOf lobjVersionIdProperty.Value Is String Then
      '      lstrVersionId = lobjVersionIdProperty.Value
      '    End If
      '  End If
      '  If Not String.IsNullOrEmpty(lstrVersionId) AndAlso PreserveIdOnImport AndAlso Helper.IsGuid(lstrVersionId, lobjCandidateIdGuid) Then
      '    ' Check to see if the guid has already been used in this object store
      '    If DocumentExists(lstrVersionId) = False Then
      '      lobjCandidateId = New FileNet.Api.Util.Id(lobjCandidateIdGuid)
      '    Else
      '      ApplicationLogging.WriteLogEntry( _
      '        String.Format("Unable to preserve id for incoming version {0}, the identifier is already in use in object store '{1}'.", _
      '                      lstrVersionId, ObjectStoreName), Reflection.MethodBase.GetCurrentMethod, TraceEventType.Warning, 63402)
      '    End If
      '  End If
      'End If

      Dim lobjCandidateId As Id = GetVersionId(lpCTSVersion)

      ' Create the new version from scratch instead of relying on the product of the checkout operation.
      If lobjCandidateId IsNot Nothing Then

        Try
          If lpVersionSeriesId IsNot Nothing Then
            lobjP8Document = Factory.Document.CreateInstance(mobjObjectStore, lpCTSVersion.Document.DocumentClass,
                                                             lobjCandidateId, lpVersionSeriesId, ReservationType.OBJECT_STORE_DEFAULT)
          Else
            lobjP8Document = Factory.Document.CreateInstance(mobjObjectStore, lpCTSVersion.Document.DocumentClass,
                                                   lobjCandidateId)
          End If

        Catch ex As Exception
          ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
          lobjP8Document = Factory.Document.CreateInstance(mobjObjectStore, lpCTSVersion.Document.DocumentClass)
        End Try

      Else
        lobjP8Document = Factory.Document.CreateInstance(mobjObjectStore, lpCTSVersion.Document.DocumentClass)
      End If

      ' Set SOME of the properties
      If lpCTSVersion.Properties.Contains(PropertyNames.DATE_CREATED) Then _
        SetPropertyValue(lobjP8Document, lobjCandidateClass, lpCTSVersion.Properties(PropertyNames.DATE_CREATED), True)

      If lpCTSVersion.Properties.Contains(PropertyNames.CREATOR) Then _
        SetPropertyValue(lobjP8Document, lobjCandidateClass, lpCTSVersion.Properties(PropertyNames.CREATOR), True)

      If HasElevatedPrivileges Then
        '!+ This next part looks useless, but we need to re-set LastModifier and 
        '!+ DateLastModified on the previous version, otherwise performing the 
        '!+ checkout updates the document being checked out.
        lpP8Document.Properties.RemoveFromCache(PropertyNames.LAST_MODIFIER)
        lpP8Document.Properties.RemoveFromCache(PropertyNames.DATE_LAST_MODIFIED)

        lpP8Document.LastModifier = lstrPreviousVersionLastModifier
        lpP8Document.DateLastModified = ldatPreviousVersionLastModified

      End If

      ' We pass the new document object as a collection of properties to be 
      ' set on the version we are about to create...
      If lobjCandidateId IsNot Nothing Then
        lpP8Document.Checkout(ReservationType.EXCLUSIVE, lobjCandidateId, Nothing, lobjP8Document.Properties)
      Else
        lpP8Document.Checkout(ReservationType.EXCLUSIVE, lobjId, Nothing, lobjP8Document.Properties)
      End If
      lpP8Document.Save(RefreshMode.REFRESH)

      Dim lobjReservation As IDocument = lpP8Document.Reservation

      '!+ This looks useless, but we need to re-set LastModifier and DateLastModified to 
      '!+ preserve them past the checkout.  Otherwise performing the checkout updates 
      '!+ the document being checked out.
      lobjReservation.Properties.RemoveFromCache(PropertyNames.LAST_MODIFIER)
      lobjReservation.Properties.RemoveFromCache(PropertyNames.DATE_LAST_MODIFIED)

      lobjReservation.ContentElements = CreateContentElementList(lpCTSVersion.Contents)

      ' Set the properties
      For Each lobjCDFVersionProperty As ECMProperty In lpCTSVersion.Properties

        If lobjCDFVersionProperty.HasValue Then
          SetPropertyValue(lobjReservation, lobjCandidateClass, lobjCDFVersionProperty, False)
        End If

      Next

      If HasElevatedPrivileges Then
        '!+ Set based on cached values
        lobjReservation.LastModifier = lstrLastModifier
        lobjReservation.DateLastModified = ldatLastModified
      End If

      If lblnIsMajorVersion = True Then
        lobjReservation.Checkin(AutoClassify.DO_NOT_AUTO_CLASSIFY, CheckinType.MAJOR_VERSION)

      Else
        lobjReservation.Checkin(AutoClassify.DO_NOT_AUTO_CLASSIFY, CheckinType.MINOR_VERSION)
      End If

      ' <Modified by: Ernie Bahr at 9/13/2012-10:39:19 on machine: ERNIEBAHR-THINK>
      ' Updated to only set the permissions if the SetPermissions parameter is set to true.
      If lpSetPermissions AndAlso lpCTSVersion.Permissions IsNot Nothing AndAlso lpCTSVersion.Permissions.Count > 0 Then
        lobjReservation.Permissions = CreateP8PermissionList(lpCTSVersion.Permissions)
      End If
      ' </Modified by: Ernie Bahr at 9/13/2012-10:39:19 on machine: ERNIEBAHR-THINK>
      ' Set the permissions if we have them

      If lpCTSVersion.Relationships IsNot Nothing Then
        For Each lobjRelationship As Relationship In lpCTSVersion.Relationships
          AddRelatedDocument(lobjReservation, lpSetPermissions, lobjRelationship)
        Next
      End If

      lobjReservation.Save(RefreshMode.REFRESH)

      lpP8Document = lobjReservation

      Return lobjReservation

      ' TODO: Verify whether or not the method below for setting the properties is using preallocated properties 
      ' of the proper type vs creating default single-valued string properties.  
      ' We are seeing cases where properties are defined in the cdf as integers but are getting passed to P8 as strings.
      ' This is causing P8 to fail the import as shown in the error message below...
      '
      ' Exception of type EngineRuntimeException occured in CENetProvider::AddVersionToP8Document: 
      ' The method invoked is inappropriate for the datatype of the property. 
      ' The property MaterialSafetyDataSheetNumber expects the datatype to be java.lang.Integer, 
      ' but the input value is java.lang.String.	2011-04-27 16:25:20Z
    Catch RuntimeEx As EngineRuntimeException
      ApplicationLogging.LogException(RuntimeEx, MethodBase.GetCurrentMethod)
      If RuntimeEx.Message.Contains("Checkout must be performed when versioning is enabled") Then
        Throw New VersioningNotEnabledException(String.Format("Versioning is not enabled for class '{0}'.", lpCTSVersion.Document.DocumentClass), RuntimeEx)
      Else
        ' Re-throw the exception to the caller
        Throw
      End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function GetVersionId(ByVal lpCTSVersion As DCore.Version) As Id

    Try

      Dim lobjVersionIdProperty As ECMProperty = Nothing
      Dim lobjCandidateId As Id = Nothing

      If lpCTSVersion.Properties.PropertyExists("VersionId", False, lobjVersionIdProperty) Then

        Dim lobjCandidateIdGuid As Guid

        Dim lstrVersionId As String = Nothing

        If lobjVersionIdProperty.Value IsNot Nothing Then

          If TypeOf lobjVersionIdProperty.Value Is Guid Then
            lstrVersionId = DirectCast(lobjVersionIdProperty.Value, Guid).ToString

          ElseIf TypeOf lobjVersionIdProperty.Value Is String Then
            lstrVersionId = lobjVersionIdProperty.Value
          End If

        End If

        If _
          Not String.IsNullOrEmpty(lstrVersionId) AndAlso PreserveIdOnImport AndAlso
          Helper.IsGuid(lstrVersionId, lobjCandidateIdGuid) Then

          ' Check to see if the guid has already been used in this object store
          If DocumentExists(lstrVersionId) = False Then
            lobjCandidateId = New Id(lobjCandidateIdGuid)

          Else
            ApplicationLogging.WriteLogEntry(
              String.Format(
                "Unable to preserve id for incoming version {0}, the identifier is already in use in object store '{1}'.",
                lstrVersionId, ObjectStoreName), MethodBase.GetCurrentMethod, TraceEventType.Warning, 63402)
          End If

        End If

      End If

      Return lobjCandidateId

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function GetVersionSeriesId(ByVal lpCTSVersion As DCore.Version) As Id

    Try

      Dim lobjVersionIdProperty As ECMProperty = Nothing
      Dim lobjCandidateId As Id = Nothing

      If lpCTSVersion.Properties.PropertyExists("VersionSeriesId", False, lobjVersionIdProperty) Then

        Dim lobjCandidateIdGuid As Guid

        Dim lstrVersionId As String = Nothing

        If lobjVersionIdProperty.Value IsNot Nothing Then

          If TypeOf lobjVersionIdProperty.Value Is Guid Then
            lstrVersionId = DirectCast(lobjVersionIdProperty.Value, Guid).ToString

          ElseIf TypeOf lobjVersionIdProperty.Value Is String Then
            lstrVersionId = lobjVersionIdProperty.Value
          End If

        End If

        'If _
        '  Not String.IsNullOrEmpty(lstrVersionId) AndAlso PreserveIdOnImport AndAlso
        '  Helper.IsGuid(lstrVersionId, lobjCandidateIdGuid) Then

        '  ' Check to see if the guid has already been used in this object store
        '  If DocumentExists(lstrVersionId) = False Then
        '    lobjCandidateId = New Id(lobjCandidateIdGuid)

        '  Else
        '    ApplicationLogging.WriteLogEntry(
        '      String.Format(
        '        "Unable to preserve id for incoming version {0}, the identifier is already in use in object store '{1}'.",
        '        lstrVersionId, ObjectStoreName), MethodBase.GetCurrentMethod, TraceEventType.Warning, 63402)
        '  End If

        'End If

        If Not String.IsNullOrEmpty(lstrVersionId) AndAlso PreserveIdOnImport AndAlso
            Helper.IsGuid(lstrVersionId, lobjCandidateIdGuid) Then
          lobjCandidateId = New Id(lobjCandidateIdGuid)
        End If
      End If

      Return lobjCandidateId

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function GetVersionId(ByVal lpFolder As Folder) As Id

    Try

      Dim lobjVersionIdProperty As ECMProperty = Nothing
      Dim lobjCandidateId As Id = Nothing

      If lpFolder.Properties.PropertyExists("Id", False, lobjVersionIdProperty) Then

        Dim lobjCandidateIdGuid As Guid

        Dim lstrVersionId As String = Nothing

        If lobjVersionIdProperty.Value IsNot Nothing Then

          If TypeOf lobjVersionIdProperty.Value Is Guid Then
            lstrVersionId = DirectCast(lobjVersionIdProperty.Value, Guid).ToString

          ElseIf TypeOf lobjVersionIdProperty.Value Is String Then
            lstrVersionId = lobjVersionIdProperty.Value
          End If

        End If

        If _
          Not String.IsNullOrEmpty(lstrVersionId) AndAlso PreserveIdOnImport AndAlso
          Helper.IsGuid(lstrVersionId, lobjCandidateIdGuid) Then

          ' Check to see if the guid has already been used in this object store
          If DocumentExists(lstrVersionId) = False Then
            lobjCandidateId = New Id(lobjCandidateIdGuid)

          Else
            ApplicationLogging.WriteLogEntry(
              String.Format(
                "Unable to preserve id for incoming version {0}, the identifier is already in use in object store '{1}'.",
                lstrVersionId, ObjectStoreName), MethodBase.GetCurrentMethod, TraceEventType.Warning, 63402)
          End If

        End If

      End If

      Return lobjCandidateId

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function


  ''' <summary>
  '''   Sets the property value in the P8 document using the supplied CTS property.
  ''' </summary>
  ''' <param name="lpP8Object">The P8 object to update.</param>
  ''' <param name="lpCandidateClass">The CTS class to validate against.</param>
  ''' <param name="lpCtsProperty">The CTS property source.</param>
  ''' <param name="lpIsNewObject">Specifies whether or not the object is in a new state.</param>
  ''' <remarks></remarks>
  Private Sub SetPropertyValue(ByVal lpP8Object As IEngineObject,
                               ByVal lpCandidateClass As RepositoryObjectClass,
                               ByVal lpCtsProperty As ECMProperty,
                               ByVal lpIsNewObject As Boolean)

    ' We are trying to locate a null reference exception
    Dim lintCodeMarker As Integer
    Dim lobjTargetProperty As ClassificationProperty
    Dim lobjPendingCreateAction As Create = Nothing
    Dim lobjP8Object As Object = lpP8Object

    Try

      If lpIsNewObject Then

        Dim lobjPendingActions As PendingAction() = lobjP8Object.GetPendingActions

        For Each lobjPendingAction As Object In lobjPendingActions

          If TypeOf lobjPendingAction Is Create Then
            lobjPendingCreateAction = lobjPendingAction
            Exit For
          End If

        Next

      End If

      lintCodeMarker += 1 ' 1

      ' Check the incoming parameters
      ArgumentNullException.ThrowIfNull(lpP8Object)

      lintCodeMarker += 1 ' 2

      ArgumentNullException.ThrowIfNull(lpCandidateClass)

      lintCodeMarker += 1 ' 3

      ArgumentNullException.ThrowIfNull(lpCtsProperty)

      lintCodeMarker += 1 ' 4

      If lpCtsProperty.Value Is Nothing OrElse lpCtsProperty.HasValue = False Then
        ' There is no value to set, exit the sub
        Exit Sub
      End If

      lintCodeMarker += 1 ' 5
      ' Let's look at how the property is defined in P8
      ' Make sure the specified class contains this property
      lobjTargetProperty = lpCandidateClass.Properties.ItemByName(lpCtsProperty.Name) '.Item(lpCtsProperty.Name)

      lintCodeMarker += 1 ' 6

      If lobjTargetProperty Is Nothing Then
        ApplicationLogging.WriteLogEntry(
          String.Format("Unable to set property '{0}'.  The property could not be found for the document class '{1}'.",
                        lpCtsProperty.Name, lpCandidateClass.Name), TraceEventType.Warning, 63825)
        Exit Sub
      End If

      lintCodeMarker += 1 ' 7

      If lobjTargetProperty.Settability = ClassificationProperty.SettabilityEnum.READ_ONLY Then
        ApplicationLogging.WriteLogEntry(
          String.Format("Unable to set the value for read only property '{0}'.", lobjTargetProperty.Name),
          MethodBase.GetCurrentMethod, TraceEventType.Warning, 62396)
        Exit Sub
      End If

      lintCodeMarker += 1 ' 8

      If lpIsNewObject = False Then

        If lobjTargetProperty.Settability = ClassificationProperty.SettabilityEnum.SETTABLE_ONLY_ON_CREATE Then

          If LogInvalidPropertyRemovals Then
            ApplicationLogging.WriteLogEntry(
              String.Format("Unable to set the value for settable only on create property '{0}'.",
                            lobjTargetProperty.Name), MethodBase.GetCurrentMethod, TraceEventType.Warning, 62397)
          End If

          Exit Sub
        End If

      Else
        'If lobjTargetProperty.Settability <> ClassificationProperty.SettabilityEnum.SETTABLE_ONLY_ON_CREATE _
        '  AndAlso lobjTargetProperty.IsRequired = False Then
        '  ' Skip this property unless it is 'DateLastModified' since that one is really to be set only on create
        '  If String.Equals(lobjTargetProperty.SystemName, PropertyNames.DATE_LAST_MODIFIED) = False Then
        '    Exit Sub
        '  End If
        'End If
      End If

      lintCodeMarker += 1 ' 9

      ' Check for predefined properties 
      Select Case lobjTargetProperty.SystemName

        Case PropertyNames.ID
          ' Skip this property
          Exit Sub

        Case PropertyNames.COMPOUND_DOCUMENT_STATE
          ' Skip this property
          Exit Sub
          'Case Constants.PropertyNames.CONTAINERS

        Case PropertyNames.CONTENT_ELEMENTS
          ' Skip this property
          Exit Sub

        Case PropertyNames.CREATOR

          If HasElevatedPrivileges Then
            lobjP8Object.Creator = lpCtsProperty.Value
            Exit Sub
          End If

        Case PropertyNames.DATE_CHECKED_IN

          If HasElevatedPrivileges AndAlso lpIsNewObject = False Then
            lobjP8Object.DateCheckedIn = lpCtsProperty.Value
            Exit Sub
          End If

        Case PropertyNames.DATE_CREATED

          If HasElevatedPrivileges Then

            If IsDate(lpCtsProperty.Value) Then
              lobjP8Object.DateCreated = Date.Parse(lpCtsProperty.Value)
            End If

            Exit Sub
          End If

        Case PropertyNames.DATE_LAST_MODIFIED

          If HasElevatedPrivileges Then
            ''If IsPendingCreate(lpP8Document) Then
            ''  If IsDate(lpCtsProperty.Value) Then
            ''    lpP8Document.DateLastModified = Date.Parse(lpCtsProperty.Value)
            ''  End If
            ''End If
            SetSingleValuedProperty(lpP8Object, lobjTargetProperty, lpCtsProperty, lpIsNewObject,
                                    lobjPendingCreateAction)
            'If lpP8Document.VersionStatus <> VersionStatus.RESERVATION Then
            ' lpP8Document.DateLastModified = lpCtsProperty.Value
            'End If
          End If

          Exit Sub

        Case PropertyNames.DOCUMENT_LIFECYCLE_POLICY
          ' TODO: Find a safe way to set this value
          'lpP8Document.DocumentLifecyclePolicy = lpCtsProperty.Value
          Exit Sub

        Case PropertyNames.LAST_MODIFIER

          If HasElevatedPrivileges Then
            ''If lpP8Document.VersionStatus <> VersionStatus.RESERVATION Then
            ''  lpP8Document.LastModifier = lpCtsProperty.Value
            ''End If
            SetSingleValuedProperty(lpP8Object, lobjTargetProperty, lpCtsProperty, lpIsNewObject,
                                    lobjPendingCreateAction)
          End If

          Exit Sub

        Case PropertyNames.MAJOR_VERSION_NUMBER, PropertyNames.MINOR_VERSION_NUMBER, "Major Version Number",
          "Minor Version Number"
          ' We do not want to set these
          Exit Sub

          'Case Constants.PropertyNames.MIME_TYPE
          '  lpP8Document.MimeType = lpCtsProperty.Value
          '  Exit Sub

          'Case Constants.PropertyNames.OWNER
          '  lpP8Document.Owner = lpCtsProperty.Value
          '  Exit Sub
        Case PropertyNames.OWNER_DOCUMENT
          ' TODO: See if there is a safe way to set this property
          Exit Sub

        Case PropertyNames.PUBLICATION_INFO
          ' Skip this property
          Exit Sub

        Case PropertyNames.PUBLISHING_SUBSIDIARY_FOLDER
          ' Skip this property
          Exit Sub

        Case PropertyNames.SECURITY_POLICY
          ' Skip this property
          Exit Sub

        Case PropertyNames.SECURITY_FOLDER, PropertyNames.SECURITY_PARENT
          ' Skip this property
          Exit Sub

        Case PropertyNames.SOURCE_DOCUMENT
          ' Skip this property for now
          Exit Sub

        Case PropertyNames.STORAGE_AREA
          ' Skip this property for now
          Exit Sub

        Case PropertyNames.STORAGE_POLICY
          ' Skip this property for now
          Exit Sub

        Case PropertyNames.PARENT
          ' Skip this property
          Exit Sub

        Case Else

          If _
            lpIsNewObject AndAlso lobjPendingCreateAction IsNot Nothing AndAlso
            lobjTargetProperty.Cardinality = DCore.Cardinality.ecmSingleValued Then
            SetSingleValuedProperty(lpP8Object, lobjTargetProperty, lpCtsProperty, lpIsNewObject,
                                    lobjPendingCreateAction)
            Exit Sub
          End If

      End Select

      Dim lobjTargetP8Property As Object = Nothing

      lintCodeMarker += 1 ' 10

      If lpCtsProperty.Cardinality = DCore.Cardinality.ecmSingleValued Then

        If lpIsNewObject = False Then
          SetSingleValuedProperty(lpP8Object, lobjTargetProperty, lpCtsProperty, False)
        End If

      Else
        lintCodeMarker += 1 ' 11
        ' Set the multi-valued property value
        'lpP8Property.SetObjectValue(CreateMultiValuedPropertyValue(lpCtsProperty))
        lobjTargetProperty = ContentProperties.ItemByName(lpCtsProperty.Name)

        lintCodeMarker += 1 ' 12

        If lobjTargetProperty Is Nothing Then
          Throw _
            New PropertyDoesNotExistException(
              String.Format("No property definition found for the property '{0}' in object store '{1}'.",
                            lpCtsProperty.Name, ObjectStoreName), lpCtsProperty.Name)

        Else

          lintCodeMarker += 1 ' 13

          If _
            lobjTargetProperty.Settability = ClassificationProperty.SettabilityEnum.READ_WRITE AndAlso
            lpCtsProperty.Name.Equals(PropertyNames.PERMISSIONS) = False Then
            lintCodeMarker += 1 ' 14
            SetMultiValuedProperty(lpP8Object, lobjTargetProperty, lpCtsProperty, lpIsNewObject, lobjPendingCreateAction)
            'lobjTargetP8Property = lpP8Document.Properties(lobjTargetProperty.SystemName)
            'If lobjTargetP8Property IsNot Nothing Then
            '  lintCodeMarker += 1 ' 15
            '  Dim lobjPropertyValue As Object = CreateMultiValuedPropertyValue(lpCtsProperty)
            '  If lobjPropertyValue IsNot Nothing Then
            '    lintCodeMarker += 1 ' 16
            '    lpP8Document.Properties(lobjTargetProperty.SystemName) = lobjPropertyValue
            '  End If
            'End If
          End If

        End If

      End If

    Catch ex As Exception

      ' Try to get a little context information for the log to help 
      ' in troubleshooting which property we had the problem with.
      Dim lobjParamBuilder As New StringBuilder

      If lpCtsProperty IsNot Nothing Then
        lobjParamBuilder.AppendFormat("PropertyName: {0}", lpCtsProperty.Name)
      End If

      If lpCandidateClass IsNot Nothing Then
        lobjParamBuilder.AppendFormat(", DocumentClass: {0}", lpCandidateClass.Name)
      End If

      lobjParamBuilder.AppendFormat(", OnlySetSettableOnlyOnCreate: {0}", lpIsNewObject)
      lobjParamBuilder.AppendFormat(" - CodeMarker: {0}", lintCodeMarker)

      Dim _
        lobjPropertyException As _
          New PropertyException(String.Format("{0} - {1}", ex.Message, lobjParamBuilder.ToString), lpCtsProperty, ex)
      ApplicationLogging.LogException(lobjPropertyException, MethodBase.GetCurrentMethod)
      Throw lobjPropertyException
    End Try
  End Sub

  Private Sub SetSingleValuedProperty(ByRef lpP8Object As IEngineObject,
                                      ByRef lpTargetProperty As ClassificationProperty,
                                      ByRef lpCtsProperty As ECMProperty,
                                      ByVal lpIsNewObject As Boolean,
                                      Optional ByRef lpPendingCreateAction As Create = Nothing)
    ' BMK SetSingleValuedProperty
    ' We are still in a create state (i.e. the document has not yet been saved for the first time)
    ' We need to set the property in a different way
    ' This is necessary to set the required properties

    Try

      If lpCtsProperty.Cardinality = DCore.Cardinality.ecmMultiValued Then
        Throw New InvalidCardinalityException(DCore.Cardinality.ecmSingleValued, lpCtsProperty)
      End If

      Dim lstrInvalidPropertyTypeMessage As String =
            String.Format("Destination property '{0}' is type '{1}', the supplied property was of type '{2}'.",
                          lpTargetProperty.Name, lpTargetProperty.Type.ToString, lpCtsProperty.Type.ToString)

      Select Case lpCtsProperty.Type

        Case PropertyType.ecmString, PropertyType.ecmHtml, PropertyType.ecmUri, PropertyType.ecmXml

          If lpTargetProperty.Type = PropertyType.ecmString Then

            If lpIsNewObject Then
              lpPendingCreateAction.PutValue(lpTargetProperty.SystemName, lpCtsProperty.Value.ToString)
            End If

            lpP8Object.Properties(lpTargetProperty.SystemName) = lpCtsProperty.Value.ToString

          Else
            'Throw New Exceptions.InvalidPropertyTypeException(lstrInvalidPropertyTypeMessage, lpCtsProperty)
          End If

        Case PropertyType.ecmBoolean

          If lpTargetProperty.Type = PropertyType.ecmBoolean Then

            Dim lblnValue As Boolean
            If TypeOf lpCtsProperty.Value Is String Then
              If Boolean.TryParse(lpCtsProperty.Value, lblnValue) = False Then
                ApplicationLogging.WriteLogEntry(
                  String.Format("Failed to convert string '{0}' to boolean for proprty '{1}'.",
                                lpCtsProperty.Value, lpCtsProperty.Name),
                  MethodBase.GetCurrentMethod(),
                  TraceEventType.Warning, 62384)

              End If
            Else
              lblnValue = lpCtsProperty.Value
            End If

            If lpIsNewObject Then
              lpPendingCreateAction.PutValue(lpTargetProperty.SystemName, lblnValue)
            End If

            lpP8Object.Properties(lpTargetProperty.SystemName) = lblnValue

          ElseIf lpTargetProperty.Type = PropertyType.ecmString Then

            If lpIsNewObject Then
              lpPendingCreateAction.PutValue(lpTargetProperty.SystemName, lpCtsProperty.Value.ToString)
            End If

            lpP8Object.Properties(lpTargetProperty.SystemName) = lpCtsProperty.Value.ToString
          End If

        Case PropertyType.ecmDate

          If IsDate(lpCtsProperty.Value) Then

            If lpTargetProperty.Type = PropertyType.ecmDate Then

              If lpIsNewObject Then
                lpPendingCreateAction.PutValue(lpTargetProperty.SystemName, lpCtsProperty.Value)
              End If

              lpP8Object.Properties(lpTargetProperty.SystemName) = lpCtsProperty.Value

            ElseIf lpTargetProperty.Type = PropertyType.ecmString Then

              If lpIsNewObject Then
                lpPendingCreateAction.PutValue(lpTargetProperty.SystemName, lpCtsProperty.Value.ToString)
              End If

              lpP8Object.Properties(lpTargetProperty.SystemName) = lpCtsProperty.Value.ToString

            Else
              Throw New InvalidPropertyTypeException(lstrInvalidPropertyTypeMessage, lpCtsProperty)
            End If

          ElseIf TypeOf lpCtsProperty.Value Is String Then

            If String.IsNullOrEmpty(lpCtsProperty.Value) Then
              Exit Select

            Else
              Throw _
                New ArgumentOutOfRangeException(lpCtsProperty.Name, lpCtsProperty.Value,
                                                String.Format(
                                                  "The value '{0}' is not a valid date, it can not be accepted.",
                                                  lpCtsProperty.Value))
            End If

          Else
            Throw _
              New ArgumentOutOfRangeException(lpCtsProperty.Name, lpCtsProperty.Value,
                                              String.Format(
                                                "The value '{0}' is not a valid date, it can not be accepted.",
                                                lpCtsProperty.Value))
          End If

        Case PropertyType.ecmLong

          If IsNumeric(lpCtsProperty.Value) Then

            If lpTargetProperty.Type = PropertyType.ecmLong Then

              If lpCtsProperty.Value > Integer.MaxValue Then
                Throw _
                  New ArgumentOutOfRangeException(lpCtsProperty.Name, lpCtsProperty.Value,
                                                  String.Format(
                                                    "The value '{0}' is too large for an integer field, it can not be accepted.",
                                                    lpCtsProperty.Value))

              Else

                If lpIsNewObject Then
                  lpPendingCreateAction.PutValue(lpTargetProperty.SystemName, Integer.Parse(lpCtsProperty.Value))
                End If

                lpP8Object.Properties(lpTargetProperty.SystemName) = Integer.Parse(lpCtsProperty.Value)
              End If

            ElseIf lpTargetProperty.Type = PropertyType.ecmDouble Then

              If lpIsNewObject Then
                lpPendingCreateAction.PutValue(lpTargetProperty.SystemName, lpCtsProperty.Value)
              End If

              lpP8Object.Properties(lpTargetProperty.SystemName) = lpCtsProperty.Value

            Else
              Throw New InvalidPropertyTypeException(lstrInvalidPropertyTypeMessage, lpCtsProperty)
            End If

          Else
            Throw _
              New ArgumentOutOfRangeException(lpCtsProperty.Name, lpCtsProperty.Value,
                                              String.Format("The value '{0}' is not numeric, it can not be accepted.",
                                                            lpCtsProperty.Value))
          End If

        Case PropertyType.ecmDouble

          If IsNumeric(lpCtsProperty.Value) Then

            If lpTargetProperty.Type = PropertyType.ecmDouble Then

              If lpCtsProperty.Value > Double.MaxValue Then
                Throw _
                  New ArgumentOutOfRangeException(lpCtsProperty.Name, lpCtsProperty.Value,
                                                  String.Format(
                                                    "The value '{0}' is too large for an integer field, it can not be accepted.",
                                                    lpCtsProperty.Value))

              Else

                If lpIsNewObject Then
                  lpPendingCreateAction.PutValue(lpTargetProperty.SystemName, Double.Parse(lpCtsProperty.Value))
                End If

                lpP8Object.Properties(lpTargetProperty.SystemName) = Double.Parse(lpCtsProperty.Value)
              End If

            Else
              Throw New InvalidPropertyTypeException(lstrInvalidPropertyTypeMessage, lpCtsProperty)
            End If

          Else
            Throw _
              New ArgumentOutOfRangeException(lpCtsProperty.Name, lpCtsProperty.Value,
                                              String.Format("The value '{0}' is not numeric, it can not be accepted.",
                                                            lpCtsProperty.Value))
          End If

        Case PropertyType.ecmGuid

          If lpTargetProperty.Type = PropertyType.ecmGuid Then

            Dim lobjId As Guid = Nothing

            If lpCtsProperty.HasValue AndAlso lpCtsProperty.Value IsNot Nothing Then

              If Helper.IsGuid(lpCtsProperty.Value.ToString, lobjId) Then

                Dim lobjFileNetId As New Id(lobjId)

                If lpIsNewObject Then
                  lpPendingCreateAction.PutValue(lpTargetProperty.SystemName, lobjFileNetId)
                End If

                lpP8Object.Properties(lpTargetProperty.SystemName) = lobjFileNetId
              End If

            End If

          Else
            Throw New InvalidPropertyTypeException(lstrInvalidPropertyTypeMessage, lpCtsProperty)
          End If

        Case PropertyType.ecmObject

          ' BMK SetSingleValuedProperty:Object
          If lpTargetProperty.Type = PropertyType.ecmObject Then

            Dim lobjPropertyTemplate As IPropertyTemplate = GetPropertyTemplate(lpTargetProperty.SystemName)

            If lobjPropertyTemplate Is Nothing Then
              Throw New PropertyDoesNotExistException(String.Format("Unable to set property value, {0} does not exist.",
                                                                    lpTargetProperty.SystemName),
                                                      lpTargetProperty.SystemName)
            End If

            Dim lstrValue As String = lpCtsProperty.Value.ToString
            Dim lintIndex As Integer = lstrValue.IndexOf(":"c)
            If lintIndex > 0 Then
              lstrValue = lstrValue.Substring(0, lintIndex)
            End If

            Dim lstrClassName As String = lpCtsProperty.Value.ToString
            Dim lintClassNameIndex = lstrClassName.IndexOf(" - ")
            If lintClassNameIndex > 0 Then
              lstrClassName = lstrClassName.Substring(lintIndex + 1, lintClassNameIndex - lintIndex)
            End If

            Dim lobjObjectId As New Id(lstrValue)

            'Dim lobjObjectValue As IIndependentObject = GetObject(lobjPropertyTemplate.ClassDescription.Id.ToString, lobjObjectId)
            Dim lobjObjectValue As IIndependentObject = GetObject(lstrClassName, lobjObjectId)

            If lpIsNewObject Then
              lpPendingCreateAction.PutValue(lpTargetProperty.SystemName, lobjObjectValue)
            End If

            lpP8Object.Properties(lpTargetProperty.SystemName) = lobjObjectValue

          End If

          ''If lobjTargetProperty.Type = PropertyType.ecmObject Then
          ''  lpP8Document.Properties(lobjTargetProperty.SystemName) = lpCtsProperty.Value
          ''Else
          ''  Throw New Exceptions.InvalidPropertyTypeException(lstrInvalidPropertyTypeMessage, lpCtsProperty)
          ''End If
          'If TypeOf lpP8Property Is IPropertyId Then
          '  lpP8Property.SetObjectValue(lpCtsProperty.Value)
          'Else
          '  Throw New Exceptions.InvalidPropertyTypeException(lstrInvalidPropertyTypeMessage, lpCtsProperty)
          'End If
        Case Else

          Throw New InvalidPropertyTypeException(lstrInvalidPropertyTypeMessage, lpCtsProperty)

      End Select

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Private Sub SetMultiValuedProperty(ByRef lpP8Object As IEngineObject,
                                     ByRef lpTargetProperty As ClassificationProperty,
                                     ByRef lpCtsProperty As ECMProperty,
                                     ByVal lpIsNewObject As Boolean,
                                     Optional ByRef lpPendingCreateAction As Create = Nothing)

    Try

      If lpCtsProperty.Cardinality = DCore.Cardinality.ecmSingleValued Then
        Throw New InvalidCardinalityException(DCore.Cardinality.ecmMultiValued, lpCtsProperty)
      End If

      Dim lstrInvalidPropertyTypeMessage As String =
            String.Format("Destination property '{0}' is type '{1}', the supplied property was of type '{2}'.",
                          lpTargetProperty.Name, lpTargetProperty.Type.ToString, lpCtsProperty.Type.ToString)

      Dim lobjPropertyValue As Object = CreateMultiValuedPropertyValue(lpCtsProperty)

      ' Even though the implementation for the property types below are all the same, we are establishing 
      ' the separate pattern now in case we need to treat them differently after further testing.

      Select Case lpCtsProperty.Type

        Case PropertyType.ecmString, PropertyType.ecmHtml, PropertyType.ecmUri, PropertyType.ecmXml

          If lpIsNewObject Then
            lpPendingCreateAction.PutValue(lpTargetProperty.SystemName, lobjPropertyValue)
          End If

          lpP8Object.Properties(lpTargetProperty.SystemName) = lobjPropertyValue

        Case PropertyType.ecmBoolean

          If lpIsNewObject Then
            lpPendingCreateAction.PutValue(lpTargetProperty.SystemName, lobjPropertyValue)
          End If

          lpP8Object.Properties(lpTargetProperty.SystemName) = lobjPropertyValue

        Case PropertyType.ecmDate

          If lpIsNewObject Then
            lpPendingCreateAction.PutValue(lpTargetProperty.SystemName, lobjPropertyValue)
          End If

          lpP8Object.Properties(lpTargetProperty.SystemName) = lobjPropertyValue

        Case PropertyType.ecmLong

          If lpIsNewObject Then
            lpPendingCreateAction.PutValue(lpTargetProperty.SystemName, lobjPropertyValue)
          End If

          lpP8Object.Properties(lpTargetProperty.SystemName) = lobjPropertyValue

        Case PropertyType.ecmDouble

          If lpIsNewObject Then
            lpPendingCreateAction.PutValue(lpTargetProperty.SystemName, lobjPropertyValue)
          End If

          lpP8Object.Properties(lpTargetProperty.SystemName) = lobjPropertyValue

        Case PropertyType.ecmGuid

          If lpIsNewObject Then
            lpPendingCreateAction.PutValue(lpTargetProperty.SystemName, lobjPropertyValue)
          End If

          lpP8Object.Properties(lpTargetProperty.SystemName) = lobjPropertyValue

        Case PropertyType.ecmObject
          'Skip this type for now

        Case Else
          Throw New InvalidPropertyTypeException(lstrInvalidPropertyTypeMessage, lpCtsProperty)

      End Select

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  'Private Sub SetSingleValuedPropertyForExistingDocument(ByVal lpP8Document As IDocument, _
  '                             ByVal lpCandidateDocumentClass As DocumentClass, _
  '                             ByVal lpTargetProperty As ClassificationProperty, _
  '                             ByVal lpCtsProperty As ECMProperty)
  '  Try
  '    Dim lstrInvalidPropertyTypeMessage As String = _
  '      String.Format("Destination property '{0}' is type '{1}', the supplied property was of type '{2}'.", _
  '              lpTargetProperty.Name, lpTargetProperty.Type.ToString, lpCtsProperty.Type.ToString)

  '    Select Case lpCtsProperty.Type
  '      Case PropertyType.ecmString, PropertyType.ecmHtml, PropertyType.ecmUri, PropertyType.ecmXml
  '        If lpTargetProperty.Type = PropertyType.ecmString Then
  '          If Not String.IsNullOrEmpty(lpCtsProperty.Value) Then
  '            lpP8Document.Properties(lpTargetProperty.SystemName) = lpCtsProperty.Value.ToString
  '          Else : Exit Select
  '          End If
  '        Else
  '          Throw New Exceptions.InvalidPropertyTypeException(lstrInvalidPropertyTypeMessage, lpCtsProperty)
  '        End If
  '        'If TypeOf lpP8Property Is IPropertyString Then
  '        '  lpP8Property.SetObjectValue(lpCtsProperty.Value.ToString)
  '        'Else
  '        '  Throw New Exceptions.InvalidPropertyTypeException(lstrInvalidPropertyTypeMessage, lpCtsProperty)
  '        'End If
  '        'lpP8Document.Properties(lpTargetProperty.SystemName) = lpCtsProperty.Value
  '      Case PropertyType.ecmBoolean
  '        If lpTargetProperty.Type = PropertyType.ecmBoolean Then
  '          lpP8Document.Properties(lpTargetProperty.SystemName) = lpCtsProperty.Value
  '        ElseIf lpTargetProperty.Type = PropertyType.ecmString Then
  '          lpP8Document.Properties(lpTargetProperty.SystemName) = lpCtsProperty.Value.ToString
  '        End If
  '        'If TypeOf lpP8Property Is IPropertyBoolean Then
  '        '  lpP8Property.SetObjectValue(lpCtsProperty.Value)
  '        'ElseIf TypeOf lpP8Property Is IPropertyString Then
  '        '  lpP8Property.SetObjectValue(lpCtsProperty.Value.ToString)
  '        'End If
  '      Case PropertyType.ecmDate
  '        If IsDate(lpCtsProperty.Value) Then
  '          If lpTargetProperty.Type = PropertyType.ecmDate Then
  '            lpP8Document.Properties(lpTargetProperty.SystemName) = lpCtsProperty.Value
  '          ElseIf lpTargetProperty.Type = PropertyType.ecmString Then
  '            lpP8Document.Properties(lpTargetProperty.SystemName) = lpCtsProperty.Value.ToString
  '          Else
  '            Throw New Exceptions.InvalidPropertyTypeException(lstrInvalidPropertyTypeMessage, lpCtsProperty)
  '          End If
  '          'If TypeOf lpP8Property Is IPropertyDateTime Then
  '          '  lpP8Property.SetObjectValue(lpCtsProperty.Value)
  '          'ElseIf TypeOf lpP8Property Is IPropertyString Then
  '          '  lpP8Property.SetObjectValue(lpCtsProperty.Value.ToString)
  '          'Else
  '          '  Throw New Exceptions.InvalidPropertyTypeException(lstrInvalidPropertyTypeMessage, lpCtsProperty)
  '          'End If
  '        ElseIf TypeOf lpCtsProperty.Value Is String Then
  '          If String.IsNullOrEmpty(lpCtsProperty.Value) Then
  '            Exit Select
  '          Else
  '            Throw New ArgumentOutOfRangeException(lpCtsProperty.Name, lpCtsProperty.Value, String.Format("The value '{0}' is not a valid date, it can not be accepted.", lpCtsProperty.Value))
  '          End If
  '        Else
  '          Throw New ArgumentOutOfRangeException(lpCtsProperty.Name, lpCtsProperty.Value, String.Format("The value '{0}' is not a valid date, it can not be accepted.", lpCtsProperty.Value))
  '        End If
  '      Case PropertyType.ecmLong
  '        If IsNumeric(lpCtsProperty.Value) Then
  '          If lpTargetProperty.Type = PropertyType.ecmLong Then
  '            If lpCtsProperty.Value > Integer.MaxValue Then
  '              Throw New ArgumentOutOfRangeException(lpCtsProperty.Name, lpCtsProperty.Value, String.Format("The value '{0}' is too large for an integer field, it can not be accepted.", lpCtsProperty.Value))
  '            Else
  '              lpP8Document.Properties(lpTargetProperty.SystemName) = Integer.Parse(lpCtsProperty.Value)
  '            End If
  '          ElseIf lpTargetProperty.Type = PropertyType.ecmDouble Then
  '            lpP8Document.Properties(lpTargetProperty.SystemName) = lpCtsProperty.Value
  '          Else
  '            Throw New Exceptions.InvalidPropertyTypeException(lstrInvalidPropertyTypeMessage, lpCtsProperty)
  '          End If

  '          'If TypeOf lpP8Property Is IPropertyInteger32 Then
  '          '  If lpCtsProperty.Value > Integer.MaxValue Then
  '          '    Throw New ArgumentOutOfRangeException(lpCtsProperty.Name, lpCtsProperty.Value, String.Format("The value '{0}' is too large for an integer field, it can not be accepted.", lpCtsProperty.Value))
  '          '  Else
  '          '    lpP8Property.SetObjectValue(Integer.Parse(lpCtsProperty.Value))
  '          '  End If
  '          'ElseIf TypeOf lpP8Property Is IPropertyFloat64 Then
  '          '  lpP8Property.SetObjectValue(lpCtsProperty.Value)
  '          'Else
  '          '  Throw New Exceptions.InvalidPropertyTypeException(lstrInvalidPropertyTypeMessage, lpCtsProperty)
  '          'End If
  '        Else
  '          Throw New ArgumentOutOfRangeException(lpCtsProperty.Name, lpCtsProperty.Value, String.Format("The value '{0}' is not numeric, it can not be accepted.", lpCtsProperty.Value))
  '        End If
  '      Case PropertyType.ecmDouble
  '        If IsNumeric(lpCtsProperty.Value) Then
  '          If lpTargetProperty.Type = PropertyType.ecmDouble Then
  '            If lpCtsProperty.Value > Double.MaxValue Then
  '              Throw New ArgumentOutOfRangeException(lpCtsProperty.Name, lpCtsProperty.Value, String.Format("The value '{0}' is too large for an integer field, it can not be accepted.", lpCtsProperty.Value))
  '            Else
  '              lpP8Document.Properties(lpTargetProperty.SystemName) = lpCtsProperty.Value
  '            End If
  '          Else
  '            Throw New Exceptions.InvalidPropertyTypeException(lstrInvalidPropertyTypeMessage, lpCtsProperty)
  '          End If

  '          'If TypeOf lpP8Property Is IPropertyFloat64 Then
  '          '  If lpCtsProperty.Value > Double.MaxValue Then
  '          '    Throw New ArgumentOutOfRangeException(lpCtsProperty.Name, lpCtsProperty.Value, String.Format("The value '{0}' is too large for an integer field, it can not be accepted.", lpCtsProperty.Value))
  '          '  Else
  '          '    lpP8Property.SetObjectValue(lpCtsProperty.Value)
  '          '  End If
  '          'Else
  '          '  Throw New Exceptions.InvalidPropertyTypeException(lstrInvalidPropertyTypeMessage, lpCtsProperty)
  '          'End If
  '        Else
  '          Throw New ArgumentOutOfRangeException(lpCtsProperty.Name, lpCtsProperty.Value, String.Format("The value '{0}' is not numeric, it can not be accepted.", lpCtsProperty.Value))
  '        End If

  '      Case PropertyType.ecmGuid
  '        If lpTargetProperty.Type = PropertyType.ecmGuid Then
  '          Dim lobjId As Guid = Nothing
  '          If lpCtsProperty.HasValue AndAlso lpCtsProperty.Value IsNot Nothing Then
  '            If Helper.IsGuid(lpCtsProperty.Value.ToString, lobjId) Then
  '              Dim lobjFileNetId As New FileNet.Api.Util.Id(lobjId)
  '              lpP8Document.Properties(lpTargetProperty.SystemName) = lobjFileNetId
  '            End If
  '          End If
  '        Else
  '          Throw New Exceptions.InvalidPropertyTypeException(lstrInvalidPropertyTypeMessage, lpCtsProperty)
  '        End If
  '      Case PropertyType.ecmObject
  '        'Skip this type for now
  '        ''If lobjTargetProperty.Type = PropertyType.ecmObject Then
  '        ''  lpP8Document.Properties(lobjTargetProperty.SystemName) = lpCtsProperty.Value
  '        ''Else
  '        ''  Throw New Exceptions.InvalidPropertyTypeException(lstrInvalidPropertyTypeMessage, lpCtsProperty)
  '        ''End If
  '        'If TypeOf lpP8Property Is IPropertyId Then
  '        '  lpP8Property.SetObjectValue(lpCtsProperty.Value)
  '        'Else
  '        '  Throw New Exceptions.InvalidPropertyTypeException(lstrInvalidPropertyTypeMessage, lpCtsProperty)
  '        'End If
  '      Case Else
  '        Throw New Exceptions.InvalidPropertyTypeException(lstrInvalidPropertyTypeMessage, lpCtsProperty)
  '    End Select
  '  Catch ex As Exception
  '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '    ' Re-throw the exception to the caller
  '    Throw
  '  End Try
  'End Sub

  'Private Sub SetPropertyValue(ByVal lpP8Property As FileNet.Api.Property.IProperty, ByVal lpCtsProperty As ECMProperty)
  '  Try

  '    Dim lstrInvalidPropertyTypeMessage As String = _
  '      String.Format("Destination property '{0}' is type '{1}', the supplied property was of type '{2}'.", _
  '                    lpP8Property.GetPropertyName, lpP8Property.GetType.Name, lpCtsProperty.Type.ToString)

  '    If lpCtsProperty.Cardinality = Cts.Core.Cardinality.ecmSingleValued Then

  '      If lpCtsProperty.Value Is Nothing Then
  '        ' There is no value to set, exit the sub
  '        Exit Sub
  '      End If

  '      Select Case lpCtsProperty.Type
  '        Case PropertyType.ecmString, PropertyType.ecmHtml, PropertyType.ecmUri, PropertyType.ecmXml
  '          If TypeOf lpP8Property Is IPropertyString Then
  '            lpP8Property.SetObjectValue(lpCtsProperty.Value.ToString)
  '          Else
  '            Throw New Exceptions.InvalidPropertyTypeException(lstrInvalidPropertyTypeMessage, lpCtsProperty)
  '          End If
  '        Case PropertyType.ecmBoolean
  '          If TypeOf lpP8Property Is IPropertyBoolean Then
  '            lpP8Property.SetObjectValue(lpCtsProperty.Value)
  '          ElseIf TypeOf lpP8Property Is IPropertyString Then
  '            lpP8Property.SetObjectValue(lpCtsProperty.Value.ToString)
  '          End If
  '        Case PropertyType.ecmDate
  '          If IsDate(lpCtsProperty.Value) Then
  '            If TypeOf lpP8Property Is IPropertyDateTime Then
  '              lpP8Property.SetObjectValue(lpCtsProperty.Value)
  '            ElseIf TypeOf lpP8Property Is IPropertyString Then
  '              lpP8Property.SetObjectValue(lpCtsProperty.Value.ToString)
  '            Else
  '              Throw New Exceptions.InvalidPropertyTypeException(lstrInvalidPropertyTypeMessage, lpCtsProperty)
  '            End If
  '          Else
  '            Throw New ArgumentOutOfRangeException(lpCtsProperty.Name, lpCtsProperty.Value, String.Format("The value '{0}' is not a valid date, it can not be accepted.", lpCtsProperty.Value))
  '          End If
  '        Case PropertyType.ecmLong
  '          If IsNumeric(lpCtsProperty.Value) Then
  '            If TypeOf lpP8Property Is IPropertyInteger32 Then
  '              If lpCtsProperty.Value > Integer.MaxValue Then
  '                Throw New ArgumentOutOfRangeException(lpCtsProperty.Name, lpCtsProperty.Value, String.Format("The value '{0}' is too large for an integer field, it can not be accepted.", lpCtsProperty.Value))
  '              Else
  '                lpP8Property.SetObjectValue(Integer.Parse(lpCtsProperty.Value))
  '              End If
  '            ElseIf TypeOf lpP8Property Is IPropertyFloat64 Then
  '              lpP8Property.SetObjectValue(lpCtsProperty.Value)
  '            Else
  '              Throw New Exceptions.InvalidPropertyTypeException(lstrInvalidPropertyTypeMessage, lpCtsProperty)
  '            End If
  '          Else
  '            Throw New ArgumentOutOfRangeException(lpCtsProperty.Name, lpCtsProperty.Value, String.Format("The value '{0}' is not numeric, it can not be accepted.", lpCtsProperty.Value))
  '          End If
  '        Case PropertyType.ecmDouble
  '          If IsNumeric(lpCtsProperty.Value) Then
  '            If TypeOf lpP8Property Is IPropertyFloat64 Then
  '              If lpCtsProperty.Value > Double.MaxValue Then
  '                Throw New ArgumentOutOfRangeException(lpCtsProperty.Name, lpCtsProperty.Value, String.Format("The value '{0}' is too large for an integer field, it can not be accepted.", lpCtsProperty.Value))
  '              Else
  '                lpP8Property.SetObjectValue(lpCtsProperty.Value)
  '              End If
  '            Else
  '              Throw New Exceptions.InvalidPropertyTypeException(lstrInvalidPropertyTypeMessage, lpCtsProperty)
  '            End If
  '          Else
  '            Throw New ArgumentOutOfRangeException(lpCtsProperty.Name, lpCtsProperty.Value, String.Format("The value '{0}' is not numeric, it can not be accepted.", lpCtsProperty.Value))
  '          End If

  '        Case PropertyType.ecmGuid, PropertyType.ecmObject
  '          If TypeOf lpP8Property Is IPropertyId Then
  '            lpP8Property.SetObjectValue(lpCtsProperty.Value)
  '          Else
  '            Throw New Exceptions.InvalidPropertyTypeException(lstrInvalidPropertyTypeMessage, lpCtsProperty)
  '          End If
  '        Case Else
  '          Throw New Exceptions.InvalidPropertyTypeException(lstrInvalidPropertyTypeMessage, lpCtsProperty)
  '      End Select

  '    Else
  '      ' Set the multi-valued property value
  '      lpP8Property.SetObjectValue(CreateMultiValuedPropertyValue(lpCtsProperty))
  '    End If
  '  Catch ex As Exception
  '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '    ' Re-throw the exception to the caller
  '    Throw
  '  End Try
  'End Sub

  Private Function CreateMultiValuedPropertyValue(ByVal lpECMProperty As ECMProperty) As Object

    Try

      Select Case lpECMProperty.Type

        Case PropertyType.ecmString

          Dim lobjP8Values As IStringList = Factory.StringList.CreateList

          For Each lstrValue As Object In lpECMProperty.Values

            If TypeOf lstrValue Is Value Then
              lobjP8Values.Add(DirectCast(lstrValue, Value).Value)
            ElseIf TypeOf lstrValue Is String Then
              lobjP8Values.Add(lstrValue)
            ElseIf TypeOf lstrValue Is IEnumerable Then
              For Each lstrIndividualValue As String In lstrValue
                lobjP8Values.Add(lstrIndividualValue)
              Next
            Else
              lobjP8Values.Add(lstrValue)
            End If

          Next

          Return lobjP8Values

        Case PropertyType.ecmBoolean

          Dim lobjP8Values As IBooleanList = Factory.BooleanList.CreateList

          For Each lblnValue As Boolean In lpECMProperty.Values
            lobjP8Values.Add(lblnValue)
          Next

          Return lobjP8Values

        Case PropertyType.ecmDate

          Dim lobjP8Values As IDateTimeList = Factory.DateTimeList.CreateList

          For Each ldatValue As Date In lpECMProperty.Values
            lobjP8Values.Add(ldatValue)
          Next

          Return lobjP8Values

        Case PropertyType.ecmBinary

          Dim lobjP8Values As IBinaryList = Factory.BinaryList.CreateList

          For Each lbinValue As Object In lpECMProperty.Values
            lobjP8Values.Add(lbinValue)
          Next

          Return lobjP8Values

        Case PropertyType.ecmObject
          ' TODO: Implement this case
          Return Nothing

        Case PropertyType.ecmLong

          Dim lobjP8Values As IInteger32List = Factory.Integer32List.CreateList

          For Each llngValue As Long In lpECMProperty.Values
            If llngValue > Integer.MaxValue Then
              Throw New ArgumentOutOfRangeException(lpECMProperty.Name, lpECMProperty.Value,
                                                    String.Format(
                                                      "The value '{0}' is too large for an integer field, it can not be accepted.",
                                                      lpECMProperty.Value))
            End If
            lobjP8Values.Add(Integer.Parse(llngValue))
          Next

          Return lobjP8Values

        Case PropertyType.ecmDouble

          Dim lobjP8Values As IFloat64List = Factory.Float64List.CreateList

          For Each ldblValue As Double In lpECMProperty.Values
            lobjP8Values.Add(ldblValue)
          Next

          Return lobjP8Values

        Case PropertyType.ecmGuid

          Dim lobjP8Values As IIdList = Factory.IdList.CreateList

          For Each lguidValue As String In lpECMProperty.Values
            lobjP8Values.Add(New Id(lguidValue))
          Next

          Return lobjP8Values

        Case Else
          Return Nothing

      End Select

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function CreateValues(ByVal lpIList As IList) As Values

    Try

      Select Case CType(lpIList, Object).GetType.ToString

      End Select

      Return Nothing

      'Select Case lpECMProperty.Type

      '  Case PropertyType.ecmString
      '    Dim lobjP8Values As FileNet.Api.Collection.IStringList = Factory.StringList.CreateList
      '    For Each lstrValue As String In lpECMProperty.Values
      '      lobjP8Values.Add(lstrValue)
      '    Next
      '    Return lobjP8Values

      '  Case PropertyType.ecmBoolean
      '    Dim lobjP8Values As FileNet.Api.Collection.IBooleanList = Factory.BooleanList.CreateList
      '    For Each lblnValue As Boolean In lpECMProperty.Values
      '      lobjP8Values.Add(lblnValue)
      '    Next
      '    Return lobjP8Values

      '  Case PropertyType.ecmDate
      '    Dim lobjP8Values As FileNet.Api.Collection.IDateTimeList = Factory.DateTimeList.CreateList
      '    For Each ldatValue As Boolean In lpECMProperty.Values
      '      lobjP8Values.Add(ldatValue)
      '    Next
      '    Return lobjP8Values

      '  Case PropertyType.ecmBinary
      '    Dim lobjP8Values As FileNet.Api.Collection.IBinaryList = Factory.BinaryList.CreateList
      '    For Each lbinValue As Boolean In lpECMProperty.Values
      '      lobjP8Values.Add(lbinValue)
      '    Next
      '    Return lobjP8Values

      '  Case PropertyType.ecmObject
      '    ' TODO: Implement this case
      '    Return Nothing

      '  Case PropertyType.ecmLong
      '    Dim lobjP8Values As FileNet.Api.Collection.IInteger32List = Factory.Integer32List.CreateList
      '    For Each llngValue As Long In lpECMProperty.Values
      '      lobjP8Values.Add(llngValue)
      '    Next
      '    Return lobjP8Values

      '  Case PropertyType.ecmDouble
      '    Dim lobjP8Values As FileNet.Api.Collection.IFloat64List = Factory.Float64List.CreateList
      '    For Each ldblValue As Double In lpECMProperty.Values
      '      lobjP8Values.Add(ldblValue)
      '    Next
      '    Return lobjP8Values

      '  Case PropertyType.ecmGuid
      '    Dim lobjP8Values As FileNet.Api.Collection.IIdList = Factory.IdList.CreateList
      '    For Each lguidValue As String In lpECMProperty.Values
      '      lobjP8Values.Add(lguidValue)
      '    Next
      '    Return lobjP8Values

      '  Case Else
      '    Return Nothing

      'End Select

      'Return Nothing
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function CreateContentElementList(ByVal lpContents As Contents) As IContentElementList

    Try

      ArgumentNullException.ThrowIfNull(lpContents)

      Dim lobjContentTransferList As IContentElementList = Factory.ContentElement.CreateList

      For Each lobjContent As Content In lpContents
        lobjContentTransferList.Add(CreateContentTransfer(lobjContent))
      Next

      Return lobjContentTransferList

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function CreateContentElementList(ByVal lpContentContainer As IContentContainer) As IContentElementList
    Try

      Dim lobjContentTransferList As IContentElementList = Factory.ContentElement.CreateList

      lobjContentTransferList.Add(CreateContentTransfer(lpContentContainer))

      Return lobjContentTransferList

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function CreateContentTransfer(ByVal lpContent As Content) As IContentTransfer

    Try

      Dim lobjContentTransfer As IContentTransfer = Factory.ContentTransfer.CreateInstance
      'Dim lobjStreamReader As New StreamReader(lpContent.ContentPath)

      Dim lintMaximumInMemoryDocumentMegabytes As Integer = SysConfig.AppSettings.Item("MaximumInMemoryDocumentMegabytes")

      With lobjContentTransfer
        '.SetCaptureSource(lobjStreamReader.BaseStream)

        'If lpContent.FileSize.Megabytes < ConnectionSettings.Instance.MaximumInMemoryDocumentMegabytes Then
        If lpContent.FileSize.Megabytes < lintMaximumInMemoryDocumentMegabytes Then
          .SetCaptureSource(lpContent.ToMemoryStream)
        Else
          ' TODO: Change this to use a file stream
          ApplicationLogging.WriteLogEntry(String.Format("Content file '{0}': {1} is larger than the configured maximum in memory document megabytes ({2}).",
                                                         lpContent.FileName,
                                                         lpContent.FileSize.ToString(),
                                                         lintMaximumInMemoryDocumentMegabytes),
                                           MethodBase.GetCurrentMethod(),
                                           TraceEventType.Warning, 29875)
          Dim lobjContentStream As Stream = lpContent.ToStream()
          ApplicationLogging.WriteLogEntry(String.Format("Using stream type '{0}'.", lobjContentStream.GetType.Name),
                                           MethodBase.GetCurrentMethod(),
                                           TraceEventType.Information,
                                           29876)
          .SetCaptureSource(lobjContentStream)
        End If
        '.SetCaptureSource(lpContent.ToMemoryStream)
        .RetrievalName = lpContent.FileName

        If lpContent.MIMEType.Length > 0 Then
          .ContentType = lpContent.MIMEType
        End If

      End With

      Return lobjContentTransfer

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function CreateContentTransfer(ByVal lpContentContainer As IContentContainer) As IContentTransfer
    Try

      Dim lobjContentTransfer As IContentTransfer = Factory.ContentTransfer.CreateInstance

      With lobjContentTransfer
        '.SetCaptureSource(lobjStreamReader.BaseStream)
        If lpContentContainer.CanRead = False AndAlso Me.AllowZeroLengthContent = False Then
          Throw New ZeroLengthContentException
        End If

        If TypeOf lpContentContainer.FileContent Is Stream Then
          .SetCaptureSource(lpContentContainer.FileContent)
        Else
          .SetCaptureSource(Helper.CopyByteArrayToStream(lpContentContainer.FileContent))
        End If

        .RetrievalName = lpContentContainer.FileName

        If lpContentContainer.MimeType.Length > 0 Then
          .ContentType = lpContentContainer.MimeType
        End If

      End With

      Return lobjContentTransfer

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function GetAvailableParentFolderByPath(ByVal lpFolderPath As String) As FileNet.Api.Core.IFolder
    Try

      Dim lobjFolderTree As New FolderTree(lpFolderPath)
      Dim lobjRootFolder As FileNet.Api.Core.IFolder = ObjectStore.RootFolder
      Dim lobjParentFolder As FileNet.Api.Core.IFolder = Nothing
      Dim lobjFolder As FileNet.Api.Core.IFolder = Nothing

      ' See if we can get the existing folder
      If FolderPathExists(lpFolderPath) Then
        lobjFolder = GetObject("Folder", lpFolderPath)
        lobjFolder.Refresh()

        If lobjFolder IsNot Nothing Then
          Return lobjFolder
        End If

      End If

      ' Clear the reference to lobjFolder to see if we can find it through recursion
      lobjFolder = Nothing

      lobjParentFolder = lobjRootFolder

      For Each lobjFolderInfo As FolderTree.FolderInfo In lobjFolderTree.Folders

        If lobjFolderInfo.Order <= 1 Then

          ' We are looking for a root level folder
          ' Try to find an existing folder
          Try

            For Each lobjSubFolder As FileNet.Api.Core.IFolder In lobjRootFolder.SubFolders

              If lobjSubFolder.Name = lobjFolderInfo.Name Then
                lobjFolder = lobjSubFolder
                Exit For
              End If

            Next

            If lobjFolder Is Nothing Then
              ' We did not find the folder, we need to create it.
              'Dim lstrFolderName As String = lobjFolderInfo.Name
              'If lstrFolderName.StartsWith("\") Then
              '  lstrFolderName = lstrFolderName.Remove(0, 1)
              'End If
              'lobjFolder = lobjRootFolder.CreateSubFolder(lstrFolderName)

              ''lobjFolder = lobjRootFolder.CreateSubFolder(lobjFolderInfo.Name)

              ''lobjFolder.Save(RefreshMode.REFRESH)
              Exit For
            Else
              lobjParentFolder = lobjFolder
            End If

          Catch ex As Exception
            ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
            'Beep()
          End Try

        Else
          ' This is not a root level folder

          Try

            For Each lobjSubFolder As FileNet.Api.Core.IFolder In lobjFolder.SubFolders

              If lobjSubFolder.Name = lobjFolderInfo.Name Then
                lobjParentFolder = lobjFolder
                lobjFolder = lobjSubFolder
                Exit For
              End If

            Next

            If lobjFolder.Name <> lobjFolderInfo.Name Then
              '  ' We did not find the folder, we need to create it.
              '  lobjFolder = lobjFolder.CreateSubFolder(lobjFolderInfo.Name)
              '  lobjFolder.Save(RefreshMode.REFRESH)
              lobjParentFolder = lobjFolder
            End If

          Catch ex As Exception
            ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
            'Beep()
          End Try

        End If

      Next

      'lobjRootFolder.Refresh()

      Return lobjParentFolder

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function GetFolderByPath(ByVal lpFolderPath As String) As FileNet.Api.Core.IFolder

    Try

      Dim lobjFolderTree As New FolderTree(lpFolderPath)
      Dim lobjRootFolder As FileNet.Api.Core.IFolder = ObjectStore.RootFolder
      Dim lobjFolder As FileNet.Api.Core.IFolder = Nothing

      ' See if we can get the existing folder
      If FolderPathExists(lpFolderPath) Then
        lobjFolder = GetObject("Folder", lpFolderPath)
        lobjFolder.Refresh()

        If lobjFolder IsNot Nothing Then
          Return lobjFolder
        End If

      End If

      'If Not lobjFolder Is Nothing Then
      '  Return lobjFolder
      'End If

      ' Clear the reference to lobjFolder to see if we can find it through recursion
      lobjFolder = Nothing

      For Each lobjFolderInfo As FolderTree.FolderInfo In lobjFolderTree.Folders

        If lobjFolderInfo.Order <= 1 Then

          ' We are looking for a root level folder
          ' Try to find an existing folder
          Try

            For Each lobjSubFolder As FileNet.Api.Core.IFolder In lobjRootFolder.SubFolders

              If lobjSubFolder.Name = lobjFolderInfo.Name Then
                lobjFolder = lobjSubFolder
                Exit For
              End If

            Next

            If lobjFolder Is Nothing Then
              ' We did not find the folder, we need to create it.
              'Dim lstrFolderName As String = lobjFolderInfo.Name
              'If lstrFolderName.StartsWith("\") Then
              '  lstrFolderName = lstrFolderName.Remove(0, 1)
              'End If
              'lobjFolder = lobjRootFolder.CreateSubFolder(lstrFolderName)

              lobjFolder = lobjRootFolder.CreateSubFolder(lobjFolderInfo.Name)

              lobjFolder.Save(RefreshMode.REFRESH)
            End If

          Catch ex As Exception
            ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
            'Beep()
          End Try

        Else
          ' This is not a root level folder

          Try

            For Each lobjSubFolder As FileNet.Api.Core.IFolder In lobjFolder.SubFolders

              If lobjSubFolder.Name = lobjFolderInfo.Name Then
                lobjFolder = lobjSubFolder
                Exit For
              End If

            Next

            If lobjFolder.Name <> lobjFolderInfo.Name Then
              ' We did not find the folder, we need to create it.
              lobjFolder = lobjFolder.CreateSubFolder(lobjFolderInfo.Name)
              lobjFolder.Save(RefreshMode.REFRESH)
            End If

          Catch ex As Exception
            ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
            'Beep()
          End Try

        End If

      Next

      lobjRootFolder.Refresh()

      Return lobjFolder

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function


#Region "Private Methods"

  Private Sub InitializeRootFolder()

    Try

      If IsInitialized Then

        Dim lobjRootIFolder As FileNet.Api.Core.IFolder = ObjectStore.RootFolder

        mobjFolder = New CENetFolder(lobjRootIFolder, Me, -1) With {
          .InvisiblePassThrough = True,
          .Provider = Me,
          .Name = lobjRootIFolder.Name
        }

      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Private Function GetConfigurationArray(lpSectionName As String) As Object
    Try
      Dim lobjBuilder As New ConfigurationBuilder()
      lobjBuilder.SetBasePath(Directory.GetCurrentDirectory())
      lobjBuilder.AddJsonFile("p8settings.json", False)

      Dim lobjConfig As IConfiguration = lobjBuilder.Build()
      Dim lobjSection As Object = lobjConfig.GetSection(lpSectionName)

      Return lobjSection

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function GetAllContentExportPropertyExclusionsCollection() As StringCollection

    Try

      Dim lobjPropertyExclusions As StringCollection

      'lobjPropertyExclusions = My.Settings.AppContentExportPropertyExclusions
      lobjPropertyExclusions = GetConfigurationArray("AppContentExportPropertyExclusions")

      'For Each lstrPropertyExclusion As String In My.Settings.UserContentExportPropertyExclusions
      '  lobjPropertyExclusions.Add(lstrPropertyExclusion)
      'Next

      For Each lstrPropertyExclusion As String In lobjPropertyExclusions
        lobjPropertyExclusions.Add(lstrPropertyExclusion)
      Next

      Return lobjPropertyExclusions

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function GetAllContentExportPropertyExclusions() As String()

    Try

      Dim lobjPropertyExclusions As StringCollection = GetAllContentExportPropertyExclusionsCollection()

      If lobjPropertyExclusions.Count > 0 Then

        Dim lstrPropertyExclusions As String()
        ReDim lstrPropertyExclusions(lobjPropertyExclusions.Count - 1)

        For lintExclusionCounter As Int16 = 0 To lobjPropertyExclusions.Count - 1
          lstrPropertyExclusions(lintExclusionCounter) = lobjPropertyExclusions(lintExclusionCounter)
        Next

        Return lstrPropertyExclusions

      Else

        Return Nothing

      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function


  ''' <summary>
  '''   Checks to see if priveledged access has been granted for setting selected system property values.
  ''' </summary>
  ''' <returns>True if current logged on user has priveledged access.</returns>
  ''' <remarks>
  '''   See https://www-304.ibm.com/support/docview.wss?uid=swg21468247&amp;wv=1 for more information.
  ''' </remarks>
  Public Function GetHasPrivelegedAccess() As Boolean

    Try

      If IsInitialized = False Then
        Return False
      End If

      Dim lstrErrorMessage As String = String.Empty
      Dim lblnHasPrivelegedAccess As Boolean
      Dim lintAccessAllowed As Integer = ObjectStore.GetAccessAllowed

      If (Constants.AccessRight.PRIVILEGED_WRITE And lintAccessAllowed) <> 0 Then
        lblnHasPrivelegedAccess = True
      End If

      If lblnHasPrivelegedAccess Then
        Return True
      End If

      Return False

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Friend Function GetObject(ByVal classIdent As String, ByVal objectId As Id) As IIndependentObject
    Try
      Return ObjectStore.GetObject(classIdent, objectId)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Friend Function GetObject(ByVal classIdent As String, ByVal objectIdent As String) As IIndependentObject
    Try
      Return ObjectStore.GetObject(classIdent, objectIdent)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Friend Function GetStorageArea(lpStorageAreaName As String) As IStorageArea
    Try
      If StorageAreaExists(lpStorageAreaName) = False Then
        Throw New StorageAreaNotFoundException(lpStorageAreaName)
      End If

      Dim lobjStorageAreas As IStorageAreaSet = ObjectStore.StorageAreas

      For Each lobjStorageArea As IStorageArea In lobjStorageAreas
        If _
          String.Equals(lobjStorageArea.DisplayName, lpStorageAreaName, StringComparison.InvariantCultureIgnoreCase) =
          True Then
          Return lobjStorageArea
        End If
      Next

      Throw New StorageAreaNotFoundException(lpStorageAreaName)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Friend Function StorageAreaExists(lpStorageAreaName As String) As Boolean
    Try
      Dim lobjStorageAreas As IStorageAreaSet = ObjectStore.StorageAreas

      For Each lobjStorageArea As IStorageArea In lobjStorageAreas
        If _
          String.Equals(lobjStorageArea.DisplayName, lpStorageAreaName, StringComparison.InvariantCultureIgnoreCase) =
          True Then
          Return True
        End If
      Next
      Return False
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function DocumentExists(ByVal lpId As String) As Boolean
    Try
      Return DocumentExists(lpId, "Document")
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function DocumentExists(ByVal lpId As String, ByVal lpQueryTarget As String) As Boolean

    Try

      Dim lobjSearch As New CENetSearch(Me)
      Dim lobjIdCriterion As New Criterion("Id", "object_id") With {
        .Value = lpId
      }

      If mblnTreatObjectIdsAsNumbers = True Then
        lobjIdCriterion.DataType = Criterion.pmoDataType.ecmLong
        ' Temporarily add the SQL statement to the log to help with debugging
        ApplicationLogging.LogInformation(lobjIdCriterion.ToString)
      End If

      lobjSearch.Criteria.Add(lobjIdCriterion)

      lobjSearch.DataSource.QueryTarget = lpQueryTarget

      '' Temporarily add the SQL statement to the log to help with debugging
      'ApplicationLogging.LogInformation(lobjSearch.DataSource.SQLStatement)

      Dim lobjSearchResultSet As DCore.SearchResultSet = lobjSearch.Execute

      If lobjSearchResultSet.Count > 0 Then
        Return True

      Else

        ApplicationLogging.WriteLogEntry(String.Format("Document '{0}' not found using search: '{1}'",
                                                       lpId, lobjSearch.DataSource.SQLStatement),
                                         MethodBase.GetCurrentMethod(),
                                         TraceEventType.Information, 51404)
        ApplicationLogging.WriteLogEntry(
          String.Format("Current connection information(Domain: {0}, Object Store: {1}, UserName: {2})",
                        lobjSearch.Domain.Name, lobjSearch.ObjectStoreName, Me.UserName),
          MethodBase.GetCurrentMethod(),
          TraceEventType.Information, 51405)

        Return False
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region


  ''' <summary>
  '''   Call this when you need a fresh search
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Overrides Function CreateSearch() As DProviders.ISearch

    Try
      Return New CENetSearch(Me)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#Region "Initialization Methods"

  Private Function CreateURL() As String

    Try
      'Dim lstrURL As String

      'lstrURL = mstrProtocol & "://" & mstrServerName & ":" & mstrPortNumber & "/wsi/FNCEWS40MTOM/"

      'Return lstrURL

      If Not mstrPortNumber = "80" Then
        Return String.Format("{0}://{1}:{2}/wsi/FNCEWS40MTOM/", mstrProtocol, mstrServerName, mstrPortNumber)
      Else
        Return String.Format("{0}://{1}/wsi/FNCEWS40MTOM/", mstrProtocol, mstrServerName)
      End If


    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Overrides Sub InitializeProvider(ByVal lpContentSource As ContentSource)

    Try
      MyBase.InitializeProvider(lpContentSource)
      InitializeProperties()
      InitializeConnection()
      mobjSearch = CreateSearch()

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Private Sub InitializeProperties()

    Try

      For Each lobjProviderProperty As ProviderProperty In ProviderProperties

        Select Case lobjProviderProperty.PropertyName
          'Case DOMAIN_NAME
          '  mstrDomainName = lobjProviderProperty.PropertyValue

          Case SERVER_NAME
            mstrServerName = lobjProviderProperty.PropertyValue

          Case OBJECT_STORE_NAME
            mstrObjectStoreName = lobjProviderProperty.PropertyValue

          Case PORT_NUMBER
            mstrPortNumber = lobjProviderProperty.PropertyValue

          Case PROTOCOL
            mstrProtocol = lobjProviderProperty.PropertyValue

          Case PING_TO_VERIFY
            mblnPingToVerify = lobjProviderProperty.PropertyValue

          Case PRESERVE_ID_ON_IMPORT
            mblnPreserveIdOnImport = Boolean.Parse(lobjProviderProperty.PropertyValue)

          Case TREAT_OBJECT_IDS_AS_NUMBERS
            ' Temporary change for Conagra
            'mblnTreatObjectIdsAsNumbers = Boolean.Parse(lobjProviderProperty.PropertyValue)
            mblnTreatObjectIdsAsNumbers = True

            ' Temporarily add to the log to help with debugging
            ApplicationLogging.LogInformation(String.Format("mblnTreatObjectIdsAsNumbers: {0}", mblnTreatObjectIdsAsNumbers.ToString()))

          Case EXPORT_SYSTEM_OBJECT_VALUED_PROPERTIES
            mblnExportSystemObjectValuedProperties = Boolean.Parse(lobjProviderProperty.PropertyValue)

          Case TRUSTED_CONNECTION
            mblnTrustedConnection = Boolean.Parse(lobjProviderProperty.PropertyValue)

          Case USER
            UserName = lobjProviderProperty.PropertyValue

          Case PWD
            Password = lobjProviderProperty.PropertyValue

          Case EXPORT_PATH
            ExportPath = lobjProviderProperty.PropertyValue

            ' <Removed by: Ernie at: 9/29/2014-11:18:48 AM on machine: ERNIE-THINK>
            '           Case IMPORT_PATH
            '             ImportPath = lobjProviderProperty.PropertyValue
            ' </Removed by: Ernie at: 9/29/2014-11:18:48 AM on machine: ERNIE-THINK>

        End Select

      Next

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Private Sub ParseConnectionString()

    'Try
    '  mstrDomainName = Helper.GetInfoFromString(ConnectionString, DOMAIN_NAME)
    'Catch ex As Exception
    '  ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '  Throw New ArgumentException("Argument not provided", DOMAIN_NAME, ex)
    'End Try

    Try
      mstrServerName = Helper.GetInfoFromString(ConnectionString, SERVER_NAME)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      Throw New ArgumentException("Argument not provided", SERVER_NAME, ex)
    End Try

    Try
      mstrObjectStoreName = Helper.GetInfoFromString(ConnectionString, OBJECT_STORE_NAME)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      Throw New ArgumentException("Argument not provided", OBJECT_STORE_NAME, ex)
    End Try

    Try
      mstrPortNumber = Helper.GetInfoFromString(ConnectionString, PORT_NUMBER)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      Throw New ArgumentException("Argument not provided", PORT_NUMBER, ex)
    End Try

    Try
      mstrProtocol = Helper.GetInfoFromString(ConnectionString, PROTOCOL)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      Throw New ArgumentException("Argument not provided", PROTOCOL, ex)
    End Try

    Try
      mstrUserName = Helper.GetInfoFromString(ConnectionString, "UserName")

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      Throw New ArgumentException("Argument not provided", "UserName", ex)
    End Try

    Try
      mstrPassword = Helper.GetInfoFromString(ConnectionString, "Password")

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      Throw New ArgumentException("Argument not provided", "Password", ex)
    End Try

    Try
      mblnPingToVerify = Boolean.Parse(Helper.GetInfoFromString(ConnectionString, PING_TO_VERIFY))
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      Throw New ArgumentException("Argument not provided", PING_TO_VERIFY, ex)
    End Try

    Try
      ExportPath = Helper.GetInfoFromString(ConnectionString, EXPORT_PATH)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ExportPath = DEFAULT_EXPORT_PATH
    End Try

    ' <Removed by: Ernie at: 9/29/2014-11:19:06 AM on machine: ERNIE-THINK>
    '     Try
    '       ImportPath = Helper.GetInfoFromString(ConnectionString, IMPORT_PATH)
    ' 
    '     Catch ex As Exception
    '       ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
    '       ImportPath = My.Settings.DefaultImportPath
    '     End Try
    ' </Removed by: Ernie at: 9/29/2014-11:19:06 AM on machine: ERNIE-THINK>
  End Sub

  Private Function InitializeConnection() As Boolean

    Debug.WriteLine("Please wait, attempting to connect...")
    Dim lstrErrorMessage As String = String.Empty

    Try


      If PingToVerify AndAlso CEServerAvailable(lstrErrorMessage) = False Then
        ApplicationLogging.WriteLogEntry("The CE Server is unavailable", MethodBase.GetCurrentMethod, TraceEventType.Warning, 13649)
        Throw New RepositoryNotAvailableException(Me.ServerName, lstrErrorMessage, ExceptionTracker.LastException)
      End If

      Dim lobjCredentials As Credentials = Nothing
      'Dim ctx As WindowsImpersonationContext = Nothing

      If TrustedConnection Then
        If Me.ContentSource.SecurityToken IsNot Nothing Then
          Try
            'Dim lobjIdent As WindowsIdentity = CType(Me.ContentSource.SecurityToken, System.Security.Principal.WindowsIdentity)
            'ApplicationLogging.WriteLogEntry(lobjIdent.Name, TraceEventType.Information)
            'ctx = lobjIdent.Impersonate()
            'lobjCredentials = New Authentication.KerberosCredentials
            'If (ctx IsNot Nothing) Then
            '  ctx.Undo()
            'End If

            'lobjCredentials = New Authentication.KerberosCredentials
            ClientContext.SetProcessCredentials(Me.ContentSource.SecurityToken)

            'Dim lobjKerbToken As Microsoft.Web.Services3.Security.Tokens.KerberosToken
            'Try
            '  lobjKerbToken = New Microsoft.Web.Services3.Security.Tokens.KerberosToken("FNCEWS/" + "GTNAHOUWAS978")
            'Catch ex As Exception
            '  ApplicationLogging.WriteLogEntry("Failed to create kerberos token " + ex.Message, TraceEventType.Error)
            'End Try

            'Try
            '  UserContext.SetProcessSecurityToken(lobjKerbToken)
            'Catch ex As Exception
            '  ApplicationLogging.WriteLogEntry("Failed to Set Process Security token " + ex.Message, TraceEventType.Error)
            'End Try

          Catch ex As Exception
            ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
          End Try
        Else
          lobjCredentials = New KerberosCredentials
          ClientContext.SetProcessCredentials(lobjCredentials)
        End If
      Else
        lobjCredentials = New UsernameCredentials(UserName, Password)
        ClientContext.SetProcessCredentials(lobjCredentials)
      End If

      'Util.ClientContext.SetProcessCredentials(lobjCredentials)
      'Util.ClientContext.SetThreadCredentials(lobjCredentials)


      'Dim lobjPropertyFilter As New PropertyFilter
      'Dim lobjEntireNetwork As IEntireNetwork = Factory.EntireNetwork.FetchInstance(lobjConnection, Nothing)

      'Console.WriteLine("Got EntireNetwork object.")

      ApplicationLogging.WriteLogEntry("Attempting to get domain", TraceEventType.Information)
      mobjDomain = GetDomain()
      ApplicationLogging.WriteLogEntry("Successfully got domain", TraceEventType.Information)

      Try
        ApplicationLogging.WriteLogEntry("Attempting to refresh domain", TraceEventType.Information)
        mobjDomain.Refresh()
        ApplicationLogging.WriteLogEntry("Successfully refreshed domain", TraceEventType.Information)

        'Catch runtimeException As FileNet.Api.Exception.EngineRuntimeException
        '  Throw New InvalidOperationException(String.Format("Invalid domain name: The domain name '{0}' is invalid.  Please supply a valid P8 domain name for the connection.", _
        '                                                    DomainName), runtimeException)

        'Catch executionEngineException As ExecutionEngineException
        '  ApplicationLogging.LogException(executionEngineException.InnerException, Reflection.MethodBase.GetCurrentMethod)
        '  Dim bDomainRefreshSuccess As Boolean = False
        '  For i As Integer = 0 To 4

        '    Try
        '      mobjDomain.Refresh()
        '      bDomainRefreshSuccess = True
        '      Exit For
        '    Catch ex As Exception
        '      ApplicationLogging.LogException(executionEngineException.InnerException, Reflection.MethodBase.GetCurrentMethod)
        '    End Try

        '  Next

        '  If (bDomainRefreshSuccess = False) Then
        '    ApplicationLogging.WriteLogEntry("Failed to refresh the Domain", TraceEventType.Error, 64298)
        '    IsConnected = False
        '    SetState(ProviderConnectionState.Unavailable)
        '    '  Re-throw the exception
        '    Throw
        '  End If

      Catch runtimeException As EngineRuntimeException
        ApplicationLogging.LogException(runtimeException, MethodBase.GetCurrentMethod)
        If runtimeException.Message.Contains("network error") Then
          If ((runtimeException.InnerException IsNot Nothing) AndAlso
              (TypeOf runtimeException.InnerException Is WebException)) Then
            Dim lobjWebEx As WebException = runtimeException.InnerException
            If ((lobjWebEx.InnerException IsNot Nothing) AndAlso
                (TypeOf lobjWebEx.InnerException Is SocketException)) Then
              ApplicationLogging.LogException(lobjWebEx.InnerException, MethodBase.GetCurrentMethod)
              IsConnected = False
              SetState(ProviderConnectionState.Unavailable)
              '  Re-throw the exception
              Throw New RepositoryNotConnectedException(
                String.Format(
                  "Unable to connect to server, make sure the server name '{0}' and port '{1}' are correct.",
                  Me.ServerName, Me.PortNumber), lobjWebEx.InnerException)
            Else
              ApplicationLogging.LogException(lobjWebEx, MethodBase.GetCurrentMethod)
              IsConnected = False
              SetState(ProviderConnectionState.Unavailable)
              '  Re-throw the exception
              Throw lobjWebEx
            End If
          Else
            ApplicationLogging.LogException(runtimeException, MethodBase.GetCurrentMethod)
            IsConnected = False
            SetState(ProviderConnectionState.Unavailable)
            '  Re-throw the exception
            Throw
          End If
        ElseIf runtimeException.Message.Contains("not authenticated") Then
          ApplicationLogging.LogException(runtimeException, MethodBase.GetCurrentMethod)
          IsConnected = False
          SetState(ProviderConnectionState.Unavailable)
          Dim lobjErrorMessageBuilder As New StringBuilder
          If Not TrustedConnection Then
            lobjErrorMessageBuilder.AppendFormat("Login failed for user '{0}', make sure the username is valid and the password is not expired.", UserName)
          Else
            lobjErrorMessageBuilder.Append("Trusted authentication failed.")
          End If
          Throw New RepositoryAuthenticationException(lobjErrorMessageBuilder.ToString(), runtimeException)
        Else
          ApplicationLogging.LogException(runtimeException, MethodBase.GetCurrentMethod)
          IsConnected = False
          SetState(ProviderConnectionState.Unavailable)
          '  Re-throw the exception
          Throw
        End If
      Catch ex As Exception

        ApplicationLogging.LogException(ex.InnerException, MethodBase.GetCurrentMethod)
        Dim bDomainRefreshSuccess As Boolean = False
        For i As Integer = 0 To 4
          ApplicationLogging.LogInformation(String.Format("Retrying Domain Refresh {0}", i))
          Try
            mobjDomain.Refresh()
            bDomainRefreshSuccess = True
            Exit For
          Catch InnerEx As Exception
            ApplicationLogging.LogException(InnerEx.InnerException, MethodBase.GetCurrentMethod)
            IsConnected = False
            SetState(ProviderConnectionState.Unavailable)
            '  Re-throw the exception
            Throw
          End Try

        Next

        If (bDomainRefreshSuccess = False) Then
          ApplicationLogging.WriteLogEntry("Failed to refresh the Domain", TraceEventType.Error, 64298)
          IsConnected = False
          SetState(ProviderConnectionState.Unavailable)
          '  Re-throw the exception
          Throw
        End If

        'ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        'IsConnected = False
        'SetState(ProviderConnectionState.Unavailable)
        ''  Re-throw the exception
        'Throw
      End Try

      'Debug.WriteLine("Got Domain Object: " & mobjDomain.Name)
      'Debug.WriteLine("ID: " & mobjDomain.Id.ToString)

      ' If we already have the object store name, complete the process
      If Not String.IsNullOrEmpty(ObjectStoreName) Then

        Try
          mobjObjectStore = Factory.ObjectStore.FetchInstance(mobjDomain, ObjectStoreName, Nothing)

        Catch runtimeException As EngineRuntimeException
          Throw _
            New InvalidOperationException(
              String.Format(
                "Invalid object store name: The object store name '{0}' is invalid.  Please supply a valid P8 object store name for the connection.",
                ObjectStoreName), runtimeException)

        Catch ex As Exception
          ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
          IsConnected = False
          SetState(ProviderConnectionState.Unavailable)
          '  Re-throw the exception
          Throw
        End Try

        IsConnected = True

        Debug.WriteLine("Got ObjectStore. ID:" & mobjObjectStore.Id.ToString)
        Debug.WriteLine("Connection Initialized")

      End If

      Return True
    Catch RepNotAvailEx As RepositoryNotAvailableException
      ApplicationLogging.LogException(RepNotAvailEx, MethodBase.GetCurrentMethod)
      IsConnected = False
      SetState(ProviderConnectionState.Unavailable)
      Throw
    Catch RepNotConEx As RepositoryNotConnectedException
      ApplicationLogging.LogException(RepNotConEx, MethodBase.GetCurrentMethod)
      IsConnected = False
      SetState(ProviderConnectionState.Unavailable)
      Throw
    Catch RepAuthEx As RepositoryAuthenticationException
      ApplicationLogging.LogException(RepAuthEx, MethodBase.GetCurrentMethod)
      IsConnected = False
      SetState(ProviderConnectionState.Unavailable)
      Throw
    Catch runtimeException As EngineRuntimeException
      ApplicationLogging.LogException(runtimeException, MethodBase.GetCurrentMethod)
      IsConnected = False
      SetState(ProviderConnectionState.Unavailable)
      If runtimeException.Message.Contains("HTTPS is required") Then
        Throw _
        (New RepositoryNotConnectedException(
          "Unable to initialize connection, if connecting to P8 5.2 or older verify that Microsoft Web Services Enhancements (WSE) 3.0 is installed.",
          runtimeException))
      Else
        Throw _
        (New RepositoryNotConnectedException(
          "Unable to initialize connection, make sure that all the elements of the connection string, including user name and password are correct.",
          runtimeException))
      End If
    Catch ex As Exception
      If ex.InnerException IsNot Nothing Then
        ApplicationLogging.WriteLogEntry(String.Format("{0} --- {1}", ex.Message, ex.InnerException.Message), TraceEventType.Error)
      Else
        ApplicationLogging.WriteLogEntry(ex.Message, TraceEventType.Error)
      End If
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      IsConnected = False
      SetState(ProviderConnectionState.Unavailable)
      Throw (New RepositoryNotConnectedException("Unable to initialize connection", ex))
    End Try
  End Function

  Private Function CEServerAvailable(Optional ByRef lpErrorMessage As String = "") As Boolean
    Dim lstrURL As String = URL '"http://" & Me.P8ContentEngine & ":" & Me.Port & "/FNCEWS35SOAP/WSDL"
    Dim lblnCanPing As Boolean = False
    Try
      'ApplicationLogging.WriteLogEntry(lstrURL, MethodBase.GetCurrentMethod, TraceEventType.Information, 13621)
      Try
        ApplicationLogging.WriteLogEntry(String.Format("Pinging Server: {0}", Me.ServerName),
                                         MethodBase.GetCurrentMethod, TraceEventType.Information, 13622)
        lblnCanPing = Helper.Ping(Me.ServerName, 5000)
      Catch PingEx As PingException
        'If PingEx.InnerException IsNot Nothing AndAlso PingEx.InnerException.Message = "No such host is known" Then
        lpErrorMessage =
          String.Format(
            "Host '{0}' Unknown, make sure that the server name supplied is correct. If so ensure that a network connection is available and DNS is functioning.",
            Me.ServerName)
        ApplicationLogging.WriteLogEntry(lpErrorMessage, MethodBase.GetCurrentMethod, TraceEventType.Error, 13631)
        Return False
        'End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

      If lblnCanPing = False Then
        lpErrorMessage = "Unable to ping Content Engine server at '" & ServerName &
                         "', verify that server name and network connection is correct."
        ApplicationLogging.WriteLogEntry(lpErrorMessage, MethodBase.GetCurrentMethod, TraceEventType.Error, 13632)
        Return False
      Else
        ApplicationLogging.WriteLogEntry(String.Format("Server ping succesfull: {0}", Me.ServerName),
                                 MethodBase.GetCurrentMethod, TraceEventType.Information, 13623)
      End If
      Dim lstrWSDLPath As String = ExportPath & "WSDL.xml"
      If File.Exists(lstrWSDLPath) Then
        File.Delete(lstrWSDLPath)
      End If

      ' Try to download the WSDL
      Dim lobjWebClient As New WebClient


      '' The next three lines are an attempt to gracefully handle certificate errors when trying to connect via HTTPS.
      '' Trust all certificates
      'ServicePointManager.ServerCertificateValidationCallback = (Function(sender, certificate, chain, sslPolicyErrors) True)

      '' Trust sender
      'ServicePointManager.ServerCertificateValidationCallback = (Function(sender, cert, chain, errors) cert.Subject.Contains(Me.ServerName))

      ' Validate cert by calling a function
      ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf ValidateRemoteCertificate)
      '' Back to normal...

      Dim lstrWSDL As String = "Empty"
      ApplicationLogging.WriteLogEntry(String.Format("Attempting to download WSDL: {0}", lstrURL),
                                 MethodBase.GetCurrentMethod, TraceEventType.Information, 13624)
      lstrWSDL = lobjWebClient.DownloadString(lstrURL)
      If lstrWSDL <> "Empty" Then
        'If (lstrWSDL IsNot Nothing) AndAlso (lstrWSDL.Length > 0) Then
        ' We got it
        ApplicationLogging.WriteLogEntry("Successfully downloaded WSDL",
                                         MethodBase.GetCurrentMethod, TraceEventType.Information, 13625)
        Return True
      Else
        ' We did not get it.
        lpErrorMessage = String.Format("{0} returned ({1})", lstrURL, lstrWSDL)
        ApplicationLogging.WriteLogEntry(lpErrorMessage, MethodBase.GetCurrentMethod, TraceEventType.Error, 13633)
        Return False
      End If
      'My.Computer.Network.DownloadFile(lstrURL, lstrWSDLPath)
      'My.Computer.FileSystem.DeleteFile(lstrWSDLPath)
      'Return True
    Catch WebEx As WebException
      If WebEx.Message.Contains("Unable to connect to the remote server") Then
        lpErrorMessage = String.Format("Unable to connect to the content engine server '{0}' using port '{1}'.  Make sure that the server name and port are correct and then test using the URL '{2}'.",
        Me.ServerName, Me.PortNumber, lstrURL)
        ApplicationLogging.WriteLogEntry(lpErrorMessage, MethodBase.GetCurrentMethod, TraceEventType.Error, 13634)
        Throw New RepositoryNotConnectedException(String.Format(
          "Unable to connect to the content engine server '{0}' using port '{1}'.  Make sure that the server name and port are correct and then test using the URL '{2}'.",
        Me.ServerName, Me.PortNumber, lstrURL), WebEx)
      Else
        ApplicationLogging.LogException(WebEx, MethodBase.GetCurrentMethod)
        lpErrorMessage = Helper.FormatCallStack(WebEx)
        ' Re-throw the exception to the caller
        Throw
      End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      lpErrorMessage = Helper.FormatCallStack(ex)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Shared Function ValidateRemoteCertificate(ByVal sender As Object, ByVal cert As X509Certificate, ByVal chain As X509Chain, ByVal policyErrors As SslPolicyErrors) As Boolean
    ' callback used to validate the certificate in an SSL conversation
    Try
      Dim lblnResult As Boolean = False

      Dim lstrExpirationDate As String = cert.GetExpirationDateString()
      Dim ldatExpirationDate As DateTime

      If DateTime.TryParse(lstrExpirationDate, ldatExpirationDate) Then
        If DateTime.Now > ldatExpirationDate Then
          Throw New InvalidOperationException("The certificate is expired")
        Else
          lblnResult = True
        End If
      Else
        Throw New InvalidOperationException("Certificate expiration date not valid")
      End If

      Return lblnResult

      Return lblnResult
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Sub InitializeObjectStore()

    Try
      mobjObjectStore = Factory.ObjectStore.FetchInstance(mobjDomain, ObjectStoreName, Nothing)

    Catch runtimeException As EngineRuntimeException
      ApplicationLogging.LogException(runtimeException, MethodBase.GetCurrentMethod)
      Throw _
        New InvalidOperationException(
          String.Format(
            "Invalid object store name: The object store name '{0}' is invalid.  Please supply a valid P8 object store name for the connection.",
            ObjectStoreName), runtimeException)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      Throw _
        New RepositoryNotConnectedException(String.Format("Unable to initialize object store '{0}'", ObjectStoreName),
                                            ex)
    End Try
  End Sub

#End Region

#Region ".NET API Implementation"

  Private Function GetDomain() As IDomain

    Try

      Dim lobjConnection As IConnection = Factory.Connection.GetConnection(URL)
      Dim lobjDomain As IDomain = Factory.Domain.GetInstance(lobjConnection, Nothing)

      Return lobjDomain

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region


  Private Function GetClassificationProperty(lpName As String) As ClassificationProperty

    Try

      'Dim lobjPropertyDefinition As Object = Nothing

      Dim lobjPropertyTemplate As IPropertyTemplate = GetPropertyTemplate(lpName)
      If (lobjPropertyTemplate IsNot Nothing) Then
        Return GetClassificationProperty(lobjPropertyTemplate)
      Else
        'lobjPropertyDefinition = ObjectStore.GetObject("PropertyDefinition", lpName)


        For Each lobjRootClass As IClassDefinition In ObjectStore.RootClassDefinitions

          If lobjRootClass.Name = "Document" Then

            For Each lobjPropertyDefinition As IPropertyDefinition In lobjRootClass.PropertyDefinitions

              If lobjPropertyDefinition.SymbolicName = lpName Then
                Return GetClassificationProperty(lobjPropertyDefinition)
              End If

            Next

            Exit For
          End If

        Next

        ' If we did not find it in the document properties, check the folder properties
        For Each lobjRootClass As IClassDefinition In ObjectStore.RootClassDefinitions

          If lobjRootClass.Name = "Folder" Then

            For Each lobjPropertyDefinition As IPropertyDefinition In lobjRootClass.PropertyDefinitions

              If lobjPropertyDefinition.SymbolicName = lpName Then
                Return GetClassificationProperty(lobjPropertyDefinition)
              End If

            Next

            Exit For
          End If

        Next


        'If lobjPropertyDefinition IsNot Nothing Then
        '  Return GetClassificationProperty(lobjPropertyDefinition)
        'End If

      End If

      Throw New PropertyDoesNotExistException("Unable to find property " + lpName, lpName)


    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function GetDocumentClass(lpClassName As String) As DocumentClass

    Dim lobjDocumentClass As DocumentClass = Nothing

    Try

      Dim lobjClassDef As IClassDefinition = ObjectStore.FetchObject("ClassDefinition", lpClassName, Nothing)
      If (lobjClassDef IsNot Nothing) Then
        lobjDocumentClass = GetObjectClass(lobjClassDef)
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

    Return lobjDocumentClass
  End Function


  Private Function GetDocumentClasses() As DocumentClasses

    Dim lobjObjectClasses As New RepositoryObjectClasses
    Dim lobjIDocumentClasses As IClassDefinitionSet
    Dim lstrClassName As String = String.Empty
    Dim lobjRootDocumentClassDefinition As IClassDefinition = Nothing

    Try

      If IsInitialized Then

        If ObjectStore Is Nothing Then
          InitializeObjectStore()
        End If

        lobjIDocumentClasses = ObjectStore.RootClassDefinitions _
        'Library.FilterClassDescriptions(IDMObjects.idmObjectType.idmObjTypeDocument,
        ' Loop through looking for Document Class
        'Dim lobjElement As IClassDefinition = Nothing
        For Each lobjElement As IClassDefinition In lobjIDocumentClasses
          'lobjElement = lobjEnumerator.Current
          lstrClassName = lobjElement.SymbolicName
          'Console.WriteLine("ClassName:" & lstrClassName)

          If String.Compare(lstrClassName, "Document", True) = 0 Then
            lobjRootDocumentClassDefinition = lobjElement
            Exit For
          End If

        Next

        'If lobjElement Is Nothing Then Return Nothing

        'Dim lobjRootDocumentClassDefinition As IClassDefinition = lobjElement

        'lobjDocumentClasses = GetDocumentClasses(lobjRootDocumentClassDefinition)
        lobjObjectClasses = GetObjectClasses(lobjRootDocumentClassDefinition)

        'For Each lobjIDocumentClass As IClassDefinition In lobjIDocumentClasses
        '  lobjDocumentClasses.Add(GetDocumentClass(lobjIDMDocumentClass))
        'Next

      End If

      'Return lobjDocumentClasses
      Return New DocumentClasses(lobjObjectClasses)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function GetCustomObjectClasses() As ObjectClasses

    Dim lobjObjectClasses As New RepositoryObjectClasses
    Dim lobjICustomObjectClasses As IClassDefinitionSet
    Dim lstrClassName As String = String.Empty
    Dim lobjRootCustomObjectClassDefinition As IClassDefinition = Nothing

    Try

      If IsInitialized Then

        If ObjectStore Is Nothing Then
          InitializeObjectStore()
        End If

        lobjICustomObjectClasses = ObjectStore.RootClassDefinitions _
        'Library.FilterClassDescriptions(IDMObjects.idmObjectType.idmObjTypeDocument,

        ' Loop through looking for CustomObject Class
        For Each lobjElement As IClassDefinition In lobjICustomObjectClasses
          lstrClassName = lobjElement.SymbolicName
          'Console.WriteLine("ClassName:" & lstrClassName)

          If String.Compare(lstrClassName, "CustomObject", True) = 0 Then
            lobjRootCustomObjectClassDefinition = lobjElement
            Exit For
          End If

        Next

        lobjObjectClasses = GetObjectClasses(lobjRootCustomObjectClassDefinition)

      End If

      Return New ObjectClasses(lobjObjectClasses)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function GetFolderClasses() As FolderClasses

    Dim lobjObjectClasses As New RepositoryObjectClasses
    Dim lobjIFolderClasses As IClassDefinitionSet
    Dim lstrClassName As String = String.Empty
    Dim lobjRootFolderClassDefinition As IClassDefinition = Nothing

    Try

      If IsInitialized Then

        If ObjectStore Is Nothing Then
          InitializeObjectStore()
        End If

        lobjIFolderClasses = ObjectStore.RootClassDefinitions _
        'Library.FilterClassDescriptions(IDMObjects.idmObjectType.idmObjTypeDocument,

        ' Loop through looking for Document Class
        For Each lobjElement As IClassDefinition In lobjIFolderClasses
          lstrClassName = lobjElement.SymbolicName
          'Console.WriteLine("ClassName:" & lstrClassName)

          If String.Compare(lstrClassName, "Folder", True) = 0 Then
            lobjRootFolderClassDefinition = lobjElement
            Exit For
          End If

        Next

        lobjObjectClasses = GetObjectClasses(lobjRootFolderClassDefinition)

      End If

      Return New FolderClasses(lobjObjectClasses)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  'Private Function GetDocumentClasses(ByVal lpParentClass As IClassDefinition, _
  '                                    Optional ByRef lpDocumentClassCollection As DocumentClasses = Nothing) As DocumentClasses

  '  Dim lobjDocumentClasses As DocumentClasses

  '  Try

  '    If lpDocumentClassCollection Is Nothing Then
  '      lobjDocumentClasses = New DocumentClasses

  '    Else
  '      lobjDocumentClasses = lpDocumentClassCollection
  '    End If

  '    ' Add the parent first
  '    ' --- Breaks here
  '    lobjDocumentClasses.Add(GetDocumentClass(lpParentClass))

  '    ' Get all the children
  '    For Each lobjIClassDefinition As IClassDefinition In lpParentClass.ImmediateSubclassDefinitions
  '      lobjDocumentClasses = GetDocumentClasses(lobjIClassDefinition, lobjDocumentClasses)
  '    Next

  '    Return lobjDocumentClasses

  '  Catch ex As Exception
  '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '    ' Re-throw the exception to the caller
  '    Throw
  '  End Try

  'End Function

  Private Function GetObjectClasses(ByVal lpParentClass As IClassDefinition,
                                    Optional ByRef lpObjectClassCollection As RepositoryObjectClasses = Nothing) _
    As RepositoryObjectClasses

    Dim lobjObjectClasses As RepositoryObjectClasses

    Try

      If lpObjectClassCollection Is Nothing Then
        lobjObjectClasses = New RepositoryObjectClasses

      Else
        lobjObjectClasses = lpObjectClassCollection
      End If

      ' Add the parent first
      ' --- Breaks here
      lobjObjectClasses.Add(GetObjectClass(lpParentClass))

      ' Get all the children
      For Each lobjIClassDefinition As IClassDefinition In lpParentClass.ImmediateSubclassDefinitions
        lobjObjectClasses = GetObjectClasses(lobjIClassDefinition, lobjObjectClasses)
      Next

      Return lobjObjectClasses

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function GetPropertyTemplate(ByVal lpSymbolicName As String) As IPropertyTemplate

    Try
      'If ObjectStore.Properties.IsPropertyPresent(lpSymbolicName) = True Then

      ' <Modified by: Ernie at 10/5/2011-10:43:42 AM on machine: ERNIE-M4400>
      ' PropertyChangedEventArgs(from)First to FirstOrDefault to resolve 'Sequence has no elements' error 
      ' Related to FogBugz case 222 for TXDOT / Bryan Phillips

      'Dim lobjIPropertyTemplate As IPropertyTemplate = (From pt In ObjectStore.PropertyTemplates Where pt.SymbolicName = lpSymbolicName Select pt).First
      Dim lobjIPropertyTemplate As IPropertyTemplate =
            (From pt In ObjectStore.PropertyTemplates Where pt.SymbolicName = lpSymbolicName Select pt).FirstOrDefault

      ' </Modified by: Ernie at 10/5/2011-10:43:42 AM on machine: ERNIE-M4400>

      Return lobjIPropertyTemplate

      'lobjIPropertyTemplate.SymbolicName
      'For Each lobjIPropertyTemplate As IPropertyTemplate In ObjectStore.PropertyTemplates
      '  If lobjIPropertyTemplate.SymbolicName.ToLower = lpSymbolicName.ToLower Then
      '    Return lobjIPropertyTemplate
      '  End If
      'Next

      'Return Nothing

      'Else
      'Return Nothing
      'End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod, 61393, lpSymbolicName)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  'Private Function GetDocumentClass(ByVal lpIClassDefinition As IClassDefinition) As DocumentClass

  '  Dim lobjECMDocumentClass As DocumentClass
  '  Dim lobjECMProperties As New ClassificationProperties
  '  Dim lobjECMProperty As ClassificationProperty

  '  Try

  '    If lpIClassDefinition Is Nothing Then
  '      Return Nothing
  '    End If

  '    Debug.WriteLine("")
  '    Debug.WriteLine("Class SymbolicName: " & lpIClassDefinition.SymbolicName)

  '    ' Let's get the properties
  '    For Each lobjIPropertyDescription As IPropertyDefinition In lpIClassDefinition.PropertyDefinitions
  '      'For Each lobjIDMPropertyDescription As IDMObjects.PropertyDescription In propDescs
  '      lobjECMProperty = GetClassificationProperty(lobjIPropertyDescription)
  '      'Debug.WriteLine("  Property SymbolicName: " & lobjIPropertyDescription.SymbolicName)

  '      If Not lobjECMProperty Is Nothing Then
  '        lobjECMProperties.Add(lobjECMProperty)

  '        If ContentProperties.Contains(lobjECMProperty.SystemName) = False Then
  '          ContentProperties.Add(lobjECMProperty)
  '        End If

  '      End If

  '    Next

  '    '    Debug.WriteLine("")

  '    ''lobjECMDocumentClass.Properties = lobjECMProperties

  '    ' Let's create the abstract
  '    lobjECMDocumentClass = New DocumentClass(lpIClassDefinition.SymbolicName, lobjECMProperties, lpIClassDefinition.Id.ToString, lpIClassDefinition.DescriptiveText)

  '    Return lobjECMDocumentClass

  '  Catch ex As Exception
  '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '    ' Re-throw the exception to the caller
  '    Throw
  '  End Try

  'End Function

  Private Function GetObjectClass(ByVal lpIClassDefinition As IClassDefinition) As RepositoryObjectClass

    Dim lobjECMDocumentClass As RepositoryObjectClass = Nothing
    Dim lobjECMProperties As New ClassificationProperties
    Dim lobjECMProperty As ClassificationProperty

    Try

      If lpIClassDefinition Is Nothing Then
        Return Nothing
      End If

      Debug.WriteLine("")
      Debug.WriteLine("Class SymbolicName: " & lpIClassDefinition.SymbolicName)

      ' Let's get the properties
      For Each lobjIPropertyDescription As IPropertyDefinition In lpIClassDefinition.PropertyDefinitions
        'For Each lobjIDMPropertyDescription As IDMObjects.PropertyDescription In propDescs
        lobjECMProperty = GetClassificationProperty(lobjIPropertyDescription)
        'Debug.WriteLine("  Property SymbolicName: " & lobjIPropertyDescription.SymbolicName)

        If lobjECMProperty IsNot Nothing Then
          lobjECMProperties.Add(lobjECMProperty)

          If ContentProperties.Contains(lobjECMProperty.SystemName) = False Then
            ContentProperties.Add(lobjECMProperty)
          End If

        End If

      Next

      '    Debug.WriteLine("")

      ''lobjECMDocumentClass.Properties = lobjECMProperties

      ' Let's create the abstract
      If lpIClassDefinition.TableDefinition.Name = "DocVersion" Then
        lobjECMDocumentClass = New DocumentClass(lpIClassDefinition.SymbolicName, lobjECMProperties,
                                                 lpIClassDefinition.Id.ToString, lpIClassDefinition.DescriptiveText)
      ElseIf lpIClassDefinition.TableDefinition.Name = "Container" Then
        lobjECMDocumentClass = New FolderClass(lpIClassDefinition.SymbolicName, lobjECMProperties,
                                               lpIClassDefinition.Id.ToString, lpIClassDefinition.DescriptiveText)
      ElseIf lpIClassDefinition.TableDefinition.Name = "Generic" Then
        lobjECMDocumentClass = New ObjectClass(lpIClassDefinition.SymbolicName, lobjECMProperties,
                                               lpIClassDefinition.Id.ToString, lpIClassDefinition.DescriptiveText)
      End If

      Return lobjECMDocumentClass

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function GetClassificationProperty(ByVal lpIPropertyDefinition As Object) As ClassificationProperty

    Dim lobjProperty As ClassificationProperty
    Dim lobjECMPropertyType As PropertyType
    Dim lenuCardinality As DCore.Cardinality
    Dim lstrErrorMessage As String = String.Empty

    Try
      ' Get the data type
      lobjECMPropertyType = GetECMPropertyType(lpIPropertyDefinition.DataType)

      ' Get the cardinality
      If lpIPropertyDefinition.Cardinality = Constants.Cardinality.SINGLE Then
        lenuCardinality = DCore.Cardinality.ecmSingleValued

      Else
        lenuCardinality = DCore.Cardinality.ecmMultiValued
      End If

      ' Create the abstract property object
      'lobjProperty = New ClassificationProperty(lobjECMPropertyType, _
      '  lpIPropertyDefinition.Name.ToString, lenuCardinality)
      lobjProperty = ClassificationPropertyFactory.Create(lobjECMPropertyType, lpIPropertyDefinition.Name.ToString,
                                                          lenuCardinality)
      ' Set the ID
      lobjProperty.SetID(lpIPropertyDefinition.Id.ToString)

      ' Set the PackedName
      lobjProperty.SetPackedName(lpIPropertyDefinition.SymbolicName)

      '' For debugging
      '' See what properties we have available
      'For Each lobjProDefProperty As Object In lpIPropertyDefinition.Properties
      '  Debug.Print(lobjProDefProperty)
      'Next

      '  Label the property as to whether or not it is a required property.
      'Try
      '  lobjProperty.IsRequired = lpIPropertyDefinition.IsValueRequired
      'Catch ex As Exception
      '  ' We were unable to get this value
      '  ' Default to false
      '  lobjProperty.IsRequired = False
      'End Try

      '  Label the property as to whether it is a system property or a custom property.
      'Try
      '  lobjProperty.IsSystemProperty = lpIPropertyDefinition.IsSystemOwned
      'Catch ex As Exception
      '  ' We were unable to get this value
      '  ' Default to false
      '  lobjProperty.IsSystemProperty = False
      'End Try

      If lpIPropertyDefinition.ChoiceList IsNot Nothing Then

        ' TODO: Add the choicelist
        ' Beep()
        If TypeOf lpIPropertyDefinition.ChoiceList Is Admin.IChoiceList Then
          lobjProperty.ChoiceList = GetChoiceList(lpIPropertyDefinition.ChoiceList, lstrErrorMessage)
        End If

      End If

      Dim lobjP8PropertyType As Type = lpIPropertyDefinition.GetType

      'Dim lobjPropertyInfo As Reflection.PropertyInfo() = lobjP8PropertyType.GetProperties
      'For Each pi As Reflection.pr lobjPropertyInfo
      Dim lobjPropertySystemOwnedInfo As PropertyInfo = lobjP8PropertyType.GetProperty("IsSystemOwned")

      If lobjPropertySystemOwnedInfo IsNot Nothing AndAlso lobjPropertySystemOwnedInfo.CanRead Then

        If lpIPropertyDefinition.IsSystemOwned IsNot Nothing Then
          lobjProperty.IsSystemProperty = lpIPropertyDefinition.IsSystemOwned
        End If

      End If

      With lobjProperty

        If lpIPropertyDefinition.IsHidden IsNot Nothing Then
          .IsHidden = lpIPropertyDefinition.IsHidden
        End If

        If lpIPropertyDefinition.IsValueRequired IsNot Nothing Then
          .IsRequired = lpIPropertyDefinition.IsValueRequired
        End If

        .Settability = lpIPropertyDefinition.Settability
        .SystemName = lpIPropertyDefinition.SymbolicName
      End With

      Select Case lobjECMPropertyType

        Case PropertyType.ecmString
          'If lpIPropertyDefinition.MaximumLengthString IsNot Nothing Then
          With DirectCast(lobjProperty, ClassificationStringProperty)
            .MaxLength = lpIPropertyDefinition.MaximumLengthString

            If lpIPropertyDefinition.PropertyDefaultString IsNot Nothing Then
              .DefaultValue = lpIPropertyDefinition.PropertyDefaultString
            End If

          End With

        Case PropertyType.ecmDate
          With DirectCast(lobjProperty, ClassificationDateTimeProperty)
            .DefaultValue = lpIPropertyDefinition.PropertyDefaultDateTime
            .MinValue = lpIPropertyDefinition.PropertyMinimumDateTime
            .MaxValue = lpIPropertyDefinition.PropertyMaximumDateTime
          End With

        Case PropertyType.ecmBoolean
          With DirectCast(lobjProperty, ClassificationBooleanProperty)
            .DefaultValue = lpIPropertyDefinition.PropertyDefaultBoolean
          End With

        Case PropertyType.ecmDouble
          With DirectCast(lobjProperty, ClassificationDoubleProperty)
            .DefaultValue = lpIPropertyDefinition.PropertyDefaultFloat64
            .MinValue = lpIPropertyDefinition.PropertyMinimumFloat64
            .MaxValue = lpIPropertyDefinition.PropertyMaximumFloat64
          End With

        Case PropertyType.ecmLong
          With DirectCast(lobjProperty, ClassificationLongProperty)

            If lpIPropertyDefinition.PropertyDefaultInteger32 IsNot Nothing Then
              .DefaultValue = New Nullable(Of Long)(lpIPropertyDefinition.PropertyDefaultInteger32)
            End If

            If lpIPropertyDefinition.PropertyMinimumInteger32 IsNot Nothing Then
              .MinValue = New Nullable(Of Long)(lpIPropertyDefinition.PropertyMinimumInteger32)
            End If

            If lpIPropertyDefinition.PropertyMaximumInteger32 IsNot Nothing Then
              .MaxValue = New Nullable(Of Long)(lpIPropertyDefinition.PropertyMaximumInteger32)
            End If

          End With

        Case PropertyType.ecmGuid
          With DirectCast(lobjProperty, ClassificationGuidProperty)
            .DefaultValue = lpIPropertyDefinition.PropertyDefaultId
          End With
      End Select

      Return lobjProperty

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod, 0, lpIPropertyDefinition.Name)
      '  Re-throw the exeption to the caller
      Throw
    End Try
  End Function

  Public Function GetChoiceList(ByVal lpP8ChoiceListId As String,
                                Optional ByRef lpErrorMessage As String = "") As ChoiceList
    Try
      Dim lobjCtsChoiceList As New ChoiceList
      Dim lobjP8ChoiceList As Admin.IChoiceList = GetP8ChoiceList(lpP8ChoiceListId)
      If lobjP8ChoiceList IsNot Nothing Then
        lobjCtsChoiceList = GetChoiceList(lobjP8ChoiceList, lpErrorMessage)
      Else
        lobjCtsChoiceList = Nothing
      End If
      Return lobjCtsChoiceList
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function GetChoiceList(ByVal lpP8ChoiceList As Admin.IChoiceList,
                                Optional ByRef lpErrorMessage As String = "") As ChoiceList

    Try

      Dim lobjCtsChoiceList As New ChoiceList

      With lobjCtsChoiceList
        .Name = lpP8ChoiceList.Name
        .DisplayName = lpP8ChoiceList.DisplayName
        .DescriptiveText = lpP8ChoiceList.DescriptiveText
        .Id = lpP8ChoiceList.Id.ToString
        .ChoiceValues = ChoiceValues(lpP8ChoiceList.ChoiceValues)
        'For Each lobjValue As Object In lpP8ChoiceList.ChoiceValues
        '  Beep()

        'Next
      End With

      Return lobjCtsChoiceList

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function ChoiceValues(ByVal lpCEChoiceValues As Object,
                                Optional ByVal lpParent As ChoiceValues = Nothing) As ChoiceValues

    Try

      Dim lobjCTSChoiceValues As ChoiceValues

      If lpParent Is Nothing Then
        lobjCTSChoiceValues = New ChoiceValues

      Else
        lobjCTSChoiceValues = lpParent
      End If

      'For lintValueCounter As Integer = 0 To lpCEChoiceValues.Value.Length - 1
      '  lobjChoiceValue = lpCEChoiceValues.Value(lintValueCounter)
      '  lobjCTSChoiceValues.Add(ChoiceItem(lobjChoiceValue))
      'Next

      For Each lobjChoiceValue As IChoice In lpCEChoiceValues
        lobjCTSChoiceValues.Add(ChoiceItem(lobjChoiceValue))
      Next

      Return lobjCTSChoiceValues

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function ChoiceItem(ByVal lpChoice As IChoice,
                              Optional ByVal lpErrorMessage As String = "") As ChoiceItem

    Try

      Dim lobjChoiceItem As ChoiceItem = Nothing

      ' Beep()

      Select Case lpChoice.ChoiceType

        Case Constants.ChoiceType.STRING
          ' This is a ChoiceValue of type String
          lobjChoiceItem = New ChoiceValue(lpChoice.ChoiceStringValue)

        Case Constants.ChoiceType.INTEGER
          ' This is a ChoiceValue of type Integer
          lobjChoiceItem = New ChoiceValue(lpChoice.ChoiceIntegerValue)

        Case Constants.ChoiceType.MIDNODE_STRING, Constants.ChoiceType.MIDNODE_INTEGER
          ' This is a ChoiceGroup
          lobjChoiceItem = New ChoiceGroup(lpChoice.Name)
          DirectCast(lobjChoiceItem, ChoiceGroup).ChoiceValues = ChoiceValues(lpChoice.ChoiceValues)

      End Select

      With lobjChoiceItem
        .Name = lpChoice.Name
        .DisplayName = lpChoice.DisplayName
        .Id = lpChoice.Id.ToString
      End With

      Return lobjChoiceItem

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function GetECMPropertyType(ByVal lpTypeID As TypeID) As PropertyType

    Try

      Select Case lpTypeID

        Case TypeID.BOOLEAN
          Return PropertyType.ecmBoolean

        Case TypeID.BINARY
          Return PropertyType.ecmBinary

        Case TypeID.STRING
          Return PropertyType.ecmString

        Case TypeID.DOUBLE
          Return PropertyType.ecmDouble

        Case TypeID.DATE
          Return PropertyType.ecmDate

        Case TypeID.GUID
          Return PropertyType.ecmGuid

        Case TypeID.LONG
          Return PropertyType.ecmLong

        Case TypeID.OBJECT
          Return PropertyType.ecmObject

        Case Else
          Return PropertyType.ecmObject

      End Select

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Protected Friend Overridable Function GetCtsPermissions(lpSecuredObject As IContainable) As IPermissions
    Try

      Dim lobjCtsPermissions As New Permissions
      Dim lobjCtsPermission As ItemPermission = Nothing

      For Each lobjP8Permission As IAccessPermission In lpSecuredObject.Permissions
        lobjCtsPermissions.Add(CreateCtsPermission(lobjP8Permission))
      Next

      Return lobjCtsPermissions

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Protected Friend Overridable Function CreateP8PermissionList(lpCtsPermissions As IPermissions) _
    As IAccessPermissionList
    Try
      Dim lobjP8PermissionList As IAccessPermissionList = Factory.AccessPermission.CreateList
      Dim lobjP8Permission As IAccessPermission = Nothing

      For Each lobjCtsPermission As DSecurity.IPermission In lpCtsPermissions
        lobjP8Permission = CreateP8Permission(lobjCtsPermission)
        If lobjP8Permission.AccessMask.HasValue Then
          lobjP8PermissionList.Add(lobjP8Permission)
        Else
          Beep()
        End If
      Next

      Return lobjP8PermissionList

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function CreateP8Permission(lpCtsPermission As DSecurity.IPermission) As IAccessPermission
    Try
      Dim lobjP8Permission As IAccessPermission = Factory.AccessPermission.CreateInstance

      If lpCtsPermission.Access.Value.HasValue Then
        lobjP8Permission.AccessMask = lpCtsPermission.Access.Value
      Else
        lobjP8Permission.AccessMask = CreateAccessMask(lpCtsPermission.Access)
      End If

      lobjP8Permission.GranteeName = lpCtsPermission.PrincipalName
      'lobjP8Permission.GranteeType = lpCtsPermission.PrincipalType
      'lobjP8Permission.PermissionSource = lpCtsPermission.PermissionSource
      lobjP8Permission.AccessType = lpCtsPermission.AccessType

      Return lobjP8Permission

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function GetRepositoryAccessMask(lpCtsPermission As DSecurity.IPermission) As Nullable(Of Integer)
    Try
      Return CreateAccessMask(lpCtsPermission.Access)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function CreateAccessMask(lpCtsAccessMask As IAccessMask) As Nullable(Of Integer)
    Try
      Dim lintAccessMask As Nullable(Of Integer) = Nothing

      If TypeOf lpCtsAccessMask Is IAccessLevel Then
        Select Case CType(lpCtsAccessMask, IAccessLevel).Level
          Case PermissionLevel.FullControl
            lintAccessMask = Constants.AccessLevel.FULL_CONTROL

          Case PermissionLevel.ModifyProperties
            lintAccessMask = Constants.AccessLevel.VIEW + Constants.AccessRight.MODIFY_OBJECTS +
                             Constants.AccessRight.LINK + Constants.AccessRight.UNLINK +
                             Constants.AccessRight.CREATE_INSTANCE +
                             Constants.AccessRight.CREATE_CHILD + Constants.AccessRight.READ_ACL +
                             Constants.AccessRight.MINOR_VERSION +
                             Constants.AccessRight.VIEW_CONTENT + Constants.AccessRight.CHANGE_STATE +
                             Constants.AccessRight.PUBLISH

          Case PermissionLevel.AddToFolder
            lintAccessMask = Constants.AccessLevel.VIEW + Constants.AccessRight.LINK +
                             Constants.AccessRight.UNLINK + Constants.AccessRight.READ_ACL

          Case PermissionLevel.ViewProperties
            lintAccessMask = Constants.AccessLevel.VIEW + Constants.AccessRight.READ_ACL

          Case PermissionLevel.Custom
            lintAccessMask = CreateAccessMask(lpCtsAccessMask.PermissionList)
          Case Else
            lintAccessMask = CreateAccessMask(lpCtsAccessMask.PermissionList)
        End Select
      Else
        lintAccessMask = CreateAccessMask(lpCtsAccessMask.PermissionList)
      End If

      Return lintAccessMask

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function CreateAccessMask(lpPermissionList As DSecurity.IPermissionList) As Nullable(Of Integer)
    Try

      Dim lintAccessMask As Nullable(Of Integer) = Nothing

      If lpPermissionList.Count > 0 Then
        ' Initialize the value so that addition operations will work as expected.
        lintAccessMask = 0
      End If

      ' <Added by: Ernie at: 7/11/2012-10:08:44 AM on machine: ERNIE-M4400>
      ' Fix for OpenText PVCS  19257
      ' The OpenText guys like to call them all to mirror the check boxes in WorkplaceXT.

      If lpPermissionList.Contains(PermissionRight.PromoteVersionRollup) AndAlso
         lpPermissionList.Contains(PermissionRight.ModifyContentRollup) AndAlso
         lpPermissionList.Contains(PermissionRight.ModifyDocumentPropertiesRollup) AndAlso
         lpPermissionList.Contains(PermissionRight.ViewContentRollup) AndAlso
         lpPermissionList.Contains(PermissionRight.ViewDocumentProperties) AndAlso
         lpPermissionList.Contains(PermissionRight.PublishRollup) Then
        lintAccessMask = ACCESS_RIGHT_ALL_BUT_OWNER_CONTROL_DOCUMENT_ROLLUP
        Return lintAccessMask
      End If

      If lpPermissionList.Contains(PermissionRight.PromoteVersionRollup) AndAlso
         lpPermissionList.Contains(PermissionRight.ModifyContentRollup) AndAlso
         lpPermissionList.Contains(PermissionRight.ModifyDocumentPropertiesRollup) AndAlso
         lpPermissionList.Contains(PermissionRight.ViewContentRollup) AndAlso
         lpPermissionList.Contains(PermissionRight.ViewDocumentProperties) Then
        lintAccessMask = ACCESS_RIGHT_ALL_BUT_OWNER_CONTROL_AND_PUBLISH_DOCUMENT_ROLLUP
        Return lintAccessMask
      End If

      If lpPermissionList.Contains(PermissionRight.ModifyContentRollup) AndAlso
         lpPermissionList.Contains(PermissionRight.ModifyDocumentPropertiesRollup) AndAlso
         lpPermissionList.Contains(PermissionRight.ViewContentRollup) AndAlso
         lpPermissionList.Contains(PermissionRight.ViewDocumentProperties) AndAlso
         lpPermissionList.Contains(PermissionRight.PublishRollup) Then
        lintAccessMask = ACCESS_RIGHT_MODIFY_CONTENT_DOCUMENT_ROLLUP
        Return lintAccessMask
      End If

      If lpPermissionList.Contains(PermissionRight.ModifyDocumentPropertiesRollup) AndAlso
         lpPermissionList.Contains(PermissionRight.ViewContentRollup) AndAlso
         lpPermissionList.Contains(PermissionRight.ViewDocumentProperties) AndAlso
         lpPermissionList.Contains(PermissionRight.PublishRollup) Then
        lintAccessMask = ACCESS_RIGHT_MODIFY_CONTENT_DOCUMENT_ROLLUP - Constants.AccessRight.MINOR_VERSION
        Return lintAccessMask
      End If

      'AR.PermissionList.Add(PermissionRight.ModifyContentRollup);
      'AR.PermissionList.Add(PermissionRight.ViewContentRollup);
      'AR.PermissionList.Add(PermissionRight.ModifyDocumentPropertiesRollup);
      'AR.PermissionList.Add(PermissionRight.ViewDocumentProperties);

      If lpPermissionList.Contains(PermissionRight.ModifyContentRollup) AndAlso
         lpPermissionList.Contains(PermissionRight.ViewContentRollup) AndAlso
         lpPermissionList.Contains(PermissionRight.ModifyDocumentPropertiesRollup) AndAlso
         lpPermissionList.Contains(PermissionRight.ViewDocumentProperties) Then
        lintAccessMask = ACCESS_RIGHT_MODIFY_CONTENT_AND_PROPERTIES_ONLY_DOCUMENT_ROLLUP
        Return lintAccessMask
      End If

      If lpPermissionList.Contains(PermissionRight.ViewContentRollup) AndAlso
         lpPermissionList.Contains(PermissionRight.ModifyDocumentPropertiesRollup) AndAlso
         lpPermissionList.Contains(PermissionRight.ViewDocumentProperties) Then
        lintAccessMask = ACCESS_RIGHT_MODIFY_PROPERTIES_AND_VIEW_CONTENT_ONLY_DOCUMENT_ROLLUP
        Return lintAccessMask
      End If

      If lpPermissionList.Contains(PermissionRight.ViewContentRollup) AndAlso
         lpPermissionList.Contains(PermissionRight.ViewDocumentProperties) Then
        lintAccessMask = ACCESS_RIGHT_VIEW_CONTENT_AND_PROPERTIES_ONLY_DOCUMENT_ROLLUP
        Return lintAccessMask
      End If

      ' </Added by: Ernie at: 7/11/2012-10:08:44 AM on machine: ERNIE-M4400>

      ' <Added by: Ernie at: 7/11/2012-9:54:37 AM on machine: ERNIE-M4400>
      ' Fix for OpenText PVCS  19260
      ' We don't want the caller to set both of these properties together.  
      ' The OpenText guys like to call them both to mirror the check boxes in WorkplaceXT.
      If lpPermissionList.Contains(PermissionRight.ModifyFolderPropertiesRollup) AndAlso
         lpPermissionList.Contains(PermissionRight.ViewFolderProperties) Then
        lpPermissionList.Remove(PermissionRight.ViewFolderProperties)
      End If
      ' </Added by: Ernie at: 7/11/2012-9:54:37 AM on machine: ERNIE-M4400>

      For Each lenuRight As PermissionRight In lpPermissionList

        Select Case lenuRight
          Case PermissionRight.ViewDocumentProperties
            lintAccessMask += Constants.AccessLevel.VIEW - Constants.AccessRight.VIEW_CONTENT

          Case PermissionRight.ViewFolderProperties
            lintAccessMask += Constants.AccessLevel.READ

          Case PermissionRight.ModifyDocumentProperties
            lintAccessMask += Constants.AccessRight.MODIFY_OBJECTS

          Case PermissionRight.ModifyDocumentPropertiesRollup
            lintAccessMask += Constants.AccessRight.WRITE + Constants.AccessRight.LINK +
                              Constants.AccessRight.CREATE_INSTANCE + Constants.AccessRight.CHANGE_STATE +
                              Constants.AccessLevel.VIEW

          Case PermissionRight.ModifyFolderPropertiesRollup
            lintAccessMask += ACCESS_LEVEL_MODIFY_PROPERTIES

            'Case Cts.Security.PermissionLevel.AllButOwnerControlDocumentRollup
            '  lintAccessMask = ACCESS_RIGHT_ALL_BUT_OWNER_CONTROL_DOCUMENT_ROLLUP

          Case PermissionRight.ViewContent
            lintAccessMask += Constants.AccessRight.VIEW_CONTENT

          Case PermissionRight.ViewContentRollup
            lintAccessMask += Constants.AccessLevel.VIEW

          Case PermissionRight.ModifyContentRollup
            lintAccessMask = Constants.AccessLevel.WRITE_DOCUMENT + Constants.AccessRight.UNLINK

          Case PermissionRight.PromoteVersionRollup
            lintAccessMask = Constants.AccessLevel.WRITE_DOCUMENT + Constants.AccessRight.UNLINK +
                             Constants.AccessRight.MAJOR_VERSION

          Case PermissionRight.Link_Annotate
            lintAccessMask += Constants.AccessRight.LINK

          Case PermissionRight.Publish
            lintAccessMask += Constants.AccessRight.PUBLISH

          Case PermissionRight.PublishRollup
            lintAccessMask += Constants.AccessRight.WRITE + Constants.AccessRight.LINK +
                              Constants.AccessRight.CREATE_INSTANCE + Constants.AccessRight.CHANGE_STATE +
                              Constants.AccessLevel.VIEW + Constants.AccessRight.PUBLISH

          Case PermissionRight.CreateInstance
            lintAccessMask += Constants.AccessRight.CREATE_INSTANCE

          Case PermissionRight.ChangeState
            lintAccessMask += Constants.AccessRight.CHANGE_STATE

          Case PermissionRight.MinorVersioning
            lintAccessMask += Constants.AccessRight.MINOR_VERSION

          Case PermissionRight.MajorVersioning
            lintAccessMask += Constants.AccessRight.MAJOR_VERSION

          Case PermissionRight.Delete
            lintAccessMask += Constants.AccessRight.DELETE

          Case PermissionRight.ReadPermissions
            lintAccessMask += Constants.AccessRight.READ_ACL

          Case PermissionRight.ModifyPermissions
            lintAccessMask += Constants.AccessRight.WRITE_ACL

          Case PermissionRight.ModifyOwner
            lintAccessMask += Constants.AccessRight.WRITE_OWNER

          Case PermissionRight.UnlinkDocument
            lintAccessMask += Constants.AccessRight.UNLINK

          Case PermissionRight.CreateSubfolder
            lintAccessMask += Constants.AccessRight.CREATE_CHILD

          Case PermissionRight.CreateSubfolderRollup
            lintAccessMask += Constants.AccessLevel.READ + Constants.AccessRight.CREATE_CHILD

          Case PermissionRight.ViewContent
            lintAccessMask += Constants.AccessLevel.VIEW

          Case PermissionRight.FileInFolder
            'lintAccessMask += AccessLevel.LINK_FOLDER
            lintAccessMask += Constants.AccessRight.LINK + Constants.AccessRight.UNLINK

          Case PermissionRight.FullControlRollup
            lintAccessMask = Constants.AccessLevel.FULL_CONTROL

        End Select

      Next

      Return lintAccessMask

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function CreateCtsPermission(lpP8Permission As IAccessPermission) As DSecurity.IPermission
    Try
      Dim lobjCtsPermission As New ItemPermission
      Dim lobjCtsAccessRight As IAccessMask = If(AvailableRights.Item(CInt(lpP8Permission.AccessMask)), New DSecurity.AccessRight("Custom", lpP8Permission.AccessMask))

      lobjCtsPermission.Access = lobjCtsAccessRight
      lobjCtsPermission.PrincipalName = lpP8Permission.GranteeName
      lobjCtsPermission.Source = lpP8Permission.PermissionSource
      lobjCtsPermission.PrincipalType = lpP8Permission.GranteeType
      'lobjCtsPermission.Access=New 
      lobjCtsPermission.AccessType = lpP8Permission.AccessType

      Return lobjCtsPermission

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function GetAllAvailableRights() As IAccessRights
    Try

      Dim lobjAccessRights As IAccessRights = New AccessRights
      Dim lobjP8AccessRights As IDictionary(Of String, Integer) =
            Helper.EnumerationDictionary(GetType(Constants.AccessRight))
      Dim lobjP8AccessLevels As IDictionary(Of String, Integer) =
            Helper.EnumerationDictionary(GetType(Constants.AccessLevel))

      ' Add all the rights
      For Each lstrRightKey As String In lobjP8AccessRights.Keys
        lobjAccessRights.Add(New DSecurity.AccessRight(lstrRightKey, lobjP8AccessRights.Item(lstrRightKey)))
      Next

      ' Add all the levels
      For Each lstrLevelKey As String In lobjP8AccessLevels.Keys
        lobjAccessRights.Add(New DSecurity.AccessLevel(lstrLevelKey, lobjP8AccessLevels.Item(lstrLevelKey)))
      Next

      If lobjAccessRights.Item(998871) Is Nothing Then
        lobjAccessRights.Add(New DSecurity.AccessLevel("FULL_CONTROL_OTHER", 998871))
      End If

      ' Add a level for the Workplace level 'ModifyContent'
      If lobjAccessRights.Item(132563) Is Nothing Then
        lobjAccessRights.Add(New DSecurity.AccessLevel("MODIFY_CONTENT_WP", 132563))
      End If

      Return lobjAccessRights

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function GetAllBaseProperties() As ClassificationProperties

    Dim lobjProperties As New ClassificationProperties
    Dim lobjProperty As ClassificationProperty

    Try

      For Each lobjIPropertyTemplate As IPropertyTemplate In ObjectStore.PropertyTemplates
        lobjProperty = GetClassificationProperty(lobjIPropertyTemplate)

        If lobjProperties.Contains(lobjProperty.Name) = False Then
          Debug.WriteLine(lobjProperty.ID & ": " & lobjProperty.Name)
          lobjProperties.Add(lobjProperty) ', lobjProperty.ID)
        End If

      Next

      ' Id
      If lobjProperties.Contains("Id") = False Then

        Dim lobjIsCurrentVersionProp As ClassificationProperty =
              ClassificationPropertyFactory.Create(PropertyType.ecmGuid, "Id", DCore.Cardinality.ecmSingleValued)
        With lobjIsCurrentVersionProp
          .IsSystemProperty = True
          .IsRequired = True
          .Settability = ClassificationProperty.SettabilityEnum.READ_ONLY
        End With
        lobjProperties.Add(lobjIsCurrentVersionProp)
      End If

      ' Hard code an add for Rick
      ' IsCurrentVersion
      If lobjProperties.Contains("IsCurrentVersion") = False Then

        Dim lobjIsCurrentVersionProp As ClassificationProperty =
              ClassificationPropertyFactory.Create(PropertyType.ecmBoolean, "IsCurrentVersion",
                                                   DCore.Cardinality.ecmSingleValued)
        With lobjIsCurrentVersionProp
          .IsSystemProperty = True
          .IsRequired = True
          .Settability = ClassificationProperty.SettabilityEnum.READ_ONLY
          .DefaultValue = False
        End With
        lobjProperties.Add(lobjIsCurrentVersionProp)
      End If

      lobjProperties.Sort()

      Return lobjProperties

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function


  Private Function GetAllFolderProperties() As ClassificationProperties

    Dim lobjProperties As New ClassificationProperties
    Dim lobjProperty As ClassificationProperty
    Try

      lobjProperties = GetAllBaseProperties()

      ' Also get any additional property definitions that may be in the root 'Document' class
      For Each lobjRootClass As IClassDefinition In ObjectStore.RootClassDefinitions

        If lobjRootClass.Name = "Folder" Then

          For Each lobjPropertyDefinition As IPropertyDefinition In lobjRootClass.PropertyDefinitions
            lobjProperty = GetClassificationProperty(lobjPropertyDefinition)

            If lobjProperties.Contains(lobjProperty.Name) = False Then
              Debug.WriteLine(lobjProperty.ID & ": " & lobjProperty.Name)
              lobjProperties.Add(lobjProperty) ', lobjProperty.ID)
            End If

          Next

          Exit For
        End If

      Next

      lobjProperties.Sort()

      Return lobjProperties

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function GetAllCustomObjectProperties() As ClassificationProperties

    Dim lobjProperties As New ClassificationProperties
    Dim lobjProperty As ClassificationProperty

    Try

      lobjProperties = GetAllBaseProperties()

      ' Also get any additional property definitions that may be in the root 'Document' class
      For Each lobjRootClass As IClassDefinition In ObjectStore.RootClassDefinitions

        If lobjRootClass.Name = "CustomObject" Then

          For Each lobjPropertyDefinition As IPropertyDefinition In lobjRootClass.PropertyDefinitions
            lobjProperty = GetClassificationProperty(lobjPropertyDefinition)

            If lobjProperties.Contains(lobjProperty.Name) = False Then
              Debug.WriteLine(lobjProperty.ID & ": " & lobjProperty.Name)
              lobjProperties.Add(lobjProperty) ', lobjProperty.ID)
            End If

          Next

          Exit For
        End If

      Next

      lobjProperties.Sort()

      Return lobjProperties

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  ''' <summary>
  ''' AccessDocumentStream
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub AccessDocumentStreams(lpDocumentId As String,
                         ByRef lpProcessedMessage As String)

    Dim lobjIDocument As IDocument = Nothing
    Dim lobjContentStream As System.IO.Stream = Nothing
    Dim lobjContentElementList As IContentElementList = Nothing

    Try

      lobjIDocument = GetIDocument(lpDocumentId)
      lobjContentElementList = lobjIDocument.ContentElements
      For Each lobjContentElement As Object In lobjContentElementList

        If TypeOf lobjContentElement Is IContentTransfer Then
          ' Get the content stream
          'This will throw an exception if content is not found or inaccessible
          'EXAMPLE exception msg: "The file does not exist \\fn\content\Data_Redirects\QA\FS_Email_Archival_Encrypted\content\FN3\FN15\1CC89617-03CF-4A01-B60D-373174EDDC40{DF6331AC-69C8-46F7-B527-0B4E6A36B792}0."
          Try
            lobjContentStream = DirectCast(lobjContentElement, IContentTransfer).AccessContentStream()
          Catch runtimeEx As EngineRuntimeException
            Dim lobjNewEx As New FileDoesNotExistException(lpDocumentId, runtimeEx.Message, runtimeEx)
            ApplicationLogging.LogException(lobjNewEx, Reflection.MethodBase.GetCurrentMethod)
            ' Re-throw the exception to the caller
            Throw
          End Try
          lpProcessedMessage = String.Format("Successfully accessed content for doc id: '{0}'.", lpDocumentId)
        ElseIf TypeOf lobjContentElement Is IContentReference Then
          Dim lstrContentLocation As String = DirectCast(lobjContentElement, IContentReference).ContentLocation
          If String.IsNullOrEmpty(lstrContentLocation) Then
            Throw New InvalidPathException(String.Format("The content location is not set for item {0}", lpDocumentId), String.Empty)
          End If
          Try
            If Helper.IsUrlAvailable(lstrContentLocation) = False Then
              Throw New UrlNotAvailableException(lstrContentLocation,
                String.Format("Invalid content location for document {0}: {1}", lpDocumentId, lstrContentLocation))
            Else
              lpProcessedMessage = String.Format("Successfully accessed external content for doc id: '{0}' at {1}.", lpDocumentId, lstrContentLocation)
            End If
          Catch ServerUnavailableEx As ServerUnavailableException
            lpProcessedMessage = String.Format("Invalid content location for document {0}, the server is unavailable for url '{1}': {2}",
                                               lpDocumentId, lstrContentLocation, ServerUnavailableEx.Message)
            Throw New ServerUnavailableException(lpProcessedMessage, lstrContentLocation, ServerUnavailableEx)
          Catch ex As Exception
            lpProcessedMessage = String.Format("Invalid content location for document {0} - '{1}': {2}",
                                               lpDocumentId, lstrContentLocation, ex.Message)
            ' Re-throw the exception to the caller
            Throw
          End Try
        Else
          Throw New UnknownItemException(String.Format("Unexpected element type {0} for document {1}",
            lobjContentElement.GetType.Name, lpDocumentId), lpDocumentId)
        End If
      Next

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    Finally

      lobjContentStream?.Dispose()

      lobjIDocument = Nothing
      lobjContentStream = Nothing
      lobjContentElementList = Nothing

    End Try

  End Sub

  ''' <summary>
  ''' ChangeState
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub ChangeState(lpDocumentId As String,
                         lpFlags As LifecycleChangeFlagEnum)

    Dim lobjIDocument As IDocument = Nothing

    Try

      lobjIDocument = GetIDocument(lpDocumentId)

      lobjIDocument.ChangeState(lpFlags)
      lobjIDocument.Save(RefreshMode.NO_REFRESH)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    Finally
      lobjIDocument = Nothing
    End Try

  End Sub

  Public Sub DemoteVersion(lpDocumentId As String)
    Try
      Dim lobjIDocument As IDocument = GetIDocument(lpDocumentId)
      If lobjIDocument.VersionStatus = VersionStatus.RELEASED Then
        If lobjIDocument.IsCurrentVersion = True Then
          If Not lobjIDocument.IsReserved Then
            lobjIDocument.DemoteVersion()
            lobjIDocument.Save(RefreshMode.NO_REFRESH)
          Else
            Throw New InvalidOperationException("Unable to demote version, this version is reserved.")
          End If
        Else
          Throw New InvalidOperationException("Unable to demote version, this is not the current version.")
        End If
      Else
        Throw New InvalidOperationException(String.Format("Unable to demote version, the version status is {0}.", lobjIDocument.VersionStatus))
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub PromoteVersion(lpDocumentId As String)
    Try
      Dim lobjIDocument As IDocument = GetIDocument(lpDocumentId)
      If lobjIDocument.VersionStatus = VersionStatus.IN_PROCESS Then
        If lobjIDocument.IsCurrentVersion = True Then
          lobjIDocument.PromoteVersion()
          lobjIDocument.Save(RefreshMode.NO_REFRESH)
        Else
          Throw New InvalidOperationException("Unable to promote version, this is not the current version.")
        End If
      Else
        Throw New InvalidOperationException(String.Format("Unable to promote version, the version status is {0}.", lobjIDocument.VersionStatus))
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub IndexContent(lpDocumentId As String,
                         ByRef lpProcessedMessage As String)
    Try
      Dim lobjIndexJob As IIndexJob = Factory.IndexJob.CreateInstance(ObjectStore)
      Dim lobjIndexJobItem As IIndexJobSingleItem = Factory.IndexJobSingleItem.CreateInstance()
      Dim lobjIndexJobItemList As IIndexJobItemList = Factory.IndexJobItem.CreateList()
      Dim lobjDocument As IDocument = GetIDocument(lpDocumentId)

      lobjIndexJobItem.SingleItem = lobjDocument

      lobjIndexJobItemList.Add(lobjIndexJobItem)

      lobjIndexJob.DescriptiveText = String.Format("Index '{0}'", lpDocumentId)
      lobjIndexJob.IndexItems = lobjIndexJobItemList

      lobjIndexJob.Save(RefreshMode.NO_REFRESH)

      lpProcessedMessage = String.Format("Document '{0}' queued for indexing.", lpDocumentId)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub MoveAnnotation(lpAnnotationId As String,
                         lpStorageAreaName As String,
                         lpPreserveLastModifiedInfo As Boolean,
                         ByRef lpProcessedMessage As String)
    Try
      If ((lpPreserveLastModifiedInfo = True) AndAlso (HasElevatedPrivileges = False)) Then
        Throw New UserDoesNotHaveElevatedPriviledgesException(Me.UserName)
      End If

      Dim lobjIAnnotation As IAnnotation = GetIAnnotation(lpAnnotationId)
      Dim lobjCurrentStorageArea As IStorageArea = Nothing
      Dim lobjDestinationStorageArea As IStorageArea = GetStorageArea(lpStorageAreaName)

      If lobjDestinationStorageArea Is Nothing Then
        Throw New ItemNotFoundException(lpStorageAreaName)
      End If

      ' Get the storage area.
      lobjCurrentStorageArea = lobjIAnnotation.StorageArea

      ' Make sure the current and destination storage areas 
      ' are not the same.  If we call move in this case FileNet 
      ' will not complain but there will be no change either.  
      ' If the intent is to move the content to force encryption 
      ' or something else, we would hate to provde a false impression 
      ' that the content had been changed.
      If ((lobjCurrentStorageArea IsNot Nothing) AndAlso
          (String.Equals(lobjCurrentStorageArea.DisplayName, lobjDestinationStorageArea.DisplayName))) Then
        Throw New InvalidOperationException(String.Format("The annotation is already in the requested storage area: {0}.", lobjCurrentStorageArea.DisplayName))
      End If

      lobjIAnnotation.MoveContent(lobjDestinationStorageArea)
      SaveObject(lobjIAnnotation, lpPreserveLastModifiedInfo)

      If lobjCurrentStorageArea IsNot Nothing Then
        lpProcessedMessage = String.Format("Successfully moved annotation '{0}' from storage area '{1}' to '{2}'",
                                           lobjIAnnotation.Id, lobjCurrentStorageArea.DisplayName,
                                           lobjDestinationStorageArea.DisplayName)
      Else
        lpProcessedMessage = String.Format("Successfully moved annotation '{0}' to '{1}'",
                                           lobjIAnnotation.Id, lobjDestinationStorageArea.DisplayName)
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub MoveContent(lpDocumentId As String,
                         lpStorageAreaName As String,
                         lpMoveAllVersions As Boolean,
                         lpPreserveLastModifiedInfo As Boolean,
                         ByRef lpProcessedMessage As String)
    Try

      If DocumentExists(lpDocumentId) = False Then
        ' See if it exists as an annotation
        If DocumentExists(lpDocumentId, "Annotation") = False Then
          Throw New DocumentDoesNotExistException(lpDocumentId)
        Else
          MoveAnnotation(lpDocumentId, lpStorageAreaName, lpPreserveLastModifiedInfo, lpProcessedMessage)
          Exit Sub
        End If
      End If

      If ((lpPreserveLastModifiedInfo = True) AndAlso (HasElevatedPrivileges = False)) Then
        Throw New UserDoesNotHaveElevatedPriviledgesException(Me.UserName)
      End If

      Dim lobjIDocument As IDocument = GetIDocument(lpDocumentId)
      Dim lobjVersionSeries As IVersionSeries = Nothing
      Dim lobjCurrentStorageArea As IStorageArea = Nothing
      Dim lobjDestinationStorageArea As IStorageArea = GetStorageArea(lpStorageAreaName)
      Dim lobjCurrentVersion As IDocument = Nothing

      If lobjDestinationStorageArea Is Nothing Then
        Throw New ItemNotFoundException(lpStorageAreaName)
      End If

      ' Make sure the document is not checked out
      If lobjIDocument.CurrentVersion.IsReserved Then
        Dim lobjCheckedOutVersion As IDocument = lobjIDocument.CurrentVersion
        Dim lstrLastModifier As String = lobjCheckedOutVersion.Properties("LastModifier")
        Throw New DocumentCheckedOutException(lpDocumentId, String.Format("Document is checked out by {0}", lstrLastModifier))
      End If

      If lpMoveAllVersions Then
        ' Get a reference to the version series
        lobjVersionSeries = lobjIDocument.VersionSeries

        ' Get a reference to the current version
        lobjCurrentVersion = lobjVersionSeries.CurrentVersion

        ' Get the storage area for the current version, 
        ' even though there could technically be multiple 
        ' versions and not all versions may be in the same 
        ' storage area, we will use the current version as 
        ' a proxy for the version series as a whole.
        lobjCurrentStorageArea = lobjCurrentVersion.StorageArea

        ' Make sure the current and destination storage areas 
        ' are not the same.  If we call move in this case FileNet 
        ' will not complain but there will be no change either.  
        ' If the intent is to move the content to force encryption 
        ' or something else, we would hate to provde a false impression 
        ' that the content had been changed.
        If ((lobjCurrentStorageArea IsNot Nothing) AndAlso
            (String.Equals(lobjCurrentStorageArea.DisplayName, lobjDestinationStorageArea.DisplayName))) Then
          Throw New InvalidOperationException(String.Format("The document is already in the requested storage area: {0}.", lobjCurrentStorageArea.DisplayName))
        End If

        ' <Modified by: Ernie at 4/8/2014-9:04:13 AM on machine: ERNIE-THINK>
        ' Move the content for the entire version series
        ''lobjVersionSeries.MoveContent(lobjDestinationStorageArea)
        ''lobjVersionSeries.Save(RefreshMode.NO_REFRESH)
        If lpPreserveLastModifiedInfo Then
          ' Loop through each version manually
          For Each lobjP8Version As IDocument In lobjVersionSeries.Versions
            lobjP8Version.MoveContent(lobjDestinationStorageArea)
            SaveObject(lobjP8Version, True)
          Next
        Else
          lobjVersionSeries.MoveContent(lobjDestinationStorageArea)
          lobjVersionSeries.Save(RefreshMode.NO_REFRESH)
        End If

        ' </Modified by: Ernie at 4/8/2014-9:04:13 AM on machine: ERNIE-THINK>


        ' Generate the processed message.
        If lobjCurrentStorageArea IsNot Nothing Then
          lpProcessedMessage = String.Format("Successfully moved version series '{0}' from storage area '{1}' to '{2}'",
                                             lobjVersionSeries.Id, lobjCurrentStorageArea.DisplayName,
                                             lobjDestinationStorageArea.DisplayName)
        Else
          lpProcessedMessage = String.Format("Successfully moved version series '{0}' to '{1}'",
                                             lobjVersionSeries.Id, lobjDestinationStorageArea.DisplayName)
        End If

        lobjVersionSeries = Nothing
        lobjCurrentVersion = Nothing
        lobjCurrentStorageArea = Nothing
        lobjDestinationStorageArea = Nothing

      Else

        ' Get the current storage area for the document.
        lobjCurrentStorageArea = lobjIDocument.StorageArea

        ' Make sure the current and destination storage areas 
        ' are not the same.  If we call move in this case FileNet 
        ' will not complain but there will be no change either.  
        ' If the intent is to move the content to force encryption 
        ' or something else, we would hate to provde a false impression 
        ' that the content had been changed.
        If ((lobjCurrentStorageArea IsNot Nothing) AndAlso
            (String.Equals(lobjCurrentStorageArea.DisplayName, lobjDestinationStorageArea.DisplayName))) Then
          Throw New InvalidOperationException("The document is already filed in the requested storage area.")
        End If

        ' Move the content for the document (note that this will only move this version).
        lobjIDocument.MoveContent(lobjDestinationStorageArea)
        lobjIDocument.Save(RefreshMode.NO_REFRESH)
        lobjIDocument = Nothing
        ' Generate the processed message.
        If lobjCurrentStorageArea IsNot Nothing Then
          lpProcessedMessage = String.Format("Successfully moved document '{0}' from storage area '{1}' to '{2}'",
                                             lobjIDocument.Id, lobjCurrentStorageArea.DisplayName,
                                             lobjDestinationStorageArea.DisplayName)
        Else
          lpProcessedMessage = String.Format("Successfully moved document '{0}' to '{1}'",
                                             lobjIDocument.Id, lobjDestinationStorageArea.DisplayName)
        End If

        lobjCurrentStorageArea = Nothing
        lobjDestinationStorageArea = Nothing

      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub


  Private Function GetP8ChoiceList(lpIdentifier As String) As Admin.IChoiceList
    Try

      Dim lblnIsGuidId As Boolean = Helper.IsGuid(lpIdentifier, Nothing)
      'If Helper.IsGuid(lpIdentifier, Nothing) Then
      '  lobjChoiceList = ObjectStore.GetObject("ChoiceList", New Util.Id(lpArgs.ID))
      'Else
      '  lobjChoiceList = ObjectStore.GetObject("ChoiceList", lpArgs.ID)
      'End If

      If ObjectStore.ChoiceLists IsNot Nothing Then
        For Each lobjChoiceList As Admin.IChoiceList In ObjectStore.ChoiceLists
          If lblnIsGuidId Then
            If String.Equals(lobjChoiceList.Id.ToString, lpIdentifier, StringComparison.InvariantCultureIgnoreCase) Then
              Return lobjChoiceList
            End If
          Else
            If String.Equals(lobjChoiceList.Name, lpIdentifier, StringComparison.InvariantCultureIgnoreCase) Then
              Return lobjChoiceList
            End If
          End If
        Next
      End If

      Return Nothing

    Catch ex As Exception
      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Sub CENetProvider_ProviderProperty_ValueChanged(sender As Object, ByRef e As ProviderPropertyValueChangedEventArgs) Handles Me.ProviderProperty_ValueChanged
    Try
      ' Beep()
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Private Sub CENetProvider_ProviderPropertyChanged(sender As Object, e As ComponentModel.PropertyChangedEventArgs) Handles Me.ProviderPropertyChanged
    Try
      If sender IsNot Nothing AndAlso TypeOf sender Is ProviderProperty Then
        Dim lobjChangedProperty As ProviderProperty = sender
        Select Case lobjChangedProperty.PropertyName
          Case "ServerName"
            mstrServerName = lobjChangedProperty.Value

          Case "UserName"
            UserName = lobjChangedProperty.Value

          Case "Password"
            Password = lobjChangedProperty.Value

          Case PROTOCOL
            ' lobjChangedProperty.SupportsValueList = True
            mstrProtocol = lobjChangedProperty.Value

        End Select
      End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

End Class
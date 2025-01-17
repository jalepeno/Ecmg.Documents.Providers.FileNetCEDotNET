
#Region "Imports"

Imports FileNet.Api.Util
Imports FileNet.Api.Authentication
Imports FileNet.Api.Core
Imports FileNet.Api.Security
Imports FileNet.Api.Collection
Imports FileNet.Api.Admin
Imports System.IO
Imports System.Text
Imports FileNet.Api.Property
Imports Documents.Exceptions
Imports Documents.Utilities


#End Region

<DebuggerDisplay("{DebuggerIdentifier(),nq}")>
Public Class StorageUtilities

#Region "Class Constants"

  Private Const ROLLING_STORAGE_AREA_PREFIX As String = "RFSA"
  Private Const FIRST_ROLLING_STORAGE_AREA As String = "RFSA1"
  Private Const ROLLING_STORAGE_POLICY_NAME As String = "Rolling Storage"

#End Region

#Region "Event Delegates"

  Public Delegate Sub StorageAreaAddedEventHandler(ByVal sender As Object, ByRef e As StorageAreaAddedEventArgs)

#End Region

#Region "Public Events"

  Public Event StorageAreaAdded As StorageAreaAddedEventHandler

#End Region

#Region "Class Variables"

  Private mobjDomain As IDomain = Nothing
  Private mobjCurrentObjectStore As IObjectStore = Nothing
  Private mobjCurrentEnvironment As Environment = Nothing
  Private mobjStorageAreaNames As List(Of String) = Nothing
  Private mobjStoragePolicyNames As List(Of String) = Nothing
  Private mobjEnvironments As Environments = Nothing
  Private mstrSharedStorageRoot As String = String.Empty
  Private mintDefaultMaxContentElements As Integer

#End Region

#Region "Public Classes"

#End Region

#Region "Public Properties"

  Public ReadOnly Property CurrentScope As String
    Get
      Return GetCurrentScope()
    End Get
  End Property

  Public Property SharedStorageRoot As String
    Get
      Return mstrSharedStorageRoot
    End Get
    Set(ByVal value As String)
      mstrSharedStorageRoot = value
    End Set
  End Property

  Public Property DefaultMaxContentElements As Integer
    Get
      Return mintDefaultMaxContentElements
    End Get
    Set(ByVal value As Integer)
      mintDefaultMaxContentElements = value
    End Set
  End Property

  Public ReadOnly Property StorageAreaNames As List(Of String)
    Get
      Try
        If mobjStorageAreaNames Is Nothing Then
          mobjStorageAreaNames = GetAllStorageAreaNames()
        End If
        Return mobjStorageAreaNames
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

  Public ReadOnly Property StoragePolicyNames As List(Of String)
    Get
      Try
        If mobjStoragePolicyNames Is Nothing Then
          mobjStoragePolicyNames = GetAllStoragePolicyNames()
        End If
        Return mobjStoragePolicyNames
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

  Public Property Environments As Environments

#End Region

#Region "Constructors"

  Public Sub New(ByVal lpEnvironments As Environments)
    Try
      Environments = lpEnvironments
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub New(ByVal lpEnvironments As Environments, ByVal lpSharedStorageRoot As String, ByVal lpDefaultMaxContentElements As Integer)
    Try
      Environments = lpEnvironments
      SharedStorageRoot = lpSharedStorageRoot
      DefaultMaxContentElements = lpDefaultMaxContentElements
      'InitializeConnection()

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  'Public Sub New(ByVal lpObjectStore As IObjectStore)
  '  Try
  '    mobjCurrentObjectStore = lpObjectStore
  '    Refresh()
  '  Catch ex As Exception
  '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '    ' Re-throw the exception to the caller
  '    Throw
  '  End Try
  'End Sub

#End Region

#Region "Public Methods"

  Public Sub Refresh()
    Try
      mobjStorageAreaNames = GetAllStorageAreaNames()
      mobjStoragePolicyNames = GetAllStoragePolicyNames()
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Function GetAllStorageAreaNames() As List(Of String)
    Try

      Dim lstrStorageAreas As New List(Of String)
      Dim lobjStorageAreas As FileNet.Api.Collection.IStorageAreaSet = mobjCurrentObjectStore.StorageAreas

      For Each lobjStorageArea As IStorageArea In lobjStorageAreas
        lstrStorageAreas.Add(lobjStorageArea.DisplayName)
      Next

      Return lstrStorageAreas

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function GetAllStoragePolicyNames() As List(Of String)
    Try

      Dim lstrStoragePolicies As New List(Of String)
      Dim lobjStoragePolicies As FileNet.Api.Collection.IStoragePolicySet = mobjCurrentObjectStore.StoragePolicies

      For Each lobjStoragePolicy As IStoragePolicy In lobjStoragePolicies
        lstrStoragePolicies.Add(lobjStoragePolicy.DisplayName)
      Next

      Return lstrStoragePolicies

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function GetAllStorageAreaNamesForPolicy(ByVal lpPolicyName As String) As List(Of String)
    Try

      Dim lstrStorageAreas As New List(Of String)
      Dim lobjPolicy As IStoragePolicy = GetStoragePolicy(lpPolicyName)

      For Each lobjStorageArea As IStorageArea In lobjPolicy.StorageAreas
        lstrStorageAreas.Add(lobjStorageArea.DisplayName)
      Next

      Return lstrStorageAreas

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function CreateRollingStoragePolicy(ByVal lpName As String,
                                        ByVal lpInitialStorageAreaCount As Integer,
                                        ByVal lpEnvironmentRootPath As String,
                                        ByVal lpMaximumContentElements As Integer,
                                        ByVal lpEnableDeDuplication As Boolean) As IStoragePolicy
    Try

      Console.WriteLine("  Creating {0} for {1}: {2}", lpName, mobjCurrentObjectStore.Name, mobjCurrentEnvironment.ToString)

      Dim lobjStoragePolicy As IStoragePolicy = CreateStoragePolicy(lpName, True)
      'Dim lobjStorageAreas As IList(Of IStorageArea) = New List(Of IStorageArea)
      Dim lobjStorageAreaIds(lpInitialStorageAreaCount - 1) As String

      If lpInitialStorageAreaCount > 0 Then
        Dim lstrStorageAreaName As String = Nothing
        Dim lobjStorageArea As IStorageArea = Nothing
        For lintStorageAreaCounter As Integer = 0 To lpInitialStorageAreaCount - 1
          lstrStorageAreaName = String.Format("{0}{1}", ROLLING_STORAGE_AREA_PREFIX, lintStorageAreaCounter + 1)
          Console.WriteLine("    Creating Storage Area {0}", lstrStorageAreaName)
          lobjStorageArea = CreateFileStorageArea(lstrStorageAreaName, lpEnvironmentRootPath, lpMaximumContentElements, lpEnableDeDuplication)

          ' We only want the first storage area to be open, the others should be in standby mode.
          If lintStorageAreaCounter > 0 Then
            lobjStorageArea.ResourceStatus = FileNet.Api.Constants.ResourceStatus.STANDBY
            lobjStorageArea.Save(FileNet.Api.Constants.RefreshMode.REFRESH)
          End If

          lobjStorageAreaIds(lintStorageAreaCounter) = lobjStorageArea.Id.ToString

        Next
        lobjStoragePolicy.FilterExpression = CreateFilterExpression(lobjStorageAreaIds)
        lobjStoragePolicy.Save(FileNet.Api.Constants.RefreshMode.REFRESH)
      End If

      Return lobjStoragePolicy

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Sub DeleteStoragePolicy(ByVal lpPolicyName As String, ByVal lpIncludeStorageAreas As Boolean, ByVal lpDeleteFolders As Boolean)
    Try

      Dim lobjStoragePolicy As IStoragePolicy = GetStoragePolicy(lpPolicyName)

      If lpIncludeStorageAreas Then
        For Each lobjStorageArea As IStorageArea In lobjStoragePolicy.StorageAreas
          Debug.Print("Deleting Storage Area {0}", lobjStorageArea.DisplayName)
          DeleteStorageArea(lobjStorageArea, lpDeleteFolders)
        Next
      End If

      Debug.Print("Deleting Storage Policy {0}", lobjStoragePolicy.DisplayName)
      lobjStoragePolicy.Delete()

      lobjStoragePolicy.Save(FileNet.Api.Constants.RefreshMode.REFRESH)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub DeleteStorageArea(ByVal lpStorageArea As IStorageArea, ByVal lpDeleteFolders As Boolean)
    Try

      If TypeOf lpStorageArea Is IFileStorageArea AndAlso lpDeleteFolders Then
        Dim lstrRootDirectory As String = CType(lpStorageArea, IFileStorageArea).RootDirectoryPath
        Directory.Delete(lstrRootDirectory, True)
      End If

      lpStorageArea.Delete()

      lpStorageArea.Save(FileNet.Api.Constants.RefreshMode.REFRESH)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Function CreateStoragePolicy(ByVal lpName As String, ByVal lpSave As Boolean) As IStoragePolicy
    Try

      Dim lobjStorageArea As IStoragePolicy = FileNet.Api.Core.Factory.StoragePolicy.CreateInstance(mobjCurrentObjectStore)

      With lobjStorageArea
        .DisplayName = lpName
        .DescriptiveText = lpName
        If lpSave Then
          .Save(FileNet.Api.Constants.RefreshMode.REFRESH)
        End If
      End With

      Return lobjStorageArea

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Sub AddStorageAreaToPolicy(ByVal lpPolicy As IStoragePolicy, ByVal lpStorageArea As IStorageArea)
    Try

      ' Create a list of all storage areas to associate with the policy.
      Dim lobjNewStorageAreaList As New List(Of IStorageArea)
      For Each lobjStorageArea As IStorageArea In lpPolicy.StorageAreas
        lobjNewStorageAreaList.Add(lobjStorageArea)
      Next
      lobjNewStorageAreaList.Add(lpStorageArea)

      SetStorageAreasForPolicy(lpPolicy, lobjNewStorageAreaList)

      Dim lstrCurrentStatus As ObjectStoreRollingPolicyStatus = GetCurrentStorageStatistics()

      RaiseEvent StorageAreaAdded(Me, New StorageAreaAddedEventArgs(mobjCurrentEnvironment, lpStorageArea, lpPolicy, lstrCurrentStatus))

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Private Sub SetStorageAreasForPolicy(ByVal lpPolicy As IStoragePolicy, ByVal lpStorageAreas As IList(Of IStorageArea))
    Try

      CType(lpPolicy.StorageAreas, Object).SetList(lpStorageAreas)

      lpPolicy.Save(FileNet.Api.Constants.RefreshMode.REFRESH)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Function CreateFileStorageArea(ByVal lpName As String,
                                        ByVal lpEnvironmentRootPath As String,
                                        ByVal lpMaximumContentElements As Integer,
                                        ByVal lpEnableDeDuplication As Boolean) As IStorageArea
    Try

      Dim lobjStorageArea As IFileStorageArea = FileNet.Api.Core.Factory.FileStorageArea.CreateInstance(mobjCurrentObjectStore, Nothing)
      lpEnvironmentRootPath = lpEnvironmentRootPath.TrimEnd("\")
      Dim lstrRootDirectoryStructure As String = String.Format("{0}\{1}\{2}", lpEnvironmentRootPath, mobjCurrentObjectStore.SymbolicName, lpName)

      If IO.Directory.Exists(lstrRootDirectoryStructure) = False Then
        IO.Directory.CreateDirectory(lstrRootDirectoryStructure)
      End If

      With lobjStorageArea
        .DisplayName = lpName
        .DescriptiveText = "Rolling Storage Area"
        .DirectoryStructure = FileNet.Api.Constants.DirectoryStructure.DIRECTORY_STRUCTURE_SMALL
        .RootDirectoryPath = lstrRootDirectoryStructure
        .DeleteMethod = FileNet.Api.Constants.AreaDeleteMethod.STANDARD
        .MaximumContentElements = lpMaximumContentElements
        .DuplicateSuppressionEnabled = lpEnableDeDuplication
      End With

      lobjStorageArea.Save(FileNet.Api.Constants.RefreshMode.REFRESH)

      Return lobjStorageArea

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

#Region "Private Methods"

  'Private Function CreateURL() As String
  '  Try
  '    Return String.Format("{0}://{1}:{2}/wsi/FNCEWS$)MTOM/", "http", ServerName, Port)
  '  Catch ex As Exception
  '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '    ' Re-throw the exception to the caller
  '    Throw
  '  End Try
  'End Function

  'Private Function GetDomain() As IDomain

  '  Try

  '    Dim lobjConnection As IConnection = Factory.Connection.GetConnection(URL)
  '    Dim lobjDomain As IDomain = Factory.Domain.GetInstance(lobjConnection, Nothing)

  '    Return lobjDomain

  '  Catch ex As Exception
  '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '    ' Re-throw the exception to the caller
  '    Throw
  '  End Try

  'End Function

  'Private Function InitializeConnection() As Boolean

  '  Console.WriteLine("Please wait, attempting to connect...")

  '  Try

  '    Dim lobjCredentials As Credentials = Nothing

  '    lobjCredentials = New UsernameCredentials(mstrUserName, mstrPassword)
  '    FileNet.Api.Util.ClientContext.SetProcessCredentials(lobjCredentials)

  '    mobjDomain = GetDomain()

  '    Try
  '      mobjDomain.Refresh()

  '    Catch ex As Exception
  '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '      IsConnected = False
  '      '  Re-throw the exception
  '      Throw
  '    End Try

  '    Return True

  '  Catch ex As Exception
  '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '    IsConnected = False
  '    Throw (New Exception("Unable to initialize connection", ex))
  '  End Try

  'End Function

  Friend Function DebuggerIdentifier() As String
    Try
      Return CurrentScope
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function GetCurrentScope() As String
    Try
      Dim lobjStringBuilder As New StringBuilder

      If mobjCurrentEnvironment IsNot Nothing Then
        lobjStringBuilder.AppendFormat("{0}: ", mobjCurrentEnvironment.Name)
      Else
        lobjStringBuilder.Append("No Current Environment: ")
      End If

      If mobjCurrentObjectStore IsNot Nothing Then
        lobjStringBuilder.AppendFormat("{0}", mobjCurrentObjectStore.Name)
      Else
        lobjStringBuilder.Append("No Current ObjectStore")
      End If

      Return lobjStringBuilder.ToString

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function GetStorageArea(ByVal lpAreaName As String) As IStorageArea
    Try

      Dim lobjStorageAreas As FileNet.Api.Collection.IStorageAreaSet = mobjCurrentObjectStore.StorageAreas

      For Each lobjStorageArea As IStorageArea In lobjStorageAreas
        If String.Equals(lobjStorageArea.DisplayName, lpAreaName) Then
          Return lobjStorageArea
        End If
      Next

      ' We did not find the storage area, throw an exception to the caller.
      Throw New Exceptions.StorageAreaNotFoundException(String.Format("The storage area '{0}' was not found in object store '{1}.", lpAreaName, mobjCurrentObjectStore.Name))

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function GetRollingStoragePolicy() As IStoragePolicy
    Try

      Dim lobjStoragePolicy As IStoragePolicy = Nothing
      Dim lobjStorageAreaNames As New List(Of String)

      For Each lstrStoragePolicyName As String In Me.StoragePolicyNames
        lobjStorageAreaNames.Clear()
        lobjStorageAreaNames = GetAllStorageAreaNamesForPolicy(lstrStoragePolicyName)
        If lobjStorageAreaNames.Contains(FIRST_ROLLING_STORAGE_AREA) Then
          lobjStoragePolicy = GetStoragePolicy(lstrStoragePolicyName)
          If lobjStoragePolicy Is Nothing Then
            Throw New Exceptions.StoragePolicyNotFoundException(lstrStoragePolicyName)
          End If
          Return lobjStoragePolicy
        End If
      Next

      Throw New Exceptions.StoragePolicyNotFoundException(String.Format("Rolling Storage Policy for {0}: {1}", mobjCurrentEnvironment.Name, mobjCurrentObjectStore.Name))

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function RollingStoragePolicyExists() As Boolean
    Try

      Dim lobjStoragePolicy As IStoragePolicy = Nothing
      Dim lobjStorageAreaNames As New List(Of String)

      For Each lstrStoragePolicyName As String In Me.StoragePolicyNames
        lobjStorageAreaNames.Clear()
        lobjStorageAreaNames = GetAllStorageAreaNamesForPolicy(lstrStoragePolicyName)
        If lobjStorageAreaNames.Contains(FIRST_ROLLING_STORAGE_AREA) Then
          Return True
        End If
      Next

      Return False

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function CheckAndUpdatePoliciesForAllObjectStores() As SiteSummary
    Try

      Dim lobjStatuses As New SiteSummary
      Dim lobjNewStorageArea As IFileStorageArea
      'Dim lobjObjectStoreRollingPolicyStatus As ObjectStoreRollingPolicyStatus
      Dim lobjEnvironmentSummary As EnvironmentSummary
      'Me.Environments.InitializeConnections()

      For Each lobjEnvironment As Environment In Me.Environments
        mobjCurrentEnvironment = lobjEnvironment
        mobjCurrentEnvironment.InitializeConnection()
        lobjEnvironmentSummary = New EnvironmentSummary(mobjCurrentEnvironment)
        For Each lobjObjectStore As IObjectStore In mobjCurrentEnvironment.Domain.ObjectStores
          mobjCurrentObjectStore = lobjObjectStore
          Refresh()
          lobjNewStorageArea = Nothing
          'lobjObjectStoreRollingPolicyStatus = GetCurrentStorageStatistics()
          'lobjStatuses.ObjectStoreSummaries.Add(GetCurrentStorageStatistics())
          Try
            lobjEnvironmentSummary.ObjectStoreSummaries.Add(CheckAndUpdateRollingStoragePolicyForCurrentObjectStore)
          Catch ex As Exception
            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
            ' Keep going
          End Try
          'lobjStatuses.Add(CheckAndUpdateRollingStoragePolicyForCurrentObjectStore())
          ' lobjObjectStoreRollingPolicyStatus =New ObjectStoreRollingPolicyStatus(lobjObjectStore.Name,
        Next
        lobjStatuses.EnvironmentSummaries.Add(lobjEnvironmentSummary)
      Next

      Return lobjStatuses

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function GetCurrentStorageStatisticsForAllEnvironments() As SiteSummary
    Try
      Dim lobjSiteSummary As New SiteSummary
      Dim lobjEnvironmentSummary As EnvironmentSummary '= New EnvironmentSummary(lobjEnvironment)

      For Each lobjEnvironment As Environment In Me.Environments
        mobjCurrentEnvironment = lobjEnvironment
        Try
          mobjCurrentEnvironment.InitializeConnection()
        Catch RepEx As RepositoryNotAvailableException
          ApplicationLogging.LogException(RepEx, Reflection.MethodBase.GetCurrentMethod)
          Console.WriteLine(RepEx.Message)
          Continue For
        End Try

        lobjEnvironmentSummary = New EnvironmentSummary(lobjEnvironment)

        For Each lobjObjectStore As IObjectStore In mobjCurrentEnvironment.Domain.ObjectStores
          mobjCurrentObjectStore = lobjObjectStore
          'Refresh()
          lobjEnvironmentSummary.ObjectStoreSummaries.Add(GetCurrentStorageStatistics())
        Next
        lobjSiteSummary.EnvironmentSummaries.Add(lobjEnvironmentSummary)
      Next

      Return lobjSiteSummary

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  ''' <summary>
  ''' Checks the rolling storage policy to see is an open area.  If not a new area will created and a reference returned using the storage area parameter.
  ''' </summary>
  ''' <returns>'Normal' if an area is currently open, otherwise 'StorageAreaAdded'.</returns>
  ''' <remarks></remarks>
  Public Function CheckAndUpdateRollingStoragePolicyForCurrentObjectStore() As ObjectStoreRollingPolicyStatus
    Try

      Dim lobjRollingPolicy As IStoragePolicy = Nothing
      Dim lobjObjectStoreRollingPolicyStatus As ObjectStoreRollingPolicyStatus = Nothing

      'Dim lobjDocumentClassDefinition As IClassDefinition = GetCurrentRootDocumentClass()

      If RollingStoragePolicyExists() Then
        lobjRollingPolicy = GetRollingStoragePolicy()
      Else
        ' We will need to create it...
        Dim lstrEnvironmentRoot As String = mobjCurrentEnvironment.CreateEnvironmentRoot(Me.SharedStorageRoot)
        lobjRollingPolicy = CreateRollingStoragePolicy(ROLLING_STORAGE_POLICY_NAME, 2, lstrEnvironmentRoot, DefaultMaxContentElements, True)
      End If

      Dim lobjFileStorageAreas As List(Of IFileStorageArea) = GetAllFileStorageAreasForPolicy(lobjRollingPolicy)
      Dim lobjFileStorageArea As IFileStorageArea = Nothing
      Dim lintFileStorageCounter As Integer = 1
      Dim lblnHasOpenArea As Boolean = False
      Dim lobjNewStorageArea As IStorageArea = Nothing

      For Each lobjFileStorageArea In lobjFileStorageAreas
        If lobjFileStorageArea.ResourceStatus = FileNet.Api.Constants.ResourceStatus.OPEN Then
          lblnHasOpenArea = True
          Exit For
        End If
        lintFileStorageCounter += 1
      Next

      If lblnHasOpenArea = False Then

        ' We do not have any open storage areas, we need to create one.
        Dim lstrNewStorageAreaName As String = String.Format("{0}{1}",
                                                             ROLLING_STORAGE_AREA_PREFIX, lintFileStorageCounter)

        ' Get the object store level of the root directory path as a starting point for the new path
        Dim lstrOSRoot As String = lobjFileStorageArea.RootDirectoryPath.Substring(0, lobjFileStorageArea.RootDirectoryPath.LastIndexOf("\"))

        ' Create the new root path by adding the new area name to the object store root path
        Dim lstrNewRootPath As String = String.Format("{0}\{1}", lstrOSRoot, lstrNewStorageAreaName)

        ' Create the new area using some of the information from the previous area.
        'Dim lobjNewStorageArea As IStorageArea = CreateFileStorageArea(lstrNewStorageAreaName, _
        '                                                               lstrNewRootPath, _
        '                                                               lobjFileStorageArea.MaximumContentElements, _
        '                                                               lobjFileStorageArea.DuplicateSuppressionEnabled)

        lobjNewStorageArea = CreateFileStorageArea(lstrNewStorageAreaName,
                                                               lstrNewRootPath,
                                                               250000,
                                                               lobjFileStorageArea.DuplicateSuppressionEnabled)

        ' Add the new area to the policy.
        AddStorageAreaToPolicy(lobjRollingPolicy, lobjNewStorageArea)

      End If


      lobjObjectStoreRollingPolicyStatus = GetCurrentStorageStatistics()
      If lobjNewStorageArea IsNot Nothing Then
        lobjObjectStoreRollingPolicyStatus.NewStorageArea = lobjNewStorageArea
        lobjObjectStoreRollingPolicyStatus.Status = RollingPolicyStatus.StorageAreaAdded
      Else
        lobjObjectStoreRollingPolicyStatus.Status = RollingPolicyStatus.Normal
      End If

      ''Dim lobjDocumentClassDefinition As IClassDefinition = GetCurrentRootDocumentClass()

      'If lobjDocumentClassDefinition IsNot Nothing Then
      '  'lobjDocumentClassDefinition.s
      'End If

      Return lobjObjectStoreRollingPolicyStatus

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function GetStoragePolicy(ByVal lpPolicyName As String) As IStoragePolicy
    Try

      Dim lobjStoragePolicies As FileNet.Api.Collection.IStoragePolicySet = mobjCurrentObjectStore.StoragePolicies

      For Each lobjStoragePolicy As IStoragePolicy In lobjStoragePolicies
        If String.Equals(lobjStoragePolicy.DisplayName, lpPolicyName) Then
          Return lobjStoragePolicy
        End If
      Next

      ' We did not find the storage policy, throw an exception to the caller.
      Throw New Exceptions.StoragePolicyNotFoundException(String.Format("The storage policy '{0}' was not found in object store '{1}.",
                                                                        lpPolicyName, mobjCurrentObjectStore.Name))

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function GetAllFileStorageAreasForPolicy(ByVal lpStoragePolicy As IStoragePolicy) As List(Of IFileStorageArea)
    Try

      Dim lobjFileStorageAreas As New List(Of IFileStorageArea)

      For Each lobjStorageArea As IStorageArea In lpStoragePolicy.StorageAreas
        If TypeOf lobjStorageArea Is IFileStorageArea Then
          lobjFileStorageAreas.Add(CType(lobjStorageArea, IFileStorageArea))
        End If
      Next

      Return lobjFileStorageAreas

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function GetCurrentStorageStatistics() As ObjectStoreRollingPolicyStatus
    Try

      Dim lobjReturn As ObjectStoreRollingPolicyStatus = Nothing
      Dim lobjStoragePolicy As IStoragePolicy = Nothing

      Dim lintTotalFilesStored As Long
      Dim lintTotalKiloBytesUsed As Long
      Dim lintCurrentTotalFiles As Integer
      Dim lintCurrentMaximumFiles As Integer

      'Dim lobjFileSize As FileSize = Nothing
      Dim lstrRollingStoragePolicyName As String = String.Empty
      Dim lstrCurrentStorageAreaName As String = String.Empty
      Dim lstrCurrentStorageAreaPath As String = String.Empty

      For Each lstrStoragePolicy As String In Me.GetAllStoragePolicyNames
        lobjStoragePolicy = Me.GetStoragePolicy(lstrStoragePolicy)
        For Each lobjStorageArea As IStorageArea In lobjStoragePolicy.StorageAreas
          lintTotalFilesStored += lobjStorageArea.ContentElementCount
          lintTotalKiloBytesUsed += lobjStorageArea.ContentElementKBytes
          If (TypeOf lobjStorageArea Is IFileStorageArea) AndAlso
            (lobjStorageArea.ResourceStatus = FileNet.Api.Constants.ResourceStatus.OPEN) _
            AndAlso lobjStorageArea.MaximumContentElements > 0 Then
            lstrRollingStoragePolicyName = lstrStoragePolicy
            lstrCurrentStorageAreaName = lobjStorageArea.DisplayName
            lstrCurrentStorageAreaPath = CType(lobjStorageArea, IFileStorageArea).RootDirectoryPath
            lintCurrentTotalFiles = lobjStorageArea.ContentElementCount
            lintCurrentMaximumFiles = lobjStorageArea.MaximumContentElements
          End If
        Next
      Next

      'lobjFileSize = New FileSize(lintTotalKiloBytesUsed)

      lobjReturn = New ObjectStoreRollingPolicyStatus(mobjCurrentEnvironment.Name,
                                                      mobjCurrentObjectStore.Name, lstrRollingStoragePolicyName, lstrCurrentStorageAreaName,
                                                      lintTotalFilesStored,
                                                      lintTotalKiloBytesUsed * 1024,
                                                      lstrCurrentStorageAreaPath,
                                                      lintCurrentTotalFiles,
                                                      lintCurrentMaximumFiles,
                                                      RollingPolicyStatus.Normal)

      Return lobjReturn

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function CreateFilterExpression(ByVal lpIdList() As String) As String
    Try

      Dim lobjFilterBuilder As New StringBuilder

      With lobjFilterBuilder
        .Append("Id IN (")
        For lintIdCounter As Integer = 0 To lpIdList.Length - 1
          If Not String.IsNullOrEmpty(lpIdList(lintIdCounter)) Then
            .AppendFormat("{0},", lpIdList(lintIdCounter))
          End If
        Next
        'For Each lstrId As String In lpIdList
        '  .AppendFormat("{0},", lstrId)
        'Next
        .Remove(.Length - 1, 1)
        .Append(")")
      End With

      Return lobjFilterBuilder.ToString

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function


  Public Function GetCurrentRootDocumentClass() As IClassDefinition
    Try
      If mobjCurrentObjectStore Is Nothing Then
        Return Nothing
      End If

      Dim lobjDocClassDefPropFilter As New PropertyFilter
      Dim lobjDefaultStoragePolicyFilterElement As New FilterElement(5, Nothing, Nothing, FileNet.Api.Constants.FilteredPropertyType.ANY, Nothing)
      'lobjDefaultStoragePolicyFilterElement.
      lobjDocClassDefPropFilter.AddIncludeProperty(lobjDefaultStoragePolicyFilterElement)
      Dim lobjClassDefinition As IClassDefinition = mobjCurrentObjectStore.FetchObject("ClassDefinition", "Document", lobjDocClassDefPropFilter)
      lobjClassDefinition.Refresh()
      Dim lobjStoragePolicyProp As IProperty = lobjClassDefinition.FetchProperty("StoragePolicy", New PropertyFilter)

      Debug.Print(lobjStoragePolicyProp.GetIdValue.ToString)

      Return Nothing

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

End Class

'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  CENetProvider_IFolderManager.vb
'   Description :  [type_description_here]
'   Created     :  4/8/2014 11:13:37 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports FileNet.Api.Core
Imports FileNet.Api.Collection
Imports FileNet.Api.Property
Imports FileNet.Api.Constants
Imports System.IO
Imports Documents.Arguments
Imports Documents.Core
Imports Documents.Exceptions
Imports Documents.Providers
Imports Documents.Utilities
Imports DCore = Documents.Core
Imports DProviders = Documents.Providers
Imports DSearch = Documents.Search
Imports DSecurity = Documents.Security

#End Region

Partial Public Class CENetProvider
  Implements IFolderManager

#Region "IFolderManager Implementation"

  Public Function UpdateFolderProperties(ByVal Args As FolderPropertyArgs) As Boolean Implements IFolderManager.UpdateFolderProperties
    Try

      ' Get document and populate property cache.
      Dim lobjIncludePropertyFilter As New PropertyFilter()

      For Each lobjProperty As DCore.IProperty In Args.Properties
        lobjIncludePropertyFilter.AddIncludeProperty(New FilterElement(Nothing, Nothing, Nothing, lobjProperty.SystemName, Nothing))
      Next

      Dim lobjIFolder As FileNet.Api.Core.IFolder

      If Not String.IsNullOrEmpty(Args.FolderId) Then
        lobjIFolder = GetIFolder(Args.FolderId)
      ElseIf Not String.IsNullOrEmpty(Args.FolderPath) Then
        lobjIFolder = GetIFolder(Args.FolderPath)
      Else
        Throw New ArgumentException("Neither a folder id or folder path was specified in the FolderPropertyArgs.", NameOf(Args))
      End If

      Dim lobjFolderClass As FolderClass = FolderClasses(lobjIFolder.GetClassName())

      For Each lobjProperty As ECMProperty In Args.Properties
        SetPropertyValue(lobjIFolder, lobjFolderClass, lobjProperty, False)
      Next

      ' Save and update property cache.
      lobjIFolder.Save(RefreshMode.REFRESH)

      Return True

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function CreateFolder(lpFolderPath As String) As DProviders.IFolder Implements IFolderManager.CreateFolder
    Try
      If FolderPathExists(lpFolderPath) Then
        ' The folder already exists, we can't create it again.
        Throw New FolderAlreadyExistsException(lpFolderPath)
      Else
        ' The method below will create the folder if it does not already exist.
        Dim lobjP8Folder As FileNet.Api.Core.IFolder = GetFolderByPath(lpFolderPath)
        Return New CENetFolder(lobjP8Folder, 0)
      End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function CreateFolder(lpParentId As String,
                               lpName As String,
                               lpClassName As String,
                               lpProperties As DCore.IProperties,
                               lpPermissions As DSecurity.IPermissions) As DProviders.IFolder _
                             Implements IFolderManager.CreateFolder
    Try

      Dim lstrExistingFolderId As String = Nothing
      Dim lobjFolderClass As FolderClass = Nothing

      ' Make sure the parent folder exists
      If FolderIDExists(lpParentId) = False Then
        Throw New FolderDoesNotExistException(lpParentId)
      End If

      ' Make sure the target folder does not exist
      If FolderExists(lpParentId, lpName, lstrExistingFolderId) Then
        Throw FolderAlreadyExistsException.Create(lpParentId, lpName, lstrExistingFolderId)
      End If

      ' Create the folder...
      Dim lobjParentFolder As FileNet.Api.Core.IFolder = ObjectStore.GetObject("Folder", lpParentId)
      Dim lobjNewFolder As FileNet.Api.Core.IFolder = lobjParentFolder.CreateSubFolder(lpName)

      ' Set the permissions if we have them
      If lpPermissions IsNot Nothing AndAlso lpPermissions.Count > 0 Then
        lobjNewFolder.Permissions = CreateP8PermissionList(lpPermissions)
      End If

      lobjNewFolder.Save(RefreshMode.REFRESH)

      If Not String.IsNullOrEmpty(lpClassName) Then
        lobjFolderClass = FolderClasses(lpClassName)
        If String.Compare(lobjNewFolder.GetClassName(), lpClassName, True) <> 0 Then
          lobjNewFolder.ChangeClass(lpClassName)
          lobjNewFolder.Save(RefreshMode.REFRESH)
        End If
      Else
        lobjFolderClass = FolderClasses("Folder")
      End If

      ' TODO: Set the folder properties...
      If lpProperties IsNot Nothing Then
        For Each lobjProperty As DCore.IProperty In lpProperties
          If lobjFolderClass.Properties.PropertyExists(lobjProperty.Name) Then
            SetPropertyValue(lobjNewFolder, lobjFolderClass, lobjProperty, False)
          Else
            Throw New InvalidPropertyException(
              String.Format("The property '{0}' is not valid for the folder class '{1}'.",
                            lobjProperty.Name, lobjFolderClass.Name), lobjProperty)
          End If
        Next
      End If

      ' TODO: Set the folder permissions...

      lobjNewFolder.Save(RefreshMode.REFRESH)

      Dim lobjReturnFolder As DProviders.IFolder = GetFolderInfoByID(lobjNewFolder.Id.ToString, 0)

      If lobjReturnFolder IsNot Nothing Then
        Return lobjReturnFolder
      Else
        Return Nothing
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Sub DeleteFolderByID(lpId As String) Implements IFolderManager.DeleteFolderByID
    Dim lobjFolder As CENetFolder = Nothing
    Try
      lobjFolder = GetFolderInfoByID(lpId, 0)
      lobjFolder.Delete()
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    Finally
      lobjFolder = Nothing
    End Try
  End Sub

  Public Sub DeleteFolderByPath(lpFolderPath As String) Implements IFolderManager.DeleteFolderByPath
    Dim lobjFolder As CENetFolder = Nothing
    Try
      lobjFolder = GetFolderInfoByPath(lpFolderPath, 0)
      lobjFolder.Delete()
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    Finally
      lobjFolder = Nothing
    End Try
  End Sub

  Public Sub FileDocumentByID(lpDocumentID As String, lpFolderID As String) Implements IFolderManager.FileDocumentByID
    Try
      If DocumentExists(lpDocumentID) Then
        If FolderPathExists(lpFolderID) Then
          FileDocument(lpDocumentID, lpFolderID)
        Else
          Throw New FolderDoesNotExistException(lpFolderID)
        End If
      Else
        Throw New DocumentDoesNotExistException(lpDocumentID)
      End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub FileDocumentByPath(lpDocumentID As String, lpFolderPath As String) Implements IFolderManager.FileDocumentByPath
    Try
      If DocumentExists(lpDocumentID) Then
        If FolderPathExists(lpFolderPath) Then
          FileDocument(lpDocumentID, lpFolderPath)
        Else
          Throw New FolderDoesNotExistException(lpFolderPath)
        End If
      Else
        Throw New DocumentDoesNotExistException(lpDocumentID)
      End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub
  Public Function FolderPathExists(lpFolderPath As String) As Boolean Implements IFolderManager.FolderPathExists
    Try

      Dim lobjFolder As FileNet.Api.Core.IFolder = Nothing

      lobjFolder = GetObject("Folder", lpFolderPath)
      lobjFolder.Refresh()

      If lobjFolder IsNot Nothing Then
        Return True
      Else
        Return False
      End If

    Catch ex As Exception
      Return False
    End Try
  End Function

  Public Function FolderIDExists(lpId As String) As Boolean Implements IFolderManager.FolderIDExists
    Try

      Dim lobjSearch As New CENetSearch(Me)
      Dim lobjIdCriterion As New DSearch.Criterion("Id", "object_id") With {
        .Value = lpId
      }
      lobjSearch.DataSource.QueryTarget = "Folder"
      lobjSearch.Criteria.Add(lobjIdCriterion)

      Dim lobjSearchResultSet As SearchResultSet = lobjSearch.Execute

      If lobjSearchResultSet.Count > 0 Then
        Return True

      Else
        Return False
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Overloads Function FolderExists(lpParentId As String, lpName As String, ByRef lpExistingFolderId As String) As Boolean Implements IFolderManager.FolderExists
    Try
      Return MyBase.FolderExists(Me, lpParentId, lpName, lpExistingFolderId)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function GetFolderIdByPath(lpFolderPath As String) As String Implements IFolderManager.GetFolderIdByPath
    Try
      If FolderPathExists(lpFolderPath) Then
        Dim lobjP8Folder As FileNet.Api.Core.IFolder = GetFolderByPath(lpFolderPath)
        Return lobjP8Folder.Id.ToString
      Else
        Throw New FolderDoesNotExistException(lpFolderPath)
      End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function GetFolderInfoByID(lpId As String, lpMaxContentCount As Integer) As DProviders.IFolder Implements IFolderManager.GetFolderInfoByID
    Try
      If FolderIDExists(lpId) Then
        Return GetFolderByID(lpId, 1, lpMaxContentCount)
      Else
        Throw New FolderDoesNotExistException(lpId)
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function GetFolderInfoByPath(lpFolderPath As String, lpMaxContentCount As Integer) As DProviders.IFolder Implements IFolderManager.GetFolderInfoByPath
    Try
      If FolderPathExists(lpFolderPath) Then
        Dim lobjP8Folder As FileNet.Api.Core.IFolder = GetFolderByPath(lpFolderPath)
        Return New CENetFolder(lobjP8Folder, lpMaxContentCount)
      Else
        Throw New FolderDoesNotExistException(lpFolderPath)
      End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Sub MoveDocumentByID(lpDocumentID As String, lpCurrentFolderID As String, lpDestinationFolderID As String) Implements IFolderManager.MoveDocumentByID
    Try

      If DocumentExists(lpDocumentID) = False Then
        Throw New DocumentDoesNotExistException(lpDocumentID)
      End If

      If FolderIDExists(lpCurrentFolderID) = False Then
        Throw New FolderDoesNotExistException(lpCurrentFolderID,
          String.Format("Current folder '{0}' does not exist.", lpCurrentFolderID))
      End If

      If FolderIDExists(lpDestinationFolderID) = False Then
        Throw New FolderDoesNotExistException(lpDestinationFolderID,
          String.Format("New folder '{0}' does not exist.", lpDestinationFolderID))
      End If

      UnFileDocument(lpDocumentID, lpCurrentFolderID)

      FileDocument(lpDocumentID, lpDestinationFolderID)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub MoveDocumentByPath(lpDocumentID As String, lpCurrentFolderPath As String, lpDestinationFolderPath As String) Implements IFolderManager.MoveDocumentByPath
    Try

      If DocumentExists(lpDocumentID) = False Then
        Throw New DocumentDoesNotExistException(lpDocumentID)
      End If

      If FolderIDExists(lpCurrentFolderPath) = False Then
        Throw New FolderDoesNotExistException(lpCurrentFolderPath,
          String.Format("Current folder '{0}' does not exist.", lpCurrentFolderPath))
      End If

      If FolderIDExists(lpDestinationFolderPath) = False Then
        Throw New FolderDoesNotExistException(lpDestinationFolderPath,
          String.Format("New folder '{0}' does not exist.", lpDestinationFolderPath))
      End If

      UnFileDocument(lpDocumentID, lpCurrentFolderPath)

      FileDocument(lpDocumentID, lpDestinationFolderPath)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub UnfileDocumentByID(lpDocumentID As String, lpFolderID As String) Implements IFolderManager.UnfileDocumentByID
    Try
      If DocumentExists(lpDocumentID) Then
        If FolderPathExists(lpFolderID) Then
          UnFileDocument(lpDocumentID, lpFolderID)
        Else
          Throw New FolderDoesNotExistException(lpFolderID)
        End If
      Else
        Throw New DocumentDoesNotExistException(lpDocumentID)
      End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub UnfileDocumentByPath(lpDocumentID As String, lpFolderPath As String) Implements IFolderManager.UnfileDocumentByPath
    Try
      If DocumentExists(lpDocumentID) Then
        If FolderPathExists(lpFolderPath) Then
          UnFileDocument(lpDocumentID, lpFolderPath)
        Else
          Throw New FolderDoesNotExistException(lpFolderPath)
        End If
      Else
        Throw New DocumentDoesNotExistException(lpDocumentID)
      End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

End Class

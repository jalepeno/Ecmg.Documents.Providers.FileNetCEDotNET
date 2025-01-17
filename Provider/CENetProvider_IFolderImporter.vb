'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  CENetProvider_IFolderImporter.vb
'   Description :  [type_description_here]
'   Created     :  3/7/2015 10:40:23 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports FileNet.Api.Constants
Imports FileNet.Api.Core
Imports Documents
Imports Documents.Core
Imports Documents.Exceptions
Imports Documents.Providers
Imports Documents.Utilities

#End Region

Partial Public Class CENetProvider
  Implements IFolderImporter

#Region "Public Methods"

  Public Function ImportFolder(lpArgs As Arguments.ImportFolderArgs) As Boolean Implements IFolderImporter.ImportFolder

    'Dim lblnResult As Boolean = False

    Try

      If lpArgs Is Nothing Then Throw New ArgumentNullException("lpArgs")

      Dim lobjFolder As Folder = lpArgs.Folder
      Dim lobjP8Folder As FileNet.Api.Core.IFolder = Nothing

      If lobjFolder Is Nothing Then Throw New InvalidOperationException("Folder reference not set")

      Dim lstrLastModifer As String = Me.UserName
      Dim ldatDateLastModified As Date = Now
      Dim ldatDateCheckedIn As Date = Now

      If HasElevatedPrivileges Then

        ' If we have last modifier and date last modified in the supplied document let's try to use it
        If lobjFolder.Properties.PropertyExists(PROP_MODIFY_USER) Then

          Dim lobjLastModifierProperty As ECMProperty = lobjFolder.Properties(PROP_MODIFY_USER)

          If lobjLastModifierProperty.HasValue Then
            lstrLastModifer = lobjLastModifierProperty.Value
          End If

        End If

        If lobjFolder.Properties.PropertyExists(PROP_MODIFY_DATE) Then

          Dim lobjLastModifiedProperty As ECMProperty = lobjFolder.Properties(PROP_MODIFY_DATE)

          If lobjLastModifiedProperty.HasValue Then
            ldatDateLastModified = lobjLastModifiedProperty.Value
          End If

        End If

        'If lobjFolder.Properties.PropertyExists(PROP_CHECK_IN_DATE) Then

        '  Dim lobjDateCheckedInProperty As ECMProperty = lobjFolder.Properties(PROP_CHECK_IN_DATE)

        '  If lobjDateCheckedInProperty.HasValue Then
        '    ldatDateCheckedIn = lobjDateCheckedInProperty.Value
        '  End If

        'End If

      End If

      ' Get the candidate document class
      Dim lobjCandidateClass As FolderClass = FolderClass(lobjFolder.FolderClass)
      Dim lobjProposedTree As New FolderTree(lpArgs.Folder.Path)
      Dim lobjAvailableParentFolder As FileNet.Api.Core.IFolder = GetAvailableParentFolderByPath(lpArgs.Folder.Path)

      If lobjAvailableParentFolder IsNot Nothing Then
        ' Make sure that the available parent is an immediate parent
        Dim lobjAvailableParentTree As New FolderTree(lobjAvailableParentFolder.PathName)
        If lobjProposedTree.Folders.Count - lobjAvailableParentTree.Folders.Count <> 1 Then
          If lobjProposedTree.Folders.Count - lobjAvailableParentTree.Folders.Count = 0 Then
            Dim lstrExistsErrorMessage As String = String.Format("Unable to create folder '{0}', the folder already exists.",
                                                                 lobjProposedTree.FolderPath)
            Throw New FolderAlreadyExistsException(lstrExistsErrorMessage)
          End If
          ' This folder is not an immediate parent
          Dim lobjExpectedFolderLevel As FolderTree.FolderInfo = lobjProposedTree.Folders(lobjProposedTree.Folders.Count - 1)
          Dim lstrErrorMessage As String = String.Format("Unable to create folder '{0}', the expected parent folder '{1}' does not exist.",
                                                         lobjProposedTree.FolderPath, lobjExpectedFolderLevel.Path)
          Throw New FolderDoesNotExistException(lobjExpectedFolderLevel.Path, lstrErrorMessage)
        End If
      End If

      lobjP8Folder = Factory.Folder.CreateInstance(mobjObjectStore, lobjCandidateClass.Name)
      lobjP8Folder.Parent = lobjAvailableParentFolder

      ' Set the properties needed on create
      For Each lobjCDFVersionProperty As ECMProperty In lobjFolder.Properties

        If lobjCDFVersionProperty.HasValue Then
          SetPropertyValue(lobjP8Folder, lobjCandidateClass, lobjCDFVersionProperty, True)
        End If

      Next

      ' Set the permissions if we have them and we were asked to.
      'If lpArgs.SetPermissions AndAlso lobjP8Folder.Permissions IsNot Nothing AndAlso lobjP8Folder.Permissions.Count > 0 Then
      '  lobjP8Folder.Permissions = CreateP8PermissionList(lobjP8Folder.Permissions)
      'End If

      If HasElevatedPrivileges Then
        lobjP8Folder.Properties.RemoveFromCache(PropertyNames.LAST_MODIFIER)
        lobjP8Folder.Properties.RemoveFromCache(PropertyNames.DATE_LAST_MODIFIED)
        lobjP8Folder.LastModifier = lstrLastModifer
        lobjP8Folder.DateLastModified = ldatDateLastModified
        'lobjP8Folder.DateCheckedIn = ldatDateCheckedIn
      End If

      ' lobjP8Folder.Save(RefreshMode.REFRESH)
      lobjP8Folder.Save(RefreshMode.NO_REFRESH)

      'lpArgs.Folder.Id = lobjP8Folder.Id.ToString()
      'lpArgs.Folder.Id = lobjP8Folder.PathName
      lpArgs.Folder.Id = lpArgs.Folder.Path

      Return True

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

End Class

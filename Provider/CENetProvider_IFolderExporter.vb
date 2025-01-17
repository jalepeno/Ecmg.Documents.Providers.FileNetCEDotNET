'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  CENetProvider_IFolderExporter.vb
'   Description :  [type_description_here]
'   Created     :  3/6/2015 9:41:00 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Arguments
Imports Documents.Core
Imports Documents.Exceptions
Imports Documents.Providers
Imports Documents.Utilities
Imports Documents

#End Region

Partial Public Class CENetProvider
  Implements IFolderExporter


#Region "Events"

  Public Event FolderExportError(ByVal sender As Object, ByVal e As Arguments.FolderExportErrorEventArgs) Implements IFolderExporter.FolderExportError

#End Region

#Region "IFolderExporter Implementation"

  Public Function ExportFolder(Args As ExportFolderEventArgs) As Boolean Implements IFolderExporter.ExportFolder
    Try
      Dim lobjFolder As FileNet.Api.Core.IFolder = Nothing
      Dim lobjReturnFolder As New Folder
      Dim lobjECMProperty As ECMProperty

      ' See if we can get the existing folder
      If FolderPathExists(Args.PrimaryIdentifier) Then
        lobjFolder = GetObject("Folder", Args.PrimaryIdentifier)
        lobjFolder.Refresh()

        If lobjFolder Is Nothing Then
          Throw New FolderDoesNotExistException(Args.PrimaryIdentifier)
        End If
      Else
        Throw New FolderDoesNotExistException(Args.PrimaryIdentifier)
      End If

      lobjReturnFolder.Id = lobjFolder.Id.ToString()
      lobjReturnFolder.FolderClass = lobjFolder.ClassDescription.SymbolicName

      Dim lobjPathProperty As ECMProperty = PropertyFactory.Create(PropertyType.ecmString, "Path", lobjFolder.PathName)
      lobjReturnFolder.Properties.Add(lobjPathProperty)

      'lobjReturnFolder.Path = lobjPathProperty.Value

      For Each lobjIProperty As FileNet.Api.Property.IProperty In lobjFolder.Properties
        lobjECMProperty = CreateECMProperty(lobjIProperty)

        If (lobjECMProperty IsNot Nothing) Then
          If lobjReturnFolder.Properties.PropertyExists(lobjECMProperty.Name) = False Then
            lobjReturnFolder.Properties.Add(lobjECMProperty)
            'If lobjECMProperty.Name = "Name" Then
            '  lobjReturnFolder.Name = lobjECMProperty.Value
            'End If
          End If
        End If
      Next

      lobjReturnFolder.Properties.Sort()

      If Args.GetPermissions Then
        ' Get all the permissions for this folder.
        lobjReturnFolder.Permissions.AddRange(GetCtsPermissions(lobjFolder))
      End If

      Args.Folder = lobjReturnFolder

      ' Return ExportFolderComplete

      Return True

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
      RaiseEvent FolderExportError(Me, New FolderExportErrorEventArgs(Args, ex))
    End Try
  End Function

#End Region

End Class

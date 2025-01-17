'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports FileNet.Api.Core
Imports Documents.Core
Imports Documents.Providers
Imports Documents.Utilities
Imports FileNet.Api.Util
Imports IDFolder = Documents.Providers.IFolder

#End Region

#Region "Explorer Classes"

Public Class CENetFolder
  Inherits CFolder

#Region "Class Variables"

  Private mobjFolderContents As New FolderContents

  Private mobjObjectStore As IObjectStore = Nothing
  Private mobjP8Folder As FileNet.Api.Core.IFolder

#End Region

#Region "Constructors"

  Public Sub New()
    MyBase.New()
  End Sub

  Public Sub New(ByVal lpFolderPath As String,
                 ByVal lpMaxContentCount As Long)
    MyBase.New(lpFolderPath, lpMaxContentCount)

    Try
      ' <Modified by: Ernie at 2/17/2012-2:08:27 PM on machine: ERNIE-M4400>
      'mobjP8Folder = GetFolderByPath(lpFolderPath, lpMaxContentCount)
      ' The method above was returning a CTS folder object, we need a P8 folder object here.
      'mobjP8Folder = CType(Provider, CENetProvider).GetFolderByPath(lpFolderPath)
      ' </Modified by: Ernie at 2/17/2012-2:08:27 PM on machine: ERNIE-M4400>

      MyBase.InitializeFolderCollection(lpFolderPath)
      InitializeFolder()

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Sub

  Public Sub New(ByVal lpP8IFolder As FileNet.Api.Core.IFolder,
                 ByVal lpMaxContentCount As Long)

    Try
      mobjP8Folder = lpP8IFolder
      Name = lpP8IFolder.Name
      Id = lpP8IFolder.Id.ToString
      MaxContentCount = lpMaxContentCount
      MyBase.InitializeFolderCollection(lpP8IFolder.PathName)
      InitializeFolder()

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Throw New Exception("Unable to create folder from IFolder Object", ex)
    End Try

  End Sub

  Public Sub New(ByVal lpFolderPath As String,
                 ByRef lpProvider As IProvider,
                 ByVal lpMaxContentCount As Long)

    MyBase.New(lpFolderPath, lpMaxContentCount)

    Try
      Provider = lpProvider
      MyBase.InitializeFolderCollection(lpFolderPath)
      InitializeFolder()
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub New(ByVal lpP8IFolder As FileNet.Api.Core.IFolder,
                 ByRef lpProvider As IProvider,
                 ByVal lpMaxContentCount As Long)

    Try
      mobjP8Folder = lpP8IFolder
      Name = lpP8IFolder.Name
      Id = lpP8IFolder.Id.ToString
      Provider = lpProvider
      MaxContentCount = lpMaxContentCount
      MyBase.InitializeFolderCollection(lpP8IFolder.PathName)
      InitializeFolder()

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Throw New Exception("Unable to create folder from IFolder Object", ex)
    End Try

  End Sub

#End Region

#Region "Public Properties"

  Public Overridable ReadOnly Property ClassDescription() As String
    Get
      Try
        If Me.Properties.PropertyExists("Class Description") = True Then
          Return Properties("Class Description").Value.ToString
        Else
          Return "UNKNOWN"
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

  Public Overrides ReadOnly Property Contents() As FolderContents
    Get
      Return mobjFolderContents
    End Get
  End Property

  Public Overrides ReadOnly Property SubFolders() As Folders
    Get

      Try
        Return GetSubFolders(True)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Get
  End Property

  Private ReadOnly Property ObjectStore() As IObjectStore
    Get

      Try
        Return CType(CType(Provider, CENetProvider).ObjectStore, IObjectStore)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return Nothing
      End Try

    End Get
  End Property

#End Region

#Region "Public Methods"

  Public Sub Delete()
    Try
      If mobjP8Folder IsNot Nothing Then
        mobjP8Folder.Delete()
        mobjP8Folder.Save(FileNet.Api.Constants.RefreshMode.REFRESH)
      End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Overrides Sub Refresh()

    Try
      InitializeFolder()

      'Throw New NotImplementedException
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Sub

#End Region

#Region "Public Overrides Methods"

  Public Overrides Function GetID() As String

    Try

      If Me.Path.Length = 0 Then
        Throw New Exception("Could not get folder id, no path is available.")
      End If

      Dim lobjIFolder As FileNet.Api.Core.IFolder = CType(Provider, CENetProvider).GetFolderByPath(Me.Path)

      If lobjIFolder Is Nothing Then
        Throw New Exception("Could not get folder object.")
      End If

      'Return lobjIFolder.Id.ToString
      Return lobjIFolder.PathName

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  Protected Overrides Function GetFolderByPath(ByVal lpFolderPath As String,
                                               ByVal lpMaxContentCount As Long) As IDFolder

    Try

      Dim lobjP8IFolder As FileNet.Api.Core.IFolder = CType(Provider, CENetProvider).GetFolderByPath(lpFolderPath)

      If lobjP8IFolder Is Nothing Then
        Return Nothing

      Else
        Return New CENetFolder(lobjP8IFolder, lpMaxContentCount)
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  Protected Overrides Function GetPath() As String

    Try

      If Not mobjP8Folder Is Nothing Then
        Return mobjP8Folder.PathName

      Else
        Return String.Empty
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  Protected Overrides Function GetSubFolderCount() As Long
    Static lintSubFolderCount As Integer = -1

    Try

      If lintSubFolderCount = -1 Then

        If Not mobjP8Folder Is Nothing Then
          lintSubFolderCount = GetFolderCount(mobjP8Folder.SubFolders)

        Else
          lintSubFolderCount = GetFolderCount(ObjectStore.TopFolders)
        End If

        Return lintSubFolderCount

      Else
        Return lintSubFolderCount
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  Public Overrides Function GetSubFolders(ByVal lpGetContents As Boolean) As Folders

    Dim lobjFolders As New Folders
    Dim lobjFolder As CFolder

    Try

      If mobjP8Folder Is Nothing Then

        For Each lobjIFolder As FileNet.Api.Core.IFolder In ObjectStore.TopFolders
          lobjFolder = New CENetFolder(lobjIFolder, Provider, MaxContentCount)
          lobjFolders.Add(lobjFolder)
        Next

        lobjFolders.Sort()
        Return lobjFolders

      Else

        For Each lobjIFolder As FileNet.Api.Core.IFolder In mobjP8Folder.SubFolders
          lobjFolder = New CENetFolder(lobjIFolder, Provider, MaxContentCount)
          lobjFolders.Add(lobjFolder)
        Next

        lobjFolders.Sort()
        Return lobjFolders

      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Return Nothing
    End Try

  End Function

  Private Function GetFolderCount(ByVal lpFolders As FileNet.Api.Collection.IFolderSet) As Integer

    Try

      Dim lintFolderCounter As Integer = 0

      For Each lobjIFolder As FileNet.Api.Core.IFolder In lpFolders
        lintFolderCounter += 1
      Next

      Return lintFolderCounter

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  Protected Overrides Sub InitializeFolder()

    Try

      If mobjP8Folder Is Nothing Then

        If Me.Path.Length > 0 Then
          mobjP8Folder = CType(Provider, CENetProvider).GetFolderByPath(Me.Path)

        Else
          ' Now what?
        End If

      End If

      If MaxContentCount = -1 Then
        InitializeFolderContents()
      End If

      ' Get the folder properties
      InitializeFolderProperties()

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Throw New Exception("Unable to initialize folder", ex)
    End Try

  End Sub

#End Region

  Private Sub InitializeFolderContents()

    mobjFolderContents = New FolderContents

    Try

      If Not mobjP8Folder Is Nothing Then

        Dim lobjFolderContent As FolderContent

        For Each lobjIDocument As IDocument In mobjP8Folder.ContainedDocuments

          If lobjIDocument.ContentElements.Count > 0 Then
            lobjFolderContent = New FolderContent(lobjIDocument.Name, lobjIDocument.Id.ToString, CLng(lobjIDocument.ContentSize), lobjIDocument.ContentElements(0).RetrievalName, CDate(lobjIDocument.DateLastModified))

          Else

            If lobjIDocument.ContentSize Is Nothing Then
              lobjFolderContent = New FolderContent(lobjIDocument.Name, lobjIDocument.Id.ToString, 0, String.Empty, CDate(lobjIDocument.DateLastModified))

            Else
              lobjFolderContent = New FolderContent(lobjIDocument.Name, lobjIDocument.Id.ToString, CLng(lobjIDocument.ContentSize), String.Empty, CDate(lobjIDocument.DateLastModified))

            End If

          End If

          mobjFolderContents.Add(lobjFolderContent)
        Next

      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Throw New Exception("Unable to Initialize Folder Contents", ex)
    End Try

  End Sub

  Private Sub InitializeFolderProperties()
    Try

      If mobjP8Folder IsNot Nothing Then

        Dim lobjP8Provider As CENetProvider = Me.Provider
        Dim lobjCtsProperty As IProperty = Nothing

        For Each lobjP8FolderProperty As FileNet.Api.Property.IProperty In mobjP8Folder.Properties
          lobjCtsProperty = lobjP8Provider.CreateECMProperty(lobjP8FolderProperty)
          If (lobjCtsProperty IsNot Nothing) AndAlso (Me.Properties.Contains(lobjCtsProperty) = False) Then
            Me.Properties.Add(lobjCtsProperty)
          End If
        Next

        ' Get the audit events
        Me.AuditEvents = lobjP8Provider.GetAuditEvents(mobjP8Folder)

        ' Get all the permissions for this folder.
        Me.Permissions.AddRange(lobjP8Provider.GetCtsPermissions(mobjP8Folder))

      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub


End Class

#End Region

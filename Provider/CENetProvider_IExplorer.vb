'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  CENetProvider_IExplorer.vb
'   Description :  [type_description_here]
'   Created     :  4/8/2014 10:26:55 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Providers
Imports Documents.Utilities
Imports DProviders = Documents.Providers
Imports DCore = Documents.Core

#End Region

Partial Public Class CENetProvider
  Implements IExplorer

#Region "IExplorer Implementation"

  Public ReadOnly Property RootFolder() As DProviders.IFolder _
                Implements IExplorer.RootFolder
    Get

      Try
        InitializeRootFolder()
        Return mobjFolder

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Get
  End Property

  Public Overrides ReadOnly Property Search() As DProviders.ISearch _
                            Implements IExplorer.Search
    Get

      Try

        If mobjSearch Is Nothing Then
          mobjSearch = New CENetSearch(Me)
        End If

        Return mobjSearch

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Get
  End Property

  Public ReadOnly Property GetFolderByID(ByVal lpFolderID As String,
                                       ByVal lpFolderLevels As Integer,
                                       ByVal lpMaxContentCount As Integer) As IFolder _
                Implements IExplorer.GetFolderByID
    Get

      Try
        'Throw New NotImplementedException("Coming Soon...")
        'Dim lobjFolder As FileNet.Api.Core.IFolder = Nothing
        'Dim lobjFolderIdentifier As New Id(lpFolderID)

        'lobjFolder = ObjectStore.GetObject("Folder", lobjFolderIdentifier)
        'Return New CENetFolder(lobjFolder, Me, lpMaxContentCount)

        Return New CENetFolder(GetFolderByPath(lpFolderID), Me, lpMaxContentCount)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Get
  End Property

  Public ReadOnly Property GetFolderContentsByID(ByVal lpFolderID As String,
                                                 ByVal lpMaxContentCount As Integer) As DCore.FolderContents _
                  Implements IExplorer.GetFolderContentsByID
    Get

      Try
        'Throw New NotImplementedException("Coming Soon...")
        Return GetFolderByID(lpFolderID, 1, lpMaxContentCount).Contents

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Get
  End Property

  Public ReadOnly Property HasSubFolders(ByVal lpFolderPath As String) As Boolean _
                  Implements IExplorer.HasSubFolders
    Get

      Try
        Throw New NotImplementedException("Coming Soon...")

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Get
  End Property

  Public ReadOnly Property IsFolderValid(ByVal lpFolderPath As String) As Boolean _
                  Implements IExplorer.IsFolderValid
    Get

      Try
        Throw New NotImplementedException("Coming Soon...")

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Get
  End Property

#End Region

End Class

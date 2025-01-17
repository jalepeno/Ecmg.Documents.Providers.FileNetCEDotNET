'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  CENetProvider_IDelete.vb
'   Description :  [type_description_here]
'   Created     :  4/8/2014 11:10:31 AM
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
Imports Documents.Exceptions
Imports Documents.Providers
Imports Documents.Utilities
Imports Documents.Core
Imports DCore = Documents.Core

#End Region

Partial Public Class CENetProvider
  Implements IDelete


#Region "IDelete Implementation"

  Public Function DeleteVersion(ByVal lpDocumentId As String,
                                ByVal lpCriterion As String,
                                 ByVal lpSessionId As String) As Boolean Implements IDelete.DeleteVersion
    Try

      ' Create a return value variable
      Dim lblnReturnValue As Boolean = False

      If ObjectStore Is Nothing Then
        InitializeConnection()
      End If

      Dim lobjTargetDocument As IDocument = GetIDocument(lpDocumentId)
      Dim lobjTargetVersion As IDocument = Nothing
      Dim lblnSuccess As Boolean

      If lobjTargetDocument IsNot Nothing Then

        Dim lobjVersionSeries As IVersionSeries = lobjTargetDocument.VersionSeries

        Dim lobjVersionDictionary As New SortedDictionary(Of String, IDocument)

        Dim lobjVersions As IVersionableSet = lobjVersionSeries.Versions
        Dim lobjVersionIdentifier As DCore.IVersionIdentifier = Nothing
        Dim lstrVersion As String = Nothing

        Select Case lpCriterion.ToUpper
          Case "ALL AFTER FIRST"
            Dim lintVersionCounter As Integer = 0
            Dim lintDeletedCounter As Integer = 0

            For Each lobjVersion As IVersionable In lobjVersionSeries.Versions
              lobjVersionIdentifier = New VersionIdentifier(lobjVersion.MajorVersionNumber, lobjVersion.MinorVersionNumber)
              lstrVersion = lobjVersionIdentifier.Version.ToString
              lobjVersionDictionary.Add(lstrVersion, lobjVersion)
              'lintVersionCounter += 1
              'If lintVersionCounter > 1 Then
              '  lobjVersionDictionary.Add(lstrVersion, lobjVersion)
              'End If
            Next

            For Each lobjDocumentVersion As IDocument In lobjVersionDictionary.Values
              lintVersionCounter += 1
              If lintVersionCounter > 1 Then
                lobjDocumentVersion.Delete()
                lobjDocumentVersion.Save(RefreshMode.NO_REFRESH)
                lintDeletedCounter += 1
              End If
            Next

            ' TODO: Figure out a way to reverse the stack and delete the versions in the required order
            If lintDeletedCounter = lobjVersionDictionary.Count - 1 Then
              lblnSuccess = True
            End If
        End Select

        ' TODO: Delete the right version
        'For Each lobjVersion As IVersionable In lobjVersions
        '  If lobjVersion.MajorVersionNumber = lpVersionId.MajorVersion AndAlso lobjVersion.MinorVersionNumber = lpVersionId.MinorVersion Then
        '    lobjTargetVersion = lobjVersion
        '    Exit For
        '  End If
        'Next

        'If lobjTargetVersion IsNot Nothing Then
        '  lobjTargetVersion.Delete()
        '  lobjTargetVersion.Save(RefreshMode.NO_REFRESH)
        '  Return True
        'Else
        '  Return False
        'End If

      End If

      Return lblnSuccess

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function DeleteVersion(ByVal lpVersion As IDocument) As Boolean
    Try
      If lpVersion IsNot Nothing Then
        lpVersion.Delete()
        lpVersion.Save(RefreshMode.NO_REFRESH)
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

  Public Function DeleteDocument(ByVal lpId As String,
                                 ByVal lpSessionId As String) As Boolean _
         Implements IDelete.DeleteDocument
    Try
      Return DeleteDocument(lpId, lpSessionId, False)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function DeleteDocument(lpId As String, lpSessionId As String, lpDeleteAllVersions As Boolean) As Boolean _
    Implements IDelete.DeleteDocument
    Try

      ' Create a return value variable
      Dim lblnReturnValue As Boolean = False

      If ObjectStore Is Nothing Then
        InitializeConnection()
      End If

      Dim lobjTargetDocument As IDocument = GetIDocument(lpId)

      ' <Modified by: Ernie at 1/11/2013-4:15:37 PM on machine: ERNIE-THINK>
      If lpDeleteAllVersions = True Then
        If lobjTargetDocument IsNot Nothing Then
          lobjTargetDocument.VersionSeries.Delete()
          lobjTargetDocument.VersionSeries.Save(RefreshMode.NO_REFRESH)
          lblnReturnValue = True
        Else
          Throw New DocumentDoesNotExistException(lpId)
        End If
      Else
        If lobjTargetDocument IsNot Nothing Then
          lobjTargetDocument.Delete()
          lobjTargetDocument.Save(RefreshMode.NO_REFRESH)
          lblnReturnValue = True
        Else
          Throw New DocumentDoesNotExistException(lpId)
        End If
      End If
      ' </Modified by: Ernie at 1/11/2013-4:15:37 PM on machine: ERNIE-THINK>

      Return lblnReturnValue

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

End Class

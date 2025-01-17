'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  CENetProvider_IBasicContentServicesProvider.vb
'   Description :  [type_description_here]
'   Created     :  4/8/2014 11:04:23 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports Ecmg.Cts.Utilities
Imports Ecmg.Cts.Core
Imports Ecmg.Cts.Arguments
Imports FileNet.Api.Core
Imports Ecmg.Cts.Exceptions
Imports FileNet.Api.Collection
Imports FileNet.Api.Property
Imports Ecmg.Cts.Migrations
Imports FileNet.Api.Constants
Imports System.IO
Imports System.Reflection
Imports Documents.Arguments
Imports Documents.Core
Imports Documents.Exceptions
Imports Documents.Providers
Imports Documents.Utilities
Imports CtsCore = Documents.Core
Imports CtsArguments = Documents.Arguments

#End Region

Partial Public Class CENetProvider
  Implements IBasicContentServicesProvider

#Region "IBasicContentServicesProvider Implementation"

#Region "Events"

  Public Event DocumentAdded(ByVal sender As Object, ByVal e As CtsArguments.DocumentAddedEventArgs) Implements IBasicContentServicesProvider.DocumentAdded

  Public Event DocumentCheckedIn(ByVal sender As Object, ByVal e As CtsArguments.DocumentCheckedInEventArgs) Implements IBasicContentServicesProvider.DocumentCheckedIn

  Public Event DocumentCheckedOut(ByVal sender As Object, ByVal e As CtsArguments.DocumentCheckedOutEventArgs) Implements IBasicContentServicesProvider.DocumentCheckedOut

  Public Event DocumentCheckOutCancelled(ByVal sender As Object, ByVal e As CtsArguments.DocumentCheckoutCancelledEventArgs) Implements IBasicContentServicesProvider.DocumentCheckOutCancelled

  Public Event DocumentCopiedOut(ByVal sender As Object, ByVal e As CtsArguments.DocumentCopiedOutEventArgs) Implements IBasicContentServicesProvider.DocumentCopiedOut

  Public Event DocumentDeleted(ByVal sender As Object, ByVal e As CtsArguments.DocumentDeletedEventArgs) Implements IBasicContentServicesProvider.DocumentDeleted

  Public Event DocumentEvent(ByVal sender As Object, ByVal e As CtsArguments.DocumentEventArgs) Implements IBasicContentServicesProvider.DocumentEvent

  Public Event DocumentFiled(ByVal sender As Object, ByVal e As CtsArguments.DocumentFiledEventArgs) Implements IBasicContentServicesProvider.DocumentFiled

  Public Event DocumentUnFiled(ByVal sender As Object, ByVal e As CtsArguments.DocumentUnFiledEventArgs) Implements IBasicContentServicesProvider.DocumentUnFiled

  Public Event DocumentUpdated(ByVal sender As Object, ByVal e As CtsArguments.DocumentUpdatedEventArgs) Implements IBasicContentServicesProvider.DocumentUpdated

#End Region

#Region "Methods"

  Public Function AddDocument(ByVal lpDocument As CtsCore.Document) As Boolean _
         Implements IBasicContentServicesProvider.AddDocument

    Try

      Return AddDocument(lpDocument, False)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  Public Function AddDocument(ByVal lpDocument As CtsCore.Document, lpAsMajorVersion As Boolean) As Boolean _
         Implements IBasicContentServicesProvider.AddDocument

    Try

      Dim lstrErrorMessage As String = String.Empty

      Dim lenuVersionType As VersionTypeEnum
      If lpAsMajorVersion = True Then
        lenuVersionType = VersionTypeEnum.Major
      Else
        lenuVersionType = VersionTypeEnum.Minor
      End If

      Dim lobjIDocument As IDocument = AddP8Document(lpDocument, lstrErrorMessage, lenuVersionType)

      ' Create a variable for the return value
      Dim lblnReturnValue As Boolean

      If lobjIDocument IsNot Nothing Then

        If String.IsNullOrEmpty(lstrErrorMessage) Then

          lpDocument.ObjectID = lobjIDocument.Id.ToString

          Dim lobjAddDocumentArgs As New DocumentAddedEventArgs(lpDocument, lobjIDocument.Id.ToString)
          RaiseEvent DocumentAdded(Me, lobjAddDocumentArgs)
          RaiseEvent DocumentEvent(Me, lobjAddDocumentArgs)
          lblnReturnValue = True

        Else
          Dim lobjDocumentException As New DocumentException(lpDocument, "0", lstrErrorMessage)
          If Not String.IsNullOrEmpty(lpDocument.ID) Then
            RaiseEvent DocumentImportError(Me,
              New DocumentImportErrorEventArgs(lpDocument.ID, lstrErrorMessage, lobjDocumentException))
          End If
          Throw lobjDocumentException
        End If

      Else
        Dim lobjImportErrorEventAgs As New DocumentImportErrorEventArgs(lstrErrorMessage, ExceptionTracker.LastException)
        RaiseEvent DocumentImportError(Me, lobjImportErrorEventAgs)
        lblnReturnValue = False
      End If

      If (lobjIDocument IsNot Nothing) Then
        lobjIDocument = Nothing
      End If

      Return lblnReturnValue

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function AddDocument(ByVal lpDocument As CtsCore.Document,
                                ByVal lpFolderPath As String) As Boolean _
         Implements IBasicContentServicesProvider.AddDocument

    Try
      Return AddDocument(lpDocument, lpFolderPath, False)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  Public Function AddDocument(ByVal lpDocument As CtsCore.Document,
                                ByVal lpFolderPath As String, lpAsMajorVersion As Boolean) As Boolean _
         Implements IBasicContentServicesProvider.AddDocument

    Try
      'If AddDocumentEx(lpDocument) = True Then
      '  ApplicationLogging.WriteLogEntry("AddDocument with folder path is not yet implemented.  The document was added as unfiled.")
      'End If

      ' Create a variable for the return value
      Dim lblnReturnValue As Boolean

      ' Create a variable to capture any errors associated with the add operation
      Dim lstrErrorMessage As String = String.Empty

      Dim lenuVersionType As VersionTypeEnum
      If lpAsMajorVersion = True Then
        lenuVersionType = VersionTypeEnum.Major
      Else
        lenuVersionType = VersionTypeEnum.Minor
      End If

      ' Attempt to add the document
      Dim lobjIDocument As IDocument = AddP8Document(lpDocument, lstrErrorMessage, lenuVersionType)

      If lobjIDocument IsNot Nothing Then

        If String.IsNullOrEmpty(lstrErrorMessage) Then
          'Dim lobjAddDocumentArgs As New DocumentAddedEventArgs(lpDocument, lobjIDocument.Id.ToString)
          'RaiseEvent DocumentAdded(Me, lobjAddDocumentArgs)
          'RaiseEvent DocumentEvent(Me, lobjAddDocumentArgs)
          lblnReturnValue = True

        Else
          Throw New DocumentException(lpDocument, "0", lstrErrorMessage)
        End If

      Else
        lblnReturnValue = False
      End If

      ' Check to see if there was an error adding the document
      If lblnReturnValue = False AndAlso lstrErrorMessage.Length > 0 Then
        Throw New DocumentException(lpDocument, "0", lstrErrorMessage)
      End If

      FileDocument(lobjIDocument, lpFolderPath, , lstrErrorMessage)

      If String.IsNullOrEmpty(lstrErrorMessage) Then

        Dim addDocumentArgs As New DocumentAddedEventArgs(New Document, lpDocument.ObjectID)
        RaiseEvent DocumentAdded(Me, addDocumentArgs)
        RaiseEvent DocumentEvent(Me, addDocumentArgs)

      Else
        RaiseEvent DocumentImportError(Me, New DocumentImportErrorEventArgs(lobjIDocument.Id.ToString, lstrErrorMessage,
                                                                            New DocumentException(lpDocument, lstrErrorMessage)))
      End If

      If (lobjIDocument IsNot Nothing) Then
        lobjIDocument = Nothing
      End If

      Return lblnReturnValue

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  Function IsCheckedOut(ByVal lpId As String) As Boolean Implements IBasicContentServicesProvider.IsCheckedOut

    Dim lobjDocument As IDocument

    Try

      lobjDocument = GetIDocument(lpId)

      Return lobjDocument.VersionSeries.IsReserved

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function CancelCheckoutDocument(ByVal lpId As String) As Boolean _
         Implements IBasicContentServicesProvider.CancelCheckoutDocument

    Dim lobjDocument As IDocument
    Dim lobjReservation As IDocument
    Dim lobjResult As IDocument

    Try

      lobjDocument = GetIDocument(lpId)

      If lobjDocument.VersionSeries.IsReserved Then
        lobjReservation = lobjDocument.VersionSeries.Reservation
        lobjResult = lobjReservation.CancelCheckout()
        lobjResult.Save(RefreshMode.NO_REFRESH)
        ' lobjReservation.Save(RefreshMode.NO_REFRESH)
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

  Public Function CheckinDocument(ByVal lpId As String,
                                  ByVal lpContentContainer As IContentContainer,
                                  ByVal lpAsMajorVersion As Boolean) As Boolean _
    Implements IBasicContentServicesProvider.CheckinDocument
    Try

      Return CheckinDocument(lpId, lpContentContainer, lpAsMajorVersion, Nothing)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Function CheckinDocument(ByVal lpId As String,
                        ByVal lpContentContainer As IContentContainer,
                        ByVal lpAsMajorVersion As Boolean,
                        ByVal lpProperties As CtsCore.IProperties) As Boolean _
         Implements IBasicContentServicesProvider.CheckinDocument
    Try

      Dim lobjDocument As IDocument = GetIDocument(lpId)

      Dim lobjReservation As IDocument = Nothing

      If lobjDocument IsNot Nothing Then
        If lobjDocument.Reservation IsNot Nothing Then
          lobjReservation = lobjDocument.Reservation

          If lpProperties IsNot Nothing AndAlso lpProperties.Count > 0 Then
            Dim lobjCandidateClass As DocumentClass = DocumentClass(lobjDocument.GetClassName)
            ' Set the properties
            For Each lobjCtsProperty As CtsCore.IProperty In lpProperties

              If lobjCtsProperty.HasValue Then
                SetPropertyValue(lobjReservation, lobjCandidateClass, lobjCtsProperty, False)
              End If

            Next
          End If


          If lpContentContainer.CanRead Then
            lobjReservation.ContentElements = CreateContentElementList(lpContentContainer)
            lobjReservation.Save(RefreshMode.NO_REFRESH)
          End If

          If lpAsMajorVersion = True Then
            lobjReservation.Checkin(AutoClassify.DO_NOT_AUTO_CLASSIFY, CheckinType.MAJOR_VERSION)
          Else
            lobjReservation.Checkin(AutoClassify.DO_NOT_AUTO_CLASSIFY, CheckinType.MINOR_VERSION)
          End If

          lobjReservation.Save(RefreshMode.NO_REFRESH)

          lobjReservation = Nothing
          lobjDocument = Nothing

        Else

          ' Throw New DocumentException(lpId, String.Format("Document {0} is not checked out,", lpId))
          Throw New DocumentNotCheckedOutException(lpId)

        End If

      Else

        Throw New DocumentDoesNotExistException(lpId)

      End If

      Return True

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function CheckinDocument(ByVal lpId As String,
                                  ByVal lpContentPath As String,
                                  ByVal lpAsMajorVersion As Boolean) As Boolean _
         Implements IBasicContentServicesProvider.CheckinDocument

    Try

      Return CheckinDocument(lpId, lpContentPath, lpAsMajorVersion, Nothing)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  Function CheckinDocument(ByVal lpId As String,
                          ByVal lpContentPath As String,
                          ByVal lpAsMajorVersion As Boolean,
                          ByVal lpProperties As CtsCore.IProperties) As Boolean _
          Implements IBasicContentServicesProvider.CheckinDocument
    Try

      Dim lobjDocument As IDocument = GetIDocument(lpId)
      'lobjDocument.Checkout(ReservationType.EXCLUSIVE, Nothing, Nothing, Nothing)
      Dim lobjReservation As IDocument = Nothing

      If lobjDocument IsNot Nothing Then
        If lobjDocument.Reservation IsNot Nothing Then
          lobjReservation = lobjDocument.Reservation

          Dim lobjContentContainer As New ContentFileContainer(lpContentPath)

          lobjReservation.ContentElements = CreateContentElementList(lobjContentContainer)

          If lpProperties IsNot Nothing AndAlso lpProperties.Count > 0 Then
            Dim lobjCandidateClass As DocumentClass = DocumentClass(lobjDocument.GetClassName)
            ' Set the properties
            For Each lobjCtsProperty As CtsCore.IProperty In lpProperties

              If lobjCtsProperty.HasValue Then
                SetPropertyValue(lobjReservation, lobjCandidateClass, lobjCtsProperty, False)
              End If

            Next
          End If

          If lpAsMajorVersion = True Then
            lobjReservation.Checkin(AutoClassify.DO_NOT_AUTO_CLASSIFY, CheckinType.MAJOR_VERSION)
          Else
            lobjReservation.Checkin(AutoClassify.DO_NOT_AUTO_CLASSIFY, CheckinType.MINOR_VERSION)
          End If

          lobjReservation.Save(RefreshMode.NO_REFRESH)

          lobjReservation = Nothing
          lobjDocument = Nothing
          lobjContentContainer.Dispose()

        Else
          Throw New DocumentException(lpId, String.Format("Document {0} is not checked out,", lpId))
        End If

      Else
        Throw New DocumentDoesNotExistException(lpId)
      End If

      Return True

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function CheckinDocument(ByVal lpId As String,
                                  ByVal lpContentPaths() As String,
                                  ByVal lpAsMajorVersion As Boolean) As Boolean _
         Implements IBasicContentServicesProvider.CheckinDocument

    Try
      Return CheckinDocument(lpId, lpContentPaths, lpAsMajorVersion, Nothing)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  Function CheckinDocument(ByVal lpId As String,
                        ByVal lpContentPaths As String(),
                        ByVal lpAsMajorVersion As Boolean,
                        ByVal lpProperties As CtsCore.IProperties) As Boolean _
         Implements IBasicContentServicesProvider.CheckinDocument
    Try

      Dim lobjDocument As IDocument = GetIDocument(lpId)
      'lobjDocument.Checkout(ReservationType.EXCLUSIVE, Nothing, Nothing, Nothing)
      Dim lobjReservation As IDocument = Nothing

      If lobjDocument IsNot Nothing Then
        If lobjDocument.Reservation IsNot Nothing Then
          lobjReservation = lobjDocument.Reservation

          Dim lobjContents As New Contents

          For Each lstrContentPath As String In lpContentPaths
            If Not String.IsNullOrEmpty(lstrContentPath) AndAlso File.Exists(lstrContentPath) Then
              lobjContents.Add(lstrContentPath)
            End If
          Next

          lobjReservation.ContentElements = CreateContentElementList(lobjContents)

          If lpProperties IsNot Nothing AndAlso lpProperties.Count > 0 Then
            Dim lobjCandidateClass As DocumentClass = DocumentClass(lobjDocument.GetClassName)
            ' Set the properties
            For Each lobjCtsProperty As CtsCore.IProperty In lpProperties

              If lobjCtsProperty.HasValue Then
                SetPropertyValue(lobjReservation, lobjCandidateClass, lobjCtsProperty, False)
              End If

            Next
          End If

          If lpAsMajorVersion = True Then
            lobjReservation.Checkin(AutoClassify.DO_NOT_AUTO_CLASSIFY, CheckinType.MAJOR_VERSION)
          Else
            lobjReservation.Checkin(AutoClassify.DO_NOT_AUTO_CLASSIFY, CheckinType.MINOR_VERSION)
          End If

          lobjReservation.Save(RefreshMode.NO_REFRESH)

          lobjReservation = Nothing
          lobjDocument = Nothing

        Else
          Throw New DocumentException(lpId, String.Format("Document {0} is not checked out,", lpId))
        End If

      Else
        Throw New DocumentDoesNotExistException(lpId)
      End If

      Return True

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function CheckoutDocument(ByVal lpId As String,
                                   ByVal lpDestinationFolder As String,
                                   ByRef lpOutputFileNames() As String) As Boolean _
         Implements IBasicContentServicesProvider.CheckoutDocument

    Try

      Dim lobjDocument As IDocument = GetIDocument(lpId)
      Dim lobjCurrentVersion As IDocument = lobjDocument.CurrentVersion

      If lobjCurrentVersion.IsReserved Then
        Throw New DocumentAlreadyCheckedOutException(lpId)
      End If

      lobjCurrentVersion.Checkout(ReservationType.EXCLUSIVE, Nothing, Nothing, Nothing)
      lobjCurrentVersion.Save(RefreshMode.NO_REFRESH)

      lobjDocument = Nothing
      lobjCurrentVersion = Nothing

      Return True

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  Public Function GetDocumentContents(lpId As String, lpVersionScope As VersionScopeEnum, lpMaxVersionCount As Integer) As Contents _
    Implements IBasicContentServicesProvider.GetDocumentContents

    'Dim lobjContents As New Contents()

    Try

      Throw New NotImplementedException

      '      Dim lobjIDocument As IDocument = GetIDocument(lpId)

      '      'TODO: Make sure this is the latest version
      '      Dim lobjILastestVersion As IDocument = Nothing

      '      Select Case lpVersionScope
      '        Case VersionScopeEnum.AllVersions
      '          ' Add some magic

      '        Case VersionScopeEnum.CurrentReleasedVersion

      '        Case VersionScopeEnum.MostCurrentVersion

      '        Case VersionScopeEnum.FirstVersion
      '          lobjILastestVersion = lobjIDocument.VersionSeries.Versions(0)

      '        Case VersionScopeEnum.FirstNVersions
      '          Throw New NotImplementedException

      '        Case VersionScopeEnum.LastNVersions
      '          Throw New NotImplementedException

      '      End Select




      '      Dim lobjContentElementList As IContentElementList
      '      Dim lobjContentTransfer As IContentTransfer
      '      lobjContentElementList = lobjILastestVersion.Properties("ContentElements")

      '      Dim lobjContentStream As System.IO.Stream
      '      Dim byteContent() As Byte
      '      Dim lintStreamLength As Integer

      '      Dim lintContentElement As Integer = 0

      '      For Each lobjContentTransfer In lobjContentElementList
      '        'Debug.Print(lobjContentTransfer.RetrievalName)
      '        lobjContentStream = lobjContentTransfer.AccessContentStream

      '        ' Changed by Ernie Bahr 3/17/2009
      '        ' Instead of writing out the file and then adding to the contents
      '        ' collection via the path, just add the stream.

      '        byteContent = New Byte(lobjContentStream.Length - 1) {}
      '        lintStreamLength = lobjContentStream.Read(byteContent, 0, lobjContentStream.Length)


      '        Dim lobjTempContent As Content = New Content(byteContent, lobjContentTransfer.RetrievalName, Content.StorageTypeEnum.Reference, False, AllowZeroLengthContent)
      '        lobjTempContent.MIMEType = lobjContentTransfer.ContentType

      '#If SupportAnnotations Then

      '        ' NOTE: The enumeration below actually iterates through starting with the newest 
      '        ' version and works towards the oldest.  We should first put the versions 
      '        ' into an array and then iterate the array in reverse.

      '        'For Each lobjNativeAnnotation As FileNet.Api.Core.IAnnotation In lobjILastestVersion.Annotations '  lobjNativeVersion.Annotations

      '        '  If lobjNativeAnnotation.AnnotatedContentElement <> lintContentElement Then Continue For

      '        '  Using lobjAnnotationContentStream As Stream = Me.RetrieveAnnotationContent(lobjNativeAnnotation)
      '        '    Dim lobjCtsAnnotation As Ecmg.Cts.Annotations.Annotation = lobjAnnotationExporter.ExportAnnotationObject(lobjAnnotationContentStream, lobjTempContent.MIMEType)
      '        '    ' lobjCtsAnnotation.AnnotatedContent = lobjTempContent
      '        '    lobjTempContent.Annotations.Add(lobjCtsAnnotation)
      '        '  End Using

      '        'Next

      '        lobjContents.Add(lobjTempContent)
      '#End If

      '        lintContentElement += 1

      '      Next

      '      Return lobjContents

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  ' <Added by: Rick Sherman at: 7/17/2012-3:03:05 PM>
  ''' <summary>
  ''' GetDocumentContents
  ''' </summary>
  ''' <param name="lpId"></param>
  ''' <returns></returns>
  ''' <remarks>Retrieves the latest version contents</remarks>
  Public Function GetDocumentContents(ByVal lpId As String) As Contents _
    Implements IBasicContentServicesProvider.GetDocumentContents


    Dim lobjContents As New Contents()

    Try


      Dim lobjIDocument As IDocument = GetIDocument(lpId)

      'TODO: Make sure this is the latest version
      Dim lobjILastestVersion As IDocument = lobjIDocument.VersionSeries.Versions(0)



      Dim lobjContentElementList As IContentElementList
      Dim lobjContentTransfer As IContentTransfer
      lobjContentElementList = lobjILastestVersion.Properties("ContentElements")

      Dim lobjContentStream As System.IO.Stream
      Dim byteContent() As Byte
      Dim lintStreamLength As Integer

      Dim lintContentElement As Integer = 0

      For Each lobjContentTransfer In lobjContentElementList
        'Debug.Print(lobjContentTransfer.RetrievalName)
        lobjContentStream = lobjContentTransfer.AccessContentStream

        ' Changed by Ernie Bahr 3/17/2009
        ' Instead of writing out the file and then adding to the contents
        ' collection via the path, just add the stream.

        byteContent = New Byte(lobjContentStream.Length - 1) {}
        lintStreamLength = lobjContentStream.Read(byteContent, 0, lobjContentStream.Length)


        Dim lobjTempContent As New Content(byteContent, lobjContentTransfer.RetrievalName, Content.StorageTypeEnum.Reference, False, AllowZeroLengthContent) With {
          .MIMEType = lobjContentTransfer.ContentType
        }

#If SupportAnnotations Then

        ' NOTE: The enumeration below actually iterates through starting with the newest 
        ' version and works towards the oldest.  We should first put the versions 
        ' into an array and then iterate the array in reverse.

        'For Each lobjNativeAnnotation As FileNet.Api.Core.IAnnotation In lobjILastestVersion.Annotations '  lobjNativeVersion.Annotations

        '  If lobjNativeAnnotation.AnnotatedContentElement <> lintContentElement Then Continue For

        '  Using lobjAnnotationContentStream As Stream = Me.RetrieveAnnotationContent(lobjNativeAnnotation)
        '    Dim lobjCtsAnnotation As Ecmg.Cts.Annotations.Annotation = lobjAnnotationExporter.ExportAnnotationObject(lobjAnnotationContentStream, lobjTempContent.MIMEType)
        '    ' lobjCtsAnnotation.AnnotatedContent = lobjTempContent
        '    lobjTempContent.Annotations.Add(lobjCtsAnnotation)
        '  End Using

        'Next

        lobjContents.Add(lobjTempContent)
#End If

        lintContentElement += 1

      Next

      Return lobjContents

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function
  ' </Added by: Rick Sherman at: 7/17/2012-3:03:05 PM>

  Public Function CopyOutDocument(ByVal lpId As String,
                                  ByVal lpDestinationFolder As String,
                                  ByRef lpOutputFileNames() As String) As Boolean _
         Implements IBasicContentServicesProvider.CopyOutDocument

    Try

      Throw New NotImplementedException

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  Public Function DeleteDocument(ByVal lpId As String) As Boolean _
         Implements IBasicContentServicesProvider.DeleteDocument

    Try
      Return DeleteDocument(lpId, String.Empty)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  Public Function DeleteVersion(ByVal lpDocumentId As String,
                               ByVal lpCriterion As String) As Boolean _
                             Implements IBasicContentServicesProvider.DeleteVersion
    Try
      Return (DeleteVersion(lpDocumentId, lpCriterion, String.Empty))
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function FileDocument(ByVal lpId As String,
                                 ByVal lpFolderPath As String) As Boolean _
         Implements IBasicContentServicesProvider.FileDocument

    Try
      If DocumentExists(lpId) = False Then
        Throw New DocumentDoesNotExistException(lpId)
      End If

      If String.IsNullOrEmpty(lpFolderPath) Then
        Throw New InvalidPathException("A folder path must be specified.", lpFolderPath)
      End If

      Dim lobjIDocument As IDocument = GetIDocument(lpId)

      Dim lobjFolders As IFolderSet = lobjIDocument.FoldersFiledIn

      For Each lobjFolder As FileNet.Api.Core.IFolder In lobjFolders
        If String.Compare(lobjFolder.PathName, lpFolderPath, True) = True Then
          ApplicationLogging.WriteLogEntry(String.Format("Document '{0}' is already filed in path '{1}'.", lpId, lpFolderPath), TraceEventType.Warning, 61209)
          Return False
        End If
      Next

      Dim lobjIFolder As FileNet.Api.Core.IFolder = FileDocument(lobjIDocument, lpFolderPath)

      If lobjIFolder IsNot Nothing Then
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

  Public Function GetDocumentWithContent(ByVal lpId As String,
                                         ByVal lpDestinationFolder As String) As CtsCore.Document _
         Implements IBasicContentServicesProvider.GetDocumentWithContent

    Try
      Dim lobjExportDocumentArgs As New ExportDocumentEventArgs(lpId) With
        {.GetContent = True, .GetContentFileNames = False, .GenerateCDF = False}

      ExportDocument(lobjExportDocumentArgs)

      If lobjExportDocumentArgs.Document IsNot Nothing Then
        Return lobjExportDocumentArgs.Document
      Else
        Return Nothing
      End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  Public Function GetDocumentWithContent(ByVal lpId As String,
                                         ByVal lpDestinationFolder As String,
                                         ByVal lpStorageType As CtsCore.Content.StorageTypeEnum) As CtsCore.Document _
         Implements IBasicContentServicesProvider.GetDocumentWithContent

    Try
      Dim lobjExportDocumentArgs As New ExportDocumentEventArgs(lpId) With
        {.GetContent = True, .GetContentFileNames = False, .GenerateCDF = False, .StorageType = lpStorageType}

      ExportDocument(lobjExportDocumentArgs)

      If lobjExportDocumentArgs.Document IsNot Nothing Then
        Return lobjExportDocumentArgs.Document
      Else
        Return Nothing
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  ''' <summary>
  ''' Gets the document metadata as a Document object from the repository
  ''' </summary>
  ''' <param name="lpId">The document identifier</param>
  ''' <param name="lpPropertyFilter">
  ''' A list of properties to return for the document.  
  ''' The properties returned will be limited to those in the list.
  ''' </param>
  ''' <returns>A Cts.Core.Document object reference</returns>
  ''' <remarks>This method always returns a document object with no Content values.  
  ''' If the content is required in addition to the metadata it is recommended to 
  ''' use the GetDocumentWithContent method.</remarks>
  Public Function GetDocumentWithoutContent(ByVal lpId As String, ByVal lpPropertyFilter As List(Of String)) As Document _
    Implements IBasicContentServicesProvider.GetDocumentWithoutContent

    Try


      Dim lobjExportDocumentArgs As New ExportDocumentEventArgs(lpId) With
         {.GetContent = False, .GetContentFileNames = True, .GenerateCDF = False, .PropertyFilter = lpPropertyFilter}

      ExportDocumentWithPropertyFilter(lobjExportDocumentArgs)

      If lobjExportDocumentArgs.Document IsNot Nothing Then
        Return lobjExportDocumentArgs.Document
      Else
        Return Nothing
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function GetDocumentWithoutContent(ByVal lpId As String) As CtsCore.Document _
         Implements IBasicContentServicesProvider.GetDocumentWithoutContent

    Try


      Dim lobjExportDocumentArgs As New ExportDocumentEventArgs(lpId) With
        {.GetContent = False, .GetContentFileNames = True, .GenerateCDF = False}

      ExportDocument(lobjExportDocumentArgs)

      If lobjExportDocumentArgs.Document IsNot Nothing Then
        Return lobjExportDocumentArgs.Document
      Else
        Return Nothing
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  Public Function UnFileDocument(ByVal lpId As String,
                               ByVal lpFolderPath As String) As Boolean _
       Implements IBasicContentServicesProvider.UnFileDocument

    Try

      If DocumentExists(lpId) = False Then
        Throw New DocumentDoesNotExistException(lpId)
      End If

      Dim lobjIDocument As IDocument = GetIDocument(lpId)

      Dim lobjFolders As IFolderSet = lobjIDocument.FoldersFiledIn

      For Each lobjFolder As FileNet.Api.Core.IFolder In lobjFolders

        If String.IsNullOrEmpty(lpFolderPath) Then

          Dim lobjRel As IReferentialContainmentRelationship = lobjFolder.Unfile(lobjIDocument)

          lobjRel?.Save(RefreshMode.NO_REFRESH)

          'If (lobjRel IsNot Nothing) Then
          '  lobjRel.Save(RefreshMode.NO_REFRESH)
          'End If

        Else

          If (String.Compare(lobjFolder.PathName, lpFolderPath, True) = True) OrElse
            ((lobjFolder.Id.ToString.Contains(lpFolderPath) = True) = True) Then

            Dim lobjRel As IReferentialContainmentRelationship = lobjFolder.Unfile(lobjIDocument)

            lobjRel?.Save(RefreshMode.NO_REFRESH)

            'If (lobjRel IsNot Nothing) Then
            '  lobjRel.Save(RefreshMode.NO_REFRESH)
            'End If

            Exit For
          End If

        End If

      Next

      Return True

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function


#End Region

#End Region

End Class

'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  CENetProvider_IDocumentImporter.vb
'   Description :  [type_description_here]
'   Created     :  4/8/2014 10:43:14 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"


Imports FileNet.Api.Constants
Imports Ecmg.Cts.Utilities
Imports Ecmg.Cts.Core
Imports Ecmg.Cts.Arguments
Imports FileNet.Api.Core
Imports Ecmg.Cts.Exceptions
Imports FileNet.Api.Collection
Imports FileNet.Api.Property
Imports Ecmg.Cts.Migrations
Imports System.IO
Imports Documents
Imports Documents.Arguments
Imports Documents.Core
Imports Documents.Exceptions
Imports Documents.Migrations
Imports Documents.Providers
Imports Documents.Utilities
Imports DCore = Documents.Core

#End Region

Partial Public Class CENetProvider
  Implements IDocumentImporter

#Region "IDocumentImporter Implementation"

#Region "Events"

  Public Event DocumentImported(ByVal sender As Object, ByVal e As Arguments.DocumentImportedEventArgs) _
    Implements IDocumentImporter.DocumentImported

  Public Event DocumentImportError(ByVal sender As Object, ByVal e As Arguments.DocumentImportErrorEventArgs) _
    Implements IDocumentImporter.DocumentImportError, IBasicContentServicesProvider.DocumentImportError

  Public Event DocumentImportMessage(ByVal sender As Object, ByVal e As Arguments.WriteMessageArgs) _
    Implements IDocumentImporter.DocumentImportMessage

#End Region

  ''' <summary>
  ''' Gets a value that determines whether or not import 
  ''' operations will only allow import of documents using 
  ''' classes defined for the destination repository
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks>Only applicable if the import provider also implements IClassification</remarks>
  Public ReadOnly Property EnforceClassificationCompliance() As Boolean _
    Implements IDocumentImporter.EnforceClassificationCompliance
    Get
      Return mblnEnforceClassificationCompliance
    End Get
  End Property

  Public Overridable Function ImportDocument(ByRef lpArgs As Arguments.ImportDocumentArgs) As Boolean _
    Implements IDocumentImporter.ImportDocument

    Dim lblnResult As Boolean = False

    Try

      ArgumentNullException.ThrowIfNull(lpArgs)

      Dim lobjIDocument As IDocument = AddP8Document(lpArgs.Document, lpArgs.SetPermissions, lpArgs.ErrorMessage,
                                                     lpArgs.VersionType)

      If lobjIDocument Is Nothing Then Exit Try ' Return False

      lpArgs.Document.ObjectID = lobjIDocument.Id.ToString

      Dim lstrDocumentName As String = lpArgs.Document.Name

      If String.IsNullOrEmpty(lstrDocumentName) Then lstrDocumentName = "Unnamed Document"

      lblnResult = FileDocument(lpArgs.Document, lobjIDocument, lstrDocumentName, lpArgs.FilingMode, lpArgs.PathFactory)

      If lobjIDocument Is Nothing Then Exit Try ' Return False

      '#If SupportAnnotations Then

      'ApplicationLogging.WriteLogEntry(String.Format("{0}: Annotations suported", lpArgs.Document.ObjectID), _
      '                                 Reflection.MethodBase.GetCurrentMethod(), TraceEventType.Information, 61541)
      'ApplicationLogging.WriteLogEntry(String.Format("lblnResult: {0},lpArgs.Document.HasAnnotations: {1}, lpArgs.SetAnnotations: {2}", _
      '                                               lblnResult.ToString(), lpArgs.Document.HasAnnotations, lpArgs.SetAnnotations), _
      '                                                Reflection.MethodBase.GetCurrentMethod(), TraceEventType.Information, 61543)

      If lblnResult = True AndAlso lpArgs.Document.HasAnnotations = True AndAlso lpArgs.SetAnnotations Then
        ' Only attempt to import annotations if the add document 
        ' operation succeeded and annotations are present in the document.
        Me.ImportAnnotations(lpArgs, lobjIDocument)
      End If
      '#Else
      '      ApplicationLogging.WriteLogEntry("Annotations not suported")
      '#End If

      lobjIDocument = Nothing

    Catch LargeContentEx As ContentTooLargeException
      ApplicationLogging.LogException(LargeContentEx, Reflection.MethodBase.GetCurrentMethod)
      lpArgs.ErrorMessage = LargeContentEx.Message
      lblnResult = False

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

    Return lblnResult
  End Function

#Region "Annotation Support"

  '#If SupportAnnotations Then

  Public Overridable Sub ImportAnnotations(ByRef lpArgs As Arguments.ImportDocumentArgs,
                                           ByVal lpDocument As IDocument)

    Dim lobjMemoryStream As MemoryStream = Nothing
    Dim lobjMemoryStreamWriter As StreamWriter = Nothing

    Try

      ArgumentNullException.ThrowIfNull(lpArgs)
      ArgumentNullException.ThrowIfNull(lpDocument)

      ' The document is specified by args.Document
      If lpArgs.Document.HasAnnotations = False Then Exit Sub

      ' Create a new builder for annotation content (XML)
      Dim lobjAnnotationXmlBuilder As New AnnotationImporter()

      Dim lobjPropertyFilter As New PropertyFilter
      lobjPropertyFilter.AddIncludeProperty(New FilterElement(Nothing, Nothing, Nothing, PropertyNames.CONTENT_ELEMENTS,
                                                              Nothing))

      ' For each Version in Document
      Dim lintVersionIndex As Integer = 0

      For Each lobjCtsVersion As DCore.Version In lpArgs.Document.Versions

        ' Get the "version id" from the version collection
        Dim lobjNativeDocument As IDocument = lpDocument.Versions(lpArgs.Document.Versions.Count - lintVersionIndex - 1)

        ' Obtain a web service object reference to the document version in question
        Dim lstrNativeDocumentVersionId As String = lobjNativeDocument.Id.ToString()

        '   For each ContentElement in Contents of Version
        Dim lintCtsContentElementIndex As Integer = 0

        For Each lobjContentElement As Content In lobjCtsVersion.Contents

          ' If this content element does not have an Annotations collection or it is empty, skip to the next content element
          If lobjContentElement.Annotations Is Nothing OrElse lobjContentElement.Annotations.Count = 0 Then _
            GoTo NextContentElement


          ' For each Annotation in the Annotations of the ContentElement
          For Each lobjCtsAnnotationElement As Annotations.Annotation In lobjContentElement.Annotations

            '     Create/persist a new annotation object for the document version and content element index number.
            Dim lobjNativeAnnotation As IAnnotation = Me.CreateAnnotation(lobjNativeDocument, lintCtsContentElementIndex,
                                                                          "Annotation", "Annotation")

            If lobjNativeAnnotation Is Nothing Then
              lpArgs.ErrorMessage = "Could not create an annotation instance for version " & lstrNativeDocumentVersionId
              RaiseEvent _
                DocumentImportError(Me,
                                    New DocumentImportErrorEventArgs(lpArgs.ErrorMessage,
                                                                     New Exception(lpArgs.ErrorMessage)))
            End If

            ' Set the "Id" property of the annotation to the ID returned from the web service.  This will become part of the XML.
            lobjCtsAnnotationElement.ID = lobjNativeAnnotation.Id.ToString()

            '   Create an IO.Stream to receive the annotation content (XML)
            lobjMemoryStream = New MemoryStream()

            If lobjCtsVersion.Contents.Count > 1 Then
              lobjCtsAnnotationElement.Layout.PageNumber = lintCtsContentElementIndex + 1
            End If

            '   Write the annotation content element out to the stream, building it along the way.
            lobjMemoryStreamWriter = New StreamWriter(lobjMemoryStream)
            lobjAnnotationXmlBuilder.WriteAnnotationContent(lobjMemoryStreamWriter, lobjCtsAnnotationElement,
                                                            lintCtsContentElementIndex + 1)

            '   Reset the stream pointer to the beginning.
            lobjMemoryStream.Seek(0, SeekOrigin.Begin)

            ' Associate the objAnnotation (web services object) with the annotation (CTS model) using the given annotation content builder.
            Me.UpdateAnnotationContentElement(lobjNativeAnnotation, lobjMemoryStream)

            lobjMemoryStreamWriter.Close()
          Next  ' Next annotation

NextContentElement:
          lintCtsContentElementIndex += 1 ' Advance to next content element
        Next ' Next ContentElement

        lintVersionIndex += 1 ' Advance to next version
      Next ' Next Version

    Catch ex As Exception

      If lobjMemoryStreamWriter IsNot Nothing Then lobjMemoryStreamWriter.Dispose()

      If lobjMemoryStream IsNot Nothing Then lobjMemoryStream.Dispose()

      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  ' Creates a native P8 Annotation object as an empty object, with appropriate reference to annotated content id and content element index.
  ' This is necessary in order to get the id and object reference of the annotation object itself.  
  ' The annotation's id has to appear in the XML content, so we have to create the object first, then update it.
  Private Function CreateAnnotation(ByVal lpNativeDocument As IDocument,
                                    ByVal lpContentElementIndex As Integer,
                                    ByVal lpAnnotationClassName As String,
                                    ByVal lpDescription As String) As IAnnotation

    Dim lobjResult As IAnnotation

    Try

      Dim lobjObjectStore As IObjectStore = lpNativeDocument.GetObjectStore()

      'Create the annotation.
      lobjResult = Factory.Annotation.CreateInstance(lobjObjectStore, lpAnnotationClassName)

      'Set the Document object to which the annotation applies on the annotation.
      lobjResult.AnnotatedObject = lpNativeDocument

      ' Specify the document's ContentElement to which the annotation applies.
      ' The ContentElement is identified by its element sequence number.
      Dim lobjNativeDocumentContentList As IContentElementList = lpNativeDocument.ContentElements
      Dim lobjNativeDocumentContentElement As IContentElement = lobjNativeDocumentContentList(lpContentElementIndex)
      Dim lintNativeElementSequenceNumber As Integer = lobjNativeDocumentContentElement.ElementSequenceNumber.Value
      lobjResult.AnnotatedContentElement = lintNativeElementSequenceNumber

      ' Set annotation's DescriptiveText property.
      lobjResult.DescriptiveText = lpDescription

      lobjResult.Save(RefreshMode.REFRESH)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try

    Return lobjResult
  End Function

  ' Populates an existing native P8 annotation object with the contents of a file.
  Private Sub UpdateAnnotationContentElement(ByVal lpNativeAnnotation As IAnnotation,
                                             ByVal lpContentStream As Stream)

    Try

      ArgumentNullException.ThrowIfNull(lpNativeAnnotation)
      ArgumentNullException.ThrowIfNull(lpContentStream)

      ' Create ContentTransfer and ContentElementList objects for the annotation.
      Dim lobjContentTransfer As IContentTransfer = Factory.ContentTransfer.CreateInstance()

      ' <Modified by: Ernie at 9/28/2011-9:10:47 PM on machine: ERNIE-M4400>
      ' Dim lobjAnnotationContentList As IContentElementList = Factory.ContentTransfer.CreateList()
      Dim lobjAnnotationContentList As IContentElementList = Factory.ContentElement.CreateList
      ' </Modified by: Ernie at 9/28/2011-9:10:47 PM on machine: ERNIE-M4400>

      lobjContentTransfer.SetCaptureSource(lpContentStream)
      lobjContentTransfer.RetrievalName = "annotation.xml"
      lobjContentTransfer.ContentType = "text/xml"

      ' Add ContentTransfer object to list and set the list on the annotation.
      lobjAnnotationContentList.Add(lobjContentTransfer)
      lpNativeAnnotation.ContentElements = lobjAnnotationContentList

      lpNativeAnnotation.Save(RefreshMode.REFRESH)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  '#End If

#End Region

  Private Function FileDocument(ByVal lpDocument As Document,
                                ByRef lpIdocument As IDocument,
                                ByVal lpContainmentName As String,
                                ByVal lpFilingMode As FilingMode,
                                Optional ByVal lpPathFactory As PathFactory = Nothing) As Boolean

    Dim lstrFilingError As String = String.Empty
    Dim lobjIFolder As FileNet.Api.Core.IFolder
    Dim lblnBatchFilingSuccess As Boolean = False
    Dim lstrFailedPaths() As String = Nothing

    Try

      Select Case lpFilingMode

        Case FilingMode.UnFiled
          ' Do nothing, we are done
          lblnBatchFilingSuccess = True

        Case FilingMode.BaseFolderPathOnly

          lobjIFolder = FileDocument(lpIdocument, lpPathFactory.BaseFolderPath, lpContainmentName, lstrFilingError)

          If (lobjIFolder IsNot Nothing) Then
            lblnBatchFilingSuccess = True
          End If

          Debug.WriteLine("Done filing doc " & lpContainmentName & " in '" & lobjIFolder.PathName & "'.")

        Case FilingMode.BaseFolderPathPlusDocumentFolderPath, FilingMode.DocumentFolderPathPlusBaseFolderPath

          If lpPathFactory IsNot Nothing Then

            Dim lstrFolderPath() As String
            lstrFolderPath = lpDocument.FolderPathArray(lpPathFactory)

            lblnBatchFilingSuccess = FileDocument(lpIdocument, lstrFolderPath, lpContainmentName, lstrFilingError,
                                                  lstrFailedPaths)

            If lblnBatchFilingSuccess = True Then

              For lintPathCounter As Int16 = 0 To lstrFolderPath.Length - 1
                Debug.WriteLine(
                  "Done filing doc " & lpContainmentName & " in '" & lstrFolderPath(lintPathCounter) & "'.")
              Next

            Else

              For lintPathCounter As Int16 = 0 To lstrFailedPaths.Length - 1
                Debug.WriteLine(
                  "Failed filing doc " & lpContainmentName & " in '" & lstrFailedPaths(lintPathCounter) & "'.")
              Next

            End If

          End If

        Case FilingMode.DocumentFolderPath

          Dim lstrFolderPath() As String
          Dim lobjPathFactory As PathFactory = lpPathFactory

          If lpPathFactory IsNot Nothing Then

            lobjPathFactory.BaseFolderPath = String.Empty
            lstrFolderPath = lpDocument.FolderPathArray(lobjPathFactory)

            lblnBatchFilingSuccess = FileDocument(lpIdocument, lstrFolderPath, lpContainmentName, lstrFilingError,
                                                  lstrFailedPaths)

            If lblnBatchFilingSuccess = True Then

              For lintPathCounter As Int16 = 0 To lstrFolderPath.Length - 1
                Debug.WriteLine(
                  "Done filing doc " & lpContainmentName & " in '" & lstrFolderPath(lintPathCounter) & "'.")
              Next

            Else

              For lintPathCounter As Int16 = 0 To lstrFailedPaths.Length - 1
                Debug.WriteLine(
                  "Failed filing doc " & lpContainmentName & " in '" & lstrFailedPaths(lintPathCounter) & "'.")
              Next

            End If

          End If

      End Select

      Return lblnBatchFilingSuccess

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function FileDocument(ByRef lpIdocument As IDocument,
                                ByVal lpPath() As String,
                                Optional ByVal lpContainmentName As String = "",
                                Optional ByRef lpErrorMessage As String = "",
                                Optional ByRef lpFailedPaths() As String = Nothing) As Boolean

    Dim lobjIFolder As FileNet.Api.Core.IFolder
    Dim lstrFolderPath As String
    Dim lblnFiled As Boolean = True

    Try

      For lintPathCounter As Int16 = 0 To lpPath.Length - 1
        lstrFolderPath = lpPath(lintPathCounter)
        lobjIFolder = FileDocument(lpIdocument, lstrFolderPath, lpContainmentName, lpErrorMessage)

        If lobjIFolder Is Nothing Then
          lblnFiled = False
          ReDim Preserve lpFailedPaths(lintPathCounter)
          lpFailedPaths(lintPathCounter) = lstrFolderPath
        End If

      Next

      ' If any single filing operation fails we will return false for the batch
      ' Check the lpFiledPaths parameter for the name of the failed path(s)
      Return lblnFiled

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  ' <Removed by: Ernie at: 9/29/2014-11:18:32 AM on machine: ERNIE-THINK>
  '   Public Shadows Property ImportPath() As String _
  '                Implements IDocumentImporter.ImportPath
  '     Get
  '       Try
  '         Return MyBase.ImportPath
  '       Catch ex As Exception
  '         ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '         ' Re-throw the exception to the caller
  '         Throw
  '       End Try
  '     End Get
  '     Set(ByVal value As String)
  '       Try
  '         MyBase.ImportPath = value
  '       Catch ex As Exception
  '         ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '         ' Re-throw the exception to the caller
  '         Throw
  '       End Try
  '     End Set
  '   End Property
  ' </Removed by: Ernie at: 9/29/2014-11:18:32 AM on machine: ERNIE-THINK>

  Public Sub OnDocumentImported(ByRef e As Arguments.DocumentImportedEventArgs) _
    Implements IDocumentImporter.OnDocumentImported
    RaiseEvent DocumentImported(Me, e)
  End Sub

  Public Sub OnDocumentImportError(ByRef e As Arguments.DocumentImportErrorEventArgs) _
    Implements IDocumentImporter.OnDocumentImportError
    RaiseEvent DocumentImportError(Me, e)
  End Sub

#End Region
End Class

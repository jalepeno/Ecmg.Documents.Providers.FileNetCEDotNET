'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  CENetProvider_IDocumentExporter.vb
'   Description :  [type_description_here]
'   Created     :  4/8/2014 10:31:20 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports FileNet.Api.Core
Imports Documents.Exceptions
Imports FileNet.Api.Collection
Imports FileNet.Api.Property
Imports System.IO
Imports Documents
Imports Documents.Arguments
Imports Documents.Core
Imports Documents.Providers
Imports Documents.Utilities

#End Region

Partial Public Class CENetProvider
  Implements IDocumentExporter

#Region "IDocumentExporter Implementation"

  Public Shadows Property ExportPath() As String _
                 Implements IDocumentExporter.ExportPath
    Get
      Try
        Return MyBase.ExportPath
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(ByVal value As String)
      Try
        MyBase.ExportPath = value
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  ' <Removed by: Ernie at: 9/29/2014-2:07:46 PM on machine: ERNIE-THINK>
  '   Public Function SetDocumentAsReadOnly(ByVal lpId As String) As Boolean _
  '          Implements IDocumentExporter.SetDocumentAsReadOnly
  ' 
  '     Try
  '       Throw New NotImplementedException
  ' 
  '     Catch ex As Exception
  '       ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '       ' Re-throw the exception to the caller
  '       Throw
  '     End Try
  ' 
  '   End Function
  ' </Removed by: Ernie at: 9/29/2014-2:07:46 PM on machine: ERNIE-THINK>

  Public Function DocumentCount(ByVal lpFolderPath As String,
                                Optional ByVal lpRecursionLevel As RecursionLevel = RecursionLevel.ecmThisLevelOnly) As Long _
         Implements IDocumentExporter.DocumentCount

    Try

      Dim lobjIFolder As FileNet.Api.Core.IFolder = GetFolderByPath(lpFolderPath)
      Dim lobjCollectionCounter As New CollectionCounter(ObjectStore, lobjIFolder.ContainedDocuments)

      Return lobjCollectionCounter.Count

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  Public Event DocumentExported(ByVal sender As Object, ByRef e As DocumentExportedEventArgs) Implements IDocumentExporter.DocumentExported
  Public Event FolderDocumentExported(ByVal sender As Object, ByRef e As Arguments.FolderDocumentExportedEventArgs) Implements IDocumentExporter.FolderDocumentExported
  Public Event FolderExported(ByVal sender As Object, ByRef e As Arguments.FolderExportedEventArgs) Implements IDocumentExporter.FolderExported
  Public Event DocumentExportError(ByVal sender As Object, ByVal e As Arguments.DocumentExportErrorEventArgs) Implements IDocumentExporter.DocumentExportError, IBasicContentServicesProvider.DocumentExportError
  Public Event DocumentExportMessage(ByVal sender As Object, ByVal e As Arguments.WriteMessageArgs) Implements IDocumentExporter.DocumentExportMessage

  Public Function ExportDocument(ByVal lpId As String) As Boolean _
         Implements IDocumentExporter.ExportDocument

    Try
      Return ExportDocument(New ExportDocumentEventArgs(lpId))

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  Public Function ExportDocument(ByVal Args As ExportDocumentEventArgs) As Boolean _
         Implements IDocumentExporter.ExportDocument

    Try
      'ApplicationLogging.WriteLogEntry(String.Format("{0}: Started Document Export {1}", _
      '																							Helper.ToDetailedDateString(Now), Args.Id), TraceEventType.Information, 1001)
      Dim lobjIDocument As IDocument
      Dim lobjECMDocument As New Core.Document(Me)
      ' Dim lobjECMVersion As Cts.Core.Version
      ' Dim lobjECMProperty As Cts.Core.ECMProperty
      Dim lobjECMValues As New Values
      Dim lstrDocumentPath As String
      Dim lstrCdfPath As String

      'Dim lstrVersionPath As String
      Dim lstrFullVersionNumber As String = String.Empty
      Dim lstrCopyPath As String = String.Empty
      Dim lintVersionCounter As Int16 = 0
      Dim lstrID As String = Args.Id
      Dim lobjRelationships As Relationships = Nothing

      Try

        ' Re-initialize the property value dictionary
        'PropertyValueStrings.Clear()

        lstrDocumentPath = ExportPath '& lstrID

        'Directory.CreateDirectory(ExportPath & lstrID)
      Catch PathEx As InvalidOperationException
        ApplicationLogging.LogException(PathEx, Reflection.MethodBase.GetCurrentMethod)
        Throw New InvalidOperationException("Unable to create directory for Document '" & lstrID & "' Export.", PathEx)
      End Try

#If SupportAnnotations Then
      Dim lobjAnnotationExporter As New AnnotationExporter
#End If

      lstrCdfPath = Helper.CleanPath(String.Format("{0}\{1}.{2}", lstrDocumentPath, lstrID, Document.CONTENT_DEFINITION_FILE_EXTENSION))

      lobjIDocument = GetIDocument(lstrID)

      ' Make sure that we did get the document
      If lobjIDocument Is Nothing Then
        Return False
      End If

      If Args.GetPermissions Then
        ' Get all the permissions for this document.
        lobjECMDocument.Permissions.AddRange(GetCtsPermissions(lobjIDocument))
      End If

      If ExportSystemObjectValuedProperties = True Then
        Dim lobjVersionsIProperty As FileNet.Api.Property.IProperty = lobjIDocument.Properties.GetProperty("Versions")
        If lobjVersionsIProperty IsNot Nothing Then
          Dim lobjEcmVersionsProperty As ECMProperty = CreateECMProperty(lobjVersionsIProperty)
          lobjECMDocument.Properties.Add(lobjEcmVersionsProperty)
        End If
      End If

      'If lobjIDocument.Permissions IsNot Nothing Then
      '  For Each lobjPermission As IAccessPermission In lobjIDocument.Permissions
      '    Debug.Print(String.Format("{0}: {1} - {2} ~ {3} ({4})", _
      '                              lobjPermission.GranteeType.ToString, _
      '                              lobjPermission.GranteeName, _
      '                              lobjPermission.PermissionSource.ToString, _
      '                              lobjPermission.AccessMask, _
      '                              lobjPermission.AccessType))
      '  Next
      'End If

      For Each lobjPermission As Security.IPermission In lobjECMDocument.Permissions
        Debug.Print(lobjPermission.ToString)
      Next

      'lobjECMDocument.StorageType = Content.StorageTypeEnum.EncodedUnCompressed

      lobjECMDocument.ID = lstrID

      Debug.Print(lobjIDocument.ClassDescription.Name)

      ' Get the 'Document Class'
      'lobjECMDocument.Properties.Add(New ECMProperty(PropertyType.ecmString, _
      '  versionableProperty.ecmPropertyScope.DocumentProperty, _
      '  "Document Class", lobjIDocument.ClassDescription.Name))
      lobjECMDocument.DocumentClass = lobjIDocument.ClassDescription.SymbolicName

      If Args.GetRelatedDocuments Then
        lobjRelationships = GetChildRelationships(lobjIDocument, Args.GetPermissions)
        If lobjRelationships IsNot Nothing AndAlso lobjRelationships.Count > 0 Then
          lobjECMDocument.Relationships = lobjRelationships
        End If
      End If

      ' Get the 'Folders Filed In'
      lobjECMValues.Clear()

      For Each lobjIFolder As FileNet.Api.Core.IFolder In lobjIDocument.VersionSeries.CurrentVersion.FoldersFiledIn
        'lobjECMValues.Add(lobjIFolder.PathName)
        ' <Modified by: Ernie at 1/14/2013-4:50:03 PM on machine: ERNIE-THINK>
        ' Added existing value check, we found a case where P8 had allowed an item to be filed in the same folder twice.
        If (lobjECMDocument.FolderPathsProperty IsNot Nothing) AndAlso
          (lobjECMDocument.FolderPathsProperty.Values.Contains(lobjIFolder.PathName) = False) Then
          lobjECMDocument.FolderPathsProperty.AddValue(lobjIFolder.PathName, False)
        End If
        ' </Modified by: Ernie at 1/14/2013-4:50:03 PM on machine: ERNIE-THINK>
      Next

      ' Get the document events.
      lobjECMDocument.AuditEvents = GetAuditEvents(lobjIDocument)

      'lobjECMDocument.Properties.Add(New ECMProperty(PropertyType.ecmString, _
      '  "Folders Filed In", lobjECMValues))
      'Dim lobjFoldersFiledInProperty As IMultiValuedProperty = PropertyFactory.Create(PropertyType.ecmString, _
      '                             Core.Document.FOLDER_PATHS_PROPERTY_NAME, Core.Cardinality.ecmMultiValued)
      'lobjFoldersFiledInProperty.Values = lobjECMValues
      'lobjECMDocument.Properties.Add(lobjFoldersFiledInProperty)

      ' It seems as though we get the versions from the VersionSeries in reverse order
      ' We need to ensure they are properly sorted

      ' Dim lobjIVersions As New List(Of IDocument)
      Dim lobjIVersions As New SortedDictionary(Of Single, IDocument)
      Dim lsngVersionNumber As Single

      For Each lobjIVersion As IDocument In lobjIDocument.VersionSeries.Versions
        lsngVersionNumber = lobjIVersion.MajorVersionNumber + (lobjIVersion.MinorVersionNumber / 10)
        lobjIVersions.Add(lsngVersionNumber, lobjIVersion)
      Next

      If Args.VersionScope IsNot Nothing Then
        Select Case Args.VersionScope.Scope
          Case VersionScopeEnum.MostCurrentVersion
            Dim lobjIVersion As IDocument = lobjIDocument.VersionSeries.CurrentVersion
            ExportVersion(Args, lobjIDocument, lobjIVersion, lobjECMDocument, lintVersionCounter)
          Case VersionScopeEnum.CurrentReleasedVersion
            If lobjIDocument.VersionSeries.ReleasedVersion IsNot Nothing Then
              Dim lobjIVersion As IDocument = lobjIDocument.VersionSeries.ReleasedVersion
              ExportVersion(Args, lobjIDocument, lobjIVersion, lobjECMDocument, lintVersionCounter)
            Else
              Throw New InvalidVersionSpecificationException(Args.Id,
                String.Format("Unable to export current released version, no released versions exist for version series for id '{0}'.",
                              Args.Id))
            End If
          Case VersionScopeEnum.AllVersions
            For Each lobjIVersion As IDocument In lobjIVersions.Values
              ExportVersion(Args, lobjIDocument, lobjIVersion, lobjECMDocument, lintVersionCounter)
            Next
        End Select
      Else
        For Each lobjIVersion As IDocument In lobjIVersions.Values
          ExportVersion(Args, lobjIDocument, lobjIVersion, lobjECMDocument, lintVersionCounter)
        Next
      End If

      ''  lobjIVersions.Reverse()
      lobjIVersions = Nothing

      'For Each lobjIVersion As IDocument In lobjIDocument.VersionSeries.Versions
      ''  For Each lobjIVersion As IDocument In lobjIVersions

      ''        lintVersionCounter += 1

      ''        ' Get the metadata
      ''        lobjECMVersion = New Cts.Core.Version()
      ''        'lstrFullVersionNumber = lobjIVersion.MajorVersionNumber.ToString & "." & lobjIVersion.MinorVersionNumber.ToString
      ''        lobjECMVersion.ID = lintVersionCounter 'lstrFullVersionNumber

      ''        lobjRelationships = GetChildRelationships(lobjIDocument, Args.GetPermissions)
      ''        If lobjRelationships IsNot Nothing AndAlso lobjRelationships.Count > 0 Then
      ''          lobjECMDocument.Relationships = lobjRelationships
      ''        End If

      ''        '' Add the Major Version Number
      ''        'lobjECMProperty = PropertyFactory.Create(Cts.Core.PropertyType.ecmLong, "MajorVersionNumber", lobjIVersion.MajorVersionNumber)
      ''        'lobjECMVersion.Properties.Add(lobjECMProperty)

      ''        '' Add the Minor Version Number
      ''        'lobjECMProperty = PropertyFactory.Create(Cts.Core.PropertyType.ecmLong, "MinorVersionNumber", lobjIVersion.MinorVersionNumber)
      ''        'lobjECMVersion.Properties.Add(lobjECMProperty)

      ''				For Each lobjIProperty As FileNet.Api.Property.IProperty In lobjIVersion.Properties
      ''					lobjECMProperty = CreateECMProperty(lobjIProperty)

      ''					If (lobjECMProperty IsNot Nothing) Then

      ''						If String.Equals(lobjECMProperty.Name, "ID") Then
      ''							lobjECMProperty.Name = "VersionId"
      ''							lobjECMProperty.SystemName = "VersionId"
      ''						End If

      ''						If Not lobjECMProperty Is Nothing Then
      ''							If lobjECMVersion.Properties.PropertyExists(lobjECMProperty.Name) = False Then
      ''								lobjECMVersion.Properties.Add(lobjECMProperty)
      ''								'	Debug.Print(String.Format("Added property: {0}", lobjECMProperty.Name))
      ''							End If

      ''						End If

      ''					End If

      ''				Next
      ''				lobjECMVersion.Properties.Sort()

      ''        ' Get the content

      ''        ' Hint: Use the IContentTransfer object

      ''        'Dim lobjContentElementsList As IStringList = lobjIVersion.ContentElementsPresent
      ''        If (Args.GetContent) Then

      ''          Dim lobjContentElementList As IContentElementList
      ''          Dim lobjContentTransfer As IContentTransfer
      ''          lobjContentElementList = lobjIVersion.Properties("ContentElements")

      ''          Dim lobjContentStream As System.IO.Stream
      ''          Dim byteContent() As Byte
      ''          Dim lintStreamLength As Integer

      ''          Dim lintContentElement As Integer = 0

      ''          For Each lobjContentTransfer In lobjContentElementList
      ''            'Debug.Print(lobjContentTransfer.RetrievalName)
      ''            lobjContentStream = lobjContentTransfer.AccessContentStream

      ''            ' Changed by Ernie Bahr 3/17/2009
      ''            ' Instead of writing out the file and then adding to the contents
      ''            ' collection via the path, just add the stream.

      ''            byteContent = New Byte(lobjContentStream.Length - 1) {}
      ''            'byteContent = Helper.CopyStreamToByteArray(lobjContentStream)
      ''            lintStreamLength = lobjContentStream.Read(byteContent, 0, lobjContentStream.Length)

      ''            'lstrCopyPath = lstrVersionPath & "\" & lobjContentTransfer.RetrievalName
      ''            'saveContentToFile(lstrCopyPath, byteContent)
      ''            'lobjECMVersion.Contents.Add(lobjContentStream, lobjContentTransfer.RetrievalName, Args.StorageType)

      ''            Dim lobjTempContent As Content = New Content(byteContent, lobjContentTransfer.RetrievalName, Args.StorageType, False, AllowZeroLengthContent)
      ''            lobjTempContent.MIMEType = lobjContentTransfer.ContentType
      ''#If SupportAnnotations Then

      ''            ' NOTE: The enumeration below actually iterates through starting with the newest 
      ''            ' version and works towards the oldest.  We should first put the versions 
      ''            ' into an array and then iterate the array in reverse.
      ''            If Args.GetAnnotations Then
      ''              'For Each lobjNativeVersion As IDocument In lobjIDocument.Versions '  IDMObjects.Version In lobjIdmDocument.Version.Series

      ''              'Dim lobjDocumentAnnotations As New Annotations.Annotations()
      ''              For Each lobjNativeAnnotation As FileNet.Api.Core.IAnnotation In lobjIVersion.Annotations '  lobjNativeVersion.Annotations

      ''                If lobjNativeAnnotation.AnnotatedContentElement <> lintContentElement Then Continue For

      ''                Using lobjAnnotationContentStream As Stream = Me.RetrieveAnnotationContent(lobjNativeAnnotation)
      ''                  Dim lobjCtsAnnotation As Ecmg.Cts.Annotations.Annotation = lobjAnnotationExporter.ExportAnnotationObject(lobjAnnotationContentStream, lobjTempContent.MIMEType)
      ''                  ' lobjCtsAnnotation.AnnotatedContent = lobjTempContent
      ''                  lobjTempContent.Annotations.Add(lobjCtsAnnotation)
      ''                End Using

      ''              Next


      ''              'If (mblnExportLatestVersionOnly) Then Exit For
      ''              'Next
      ''            End If

      ''#End If

      ''            lobjECMVersion.Contents.Add(lobjTempContent)
      ''            'lobjECMVersion.Contents.Add(lstrCopyPath)
      ''            'lobjECMVersion.Contents.Add(byteContent, _
      ''            '                            lobjContentTransfer.RetrievalName, _
      ''            '                            Content.StorageTypeEnum.EncodedUnCompressed)
      ''            lintContentElement += 1

      ''          Next

      ''          ' <Added by: Ernie at: 7/10/2012-4:20:33 PM on machine: ERNIE-M4400>
      ''        ElseIf Args.GetContentFileNames Then
      ''          ' We need to get the content file names but not the content itself
      ''          ' This was requested by OpenText for support of OTIC.

      ''          Dim lobjContentElementList As IContentElementList
      ''          lobjContentElementList = lobjIVersion.Properties("ContentElements")

      ''          For Each lobjContentTransfer As IContentTransfer In lobjContentElementList
      ''            lobjECMVersion.FileNames.Add(lobjContentTransfer.RetrievalName)
      ''          Next
      ''          ' </Added by: Ernie at: 7/10/2012-4:20:33 PM on machine: ERNIE-M4400>
      ''        End If


      ''        lobjECMVersion.Properties.Sort()
      ''        lobjECMDocument.Versions.Add(lobjECMVersion)

      ''        'lobjECMDocument.Serialize(lstrDocumentPath & "\" & lobjECMDocument.ID & "." & Document.CONTENT_DEFINITION_FILE_EXTENSION)

      ''  Next

      'lobjECMDocument.Properties.Sort()
      'lobjECMDocument.Versions.Sort()
      'lobjECMDocument.StorageType = Content.StorageTypeEnum.Reference
      ''lobjECMDocument.Serialize(lstrDocumentPath & "\" & lobjECMDocument.ID & "." & Document.CONTENT_DEFINITION_FILE_EXTENSION)
      'lobjECMDocument.Serialize(lstrCdfPath)

      Args.Document = lobjECMDocument

      Return ExportDocumentComplete(Me, Args)

      'RaiseEvent DocumentExported(lobjECMDocument, Now, Args.Worker, Args.DoWorkEventArgs)
      'RaiseEvent DocumentExported(Me, New DocumentExportedEventArgs(Args))

      'Return True

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Args.ErrorMessage = ex.Message
      RaiseEvent DocumentExportError(Me, New DocumentExportErrorEventArgs(Args, ex))
      Return False
    End Try

  End Function

  Private Sub ExportVersion(ByRef Args As ExportDocumentEventArgs, ByRef lobjIDocument As IDocument, ByRef lobjIVersion As IDocument, ByRef lobjECMDocument As Core.Document, lpVersionCounter As Integer)
    Try

      Dim lobjRelationships As Relationships = Nothing
      Dim lobjECMProperty As Core.ECMProperty

      ' Get the metadata
      Dim lobjECMVersion As New Core.Version()
      'lstrFullVersionNumber = lobjIVersion.MajorVersionNumber.ToString & "." & lobjIVersion.MinorVersionNumber.ToString
      lobjECMVersion.ID = lpVersionCounter 'lstrFullVersionNumber

      lobjRelationships = GetChildRelationships(lobjIDocument, Args.GetPermissions)
      If lobjRelationships IsNot Nothing AndAlso lobjRelationships.Count > 0 Then
        lobjECMDocument.Relationships = lobjRelationships
      End If

      '' Add the Major Version Number
      'lobjECMProperty = PropertyFactory.Create(Cts.Core.PropertyType.ecmLong, "MajorVersionNumber", lobjIVersion.MajorVersionNumber)
      'lobjECMVersion.Properties.Add(lobjECMProperty)

      '' Add the Minor Version Number
      'lobjECMProperty = PropertyFactory.Create(Cts.Core.PropertyType.ecmLong, "MinorVersionNumber", lobjIVersion.MinorVersionNumber)
      'lobjECMVersion.Properties.Add(lobjECMProperty)

      ' Add the VersionSeriesId
      lobjECMProperty = PropertyFactory.Create(Core.PropertyType.ecmString, "VersionSeriesId", lobjIVersion.VersionSeries.Id.ToString())
      lobjECMVersion.Properties.Add(lobjECMProperty)

      For Each lobjIProperty As FileNet.Api.Property.IProperty In lobjIVersion.Properties
        lobjECMProperty = CreateECMProperty(lobjIProperty)

        If (lobjECMProperty IsNot Nothing) Then

          If String.Equals(lobjECMProperty.Name, "ID") Then
            lobjECMProperty.Name = "VersionId"
            lobjECMProperty.SystemName = "VersionId"
          End If

          If Not lobjECMProperty Is Nothing Then
            If lobjECMVersion.Properties.PropertyExists(lobjECMProperty.Name) = False Then
              lobjECMVersion.Properties.Add(lobjECMProperty)
              '	Debug.Print(String.Format("Added property: {0}", lobjECMProperty.Name))
            End If

          End If

        End If
        lobjECMProperty = Nothing
        lobjIProperty = Nothing
      Next
      lobjECMVersion.Properties.Sort()

      ' Get the content

      ' Hint: Use the IContentTransfer object

      'Dim lobjContentElementsList As IStringList = lobjIVersion.ContentElementsPresent
      If (Args.GetContent) Then

        Dim lobjContentElementList As IContentElementList
        Dim lobjContentTransfer As IContentTransfer
        lobjContentElementList = lobjIVersion.Properties("ContentElements")

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
          'byteContent = Helper.CopyStreamToByteArray(lobjContentStream)
          lintStreamLength = lobjContentStream.Read(byteContent, 0, lobjContentStream.Length)

          'lstrCopyPath = lstrVersionPath & "\" & lobjContentTransfer.RetrievalName
          'saveContentToFile(lstrCopyPath, byteContent)
          'lobjECMVersion.Contents.Add(lobjContentStream, lobjContentTransfer.RetrievalName, Args.StorageType)

          Dim lobjTempContent As Content = New Content(byteContent, lobjContentTransfer.RetrievalName, Args.StorageType, False, AllowZeroLengthContent)
          lobjTempContent.MIMEType = lobjContentTransfer.ContentType
#If SupportAnnotations Then

          Dim lobjAnnotationExporter As New AnnotationExporter

          ' NOTE: The enumeration below actually iterates through starting with the newest 
          ' version and works towards the oldest.  We should first put the versions 
          ' into an array and then iterate the array in reverse.
          If Args.GetAnnotations Then
            'For Each lobjNativeVersion As IDocument In lobjIDocument.Versions '  IDMObjects.Version In lobjIdmDocument.Version.Series

            'Dim lobjDocumentAnnotations As New Annotations.Annotations()
            For Each lobjNativeAnnotation As FileNet.Api.Core.IAnnotation In lobjIVersion.Annotations '  lobjNativeVersion.Annotations

              If lobjNativeAnnotation.AnnotatedContentElement <> lintContentElement Then Continue For

              Using lobjAnnotationContentStream As Stream = Me.RetrieveAnnotationContent(lobjNativeAnnotation)
                Dim lobjCtsAnnotation As Ecmg.Cts.Annotations.Annotation = lobjAnnotationExporter.ExportAnnotationObject(lobjAnnotationContentStream, lobjTempContent.MIMEType)
                ' lobjCtsAnnotation.AnnotatedContent = lobjTempContent
                If lobjCtsAnnotation IsNot Nothing Then
                  lobjTempContent.Annotations.Add(lobjCtsAnnotation)
                End If
              End Using

            Next


            'If (mblnExportLatestVersionOnly) Then Exit For
            'Next
          End If

#End If

          lobjECMVersion.Contents.Add(lobjTempContent)
          'lobjECMVersion.Contents.Add(lstrCopyPath)
          'lobjECMVersion.Contents.Add(byteContent, _
          '                            lobjContentTransfer.RetrievalName, _
          '                            Content.StorageTypeEnum.EncodedUnCompressed)
          lintContentElement += 1

        Next

        ' <Added by: Ernie at: 7/10/2012-4:20:33 PM on machine: ERNIE-M4400>
      ElseIf Args.GetContentFileNames Then
        ' We need to get the content file names but not the content itself
        ' This was requested by OpenText for support of OTIC.

        Dim lobjContentElementList As IContentElementList
        lobjContentElementList = lobjIVersion.Properties("ContentElements")

        For Each lobjContentTransfer As IContentTransfer In lobjContentElementList
          lobjECMVersion.FileNames.Add(lobjContentTransfer.RetrievalName)
        Next
        ' </Added by: Ernie at: 7/10/2012-4:20:33 PM on machine: ERNIE-M4400>
      End If


      lobjECMVersion.Properties.Sort()
      lobjECMDocument.Versions.Add(lobjECMVersion)


    Catch Ex As Exception
      ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  ''' <summary>Creates a CTS Relationships collection based on the native ChildRelationships in the specified P8 document reference.</summary>
  ''' <param name="lpParentDocument">
  '''  <para>A native P8 .NET API document object reference.</para>
  ''' </param>
  ''' <param name="lpGetPermissions">
  '''  <para>Pass through parameter which specifies whether or not to get the permissions of the child document(s).</para>
  ''' </param>
  ''' <remarks>If there are no child relationships of the P8 document, an empty collection is returned.</remarks>
  ''' <returns>A CTS Relationships collection.</returns>
  Private Function GetChildRelationships(lpParentDocument As IDocument, lpGetPermissions As Boolean) As Relationships

    Try

      Dim lobjRelationships As New Relationships
      Dim lblnExportChildSuccess As Boolean

      If lpParentDocument.ChildRelationships IsNot Nothing AndAlso lpParentDocument.ChildRelationships.IsEmpty = False Then

        For Each lobjChildRelationship As IComponentRelationship In lpParentDocument.ChildRelationships

          If lobjChildRelationship.ChildComponent Is Nothing Then
            ApplicationLogging.WriteLogEntry(
              String.Format("Unable to export child relationship '{0}' of Parent '{1}:{2}', the child component is missing.",
                            lobjChildRelationship.Id, lpParentDocument.Id, lpParentDocument.Name), TraceEventType.Warning, 61231)
            Continue For
          End If
          Dim lobjRelationShip As New Relationship(lobjChildRelationship.Id.ToString,
                                                   lpParentDocument.Id.ToString,
                                                   lobjChildRelationship.ChildComponent.Id.ToString,
                                                   RelatedObjectType.Document,
                                                   RelatedObjectType.Version,
                                                   RelationshipType.Child,
                                                   RelationshipStrength.Strong,
                                                   lobjChildRelationship.ComponentRelationshipType)

          Dim lobjExportDocumentArgs As New ExportDocumentEventArgs(lobjChildRelationship.ChildComponent.Id.ToString)
          With lobjExportDocumentArgs
            .GenerateCDF = False
            .GetPermissions = lpGetPermissions
          End With
          lblnExportChildSuccess = ExportDocument(lobjExportDocumentArgs)
          If lblnExportChildSuccess Then
            lobjRelationShip.RelatedDocument = lobjExportDocumentArgs.Document
          End If

          lobjRelationships.Add(lobjRelationShip)

        Next

      End If

      Return lobjRelationships

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  Private Function ExportDocumentWithPropertyFilter(ByVal Args As ExportDocumentEventArgs) As Boolean

    Try

      Dim lobjIDocument As IDocument
      Dim lobjECMDocument As New Core.Document(Me)
      Dim lobjECMVersion As Core.Version
      Dim lobjECMProperty As Core.ECMProperty
      Dim lobjECMValues As New Values
      Dim lstrDocumentPath As String
      Dim lstrCdfPath As String

      'Dim lstrVersionPath As String
      Dim lstrFullVersionNumber As String = String.Empty
      Dim lstrCopyPath As String = String.Empty
      Dim lintVersionCounter As Int16 = 0
      Dim lstrID As String = Args.Id

      Try
        lstrDocumentPath = ExportPath '& lstrID

        'Directory.CreateDirectory(ExportPath & lstrID)
      Catch PathEx As InvalidOperationException
        ApplicationLogging.LogException(PathEx, Reflection.MethodBase.GetCurrentMethod)
        Throw New InvalidOperationException("Unable to create directory for Document '" & lstrID & "' Export.", PathEx)
      End Try

#If SupportAnnotations Then
      Dim lobjAnnotationExporter As New AnnotationExporter
#End If

      lstrCdfPath = Helper.CleanPath(String.Format("{0}\{1}.{2}", lstrDocumentPath, lstrID, Document.CONTENT_DEFINITION_FILE_EXTENSION))

      If Args.PropertyFilter IsNot Nothing Then

        ' Get document and populate property cache.
        Dim lobjIncludePropertyFilter As PropertyFilter = New PropertyFilter()

        For Each lstrPropertyName As String In Args.PropertyFilter
          lobjIncludePropertyFilter.AddIncludeProperty(New FilterElement(10, Nothing, Nothing, lstrPropertyName, Nothing))
        Next

        lobjIDocument = GetIDocument(lstrID, lobjIncludePropertyFilter)
      Else
        lobjIDocument = GetIDocument(lstrID)
      End If

      ' Make sure that we did get the document
      If lobjIDocument Is Nothing Then
        Return False
      End If


      If lobjIDocument.Properties.IsPropertyPresent("Permissions") Then

        ' Get all the permissions for this document.
        lobjECMDocument.Permissions.AddRange(GetCtsPermissions(lobjIDocument))


        For Each lobjPermission As Security.IPermission In lobjECMDocument.Permissions
          Debug.Print(lobjPermission.ToString)

        Next
      End If



      'lobjECMDocument.StorageType = Content.StorageTypeEnum.EncodedUnCompressed

      lobjECMDocument.ID = lstrID


      If lobjIDocument.Properties.IsPropertyPresent("ClassDescription") Then
        Debug.Print(lobjIDocument.ClassDescription.Name)

        ' Get the 'Document Class'
        'lobjECMDocument.Properties.Add(New ECMProperty(PropertyType.ecmString, _
        '  versionableProperty.ecmPropertyScope.DocumentProperty, _
        '  "Document Class", lobjIDocument.ClassDescription.Name))
        lobjECMDocument.DocumentClass = lobjIDocument.ClassDescription.SymbolicName
      End If


      If lobjIDocument.Properties.IsPropertyPresent("FoldersFiledIn") Then
        ' Get the 'Folders Filed In'
        lobjECMValues.Clear()

        For Each lobjIFolder As FileNet.Api.Core.IFolder In lobjIDocument.Properties("FoldersFiledIn")
          'lobjECMValues.Add(lobjIFolder.PathName)
          lobjECMDocument.FolderPathsProperty.AddValue(lobjIFolder.PathName, False)
        Next
      End If

      If lobjIDocument.Properties.IsPropertyPresent("AuditEvents") Then
        ' Get the document events.
        lobjECMDocument.AuditEvents = GetAuditEvents(lobjIDocument)
      End If



      ' It seems as though we get the versions from the VersionSeries in reverse order
      ' We need to flip them in the export
      Dim lobjIVersions As New List(Of IDocument)

      If lobjIDocument.Properties.IsPropertyPresent("VersionSeries") Then
        For Each lobjIVersion As IDocument In lobjIDocument.VersionSeries.Versions
          lobjIVersions.Add(lobjIVersion)
        Next

        lobjIVersions.Reverse()
      End If


      'For Each lobjIVersion As IDocument In lobjIDocument.VersionSeries.Versions
      For Each lobjIVersion As IDocument In lobjIVersions

        lintVersionCounter += 1

        ' Get the metadata
        lobjECMVersion = New Core.Version()
        'lstrFullVersionNumber = lobjIVersion.MajorVersionNumber.ToString & "." & lobjIVersion.MinorVersionNumber.ToString
        lobjECMVersion.ID = lintVersionCounter 'lstrFullVersionNumber

        '' Add the Major Version Number
        'lobjECMProperty = PropertyFactory.Create(Cts.Core.PropertyType.ecmLong, "MajorVersionNumber", lobjIVersion.MajorVersionNumber)
        'lobjECMVersion.Properties.Add(lobjECMProperty)

        '' Add the Minor Version Number
        'lobjECMProperty = PropertyFactory.Create(Cts.Core.PropertyType.ecmLong, "MinorVersionNumber", lobjIVersion.MinorVersionNumber)
        'lobjECMVersion.Properties.Add(lobjECMProperty)

        For Each lobjIProperty As FileNet.Api.Property.IProperty In lobjIVersion.Properties

          lobjECMProperty = CreateECMProperty(lobjIProperty)


          If (lobjECMProperty IsNot Nothing) Then

            If String.Equals(lobjECMProperty.Name, "ID") Then
              lobjECMProperty.Name = "VersionId"
              lobjECMProperty.SystemName = "VersionId"
            End If

            If Not lobjECMProperty Is Nothing Then

              If lobjECMVersion.Properties.PropertyExists(lobjECMProperty.Name) = False Then
                lobjECMVersion.Properties.Add(lobjECMProperty)
              End If

            End If

          End If

        Next

        lobjECMVersion.Properties.Sort()

        ' Get the content

        ' Hint: Use the IContentTransfer object

        'Dim lobjContentElementsList As IStringList = lobjIVersion.ContentElementsPresent
        If (Args.GetContent) Then

          Dim lobjContentElementList As IContentElementList
          Dim lobjContentTransfer As IContentTransfer
          lobjContentElementList = lobjIVersion.Properties("ContentElements")

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
            'byteContent = Helper.CopyStreamToByteArray(lobjContentStream)
            lintStreamLength = lobjContentStream.Read(byteContent, 0, lobjContentStream.Length)

            'lstrCopyPath = lstrVersionPath & "\" & lobjContentTransfer.RetrievalName
            'saveContentToFile(lstrCopyPath, byteContent)
            'lobjECMVersion.Contents.Add(lobjContentStream, lobjContentTransfer.RetrievalName, Args.StorageType)

            Dim lobjTempContent As Content = New Content(byteContent, lobjContentTransfer.RetrievalName, Args.StorageType, False, AllowZeroLengthContent)
            lobjTempContent.MIMEType = lobjContentTransfer.ContentType
#If SupportAnnotations Then

            ' NOTE: The enumeration below actually iterates through starting with the newest 
            ' version and works towards the oldest.  We should first put the versions 
            ' into an array and then iterate the array in reverse.
            If Args.GetAnnotations Then
              'For Each lobjNativeVersion As IDocument In lobjIDocument.Versions '  IDMObjects.Version In lobjIdmDocument.Version.Series

              'Dim lobjDocumentAnnotations As New Annotations.Annotations()
              For Each lobjNativeAnnotation As FileNet.Api.Core.IAnnotation In lobjIVersion.Annotations '  lobjNativeVersion.Annotations

                If lobjNativeAnnotation.AnnotatedContentElement <> lintContentElement Then Continue For

                Using lobjAnnotationContentStream As Stream = Me.RetrieveAnnotationContent(lobjNativeAnnotation)
                  Dim lobjCtsAnnotation As Ecmg.Cts.Annotations.Annotation = lobjAnnotationExporter.ExportAnnotationObject(lobjAnnotationContentStream, lobjTempContent.MIMEType)
                  ' lobjCtsAnnotation.AnnotatedContent = lobjTempContent
                  If lobjCtsAnnotation IsNot Nothing Then
                    lobjTempContent.Annotations.Add(lobjCtsAnnotation)
                  End If
                End Using

              Next


              'If (mblnExportLatestVersionOnly) Then Exit For
              'Next
            End If

#End If

            lobjECMVersion.Contents.Add(lobjTempContent)
            'lobjECMVersion.Contents.Add(lstrCopyPath)
            'lobjECMVersion.Contents.Add(byteContent, _
            '                            lobjContentTransfer.RetrievalName, _
            '                            Content.StorageTypeEnum.EncodedUnCompressed)
            lintContentElement += 1

          Next

          ' <Added by: Ernie at: 7/10/2012-4:20:33 PM on machine: ERNIE-M4400>
        ElseIf Args.GetContentFileNames Then
          ' We need to get the content file names but not the content itself
          ' This was requested by OpenText for support of OTIC.

          If lobjIDocument.Properties.IsPropertyPresent("ContentElements") Then
            Dim lobjContentElementList As IContentElementList
            lobjContentElementList = lobjIVersion.Properties("ContentElements")

            For Each lobjContentTransfer As IContentTransfer In lobjContentElementList
              lobjECMVersion.FileNames.Add(lobjContentTransfer.RetrievalName)
            Next
            ' </Added by: Ernie at: 7/10/2012-4:20:33 PM on machine: ERNIE-M4400>
          End If


        End If


        lobjECMVersion.Properties.Sort()
        lobjECMDocument.Versions.Add(lobjECMVersion)

        'lobjECMDocument.Serialize(lstrDocumentPath & "\" & lobjECMDocument.ID & "." & Document.CONTENT_DEFINITION_FILE_EXTENSION)

      Next

      'lobjECMDocument.Properties.Sort()
      'lobjECMDocument.Versions.Sort()
      'lobjECMDocument.StorageType = Content.StorageTypeEnum.Reference
      ''lobjECMDocument.Serialize(lstrDocumentPath & "\" & lobjECMDocument.ID & "." & Document.CONTENT_DEFINITION_FILE_EXTENSION)
      'lobjECMDocument.Serialize(lstrCdfPath)

      Args.Document = lobjECMDocument

      Return ExportDocumentComplete(Me, Args)

      'RaiseEvent DocumentExported(lobjECMDocument, Now, Args.Worker, Args.DoWorkEventArgs)
      'RaiseEvent DocumentExported(Me, New DocumentExportedEventArgs(Args))

      'Return True

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Args.ErrorMessage = ex.Message
      RaiseEvent DocumentExportError(Me, New DocumentExportErrorEventArgs(Args, ex))
      Return False
    End Try

  End Function




  ' <Removed by: Ernie at: 9/29/2014-1:57:03 PM on machine: ERNIE-THINK>
  '   Public Overloads Sub ExportFolder(ByVal Args As ExportFolderEventArgs) _
  '                    Implements IDocumentExporter.ExportFolder
  ' 
  '     Try
  '       MyBase.ExportFolder(Me, Args)
  ' 
  '     Catch ex As Exception
  '       ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '       ' Re-throw the exception to the caller
  '       Throw
  '     End Try
  ' 
  '   End Sub
  ' </Removed by: Ernie at: 9/29/2014-1:57:03 PM on machine: ERNIE-THINK>

  ' <Removed by: Ernie at: 9/26/2014-10:41:31 AM on machine: ERNIE-THINK>
  '   Public Overloads Function ExportDocuments(ByVal Args As ExportDocumentsEventArgs) As Boolean _
  '                    Implements IDocumentExporter.ExportDocuments
  ' 
  '     Try
  '       Return MyBase.ExportDocuments(Me, Args, AddressOf ExportDocument)
  ' 
  '     Catch ex As Exception
  '       ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
  '       ' Re-throw the exception to the caller
  '       Throw
  '     End Try
  ' 
  '   End Function
  ' </Removed by: Ernie at: 9/26/2014-10:41:31 AM on machine: ERNIE-THINK>

  Protected Sub OnDocumentExported(ByRef e As Arguments.DocumentExportedEventArgs) _
            Implements IDocumentExporter.OnDocumentExported
    RaiseEvent DocumentExported(Me, e)
  End Sub

  Protected Sub OnFolderDocumentExported(ByRef e As Arguments.FolderDocumentExportedEventArgs) _
            Implements IDocumentExporter.OnFolderDocumentExported
    RaiseEvent FolderDocumentExported(Me, e)
  End Sub

  Protected Sub OnDocumentExportError(ByRef e As Arguments.DocumentExportErrorEventArgs) _
            Implements IDocumentExporter.OnDocumentExportError
    RaiseEvent DocumentExportError(Me, e)
  End Sub

  Public Sub OnDocumentExportMessage(ByRef e As Arguments.WriteMessageArgs) _
         Implements IDocumentExporter.OnDocumentExportMessage
    RaiseEvent DocumentExportMessage(Me, e)
  End Sub

  Protected Sub OnFolderExported(ByRef e As Arguments.FolderExportedEventArgs) _
            Implements IDocumentExporter.OnFolderExported
    RaiseEvent FolderExported(Me, e)
  End Sub

#End Region

End Class

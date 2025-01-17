'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Ecmg.Cts.Core
Imports Ecmg.Cts.Annotations
Imports Ecmg.Cts.Annotations.Decoration
Imports Ecmg.Cts.Annotations.Common
Imports Ecmg.Cts.Annotations.Highlight
Imports System.Xml
Imports Ecmg.Cts.Annotations.Exception
Imports Ecmg.Cts.Annotations.Text
Imports Ecmg.Cts.Annotations.Special
Imports System.Text
Imports System.Collections.Generic
Imports System.IO
Imports Ecmg.Cts.Annotations.Shape
Imports Ecmg.Cts.Utilities
Imports Documents.Annotations.Common
Imports Documents.Annotations.Exception
Imports Documents.Annotations.Highlight
Imports Documents.Annotations.Shape
Imports Documents.Annotations.Special
Imports Documents.Annotations.Text
Imports Documents.Utilities
Imports Documents.Annotations


#End Region
''' <summary>
''' Builds an XML document / string suitable for Daeja ViewOne, based on the CTS Annotation model.
''' </summary>
Partial Public Class AnnotationImporter

#Region "Public properties"

  Protected Property Dpi As Single = 0.0

#End Region

#Region "IAnnotationImporter"

  ''' <summary>
  ''' Converts a CTS Annotation instance into Daeja ViewOne XML, writing it to the provided I/O stream.
  ''' </summary>
  ''' <param name="lpStream">The I/O stream.</param>
  ''' <param name="lpCtsAnnotation">The CTS annotation.</param>
  ''' <param name="lpMimeType">The MIME type</param>
  ''' <param name="lpContentElementIndex">The content element index</param>
  Public Sub WriteAnnotationContent(ByVal lpStream As IO.StreamWriter,
                                    ByVal lpCtsAnnotation As Annotation,
                                    ByVal lpMimeType As String,
                                    Optional ByVal lpContentElementIndex As Integer = 1)

    Try

      ArgumentNullException.ThrowIfNull(lpStream, NameOf(lpStream))
      ArgumentNullException.ThrowIfNull(lpCtsAnnotation, NameOf(lpCtsAnnotation))
      ArgumentNullException.ThrowIfNullOrEmpty(lpMimeType, NameOf(lpMimeType))

      ' Import annotation
      Dim lblnIsMultiPageTiff As Boolean = False
      Dim lstrNormalizedMimeType As String = lpMimeType.ToLowerInvariant()
      If lstrNormalizedMimeType.Equals("image/tiff") OrElse lstrNormalizedMimeType.Equals("image/x-tiff") Then lblnIsMultiPageTiff = True

      Me.Dpi = IIf(lpCtsAnnotation.Dpi = 0, 1, lpCtsAnnotation.Dpi)

      ' <FnAnno>
      Dim xmlDoc As New XmlDocument
      Dim FnAnno As XmlElement = xmlDoc.CreateElement("FnAnno")
      xmlDoc.AppendChild(FnAnno)

      ' <PropDesc 
      Dim PropDesc As XmlElement = xmlDoc.CreateElement("PropDesc")
      Me.ImportCommonMetadata(xmlDoc, PropDesc, lpCtsAnnotation, lblnIsMultiPageTiff, lpContentElementIndex)

      Dim processed As Boolean = False

      If TypeOf lpCtsAnnotation Is TextAnnotation Then
        Me.Import(xmlDoc, PropDesc, DirectCast(lpCtsAnnotation, TextAnnotation))
        processed = True
      End If

      If TypeOf lpCtsAnnotation Is HighlightRectangle Then
        Me.Import(xmlDoc, PropDesc, DirectCast(lpCtsAnnotation, HighlightRectangle))
        processed = True
      End If

      ' Arrow
      If TypeOf lpCtsAnnotation Is ArrowAnnotation Then
        Me.Import(xmlDoc, PropDesc, DirectCast(lpCtsAnnotation, ArrowAnnotation))
        processed = True
      End If

      If TypeOf lpCtsAnnotation Is StickyNoteAnnotation Then
        Me.Import(xmlDoc, PropDesc, DirectCast(lpCtsAnnotation, StickyNoteAnnotation))
        processed = True
      End If

      If TypeOf lpCtsAnnotation Is RectangleAnnotation Then
        Me.Import(xmlDoc, PropDesc, DirectCast(lpCtsAnnotation, RectangleAnnotation))
        processed = True
      End If

      If TypeOf lpCtsAnnotation Is EllipseAnnotation Then
        Me.Import(xmlDoc, PropDesc, DirectCast(lpCtsAnnotation, EllipseAnnotation))
        processed = True
      End If

      If TypeOf lpCtsAnnotation Is PointCollectionAnnotation Then
        Me.ImportLine(xmlDoc, PropDesc, DirectCast(lpCtsAnnotation, PointCollectionAnnotation))
      End If

#If ImportPen Then

    If TypeOf ctsAnnotation Is PointCollectionAnnotation Then
      Me.Import(xmlDoc, PropDesc, DirectCast(ctsAnnotation, PointCollectionAnnotation))
      processed = True
    End If

#End If

      If TypeOf lpCtsAnnotation Is StampAnnotation Then
        Me.Import(xmlDoc, PropDesc, DirectCast(lpCtsAnnotation, StampAnnotation))
        processed = True
      End If

      If Not processed Then
        Throw New UnsupportedAnnotationException("The annotation was not recognized during import.")
      End If

      ' </FnAnno>
      ' For Reference Only - "docaccesslevel" is not present in P8
      'Dim docaccesslevel As XmlElement = xmlDoc.CreateElement("docaccesslevel")
      'Dim stdText As XmlText = xmlDoc.CreateTextNode("full")
      'xmlDoc.LastChild.AppendChild(docaccesslevel)
      'xmlDoc.LastChild.LastChild.AppendChild(stdText)

      ' Load into annotation content element
      Dim annotationXml As String = xmlDoc.InnerXml
      lpStream.Write(annotationXml)

      ' Never close the stream that we're given, just flush the data through.
      lpStream.Flush()

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Debug.WriteLine(ex.Message)
      Throw
    End Try

  End Sub

#End Region

End Class

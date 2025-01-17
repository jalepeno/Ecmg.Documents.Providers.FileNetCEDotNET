'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml
Imports System.Text
Imports System.Collections.Generic
Imports System.IO
Imports Documents.Annotations.Highlight
Imports Documents.Utilities

#End Region

Partial Public Class AnnotationImporter

#Region "Annotation Import Methods"

  ''' <summary>
  ''' Imports the highlight rectangle.
  ''' </summary>
  ''' <param name="lpXmlDocument">The doc.</param>
  ''' <param name="lpXmlElement">Name of the attribute.</param>
  ''' <param name="lpCtsAnnotation">The attribute value.</param>
  Private Sub Import(ByVal lpXmlDocument As XmlDocument,
                     ByVal lpXmlElement As XmlElement,
                     ByVal lpCtsAnnotation As HighlightRectangle)

    Try

      If lpXmlDocument Is Nothing Then Throw New ArgumentNullException("lpXmlDocument", "Parameter lpXmlDocument can't be null.")
      If lpXmlElement Is Nothing Then Throw New ArgumentNullException("lpXmlElement", "Parameter lpXmlElement can't be null.")
      If lpCtsAnnotation Is Nothing Then Throw New ArgumentNullException("lpCtsAnnotation", "Parameter lpCtsAnnotation can't be null.")

      '	<PropDesc
      ' F_CLASSNAME = "Highlight"
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_CLASSNAME", "Highlight"))

      ' F_CLASSID = "{5CF11942-018F-11D0-A87A-00A0246922A5}"
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_CLASSID", "{5CF11942-018F-11D0-A87A-00A0246922A5}"))

      ' F_TEXT_BACKMODE = "1"
      ' F_TEXT_BACKMODE - P8 really sets this, although it doesn't make sense why.
      Me.ImportTextBackgroundMode(lpXmlDocument, lpXmlElement, lpCtsAnnotation)

      If lpCtsAnnotation.Display.Border IsNot Nothing Then
        ' F_LINE_COLOR = "65535"
        lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_LINE_COLOR", lpCtsAnnotation.Display.Border.Color))

        ' F_LINE_WIDTH = "8"
        lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_LINE_WIDTH", lpCtsAnnotation.Display.Border.LineStyle.LineWeight))
      End If

      ' F_BRUSHCOLOR = "39423"
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_BRUSHCOLOR", lpCtsAnnotation.HighlightColor))

      ' <F_CUSTOM_BYTES/>
      lpXmlDocument.LastChild.AppendChild(lpXmlElement)
      Me.ImportCustomBytes(lpXmlDocument, lpCtsAnnotation)

      ' <F_POINTS/>
      Me.ImportPoints(lpXmlDocument, lpCtsAnnotation)

      ' <F_TEXT/>
      Me.ImportText(lpXmlDocument, lpCtsAnnotation)

      ' </PropDesc> implied
    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Debug.WriteLine(ex.Message)
      Throw
    End Try

  End Sub

#End Region


End Class

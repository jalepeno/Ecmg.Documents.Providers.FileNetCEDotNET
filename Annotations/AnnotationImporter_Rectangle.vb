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
Imports Documents.Annotations.Shape
Imports Documents.Utilities

#End Region

Partial Public Class AnnotationImporter

#Region "Annotation Import Methods"

  ' Rectangle
  ''' <param name="lpXmlDocument">The doc.</param>
  ''' <param name="lpXmlElement">The node.</param>
  ''' <param name="lpCtsAnnotation">The CTS annotation.</param>
  Private Sub Import(ByVal lpXmlDocument As XmlDocument,
                     ByVal lpXmlElement As XmlElement,
                     ByVal lpCtsAnnotation As RectangleAnnotation)

    Try

      If lpXmlDocument Is Nothing Then Throw New ArgumentNullException("lpXmlDocument", "Parameter lpXmlDocument can't be null.")
      If lpXmlElement Is Nothing Then Throw New ArgumentNullException("lpXmlElement", "Parameter lpXmlElement can't be null.")
      If lpCtsAnnotation Is Nothing Then Throw New ArgumentNullException("lpCtsAnnotation", "Parameter lpCtsAnnotation can't be null.")

      Dim fillMode As Integer = 0

      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_CLASSNAME", "Proprietary"))
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_SUBCLASS", "v1-Rectangle"))
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_CLASSID", "{A91E5DF2-6B7B-11D1-B6D7-00609705F027}"))
      If lpCtsAnnotation.Display.Foreground IsNot Nothing AndAlso lpCtsAnnotation.Display.Foreground.Opacity > 0 Then
        lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_BRUSHCOLOR", lpCtsAnnotation.Display.Foreground))
        If lpCtsAnnotation.Display.Foreground.Opacity > 25 Then fillMode = 1
        If lpCtsAnnotation.IsFilled OrElse lpCtsAnnotation.Display.Foreground.Opacity > 75 Then fillMode = 2
      Else
        fillMode = 2
      End If
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_LINE_COLOR", lpCtsAnnotation.Display.Border.Color))
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_LINE_WIDTH", lpCtsAnnotation.LineStyle.LineWeight))
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_TEXT_BACKMODE", fillMode))



      lpXmlDocument.LastChild.AppendChild(lpXmlElement)
      Me.ImportCustomBytes(lpXmlDocument, lpCtsAnnotation)
      Me.ImportPoints(lpXmlDocument, lpCtsAnnotation)
      Me.ImportText(lpXmlDocument, lpCtsAnnotation)

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Debug.WriteLine(ex.Message)
      Throw
    End Try
  End Sub

#End Region


End Class

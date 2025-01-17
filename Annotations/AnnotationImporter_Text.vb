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
Imports Documents.Annotations.Text
Imports Documents.Utilities

#End Region

Partial Public Class AnnotationImporter

#Region "Annotation Import Methods"

  ' Text
  ''' <summary>
  ''' Imports the specified doc.
  ''' </summary>
  ''' <param name="lpXmlDocument">The doc.</param>
  ''' <param name="lpXmlElement">The node.</param>
  ''' <param name="lpCtsAnnotation">The CTS annotation.</param>
  Private Sub Import(ByVal lpXmlDocument As XmlDocument,
                     ByVal lpXmlElement As XmlElement,
                     ByVal lpCtsAnnotation As TextAnnotation)

    If lpXmlDocument Is Nothing Then Throw New ArgumentNullException("lpXmlDocument", "Parameter lpXmlDocument can't be null.")
    If lpXmlElement Is Nothing Then Throw New ArgumentNullException("lpXmlElement", "Parameter lpXmlElement can't be null.")
    If lpCtsAnnotation Is Nothing Then Throw New ArgumentNullException("lpCtsAnnotation", "Parameter lpCtsAnnotation can't be null.")

    Try
      'Dim attrib As XmlAttribute

      ' F_CLASSNAME = "Text"
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_CLASSNAME", "Text"))

      ' F_CLASSID = "{5CF11941-018F-11D0-A87A-00A0246922A5}"
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_CLASSID", "{5CF11941-018F-11D0-A87A-00A0246922A5}"))

      ' F_HASBORDER = "true" 
      ' F_BACKCOLOR = "34026" 
      ' F_BORDER_BACKMODE = "2" 
      ' F_BORDER_COLOR = "65280" 
      ' F_BORDER_STYLE = "0" 
      ' F_BORDER_WIDTH = "1"
      Me.ImportBorderInfo(lpXmlDocument, lpXmlElement, lpCtsAnnotation)

      ' F_FONT_BOLD = "true"
      ' F_FONT_ITALIC = "false"
      ' F_FONT_NAME = "arial"
      ' F_FONT_SIZE = "12"
      ' F_FONT_STRIKETHROUGH = "false"
      ' F_FONT_UNDERLINE = "false"
      ' F_FORECOLOR = "65280"
      Me.ImportFontMetadata(lpXmlDocument, lpXmlElement, lpCtsAnnotation.TextMarkups(0))

      Me.ImportTextBackgroundMode(lpXmlDocument, lpXmlElement, lpCtsAnnotation)

      '	Commit <PropDesc> element
      lpXmlDocument.LastChild.AppendChild(lpXmlElement)

      ' <F_CUSTOM_BYTES/>
      Me.ImportCustomBytes(lpXmlDocument, lpCtsAnnotation)

      ' <F_POINTS/>
      Me.ImportPoints(lpXmlDocument, lpCtsAnnotation)

      '	<F_TEXT Encoding="unicode">0054006500730074000A0054006500780074</F_TEXT>
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

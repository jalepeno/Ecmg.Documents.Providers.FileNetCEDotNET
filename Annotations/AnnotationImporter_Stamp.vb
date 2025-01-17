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
Imports Documents.Annotations.Special
Imports Documents.Utilities

#End Region

Partial Public Class AnnotationImporter

#Region "Annotation Import Methods"

  ' Stamp
  ''' <param name="lpXmlDocument">The doc.</param>
  ''' <param name="lpXmlElement">The node.</param>
  ''' <param name="lpCtsAnnotation">The CTS annotation.</param>
  Private Sub Import(ByVal lpXmlDocument As XmlDocument,
                     ByVal lpXmlElement As XmlElement,
                     ByVal lpCtsAnnotation As StampAnnotation)

    Try

      If lpXmlDocument Is Nothing Then Throw New ArgumentNullException("lpXmlDocument", "Parameter lpXmlDocument can't be null.")
      If lpXmlElement Is Nothing Then Throw New ArgumentNullException("lpXmlElement", "Parameter lpXmlElement can't be null.")
      If lpCtsAnnotation Is Nothing Then Throw New ArgumentNullException("lpCtsAnnotation", "Parameter lpCtsAnnotation can't be null.")

      '<!--  Stamp annotation
      'Yellow box, red border, reading "Urgent!" with dots missing from text
      'text is at a diagonal from lower-left to upper right -->
      '	<PropDesc F_ANNOTATEDID="{9E8F50B6-CE56-4E9C-A9F4-6E637D8BFA3E}" F_BACKCOLOR="65535" F_BORDER_BACKMODE="2" F_BORDER_COLOR="255" F_BORDER_STYLE="0" F_BORDER_WIDTH="1" F_CLASSID="{5CF1194C-018F-11D0-A87A-00A0246922A5}" F_CLASSNAME="Stamp" F_ENTRYDATE="2010-07-01T21:38:34.0000000-05:00" F_FONT_BOLD="true" F_FONT_ITALIC="false" F_FONT_NAME="arial" F_FONT_SIZE="22" F_FONT_STRIKETHROUGH="false" F_FONT_UNDERLINE="false" F_FORECOLOR="255" F_HASBORDER="true" F_HEIGHT="1.37" F_ID="{9E8F50B6-CE56-4E9C-A9F4-6E637D8BFA3E}" F_LEFT="0.3433333333333333" F_MODIFYDATE="2010-07-01T21:39:03.0000000-05:00" F_MULTIPAGETIFFPAGENUMBER="0" F_NAME="-1-1" F_PAGENUMBER="1" F_ROTATION="45" F_TEXT_BACKMODE="2" F_TOP="1.19" F_WIDTH="1.5033333333333334">

      '	<PropDesc
      ' F_CLASSNAME = "Stamp"
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_CLASSNAME", "Stamp"))

      ' F_CLASSID = "{5CF11942-018F-11D0-A87A-00A0246922A5}"
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_CLASSID", "{5CF1194C-018F-11D0-A87A-00A0246922A5}"))

      Me.ImportBorderInfo(lpXmlDocument, lpXmlElement, lpCtsAnnotation)
      Me.ImportFontMetadata(lpXmlDocument, lpXmlElement, lpCtsAnnotation.TextElement)

      ' F_BACKCOLOR - should already be set
      'node.Attributes.Append(Me.CreateAttribute(doc, "F_BACKCOLOR", ctsAnnotation.Display.Background))

      ' F_FORECOLOR = "255" - should already be set

      ' F_ROTATION = "45"
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_ROTATION", 360.0! - lpCtsAnnotation.TextElement.TextRotation))

      ' F_TEXT_BACKMODE = "2" - should already be set
      Me.ImportTextBackgroundMode(lpXmlDocument, lpXmlElement, lpCtsAnnotation)

      ''	Commit <PropDesc> element
      lpXmlDocument.LastChild.AppendChild(lpXmlElement)

      ''       <F_CUSTOM_BYTES/>
      Me.ImportCustomBytes(lpXmlDocument, lpCtsAnnotation)

      ''       <F_POINTS/>
      Me.ImportPoints(lpXmlDocument, lpCtsAnnotation)

      '	<F_TEXT Encoding="unicode">0055007200670065006E00740021</F_TEXT>
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

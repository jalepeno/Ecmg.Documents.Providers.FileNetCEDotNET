'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Ecmg.Cts.Core
Imports Ecmg.Cts.Annotations
Imports Documents.Annotations.Decoration
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
Imports Documents.Annotations
Imports Documents.Annotations.Common
Imports Documents.Annotations.Special
Imports Documents.Annotations.Text
Imports Documents.Utilities


#End Region

Partial Public Class AnnotationImporter

#Region "XML builder helper methods"

  ''' <summary>
  ''' Processes the "custom bytes" portion of the annotation.
  ''' Every annotation in this system has at least an empty XML tag for this.
  ''' </summary>
  ''' <param name="lpXmlDocument">The doc.</param>
  ''' <param name="lpCtsAnnotation">The CTS annotation.</param>
  Private Sub ImportCustomBytes(ByVal lpXmlDocument As XmlDocument,
                                ByVal lpCtsAnnotation As Annotation)

    Try

      ArgumentNullException.ThrowIfNull(lpXmlDocument, NameOf(lpXmlDocument))
      ArgumentNullException.ThrowIfNull(lpCtsAnnotation, NameOf(lpCtsAnnotation))

      Dim node As XmlNode = lpXmlDocument.CreateElement("F_CUSTOM_BYTES")

      lpXmlDocument.DocumentElement.LastChild.AppendChild(node)

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Debug.WriteLine(ex.Message)
      Throw
    End Try

  End Sub

  ''' <summary>
  ''' Processes the "data points" portion of the annotation.
  ''' Every annotation in this system has at least an empty XML tag for this.
  ''' </summary>
  ''' <param name="lpXmlDocument">The doc.</param>
  ''' <param name="lpCtsAnnotation">The CTS annotation.</param>
  Private Sub ImportPoints(ByVal lpXmlDocument As XmlDocument,
                           ByVal lpCtsAnnotation As Annotation)

    Try

      ArgumentNullException.ThrowIfNull(lpXmlDocument, NameOf(lpXmlDocument))
      ArgumentNullException.ThrowIfNull(lpCtsAnnotation, NameOf(lpCtsAnnotation))

      Dim node As XmlNode = lpXmlDocument.CreateElement("F_POINTS")

      lpXmlDocument.DocumentElement.LastChild.AppendChild(node)

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Debug.WriteLine(ex.Message)
      Throw
    End Try

  End Sub

  ''' <summary>
  ''' Processes the "text string" portion of the annotation.
  ''' Every annotation in this system has at least an empty XML tag for this.
  ''' </summary>
  ''' <param name="lpXmlDocument">The doc.</param>
  ''' <param name="lpCtsAnnotation">The CTS annotation.</param>
  Private Sub ImportText(ByVal lpXmlDocument As XmlDocument,
                         ByVal lpCtsAnnotation As Annotation)

    Try

      ArgumentNullException.ThrowIfNull(lpXmlDocument, NameOf(lpXmlDocument))
      ArgumentNullException.ThrowIfNull(lpCtsAnnotation, NameOf(lpCtsAnnotation))

      Dim element As XmlElement = lpXmlDocument.CreateElement("F_TEXT")

      Dim encodedOutput As String = Nothing

      If TypeOf lpCtsAnnotation Is TextAnnotation Then

        Dim annotation As TextAnnotation = DirectCast(lpCtsAnnotation, TextAnnotation)

        ' For P8, only one text markup per annotation is possible at this time in Daeja
        Dim markup = annotation.TextMarkups(0)
        encodedOutput = Me.FormatUnicodeHexString(markup.Text)
      End If

      If TypeOf lpCtsAnnotation Is StampAnnotation Then

        Dim annotation As StampAnnotation = DirectCast(lpCtsAnnotation, StampAnnotation)
        encodedOutput = Me.FormatUnicodeHexString(annotation.TextElement.Text)
      End If

      If TypeOf lpCtsAnnotation Is StickyNoteAnnotation Then

        Dim annotation As StickyNoteAnnotation = DirectCast(lpCtsAnnotation, StickyNoteAnnotation)
        encodedOutput = Me.FormatUnicodeHexString(annotation.TextNote.Text)
      End If

      Dim previousParent = lpXmlDocument.DocumentElement.LastChild

      If encodedOutput IsNot Nothing Then
        element.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "Encoding", "unicode"))
        element.InnerText = encodedOutput
      End If

      previousParent.AppendChild(element)

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Debug.WriteLine(ex.Message)
      Throw
    End Try

  End Sub

#End Region

#Region "Data exchange helper methods"

  ''' <summary>
  ''' Formats the color for this patform using RGB levels.
  ''' </summary>
  ''' <param name="lpColorInfo">The ColorInfo to format as an RGB integer</param>
  ''' <returns></returns>
  Private Function FormatColor(ByVal lpColorInfo As ColorInfo) As String

    Try

      If lpColorInfo Is Nothing Then Throw New ArgumentNullException(NameOf(lpColorInfo), "Parameter lpColorInfo can't be null.")

      ' left as multiple statements for debugging.
      Dim result As Integer
      result = lpColorInfo.Blue
      result *= 256
      result += lpColorInfo.Green
      result *= 256
      result += lpColorInfo.Red
      Return result.ToString

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Debug.WriteLine(ex.Message)
      Throw
    End Try

  End Function

  ''' <summary>
  ''' Formats the date.
  ''' </summary>
  ''' <param name="lpDateTimeOffset">The CTS date.</param>
  ''' <returns></returns>
  Private Function FormatDate(ByVal lpDateTimeOffset As DateTimeOffset) As String

    Try

      ' F_ENTRYDATE =  "2010-07-01T21:20:36.0000000-05:00"
      ' F_MODIFYDATE = "2010-07-01T21:20:50.0000000-05:00"
      Dim result As String = String.Format("{0:yyyy-MM-dd}T{0:HH:mm:ss.fffffffK}", lpDateTimeOffset)
      Return result

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Debug.WriteLine(ex.Message)
      Throw
    End Try

  End Function

  Private Function FormatUnicodeHexString(ByVal lpString As String) As String

    Try

      If lpString Is Nothing Then
        Return String.Empty
      End If

      Dim builder As New StringBuilder()

      For i As Integer = 0 To lpString.Length - 1
        builder.AppendFormat("{0:X4}", Asc(lpString(i)))
      Next

      Return builder.ToString()

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Debug.WriteLine(ex.Message)
      Throw
    End Try

  End Function

#Region "Overloads for CreateAttribute"

  ''' <summary>
  ''' Creates an XML attribute having the name and value provided.
  ''' </summary>
  ''' <param name="lpXmlDocument">The doc.</param>
  ''' <param name="lpAttributeName">Name of the attribute.</param>
  ''' <param name="lpAttributeValue">The attribute value.</param>
  ''' <returns></returns>
  Private Function CreateAttribute(ByVal lpXmlDocument As XmlDocument,
                                   ByVal lpAttributeName As String,
                                   ByVal lpAttributeValue As Boolean)

    Try
      If lpXmlDocument Is Nothing Then Throw New ArgumentNullException(NameOf(lpXmlDocument), "Parameter lpXmlDocument can't be null.")
      If lpAttributeName Is Nothing Then Throw New ArgumentNullException(NameOf(lpAttributeName), "Parameter lpAttributeName can't be null.")
      ' The only way Daeja P8 will recognize Boolean attributes is if they are in *lowercase*
      Dim attributeValueString As String

      If lpAttributeValue Then
        attributeValueString = "true"

      Else
        attributeValueString = "false"
      End If

      Return Me.CreateAttribute(lpXmlDocument, lpAttributeName, attributeValueString)

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Debug.WriteLine(ex.Message)
      Throw
    End Try

  End Function

  ''' <summary>
  ''' Creates an XML attribute having the name and value provided.
  ''' </summary>
  ''' <param name="lpXmlDocument">The doc.</param>
  ''' <param name="lpAttributeName">Name of the attribute.</param>
  ''' <param name="lpAttributeValue">The attribute value.</param>
  ''' <returns></returns>
  Private Function CreateAttribute(ByVal lpXmlDocument As XmlDocument,
                                   ByVal lpAttributeName As String,
                                   ByVal lpAttributeValue As ColorInfo)

    Try
      If lpXmlDocument Is Nothing Then Throw New ArgumentNullException(NameOf(lpXmlDocument), "Parameter lpXmlDocument can't be null.")
      If lpAttributeName Is Nothing Then Throw New ArgumentNullException(NameOf(lpAttributeName), "Parameter lpAttributeName can't be null.")
      If lpAttributeValue Is Nothing Then Throw New ArgumentNullException(NameOf(lpAttributeValue), "Parameter lpAttributeValue can't be null.")

      Return Me.CreateAttribute(lpXmlDocument, lpAttributeName, Me.FormatColor(lpAttributeValue))

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Debug.WriteLine(ex.Message)
      Throw
    End Try

  End Function

  ''' <summary>
  ''' Creates an XML attribute having the name and value provided.
  ''' </summary>
  ''' <param name="lpXmlDocument">The doc.</param>
  ''' <param name="lpAttributeName">Name of the attribute.</param>
  ''' <param name="lpAttributeValue">The attribute value.</param>
  ''' <returns></returns>
  Private Function CreateAttribute(ByVal lpXmlDocument As XmlDocument,
                                   ByVal lpAttributeName As String,
                                   ByVal lpAttributeValue As Single)

    Try
      ArgumentNullException.ThrowIfNull(lpXmlDocument, NameOf(lpXmlDocument))
      ArgumentNullException.ThrowIfNull(lpAttributeName, NameOf(lpAttributeName))

      Return Me.CreateAttribute(lpXmlDocument, lpAttributeName, lpAttributeValue.ToString())

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Debug.WriteLine(ex.Message)
      Throw
    End Try

  End Function

  ''' <summary>
  ''' Creates an XML attribute having the name and value provided.
  ''' </summary>
  ''' <param name="lpXmlDocument">The doc.</param>
  ''' <param name="lpAttributeName">Name of the attribute.</param>
  ''' <param name="lpAttributeValue">The attribute value.</param>
  ''' <returns></returns>
  Private Function CreateAttribute(ByVal lpXmlDocument As XmlDocument,
                                   ByVal lpAttributeName As String,
                                   ByVal lpAttributeValue As String)

    Try
      ArgumentNullException.ThrowIfNull(lpXmlDocument, NameOf(lpXmlDocument))
      ArgumentNullException.ThrowIfNull(lpAttributeName, NameOf(lpAttributeName))
      If lpAttributeValue Is Nothing Then
        Return String.Empty
      End If

      Dim attrib As XmlAttribute
      attrib = lpXmlDocument.CreateAttribute(lpAttributeName)
      If lpAttributeValue Is Nothing Then
        attrib.Value = String.Empty
      Else
        attrib.Value = lpAttributeValue
      End If

      Return attrib

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Debug.WriteLine(ex.Message)
      Throw
    End Try

  End Function

  ''' <summary>
  ''' Creates an XML attribute having the name and value provided.
  ''' </summary>
  ''' <param name="lpXmlDocument">The doc.</param>
  ''' <param name="lpAttributeName">Name of the attribute.</param>
  ''' <param name="lpAttributeValue">The attribute value.</param>
  ''' <returns></returns>
  Private Function CreateAttribute(ByVal lpXmlDocument As XmlDocument,
                                   ByVal lpAttributeName As String,
                                   ByVal lpAttributeValue As DateTimeOffset)

    Try
      ArgumentNullException.ThrowIfNull(lpXmlDocument, NameOf(lpXmlDocument))
      ArgumentNullException.ThrowIfNull(lpAttributeName, NameOf(lpAttributeName))

      Return Me.CreateAttribute(lpXmlDocument, lpAttributeName, Me.FormatDate(lpAttributeValue))

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Debug.WriteLine(ex.Message)
      Throw
    End Try

  End Function

#End Region

#End Region
#Region "Annotation metadata helpers"

  ''' <summary>
  ''' Imports the border info.
  ''' Only valid for Stamp and Text annotations on this platform.
  ''' </summary>
  ''' <param name="lpXmlDocument">The doc.</param>
  ''' <param name="lpXmlNode">Name of the attribute.</param>
  ''' <param name="lpCtsAnnotation">The attribute value.</param>
  Private Sub ImportBorderInfo(ByVal lpXmlDocument As XmlDocument,
                               ByVal lpXmlNode As XmlNode,
                               ByVal lpCtsAnnotation As Annotation)

    Try

      ArgumentNullException.ThrowIfNull(lpXmlDocument, NameOf(lpXmlDocument))
      ArgumentNullException.ThrowIfNull(lpXmlNode, NameOf(lpXmlNode))
      ArgumentNullException.ThrowIfNull(lpCtsAnnotation, NameOf(lpCtsAnnotation))

      ' F_HASBORDER 
      If lpCtsAnnotation.Display.Border Is Nothing Then
        lpXmlNode.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_HASBORDER", False))
        Return
      End If

      lpXmlNode.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_HASBORDER", True))

      ' F_BACKCOLOR
      If lpCtsAnnotation.Display.Background Is Nothing Then
        lpXmlNode.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_BACKCOLOR", "256"))
      Else
        lpXmlNode.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_BACKCOLOR", Me.FormatColor(lpCtsAnnotation.Display.Background)))
      End If

      ' F_BORDER_BACKMODE
      Dim borderBackMode As Integer = 0

      If lpCtsAnnotation.Display.Border.Color.Opacity > 75 Then
        borderBackMode = 2

      ElseIf lpCtsAnnotation.Display.Border.Color.Opacity > 25 Then
        borderBackMode = 1
      End If

      lpXmlNode.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_BORDER_BACKMODE", borderBackMode))

      ' F_BORDER_COLOR
      lpXmlNode.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_BORDER_COLOR", Me.FormatColor(lpCtsAnnotation.Display.Border.Color)))

      ' F_BORDER_WIDTH
      lpXmlNode.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_BORDER_WIDTH", lpCtsAnnotation.Display.Border.LineStyle.LineWeight))

      ' F_BORDER_STYLE
      Dim borderStyle As Integer = 0

      Select Case lpCtsAnnotation.Display.Border.LineStyle.Pattern

        Case LineStyleInfo.LinePattern.Solid
          borderStyle = 1

        Case LineStyleInfo.LinePattern.Dash
          borderStyle = 2

        Case LineStyleInfo.LinePattern.Dot
          borderStyle = 3

        Case LineStyleInfo.LinePattern.DashDot
          borderStyle = 4

        Case LineStyleInfo.LinePattern.DashDotDot
          borderStyle = 5

        Case Else
          ' TODO: Throw exception or favor robustness?

      End Select

      lpXmlNode.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_BORDER_STYLE", borderStyle))

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Debug.WriteLine(ex.Message)
      Throw
    End Try

  End Sub

  ''' <summary>
  ''' Imports the font metadata.
  ''' </summary>
  ''' <param name="lpXmlDocument">The doc.</param>
  ''' <param name="lpXmlNode">Name of the attribute.</param>
  ''' <param name="lpTextMarkup">The TextMarkup object to serialize.</param>
  Private Sub ImportFontMetadata(ByVal lpXmlDocument As XmlDocument,
                                 ByVal lpXmlNode As XmlNode,
                                 ByVal lpTextMarkup As TextMarkup)

    Try

      ArgumentNullException.ThrowIfNull(lpXmlDocument)
      ArgumentNullException.ThrowIfNull(lpXmlNode)
      ArgumentNullException.ThrowIfNull(lpTextMarkup)

      'If lpXmlDocument Is Nothing Then Throw New ArgumentNullException("lpXmlDocument", "Parameter lpXmlDocument can't be null.")
      'If lpXmlNode Is Nothing Then Throw New ArgumentNullException("lpXmlNode", "Parameter lpXmlNode can't be null.")
      'If lpTextMarkup Is Nothing Then Throw New ArgumentNullException("lpTextMarkup", "Parameter lpTextMarkup can't be null.")

      ' P8 provides support for only one text markup within an annotation.

      ' F_FONT_BOLD = "true"
      lpXmlNode.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_FONT_BOLD", lpTextMarkup.Font.IsBold))

      ' F_FONT_ITALIC = "false"
      lpXmlNode.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_FONT_ITALIC", lpTextMarkup.Font.IsItalic))

      ' F_FONT_NAME = "arial"
      ' Ideally, we should do a check on the local platform to see if the font is installed and if not, select another based on the FontFamily property.
      ' The operation is expensive, so the font lookup wrapper should be memoized.
      lpXmlNode.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_FONT_NAME", lpTextMarkup.Font.FontName.ToLowerInvariant()))

      ' F_FONT_SIZE = "12"
      Dim lsglFontSize As Single = lpTextMarkup.Font.FontSize
      If Me.Dpi > 1 Then ' A custom DPI was set for a font based in pixels, not font points
        lsglFontSize /= (Me.Dpi * 0.013889)
      End If
      lpXmlNode.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_FONT_SIZE", lsglFontSize))

      ' F_FONT_STRIKETHROUGH = "false"
      lpXmlNode.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_FONT_STRIKETHROUGH", lpTextMarkup.Font.IsStrikethrough))

      ' F_FONT_UNDERLINE = "false"
      lpXmlNode.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_FONT_UNDERLINE", lpTextMarkup.Font.IsUnderline))

      ' F_FORECOLOR = "65280"
      lpXmlNode.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_FORECOLOR", Me.FormatColor(lpTextMarkup.Font.Color)))

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Sub

  ''' <summary>
  ''' Imports the line metadata.
  ''' </summary>
  ''' <param name="lpXmlDocument">The doc.</param>
  ''' <param name="lpXmlNode">Name of the attribute.</param>
  ''' <param name="lpCtsAnnotation">The attribute value.</param>
  Private Sub ImportLineMetadata(ByVal lpXmlDocument As XmlDocument,
                                 ByVal lpXmlNode As XmlNode,
                                 ByVal lpCtsAnnotation As LineBase)

    Try

      ArgumentNullException.ThrowIfNull(lpXmlDocument, NameOf(lpXmlDocument))
      ArgumentNullException.ThrowIfNull(lpXmlNode, NameOf(lpXmlNode))
      ArgumentNullException.ThrowIfNull(lpCtsAnnotation, NameOf(lpCtsAnnotation))


      ' F_LINE_BACKMODE="2"
      If lpCtsAnnotation.Display.Foreground.Opacity > 50 Then
        lpXmlNode.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_LINE_BACKMODE", 2))

      Else
        lpXmlNode.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_LINE_BACKMODE", 1))
      End If

      ' F_LINE_COLOR="16711680"
      lpXmlNode.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_LINE_COLOR", lpCtsAnnotation.Display.Foreground))

      ' F_LINE_STYLE="0"
      Dim lineStyle As Integer = 0

      Select Case lpCtsAnnotation.LineStyle.Pattern

        Case LineStyleInfo.LinePattern.Solid
          lineStyle = 0

        Case LineStyleInfo.LinePattern.Dash
          lineStyle = 1

        Case LineStyleInfo.LinePattern.Dot
          lineStyle = 2

        Case LineStyleInfo.LinePattern.DashDot
          lineStyle = 3

        Case LineStyleInfo.LinePattern.DashDotDot
          lineStyle = 4

        Case Else
          ' TODO: Throw exception or favor robustness?

      End Select

      lpXmlNode.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_LINE_STYLE", lineStyle))

      ' F_LINE_WIDTH="12"
      lpXmlNode.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_LINE_WIDTH", lpCtsAnnotation.LineStyle.LineWeight))

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Debug.WriteLine(ex.Message)
      Throw
    End Try

  End Sub

  ''' <summary>
  ''' Imports the text background mode.
  ''' </summary>
  ''' <param name="lpXmlDocument">The doc.</param>
  ''' <param name="lpXmlNode">Name of the attribute.</param>
  ''' <param name="lpCtsAnnotation">The attribute value.</param>
  Private Sub ImportTextBackgroundMode(ByVal lpXmlDocument As XmlDocument,
                                       ByVal lpXmlNode As XmlNode,
                                       ByVal lpCtsAnnotation As Annotation)

    Try

      ArgumentNullException.ThrowIfNull(lpXmlDocument)
      ArgumentNullException.ThrowIfNull(lpXmlNode)
      ArgumentNullException.ThrowIfNull(lpCtsAnnotation)

      'If lpXmlDocument Is Nothing Then Throw New ArgumentNullException("lpXmlDocument", "Parameter lpXmlDocument can't be null.")
      'If lpXmlNode Is Nothing Then Throw New ArgumentNullException("lpXmlNode", "Parameter lpXmlNode can't be null.")
      'If lpCtsAnnotation Is Nothing Then Throw New ArgumentNullException("lpCtsAnnotation", "Parameter lpCtsAnnotation can't be null.")


      If lpCtsAnnotation.Display.Background Is Nothing Then Return

      ' F_BACKCOLOR
      lpXmlNode.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_BACKCOLOR", Me.FormatColor(lpCtsAnnotation.Display.Background)))

      ' F_TEXT_BACKMODE = "2"
      If lpCtsAnnotation.Display.Background.Opacity <= 50 Then
        lpXmlNode.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_TEXT_BACKMODE", 1))

      Else
        lpXmlNode.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_TEXT_BACKMODE", 2))
      End If

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Debug.WriteLine(ex.Message)
      Throw
    End Try

  End Sub

#End Region
  ''' <summary>
  ''' Imports the common metadata.
  ''' </summary>
  ''' <param name="lpXmlDocument">The doc.</param>
  ''' <param name="lpXmlElement"></param>
  ''' <param name="lpCtsAnnotation"></param>
  ''' <param name="lpIsMultiPageTiff"></param>
  ''' <param name="lpContentElementIndex"></param>
  Private Sub ImportCommonMetadata(ByVal lpXmlDocument As XmlDocument,
                                   ByVal lpXmlElement As XmlElement,
                                   ByVal lpCtsAnnotation As Annotation,
                                   ByVal lpIsMultiPageTiff As Boolean,
                                   ByVal lpContentElementIndex As Integer)

    Try

      ArgumentNullException.ThrowIfNull(lpXmlDocument, NameOf(lpXmlDocument))
      ArgumentNullException.ThrowIfNull(lpXmlElement, NameOf(lpXmlElement))
      ArgumentNullException.ThrowIfNull(lpCtsAnnotation, NameOf(lpCtsAnnotation))

      ' F_ANNOTATEDID = "{28EADBFC-CACE-4882-B9B4-9068B2D894EB}" 
      ' Daeja stores the id of the annotation here, not the id of the document or of the content element.
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_ANNOTATEDID", lpCtsAnnotation.ID))

      ' F_CLASSID
      ' Specified in each annotation import method.

      ' F_CLASSNAME
      ' Specified in each annotation import method.

      ' F_ENTRYDATE = "2010-07-01T21:20:36.0000000-05:00"
      If lpCtsAnnotation.AuditEvents IsNot Nothing AndAlso lpCtsAnnotation.AuditEvents.Created IsNot Nothing Then
        lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_ENTRYDATE", lpCtsAnnotation.AuditEvents.Created.EventTime))
      Else
        lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_ENTRYDATE", System.DateTimeOffset.Now))
      End If

      ' F_HEIGHT = "0.3333333333333333"
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_HEIGHT", (lpCtsAnnotation.Layout.LowerRightExtent.Second / Me.Dpi) - (lpCtsAnnotation.Layout.UpperLeftExtent.Second / Me.Dpi)))

      ' F_ID = "{28EADBFC-CACE-4882-B9B4-9068B2D894EB}"  (Non-portable, must be set to the id in the target system.)
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_ID", lpCtsAnnotation.ID))

      ' F_LEFT = "0.26666666666666666"
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_LEFT", lpCtsAnnotation.Layout.UpperLeftExtent.First / Me.Dpi))

      ' F_MODIFYDATE = "2010-07-01T21:20:50.0000000-05:00"
      If lpCtsAnnotation.AuditEvents IsNot Nothing AndAlso lpCtsAnnotation.AuditEvents.Modified IsNot Nothing Then
        lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_MODIFYDATE", lpCtsAnnotation.AuditEvents.Modified.EventTime))
      Else
        lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_MODIFYDATE", System.DateTimeOffset.Now))
      End If

      If lpIsMultiPageTiff Then
        lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_MULTIPAGETIFFPAGENUMBER", lpCtsAnnotation.Layout.PageNumber))
      Else
        lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_PAGENUMBER", lpCtsAnnotation.Layout.PageNumber))
        lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_MULTIPAGETIFFPAGENUMBER", 0))

      End If

      ' Set F_MULTIPAGETIFFPAGENUMBER if TIFF, otherwise set F_PAGENUMBER
      '       F_MULTIPAGETIFFPAGENUMBER = "0"
      'attrib = doc.CreateAttribute("F_MULTIPAGETIFFPAGENUMBER")
      'attrib.Value = Nothing
      'node.Attributes.Append(attrib)
      '       F_PAGENUMBER = "1"
      'attrib = doc.CreateAttribute("F_PAGENUMBER")
      'attrib.Value = Nothing
      'node.Attributes.Append(attrib)

      ' F_NAME = "-1-1"
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_NAME", "-" & lpContentElementIndex & "-" & lpCtsAnnotation.ID))

      ' F_TOP = "0.3"
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_TOP", lpCtsAnnotation.Layout.UpperLeftExtent.Second / Me.Dpi))

      ' F_WIDTH="0.5833333333333334">
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_WIDTH", (lpCtsAnnotation.Layout.LowerRightExtent.First / Me.Dpi) - (lpCtsAnnotation.Layout.UpperLeftExtent.First / Me.Dpi)))

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Debug.WriteLine(ex.Message)
      Throw
    End Try

  End Sub

End Class

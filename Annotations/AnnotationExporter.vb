'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.IO
Imports System.Xml
Imports System.Collections.ObjectModel
Imports System.Text
Imports Documents.Annotations
Imports Documents.Annotations.Auditing
Imports Documents.Annotations.Common
Imports Documents.Annotations.Decoration
Imports Documents.Annotations.Exception
Imports Documents.Annotations.Highlight
Imports Documents.Annotations.Shape
Imports Documents.Annotations.Special
Imports Documents.Annotations.Text
Imports Documents.Utilities

#End Region

Public Class AnnotationExporter

#Region "Class constants"

  Private Const PROP_DESC_PATH As String = "/FnAnno/PropDesc"
  Private Const TEXT_PATH As String = PROP_DESC_PATH + "/F_TEXT"

#End Region

#Region "Public Properties"

  Public Property ScaleX As Single = 96.0
  Public Property ScaleY As Single = 96.0

#End Region

#Region "Public methods"

  Public Function ExportAnnotationObject(ByVal lpAnnotation As Stream, ByVal lpMimeType As String) As Annotation
    Dim lobjResult As Annotation = Nothing

    Try
      ArgumentNullException.ThrowIfNull(lpAnnotation)
      ArgumentNullException.ThrowIfNull(lpMimeType)

      Dim lobjAnnotationXml As New CtsXmlDocument()
      lobjAnnotationXml.Load(lpAnnotation)

      Dim lobjAnnotationType As Type = GetAnnotationType(lobjAnnotationXml)
      If lobjAnnotationType Is Nothing Then
        ApplicationLogging.WriteLogEntry("Could not determine CTS annotation class")
        Exit Try
      End If

      Dim lblnProcessed As Boolean = False
      Dim lblnIsMultiPageTiff As Boolean = False
      Dim lstrNormalizedMimeType As String = lpMimeType.ToLowerInvariant()
      If lstrNormalizedMimeType.Equals("image/tiff") OrElse lstrNormalizedMimeType.Equals("image/x-tiff") Then lblnIsMultiPageTiff = True

      If lobjAnnotationType Is GetType(ArrowAnnotation) Then
        lobjResult = New ArrowAnnotation()
        Me.ExportCommonMetadata(lobjResult, lobjAnnotationXml, lblnIsMultiPageTiff)
        Me.ExportArrow(lobjResult, lobjAnnotationXml)
        lblnProcessed = True
      End If

      If lobjAnnotationType Is GetType(EllipseAnnotation) Then
        lobjResult = New EllipseAnnotation()
        Me.ExportCommonMetadata(lobjResult, lobjAnnotationXml, lblnIsMultiPageTiff)
        Me.ExportEllipse(lobjResult, lobjAnnotationXml)
        lblnProcessed = True
      End If

      If lobjAnnotationType Is GetType(HighlightRectangle) Then
        lobjResult = New HighlightRectangle()
        Me.ExportCommonMetadata(lobjResult, lobjAnnotationXml, lblnIsMultiPageTiff)
        Me.ExportHighlightRectangle(lobjResult, lobjAnnotationXml)
        lblnProcessed = True
      End If

      If lobjAnnotationType Is GetType(RectangleAnnotation) Then
        lobjResult = New RectangleAnnotation()
        Me.ExportCommonMetadata(lobjResult, lobjAnnotationXml, lblnIsMultiPageTiff)
        Me.ExportRectangle(lobjResult, lobjAnnotationXml)
        lblnProcessed = True
      End If

      If lobjAnnotationType Is GetType(StampAnnotation) Then
        lobjResult = New StampAnnotation()
        Me.ExportCommonMetadata(lobjResult, lobjAnnotationXml, lblnIsMultiPageTiff)
        Me.ExportStamp(lobjResult, lobjAnnotationXml)
        lblnProcessed = True
      End If

      If lobjAnnotationType Is GetType(StickyNoteAnnotation) Then
        lobjResult = New StickyNoteAnnotation()
        Me.ExportCommonMetadata(lobjResult, lobjAnnotationXml, lblnIsMultiPageTiff)
        ExportStickyNote(lobjResult, lobjAnnotationXml)
        lblnProcessed = True
      End If

      If lobjAnnotationType Is GetType(TextAnnotation) Then
        lobjResult = New TextAnnotation()
        Me.ExportCommonMetadata(lobjResult, lobjAnnotationXml, lblnIsMultiPageTiff)
        Me.ExportText(lobjResult, lobjAnnotationXml)
        lblnProcessed = True
      End If

      If lobjAnnotationType Is GetType(PointCollectionAnnotation) Then
        lobjResult = New PointCollectionAnnotation()
        Me.ExportCommonMetadata(lobjResult, lobjAnnotationXml, lblnIsMultiPageTiff)
        ExportPointCollection(lobjResult, lobjAnnotationXml)
      End If

      If Not lblnProcessed Then
        Throw New UnsupportedAnnotationException("The annotation was not recognized during import.")
      End If


    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try

    Return lobjResult

  End Function

#End Region

#Region "Private Methods"

#Region "Annotation exports"

  Private Shared Function GetAnnotationType(ByVal lpXmlAnnotation As CtsXmlDocument) As Type
    Dim lobjResult As Type = Nothing
    Try
      ArgumentNullException.ThrowIfNull(lpXmlAnnotation)

      Dim lstrClassId As String = lpXmlAnnotation.QuerySingleAttribute(PROP_DESC_PATH, "F_CLASSID")
      Dim lstrClassName As String = lpXmlAnnotation.QuerySingleAttribute(PROP_DESC_PATH, "F_CLASSNAME")
      Dim lstrSubclassName As String = lpXmlAnnotation.QuerySingleAttribute(PROP_DESC_PATH, "F_SUBCLASS")

      For Each lobjItem As PlatformAnnotation In SupportedAnnotations.Instance.Items
        If lobjItem.ClassId.CompareTo(lstrClassId) <> 0 Then Continue For
        If lobjItem.ClassName.CompareTo(lstrClassName) <> 0 Then Continue For
        If String.IsNullOrEmpty(lobjItem.SubClassName) Then
          lobjResult = lobjItem.AnnotationType
          Exit For
        End If

        If lobjItem.SubClassName.CompareTo(lstrSubclassName) <> 0 Then Continue For
        lobjResult = lobjItem.AnnotationType

      Next


      '' Note: the next check will not execute if the F_CLASSNAME element was not found.
      'If String.IsNullOrEmpty(lstrClassName) Then
      '    ApplicationLogging.WriteLogEntry("Could not determine annotation class name.")
      '    Exit Try
      'End If

      'lstrClassName = lstrClassName.ToLowerInvariant()
      'Select Case lstrClassName
      '    Case "arrow" : lobjResult = New ArrowAnnotation()
      '    Case "stamp" : lobjResult = New StampAnnotation()
      '    Case "stickynote" : lobjResult = New StickyNoteAnnotation()
      '    Case "text" : lobjResult = New TextAnnotation()
      '    Case "highlight" : lobjResult = New HighlightRectangle()


      'End Select

      If lobjResult Is Nothing Then
        If String.IsNullOrEmpty(lstrSubclassName) Then lstrSubclassName = String.Empty
        ApplicationLogging.WriteLogEntry(String.Format("Could not map annotation type.  ClassId='{0}', ClassName='{1}', SubClass='{2}'", lstrClassId, lstrClassName, lstrSubclassName))
      End If

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try

    Return lobjResult
  End Function

  Private Sub ExportCommonMetadata(ByVal lpAnnotation As Annotation, ByVal lpXmlAnnotation As CtsXmlDocument, ByVal lpIsMultiPageTiff As Boolean)
    Try
      ArgumentNullException.ThrowIfNull(lpAnnotation)
      ArgumentNullException.ThrowIfNull(lpXmlAnnotation)

      lpAnnotation.ID = lpXmlAnnotation.QuerySingleAttribute(PROP_DESC_PATH, "F_ANNOTATEDID")   ' F_ID is also set for the same value

      With lpAnnotation.AuditEvents
        .Created = New CreateEvent With {
          .EventTime = lpXmlAnnotation.QuerySingleAttributeAsDate(PROP_DESC_PATH, "F_ENTRYDATE")
        }
        .Modified = New ModifyEvent With {
          .EventTime = lpXmlAnnotation.QuerySingleAttributeAsDate(PROP_DESC_PATH, "F_MODIFYDATE")
        }
      End With

      With lpAnnotation.Layout
        If lpIsMultiPageTiff Then
          .PageNumber = lpXmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_MULTIPAGETIFFPAGENUMBER")
        Else
          .PageNumber = lpXmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_PAGENUMBER")
        End If

        .UpperLeftExtent = New Point With {
          .First = ScaleX * lpXmlAnnotation.QuerySingleAttributeAsSingle(PROP_DESC_PATH, "F_LEFT"),
          .Second = ScaleY * lpXmlAnnotation.QuerySingleAttributeAsSingle(PROP_DESC_PATH, "F_TOP")
        }
        .LowerRightExtent = New Point With {
          .First = ScaleX * lpXmlAnnotation.QuerySingleAttributeAsSingle(PROP_DESC_PATH, "F_WIDTH") + lpAnnotation.Layout.UpperLeftExtent.First,
          .Second = ScaleY * lpXmlAnnotation.QuerySingleAttributeAsSingle(PROP_DESC_PATH, "F_HEIGHT") + lpAnnotation.Layout.UpperLeftExtent.Second
        }
      End With

      '' F_NAME = "-1-1"
      'element.Attributes.Append(Me.CreateAttribute(doc, "F_NAME", "-" & contentElementIndex & "-" & ctsAnnotation.ID))
      ' This extracts the content element index from F_NAME attribute and passes it back to the provider as a platform-specific property.
      ' The provider will check this value as it is iterating over the contents collection for the version.
      'Dim lstrName As String = lpXmlAnnotation.QuerySingleAttribute(PROP_DESC_PATH, "F_NAME")
      'Dim laryItems() As String = lstrName.Split("-")

      'Dim lintResult As Integer
      'If Integer.TryParse(laryItems(1), lintResult) Then

      '    ' lpAnnotation.Properties.Add(PropertyFactory.Create(PropertyType.ecmLong, "ContentIndex", lintResult, Cardinality.ecmSingleValued))
      'End If

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Private Sub ExportArrow(ByVal lpAnnotation As ArrowAnnotation, ByVal lpXmlAnnotation As CtsXmlDocument)
    Try
      ArgumentNullException.ThrowIfNull(lpAnnotation)
      ArgumentNullException.ThrowIfNull(lpXmlAnnotation)

      Me.ExportLineMetadata(lpAnnotation, lpXmlAnnotation)
      lpAnnotation.Size = lpXmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_ARROWHEAD_SIZE")
      lpAnnotation.StartPoint = New Point With {
        .First = lpXmlAnnotation.QuerySingleAttributeAsSingle(PROP_DESC_PATH, "F_LINE_START_X") * ScaleX,
        .Second = lpXmlAnnotation.QuerySingleAttributeAsSingle(PROP_DESC_PATH, "F_LINE_START_Y") * ScaleY
      }

      lpAnnotation.EndPoint = New Point With {
        .First = lpXmlAnnotation.QuerySingleAttributeAsSingle(PROP_DESC_PATH, "F_LINE_END_X") * ScaleX,
        .Second = lpXmlAnnotation.QuerySingleAttributeAsSingle(PROP_DESC_PATH, "F_LINE_END_Y") * ScaleY
      }

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Private Sub ExportEllipse(ByVal lpAnnotation As EllipseAnnotation, ByVal lpXmlAnnotation As CtsXmlDocument)
    Try
      ArgumentNullException.ThrowIfNull(lpAnnotation)
      ArgumentNullException.ThrowIfNull(lpXmlAnnotation)

      With lpAnnotation.Display
        .Foreground = ParseColor(lpXmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_LINE_COLOR"))
        If lpXmlAnnotation.QuerySingleAttribute(PROP_DESC_PATH, "F_LINE_BACK_MODE") Is Nothing Then
          .Foreground.Opacity = 100
        Else
          Dim lintLineBackMode As Integer = lpXmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_LINE_BACK_MODE")
          .Foreground.Opacity = 50 * lintLineBackMode
        End If
      End With

      lpAnnotation.LineStyle = New LineStyleInfo()
      With lpAnnotation.LineStyle
        Dim lintLinePattern As Integer = lpXmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_LINE_STYLE")
        Select Case lintLinePattern
          Case 0 : .Pattern = LineStyleInfo.LinePattern.Solid
          Case 1 : .Pattern = LineStyleInfo.LinePattern.Dash
          Case 2 : .Pattern = LineStyleInfo.LinePattern.Dot
          Case 3 : .Pattern = LineStyleInfo.LinePattern.DashDot
          Case 4 : .Pattern = LineStyleInfo.LinePattern.DashDotDot
          Case Else
            ApplicationLogging.WriteLogEntry(String.Format("Unknown line pattern value {0}", lintLinePattern))
        End Select

        Dim lintLineWeight As Integer = lpXmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_LINE_WIDTH")
        .LineWeight = lintLineWeight
      End With

      ' LineColor          ctsAnnotation.Display.Foreground
      ' Transparency       ctsAnnotation.Display.Background.Opacity <50 if 1 otherwise 2
      ' LineWidth          ctsAnnotation.Display.Border.LineStyle.LineWeight
    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Private Sub ExportHighlightRectangle(ByVal lpAnnotation As HighlightRectangle, ByVal lpXmlAnnotation As CtsXmlDocument)
    Try
      ArgumentNullException.ThrowIfNull(lpAnnotation)
      ArgumentNullException.ThrowIfNull(lpXmlAnnotation)

      ExportTextBackgroundMode(lpAnnotation, lpXmlAnnotation)

      lpAnnotation.HighlightColor = ParseColor(lpXmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_BRUSHCOLOR"))
      'Dim lintOpacity As Integer = lpXmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_TEXT_BACKMODE")
      'lpAnnotation.HighlightColor.Opacity = 50 * lintOpacity

      If String.IsNullOrEmpty(lpXmlAnnotation.QuerySingleAttribute(PROP_DESC_PATH, "F_LINE_WIDTH")) Then Exit Try

      ' optional
      Dim lintBorderWidth As Integer = lpXmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_LINE_WIDTH")
      If lintBorderWidth <> 0 Then
        Dim lintBorderColor As Integer = lpXmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_LINE_COLOR")
        lpAnnotation.Display.Border = New BorderInfo()
        With lpAnnotation.Display.Border
          .Color = ParseColor(lintBorderColor)
          .LineStyle = New LineStyleInfo() With {.LineWeight = lintBorderWidth, .Pattern = LineStyleInfo.LinePattern.Solid}
        End With
      End If

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Private Sub ExportRectangle(ByVal lpAnnotation As RectangleAnnotation, ByVal lpXmlAnnotation As CtsXmlDocument)
    Try
      ArgumentNullException.ThrowIfNull(lpAnnotation)
      ArgumentNullException.ThrowIfNull(lpXmlAnnotation)

      With lpAnnotation.Display
        If lpXmlAnnotation.QuerySingleAttribute(PROP_DESC_PATH, "F_BRUSHCOLOR") IsNot Nothing Then
          .Foreground = ParseColor(lpXmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_BRUSHCOLOR"))
          .Foreground.Opacity = 50 * lpXmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_TEXT_BACKMODE")
        Else
          .Foreground = Nothing '  New ColorInfo() With {.Blue = 255, .Green = 255, .Red = 255, .Opacity = 0}
        End If

        .Border = New BorderInfo With {
          .Color = ParseColor(lpXmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_LINE_COLOR"))
        }

      End With

      lpAnnotation.LineStyle.LineWeight = lpXmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_LINE_WIDTH")

      ' Me.ExportTextMetadata(lpAnnotation, lpXmlAnnotation)

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Private Sub ExportStamp(ByVal lpAnnotation As StampAnnotation, ByVal lpXmlAnnotation As CtsXmlDocument)
    Try
      ArgumentNullException.ThrowIfNull(lpAnnotation)
      ArgumentNullException.ThrowIfNull(lpXmlAnnotation)

      Me.ExportBorderInfo(lpAnnotation, lpXmlAnnotation)
      lpAnnotation.TextElement = New TextMarkup()
      Me.ExportFontMetadata(lpAnnotation.TextElement, lpXmlAnnotation)
      lpAnnotation.TextElement.TextRotation = 360.0! - lpXmlAnnotation.QuerySingleAttributeAsSingle(PROP_DESC_PATH, "F_ROTATION")
      ExportTextBackgroundMode(lpAnnotation, lpXmlAnnotation)
      ExportTextMetadata(lpAnnotation.TextElement, lpXmlAnnotation)
      ' Import Code

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Private Shared Sub ExportStickyNote(ByVal lpAnnotation As StickyNoteAnnotation, ByVal lpXmlAnnotation As CtsXmlDocument)
    Try
      ArgumentNullException.ThrowIfNull(lpAnnotation)
      ArgumentNullException.ThrowIfNull(lpXmlAnnotation)

      ' Daeja on P8 does not support setting the color of the stickynote, but it does specify the color.
      ' F_FORECOLOR = "10092543"
      lpAnnotation.Display.Background = ParseColor(lpXmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_FORECOLOR"))

      ' F_ORDINAL is not used in P8.  CS and IS used this.
      ' lpAnnotation.NoteOrder = lpXmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_ORDINAL")
      ExportTextMetadata(lpAnnotation.TextNote, lpXmlAnnotation)
      ' Import code

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Private Sub ExportText(ByVal lpAnnotation As TextAnnotation, ByVal lpXmlAnnotation As CtsXmlDocument)
    Try
      ArgumentNullException.ThrowIfNull(lpAnnotation)
      ArgumentNullException.ThrowIfNull(lpXmlAnnotation)

      Dim lobjTextMarkup As New TextMarkup
      Me.ExportBorderInfo(lpAnnotation, lpXmlAnnotation)
      Me.ExportFontMetadata(lobjTextMarkup, lpXmlAnnotation)
      ExportTextBackgroundMode(lpAnnotation, lpXmlAnnotation)
      ExportTextMetadata(lobjTextMarkup, lpXmlAnnotation)
      lpAnnotation.TextMarkups.Add(lobjTextMarkup)

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Private Shared Sub ExportPointCollection(
    ByVal lpAnnotation As PointCollectionAnnotation,
    ByVal lpXmlAnnotation As CtsXmlDocument)

    Try

      ArgumentNullException.ThrowIfNull(lpAnnotation)
      ArgumentNullException.ThrowIfNull(lpXmlAnnotation)

      ' PointCollection maps to more than one annotation type.  We need to determine the type for the target
      Dim lstrSubclassName As String = lpXmlAnnotation.QuerySingleAttribute(PROP_DESC_PATH, "F_SUBCLASS")

      If String.Compare(lstrSubclassName, "v1-Line", True, System.Globalization.CultureInfo.InvariantCulture) = 0 Then
        ExportLineAsPointCollection(lpAnnotation, lpXmlAnnotation)
        Exit Try
      End If

      If String.Compare(lstrSubclassName, "Pen", True, System.Globalization.CultureInfo.InvariantCulture) = 0 Then
        ExportPenAsPointCollection(lpAnnotation, lpXmlAnnotation)
        Exit Try
      End If

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw

    End Try

  End Sub

  Private Shared Sub ExportLineAsPointCollection(
    ByVal lpAnnotation As PointCollectionAnnotation,
    ByVal lpXmlAnnotation As CtsXmlDocument)

    Try
      ArgumentNullException.ThrowIfNull(lpAnnotation)
      ArgumentNullException.ThrowIfNull(lpXmlAnnotation)

      Dim lintLineWeight As Single = lpXmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_LINE_WIDTH")
      Dim lsglX1 As Single = lpXmlAnnotation.QuerySingleAttributeAsSingle(PROP_DESC_PATH, "F_LINE_START_X")
      Dim lsglY1 As Single = lpXmlAnnotation.QuerySingleAttributeAsSingle(PROP_DESC_PATH, "F_LINE_START_Y")
      Dim lsglX2 As Single = lpXmlAnnotation.QuerySingleAttributeAsSingle(PROP_DESC_PATH, "F_LINE_END_X")
      Dim lsglY2 As Single = lpXmlAnnotation.QuerySingleAttributeAsSingle(PROP_DESC_PATH, "F_LINE_END_Y")
      Dim lobjPointSytle As New PointStyle With {.Endpoint = PointStyle.EndpointStyle.None, .Filled = False, .Thickness = lintLineWeight}

      lpAnnotation.SetStartPoint(lsglX1, lsglY1, lobjPointSytle)
      lpAnnotation.AddSegment(lsglX2, lsglY2, lobjPointSytle)
      lpAnnotation.Display.Border.LineStyle.LineWeight = lintLineWeight

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Throw
    End Try

  End Sub

  Private Shared Sub ExportPenAsPointCollection(
    ByVal lpAnnotation As PointCollectionAnnotation,
    ByVal lpXmlAnnotation As CtsXmlDocument)

    Try
      ArgumentNullException.ThrowIfNull(lpAnnotation)
      ArgumentNullException.ThrowIfNull(lpXmlAnnotation)

      ' To be developed
      Dim lintLineWeight As Single = lpXmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_LINE_WIDTH")
      Dim lsglX1 As Single = lpXmlAnnotation.QuerySingleAttributeAsSingle(PROP_DESC_PATH, "F_LINE_START_X")
      Dim lsglY1 As Single = lpXmlAnnotation.QuerySingleAttributeAsSingle(PROP_DESC_PATH, "F_LINE_START_Y")
      Dim lsglX2 As Single = lpXmlAnnotation.QuerySingleAttributeAsSingle(PROP_DESC_PATH, "F_LINE_END_X")
      Dim lsglY2 As Single = lpXmlAnnotation.QuerySingleAttributeAsSingle(PROP_DESC_PATH, "F_LINE_END_Y")
      Dim lobjPointSytle As New PointStyle With {.Endpoint = PointStyle.EndpointStyle.None, .Filled = False, .Thickness = lintLineWeight}

      lpAnnotation.SetStartPoint(lsglX1, lsglY1, lobjPointSytle)
      lpAnnotation.AddSegment(lsglX2, lsglY2, lobjPointSytle)
      lpAnnotation.Display.Border.LineStyle.LineWeight = lintLineWeight

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Throw
    End Try

  End Sub

#End Region

#Region "Annotation Metadata helpers"

  Private Sub ExportBorderInfo(ByVal lpAnnotation As Annotation, ByVal lpXmlAnnotation As CtsXmlDocument)
    Try
      ArgumentNullException.ThrowIfNull(lpAnnotation)
      ArgumentNullException.ThrowIfNull(lpXmlAnnotation)

      Dim lblnHasBorder As Boolean = lpXmlAnnotation.QuerySingleAttributeAsBoolean(PROP_DESC_PATH, "F_HASBORDER")
      If Not lblnHasBorder Then
        lpAnnotation.Display.Border = Nothing
        Exit Try
      End If

      lpAnnotation.Display.Border = New BorderInfo()
      With lpAnnotation.Display
        .Background = ParseColor(lpXmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_BACKCOLOR"))
        .Border.Color = ParseColor(lpXmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_BORDER_COLOR"))
        .Border.Color.Opacity = 50 * lpXmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_BORDER_BACKMODE")
      End With

      With lpAnnotation.Display.Border
        .LineStyle = New LineStyleInfo With {
          .LineWeight = lpXmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_BORDER_WIDTH")
        }
        Dim lintBorderStyle As Integer = lpXmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_BORDER_STYLE")
        Select Case lintBorderStyle
          Case 0 : .LineStyle.Pattern = LineStyleInfo.LinePattern.Solid
          Case 1 : .LineStyle.Pattern = LineStyleInfo.LinePattern.Dash
          Case 2 : .LineStyle.Pattern = LineStyleInfo.LinePattern.Dot
          Case 3 : .LineStyle.Pattern = LineStyleInfo.LinePattern.DashDot
          Case 4 : .LineStyle.Pattern = LineStyleInfo.LinePattern.DashDotDot
          Case Else
            ApplicationLogging.WriteLogEntry(String.Format("Unknown border linestyle value {0}", lintBorderStyle))
        End Select

      End With

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Private Sub ExportFontMetadata(ByVal lpTextMarkup As TextMarkup, ByVal lpXmlAnnotation As CtsXmlDocument)
    Try
      ArgumentNullException.ThrowIfNull(lpTextMarkup)
      ArgumentNullException.ThrowIfNull(lpXmlAnnotation)

      lpTextMarkup.Font = New FontInfo()
      With lpTextMarkup.Font
        .IsBold = lpXmlAnnotation.QuerySingleAttributeAsBoolean(PROP_DESC_PATH, "F_FONT_BOLD")
        .IsItalic = lpXmlAnnotation.QuerySingleAttributeAsBoolean(PROP_DESC_PATH, "F_FONT_ITALIC")
        .IsStrikethrough = lpXmlAnnotation.QuerySingleAttributeAsBoolean(PROP_DESC_PATH, "F_FONT_STRIKETHROUGH")
        .IsUnderline = lpXmlAnnotation.QuerySingleAttributeAsBoolean(PROP_DESC_PATH, "F_FONT_UNDERLINE")
        .FontName = lpXmlAnnotation.QuerySingleAttribute(PROP_DESC_PATH, "F_FONT_NAME")
        .FontSize = lpXmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_FONT_SIZE")
        .Color = ParseColor(lpXmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_FORECOLOR"))

        ' Preserve the font family, in case the target platform does not know the font by the exact name specified.
        ' If the font size is 0.0, then we can't initialize the platform's font info for font family.
        ' Check for fonts less than 0.5, because 0.0 may not test true.
        Dim lintSafeFontSize As Single = IIf(.FontSize < 0.5, 12.0, .FontSize)
        Dim lobjPlatformFont As New System.Drawing.Font(.FontName, lintSafeFontSize)
        .FontFamily = lobjPlatformFont.FontFamily.Name

      End With
      ' P8 provides support for only one text markup within an annotation.

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Private Sub ExportLineMetadata(ByVal lpAnnotation As ArrowAnnotation, ByVal lpXmlAnnotation As CtsXmlDocument)
    Try
      ArgumentNullException.ThrowIfNull(lpAnnotation)
      ArgumentNullException.ThrowIfNull(lpXmlAnnotation)

      With lpAnnotation.Display
        .Foreground = ParseColor(lpXmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_LINE_COLOR"))
        If lpXmlAnnotation.QuerySingleAttribute(PROP_DESC_PATH, "F_LINE_BACK_MODE") Is Nothing Then
          .Foreground.Opacity = 100
        Else
          Dim lintLineBackMode As Integer = lpXmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_LINE_BACK_MODE")
          .Foreground.Opacity = 50 * lintLineBackMode
        End If
      End With

      lpAnnotation.LineStyle = New LineStyleInfo()
      With lpAnnotation.LineStyle

        Dim lintLinePattern As Integer = lpXmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_LINE_STYLE")
        Select Case lintLinePattern
          Case 0 : .Pattern = LineStyleInfo.LinePattern.Solid
          Case 1 : .Pattern = LineStyleInfo.LinePattern.Dash
          Case 2 : .Pattern = LineStyleInfo.LinePattern.Dot
          Case 3 : .Pattern = LineStyleInfo.LinePattern.DashDot
          Case 4 : .Pattern = LineStyleInfo.LinePattern.DashDotDot
          Case Else
            ApplicationLogging.WriteLogEntry(String.Format("Unknown line pattern value {0}", lintLinePattern))
        End Select

        Dim lintLineWeight As Integer = lpXmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_LINE_WIDTH")
        .LineWeight = lintLineWeight

      End With

      ' Import code

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Private Shared Sub ExportTextBackgroundMode(ByVal lpAnnotation As Annotation, ByVal lpXmlAnnotation As CtsXmlDocument)
    Try

      ArgumentNullException.ThrowIfNull(lpAnnotation)
      ArgumentNullException.ThrowIfNull(lpXmlAnnotation)

      If lpXmlAnnotation.QuerySingleAttribute(PROP_DESC_PATH, "F_BACKCOLOR") IsNot Nothing Then
        Dim lintBackgroundColor As Integer = lpXmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_BACKCOLOR")
        lpAnnotation.Display.Background = ParseColor(lintBackgroundColor)
        lpAnnotation.Display.Background.Opacity = 100

        If lpXmlAnnotation.QuerySingleAttribute(PROP_DESC_PATH, "F_TEXT_BACKMODE") IsNot Nothing Then
          Dim lintTextBackgroundMode As Integer = lpXmlAnnotation.QuerySingleAttributeAsInteger(PROP_DESC_PATH, "F_TEXT_BACKMODE")
          lpAnnotation.Display.Background.Opacity = 50 * lintTextBackgroundMode
        End If
      End If


    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Private Shared Sub ExportTextMetadata(ByVal lpTextMarkup As TextMarkup, ByVal lpXmlAnnotation As CtsXmlDocument)
    Try
      ArgumentNullException.ThrowIfNull(lpTextMarkup)
      ArgumentNullException.ThrowIfNull(lpXmlAnnotation)

      Dim lstrHexString As String = lpXmlAnnotation.QuerySingleString(TEXT_PATH)
      If String.IsNullOrEmpty(lstrHexString) Then Exit Try


      lpTextMarkup.Text = DecodeUnicodeHexString(lstrHexString)

      'Import code
      'Dim encodedOutput As String = Nothing

      'If TypeOf ctsAnnotation Is TextAnnotation Then

      '    Dim annotation As TextAnnotation = DirectCast(ctsAnnotation, TextAnnotation)

      '    ' For P8, only one text markup per annotation is possible at this time in Daeja
      '    Dim markup = annotation.TextMarkups(0)
      '    encodedOutput = Me.FormatUnicodeHexString(markup.Text)
      'End If

      'If TypeOf ctsAnnotation Is StampAnnotation Then

      '    Dim annotation As StampAnnotation = DirectCast(ctsAnnotation, StampAnnotation)
      '    encodedOutput = Me.FormatUnicodeHexString(annotation.TextElement.Text)
      'End If

      'If TypeOf ctsAnnotation Is StickyNoteAnnotation Then

      '    Dim annotation As StickyNoteAnnotation = DirectCast(ctsAnnotation, StickyNoteAnnotation)
      '    encodedOutput = Me.FormatUnicodeHexString(annotation.TextNote.Text)
      'End If

      'Dim previousParent = doc.DocumentElement.LastChild

      'If encodedOutput IsNot Nothing Then
      '    element.Attributes.Append(Me.CreateAttribute(doc, "Encoding", "unicode"))
      '    element.InnerText = encodedOutput
      'End If

      'previousParent.AppendChild(element)


    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Private Shared Function ParseColor(ByVal lpRGBColor As Integer) As ColorInfo
    Dim lobjResult As ColorInfo
    Try
      ' This isn't accurate
      lobjResult = New ColorInfo With {
        .Red = lpRGBColor And 255,
        .Green = (lpRGBColor >> 8) And 255,
        .Blue = (lpRGBColor >> 16) And 255
      }

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try
    Return lobjResult
  End Function

  Private Shared Function DecodeUnicodeHexString(ByVal lpSource As String) As String
    Dim lstrResult As String = String.Empty
    Try
      If String.IsNullOrEmpty(lpSource) Then Exit Try
      If lpSource.Length Mod 4 <> 0 Then Throw New ArgumentException("The Unicode hexadecimal string length must be evenly divisible by 4", NameOf(lpSource))

      Dim lobjBuilder As New StringBuilder

      For lintPosition As Integer = 0 To lpSource.Length - 1 Step 4
        Dim lintCharValue As Integer = Integer.Parse(lpSource.Substring(lintPosition, 4), System.Globalization.NumberStyles.HexNumber)
        lobjBuilder.Append(ChrW(lintCharValue))
      Next
      lstrResult = lobjBuilder.ToString()
    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try
    Return lstrResult
  End Function

#End Region

#End Region

End Class

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
Imports Documents.Annotations.Common
Imports Documents.Utilities

#End Region

Partial Public Class AnnotationImporter

#Region "Annotation Import Methods"

  ' Line
  ''' <param name="lpXmlDocument">The doc.</param>
  ''' <param name="lpXmlElement">The node.</param>
  ''' <param name="lpCtsAnnotation">The CTS annotation.</param>
  Private Sub ImportLine(ByVal lpXmlDocument As XmlDocument,
                         ByVal lpXmlElement As XmlElement,
                         ByVal lpCtsAnnotation As PointCollectionAnnotation)
    Try

      If lpXmlDocument Is Nothing Then Throw New ArgumentNullException("lpXmlDocument", "Parameter lpXmlDocument can't be null.")
      If lpXmlElement Is Nothing Then Throw New ArgumentNullException("lpXmlElement", "Parameter lpXmlElement can't be null.")
      If lpCtsAnnotation Is Nothing Then Throw New ArgumentNullException("lpCtsAnnotation", "Parameter lpCtsAnnotation can't be null.")

      If lpCtsAnnotation.Points.Count <> 2 Then
        ApplicationLogging.WriteLogEntry(
          String.Format("Error: Line annotation not properly constructed (Content Id={0}, Annotation Id={1}, Points={2})",
                        lpCtsAnnotation.AnnotatedContentId, lpCtsAnnotation.ID, lpCtsAnnotation.Points.Count))
        Exit Sub
      End If

      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_CLASSNAME", "Proprietary"))
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_SUBCLASS", "v1-Line"))
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_CLASSID", "{A91E5DF2-6B7B-11D1-B6D7-00609705F027}"))

      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_LINE_COLOR", Me.FormatColor(lpCtsAnnotation.Display.Foreground)))
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_LINE_START_X", lpCtsAnnotation.Points(0).First / Me.Dpi))
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_LINE_START_Y", lpCtsAnnotation.Points(0).Second / Me.Dpi))
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_LINE_END_X", lpCtsAnnotation.Points(1).First / Me.Dpi))
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_LINE_END_Y", lpCtsAnnotation.Points(1).Second / Me.Dpi))

      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_LINE_WIDTH", lpCtsAnnotation.Display.Border.LineStyle.LineWeight))

      '	Commit <PropDesc> element
      lpXmlDocument.LastChild.AppendChild(lpXmlElement)

      ' <F_CUSTOM_BYTES/>
      Me.ImportCustomBytes(lpXmlDocument, lpCtsAnnotation)

      '	<F_POINTS>0 0 0 0 7 4 11 8 11 13 11 20 15 25 19 31 26 37 30 42 33 47 37 51 45 55 48 59 52 64 63 68 67 72 85 79 93 83 108 88 119 91 134 95 145 98 160 100 171 103 182 105 193 106 200 110 204 115 204 120 208 126 208 130 211 136 211 140 215 144 219 149 219 153 219 157 219 161 219 165 219 170 215 174 215 178 215 182 215 187 215 191 211 195 211 199 211 205 211 209 215 214 215 218 219 222 223 229 226 235 230 239 241 246 245 250 255 255</F_POINTS>
      Me.ImportPoints(lpXmlDocument, lpCtsAnnotation)

      ' <F_TEXT/>
      Me.ImportText(lpXmlDocument, lpCtsAnnotation)

      ' </PropDesc> implied

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Debug.WriteLine(ex.Message)
      Throw
    End Try
    '<FnAnno>                                                                          ' From caller : WriteAnnotationContent
    '    <PropDesc                                                                     ' From caller : WriteAnnotationContent
    '        F_ANNOTATEDID               = "{9B7DC449-0831-467B-AA51-A2013C65D355}"    ' From caller's : ImportCommonMetadata
    '        F_CLASSID                   = "{A91E5DF2-6B7B-11D1-B6D7-00609705F027}"    ' In this method
    '        F_CLASSNAME                 = "Proprietary"                               ' In this method
    '        F_CREATOR                   = "suser"                                     ' ??????????????
    '        F_ENTRYDATE                 = "2011-08-16T17:28:04.0000000-07:00"         ' From caller's : ImportCommonMetadata
    '        F_HEIGHT                    = "2.92"                                      ' From caller's : ImportCommonMetadata
    '        F_ID                        = "{9B7DC449-0831- 467B-AA51-A2013C65D355}"   ' From caller's : ImportCommonMetadata
    '        F_LEFT                      = "4.74"                                      ' From caller's : ImportCommonMetadata
    '        F_LINE_COLOR                = "255"                                       ' In this method
    '        F_LINE_END_X                = "4.74"                                      ' In this method
    '        F_LINE_END_Y                = "1.59"                                      ' In this method
    '        F_LINE_START_X              = "19.06"                                     ' In this method
    '        F_LINE_START_Y              = "4.51"                                      ' In this method
    '        F_LINE_WIDTH                = "16"                                        ' In this method
    '        F_MODIFYDATE                = "2011-08-16T17:28:04.0000000-07:00"         ' From caller's : ImportCommonMetadata
    '        F_MULTIPAGETIFFPAGENUMBER   = "0"                                         ' From caller's : ImportCommonMetadata
    '        F_NAME                      = "-1-3"                                      ' From caller's : ImportCommonMetadata
    '        F_PAGENUMBER                = "1"                                         ' From caller's : ImportCommonMetadata (conditional)
    '        F_SUBCLASS                  = "v1-Line"                                   ' In this method
    '        F_TOP                       = "1.59"                                      ' From caller's : ImportCommonMetadata
    '        F_WIDTH                     ="14.32">                                     ' From caller's : ImportCommonMetadata

    '        <F_CUSTOM_BYTES/>                                                         ' Me.ImportCustomBytes(doc, ctsAnnotation)
    '        <F_POINTS/>                                                               ' Me.ImportPoints(doc, ctsAnnotation)
    '        <F_TEXT/>                                                                 ' Me.ImportText(doc, ctsAnnotation)
    '    </PropDesc>                                                                   ' Implied
    '</FnAnno>
  End Sub

#End Region


End Class

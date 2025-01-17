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

  ' Pen
  ''' <param name="lpXmlDocument">The doc.</param>
  ''' <param name="lpXmlElement">The node.</param>
  ''' <param name="lpCtsAnnotation">The CTS annotation.</param>
  Private Sub Import(ByVal lpXmlDocument As XmlDocument,
                     ByVal lpXmlElement As XmlElement,
                     ByVal lpCtsAnnotation As PointCollectionAnnotation)

    Try

      If lpXmlDocument Is Nothing Then Throw New ArgumentNullException("lpXmlDocument", "Parameter lpXmlDocument can't be null.")
      If lpXmlElement Is Nothing Then Throw New ArgumentNullException("lpXmlElement", "Parameter lpXmlElement can't be null.")
      If lpCtsAnnotation Is Nothing Then Throw New ArgumentNullException("lpCtsAnnotation", "Parameter lpCtsAnnotation can't be null.")

      '	<PropDesc
      '	F_ANNOTATEDID="{F284249D-D541-43D6-A1AD-7D5A4286C876}" F_CLASSID="{5CF11949-018F-11D0-A87A-00A0246922A5}" F_CLASSNAME="Pen" F_ENTRYDATE="2010-07-01T21:34:56.0000000-05:00" F_HEIGHT="0.6033333333333334" F_ID="{F284249D-D541-43D6-A1AD-7D5A4286C876}" F_LEFT="0.16333333333333333" F_LINE_BACKMODE="2" F_LINE_COLOR="65280" F_LINE_STYLE="0" F_LINE_WIDTH="12" F_MODIFYDATE="2010-07-01T21:35:14.0000000-05:00" F_MULTIPAGETIFFPAGENUMBER="0" F_NAME="-1-1" F_PAGENUMBER="1" F_TOP="0.16666666666666666" F_WIDTH="0.23">

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

  End Sub

#End Region


End Class

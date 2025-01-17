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
Imports Documents.Utilities


#End Region

Partial Public Class AnnotationImporter

#Region "Annotation Import Methods"

  ' Arrow
  ''' <param name="lpXmlDocument">The doc.</param>
  ''' <param name="lpXmlElement">The node.</param>
  ''' <param name="lpCtsAnnotation">The CTS annotation.</param>
  Protected Sub Import(ByVal lpXmlDocument As XmlDocument,
                    ByVal lpXmlElement As XmlElement,
                    ByVal lpCtsAnnotation As ArrowAnnotation)

    Try

      If lpXmlDocument Is Nothing Then Throw New ArgumentNullException("lpXmlDocument", "Parameter lpXmlDocument can't be null.")
      If lpXmlElement Is Nothing Then Throw New ArgumentNullException("lpXmlElement", "Parameter lpXmlElement can't be null.")
      If lpCtsAnnotation Is Nothing Then Throw New ArgumentNullException("lpCtsAnnotation", "Parameter lpCtsAnnotation can't be null.")

      '	<PropDesc
      ' F_LINE_BACKMODE="2"
      ' F_LINE_COLOR="16711680"
      ' F_LINE_STYLE="0"
      ' F_LINE_WIDTH="12"
      Me.ImportLineMetadata(lpXmlDocument, lpXmlElement, lpCtsAnnotation)

      ' F_CLASSNAME="Arrow"
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_CLASSNAME", "Arrow"))

      ' F_CLASSID="{5CF11946-018F-11D0-A87A-00A0246922A5}"
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_CLASSID", "{5CF11946-018F-11D0-A87A-00A0246922A5}"))

      ' F_ARROWHEAD_SIZE="1" 
      ' Size is a non-portable enumeration (1,2,3) This is fine for FileNet/Daeja systems but may require a formal enumeration (small, medium, large) later on.
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_ARROWHEAD_SIZE", lpCtsAnnotation.Size))

      ' F_LINE_START_X="1.7966666666666666"
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_LINE_START_X", lpCtsAnnotation.StartPoint.First / Me.Dpi))

      '    F_LINE_START_Y="0.8333333333333334"
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_LINE_START_Y", lpCtsAnnotation.StartPoint.Second / Me.Dpi))

      ' F_LINE_END_X="1.28"
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_LINE_END_X", lpCtsAnnotation.EndPoint.First / Me.Dpi))

      ' F_LINE_END_Y="0.7366666666666667"
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_LINE_END_Y", lpCtsAnnotation.EndPoint.Second / Me.Dpi))

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

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
Imports Documents.Annotations.Special
Imports Documents.Utilities

#End Region

Partial Public Class AnnotationImporter

#Region "Annotation Import Methods"

  ' Stickynote
  ''' <param name="lpXmlDocument">The doc.</param>
  ''' <param name="lpXmlElement">The node.</param>
  ''' <param name="lpCtsAnnotation">The CTS annotation.</param>
  Private Sub Import(ByVal lpXmlDocument As XmlDocument,
                     ByVal lpXmlElement As XmlElement,
                     ByVal lpCtsAnnotation As StickyNoteAnnotation)

    If lpXmlDocument Is Nothing Then Throw New ArgumentNullException("lpXmlDocument", "Parameter lpXmlDocument can't be null.")
    If lpXmlElement Is Nothing Then Throw New ArgumentNullException("lpXmlElement", "Parameter lpXmlElement can't be null.")
    If lpCtsAnnotation Is Nothing Then Throw New ArgumentNullException("lpCtsAnnotation", "Parameter lpCtsAnnotation can't be null.")

    Try
      '	<PropDesc
      ' F_CLASSNAME = "StickyNote"
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_CLASSNAME", "StickyNote"))

      ' F_CLASSID = "{5CF11945-018F-11D0-A87A-00A0246922A5}"
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_CLASSID", "{5CF11945-018F-11D0-A87A-00A0246922A5}"))

      ' F_FORECOLOR = "10092543"
      ' This value is fixed for now, as CS does not export it.
      If lpCtsAnnotation.Display.Background Is Nothing Then
        lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_FORECOLOR", Me.FormatColor(New ColorInfo With {.Red = 255, .Green = 255, .Blue = 153, .Opacity = 100})))
      Else
        lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_FORECOLOR", Me.FormatColor(lpCtsAnnotation.Display.Background)))
      End If

      ' F_ORDINAL
      lpXmlElement.Attributes.Append(Me.CreateAttribute(lpXmlDocument, "F_ORDINAL", lpCtsAnnotation.NoteOrder))

      '	Commit <PropDesc> element
      lpXmlDocument.LastChild.AppendChild(lpXmlElement)

      ' <F_CUSTOM_BYTES/>
      Me.ImportCustomBytes(lpXmlDocument, lpCtsAnnotation)

      ' <F_POINTS/>
      Me.ImportPoints(lpXmlDocument, lpCtsAnnotation)

      ' <F_TEXT Encoding="unicode">0053007400690063006B00790020006E006F00740065</F_TEXT>
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

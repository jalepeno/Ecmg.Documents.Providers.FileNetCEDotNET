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
Imports Documents.Annotations.Common
Imports Documents.Annotations.Highlight
Imports Documents.Annotations.Shape
Imports Documents.Annotations.Special
Imports Documents.Annotations.Text

#End Region

Friend Class SupportedAnnotations

#Region "Class variables"

  Private ReadOnly mobjItems As Collection(Of PlatformAnnotation)
  Private Shared ReadOnly mobjInstance As SupportedAnnotations = New SupportedAnnotations()

#End Region

#Region "Friend Properties"

  Friend ReadOnly Property Items As ReadOnlyCollection(Of PlatformAnnotation)
    Get
      Return New ReadOnlyCollection(Of PlatformAnnotation)(Me.mobjItems)
    End Get
  End Property

  Friend Shared ReadOnly Property Instance As SupportedAnnotations
    Get
      Return mobjInstance
    End Get
  End Property

#End Region

#Region "Constructors"

  Private Sub New()

    Me.mobjItems = New Collection(Of PlatformAnnotation)()
    Me.mobjItems.Add(New PlatformAnnotation("{5CF11946-018F-11D0-A87A-00A0246922A5}", "Arrow", GetType(ArrowAnnotation)))                           'Verified
    Me.mobjItems.Add(New PlatformAnnotation("{5CF1194C-018F-11D0-A87A-00A0246922A5}", "Stamp", GetType(StampAnnotation)))                           'Verified
    Me.mobjItems.Add(New PlatformAnnotation("{5CF11945-018F-11D0-A87A-00A0246922A5}", "StickyNote", GetType(StickyNoteAnnotation)))                 'Verified
    Me.mobjItems.Add(New PlatformAnnotation("{5CF11941-018F-11D0-A87A-00A0246922A5}", "Text", GetType(TextAnnotation)))                             'Verified
    Me.mobjItems.Add(New PlatformAnnotation("{5CF11942-018F-11D0-A87A-00A0246922A5}", "Highlight", GetType(HighlightRectangle)))                    'Verified
    Me.mobjItems.Add(New PlatformAnnotation("{A91E5DF2-6B7B-11D1-B6D7-00609705F027}", "Proprietary", "v1-Rectangle", GetType(RectangleAnnotation))) 'Verified
    Me.mobjItems.Add(New PlatformAnnotation("{A91E5DF2-6B7B-11D1-B6D7-00609705F027}", "Proprietary", "v1-Line", GetType(PointCollectionAnnotation)))
    Me.mobjItems.Add(New PlatformAnnotation("{A91E5DF2-6B7B-11D1-B6D7-00609705F027}", "Proprietary", "v1-Oval", GetType(EllipseAnnotation)))
    Me.mobjItems.Add(New PlatformAnnotation("{5CF11949-018F-11D0-A87A-00A0246922A5}", "Pen", GetType(PointCollectionAnnotation)))

  End Sub

#End Region

End Class


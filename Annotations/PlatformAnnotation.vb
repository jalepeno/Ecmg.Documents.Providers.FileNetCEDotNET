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
Imports Documents.Utilities

#End Region

Friend Class PlatformAnnotation

#Region "Class Variables"

  Private ReadOnly mstrClassName As String
  Private ReadOnly mstrClassId As String
  Private ReadOnly mstrSubclassName As String
  Private ReadOnly mobjAnnotationType As Type

#End Region

#Region "Public Properties"

  Friend ReadOnly Property ClassName As String
    Get
      Return Me.mstrClassName
    End Get
  End Property

  Friend ReadOnly Property SubClassName As String
    Get
      Return Me.mstrSubclassName
    End Get
  End Property

  Friend ReadOnly Property ClassId As String
    Get
      Return Me.mstrClassId
    End Get
  End Property

  Friend ReadOnly Property AnnotationType As Type
    Get
      Return Me.mobjAnnotationType
    End Get
  End Property

#End Region

#Region "Constructors"

  Public Sub New(ByVal lpClassId As String, lpClassName As String, lpAnnotationType As Type)
    Try

      If String.IsNullOrEmpty(lpClassId) Then Throw New ArgumentNullException("lpClassId")
      If String.IsNullOrEmpty(lpClassName) Then Throw New ArgumentNullException("lpClassName")
      If lpAnnotationType Is Nothing Then Throw New ArgumentNullException("lpAnnotationType")

      Me.mstrClassId = lpClassId
      Me.mstrClassName = lpClassName
      Me.mobjAnnotationType = lpAnnotationType

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub New(ByVal lpClassId As String, lpClassName As String, lpSubclassName As String, lpAnnotationType As Type)
    Try

      If String.IsNullOrEmpty(lpClassId) Then Throw New ArgumentNullException("lpClassId")
      If String.IsNullOrEmpty(lpClassName) Then Throw New ArgumentNullException("lpClassName")
      If String.IsNullOrEmpty(lpSubclassName) Then Throw New ArgumentNullException("lpSubclassName")
      If lpAnnotationType Is Nothing Then Throw New ArgumentNullException("lpAnnotationType")

      Me.mstrClassId = lpClassId
      Me.mstrClassName = lpClassName
      Me.mobjAnnotationType = lpAnnotationType
      Me.mstrSubclassName = lpSubclassName

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

End Class

'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Ecmg.Cts.Annotations.Auditing
Imports Ecmg.Cts.Annotations.Decoration
Imports Ecmg.Cts.Annotations.Common
Imports Ecmg.Cts.Core
Imports Ecmg.Cts.Utilities
Imports Ecmg.Cts.Annotations
Imports Ecmg.Cts.Annotations.Text
Imports Ecmg.Cts.Annotations.Special
Imports Ecmg.Cts.Annotations.Highlight
Imports Ecmg.Cts.Annotations.Decoration.LineStyleInfo
Imports Ecmg.Cts.Annotations.Decoration.PointStyle

Imports System.IO
Imports System.Xml
Imports System.Collections.ObjectModel
Imports Ecmg.Cts.Annotations.Shape
Imports Ecmg.Cts.Annotations.Exception
Imports System.Text
Imports Documents.Utilities


#End Region

Friend Class CtsXmlDocument
  Inherits XmlDocument

#Region "XML Helpers"

  ''' <summary>
  ''' Queries the XML response for the existence of a particular node.
  ''' </summary>
  ''' <param name="lpXPath">The XML path.</param>
  ''' <returns>True if exists, otherwise false.</returns>
  Public Function QueryExists(ByVal lpXPath As String) As Boolean

    Try
      If String.IsNullOrEmpty(lpXPath) Then Throw New ArgumentNullException("lpXPath")

      Dim lobjNode As XmlNodeList = Me.SelectNodes(lpXPath)
      If lobjNode Is Nothing OrElse lobjNode.Item(0) Is Nothing Then
        Return False
      End If

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    End Try

    Return True

  End Function

  ''' <summary>
  ''' Queries the XML for a single string.
  ''' </summary>
  ''' <param name="lpXPath">The XML path.</param>
  ''' <returns>The text of a single node.</returns>
  Public Function QuerySingleString(ByVal lpXPath As String) As String

    Dim lstrResult As String = Nothing
    Try

      If String.IsNullOrEmpty(lpXPath) Then Throw New ArgumentNullException("lpXPath")

      Dim lobjNode As XmlNode = Me.SelectSingleNode(lpXPath)
      lstrResult = lobjNode.InnerText

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    End Try

    Return lstrResult

  End Function

  Public Function QuerySingleBoolean(ByVal lpXPath As String) As Boolean
    Dim lblnResult As Boolean = False

    Try
      If String.IsNullOrEmpty(lpXPath) Then Throw New ArgumentNullException("lpXPath")

      Dim lstrValue As String = Me.QuerySingleString(lpXPath)
      If String.IsNullOrEmpty(lstrValue) Then Exit Try
      If lstrValue.ToUpperInvariant().CompareTo("TRUE") = 0 Then lblnResult = True
    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    End Try

    Return lblnResult
  End Function

  Public Function QuerySingleDate(ByVal lpXPath As String) As Date
    Dim lobjResult As Date

    Try
      If String.IsNullOrEmpty(lpXPath) Then Throw New ArgumentNullException("lpXPath")

      Dim lstrValue As String = Me.QuerySingleString(lpXPath)

      If String.IsNullOrEmpty(lstrValue) Then Exit Try

      If Not Date.TryParse(lstrValue, lobjResult) Then
        ApplicationLogging.WriteLogEntry("Could not parse date {0}", lstrValue)
      End If
    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    End Try

    Return lobjResult
  End Function

  Public Function QuerySingleInteger(ByVal lpXPath As String) As Integer
    Dim lobjResult As Integer

    Try
      If String.IsNullOrEmpty(lpXPath) Then Throw New ArgumentNullException("lpXPath")

      Dim lstrValue As String = Me.QuerySingleString(lpXPath)
      If String.IsNullOrEmpty(lstrValue) Then Exit Try
      If Not Integer.TryParse(lstrValue, lobjResult) Then
        ApplicationLogging.WriteLogEntry("Could not parse integer {0}", lstrValue)
      End If
    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    End Try

    Return lobjResult
  End Function

  ''' <summary>
  ''' Queries the XML for the text of an attribute.
  ''' </summary>
  ''' <param name="lpXPath">The XML path.</param>
  ''' <param name="lpAttribute">The XML attribute.</param>
  ''' <returns>The text of the specified attribute.</returns>
  Public Function QuerySingleAttribute(ByVal lpXPath As String, ByVal lpAttribute As String) As String

    Dim lstrResult As String = Nothing

    Try
      If String.IsNullOrEmpty(lpXPath) Then Throw New ArgumentNullException("lpXPath")
      If String.IsNullOrEmpty(lpAttribute) Then Throw New ArgumentNullException("lpAttribute")

      Dim lobjNode As XmlNode = Me.SelectSingleNode(lpXPath)
      If lobjNode Is Nothing Then Exit Try

      Dim lobjItem As XmlAttribute = lobjNode.Attributes(lpAttribute)
      If lobjItem Is Nothing Then Exit Try

      lstrResult = lobjItem.Value

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    End Try

    Return lstrResult

  End Function

  Public Function QuerySingleAttributeAsDate(ByVal lpXPath As String, ByVal lpAttribute As String) As Date
    Dim lobjResult As Date = Nothing
    Try
      If String.IsNullOrEmpty(lpXPath) Then Throw New ArgumentNullException("lpXPath")
      If String.IsNullOrEmpty(lpAttribute) Then Throw New ArgumentNullException("lpAttribute")

      lobjResult = New Date
      Dim lstrValue As String = Me.QuerySingleAttribute(lpXPath, lpAttribute)
      If Not Date.TryParse(lstrValue, lobjResult) Then
        ApplicationLogging.WriteLogEntry("Could not parse date {0}", lstrValue)
      End If

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    End Try
    Return lobjResult
  End Function

  Public Function QuerySingleAttributeAsInteger(ByVal lpXPath As String, ByVal lpAttribute As String) As Integer
    Dim lintResult As Integer
    Try
      If String.IsNullOrEmpty(lpXPath) Then Throw New ArgumentNullException("lpXPath")
      If String.IsNullOrEmpty(lpAttribute) Then Throw New ArgumentNullException("lpAttribute")

      Dim lstrValue As String = Me.QuerySingleAttribute(lpXPath, lpAttribute)
      If Not Integer.TryParse(lstrValue, lintResult) Then
        ApplicationLogging.WriteLogEntry("Could not parse integer {0}", lstrValue)
      End If
    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    End Try
    Return lintResult
  End Function

  Public Function QuerySingleAttributeAsSingle(ByVal lpXPath As String, ByVal lpAttribute As String) As Single
    Dim lintResult As Single
    Try
      If String.IsNullOrEmpty(lpXPath) Then Throw New ArgumentNullException("lpXPath")
      If String.IsNullOrEmpty(lpAttribute) Then Throw New ArgumentNullException("lpAttribute")

      Dim lstrValue As String = Me.QuerySingleAttribute(lpXPath, lpAttribute)
      If Not Single.TryParse(lstrValue, lintResult) Then
        ApplicationLogging.WriteLogEntry("Could not parse single {0}", lstrValue)
      End If
    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    End Try
    Return lintResult
  End Function

  Public Function QuerySingleAttributeAsBoolean(ByVal lpXPath As String, ByVal lpAttribute As String) As Boolean
    Dim lblnResult As Boolean
    Try
      If String.IsNullOrEmpty(lpXPath) Then Throw New ArgumentNullException("lpXPath")
      If String.IsNullOrEmpty(lpAttribute) Then Throw New ArgumentNullException("lpAttribute")

      Dim lstrValue As String = Me.QuerySingleAttribute(lpXPath, lpAttribute)
      If String.IsNullOrEmpty(lstrValue) Then Exit Try
      If lstrValue.ToUpperInvariant().CompareTo("TRUE") = 0 Then lblnResult = True

    Catch ex As System.Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    End Try
    Return lblnResult
  End Function

#End Region

End Class


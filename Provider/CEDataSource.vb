'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Data
Imports Documents.Utilities

#End Region

Public Class CEDataSource
  Inherits DataSource

#Region "Constructors"

  Public Sub New()
    MyBase.New()
  End Sub

  Public Sub New(ByVal lpConnectionString As String,
                 ByVal lpQueryTarget As String,
                 ByVal lpSourceColumn As String,
                 ByVal lpCriteria As Criteria)

    MyBase.New(lpConnectionString, lpQueryTarget, lpSourceColumn, lpCriteria)

  End Sub

  Public Sub New(ByVal lpXMLFilePath As String)
    MyBase.New(lpXMLFilePath)
  End Sub

#End Region

#Region "Public Methods"

  Public Overrides Function BuildSQLString(Optional ByRef lpErrorMessage As String = "") As String

    Try

      If (Me.SearchType = SearchType.ContentSearch) Then
        Return BuildSQLStringContent(lpErrorMessage)

      Else
        Return MyBase.BuildSQLString(lpErrorMessage)
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

#End Region

#Region "Private Methods"

  ''' <summary>
  ''' 
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function BuildSQLStringContent() As String

    Try

      Dim lstrSQL As String = String.Empty

      lstrSQL = "SELECT "

      If (Me.LimitResults > 0) Then
        lstrSQL &= " TOP " & Me.LimitResults.ToString() & " "
      End If

      lstrSQL &= ResultColumnsString & " FROM [" & QueryTarget & "] d INNER JOIN VerityContentSearch v ON d.This = v.QueriedObject"

      lstrSQL &= " WHERE CONTAINS(v.Content,'"
      lstrSQL &= Me.ContentKeywords
      lstrSQL &= "') "

      Return lstrSQL

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

#End Region

End Class

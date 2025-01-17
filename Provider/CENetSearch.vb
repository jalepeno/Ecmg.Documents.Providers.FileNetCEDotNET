'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports FileNet.Api.Core
Imports FileNet.Api.Query
Imports FileNet.Api.Collection
Imports Documents.Arguments
Imports Documents.Core
Imports Documents.Providers
Imports Documents.Utilities
Imports DProviders = Documents.Providers
Imports DCore = Documents.Core
Imports Documents

#End Region

Public Class CENetSearch
  Inherits CSearch
  Implements ISQLPassThroughSearch

#Region "Class Variables"

  'Private mobjDataSource As New CEDataSource

  Private mstrURL As String
  Private mstrUserName As String
  Private mstrPassword As String
  Private mobjDomain As IDomain
  Private mstrObjectStoreName As String

#End Region

#Region "Class Constants"

  Public Shadows Const ID_COLUMN As String = "Id"
  Public Shadows Const DOCUMENT_QUERY_TARGET As String = "Document"

#End Region

#Region "Constructors"

  Public Sub New()

  End Sub

  Public Sub New(ByVal lpProvider As CENetProvider)
    Me.New(lpProvider, New Data.Criteria)
  End Sub

  Public Sub New(ByVal lpProvider As CENetProvider,
                 ByVal lpCriteria As Data.Criteria)
    MyBase.New(CType(lpProvider, CProvider), lpCriteria, ID_COLUMN, DOCUMENT_QUERY_TARGET)

    Try

      Dim lobjProvider As Object = lpProvider

      If lpProvider.IsConnected = False AndAlso lpProvider.ContentSource IsNot Nothing Then
        lpProvider.Connect(lpProvider.ContentSource)
      End If

      UserName = lpProvider.ProviderProperties("UserName").PropertyValue
      Password = lpProvider.ProviderProperties("Password").PropertyValue
      ObjectStoreName = lpProvider.ProviderProperties("ObjectStore").PropertyValue

      Domain = lpProvider.Domain
      URL = lpProvider.URL

    Catch ex As Exception
      ' ApplicationLogging.LogException(ex, "CENetProvider::New(lpProvider, lpCriteria)")
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Sub

#End Region

#Region "Public Properties"

  Public Property URL() As String
    Get
      Return mstrURL
    End Get
    Set(ByVal value As String)
      mstrURL = value
    End Set
  End Property

  Public Property UserName() As String
    Get
      Return mstrUserName
    End Get
    Set(ByVal value As String)
      mstrUserName = value
    End Set
  End Property

  Protected Property Password() As String
    Get
      Return mstrPassword
    End Get
    Set(ByVal value As String)
      mstrPassword = value
    End Set
  End Property

  Public Property ObjectStoreName() As String
    Get
      Return mstrObjectStoreName
    End Get
    Set(ByVal value As String)
      mstrObjectStoreName = value
    End Set
  End Property

  Public Property Domain() As IDomain
    Get
      Return mobjDomain
    End Get
    Set(ByVal value As IDomain)
      mobjDomain = value
    End Set
  End Property

  Protected Overrides ReadOnly Property DefaultDelimitedResultColumns As String
    Get

      Try
        Return "Id,DocumentTitle"

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Get
  End Property

  Public Overrides ReadOnly Property DefaultQueryTarget As String
    Get

      Try
        Return DOCUMENT_QUERY_TARGET

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Get
  End Property

#End Region

#Region "Public Methods"

#End Region

#Region "Public Overrides Methods"

#Region "ISQLPassThroughSearch Implementation"

  Public Overloads Function Execute(Sql As String) As DCore.SearchResultSet Implements ISQLPassThroughSearch.Execute
    Try

      Dim lobjObjectStore As IObjectStore = Factory.ObjectStore.GetInstance(Me.Domain, Me.ObjectStoreName)
      Dim lobjSearchScope As New SearchScope(lobjObjectStore)
      Dim lobjSQLObject As New SearchSQL()
      lobjSQLObject.SetQueryString(Sql)

      Dim lobjValues As New Core.Values
      Dim lobjSearchResults As New Core.SearchResults
      Dim lobjSearchResultSet As Core.SearchResultSet

      Dim lobjRowSet As IRepositoryRowSet = lobjSearchScope.FetchRows(lobjSQLObject, Nothing, Nothing, True)
      Dim lstrId As String = Nothing
      Dim lobjEcmProperty As ECMProperty = Nothing
      Dim lobjDataItems As Data.DataItems = Nothing
      Dim lobjDataItem As Data.DataItem = Nothing

      For Each lobjRow As IRepositoryRow In lobjRowSet
        lobjDataItems = New Data.DataItems
        lobjEcmProperty = Nothing
        lstrId = lobjRow.Properties.GetProperty("Id").GetIdValue().ToString
        For Each lobjP8ObjectProperty As FileNet.Api.Property.IProperty In lobjRow.Properties
          lobjEcmProperty = CType(Me.Provider, CENetProvider).CreateECMProperty(lobjP8ObjectProperty)
          lobjDataItems.Add(New Data.DataItem(lobjEcmProperty))
        Next
        lobjSearchResults.Add(New Core.SearchResult(lstrId, lobjDataItems))
        Debug.Print("Added {0}", lstrId)
      Next

      lobjSearchResultSet = New Core.SearchResultSet(lobjSearchResults)

      Return lobjSearchResultSet

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

  Public Overrides Function Execute(ByVal Args As SearchArgs) As SearchResultSet

    Dim lstrErrorMessage As String = String.Empty
    Dim lobjSearchResultSet As Core.SearchResultSet

    Try
      InitializeSearch(Args)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, "CENetProvider::ExecuteSearch(Args)_InitializeSearch")
      lobjSearchResultSet = New Core.SearchResultSet(ex)
      Return lobjSearchResultSet
    End Try

    Try

      ' Initialize the DataSource with the correct QueryTarget and IDColumn
      If Args.Document IsNot Nothing Then
        InitializeDataSource(CENetSearch.ID_COLUMN, Args.Document.DocumentClass)

      Else
        InitializeDataSource(CENetSearch.ID_COLUMN, CENetSearch.DOCUMENT_QUERY_TARGET)
      End If

      ' If the Criteria includes the Document Class we need to remove it.
      ' It will be set as the query target in this case.
      For Each lobjCriterion As Data.Criterion In Criteria

        If lobjCriterion.PropertyName = "Document Class" Then
          Me.Criteria.Remove(lobjCriterion)
          Exit For
        End If

      Next

      ''DataSource.Criteria = Me.Criteria

      ' Copy the document object
      Dim lobjDocument As Core.Document = Args.Document

      ' Clear the document class so that it will not be parsed as a where clause.
      ' By passing it to the InitializeDataSource method above we effectively
      ' set the query target to the document class.
      If lobjDocument IsNot Nothing Then
        lobjDocument.DocumentClass = ""

      Else
        ' Let's see what happens
      End If

      Dim lstrSQL As String

      If Args.UseDocumentValuesInCriteriaValues Then
        lstrSQL = Me.DataSource.BuildSQLString(lobjDocument, Args.VersionIndex, lstrErrorMessage)

      Else
        lstrSQL = Me.DataSource.BuildSQLString(lstrErrorMessage)
      End If

      Me.DataSource.SQLStatement = lstrSQL

#If DEBUG Then
      ' Write the SQL statement to the log for debugging
      ApplicationLogging.WriteLogEntry(String.Format("CENetProvider::ExecuteSearch SQL Initialized as '{0}'", lstrSQL), TraceEventType.Information, 9411)
#End If

      If lstrErrorMessage.Length > 0 Then
        lobjSearchResultSet = New Core.SearchResultSet(New ApplicationException("Error Creating SQL Statement: " & lstrErrorMessage))
        Return lobjSearchResultSet
      End If

      Try
        lobjSearchResultSet = GetDocumentIDSet(Me, lstrErrorMessage)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, "CENetProvider::ExecuteSearch(Args)_GetDocumentIDSet")
        Args.ErrorMessage += String.Format("Exception: '{0}': SQL Statement: '{1}'", ex.Message, lstrSQL)
        Throw New ApplicationException(String.Format("Unable to Execute Search: {0}, ErrorMessage: {1}", lstrSQL, Args.ErrorMessage))
      End Try

      If lobjSearchResultSet.Exception IsNot Nothing Then
        lobjSearchResultSet = New Core.SearchResultSet(New Exception(String.Format("Unable to Execute Search: '{0}'", lstrSQL), lobjSearchResultSet.Exception))
        Return lobjSearchResultSet
      End If

      If lstrErrorMessage.Length > 0 Then
        lobjSearchResultSet = New Core.SearchResultSet(New Exception(String.Format("Unable to Execute Search: '{0}'", lstrSQL) & lstrErrorMessage))
        Return lobjSearchResultSet
      End If

      Return lobjSearchResultSet

    Catch ex As Exception
      ApplicationLogging.LogException(ex, "CENetProvider::ExecuteSearch(Args)")
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  Public Overrides Function Execute() As DCore.SearchResultSet

    Dim lstrErrorMessage As String = String.Empty
    Dim lobjSearchResultSet As DCore.SearchResultSet

    Try


      InitializeSearch()

    Catch ex As Exception
      ApplicationLogging.LogException(ex, "CENetProvider::ExecuteSearch()_InitializeSearch")
      lobjSearchResultSet = New DCore.SearchResultSet(ex)
      Return lobjSearchResultSet
    End Try

    Try

      ' '' Initialize the DataSource with the correct QueryTarget and IDColumn
      ''InitializeDataSource(Me.ID_COLUMN, Me.DOCUMENT_QUERY_TARGET)

      ' If the Criteria includes the Document Class we need to remove it.
      ' It will be set as the query target in this case.
      For Each lobjCriterion As Data.Criterion In Criteria

        If lobjCriterion.PropertyName = "Document Class" Then
          Me.Criteria.Remove(lobjCriterion)
          Exit For
        End If

      Next

      If lstrErrorMessage.Length > 0 Then
        lobjSearchResultSet = New DCore.SearchResultSet(New ApplicationException("Error Creating SQL Statement: " & lstrErrorMessage))
        Return lobjSearchResultSet
      End If

      lobjSearchResultSet = ExecuteSearch(Me, lstrErrorMessage)

      If lstrErrorMessage.Length > 0 Then
        lobjSearchResultSet = New DCore.SearchResultSet(New Exception("Unable to Execute Search: " & lstrErrorMessage))
        Return lobjSearchResultSet
      End If

      Return lobjSearchResultSet

    Catch ex As Exception
      ApplicationLogging.LogException(ex, "CENetProvider::ExecuteSearch()")
      lobjSearchResultSet = New DCore.SearchResultSet(New Exception("Unable to Execute Search: " & ex.Message))
      Return lobjSearchResultSet
    End Try

  End Function

  Public Overrides Function SimpleSearch(ByVal Args As SimpleSearchArgs) As System.Data.DataTable

    Try

      Dim lobjArgs As New SearchEventArgs()
      lobjArgs.Results(0) = Nothing
      lobjArgs.Exception = Nothing
      lobjArgs.UserState = "Search not implemented"
      Me.RaiseSearchCompleteEvent(Me, lobjArgs)

      Return Nothing

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

#End Region

#Region "Private Methods"

  ''' <summary>
  ''' ExecuteSearch
  ''' </summary>
  ''' <param name="lpSearch"></param>
  ''' <param name="lpErrorMessage"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function ExecuteSearch(ByVal lpSearch As DProviders.ISearch,
                                 Optional ByRef lpErrorMessage As String = "") As DCore.SearchResultSet

    Try

      Dim lobjResultSet As New DCore.SearchResultSet

      Dim lobjObjectStore As IObjectStore = Factory.ObjectStore.GetInstance(Me.Domain, Me.ObjectStoreName)
      Dim lobjSearchScope As New SearchScope(lobjObjectStore)
      Dim lobjSQLObject As New SearchSQL()
      lobjSQLObject.SetQueryString(lpSearch.DataSource.SQLStatement)

      Dim lobjRowSet As IRepositoryRowSet = lobjSearchScope.FetchRows(lobjSQLObject, Nothing, Nothing, True)

      lobjResultSet = BuildSearchResultSet(lobjRowSet)

      Return lobjResultSet

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  ''' <summary>
  ''' BuildSearchResultSet
  ''' </summary>
  ''' <param name="lobjRowSet"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Shared Function BuildSearchResultSet(ByVal lobjRowSet As IRepositoryRowSet) As DCore.SearchResultSet

    Try

      Dim lobjSearchResultSet As New DCore.SearchResultSet

      ' Make sure we have some actual data to return
      If lobjRowSet Is Nothing Then
        ' Simply return the empty result set
        Return lobjSearchResultSet
      End If

      Dim lobjSearchResult As DCore.SearchResult

      ' Iterate through each object returned from the search

      For Each lobjRow As IRepositoryRow In lobjRowSet

        lobjSearchResult = New DCore.SearchResult()

        'Loop through each property
        For Each lobjProperty As FileNet.Api.Property.IProperty In lobjRow.Properties

          If (lobjProperty.GetPropertyName.ToLower = "id") Then
            'Dim lobjDocId As Id = lobjRow.Properties.GetProperty("Id").GetIdValue()
            lobjSearchResult.ID = lobjRow.Properties.GetProperty(lobjProperty.GetPropertyName).GetIdValue().ToString
          End If

          ' Add the property to the search result
          lobjSearchResult.Values.Add(GetDataItem(lobjProperty))

        Next

        lobjSearchResultSet.Results.Add(lobjSearchResult)

      Next

      Return lobjSearchResultSet

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  ''' <summary>
  ''' GetDataItem
  ''' </summary>
  ''' <param name="lpResultProperty"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Shared Function GetDataItem(ByVal lpResultProperty As FileNet.Api.Property.IProperty) As Data.DataItem

    Try

      Dim lobjDataItem As Data.DataItem = Nothing

      Select Case lpResultProperty.GetType.Name

        Case "PropertyIdImpl"

          If lpResultProperty.GetPropertyName.ToLower = "id" Then
            lobjDataItem = New Data.DataItem(lpResultProperty.GetPropertyName, Core.PropertyType.ecmGuid, lpResultProperty.GetIdValue, True)

          Else
            lobjDataItem = New Data.DataItem(lpResultProperty.GetPropertyName, Core.PropertyType.ecmGuid, lpResultProperty.GetIdValue, False)
          End If

        Case "PropertyStringImpl"
          lobjDataItem = New Data.DataItem(lpResultProperty.GetPropertyName, Core.PropertyType.ecmString, lpResultProperty.GetStringValue, False)

        Case "PropertyBooleanImpl"
          lobjDataItem = New Data.DataItem(lpResultProperty.GetPropertyName, Core.PropertyType.ecmBoolean, lpResultProperty.GetBooleanValue, False)

        Case "PropertyDateTimeImpl"
          lobjDataItem = New Data.DataItem(lpResultProperty.GetPropertyName, Core.PropertyType.ecmDate, lpResultProperty.GetDateTimeValue, False)

        Case "PropertyFloat64Impl"
          lobjDataItem = New Data.DataItem(lpResultProperty.GetPropertyName, Core.PropertyType.ecmDouble, lpResultProperty.GetFloat64Value, False)

        Case "PropertyInteger32Impl"
          lobjDataItem = New Data.DataItem(lpResultProperty.GetPropertyName, Core.PropertyType.ecmLong, lpResultProperty.GetInteger32Value, False)

        Case "PropertyEngineObjectImpl"
          lobjDataItem = New Data.DataItem(lpResultProperty.GetPropertyName, Core.PropertyType.ecmObject, lpResultProperty.GetObjectValue, False)

        Case "PropertyBinaryImpl"
          lobjDataItem = New Data.DataItem(lpResultProperty.GetPropertyName, Core.PropertyType.ecmBinary, lpResultProperty.GetBinaryValue, False)

      End Select

      Return lobjDataItem

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  ''' <summary>
  ''' GetDocumentIDSet
  ''' </summary>
  ''' <param name="lpSearch"></param>
  ''' <param name="lpErrorMessage"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function GetDocumentIDSet(ByVal lpSearch As DProviders.ISearch,
                                    Optional ByRef lpErrorMessage As String = "") As DCore.SearchResultSet

    Dim lobjSearchResults As New DCore.SearchResults
    Dim lobjSearchResultSet As DCore.SearchResultSet

    Try

      Dim lobjValues As DCore.Values = GetDocumentIDs(lpSearch, lpErrorMessage)

      For Each lstrValue As String In lobjValues
        lobjSearchResults.Add(New DCore.SearchResult(lstrValue))
      Next

      lobjSearchResultSet = New DCore.SearchResultSet(lobjSearchResults)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      lpErrorMessage += ex.Message
      lobjSearchResultSet = New DCore.SearchResultSet(lobjSearchResults, ex)
    End Try

    Return lobjSearchResultSet

  End Function

  ''' <summary>
  ''' GetDocumentIDs
  ''' </summary>
  ''' <param name="lpSearch"></param>
  ''' <param name="lpErrorMessage"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function GetDocumentIDs(ByVal lpSearch As DProviders.ISearch,
                                  Optional ByRef lpErrorMessage As String = "") As DCore.Values

    Try

      Dim lobjValues As New DCore.Values
      Dim lstrErrorMessage As String = String.Empty

#If DEBUG Then

      Dim lstrDebugMessage As String = String.Format("Performing an ExecuteSearch, SQL: {0}", lpSearch.DataSource.SQLStatement)
      Debug.WriteLine(lstrDebugMessage)
      ApplicationLogging.WriteLogEntry(lstrDebugMessage, TraceEventType.Information, 4693)
#End If

      Dim lobjObjectStore As IObjectStore = Factory.ObjectStore.GetInstance(Me.Domain, Me.ObjectStoreName)
      Dim lobjSearchScope As New SearchScope(lobjObjectStore)
      Dim lobjSQLObject As New SearchSQL()
      lobjSQLObject.SetQueryString(lpSearch.DataSource.SQLStatement)

      Dim lobjRowSet As IRepositoryRowSet = lobjSearchScope.FetchRows(lobjSQLObject, Nothing, Nothing, True)

      For Each lobjRow As IRepositoryRow In lobjRowSet
        lobjValues.Add(lobjRow.Properties.GetProperty("Id").GetIdValue().ToString)
      Next

      Return lobjValues

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

#End Region

End Class

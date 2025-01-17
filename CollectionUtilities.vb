'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System
Imports System.IO
Imports Microsoft.Web.Services3
Imports FileNet.Api
Imports FileNet.Api.Core
Imports FileNet.Api.Property
Imports FileNet.Api.Collection
Imports Documents.Utilities


#End Region

Public Class CollectionCounter

#Region "Class Variables"

  Private mobjObjectStore As IObjectStore = Nothing
  Private mobjObjectSet As IIndependentObjectSet = Nothing
  Private mintPageSize As Int16 = 0
  Private mlngItemCount As Long = -1

#End Region

#Region "Public Properties"

  Private ReadOnly Property ObjectStore() As IObjectStore
    Get
      Return mobjObjectStore
    End Get
  End Property

  Public ReadOnly Property Count() As Long
    Get

      Try

        If mlngItemCount = -1 Then
          mlngItemCount = GetCount()
        End If

        Return mlngItemCount

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Get
  End Property

  Private Property PageSize() As Int16
    Get
      Return mintPageSize
    End Get
    Set(ByVal value As Int16)
      mintPageSize = value
    End Set
  End Property

#End Region

#Region "Constructors"

  Public Sub New(ByVal lpObjectStore As IObjectStore,
                 ByVal lpCollection As IIndependentObjectSet,
                 Optional ByVal lpPageSize As Int16 = 10)

    Try
      mobjObjectStore = lpObjectStore
      mobjObjectSet = lpCollection
      mintPageSize = lpPageSize

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Sub

#End Region

#Region "Public Methods"

  Private Function GetCount() As Long

    Try

      ' Iterate through the items
      'Dim lintPageCount As Int16 = 0
      Dim lobjPageEnumerator As IPageEnumerator = mobjObjectSet.GetPageEnumerator
      Dim lintElementCount As Integer = 0
      Dim llngTotalItemCount As Long = 0

      lobjPageEnumerator.PageSize = PageSize

      While lobjPageEnumerator.NextPage = True
        ' Get the number of objects on this page
        'lintPageCount += 1
        lintElementCount = lobjPageEnumerator.ElementCount
        llngTotalItemCount += lintElementCount
      End While

      Return llngTotalItemCount

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

#End Region

End Class

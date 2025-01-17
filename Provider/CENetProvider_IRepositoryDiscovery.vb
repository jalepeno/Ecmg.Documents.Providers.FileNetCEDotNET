'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  CENetProvider_IRepositoryDiscovery.vb
'   Description :  [type_description_here]
'   Created     :  4/8/2014 11:08:49 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports Ecmg.Cts.Utilities
Imports Ecmg.Cts.Core
Imports Ecmg.Cts.Arguments
Imports FileNet.Api.Core
Imports Ecmg.Cts.Exceptions
Imports FileNet.Api.Collection
Imports FileNet.Api.Property
Imports Ecmg.Cts.Migrations
Imports FileNet.Api.Constants
Imports System.IO
Imports Documents.Core
Imports Documents.Providers
Imports Documents.Utilities
Imports DCore = Documents.Core

#End Region

Partial Public Class CENetProvider
  Implements IRepositoryDiscovery

#Region "IRepositoryDiscovery Implementation"

  Public Function GetRepositories() As DCore.RepositoryIdentifiers _
         Implements IRepositoryDiscovery.GetRepositories

    Try
      InitializeProperties()
      Return GetObjectStoreIdentifiers()

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  Private Function GetObjectStoreIdentifiers() As RepositoryIdentifiers

    Try

      Dim lstrErrorMessage As String = String.Empty
      Dim lobjRepositories As New RepositoryIdentifiers

      InitializeConnection()

      For Each lobjObjectStore As IObjectStore In Domain.ObjectStores
        lobjRepositories.Add(New RepositoryIdentifier(lobjObjectStore.Id.ToString, lobjObjectStore.Name, lobjObjectStore.DisplayName))
      Next

      Return lobjRepositories

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

#End Region
End Class

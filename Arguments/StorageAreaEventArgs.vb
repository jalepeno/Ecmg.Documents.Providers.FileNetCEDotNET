
#Region "Imports"

Imports FileNet.Api.Admin
Imports Documents.Arguments
Imports Documents.Utilities


#End Region

Public Class StorageAreaEventArgs
  Inherits CtsEventArgs

#Region "Class Variables"

  Private mobjEnvironment As Environment
  Private mobjObjectStoreStatus As ObjectStoreRollingPolicyStatus
  Private mobjStorageArea As IStorageArea
  Private mobjStoragePolicy As IStoragePolicy

#End Region

#Region "Public Properties"

  Public ReadOnly Property Environment As Environment
    Get
      Return mobjEnvironment
    End Get
  End Property

  Public ReadOnly Property ObjectStoreStatus As ObjectStoreRollingPolicyStatus
    Get
      Return mobjObjectStoreStatus
    End Get
  End Property

  Public ReadOnly Property StorageArea As IStorageArea
    Get
      Return mobjStorageArea
    End Get
  End Property

  Public ReadOnly Property StoragePolicy As IStoragePolicy
    Get
      Return mobjStoragePolicy
    End Get
  End Property

#End Region

#Region "Constructors"

  Public Sub New(ByVal lpEnvironment As Environment, ByVal lpStorageArea As IStorageArea, ByVal lpStoragePolicy As IStoragePolicy)
    Me.New(lpEnvironment, lpStorageArea, lpStoragePolicy, Nothing)
  End Sub

  Public Sub New(ByVal lpEnvironment As Environment, ByVal lpStorageArea As IStorageArea, ByVal lpStoragePolicy As IStoragePolicy, ByVal lpObjectStoreStatus As ObjectStoreRollingPolicyStatus)
    Try
      mobjEnvironment = lpEnvironment
      mobjStorageArea = lpStorageArea
      mobjStoragePolicy = lpStoragePolicy
      mobjObjectStoreStatus = lpObjectStoreStatus
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

End Class

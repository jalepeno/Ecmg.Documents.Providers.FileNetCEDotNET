#Region "Imports"

Imports FileNet.Api.Admin

#End Region

Public Class StorageAreaAddedEventArgs
  Inherits StorageAreaEventArgs

#Region "Constructors"

  Public Sub New(ByVal lpEnvironment As Environment, ByVal lpStorageArea As IStorageArea, ByVal lpStoragePolicy As IStoragePolicy)
    MyBase.New(lpEnvironment, lpStorageArea, lpStoragePolicy, Nothing)
  End Sub

  Public Sub New(ByVal lpEnvironment As Environment, ByVal lpStorageArea As IStorageArea, ByVal lpStoragePolicy As IStoragePolicy, ByVal lpObjectStoreStatus As ObjectStoreRollingPolicyStatus)
    MyBase.New(lpEnvironment, lpStorageArea, lpStoragePolicy, lpObjectStoreStatus)
  End Sub

#End Region

End Class

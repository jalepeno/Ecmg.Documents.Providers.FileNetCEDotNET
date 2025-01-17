'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  CENetProvider_ISecurityClassification.vb
'   Description :  [type_description_here]
'   Created     :  4/8/2014 10:58:21 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports FileNet.Api.Core
Imports FileNet.Api.Collection
Imports FileNet.Api.Property
Imports FileNet.Api.Admin
Imports Documents.Providers
Imports Documents.Utilities
Imports Documents

#End Region

Partial Public Class CENetProvider
  Implements ISecurityClassification

#Region "Class Variables"

  ' For ISecurityClassification
  Private mobjAvailableRights As Security.IAccessRights

#End Region

#Region "ISecurityClassification Implementation"

  Public ReadOnly Property AvailableRights As Security.IAccessRights Implements ISecurityClassification.AvailableRights
    Get
      Try
        If mobjAvailableRights Is Nothing Then
          mobjAvailableRights = GetAllAvailableRights()
        End If
        Return mobjAvailableRights
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

#End Region

End Class

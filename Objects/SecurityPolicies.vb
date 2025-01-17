' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  SecurityPolicies.vb
'  Description :  Used to manage a cache of P8 Security Policy objects.
'  Created     :  11/7/2011 4:24:28 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Utilities
Imports FileNet.Api.Security

#End Region

Public Class SecurityPolicies
  Inherits CCollection(Of SecurityPolicy)

  Public Overrides Function Contains(ByVal name As String) As Boolean
    Try
      Return Contains(name, Nothing)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Overloads Function Contains(ByVal lpReference As String, ByRef lpFoundPolicy As SecurityPolicy) As Boolean
    Try
      For Each lobjPolicy As SecurityPolicy In Me
        If lobjPolicy.IsMatch(lpReference) Then
          lpFoundPolicy = lobjPolicy
          Return True
        End If
      Next

      ' We did not find a match
      lpFoundPolicy = Nothing
      Return False

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Overloads Sub Add(ByVal lpPolicy As ISecurityPolicy)
    Try
      MyBase.Add(New SecurityPolicy(lpPolicy))
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#Region "Constructors"

  Public Sub New()

  End Sub

#End Region

#Region "Private Methods"

#End Region

End Class

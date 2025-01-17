' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ObjectReference.vb
'  Description :  Used to manage a cache of P8 Security Policy objects.
'  Created     :  11/8/2011 6:45:43 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities
Imports FileNet.Api.Security

#End Region

Public Class SecurityPolicy

#Region "Class Variables"

  Private mobjObjectReference As ObjectReference = Nothing
  Private mobjPolicy As ISecurityPolicy

#End Region

#Region "Public Properties"

  Public Property ObjectReference As ObjectReference
    Get
      Return mobjObjectReference
    End Get
    Set(ByVal value As ObjectReference)
      mobjObjectReference = value
    End Set
  End Property

  Public Property Policy As ISecurityPolicy
    Get
      Return mobjPolicy
    End Get
    Set(ByVal value As ISecurityPolicy)
      mobjPolicy = value
    End Set
  End Property

#End Region

#Region "Constructors"

  Public Sub New()

  End Sub

  Public Sub New(ByVal lpPolicy As ISecurityPolicy)
    Try
      Policy = lpPolicy
      ObjectReference = ParseObjectReferenceFromPolicy(lpPolicy)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

#Region "Public Methods"

  Public Function IsMatch(ByVal lpReference As String) As Boolean
    Try
      If ObjectReference Is Nothing Then
        Return False
      Else
        Return ObjectReference.IsMatch(lpReference)
      End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Overrides Function ToString() As String
    Try
      If ObjectReference IsNot Nothing Then
        Return ObjectReference.ToString
      Else
        Return MyBase.ToString
      End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

#Region "Private Methods"

  Private Function ParseObjectReferenceFromPolicy(ByVal lpPolicy As ISecurityPolicy) As ObjectReference
    Try
      Dim lobjReturnReference As New ObjectReference

      With lobjReturnReference
        .Id = lpPolicy.Id.ToString
        .Name = lpPolicy.Name
      End With

      Return lobjReturnReference

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

End Class

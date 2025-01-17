' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ObjectReference.vb
'  Description :  Used to manage a cache of P8 Security Policy objects.
'  Created     :  11/8/2011 6:24:14 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Utilities

#End Region

Public Class ObjectReference

#Region "Class Variables"

  Private mstrId As String = String.Empty
  Private mobjGuid As Guid = Nothing
  Private mstrName As String = String.Empty

#End Region

#Region "Public Properties"

  Public Property Id As String
    Get
      Return mstrId
    End Get
    Set(ByVal value As String)
      Try
        mstrId = value
        Helper.IsGuid(mstrId, Guid)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  Public Property Guid As Guid
    Get
      Return mobjGuid
    End Get
    Set(ByVal value As Guid)
      mobjGuid = value
    End Set
  End Property

  Public Property Name As String
    Get
      Return mstrName
    End Get
    Set(ByVal value As String)
      mstrName = value
    End Set
  End Property

#End Region

#Region "Constructors"

  Public Sub New()

  End Sub

  Public Sub New(ByVal lpObjectReference As String)
    Try
      ' Try to split up the reference into the component parts
      SplitReference(lpObjectReference)
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
      If lpReference.Contains(Id) OrElse lpReference.Contains(Name) Then
        Return True
      Else
        Return False
      End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Overrides Function ToString() As String
    Try
      Return String.Format("{0}: {1}", Id, Name)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

#Region "Private Methods"

  Private Sub SplitReference(ByVal lpObjectReference As String)
    Try

      If Helper.HasGuid(lpObjectReference, Guid) Then
        Name = lpObjectReference.Replace(Guid.ToString, String.Empty)
      Else
        Name = lpObjectReference
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

End Class

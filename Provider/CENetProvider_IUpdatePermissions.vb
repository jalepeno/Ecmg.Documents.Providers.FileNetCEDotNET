'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  CENetProvider_IUpdatePermissions.vb
'   Description :  [type_description_here]
'   Created     :  4/8/2014 11:00:30 AM
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
Imports FileNet.Api.Security
Imports FileNet.Api.Constants
Imports Documents.Arguments
Imports Documents.Exceptions
Imports Documents.Providers
Imports Documents.Utilities

#End Region

Partial Public Class CENetProvider
  Implements IUpdatePermissions


#Region "IUpdatePermissions Implementation"

  Public Function UpdatePermissions(Args As ObjectSecurityArgs) As Boolean Implements IUpdatePermissions.UpdatePermissions
    Try

      Dim lobjP8Permissions As IAccessPermissionList = CreateP8PermissionList(Args.Permissions)

      If TypeOf Args Is DocumentSecurityArgs Then
        ' The target is a document
        If Me.DocumentExists(Args.ObjectID) Then
          Dim lobjP8Document As IDocument = GetIDocument(Args.ObjectID)
          If UpdatePermissions(lobjP8Document, lobjP8Permissions, Args.Mode) Then
            lobjP8Document.Save(RefreshMode.REFRESH)
            Return True
          Else
            Return False
          End If
        Else
          Throw New DocumentDoesNotExistException(Args.ObjectID)
        End If
      ElseIf TypeOf Args Is FolderSecurityArgs Then
        ' The target is a folder
        If Me.FolderIDExists(Args.ObjectID) Then
          Dim lobjFolder As FileNet.Api.Core.IFolder = ObjectStore.FetchObject("Folder", Args.ObjectID, Nothing) ' GetObject("Folder", Args.ObjectID)
          If UpdatePermissions(lobjFolder, lobjP8Permissions, Args.Mode) Then
            lobjFolder.Save(RefreshMode.REFRESH)
            Return True
          Else
            Return False
          End If
        Else
          Throw New FolderDoesNotExistException(Args.ObjectID)
        End If
      Else
        ' Assume the target is a document
        If Me.DocumentExists(Args.ObjectID) Then
          Dim lobjP8Document As IDocument = GetIDocument(Args.ObjectID)
          If UpdatePermissions(lobjP8Document, lobjP8Permissions, Args.Mode) Then
            lobjP8Document.Save(RefreshMode.REFRESH)
            Return True
          Else
            Return False
          End If
        Else
          Throw New DocumentDoesNotExistException(Args.ObjectID)
        End If
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function UpdatePermissions(lpTarget As IContainable,
                                     lpPermissions As IAccessPermissionList,
                                     lpMode As ObjectSecurityArgs.UpdateMode) As Boolean
    Try
      Select Case lpMode
        Case ObjectSecurityArgs.UpdateMode.Replace
          If lpPermissions IsNot Nothing Then
            lpTarget.Permissions = lpPermissions
          Else
            Return False
          End If
          Return True
        Case ObjectSecurityArgs.UpdateMode.Append
          If lpPermissions IsNot Nothing AndAlso lpPermissions.Count > 0 Then
            For Each lobjP8Permission As IAccessPermission In lpPermissions
              lpTarget.Permissions.Add(lobjP8Permission)
            Next
          Else
            Return False
          End If
          Return True
        Case ObjectSecurityArgs.UpdateMode.Update
          If lpPermissions IsNot Nothing AndAlso lpPermissions.Count > 0 Then
            Dim lblnExistingPermissionFound As Boolean = False
            Dim lobjExistingP8Permission As IAccessPermission = Nothing
            For Each lobjCandidateP8Permission As IAccessPermission In lpPermissions
              lblnExistingPermissionFound = False
              For lintExistingPermissionCounter As Integer = 0 To lpTarget.Permissions.Count - 1
                lobjExistingP8Permission = lpTarget.Permissions(lintExistingPermissionCounter)
                If String.Compare(lobjExistingP8Permission.GranteeName, lobjCandidateP8Permission.GranteeName, True) = 0 Then
                  lblnExistingPermissionFound = True
                  lobjExistingP8Permission.AccessMask = lobjCandidateP8Permission.AccessMask
                  Continue For
                End If
              Next
              If lblnExistingPermissionFound = False Then
                lpTarget.Permissions.Add(lobjCandidateP8Permission)
              End If
            Next
          Else
            Return False
          End If
          Return True
      End Select
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

End Class

'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  UserDoesNotHaveElevatedPriviledgesException.vb
'   Description :  [type_description_here]
'   Created     :  4/8/2014 9:49:26 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"



#End Region

Public Class UserDoesNotHaveElevatedPriviledgesException
  Inherits Exception

#Region "Class Constants"

  Private Const MSG_TEMPLATE As String =
    "User '{0}' does not have elevated priviledges, in order to preserve last modified info grant 'Modify certain system properties'."

#End Region

#Region "Class Variables"

  Private mstrUserName As String

#End Region

#Region "Public Properties"

  Public ReadOnly Property UserName As String
    Get
      Return mstrUserName
    End Get
  End Property

#End Region

#Region "Constructors"

  Public Sub New(ByVal userName As String)
    MyBase.New(String.Format(MSG_TEMPLATE, userName))
    mstrUserName = userName
  End Sub

  Public Sub New(ByVal message As String, ByVal userName As String)
    MyBase.New(message)
    mstrUserName = userName
  End Sub

  Public Sub New(ByVal userName As String, ByVal innerException As Exception)
    MyBase.New(String.Format(MSG_TEMPLATE, userName), innerException)
    mstrUserName = userName
  End Sub

  Public Sub New(ByVal message As String, ByVal userName As String, ByVal innerException As Exception)
    MyBase.New(message, innerException)
    mstrUserName = userName
  End Sub

#End Region

End Class

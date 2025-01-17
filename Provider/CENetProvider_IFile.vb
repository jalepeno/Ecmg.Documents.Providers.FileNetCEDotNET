'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  CENetProvider_IFile.vb
'   Description :  [type_description_here]
'   Created     :  4/8/2014 11:12:18 AM
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
Imports FileNet.Api.Constants
Imports System.IO
Imports Documents.Providers
Imports Documents.Utilities

#End Region

Partial Public Class CENetProvider
  Implements IFile

#Region "IFile Implementastion"

  Public Function FileDocument(ByVal lpId As String,
                               ByVal lpFolderPath As String,
                               ByVal lpSessionId As String) As Boolean _
         Implements IFile.FileDocument

    Try

      Return FileDocument(lpId, lpFolderPath)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

  Public Function UnFileDocument(ByVal lpId As String,
                                 ByVal lpFolderPath As String,
                                 ByVal lpSessionId As String) As Boolean _
         Implements IFile.UnFileDocument

    Try
      Return UnFileDocument(lpId, lpFolderPath)

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

#End Region

End Class

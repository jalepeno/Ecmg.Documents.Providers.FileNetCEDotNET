'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  CENetProvider_IFolderClassification.vb
'   Description :  [type_description_here]
'   Created     :  4/8/2014 10:55:55 AM
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
Imports Documents.Core
Imports Documents.Exceptions
Imports Documents.Providers
Imports Documents.Utilities
Imports Documents

#End Region

Partial Public Class CENetProvider
  Implements IFolderClassification

#Region "Class Variables"

  ' For IFolderClassification
  Private mobjFolderClasses As FolderClasses
  Private mobjRequestedFolderClasses As FolderClasses
  Private mobjFolderProperties As ClassificationProperties

#End Region

#Region "IFolderClassification Implementation"

  Public ReadOnly Property FolderClass(lpFolderClassName As String) As Core.FolderClass Implements IFolderClassification.FolderClass
    Get
      Try

        If mobjFolderClasses Is Nothing Then
          mobjFolderClasses = GetFolderClasses()
        End If

        If mobjFolderClasses Is Nothing OrElse mobjFolderClasses.Count = 0 Then
          Throw New FolderClassNotInitializedException("No folder classes are initialized.", lpFolderClassName, Me.ContentSource)
        End If

        If mobjFolderClasses.Contains(lpFolderClassName) Then
          Return mobjFolderClasses(lpFolderClassName)

        Else

          If Me.ContentSource IsNot Nothing Then
            Throw New FolderClassNotInitializedException(String.Format("The FolderClass '{0}' was not found in the folder class collection for {1}.", lpFolderClassName, Me.ContentSource.Name), lpFolderClassName, Me.ContentSource)

          Else
            Throw New FolderClassNotInitializedException(String.Format("The FolderClass '{0}' was not found in the document class collection.", lpFolderClassName), lpFolderClassName, Nothing)
          End If

        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod, 62714, lpFolderClassName)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

  Public ReadOnly Property FolderClasses As Core.FolderClasses Implements IFolderClassification.FolderClasses
    Get

      Try

        If mobjFolderClasses Is Nothing Then
          mobjFolderClasses = GetFolderClasses()
        End If

        Return mobjFolderClasses

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Get
  End Property

  Public ReadOnly Property FolderProperties As Core.ClassificationProperties Implements IFolderClassification.FolderProperties
    Get

      Try

        If mobjProperties Is Nothing Then
          mobjProperties = GetAllFolderProperties()
        End If

        Return mobjProperties

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Get
  End Property

#End Region

End Class

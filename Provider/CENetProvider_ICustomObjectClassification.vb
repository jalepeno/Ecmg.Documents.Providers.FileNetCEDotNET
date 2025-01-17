'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  CENetProvider_ICustomObjectClassification.vb
'   Description :  [type_description_here]
'   Created     :  9/2/2015 1:46:55 AM
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

#End Region

Partial Public Class CENetProvider
  Implements ICustomObjectClassification

#Region "Class Variables"

  ' For IFolderClassification
  Private mobjObjectClasses As ObjectClasses
  Private mobjRequestedObjectClasses As ObjectClasses

#End Region

#Region "ICustomObjectClassification Implementation"

  Public ReadOnly Property ObjectProperties As ClassificationProperties Implements ICustomObjectClassification.ObjectProperties
    Get

      Try

        If mobjProperties Is Nothing Then
          mobjProperties = GetAllCustomObjectProperties()
        End If

        Return mobjProperties

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Get
  End Property

  Public ReadOnly Property ObjectClasses As ObjectClasses Implements ICustomObjectClassification.ObjectClasses
    Get

      Try

        If mobjObjectClasses Is Nothing Then
          mobjObjectClasses = GetCustomObjectClasses()
        End If

        Return mobjObjectClasses

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Get
  End Property

  Public ReadOnly Property ObjectClass(lpObjectClassName As String) As ObjectClass Implements ICustomObjectClassification.ObjectClass
    Get
      Try

        If mobjObjectClasses Is Nothing Then
          mobjObjectClasses = GetCustomObjectClasses()
        End If

        If mobjObjectClasses Is Nothing OrElse mobjObjectClasses.Count = 0 Then
          Throw New FolderClassNotInitializedException("No custom object classes are initialized.", lpObjectClassName, Me.ContentSource)
        End If

        If mobjObjectClasses.Contains(lpObjectClassName) Then
          Return mobjObjectClasses(lpObjectClassName)

        Else

          If Me.ContentSource IsNot Nothing Then
            Throw New FolderClassNotInitializedException(String.Format("The ObjectClass '{0}' was not found in the custom object class collection for {1}.", lpObjectClassName, Me.ContentSource.Name), lpObjectClassName, Me.ContentSource)

          Else
            Throw New FolderClassNotInitializedException(String.Format("The ObjectClass '{0}' was not found in the custom object class collection.", lpObjectClassName), lpObjectClassName, Nothing)
          End If

        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod, 62814, lpObjectClassName)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

#End Region

End Class

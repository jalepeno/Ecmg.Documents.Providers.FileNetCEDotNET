'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  CENetProvider_IClassification.vb
'   Description :  [type_description_here]
'   Created     :  4/8/2014 10:49:52 AM
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
Imports Documents.Providers
Imports Documents.Utilities
Imports DCore = Documents.Core

#End Region

Partial Public Class CENetProvider
  Implements IClassification

#Region "Class Variables"

  Private mobjDocumentClasses As DocumentClasses
  Private mobjRequestedDocumentClasses As DocumentClasses
  Private mobjProperties As ClassificationProperties
  Private mobjRequestedProperties As ClassificationProperties

#End Region

#Region "IClassification Implementation"

  Public ReadOnly Property ContentProperty(lpName As String) As ClassificationProperty
    Get

      Dim lobjContentProperty As ClassificationProperty = Nothing

      Try
        If mobjRequestedProperties Is Nothing Then
          mobjRequestedProperties = New ClassificationProperties
        End If

        lobjContentProperty = mobjRequestedProperties.ItemByName(lpName)

        If lobjContentProperty Is Nothing Then
          lobjContentProperty = GetClassificationProperty(lpName)
          mobjRequestedProperties.Add(lobjContentProperty)
        End If

        Return lobjContentProperty

        'If (mobjRequestedProperties Is Nothing OrElse mobjRequestedProperties.ItemByName(lpName) Is Nothing) Then
        '	Dim lobjProperty As Core.ClassificationProperty = GetClassificationProperty(lpName)
        '	If (mobjRequestedProperties Is Nothing) Then
        '		mobjRequestedProperties = New ClassificationProperties
        '	End If
        '	If (lobjProperty IsNot Nothing) Then
        '		mobjRequestedProperties.Add(lobjProperty)
        '	End If
        '	Return lobjProperty
        'Else
        '	Return mobjRequestedProperties(lpName)
        'End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Get
  End Property

  Public ReadOnly Property ContentProperties() As ClassificationProperties _
                  Implements IClassification.ContentProperties
    Get

      Try

        If mobjProperties Is Nothing Then
          mobjProperties = GetAllContentProperties()
        End If

        Return mobjProperties

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Get
  End Property

  Public ReadOnly Property DocumentClasses() As DocumentClasses _
                  Implements IClassification.DocumentClasses
    Get

      Try

        If mobjDocumentClasses Is Nothing Then
          mobjDocumentClasses = GetDocumentClasses()
        End If

        Return mobjDocumentClasses

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Get
  End Property

  Public ReadOnly Property DocumentClass(ByVal lpDocumentClassName As String) As DCore.DocumentClass _
                  Implements IClassification.DocumentClass
    Get


      Try
        If (mobjRequestedDocumentClasses Is Nothing OrElse mobjRequestedDocumentClasses(lpDocumentClassName) Is Nothing) Then
          Dim lobjDocumentClass As DCore.DocumentClass = GetDocumentClass(lpDocumentClassName)
          If (mobjRequestedDocumentClasses Is Nothing) Then
            mobjRequestedDocumentClasses = New DCore.DocumentClasses
          End If
          If (lobjDocumentClass IsNot Nothing) Then
            mobjRequestedDocumentClasses.Add(lobjDocumentClass)
          End If
          Return lobjDocumentClass
        Else
          Return mobjRequestedDocumentClasses(lpDocumentClassName)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try


      'Try

      '  If mobjDocumentClasses Is Nothing Then
      '    mobjDocumentClasses = GetDocumentClasses()
      '  End If

      '  If mobjDocumentClasses Is Nothing OrElse mobjDocumentClasses.Count = 0 Then
      '    Throw New Exceptions.DocumentClassNotInitializedException("No document classes are initialized.", lpDocumentClassName, Me.ContentSource)
      '  End If

      '  If mobjDocumentClasses.Contains(lpDocumentClassName) Then
      '    Return mobjDocumentClasses(lpDocumentClassName)

      '  Else

      '    If Me.ContentSource IsNot Nothing Then
      '      Throw New DocumentClassNotInitializedException(String.Format("The DocumentClass '{0}' was not found in the document class collection for {1}.", lpDocumentClassName, Me.ContentSource.Name), lpDocumentClassName, Me.ContentSource)

      '    Else
      '      Throw New Exceptions.DocumentClassNotInitializedException(String.Format("The DocumentClass '{0}' was not found in the document class collection.", lpDocumentClassName), lpDocumentClassName, Nothing)
      '    End If

      '  End If

      'Catch ex As Exception
      '  ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod, 62713, lpDocumentClassName)
      '  ' Re-throw the exception to the caller
      '  Throw
      'End Try

    End Get
  End Property

#Region "IClassification Worker Methods"

  Private Function GetAllContentProperties() As ClassificationProperties

    Dim lobjProperties As New ClassificationProperties
    Dim lobjProperty As ClassificationProperty

    Try

      If IsInitialized Then

        lobjProperties = GetAllBaseProperties()

        ' Also get any additional property definitions that may be in the root 'Document' class
        For Each lobjRootClass As IClassDefinition In ObjectStore.RootClassDefinitions

          If lobjRootClass.Name = "Document" Then

            For Each lobjPropertyDefinition As IPropertyDefinition In lobjRootClass.PropertyDefinitions
              lobjProperty = GetClassificationProperty(lobjPropertyDefinition)

              If lobjProperties.Contains(lobjProperty.Name) = False Then
                Debug.WriteLine(lobjProperty.ID & ": " & lobjProperty.Name)
                lobjProperties.Add(lobjProperty) ', lobjProperty.ID)
              End If

            Next

            Exit For
          End If

        Next

      End If

      lobjProperties.Sort()

      Return lobjProperties

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

#End Region

#End Region

End Class

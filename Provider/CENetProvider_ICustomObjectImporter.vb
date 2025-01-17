'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  CENetProvider_ICustomObjectImporter.vb
'   Description :  [type_description_here]
'   Created     :  9/2/2015 12:57:23 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports FileNet.Api.Constants
Imports FileNet.Api.Core
Imports Documents.Arguments
Imports Documents.Core
Imports Documents.Providers
Imports Documents.Utilities

#End Region

Partial Public Class CENetProvider
  Implements ICustomObjectImporter

#Region "Public Methods"

  Public Function ImportObject(lpArgs As ImportObjectArgs) As Boolean Implements ICustomObjectImporter.ImportObject
    Try

      If lpArgs Is Nothing Then Throw New ArgumentNullException("lpArgs")

      Dim lobjSourceObject As CustomObject = lpArgs.Object
      Dim lobjP8CustomObject As FileNet.Api.Core.ICustomObject = Nothing

      If lobjSourceObject Is Nothing Then Throw New InvalidOperationException("Object reference not set")

      Dim lstrLastModifer As String = Me.UserName
      Dim ldatDateLastModified As Date = Now
      Dim ldatDateCheckedIn As Date = Now

      If HasElevatedPrivileges Then

        ' If we have last modifier and date last modified in the supplied document let's try to use it
        If lobjSourceObject.Properties.PropertyExists(PROP_MODIFY_USER) Then

          Dim lobjLastModifierProperty As ECMProperty = lobjSourceObject.Properties(PROP_MODIFY_USER)

          If lobjLastModifierProperty.HasValue Then
            lstrLastModifer = lobjLastModifierProperty.Value
          End If

        End If

        If lobjSourceObject.Properties.PropertyExists(PROP_MODIFY_DATE) Then

          Dim lobjLastModifiedProperty As ECMProperty = lobjSourceObject.Properties(PROP_MODIFY_DATE)

          If lobjLastModifiedProperty.HasValue Then
            ldatDateLastModified = lobjLastModifiedProperty.Value
          End If

        End If

      End If

      ' Get the candidate class
      Dim lobjCandidateClass As ObjectClass = ObjectClass(lobjSourceObject.ClassName)

      lobjP8CustomObject = Factory.CustomObject.CreateInstance(mobjObjectStore, lobjCandidateClass.Name)

      ' Set the properties needed on create
      For Each lobjCDFVersionProperty As ECMProperty In lobjSourceObject.Properties

        If lobjCDFVersionProperty.HasValue Then
          SetPropertyValue(lobjP8CustomObject, lobjCandidateClass, lobjCDFVersionProperty, True)
        End If

      Next

      ' Set the permissions if we have them and we were asked to.
      'If lpArgs.SetPermissions AndAlso lobjP8Folder.Permissions IsNot Nothing AndAlso lobjP8Folder.Permissions.Count > 0 Then
      '  lobjP8Folder.Permissions = CreateP8PermissionList(lobjP8Folder.Permissions)
      'End If

      If HasElevatedPrivileges Then
        lobjP8CustomObject.Properties.RemoveFromCache(PropertyNames.LAST_MODIFIER)
        lobjP8CustomObject.Properties.RemoveFromCache(PropertyNames.DATE_LAST_MODIFIED)
        lobjP8CustomObject.LastModifier = lstrLastModifer
        lobjP8CustomObject.DateLastModified = ldatDateLastModified
      End If

      lobjP8CustomObject.Save(RefreshMode.REFRESH)

      lpArgs.Object.Id = lobjP8CustomObject.Id.ToString()

      Return True

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

End Class

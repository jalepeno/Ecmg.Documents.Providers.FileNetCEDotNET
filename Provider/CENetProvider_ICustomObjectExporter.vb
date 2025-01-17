'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  CENetProvider_ICustomObjectExporter.vb
'   Description :  [type_description_here]
'   Created     :  9/1/2015 3:39:00 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Arguments
Imports Documents.Exceptions
Imports Documents.Providers
Imports Documents.Utilities
Imports DCore = Documents.Core

#End Region

Partial Public Class CENetProvider
  Implements ICustomObjectExporter

#Region "Events"

  Public Event ObjectExportError As ObjectExportErrorEventHandler Implements ICustomObjectExporter.ObjectExportError

#End Region

#Region "ICustomObjectExporter Implementation"

  Public Function ExportObject(Args As ExportObjectEventArgs) As Boolean Implements ICustomObjectExporter.ExportObject
    Try
      Dim lobjCustomObject As FileNet.Api.Core.ICustomObject = Nothing
      Dim lobjReturnObject As New DCore.CustomObject
      Dim lobjECMProperty As DCore.ECMProperty

      lobjCustomObject = GetObject("CustomObject", Args.ObjectId)
      lobjCustomObject.Refresh()

      If lobjCustomObject Is Nothing Then
        Throw New ItemDoesNotExistException(Args.ObjectId)
      End If

      lobjReturnObject.Id = lobjCustomObject.Id.ToString()
      lobjReturnObject.ClassName = lobjCustomObject.ClassDescription.SymbolicName

      For Each lobjIProperty As FileNet.Api.Property.IProperty In lobjCustomObject.Properties
        lobjECMProperty = CreateECMProperty(lobjIProperty)

        If (lobjECMProperty IsNot Nothing) Then
          If lobjReturnObject.Properties.PropertyExists(lobjECMProperty.Name) = False Then
            lobjReturnObject.Properties.Add(lobjECMProperty)
          End If
          If lobjECMProperty.Name.Equals("Name") Then
            lobjReturnObject.Name = lobjECMProperty.Value
          End If
        End If
      Next

      lobjReturnObject.Properties.Sort()

      If Args.GetPermissions Then
        ' Get all the permissions for this object.
        lobjReturnObject.Permissions.AddRange(GetCtsPermissions(lobjCustomObject))
      End If

      Args.Object = lobjReturnObject

      Return True

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
      RaiseEvent ObjectExportError(Me, New ObjectExportErrorEventArgs(Args, ex))
    End Try
  End Function

#End Region

End Class

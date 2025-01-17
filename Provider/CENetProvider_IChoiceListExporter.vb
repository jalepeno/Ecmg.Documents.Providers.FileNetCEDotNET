'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  CENetProvider_IChoiceListExporter.vb
'   Description :  [type_description_here]
'   Created     :  4/8/2014 11:25:13 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports FileNet.Api.Constants
Imports FileNet.Api
Imports FileNet.Api.Admin
Imports FileNet.Api.Util
Imports System
Imports Documents.Arguments
Imports Documents.Core.ChoiceLists
Imports Documents.Utilities
Imports DCore = Documents.Core

#End Region

Partial Public Class CENetProvider
  Implements IChoiceListExporter

#Region "IChoiceListExporter Implementation"

  Public Event ChoiceListExported(sender As Object, e As ChoiceListExportedEventArgs) Implements IChoiceListExporter.ChoiceListExported

  Public Event ChoiceListExportError(sender As Object, e As ChoiceListExportErrorEventArgs) Implements IChoiceListExporter.ChoiceListExportError

  Public ReadOnly Property ChoiceListNames As List(Of String) Implements IChoiceListExporter.ChoiceListNames
    Get
      Try
        Return GetChoiceListNames()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
  End Property

  Public Function ExportChoiceList(ByRef lpArgs As ExportChoiceListEventArgs) As Boolean Implements IChoiceListExporter.ExportChoiceList
    Try
      If lpArgs Is Nothing Then
        Throw New ArgumentNullException("lpArgs")
      End If
      If String.IsNullOrEmpty(lpArgs.ID) Then
        Throw New ArgumentOutOfRangeException("lpArgs.ID", "The ID property must contain a value.")
      End If

      Dim lobjP8ChoiceList As Admin.IChoiceList = GetP8ChoiceList(lpArgs.ID)
      Dim lobjChoiceList As New ChoiceList
      Select Case lobjP8ChoiceList.DataType
        Case TypeID.STRING
          lobjChoiceList.Type = DCore.ChoiceLists.ChoiceType.ChoiceString
        Case TypeID.LONG
          lobjChoiceList.Type = DCore.ChoiceLists.ChoiceType.ChoiceInteger
        Case Else
          Beep()
      End Select

      ' TODO: Handle choice groups
      With lobjChoiceList
        .Id = lobjP8ChoiceList.Id.ToString
        .Name = lobjP8ChoiceList.Name
        .DisplayName = lobjP8ChoiceList.DisplayName
        .DescriptiveText = lobjP8ChoiceList.DescriptiveText
        For Each lobjChoiceItem As Admin.IChoice In lobjP8ChoiceList.ChoiceValues
          .ChoiceValues.Add(GetChoiceValue(lobjChoiceItem))
        Next
      End With

      lpArgs.SetChoiceList(lobjChoiceList)

      OnChoiceListExported(New ChoiceListExportedEventArgs(lpArgs))

      Return True

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      lpArgs.SetErrorMessage(ex.Message)
      OnChoiceListExportError(New ChoiceListExportErrorEventArgs(lpArgs, ex))
      Return False
    End Try
  End Function

  Public Sub OnChoiceListExported(ByRef e As ChoiceListExportedEventArgs) Implements IChoiceListExporter.OnChoiceListExported
    Try
      RaiseEvent ChoiceListExported(Me, e)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub OnChoiceListExportError(ByRef e As ChoiceListExportErrorEventArgs) Implements IChoiceListExporter.OnChoiceListExportError
    Try
      RaiseEvent ChoiceListExportError(Me, e)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Private Function GetChoiceListNames() As List(Of String)
    Try
      Dim lobjChoiceListNames As New List(Of String)

      If IsInitialized AndAlso ObjectStore.ChoiceLists IsNot Nothing Then
        For Each lobjChoiceList As Admin.IChoiceList In ObjectStore.ChoiceLists
          lobjChoiceListNames.Add(lobjChoiceList.Name)
        Next
      End If

      ' Sort the list
      lobjChoiceListNames.Sort()

      Return lobjChoiceListNames

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function GetChoiceValue(lpP8ChoiceItem As Admin.IChoice) As ChoiceValue
    Try
      Dim lobjChoiceValue As ChoiceValue
      lobjChoiceValue = New ChoiceValue
      With lobjChoiceValue
        .Id = lpP8ChoiceItem.Id.ToString
        .Name = lpP8ChoiceItem.Name
        .DisplayName = lpP8ChoiceItem.DisplayName
      End With

      Select Case lpP8ChoiceItem.ChoiceType
        Case Constants.ChoiceType.STRING
          lobjChoiceValue.ChoiceType = DCore.ChoiceLists.ChoiceType.ChoiceString
          lobjChoiceValue.Value = lpP8ChoiceItem.ChoiceStringValue

        Case Constants.ChoiceType.INTEGER
          lobjChoiceValue.ChoiceType = DCore.ChoiceLists.ChoiceType.ChoiceInteger
          lobjChoiceValue.Value = lpP8ChoiceItem.ChoiceIntegerValue

        Case Constants.ChoiceType.MIDNODE_STRING
          Beep()

        Case Constants.ChoiceType.MIDNODE_INTEGER
          Beep()

      End Select

      Return lobjChoiceValue

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

End Class

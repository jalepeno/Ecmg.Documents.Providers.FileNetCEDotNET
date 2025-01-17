'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  CENetProvider_IChoiceListImporter.vb
'   Description :  [type_description_here]
'   Created     :  4/8/2014 11:36:20 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core.ChoiceLists.Exceptions
Imports FileNet.Api.Core
Imports FileNet.Api.Constants
Imports FileNet.Api
Imports FileNet.Api.Admin
Imports FileNet.Api.Util
Imports System
Imports Documents.Arguments
Imports Documents.Core
Imports Documents.Core.ChoiceLists
Imports Documents.Utilities
Imports DCore = Documents.Core

#End Region

Partial Public Class CENetProvider
  Implements IChoiceListImporter

#Region "IChoiceListImporter Implementation"


  Public Function ChoiceListExists(ByRef lpId As ObjectIdentifier, Optional ByRef lpReturnedObjectId As String = "") As Boolean Implements IChoiceListImporter.ChoiceListExists
    Try
      Dim lobjP8ChoiceList As Admin.IChoiceList = GetP8ChoiceList(lpId.IdentifierValue)
      If lobjP8ChoiceList IsNot Nothing Then
        lpReturnedObjectId = lobjP8ChoiceList.Id.ToString
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

  Public Event ChoiceListImported(sender As Object, e As ChoiceListImportedEventArgs) Implements IChoiceListImporter.ChoiceListImported

  Public Event ChoiceListImportError(sender As Object, e As ChoiceListImportErrorEventArgs) Implements IChoiceListImporter.ChoiceListImportError

  Public Function ImportChoiceList(ByRef Args As ImportChoiceListEventArgs) As Boolean Implements IChoiceListImporter.ImportChoiceList
    Try
      Dim lobjChoiceListIdentifier As DCore.ObjectIdentifier
      Dim lblnChoiceListAlreadyExists As Boolean
      Dim lstrExistingChoiceListId As String = String.Empty
      Dim lblnSuccess As Boolean

      lobjChoiceListIdentifier = New DCore.ObjectIdentifier(Args.ChoiceList.Name, DCore.ObjectIdentifier.IdTypeEnum.Name)
      lblnChoiceListAlreadyExists = ChoiceListExists(lobjChoiceListIdentifier, lstrExistingChoiceListId)

      ' Check to see if we need to add it or replace it.
      If lblnChoiceListAlreadyExists = True Then
        Dim lobjChoiceList As ChoiceList = GetChoiceList(lstrExistingChoiceListId)
        If Args.ReplaceExisting = True Then
          Dim lenuReplaceResult As DCore.ReplaceResult = ReplaceChoiceList(lobjChoiceListIdentifier, Args.ChoiceList, Args.ErrorMessage)
          Select Case lenuReplaceResult
            Case DCore.ReplaceResult.ObjectReplacedPreservingOriginalID
              RaiseEvent ChoiceListImported(Me, New ChoiceListImportedEventArgs(Args, lobjChoiceListIdentifier.IdentifierValue))
              Return True
            Case Else
              RaiseEvent ChoiceListImportError(Me,
                New ChoiceListImportErrorEventArgs(String.Format("Replace choice list operation failed: {0}",
                                                                 lenuReplaceResult.ToString), Nothing))
              Return False
          End Select
        Else
          ' Throw an exception indicating the choicelist already exists
          Throw New ChoiceListAlreadyExistsException(lobjChoiceList, Me.ContentSource.Name)
        End If
      Else
        lblnSuccess = AddChoiceList(Args.ChoiceList, Args.ErrorMessage)
        If lblnSuccess = True Then
          RaiseEvent ChoiceListImported(Me, New ChoiceListImportedEventArgs(Args, Args.ChoiceList.Id))
        Else
          RaiseEvent ChoiceListImportError(Me, New ChoiceListImportErrorEventArgs(Args, Nothing))
        End If

        Return lblnSuccess
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Args.ErrorMessage = ex.Message
      RaiseEvent ChoiceListImportError(Me, New ChoiceListImportErrorEventArgs(Args, ex))
      Return False
    End Try
  End Function

  Public Sub OnChoiceListImported(ByRef e As ChoiceListImportedEventArgs) Implements IChoiceListImporter.OnChoiceListImported
    Try
      RaiseEvent ChoiceListImported(Me, e)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Public Sub OnChoiceListImportError(ByRef e As ChoiceListImportErrorEventArgs) Implements IChoiceListImporter.OnChoiceListImportError
    Try
      RaiseEvent ChoiceListImportError(Me, e)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

  Private Function AddChoiceList(ByRef lpChoiceList As ChoiceList,
    Optional ByRef lpErrorMessage As String = "",
    Optional ByVal lpIncludeObjectID As Boolean = False,
    Optional ByRef lpNewObjectID As String = "") As Boolean

    Try
      Dim lobjP8ChoiceList As Admin.IChoiceList = CreateP8ChoiceList(lpChoiceList)

      lobjP8ChoiceList.Save(RefreshMode.REFRESH)
      lpNewObjectID = lobjP8ChoiceList.Id.ToString

      Return True

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      lpErrorMessage &= ex.Message.Replace("'", "^")
      Return False
    End Try
  End Function

  Private Function ReplaceChoiceList(ByVal lpId As DCore.ObjectIdentifier,
  ByVal lpNewChoiceList As ChoiceList,
  Optional ByRef lpErrorMessage As String = "") As DCore.ReplaceResult
    Try
      Dim lobjP8ChoiceList As Admin.IChoiceList = GetP8ChoiceList(lpId.IdentifierValue)
      If lobjP8ChoiceList IsNot Nothing Then
        'lobjP8ChoiceList.Delete()
        'lobjP8ChoiceList.Save(RefreshMode.NO_REFRESH)
        lobjP8ChoiceList.ChoiceValues.Clear()

        For Each lobjChoiceValue As ChoiceValue In lpNewChoiceList.ChoiceValues
          lobjP8ChoiceList.ChoiceValues.Add(CreateP8ChoiceValue(lobjChoiceValue))
        Next
      End If

      lobjP8ChoiceList.Save(RefreshMode.NO_REFRESH)

      Return DCore.ReplaceResult.ObjectReplacedPreservingOriginalID

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      lpErrorMessage = Helper.FormatCallStack(ex)
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Return DCore.ReplaceResult.ObjectNotFoundNoAction
    End Try
  End Function

  Private Function CreateP8ChoiceList(lpChoiceList As ChoiceList) As Admin.IChoiceList
    Try
      Dim lobjP8ChoiceList As Admin.IChoiceList = Factory.ChoiceList.CreateInstance(ObjectStore)
      lobjP8ChoiceList.DisplayName = lpChoiceList.Name
      lobjP8ChoiceList.DescriptiveText = lpChoiceList.DescriptiveText

      Select Case lpChoiceList.Type
        Case DCore.ChoiceLists.ChoiceType.ChoiceString
          lobjP8ChoiceList.DataType = Constants.TypeID.STRING
        Case DCore.ChoiceLists.ChoiceType.ChoiceInteger
          lobjP8ChoiceList.DataType = Constants.TypeID.LONG
      End Select

      lobjP8ChoiceList.ChoiceValues = Factory.Choice.CreateList

      For Each lobjChoiceValue As ChoiceValue In lpChoiceList.ChoiceValues
        lobjP8ChoiceList.ChoiceValues.Add(CreateP8ChoiceValue(lobjChoiceValue))
      Next

      Return lobjP8ChoiceList

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Private Function CreateP8ChoiceValue(lpChoiceValue As ChoiceValue) As IChoice
    ' http://pic.dhe.ibm.com/infocenter/p8docs/v5r0m0/index.jsp?topic=%2Fcom.ibm.p8.ce.dev.ce.doc%2Fchoicelist_procedures.htm
    Try
      Dim lobjP8Choice As IChoice = Factory.Choice.CreateInstance

      With lobjP8Choice
        .DisplayName = lpChoiceValue.DisplayName
      End With

      Select Case lpChoiceValue.ChoiceType
        Case DCore.ChoiceLists.ChoiceType.ChoiceString
          lobjP8Choice.ChoiceType = Constants.ChoiceType.STRING
          lobjP8Choice.ChoiceStringValue = lpChoiceValue.Value.ToString

        Case DCore.ChoiceLists.ChoiceType.ChoiceInteger
          lobjP8Choice.ChoiceType = Constants.ChoiceType.INTEGER
          lobjP8Choice.ChoiceIntegerValue = lpChoiceValue.Value

      End Select

      Return lobjP8Choice

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

End Class

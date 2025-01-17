'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  CENetProvider_Extras.vb
'   Description :  [type_description_here]
'   Created     :  4/8/2014 8:30:08 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities
Imports Ecmg.Cts
Imports Ecmg.Cts.Utilities
Imports FileNet.Api
Imports FileNet.Api.Admin
Imports FileNet.Api.Constants
Imports FileNet.Api.Core
Imports System

#End Region

Partial Public Class CENetProvider

  Private Sub SaveObject(lpP8Object As IIndependentlyPersistableObject, lpPreserveLastModifiedInfo As Boolean)
    Try

      Dim lstrLastModifer As String = Me.UserName
      Dim ldatDateLastModified As Date = Now

      If lpPreserveLastModifiedInfo Then

        ' In order to back fill the modified properties we require special priviledges on the object store.
        If HasElevatedPrivileges Then

          lstrLastModifer = lpP8Object.Properties(PropertyNames.LAST_MODIFIER)
          ldatDateLastModified = lpP8Object.Properties(PropertyNames.DATE_LAST_MODIFIED)

          ' We will save the object once to persist any updates prior to this method call.
          lpP8Object.Save(RefreshMode.REFRESH)

          ' We are removing the last modifier and last modified date properties from the local property cache
          lpP8Object.Properties.RemoveFromCache(PropertyNames.LAST_MODIFIER)
          lpP8Object.Properties.RemoveFromCache(PropertyNames.DATE_LAST_MODIFIED)
          ' We are replacing those properties with the values we retrieved before our initial save.
          lpP8Object.Properties(PropertyNames.LAST_MODIFIER) = lstrLastModifer
          lpP8Object.Properties(PropertyNames.DATE_LAST_MODIFIED) = ldatDateLastModified

          ' We are saving again to persist the original modified properties we just changed above.
          lpP8Object.Save(RefreshMode.REFRESH)

        Else
          Throw New UserDoesNotHaveElevatedPriviledgesException(Me.UserName)
        End If
      Else
        ' Just save the document, we are not trying to preserve the last modified information.
        lpP8Object.Save(RefreshMode.REFRESH)
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

End Class

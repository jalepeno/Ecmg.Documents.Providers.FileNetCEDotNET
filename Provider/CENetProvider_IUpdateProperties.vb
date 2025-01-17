'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  CENetProvider_IUpdateProperties.vb
'   Description :  [type_description_here]
'   Created     :  4/8/2014 11:02:53 AM
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
Imports FileNet.Api.Security
Imports FileNet.Api.Constants
Imports Documents
Imports Documents.Core
Imports Documents.Providers
Imports Documents.Utilities

#End Region

Partial Public Class CENetProvider
  Implements IUpdateProperties

#Region "IUpdateProperties Implementation"

  Public Function UpdateDocumentProperties(ByVal Args As Arguments.DocumentPropertyArgs) As Boolean _
         Implements IUpdateProperties.UpdateDocumentProperties, IBasicContentServicesProvider.UpdateDocumentProperties

    Try

      If Args Is Nothing Then
        Throw New ArgumentNullException("Args")
      End If

      ' Make sure we have a valid doc id
      If String.IsNullOrEmpty(Args.DocumentID) Then
        Throw New ArgumentNullException("Args.DocumentID")
      End If

      ' Get document and populate property cache.
      Dim lobjIncludePropertyFilter As PropertyFilter = New PropertyFilter()

      'For Each lobjProperty As Core.IProperty In Args.Properties
      '  lobjIncludePropertyFilter.AddIncludeProperty(New FilterElement(Nothing, Nothing, Nothing, lobjProperty.SystemName, Nothing))
      'Next

      ' Dim lobjIDocument As IDocument = GetIDocument(Args.DocumentID, lobjIncludePropertyFilter)
      Dim lobjIDocument As IDocument = GetIDocument(Args.DocumentID)

      Dim lobjTargetDocumentClass As DocumentClass = DocumentClass(lobjIDocument.GetClassName())

      If String.IsNullOrEmpty(Args.VersionID) Then
        Select Case Args.VersionScope
          Case VersionScopeEnum.MostCurrentVersion
            lobjIDocument = lobjIDocument.VersionSeries.CurrentVersion
            For Each lobjProperty As ECMProperty In Args.Properties
              SetPropertyValue(lobjIDocument, lobjTargetDocumentClass, lobjProperty, False)
            Next
            ' Save and update property cache.
            lobjIDocument.Save(RefreshMode.REFRESH)
          Case VersionScopeEnum.CurrentReleasedVersion
            lobjIDocument = lobjIDocument.VersionSeries.ReleasedVersion
            For Each lobjProperty As ECMProperty In Args.Properties
              SetPropertyValue(lobjIDocument, lobjTargetDocumentClass, lobjProperty, False)
            Next
            ' Save and update property cache.
            lobjIDocument.Save(RefreshMode.REFRESH)
          Case VersionScopeEnum.AllVersions
            For Each lobjIVersion As IDocument In lobjIDocument.VersionSeries.Versions
              lobjIVersion = lobjIDocument.VersionSeries.ReleasedVersion
              For Each lobjProperty As ECMProperty In Args.Properties
                SetPropertyValue(lobjIVersion, lobjTargetDocumentClass, lobjProperty, False)
              Next
              ' Save and update property cache.
              lobjIVersion.Save(RefreshMode.REFRESH)
            Next
        End Select
      End If



      Return True

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function

#End Region

End Class

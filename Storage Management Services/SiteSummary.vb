#Region "Imports"

Imports Documents.Utilities
Imports System.Xml.Serialization

#End Region

Public Class SiteSummary

#Region "Class Variables"

  Private mobjEnvironmentSummaries As New EnvironmentSummaries
  Private mobjDateCreated As DateTime = Now

#End Region

#Region "Public Properties"

  Public Property EnvironmentSummaries As EnvironmentSummaries
    Get
      Return mobjEnvironmentSummaries
    End Get
    Set(ByVal value As EnvironmentSummaries)
      mobjEnvironmentSummaries = value
    End Set
  End Property

  <XmlAttribute()>
  Public Property DateCreated As DateTime
    Get
      Return mobjDateCreated
    End Get
    Set(ByVal value As DateTime)
      mobjDateCreated = value
    End Set
  End Property

  <XmlAttribute()>
  Public Property SpaceUsed As String
    Get
      Return GetTotalSpacedUsed()
    End Get
    Set(ByVal value As String)

    End Set
  End Property

  <Xml.Serialization.XmlAttribute()>
  Public Property FileCount As String
    Get
      Return FormatNumber(GetTotalFileCount(), 0)
    End Get
    Set(ByVal value As String)

    End Set
  End Property

#End Region

#Region "Public Methods"

  Public Function GetTotalSpacedUsed() As String
    Try

      Return GetTotalFileSize.ToString

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function GetTotalFileSize() As FileSize
    Try
      Dim lobjFileSize As New FileSize

      For Each lobjStatus As EnvironmentSummary In Me.EnvironmentSummaries
        lobjFileSize += lobjStatus.GetTotalFileSize
      Next

      Return lobjFileSize

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function GetTotalFileCount() As Long
    Try
      Dim llngTotalFiles As Long

      For Each lobjStatus As EnvironmentSummary In Me.EnvironmentSummaries
        llngTotalFiles += lobjStatus.FileCount
      Next

      Return llngTotalFiles

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

End Class

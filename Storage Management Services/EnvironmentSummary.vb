#Region "Imports"

Imports Documents.Utilities
Imports System.Xml.Serialization

#End Region


Public Class EnvironmentSummary

#Region "Class Variables"

  Private mobjEnvironment As Environment = Nothing
  Private mobjObjectStoreSummaries As New List(Of ObjectStoreRollingPolicyStatus)

#End Region

#Region "Public Properties"

  Public Property Environment As Environment
    Get
      Return mobjEnvironment
    End Get
    Set(ByVal value As Environment)
      mobjEnvironment = value
    End Set
  End Property

  <XmlAttribute()>
  Public Property Name As String
    Get
      Return Environment.Name
    End Get
    Set(ByVal value As String)

    End Set
  End Property

  <Xml.Serialization.XmlAttribute()>
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

  Public Property ObjectStoreSummaries As List(Of ObjectStoreRollingPolicyStatus)
    Get
      Return mobjObjectStoreSummaries
    End Get
    Set(ByVal value As List(Of ObjectStoreRollingPolicyStatus))
      mobjObjectStoreSummaries = value
    End Set
  End Property

#End Region

#Region "Constructors"

  Public Sub New()

  End Sub

  Public Sub New(ByVal lpEnvironment As Environment)
    Try
      Environment = lpEnvironment
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

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

      For Each lobjStatus As ObjectStoreRollingPolicyStatus In Me.ObjectStoreSummaries
        lobjFileSize += lobjStatus.TotalSpaceUsed
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

      For Each lobjStatus As ObjectStoreRollingPolicyStatus In Me.ObjectStoreSummaries
        llngTotalFiles += lobjStatus.TotalFiles
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

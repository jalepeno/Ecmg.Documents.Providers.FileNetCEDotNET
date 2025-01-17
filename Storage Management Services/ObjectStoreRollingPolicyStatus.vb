#Region "Imports"

Imports Documents.Utilities
Imports FileNet.Api.Admin
Imports System.Text
Imports System.Xml.Serialization

#End Region

<DebuggerDisplay("{DebuggerIdentifier(),nq}")>
<Xml.Serialization.XmlRoot("StorageStatus")>
Public Class ObjectStoreRollingPolicyStatus

#Region "Class Variables"

  Private mstrEnvironmentName As String = String.Empty
  Private mstrObjectStoreName As String = String.Empty
  Private mintTotalFiles As Integer
  Private mobjTotalSpaceUsed As FileSize
  Private mstrTotalSpaceUsed As String = String.Empty
  Private mstrCurrentStorageAreaPath As String = String.Empty
  Private mintCurrentAreaTotalFiles As Integer
  Private mintCurrentAreaMaximumFiles As Integer
  Private menuStatus As RollingPolicyStatus
  Private mobjNewStorageArea As IFileStorageArea
  Private mstrRollingPolicyName As String = String.Empty
  Private msrCurrentRollingStorageAreaName As String = String.Empty

#End Region

#Region "Public Properties"

  <Xml.Serialization.XmlAttribute()>
  Public Property ObjectStore As String
    Get
      Return String.Format("{0} - {1}", EnvironmentName, ObjectStoreName)
    End Get
    Set(ByVal value As String)

    End Set
  End Property

  <Xml.Serialization.XmlAttribute()>
  Public Property FileCount As String
    Get
      Return FormatNumber(TotalFiles, 0)
    End Get
    Set(ByVal value As String)

    End Set
  End Property


  <Xml.Serialization.XmlAttribute()>
  Public Property SpaceUsed As String
    Get
      Return TotalSpaceUsedLabel
    End Get
    Set(ByVal value As String)

    End Set
  End Property

  Public Property EnvironmentName As String
    Get
      Return mstrEnvironmentName
    End Get
    Set(ByVal value As String)
      mstrEnvironmentName = value
    End Set
  End Property

  Public Property ObjectStoreName As String
    Get
      Return mstrObjectStoreName
    End Get
    Set(ByVal value As String)
      mstrObjectStoreName = value
    End Set
  End Property

  Public Property TotalFiles As Integer
    Get
      Return mintTotalFiles
    End Get
    Set(ByVal value As Integer)
      mintTotalFiles = value
    End Set
  End Property

  <XmlIgnore()>
  Public Property TotalSpaceUsed As FileSize
    Get
      Return mobjTotalSpaceUsed
    End Get
    Set(ByVal value As FileSize)
      mobjTotalSpaceUsed = value
    End Set
  End Property

  <XmlElement("TotalSpaceUsed")>
  Public Property TotalSpaceUsedLabel As String
    Get
      Return mobjTotalSpaceUsed.ToString
    End Get
    Set(ByVal value As String)
      Try
        If mobjTotalSpaceUsed Is Nothing Then
          mobjTotalSpaceUsed = FileSize.FromString(value)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Set
  End Property

  Public Property CurrentStorageAreaPath As String
    Get
      Return mstrCurrentStorageAreaPath
    End Get
    Set(ByVal value As String)
      mstrCurrentStorageAreaPath = value
    End Set
  End Property

  Public Property CurrentAreaTotalFiles As Integer
    Get
      Return mintCurrentAreaTotalFiles
    End Get
    Set(ByVal value As Integer)
      mintCurrentAreaTotalFiles = value
    End Set
  End Property

  Public Property CurrentAreaMaximumFiles As Integer
    Get
      Return mintCurrentAreaMaximumFiles
    End Get
    Set(ByVal value As Integer)
      mintCurrentAreaMaximumFiles = value
    End Set
  End Property

  Public Property CurrentPercentageUsed As Single
    Get
      Try
        Return CalculateCurrentPercentageUsed()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Get
    Set(ByVal value As Single)

    End Set
  End Property

  Public Property Status As RollingPolicyStatus
    Get
      Return menuStatus
    End Get
    Set(ByVal value As RollingPolicyStatus)
      menuStatus = value
    End Set
  End Property

  <Xml.Serialization.XmlIgnore()>
  Public Property NewStorageArea As IFileStorageArea
    Get
      Return mobjNewStorageArea
    End Get
    Set(ByVal value As IFileStorageArea)
      mobjNewStorageArea = value
    End Set
  End Property

  Public Property RollingPolicyName As String
    Get
      Return mstrRollingPolicyName
    End Get
    Set(ByVal value As String)
      mstrRollingPolicyName = value
    End Set
  End Property

  Public Property CurrentRollingStorageAreaName As String
    Get
      Return msrCurrentRollingStorageAreaName
    End Get
    Set(ByVal value As String)
      msrCurrentRollingStorageAreaName = value
    End Set
  End Property

#End Region

#Region "Constructors"

  Public Sub New()
    MyBase.New()
  End Sub

  Public Sub New(ByVal lpEnvironmentName As String,
                 ByVal lpObjectStoreName As String,
                 ByVal lpRollingPolicyName As String,
                 ByVal lpCurrentRollingStorageAreaName As String,
                 ByVal lpTotalFiles As Integer,
                 ByVal lpTotalSpaceUsed As Long,
                 ByVal lpCurrentStorageAreaPath As String,
                 ByVal lpCurrentTotalFiles As Integer,
                 ByVal lpCurrentMaximumFiles As Integer,
                 ByVal lpStatus As RollingPolicyStatus)
    Me.New(lpEnvironmentName, lpObjectStoreName, lpRollingPolicyName,
           lpCurrentRollingStorageAreaName, lpTotalFiles, lpTotalSpaceUsed,
           lpCurrentStorageAreaPath, lpCurrentTotalFiles, lpCurrentMaximumFiles,
           lpStatus, Nothing)
  End Sub

  Public Sub New(ByVal lpEnvironmentName As String,
                 ByVal lpObjectStoreName As String,
                 ByVal lpRollingPolicyName As String,
                 ByVal lpCurrentRollingStorageAreaName As String,
                 ByVal lpTotalFiles As Integer,
                 ByVal lpTotalSpaceUsed As Long,
                 ByVal lpCurrentStorageAreaPath As String,
                 ByVal lpCurrentTotalFiles As Integer,
                 ByVal lpCurrentMaximumFiles As Integer,
                 ByVal lpStatus As RollingPolicyStatus,
                 ByVal lpNewStorageArea As IFileStorageArea)
    Try
      mstrEnvironmentName = lpEnvironmentName
      mstrObjectStoreName = lpObjectStoreName
      mintTotalFiles = lpTotalFiles
      If Not String.IsNullOrEmpty(lpTotalSpaceUsed) Then
        mobjTotalSpaceUsed = New FileSize(lpTotalSpaceUsed)
      Else
        mobjTotalSpaceUsed = New FileSize()
      End If
      mstrCurrentStorageAreaPath = lpCurrentStorageAreaPath
      mintCurrentAreaTotalFiles = lpCurrentTotalFiles
      mintCurrentAreaMaximumFiles = lpCurrentMaximumFiles
      menuStatus = lpStatus
      If lpNewStorageArea IsNot Nothing Then
        NewStorageArea = lpNewStorageArea
      End If
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

#Region "Private Methods"

  Private Function CalculateCurrentPercentageUsed() As Single
    Try
      If CurrentAreaMaximumFiles = 0 Then
        Return 0
      End If

      Return (CurrentAreaTotalFiles / CurrentAreaMaximumFiles) * 100

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Friend Function DebuggerIdentifier() As String

    Dim lstrReturnBuilder As New StringBuilder

    Try

      If Not String.IsNullOrEmpty(Me.RollingPolicyName) Then
        lstrReturnBuilder.AppendFormat("{0}: {1} <{2} {3} space used; {4}% used>",
                                       Me.EnvironmentName,
                                       Me.ObjectStoreName,
                                       Me.RollingPolicyName,
                                       Me.TotalSpaceUsed,
                                       Me.CurrentPercentageUsed)

      Else
        lstrReturnBuilder.AppendFormat("{0}: {1} <{2}space used>",
                               Me.EnvironmentName,
                               Me.ObjectStoreName,
                               Me.TotalSpaceUsed)

      End If

      Return lstrReturnBuilder.ToString

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

End Class

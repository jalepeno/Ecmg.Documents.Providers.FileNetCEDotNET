#Region "Imports"

Imports Documents.Utilities

#End Region

Public Class Environments
  Inherits List(Of Environment)

  Public Sub InitializeConnections()
    Try
      For Each lobjEnvironment As Environment In Me
        lobjEnvironment.InitializeConnection()
      Next
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

End Class
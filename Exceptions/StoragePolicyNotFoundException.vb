Namespace Exceptions

  Public Class StoragePolicyNotFoundException
    Inherits ItemNotFoundException

#Region "Constructors"

    Public Sub New(ByVal itemName As String)
      MyBase.New(String.Format("Storage policy '{0}' not found.", itemName), itemName)
    End Sub

    Public Sub New(ByVal message As String, ByVal itemName As String)
      MyBase.New(message, itemName)
    End Sub

    Public Sub New(ByVal itemName As String, ByVal innerException As Exception)
      MyBase.New(String.Format("Storage policy '{0}' not found.", itemName), innerException)
    End Sub

    Public Sub New(ByVal message As String, ByVal itemName As String, ByVal innerException As Exception)
      MyBase.New(message, itemName, innerException)
    End Sub

#End Region

  End Class

End Namespace
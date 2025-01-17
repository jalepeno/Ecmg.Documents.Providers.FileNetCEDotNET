
Imports Documents.Exceptions


Namespace Exceptions

  Public Class ItemNotFoundException
    Inherits CtsException

#Region "Class Variables"

    Private mstrItemName As String

#End Region

#Region "Public Properties"

    Public ReadOnly Property ItemName As String
      Get
        Return mstrItemName
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal itemName As String)
      MyBase.New(String.Format("Item '{0}' not found.", itemName))
      mstrItemName = itemName
    End Sub

    Public Sub New(ByVal message As String, ByVal itemName As String)
      MyBase.New(message)
      mstrItemName = itemName
    End Sub

    Public Sub New(ByVal itemName As String, ByVal innerException As Exception)
      MyBase.New(String.Format("Item '{0}' not found.", itemName), innerException)
      mstrItemName = itemName
    End Sub

    Public Sub New(ByVal message As String, ByVal itemName As String, ByVal innerException As Exception)
      MyBase.New(message, innerException)
      mstrItemName = itemName
    End Sub

#End Region

  End Class

End Namespace
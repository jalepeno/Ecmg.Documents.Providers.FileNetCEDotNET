' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  CENetProvider_IRecordsManager.vb
'  Description :  [type_description_here]
'  Created     :  7/25/2012 2:05:16 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents
Imports Documents.Exceptions
Imports Documents.Records
Imports Documents.Annotations
Imports Documents.Utilities
Imports ExcelDataReader.Core.OpenXmlFormat
Imports System

#End Region

Partial Public Class CENetProvider
  Implements IRecordsManager

#Region "IRecordsManager Implementation"

  Public Function AddPhysicalRecord(ByRef Args As Records.AddPhysicalRecordArgs) As Boolean _
    Implements Records.IRecordsManager.AddPhysicalRecord
    Try
      Dim lobjDeclaration As New Declaration
      lobjDeclaration.AddPhysicalRecord(Args)
      Return True
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  Public Function DeclareRecord(ByRef Args As Records.DeclareRecordArgs) As Boolean _
    Implements Records.IRecordsManager.DeclareRecord
    Try
      Throw New NotImplementedException
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

#End Region

End Class

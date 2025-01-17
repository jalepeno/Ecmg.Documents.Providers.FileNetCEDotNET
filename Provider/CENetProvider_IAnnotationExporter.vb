'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  CENetProvider_IAnnotationExporter.vb
'   Description :  [type_description_here]
'   Created     :  4/8/2014 11:20:38 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Providers
Imports Documents.Utilities
Imports Documents.Arguments

#End Region

Partial Public Class CENetProvider
  Implements IAnnotationExporter

#Region "IAnnotationExporter Implementation"
#Region "Events"

  Event AnnotationExported As AnnotationExportEventHandler Implements IAnnotationExporter.AnnotationExported

  Event AnnotationExportError As AnnotationExportEventHandler Implements IAnnotationExporter.AnnotationExportError

  Event AnnotationExportMessage As AnnotationExportMessageEventHandler Implements IAnnotationExporter.AnnotationExportMessage

#End Region

#Region "Methods"

#Region "Event Handler Methods"

  Sub OnAnnotationExported(ByRef e As AnnotationExportedEventArgs) _
  Implements IAnnotationExporter.OnAnnotationExported
    Try
      Throw New NotImplementedException()
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Sub

  Sub OnAnnotationExportError(ByRef e As AnnotationExportErrorEventArgs) _
  Implements IAnnotationExporter.OnAnnotationExportError
    Try
      Throw New NotImplementedException()
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

#Region "Methods"

  Public Function ExportAnnotation(ByVal lpId As String) As Boolean _
         Implements IAnnotationExporter.ExportAnnotation
    'Try
    '  Return ExportAnnotation(New ExportAnnotationEventArgs(lpId))
    'Catch ex As System.Exception
    '  ApplicationLogging.LogException(ex, String.Format("CSProvider::ExportAnnotation('{0}')", lpId))
    'End Try   
    Try
      Throw New NotImplementedException()
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try


  End Function

  Public Function ExportAnnotation(ByVal Args As ExportAnnotationEventArgs) As Boolean _
         Implements IAnnotationExporter.ExportAnnotation
    'Dim result As Boolean = False

    'Try
    '  Dim csAnnotation As IDMObjects.Annotation = Me.GetIdmAnnotation(Args.ID)
    '  ' Args.Annotation = Me.lobjAnnotationExporter.ExportAnnotationObject(csAnnotation)
    '  result = True
    'Catch e As COMException
    '  ApplicationLogging.LogException(e, String.Format("CSProvider::ExportAnnotation('{0}')", Args.ID))
    'End Try

    'Return result
    Try
      Throw New NotImplementedException()
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try


  End Function

  '    Function ExportAnnotations(ByVal Args As ExportAnnotationsEventArgs) As Boolean

#End Region

#End Region

#End Region

End Class

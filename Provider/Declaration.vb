
#Region "Imports"

Imports Documents
Imports Documents.Core
Imports Documents.Providers
Imports Documents.Records
Imports Documents.Utilities
Imports System.Text
Imports System.Data.OleDb
Imports System.Data
Imports FileNet.Api.Admin
Imports FileNet.Api.Core
Imports FileNet.Api.Property

#End Region

Public Class DeclarationException
  Inherits InvalidOperationException

#Region "Class Variables"

  Private lenuErrorType As Declaration.ErrorType

#End Region

#Region "Public Properties"

  Public ReadOnly Property ErrorType() As Declaration.ErrorType
    Get
      Return lenuErrorType
    End Get
  End Property

#End Region

#Region "Constructors"

  Public Sub New(ByVal errorType As Declaration.ErrorType)
    MyBase.New()
    lenuErrorType = errorType
  End Sub

  Public Sub New(ByVal message As String, ByVal errorType As Declaration.ErrorType)
    MyBase.New(message)
    lenuErrorType = errorType
  End Sub

  Public Sub New(ByVal message As String, ByVal innerException As System.Exception, ByVal errorType As Declaration.ErrorType)
    MyBase.New(message, innerException)
    lenuErrorType = errorType
  End Sub

#End Region

End Class

Public Class Declaration

#Region "Class Enumerations"

  Public Enum ImplementationType
    '            If (lpImplementationType = "DOD-5015.2" Or lpImplementationType = "Base" Or lpImplementationType = "DOD-5015.2CHAPTER4") Then  'DOD
    Base = 0
    Pro = 1
    DOD5015Point2 = 3
    DOD5015Point2Chapter4 = 4
  End Enum

  Public Enum RecordType
    Invalid = -1
    Generic = 100
    RecordCategory = 101
    RecordFolder = 102
    Volume = 103
    ElectronicRecordFolder = 105
    PhysicalContainer = 106
    HybridRecordFolder = 108
    RMFolder = 109
    PhysicalRecordFolder = 110
    Record = 300
    ElectronicRecordInfo = 301
    EmailRecord = 302
    Marker = 303
  End Enum

  ''' <summary>
  ''' Enumeration of the various errors associated with a DeclarationException
  ''' </summary>
  ''' <remarks></remarks>
  Public Enum ErrorType

    MissingRequiredParameter_RecordClass = 1002
    MissingRequiredParameter_RecordType = 1004
    MissingRequiredParameter_RecordFolderPath = 1005
    MissingRequiredParameter_SourceDocument = 1006

    InvalidRecordType_RecordTypeNotValidForHybridRecordFolder = 1024
    UnableToSetThePropertiesOnTheDocument = 1042

    UnableToFindTargetRecordFolder = 1090
    UnableToCreateTheRecordObjectWithTheGivenClassID = 1099

    DocumentCantBeDeclaredAsCanDeclarePropertyIsNotDefinedOrIsFalse = 1100
    DocumentIsAlreadyDeclaredAsRecord = 1101
    DocumentDannotBeDeclaredAsTheDocumentIsCurrentlyInReservedState = 1102

    UnableToLocateFilePlanObjectStore = 1200
    TriggerAndDisposalActionPropertiesRequiredForVitalRecord = 1400

    RecordCategoryIsClosedOrInactive = 1501
    RecordMayNotBeDeclaredInTheRecordCategoryInDoDFilePlanObjectStore = 1502
    ContainerIsClosed = 1503
    InvalidRecordType_RecordTypeNotValidForElectronicRecordFolder = 1504
    InvalidRecordType_RecordTypeNotValidForPhysicalRecordFolder = 1505
    InvalidRecordType_RecordTypeNotValidForPhysicalContainer = 1506
    InvalidRecordType_RecordTypeNotAllowedForContainer = 1507

    AccessDeniedForFolder = 1510

    GenericConnectorRegistrationObjectNotFound = 1600

    UnableToSetTheSecurityParentForTheRecord = 1800
    'throwError(1900, "getActiveVolume", "Record can not be declared in the Volume at Path " + loLastCreatedVolume.Name + " because it is not active. It is closed.")

    UnableToAddDocument = 1198
    UnableToDeclareRecord = 1199
    ErrorDeletingRecord = 4000

    UnableToRetrieveLocation = 5001

  End Enum

#End Region

#Region "Class Constants"

  Const RM_TYPE_ELECTRONICRECORDFOLDER_CLASS As String = "Electronic Record Folder"
  Const RM_TYPE_PHYSICALCONTAINER_CLASS As String = "Physical Box"
  Const RM_TYPE_BOX_CLASS As String = "Box"
  Const RM_TYPE_HYBRIDRECORDFOLDER_CLASS As String = "Hybrid Record Folder"
  Const RM_TYPE_FOLDER_CLASS As String = "Folder"
  Const RM_TYPE_PHYSICALRECORDFOLDER_CLASS As String = "Physical Record Folder"
  Const RM_TYPE_VOLUME_CLASS As String = "Volume"
  Const RM_TYPE_RECORDCATEGORY_CLASS As String = "Record Category"

  Const VITAL_RECORD_DISPOSAL_TRIGGER As String = "VitalRecordDisposalTrigger"
  Const VITAL_RECORD_REVIEW_ACTION As String = "VitalRecordReviewAction"

  Const IDM_CHANGE_NO_REFRESH = 1
  Const IDM_CONTAINMENT_DEFINE_SECURITY_PARENTAGE = 1024
  Const IDM_TAKE_FEDERATED_OWNERSHIP = 262144


#End Region

#Region "Class Variables"

  Private mobjROSContentSource As ContentSource
  Private mobjFPOSContentSource As ContentSource
  Private mobjArguments As Records.DeclareRecordArgs
  Private mobjFolderClasses As New HierarchyItemCollection

#End Region

#Region "Public Properties"

  Public Property ROS_ContentSource() As ContentSource
    Get
      Return mobjROSContentSource
    End Get
    Set(ByVal value As ContentSource)
      mobjROSContentSource = value
    End Set
  End Property

  Public Property FilePlanContentSource() As ContentSource
    Get
      Return mobjFPOSContentSource
    End Get
    Set(ByVal value As ContentSource)
      mobjFPOSContentSource = value
    End Set
  End Property

  Public ReadOnly Property FilePlanObjectStoreName() As String
    Get
      Try
        Return FilePlanContentSource.Properties("ObjectStore").PropertyValue
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return "Object Store Name Unavailable"
      End Try
    End Get
  End Property

  Public ReadOnly Property RecordObjectStoreName() As String
    Get
      Try
        Return ROS_ContentSource.Properties("ObjectStore").PropertyValue
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return "Object Store Name Unavailable"
      End Try
    End Get
  End Property

  Public ReadOnly Property Arguments() As Records.DeclareRecordArgs
    Get
      Return mobjArguments
    End Get
  End Property

#End Region

#Region "Private Methods"

  Private Function CheckRequiredArgs(ByVal Args As Records.DeclareRecordArgs) As Boolean
    Try

      ' First set the Arguments property of the class
      mobjArguments = Args

      ' Check the inputs to the declare method
      If Helper.CheckForNullString(Args.RecordClass) = "" Then
        Throw New DeclarationException("Provide the RecordInfo Class ID using which record is to created.  It is required parameter.",
                                       ErrorType.MissingRequiredParameter_RecordClass)
      End If

      If Helper.CheckForNullString(Args.RecordType) = "" Then
        Throw New DeclarationException("Provide the record type of which the record is to be created.  It is required parameter.",
                                       ErrorType.MissingRequiredParameter_RecordType)
      End If

      If Helper.CheckForNullString(Args.RecordFolderPath) = "" Then
        Throw New DeclarationException("Provide the name of Record Folder or Record category or Volume  where record is to be created. It is required parameter.",
                                       ErrorType.MissingRequiredParameter_RecordFolderPath)
      End If

      Select Case Args.GetType.Name
        Case "DeclareRecordArgs"
          If Args.SourceDocument Is Nothing Then
            Throw New DeclarationException("Provide the document which is to be declared as record. It is required parameter.",
                                           ErrorType.MissingRequiredParameter_SourceDocument)
          End If
        Case "AddPhysicalRecordArgs"
          ' We will skip this check as we do not require the source document in this case
      End Select

      Return True

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  ''' <summary>
  ''' This method declares the document specified by ‘lpDocumentLocationURI’ as a record and 
  ''' files it to the given object.   The 'lpRecordInfoClassID' specifies the type of Record 
  ''' to be created (i.e. the class / subclass of ElectronicRecordInfo, Marker, EmailRecordInfo).  
  ''' During the Multiple filing operations, in case one of the filing fails, 
  ''' the whole declare operation will be rolled back.
  ''' </summary>
  ''' <param name="lpFilePlanObjectStore">The Object store of the Document</param>
  ''' <param name="lpDocumentLocationURI">The URL of the document. The NUll or empty URL will not be allowed</param>
  ''' <param name="lpRecordRMType">The Type of the Record</param>
  ''' <param name="lpRecordInfoClassID">The ID of the ElectronicRecordInfo,  Marker, and EmailRecordInfo. The NUll or empty ClassID will not be allowed</param>
  ''' <param name="lpCreateRecordInFolder">The  RecordFolder where this record needs to filed in during declaration</param>
  ''' <param name="lpSourceDocument"></param>
  ''' <param name="lpRecordProperties">Properties of the Record</param>
  ''' <param name="lpImplementationType">The implementation type of the Object Store i.e. DoD, Base or PRO</param>
  ''' <returns>RecordInfo objects created during the declaration</returns>
  ''' <remarks></remarks>
  Private Function createRecordInfo(ByVal lpFilePlanObjectStore As String,
                                    ByVal lpDocumentLocationURI As String,
                                    ByVal lpRecordRMType As RecordType,
                                    ByVal lpRecordInfoClassID As String,
                                    ByVal lpCreateRecordInFolder As Object,
                                    ByVal lpSourceDocument As Object,
                                    ByVal lpRecordProperties As Object,
                                    ByVal lpImplementationType As String) As Records.Record
    ' TODO: Port, test and clean the createRecordInfo method
    Try

      Dim lobjRecord As Records.Record = Arguments.NewRecord
      lobjRecord.DocumentClass = Arguments.RecordClass

      ' Set a property for the date filed in the record
      If (lobjRecord.PropertyExists("DateDeclared", False) = False) Then
        lobjRecord.SetPropertyValue(PropertyScope.VersionProperty, "DateDeclared", Now.ToUniversalTime, True, PropertyType.ecmDate)
      End If


      lobjRecord.ContentSource = Me.FilePlanContentSource

      ' Create a new RecordInfo using the the ClassID provided
      'lobjRecord = lpFilePlanObjectStore.CreateObject(lpRecordInfoClassID)

      If AddRecord(lobjRecord, lpDocumentLocationURI, lpRecordRMType, lpCreateRecordInFolder, lpSourceDocument, lpRecordProperties, lpImplementationType) Then
        Return lobjRecord
      Else
        Return Nothing
      End If

    Catch DeclareEx As DeclarationException
      ApplicationLogging.LogException(DeclareEx, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      Throw New DeclarationException(String.Format("Unable to create the Record object using class ID '{0}' in ObjectStore '{1}'",
                                          lpRecordInfoClassID, lpFilePlanObjectStore), ex, ErrorType.UnableToCreateTheRecordObjectWithTheGivenClassID)

    End Try

  End Function

  Private Function AddRecord(ByVal lpRecord As Record,
                                       ByVal lpDocumentLocationURI As String,
                                       ByVal lpRecordRMType As RecordType,
                                       ByVal lpCreateRecordInFolder As CFolder,
                                       ByVal lpSourceDocument As Document,
                                       ByVal lpRecordProperties As ECMProperties,
                                       ByVal lpImplementationType As ImplementationType) As Boolean
    ' TODO: Document the setRecordProperties method
    ' TODO: Port, test and clean setRecordProperties method
    Try

      Dim lstrErrorMessage As String = ""
      Dim lobjProvider As CENetProvider = Me.FilePlanContentSource.Provider

      ' Make sure the record is marked for creation as a major version
      lpRecord.SetPropertyValue(PropertyScope.VersionProperty, "MajorVersion", True, True, PropertyType.ecmBoolean)

      'Dim RM_TYPE_ELECTRONICRECORD
      'Dim RM_TYPE_EMAILRECORD
      'Dim RM_TYPE_MARKER
      'RM_TYPE_ELECTRONICRECORD = 301  'Electronic Record Info
      'RM_TYPE_EMAILRECORD = 302   'Email Record Info
      'RM_TYPE_MARKER = 303   'Marker

      ' Set the important Record related properties
      ''LogError "Document URI " + asDocumentLocationURI
      ' CRITICAL SECTION
      ' Do NOT Move lines in between up and down this has been done after good amount of consideration

      ' First set the Properties as the default values
      ' some of the record object properties are set to default values.
      ' the from property on the record object is set with the creator of the source document.

      For Each loPassedProperty As ECMProperty In lpRecordProperties
        If loPassedProperty.Name <> "RMEntityType" Then
          'Dim loProperty
          'loProperty = lpRecord.Properties.Item(loPassedProperty)
          'Dim loPropDesc
          'loPropDesc = loProperty.PropertyDescription


          lpRecord.GetLatestVersion.Properties.Add(loPassedProperty)
          'If loPropDesc.Cardinality = 0 Then
          '  lpRecord.Properties.Item(loPassedProperty) = lpRecordProperties.Item(loPassedProperty)
          'ElseIf loPropDesc.Cardinality = 1 Or loPropDesc.Cardinality = 2 Then

          '  Dim loRecVals
          '  loRecVals = loProperty.Value
          '  Dim liCounter
          '  Dim loArray
          '  loArray = aoRecordProperties.Item(loPassedProperty)
          '  For liCounter = 1 To UBound(loArray)
          '    loRecVals.Add(loArray(liCounter), liCounter)
          '  Next

          'End If
        End If
      Next

      If lpRecord.GetType.Name <> "PhysicalRecord" Then
        ' This is not a physical record

        lpRecord.SetPropertyValue(PropertyScope.VersionProperty, "From", lpSourceDocument.GetLatestVersion.Properties("Creator").Value, True, PropertyType.ecmString)
        ' the document title of the record object is set to the document title of the source document.
        lpRecord.SetPropertyValue(PropertyScope.VersionProperty, "DocumentTitle", lpSourceDocument.Name, True, PropertyType.ecmString)
        ' the Subject of the record object is changed to EmailSubject of the source document.
        lpRecord.SetPropertyValue(PropertyScope.VersionProperty, "EmailSubject", lpSourceDocument.Name, True, PropertyType.ecmString)
        ' Publication Date name is changed to Sent On
        If (lpRecord.PropertyExists("SentOn", False) = False) Then
          lpRecord.SetPropertyValue(PropertyScope.VersionProperty, "SentOn", lpSourceDocument.GetLatestVersion.Properties("DateCreated").Value, True, PropertyType.ecmDate)
        End If

        ' the owner of record object is set to the same value as the owner of the source document.
        'lprecord.Owner = SourceDocument.Owner 
        ' the reviewer of the record object is set to the creator of the source document.
        lpRecord.SetPropertyValue(PropertyScope.VersionProperty, "Reviewer", lpSourceDocument.GetLatestVersion.Properties("Creator").Value, True, PropertyType.ecmString)

        ' Set all the required properties as per the implementation type
        'If (asImplementationType = "DOD-5015.2" Or asImplementationType = "Base" Or asImplementationType = "DOD-5015.2CHAPTER4") Then  'DOD or Base
        If (lpImplementationType = ImplementationType.DOD5015Point2 OrElse
            lpImplementationType = ImplementationType.Base OrElse
            lpImplementationType = ImplementationType.DOD5015Point2Chapter4) Then  'DOD or Base
          If (lpSourceDocument.GetLatestVersion.Contents.Count > 0) Then
            lpRecord.SetPropertyValue(PropertyScope.VersionProperty, "MediaType", lpSourceDocument.GetLatestVersion.Contents(0).MIMEType + "", True, PropertyType.ecmString)
            lpRecord.SetPropertyValue(PropertyScope.VersionProperty, "Format", lpSourceDocument.GetLatestVersion.Contents(0).MIMEType + "", True, PropertyType.ecmString)
          ElseIf (lpSourceDocument.PropertyExists(PropertyScope.VersionProperty, "MIMEType", False)) Then
            lpRecord.SetPropertyValue(PropertyScope.VersionProperty, "MediaType", lpSourceDocument.GetLatestVersion.Properties("MIMEType").Value + "", True, PropertyType.ecmString)
            lpRecord.SetPropertyValue(PropertyScope.VersionProperty, "Format", lpSourceDocument.GetLatestVersion.Properties("MIMEType").Value + "", True, PropertyType.ecmString)
          End If
          ' TODO: Create and add the CE Location property

        Else
          'IN PRO no extra properties are otherwise we need to put those here . OK.
        End If
        ' Populate the Record Propertiesa as per the Values Passed by the Caller

        ' Set all values which must be as per the declaration
        If Trim(lpDocumentLocationURI) = "" Then
          lpRecord.SetPropertyValue(PropertyScope.VersionProperty, "DocURI", "", True, PropertyType.ecmString)
        Else
          lpRecord.SetPropertyValue(PropertyScope.VersionProperty, "DocURI", lpDocumentLocationURI, True, PropertyType.ecmString)
        End If
        If (lpRecord.PropertyExists("DateDeclared", False) = False) Then
          lpRecord.SetPropertyValue(PropertyScope.VersionProperty, "DateDeclared", Now(), True, PropertyType.ecmDate)
        End If

        ' Set the mimeType of the Record
        If lpRecordRMType = RecordType.EmailRecord Then
          'lpRecord.ChangeContentMimeType("application/x-filenet-rm-emailrecord", Document.ALL_VERSIONS, 0)
          lpRecord.SetPropertyValue(PropertyScope.VersionProperty, "MimeType", "application/x-filenet-rm-emailrecord", True, PropertyType.ecmString)
        ElseIf lpRecordRMType = RecordType.ElectronicRecordInfo Then
          'lpRecord.ChangeContentMimeType("application/x-filenet-rm-electronicrecord", Document.ALL_VERSIONS, 0)
          lpRecord.SetPropertyValue(PropertyScope.VersionProperty, "MimeType", "application/x-filenet-rm-electronicrecord", True, PropertyType.ecmString)
        ElseIf lpRecordRMType = RecordType.Marker Then
          'lpRecord.ChangeContentMimeType("application/x-filenet-rm-physicalrecord", Document.ALL_VERSIONS, 0)
          'lpRecord.SetPropertyValue(PropertyScope.VersionProperty, "MimeType", "application/x-filenet-rm-physicalrecord", True, PropertyType.ecmString)
        End If

      Else
        ' This is a physical record
        Dim lobjCELocation As Object

        ' Make sure we set a value for 'Reviewer'
        If lpRecord.PropertyExists("Reviewer") = False Then
          'lpRecord.SetPropertyValue(PropertyScope.VersionProperty, "Reviewer", lpSourceDocument.GetLatestVersion.Properties("Creator").Value, True, PropertyType.ecmString)
          'Beep()
        End If

        lpRecord.SetPropertyValue(PropertyScope.VersionProperty, "DocURI", "Physical Record", True, PropertyType.ecmString)


        Dim lobjPhysicalRecord As PhysicalRecord = CType(lpRecord, PhysicalRecord)

        Try
          lobjCELocation = GetCELocation(lobjPhysicalRecord.Location, lstrErrorMessage)
        Catch ex As Exception
          Throw New DeclarationException(ex.Message, ex, ErrorType.UnableToRetrieveLocation)
        End Try

        'Dim lobjCEWebServices As New Utilities.CEWebServices(lobjProvider.URL, lobjProvider.UserName, lobjProvider.Password, lobjProvider.ObjectStore)
        'Dim lobjLocation As Object = lobjCEWebServices.GetObject(lobjCELocation.objectid, lobjCELocation.classid, lstrErrorMessage)

        ' Create the properties of the new Location

        'Dim lobjPropertyTemplate As IPropertyTemplate = (From pt In CType(Me.FilePlanContentSource.Provider, CENetProvider).ObjectStore.PropertyTemplates Where pt.SymbolicName = "Location" Select pt).FirstOrDefault
        'Dim lobjObjectValue As IIndependentObject = GetObject(lobjPropertyTemplate.ClassDescription.Id.ToString, lobjCELocation.Id.ToString)

        Dim lobjLocationReference As String = CType(Me.FilePlanContentSource.Provider, CENetProvider).GetObjectDescriptor(lobjCELocation)

        'Dim lobjLocationReference As IPropertyEngineObject = Factory.CustomObject.CreateInstance(CType(Me.FilePlanContentSource.Provider, CENetProvider).ObjectStore, "Location")
        'Dim lobjLocation As Object = New Object
        'With lobjLocationReference
        '  .classId = lobjCELocation.ClassDescription.Id
        '  .Id = lobjCELocation.Id
        '  '.objectStore = FilePlanContentSource.Properties("ObjectStore").PropertyValue
        'End With

        'lobjLocation.propertyId = "Location"
        'lobjLocation.Value = lobjLocationReference

        'Dim lobjCleanLocation As Object = lobjCEWebServices.ClearUnevaluatedValues(lobjLocation)

        lpRecord.SetPropertyValue(PropertyScope.VersionProperty, "Location", lobjLocationReference, True, PropertyType.ecmObject)
        lpRecord.SetPropertyValue(PropertyScope.VersionProperty, "HomeLocation", lobjLocationReference, True, PropertyType.ecmObject)

      End If

      'lpRecord.checkin()

      ' Create a reference to the web service utility
      'Dim lobjCEWebServices As New Utilities.CEWebServices(lobjProvider.URL, lobjProvider.UserName, lobjProvider.Password, lobjProvider.ObjectStore)
      'Dim lobjCEWebServices As Utilities.CEWebServices = Utilities.CEWebServices.CreateInstance(lobjProvider)
      ' Create a variable for the return value
      Dim lblnReturnValue As Boolean

      Try


        ' Then check to see if the properties are valid
        'Dim lobjInvalidProperties As ECMProperties = CType(lobjProvider, CProvider). _
        'FindInvalidProperties(lpRecord, PropertyScope.BothDocumentAndVersionProperties)

        'lblnReturnValue = DeclareRecordUsingJavaAPI(lpRecord.GetLatestVersion.Properties, lpCreateRecordInFolder, lpRecord.DocumentClass, lpRecord)

        ' Attempt to add the document
        lblnReturnValue = CType(Me.FilePlanContentSource.Provider, CENetProvider).AddDocument(lpRecord, lpCreateRecordInFolder.Path, True)

        'lobjProvider.ImportDocument .AddDocument(lpRecord,lpCreateRecordInFolder.Path,
        'lpRecord.Add(lpCreateRecordInFolder.Path)
      Catch ex As Exception
        'DeleteRecord(lpRecord)
        lstrErrorMessage = ex.Message
        Throw New DeclarationException(lstrErrorMessage, ErrorType.UnableToAddDocument)
      End Try

      If lblnReturnValue = False OrElse lstrErrorMessage.Length > 0 AndAlso lstrErrorMessage.ToLower.StartsWith("successfully filed a document") = False Then
        If lstrErrorMessage.ToLower.StartsWith("access is denied") Then
          Throw New DeclarationException(lstrErrorMessage, ErrorType.AccessDeniedForFolder)
        Else
          Throw New DeclarationException(lstrErrorMessage, ErrorType.UnableToDeclareRecord)
        End If
      End If



      'Now that we have declared(added the record), we need to go back to the source document
      'and fill in the RecordInformation fields.
      'Try
      '  Dim lstrROSObjectStoreGuid As String = String.Empty
      '  If (lpSourceDocument IsNot Nothing) Then
      '    If (lpSourceDocument.ContentSource IsNot Nothing) Then
      '      If (lpSourceDocument.ContentSource.Provider IsNot Nothing) Then
      '        Dim lobjROSProvider As CEWSI35Provider = lpSourceDocument.ContentSource.Provider
      '        lstrROSObjectStoreGuid = lobjCEWebServices.GetObjectStoreGuid(lobjROSProvider.ObjectStore)
      '      End If
      '    End If
      '  End If
      '  If (lstrROSObjectStoreGuid = String.Empty) Then
      '    lstrROSObjectStoreGuid = "00000000-0000-0000-0000-000000000000"
      '    Helper.WriteLogEntry("AddRecord:: The ROSObjectStoreGuid is empty because the source document's contentsource is set to nothing", TraceEventType.Warning)
      '  End If
      '  UpdateRecordInformationFields(lobjProvider.ROSDBConnection, lobjProvider.FPOSDBConnection, lobjCEWebServices.GetObjectStoreGuid(lobjProvider.ObjectStore), lpRecord.ObjectID, lpSourceDocument.ObjectID, lstrROSObjectStoreGuid)
      'Catch ex As Exception
      '  ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Throw New DeclarationException(ex.Message, ErrorType.UnableToDeclareRecord)
      'End Try



      'Dim lobjRIObject As Object = Nothing
      'Dim lobjDocumentObject As Object = Nothing
      'Dim lobjErrormsg As String = String.Empty
      'Dim lpUpdateArgs As New Arguments.DocumentPropertyArgs(lpSourceDocument.ObjectID)
      'Dim lobjROSCEWebServices As New Utilities.CEWebServices(lobjProvider.URL, lobjProvider.UserName, lobjProvider.Password, "ROSSpectraDev")
      'Try
      '  lobjRIObject = lobjCEWebServices.GetObject(lpRecord.ObjectID, "RecordInfo", lobjErrormsg)
      '  lobjDocumentObject = lobjROSCEWebServices.GetRecordInformation(lpSourceDocument.ObjectID)
      '  Dim lobjProperty As New ECMProperty(Core.PropertyType.ecmObject, "RecordInformation", Cardinality.ecmSingleValued, lobjRIObject, False)
      '  lobjProperty.Value = lobjDocumentObject
      '  lpUpdateArgs.Properties.Add(lobjProperty)
      '  lpSourceDocument.UpdateProperties(lpUpdateArgs)
      '  'TODO: Update the source document with this RecordInfo
      'Catch ex As Exception
      '  ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Throw ex
      'End Try


      Return lblnReturnValue

      ' TODO: Port, test and clean setDocumentProperties, then uncomment the block below
      ''If setDocumentProperties(lpSourceDocument, lpRecord) Then
      ''  Return True
      ''End If

      'File the Record in the Folder Passed
      'For setSecurityParent Update- Value of 1024 sets the SecurityParent property of the object 

      'Dim objDRCR
      'Try
      '  objDRCR = lpCreateRecordInFolder.File(lpRecord, IDM_CONTAINMENT_DEFINE_SECURITY_PARENTAGE)
      'Catch ex As Exception
      '  Throw New DeclarationException(String.Format("Unable to set the security parent for the record '{0}'.", _
      '                              lpRecord.Name), ex, ErrorType.UnableToSetTheSecurityParentForTheRecord)
      'End Try

      'lpRecord.Save()

      'If Err.Number <> 0 Then
      '  ' Delete the Created Record in case we are not able to save the properties
      '  'LogError "Deleting the record created above !" + Err.Description
      '  DeleteRecord(lpRecord)
      'Else

      '  If setDocumentProperties(lpSourceDocument, lpRecord) Then
      '    setRecordProperties = True
      '  End If

      'End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  ''' <summary>
  ''' Finds an existing location or creates a new location using the specified location name
  ''' </summary>
  ''' <param name="lpLocation">A Location object reference</param>
  ''' <returns>The Location custom object from the Content Engine Object Store</returns>
  ''' <remarks></remarks>
  Private Function GetCELocation(ByVal lpLocation As Location, Optional ByRef lpErrorMessage As String = "") As Object
    Try
      Dim lobjProvider As CENetProvider = Me.FilePlanContentSource.Provider
      'Dim lobjCEWebServices As New Utilities.CEWebServices(lobjProvider.URL, lobjProvider.UserName, lobjProvider.Password, lobjProvider.ObjectStore)
      Dim lobjCELocationObject As Object
      Dim lblnReturnValue As Boolean

      ' Check the provided Location parameter
      If lpLocation Is Nothing Then
        Throw New ArgumentNullException("lpLocation", "A valid value is required for the parameter lpLocation")
      Else
        If lpLocation.Name = String.Empty Then
          Throw New ArgumentException("The 'Name' property of the parameter 'lpLocation' is not set.  A valid name is required", "lpLocation")
        End If
      End If

      '  First we need to look for an existing location object matching the location parameter
      Dim lobjSearch As ISearch
      lobjSearch = New CENetSearch(CType(Me.FilePlanContentSource.Provider, CProvider))
      lobjSearch.DataSource.QueryTarget = "Location"
      Dim lobjCriterion As Data.Criterion

      '  Add the Name
      lobjCriterion = New Data.Criterion("LocationName", "LocationName", PropertyScope.BothDocumentAndVersionProperties, Data.Criterion.pmoOperator.opEquals, SetEvaluation.seAnd)
      With lobjCriterion
        .Value = lpLocation.Name
        .DataType = Data.Criterion.pmoDataType.ecmString
        .Cardinality = Cardinality.ecmSingleValued
      End With

      'lobjSearch.DataSource.Criteria.Add(lobjCriterion)
      lobjSearch.Criteria.Add(lobjCriterion)

      '  Add the Barcode
      lobjCriterion = New Data.Criterion("BarcodeID", "BarcodeID", PropertyScope.BothDocumentAndVersionProperties, Data.Criterion.pmoOperator.opEquals, SetEvaluation.seAnd)
      With lobjCriterion
        .Value = lpLocation.Barcode
        .DataType = Data.Criterion.pmoDataType.ecmString
        .Cardinality = Cardinality.ecmSingleValued
      End With
      'lobjSearch.DataSource.Criteria.Add(lobjCriterion)
      lobjSearch.Criteria.Add(lobjCriterion)

      ''  Add the Description
      'lobjCriterion = New Data.Criterion("Description", "Description", PropertyScope.BothDocumentAndVersionProperties, Data.Criterion.pmoOperator.opEquals, SetEvaluation.seAnd)
      'With lobjCriterion
      '  .Value = lpLocation.Description
      '  .DataType = Data.Criterion.pmoDataType.ecmString
      '  .Cardinality = Cardinality.ecmSingleValued
      'End With
      ''lobjSearch.DataSource.Criteria.Add(lobjCriterion)
      'lobjSearch.Criteria.Add(lobjCriterion)

      '  Add the Reviewer
      lobjCriterion = New Data.Criterion("Reviewer", "Reviewer", PropertyScope.BothDocumentAndVersionProperties, Data.Criterion.pmoOperator.opEquals, SetEvaluation.seAnd)
      With lobjCriterion
        .Value = lpLocation.Reviewer
        .DataType = Data.Criterion.pmoDataType.ecmString
        .Cardinality = Cardinality.ecmSingleValued
      End With
      'lobjSearch.DataSource.Criteria.Add(lobjCriterion)
      lobjSearch.Criteria.Add(lobjCriterion)

      Dim lobjSearchResultSet As SearchResultSet = lobjSearch.Execute

      If lobjSearchResultSet.HasException = False AndAlso lobjSearchResultSet.Results.Count > 0 Then
        '  We found one
        Dim lobjCELocationID As String = lobjSearchResultSet.Results(0).ID
        'lobjCELocationObject = CType(Me.FilePlanContentSource.Provider, CENetProvider).GetObject("Location", lobjCELocationID)
        lobjCELocationObject = CType(Me.FilePlanContentSource.Provider, CENetProvider).ObjectStore.FetchObject("Location", lobjCELocationID, Nothing)

        Return lobjCELocationObject

      Else
        '  We did not find one, we need to create it

        Dim lobjLocationCandidate As New Document
        lobjLocationCandidate.DocumentClass = "Location"

        Dim lobjFirstVersion As New Version()

        With lobjFirstVersion
          ' <Changed by Ernie Bahr - 8/15/2012>
          '.Properties.Add(New ECMProperty(PropertyType.ecmString, "LocationName", lpLocation.Name))
          '.Properties.Add(New ECMProperty(PropertyType.ecmString, "BarcodeID", lpLocation.Barcode))
          '.Properties.Add(New ECMProperty(PropertyType.ecmString, "RMEntityDescription", lpLocation.Description))
          '.Properties.Add(New ECMProperty(PropertyType.ecmString, "Reviewer", lpLocation.Reviewer))
          .Properties.Add(PropertyFactory.Create(PropertyType.ecmString, "LocationName", lpLocation.Name))
          .Properties.Add(PropertyFactory.Create(PropertyType.ecmString, "BarcodeID", lpLocation.Barcode))
          .Properties.Add(PropertyFactory.Create(PropertyType.ecmString, "RMEntityDescription", lpLocation.Description))
          .Properties.Add(PropertyFactory.Create(PropertyType.ecmString, "Reviewer", lpLocation.Reviewer))
          ' <End Changed by Ernie Bahr - 8/15/2012>
        End With

        lobjLocationCandidate.Versions.Add(lobjFirstVersion)
        lobjCELocationObject = CType(Me.FilePlanContentSource.Provider, CENetProvider).AddP8Document(lobjLocationCandidate, lpErrorMessage, VersionTypeEnum.Major)
        lblnReturnValue = True
        If lblnReturnValue = True Then
          Return lobjCELocationObject
        Else
          Return ""
        End If

      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  ''' <summary>
  ''' This method validates for the folder to a be a valid folder for declaring the record, 
  ''' depending on the business logic. E.x if the objecstore is a DOD based then the record can be devclared in the 
  ''' record category. If the record is to be filed in a Volume then chcek that the volume is active or not.
  ''' </summary>
  ''' <param name="lpCreateRecordInFolderPath"></param>
  ''' <param name="lpRecordRMType"></param>
  ''' <param name="lpImplementationType"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function getFolderWhereRecordWillBeFiled(ByVal lpCreateRecordInFolderPath As String,
                                                   ByVal lpRecordRMType As RecordType,
                                                   ByVal lpImplementationType As ImplementationType) As CFolder

    Try

      Dim lobjFolderPassedIn As CENetFolder

      ' Create a reference to the web service utility
      Dim lobjProvider As CENetProvider = Me.FilePlanContentSource.Provider
      'Dim lobjCEWebServices As New Utilities.CEWebServices(lobjProvider.URL, lobjProvider.UserName, lobjProvider.Password, lobjProvider.ObjectStore)

      '// Get the Folder Passed in by the script where it feels we should declare
      'Let See do we need to get volume or not
      '// We can improve the logic here by using AllowedRMtypes

      'LogError "inside getFolderWhereRecordWillBeFiled " + asCreateRecordInFolderPath + "object Store Name " + aoFPOSObjectStore.Name
      lobjFolderPassedIn = Me.FilePlanContentSource.Provider.GetFolder(lpCreateRecordInFolderPath, 0)

      'TEST
      'Dim lbReturn = lobjCEWebServices.IsDescendantOf(lobjFolderPassedIn.FolderClass, RM_TYPE_HYBRIDRECORDFOLDER_CLASS)


      Dim lstrAllowedTypes As String = String.Empty

      If checkAllowedRMTypes(lobjFolderPassedIn, lpRecordRMType, lstrAllowedTypes) Then
        'LogError "Got Handle to the Folder "
        If lobjFolderPassedIn Is Nothing Then
          Throw New DeclarationException(String.Format("Not able to find the Folder where record is to be created. Path '{0}'. FilePlan Objectstore '{1}'",
                                                       lpCreateRecordInFolderPath, FilePlanObjectStoreName), ErrorType.UnableToFindTargetRecordFolder)
        Else
          ' Record Categories will allow filing in case of DoD only
          'If lobjFolderPassedIn.FolderClass = RM_TYPE_RECORDCATEGORY_CLASS Then
          If lobjFolderPassedIn.ClassDescription.Contains(RM_TYPE_RECORDCATEGORY_CLASS) Then
            ' Only for the DoD implementation it will allow records in the Record Category
            ' Can improve the performance here
            'If (lpImplementationType = "DOD-5015.2" Or lpImplementationType = "Base" Or lpImplementationType = "DOD-5015.2CHAPTER4") Then  'DOD
            If (lpImplementationType = ImplementationType.DOD5015Point2 OrElse
                lpImplementationType = ImplementationType.Base OrElse
                lpImplementationType = ImplementationType.DOD5015Point2Chapter4) Then  'DOD
              '// Check if the Record Category is inactive or closed then record cannpot be filed in this

              If ((lobjFolderPassedIn.Properties("InActive").Value = True) OrElse
                  (lobjFolderPassedIn.Properties("DateClosed").Value <> "1/1/0001" AndAlso lobjFolderPassedIn.Properties("ReOpenedDate").Value = "1/1/0001")) Then
                Throw New DeclarationException(String.Format("Record can not be declared in the Record category at Path '{0}' in ObjectStore '{1}' because it is closed or inactive.",
                                                             lpCreateRecordInFolderPath, FilePlanObjectStoreName), ErrorType.RecordCategoryIsClosedOrInactive)
              End If

            Else
              Throw New DeclarationException(String.Format("Record can not be declared in the Record category in DoD FilePlan ObjectStore '{0}'",
                                                           FilePlanObjectStoreName), ErrorType.RecordMayNotBeDeclaredInTheRecordCategoryInDoDFilePlanObjectStore)
            End If

            'ElseIf lobjFolderPassedIn.FolderClass = RM_TYPE_VOLUME_CLASS Then
          ElseIf lobjFolderPassedIn.ClassDescription.Contains(RM_TYPE_VOLUME_CLASS) Then
            ' Check if the Volume is inactive or closed then record cannpot be filed in this
            If ((lobjFolderPassedIn.Properties("DateClosed").Value <> "1/1/0001" And lobjFolderPassedIn.Properties("ReOpenedDate").Value = "1/1/0001")) Then
              Throw New DeclarationException(String.Format("Record can not be declared in the Volume at Path '{0}' in ObjectStore '{1}' because it is closed.",
                                                           lpCreateRecordInFolderPath, FilePlanObjectStoreName), ErrorType.ContainerIsClosed)
            Else
              'Check in the Physical Volume only markers are allowed
            End If
            'ElseIf lobjFolderPassedIn.FolderClass = RM_TYPE_ELECTRONICRECORDFOLDER_CLASS Then
          ElseIf lobjFolderPassedIn.ClassDescription.Contains(RM_TYPE_ELECTRONICRECORDFOLDER_CLASS) Then
            If lpRecordRMType = RecordType.ElectronicRecordInfo Or
              lpRecordRMType = RecordType.EmailRecord Then
              '// Set the Folder to the active volume
              lobjFolderPassedIn = getActiveVolume(lobjFolderPassedIn)
            Else
              Throw New DeclarationException(String.Format("Record Can not be declared in the Electronic Record Folder '{0}'",
                                                           lpCreateRecordInFolderPath), ErrorType.InvalidRecordType_RecordTypeNotValidForElectronicRecordFolder)
            End If
            'ElseIf lobjFolderPassedIn.FolderClass = RM_TYPE_PHYSICALRECORDFOLDER_CLASS Then
          ElseIf lobjFolderPassedIn.ClassDescription.Contains(RM_TYPE_PHYSICALRECORDFOLDER_CLASS) Then
            If lpRecordRMType = RecordType.Marker Then
              '// Set the Folder to the active volume
              lobjFolderPassedIn = getActiveVolume(lobjFolderPassedIn)
            Else
              Throw New DeclarationException(String.Format("Record can not be declared in the Physical Record Folder '{0}'",
                                                           lobjFolderPassedIn.Name), ErrorType.InvalidRecordType_RecordTypeNotValidForPhysicalRecordFolder)
            End If
            'ElseIf lobjFolderPassedIn.FolderClass = RM_TYPE_HYBRIDRECORDFOLDER_CLASS Then
          ElseIf lobjFolderPassedIn.ClassDescription.Contains(RM_TYPE_HYBRIDRECORDFOLDER_CLASS) Then
            If lpRecordRMType = RecordType.ElectronicRecordInfo Or
                lpRecordRMType = RecordType.Marker Or
                lpRecordRMType = RecordType.EmailRecord Then
              lobjFolderPassedIn = getActiveVolume(lobjFolderPassedIn)
            Else
              Throw New DeclarationException(String.Format("Record Can not be declared in the Hybrid Record Folder '{0}'",
                                                           lobjFolderPassedIn.Name), ErrorType.InvalidRecordType_RecordTypeNotValidForHybridRecordFolder)
            End If
            'ElseIf lobjFolderPassedIn.FolderClass = RM_TYPE_PHYSICALCONTAINER_CLASS Then
          ElseIf lobjFolderPassedIn.ClassDescription.Contains(RM_TYPE_PHYSICALCONTAINER_CLASS) Then
            If lpRecordRMType = RecordType.Marker Then
              If ((lobjFolderPassedIn.Properties("DateClosed").Value <> "1/1/0001" And lobjFolderPassedIn.Properties("ReOpenedDate").Value = "1/1/0001")) Then
                Throw New DeclarationException(String.Format("Record can not be declared in the Box at Path '{0}' in ObjectStore '{1}' because it is closed.",
                                             lpCreateRecordInFolderPath, FilePlanObjectStoreName), ErrorType.ContainerIsClosed)
              End If
              '// Record cannot be declared if a container has another container
              Dim SubFolders
              For Each SubFolders In lobjFolderPassedIn.SubFolders
                Throw New DeclarationException(String.Format("Record Can not be declared in the Physical Container '{0}'.",
                             lobjFolderPassedIn.Name), ErrorType.InvalidRecordType_RecordTypeNotValidForPhysicalContainer)
                Exit For
              Next
            Else
              Throw New DeclarationException(String.Format("Only Marker type record can declare in the Physical Container '{0}'.",
                           lobjFolderPassedIn.Name), ErrorType.InvalidRecordType_RecordTypeNotAllowedForContainer)
            End If
          Else
            lobjFolderPassedIn = Nothing
          End If
        End If
      Else
        If lpRecordRMType = RecordType.Marker Then
          Throw New DeclarationException(String.Format("The RM Container '{0}' does not allow the declaration of Marker Record!  The only allowed types are '{1}'",
                                                       lobjFolderPassedIn.Name, lstrAllowedTypes), ErrorType.InvalidRecordType_RecordTypeNotAllowedForContainer)
        ElseIf lpRecordRMType = RecordType.ElectronicRecordInfo Then
          Throw New DeclarationException(String.Format("The RM Container '{0}'  does not allow the declaration of Electronic Record!  The only allowed types are '{1}'",
                                                       lobjFolderPassedIn.Name, lstrAllowedTypes), ErrorType.InvalidRecordType_RecordTypeNotAllowedForContainer)
        ElseIf lpRecordRMType = RecordType.EmailRecord Then
          Throw New DeclarationException(String.Format("The RM Container '{0}'  does not allow the declaration of Email Record!  The only allowed types are '{1}'",
                                                       lobjFolderPassedIn.Name, lstrAllowedTypes), ErrorType.InvalidRecordType_RecordTypeNotAllowedForContainer)
        End If
        lobjFolderPassedIn = Nothing
      End If

      Return lobjFolderPassedIn

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function


  ''' <summary>
  ''' This method checks whether the volume in which the record is going to be declared is active or not, 
  ''' by checking that the date closed and re open date is not NULL.
  ''' </summary>
  ''' <param name="lpRecordFolder"></param>
  ''' <returns></returns>
  ''' <remarks>The current active volume for the record folder.</remarks>
  Private Function getActiveVolume(ByVal lpRecordFolder As CFolder) As CFolder

    Try

      Dim lsDateVolumeCreated As DateTime = DateTime.MinValue
      'lsDateVolumeCreated = lpRecordFolder.Properties("DateCreated").Value
      Dim loLastCreatedVolume As CFolder = Nothing
      For Each loVolumeInFolder As CFolder In lpRecordFolder.SubFolders
        If loVolumeInFolder.Properties("DateCreated").Value > lsDateVolumeCreated Then
          loLastCreatedVolume = loVolumeInFolder
          lsDateVolumeCreated = loVolumeInFolder.Properties("DateCreated").Value
        End If
      Next

      If ((loLastCreatedVolume.Properties("DateClosed").Value <> "1/1/0001" And loLastCreatedVolume.Properties("ReOpenedDate").Value = "1/1/0001")) Then
        Throw New DeclarationException(String.Format("Record can not be declared in the Volume at Path '{0}' because it is not active. It is closed.",
                                             loLastCreatedVolume.Name), ErrorType.ContainerIsClosed)
        loLastCreatedVolume = Nothing
      End If

      getActiveVolume = loLastCreatedVolume
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try

  End Function


  ''' <summary>
  ''' 
  ''' </summary>
  ''' <param name="lpCreateRecordInFolder"></param>
  ''' <param name="lpRecordRMType"></param>
  ''' <param name="lpAllowedTypes"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function checkAllowedRMTypes(ByVal lpCreateRecordInFolder As CFolder,
                                       ByVal lpRecordRMType As RecordType,
                                       Optional ByRef lpAllowedTypes As String = "") As Boolean

    Try

      'LogError "Script is in checkAllowedRMTypes"
      Dim lblnFilingAllowedHere As Boolean
      Dim lenuAllowedValue As RecordType

      ' By Default function assumes you cannot file here. But Still it will give it a try !
      lblnFilingAllowedHere = False
      ' Get the values of the allowed RM Types and loop through all the values till the 
      ' match for passed in value is find.
      Dim lsAllowedRMTypesValues As Core.Values
      lsAllowedRMTypesValues = lpCreateRecordInFolder.Properties("AllowedRMTypes").Values
      'LogError "Number of Values in the Category " & Err.Description & "Count " & lsAllowedRMTypesValues.Count
      For Each lobjValue As Object In lsAllowedRMTypesValues
        lenuAllowedValue = lobjValue
        lpAllowedTypes &= String.Format("{0}, ", lenuAllowedValue.ToString)
        If lenuAllowedValue = lpRecordRMType Then
          lblnFilingAllowedHere = True
          Exit For
        End If
      Next

      If lpAllowedTypes.EndsWith(", ") Then
        lpAllowedTypes = lpAllowedTypes.Remove(lpAllowedTypes.Length - 2)
      End If

      If lsAllowedRMTypesValues.Count = 2 Then
        lpAllowedTypes = lpAllowedTypes.Replace(", ", " & ")
      End If

      Return lblnFilingAllowedHere
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  ''' <summary>
  ''' This method get the Implementation Type either DOD or PRO.  
  ''' It searches for the System Configuration object "FPOS Setup" 
  ''' and checks the value of the metadata "PropertyValue"
  ''' </summary>
  ''' <returns>Implementation Type</returns>
  ''' <remarks>Need to implement later...
  ''' For Implementation (whether DOD or PRO).  
  ''' User need to set this properties if the implementation is other than DOD
  ''' </remarks>
  Public Function getRMImplementationType() As ImplementationType '(ByVal lpFilePlanObjectStore As Object) As String
    Try
      Dim lstrQueryResult As String
      'Dim lenuImplementationType As ImplementationType
      lstrQueryResult = getSingleQueryResult("SystemConfiguration", "propertyname", "FPOS Setup", "propertyvalue")
      Select Case lstrQueryResult
        Case "Base"
          Return ImplementationType.Base
        Case "Pro"
          Return ImplementationType.Pro
        Case "DOD-5015.2", "DOD"
          Return ImplementationType.DOD5015Point2
        Case "DOD-5015.2CHAPTER4"
          Return ImplementationType.DOD5015Point2Chapter4
      End Select
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  ''' <summary>
  ''' 
  ''' </summary>
  ''' <param name="lpQueryTarget"></param>
  ''' <param name="lpCriteria"></param>
  ''' <param name="lpReturnColumnName"></param>
  ''' <param name="lpIncludeIdInResults"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function getSingleQueryResult(ByVal lpQueryTarget As String, ByVal lpCriteria As Data.Criteria,
                                        ByVal lpReturnColumnName As String,
                                        Optional ByVal lpIncludeIdInResults As Boolean = False) As String
    Try
      Dim lobjSearch As ISearch
      lobjSearch = New CENetSearch(CType(Me.FilePlanContentSource.Provider, CProvider))

      lobjSearch.DataSource.QueryTarget = lpQueryTarget


      lobjSearch.Criteria = lpCriteria

      If lpReturnColumnName.ToLower <> "id" Then
        lobjSearch.DataSource.ResultColumns.Add(lpReturnColumnName)
      End If

      If lpIncludeIdInResults = False Then
        lobjSearch.DataSource.ResultColumns.Remove("Id")
      End If

      Dim srs As SearchResultSet = lobjSearch.Execute()

      If srs.HasException = True Then
        Throw New System.Data.DataException(String.Format("The '{0}' could not be determined in getSingleQueryResult.", lpReturnColumnName), srs.Exception)
      End If

      If srs.Count > 0 Then
        Dim lobjDataTable As System.Data.DataTable = srs.ToDataTable
        Return lobjDataTable.Rows(0).Item(0).ToString
      Else
        Throw New System.Data.DataException(String.Format("The '{0}' could not be determined.", lpReturnColumnName))
      End If

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod, 0,
                          String.Format("lpQueryTarget:{0}", lpQueryTarget),
                          String.Format("lpCriteria:{0}", lpCriteria.ToString),
                          String.Format("lpReturnColumnName:{0}", lpReturnColumnName))
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function

  ''' <summary>
  ''' 
  ''' </summary>
  ''' <param name="lpQueryTarget"></param>
  ''' <param name="lpCriterionName"></param>
  ''' <param name="lpCriterionValue"></param>
  ''' <param name="lpReturnColumnName"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function getSingleQueryResult(ByVal lpQueryTarget As String,
                                        ByVal lpCriterionName As String,
                                        ByVal lpCriterionValue As String,
                                        ByVal lpReturnColumnName As String) As String
    Try

      Dim lobjCriterion As Data.Criterion = New Data.Criterion(lpCriterionName)
      lobjCriterion.Value = lpCriterionValue
      lobjCriterion.DataType = Data.Criterion.pmoDataType.ecmString

      Dim lobjCriteria As New Data.Criteria

      lobjCriteria.Add(lobjCriterion)

      Return getSingleQueryResult(lpQueryTarget, lobjCriteria, lpReturnColumnName)

      'getRMImplementationType = getSingleQueryResult(lpFilePlanObjectStore, "SELECT propertyvalue FROM SystemConfiguration Where propertyname = 'FPOS Setup'")

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Function


  Private Function IsDescendantOf(ByVal lstrFolderName As String, ByVal lstrAnscenstorFolderName As String) As Boolean

    Dim lbReturn As Boolean = False
    Dim lobjClassInfo As Object
    Try
      If (mobjFolderClasses.Count = 0) Then
        'Build the collection
        lobjClassInfo = GetClassInfo("{01A3A8CA-7AEC-11D1-A31B-0020AF9FBB1C}") 'Root Folder Id
        PopulateFolderDefinitions(lobjClassInfo, "{01A3A8CA-7AEC-11D1-A31B-0020AF9FBB1C}", "{01A3A8CA-7AEC-11D1-A31B-0020AF9FBB1C}")
      End If

      lbReturn = mobjFolderClasses.IsDescendantOf(lstrFolderName, lstrAnscenstorFolderName)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try

    Return lbReturn

  End Function


  ''' <summary>
  ''' Populates the collection mobjFolderClasses
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub PopulateFolderDefinitions(ByVal lpFolderInfo As Object, ByVal lpParentId As String, ByVal lpId As String)

    Dim lobjHierarchyItem As HierarchyItem = New HierarchyItem
    Dim lobjEcmProperty As ECMProperty
    Dim lobjObjectValue As Object
    Dim lstrId As String

    Try

      For Each lobjProperty As FileNet.Api.Property.IProperty In lpFolderInfo.Property
        lobjEcmProperty = CType(Me.mobjFPOSContentSource.Provider, CENetProvider).CreateECMProperty(lobjProperty)
        Select Case lobjEcmProperty.Name
          Case "Id"
            lobjHierarchyItem.Id = lobjEcmProperty.Value
            lobjHierarchyItem.ParentId = lpParentId
            'Case "PrimaryId"
            '  lobjHierarchyItem.ParentId = lobjEcmProperty.Value
          Case "SymbolicName"
            lobjHierarchyItem.Name = lobjEcmProperty.Value.ToString
          Case "ImmediateSubclassDefinitions" 'Big assumption, properties always come back in order Id, Name, ImmediateSubClassDefinitions
            For Each lobjValue As Object In lobjEcmProperty.Values
              lstrId = lobjValue.Value.ToString
              lobjObjectValue = GetClassInfo(lstrId)
              mobjFolderClasses.Add(lobjHierarchyItem)
              PopulateFolderDefinitions(lobjObjectValue, lobjHierarchyItem.Id, lstrId)
            Next
        End Select
      Next

      mobjFolderClasses.Add(lobjHierarchyItem)
    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      '  Re-throw the exception to the caller
      Throw
    End Try


  End Sub

  Private Function GetClassInfo(ByVal lpId As String) As Object

    Try

      Return Nothing

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod, lpId)
      '  Re-throw the exception to the caller
      Throw
    End Try

  End Function


#End Region

#Region "Public Methods"

  Public Sub AddPhysicalRecord(ByRef Args As Records.AddPhysicalRecordArgs)
    Try

      ' Check to make sure all the required argument values are present
      CheckRequiredArgs(Args)

      Dim lstrErrorMessage As String = String.Empty
      Dim lobjCreateRecordInFolder As CFolder

      If Me.FilePlanContentSource Is Nothing Then
        Me.FilePlanContentSource = Args.FilePlanContentSource
      End If

      ' Determine the Implementation Type
      Dim lsImplementationType As ImplementationType
      'LogError "Before getting the Implementation Type 
      lsImplementationType = getRMImplementationType() '(lstrFilePlanObjectStoreName)

      lobjCreateRecordInFolder = getFolderWhereRecordWillBeFiled(Args.RecordFolderPath, Args.RecordType, lsImplementationType)

      ' Its time to create the record!
      Dim lobjNewRecord As PhysicalRecord

      lobjNewRecord = createRecordInfo(FilePlanObjectStoreName, "",
                                          Args.RecordType, Args.RecordClass,
                                          lobjCreateRecordInFolder, Args.SourceDocument, Args.RecordProperties, lsImplementationType)

      Args.NewRecord = lobjNewRecord

    Catch ex As Exception
      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      ' Re-throw the exception to the caller
      Throw
    End Try
  End Sub

#End Region

End Class



''' <summary>
''' Represents an item in a Hierarchy
''' </summary>
''' <remarks></remarks>
Public Class HierarchyItem

#Region "Class Variables"
  Dim mstrParentId As String
  Dim mstrId As String
  Dim mstrName As String
#End Region

#Region "Public Properties"

  Public Property ParentId() As String
    Get
      Return mstrParentId
    End Get
    Set(ByVal value As String)
      mstrParentId = value
    End Set
  End Property

  Public Property Id() As String
    Get
      Return mstrId
    End Get
    Set(ByVal value As String)
      mstrId = value
    End Set
  End Property

  Public Property Name() As String
    Get
      Return mstrName
    End Get
    Set(ByVal value As String)
      mstrName = value
    End Set
  End Property

#End Region

#Region "Constructors"

  Public Sub New()

  End Sub

  Public Sub New(ByVal Id As String, ByVal ParentId As String, ByVal Name As String)
    mstrId = Id
    mstrParentId = ParentId
    mstrName = Name
  End Sub

#End Region

End Class

''' <summary>
''' A Collection of HiearchyItems
''' </summary>
''' <remarks></remarks>
Public Class HierarchyItemCollection
  Inherits CollectionBase

#Region "Public Propreties"

  Default Public Property Item(ByVal index As Integer) As HierarchyItem
    Get
      Return CType(List(index), HierarchyItem)
    End Get
    Set(ByVal value As HierarchyItem)
      List(index) = value
    End Set
  End Property

  Default Public Property Item(ByVal name As String) As HierarchyItem
    Get
      For Each hitem As HierarchyItem In List
        If hitem.Name = name Then
          Return hitem
        End If
      Next
      Return Nothing
    End Get
    Set(ByVal value As HierarchyItem)
      For Each hitem As HierarchyItem In List
        If hitem.Name = name Then
          List(List.IndexOf(hitem)) = value
        End If
      Next
    End Set
  End Property

#End Region

#Region "Public Functions"

  Public Function IsDescendantOf(ByVal FolderName As String, ByVal AnscenstorName As String) As Boolean

    If (FolderName = AnscenstorName) Then
      Return True
    End If

    ' We've reached the top root and no match so bail out
    If (FolderName = "Folder") Then
      Return False
    End If

    Dim lobjHierarchyItem As HierarchyItem = Me.Item(FolderName)
    Dim lobjHierCheck As HierarchyItem

    If (lobjHierarchyItem Is Nothing) Then
      Return False
    End If

    'Check Parents
    For i As Integer = 0 To Me.List.Count - 1
      lobjHierCheck = Me.List(i)
      If (lobjHierCheck.Id = lobjHierarchyItem.ParentId) Then
        Return IsDescendantOf(lobjHierCheck.Name, AnscenstorName)
      End If
    Next

    Return False

  End Function

  Public Function Add(ByVal value As HierarchyItem) As Integer
    Return List.Add(value)
  End Function 'Add

  Public Function IndexOf(ByVal value As HierarchyItem) As Integer
    Return List.IndexOf(value)
  End Function 'IndexOf


  Public Sub Insert(ByVal index As Integer, ByVal value As HierarchyItem)
    List.Insert(index, value)
  End Sub 'Insert

  Public Shadows Function Count() As Integer
    Return List.Count
  End Function

  Public Sub Remove(ByVal value As HierarchyItem)
    List.Remove(value)
  End Sub 'Remove


  Public Function Contains(ByVal value As HierarchyItem) As Boolean
    ' If value is not of type Int16, this will return false.
    Return List.Contains(value)
  End Function 'Contains


  Protected Overrides Sub OnInsert(ByVal index As Integer, ByVal value As Object)
    ' Insert additional code to be run only when inserting values.
  End Sub 'OnInsert


  Protected Overrides Sub OnRemove(ByVal index As Integer, ByVal value As Object)
    ' Insert additional code to be run only when removing values.
  End Sub 'OnRemove


  Protected Overrides Sub OnSet(ByVal index As Integer, ByVal oldValue As Object, ByVal newValue As Object)
    ' Insert additional code to be run only when setting values.
  End Sub 'OnSet


  Protected Overrides Sub OnValidate(ByVal value As Object)
    If Not GetType(HierarchyItem).IsAssignableFrom(value.GetType()) Then
      Throw New ArgumentException("value must be of type HierarchyItem.", "value")
    End If
  End Sub 'OnValidate 

#End Region

End Class

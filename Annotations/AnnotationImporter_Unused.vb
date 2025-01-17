'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "UNUSED: Public Class FileNetAnnotationHelper"

' ''' <summary>
' ''' Assists with the creation and content update of native P8 annotations.
' ''' The code is in this class is mostly commented-out because it was written (and tested) against the .NET API, whose P8 object wrappers are not exposed in the web services API.
' ''' </summary>
'Public Class FileNetAnnotationHelper

'  ' Main execution block sample for getting a Daeja-style XML representation of the CTS Annotation model and loading it into P8
'  ' This method uses the .NET API, and so needs to be translated to the web services API
'  Sub LoadAnnotation_SampleCode(ByVal lpObjCEWebServices As Utilities.CEWebServices)

'    ' * INTERCHANGE 1 - QUERY FOR DOMAIN

'    'Dim ce As New CEConnection()
'    'ce.EstablishCredidentials("p8admin", "sword", "http://vm01:9080/wsi/FNCEWS40MTOM/")
'    Dim wseService As FNCEWS35ServiceWse = Me.ObtainWebServiceObject("http://vm01:9080/wsi/FNCEWS40MTOM/")
'    Dim soapContextObject As SoapContext = Me.ObtainSoapContext(wseService, "p8admin", "sword")

'    ' * INTERCHANGE 2 - GET OBJECT STORES LIST
'    ' * INTERCHANGE 3 - GET "XP" OBJECT STORE REFERENCE

'    'Dim os = ce.FetchObjectStore("XP")

'    ' * INTERCHANGE 4 - GET DOCUMENT REFERENCE

'    '' Create property filter for document's content elements, 
'    '' which are needed to get element sequence numbers that identify elements.
'    'Dim pf As PropertyFilter = New PropertyFilter()
'    'pf.AddIncludeProperty(New FilterElement(Nothing, Nothing, Nothing, PropertyNames.CONTENT_ELEMENTS, Nothing))

'    ''Fetch Document object.
'    'Dim doc As IDocument = Factory.Document.FetchInstance(os, "/Annotations/Test 1", Nothing)
'    Dim objectDocument As ObjectValue = Me.FetchInstance(wseService, "XP", "Document", "/Annotations/Test 1")

'    ' * INTERCHANGE 5 - GET CONTENT ELEMENTS
'    ' * INTERCHANGE 6  - CREATE ANNOTATION

'    'Dim nativeAnnotation As IAnnotation = CreateAnnotation(os, doc, 0)
'    Dim objectAnnotation As ObjectValue = Me.CreateAnnotation(wseService, objectDocument, 0)

'    'GenerateAnnotationContent("{5CF11946-018F-11D0-A87A-00A0246922A5}", nativeAnnotation.Id.ToString(), "C:\\annotation.xml")
'    Me.GenerateAnnotationContent("{5CF11946-018F-11D0-A87A-00A0246922A5}", objectAnnotation.Property("Id").ToString(), "C:\\annotation.xml")

'    ' * INTERCHANGE 7  - POPULATE ANNOTATION CONTENT

'    'PopulateAnnotationContent(nativeAnnotation, "C:\\annotation.xml")
'    Me.PopulateAnnotationContent(wseService, objectAnnotation, "C:\\annotation.xml", "text/xml")

'  End Sub

'  ' This method uses the .NET API, and so needs to be translated to the web services API
'  'Public Function CreateAnnotation(ByVal os As IObjectStore, ByVal doc As IDocument, ByVal contentElementIndex As Integer) As IAnnotation
'  '  Return CreateAnnotation(os, doc, contentElementIndex, "Annotation")
'  'End Function

'  ' This method uses the .NET API, and so needs to be translated to the web services API
'  'Public Function CreateAnnotation(ByVal os As IObjectStore, ByVal doc As IDocument, ByVal contentElementIndex As Integer, ByVal annotationClassName As String) As IAnnotation
'  '  Return CreateAnnotation(os, doc, contentElementIndex, annotationClassName, String.Empty)
'  'End Function

'  ' Creates a native P8 Annotation object as an empty object, with appropriate reference to annotated content id and content element index.
'  ' This is necessary in order to get the id and object reference of the annotation object itself.  
'  ' The annotation's id has to appear in the XML content, so we have to create the object first, then update it.
'  ' This method uses the .NET API, and so needs to be translated to the web services API
'  'Public Function CreateAnnotation(ByVal os As IObjectStore, ByVal doc As IDocument, ByVal contentElementIndex As Integer, ByVal annotationClassName As String, ByVal description As String) As IAnnotation
'  '  'Create the annotation.
'  '  Dim annObject As IAnnotation = Factory.Annotation.CreateInstance(os, annotationClassName)

'  '  'Set the Document object to which the annotation applies on the annotation.
'  '  annObject.AnnotatedObject = doc

'  '  ' Specify the document's ContentElement to which the annotation applies.
'  '  ' The ContentElement is identified by its element sequence number.
'  '  Dim docContentList As IContentElementList = doc.ContentElements
'  '  Dim contentElement As IContentElement = docContentList(contentElementIndex)
'  '  Dim elementSequenceNumber As Integer = contentElement.ElementSequenceNumber.Value
'  '  annObject.AnnotatedContentElement = elementSequenceNumber

'  '  ' Set annotation's DescriptiveText property.
'  '  annObject.DescriptiveText = description

'  '  annObject.Save(RefreshMode.REFRESH)

'  '  Return annObject
'  'End Function

'  ' Populates an existing native P8 annotation object with the contents of a file.
'  ' This method uses the .NET API, and so needs to be translated to the web services API
'  'Public Sub PopulateAnnotationContent(ByVal nativeAnnotation As IAnnotation, ByVal filePath As String)
'  '  ' Create Stream object with annotation content.
'  '  Dim fileStream As Stream = File.OpenRead(filePath)

'  '  ' Create ContentTransfer and ContentElementList objects for the annotation.
'  '  Dim ctObject As IContentTransfer = Factory.ContentTransfer.CreateInstance()
'  '  Dim annContentList As IContentElementList = Factory.ContentTransfer.CreateList()
'  '  ctObject.SetCaptureSource(fileStream)
'  '  ctObject.RetrievalName = "annotation.xml"
'  '  ctObject.ContentType = "text/xml"

'  '  ' Add ContentTransfer object to list and set the list on the annotation.
'  '  annContentList.Add(ctObject)
'  '  nativeAnnotation.ContentElements = annContentList

'  '  nativeAnnotation.Save(RefreshMode.REFRESH)

'  'End Sub

'  ' Stub to serve the same purpose as AnnotationImporter.
'  Public Function GenerateAnnotationXML(ByVal annotationId As String, ByVal annotationClassId As String) As String
'    ' <FnAnno>
'    Dim xml As New StringBuilder()
'    xml.Append("<FnAnno>")
'    ' <PropDesc F_ANNOTATEDID="{D94DE8C2-783E-4F30-9EE5-41C907FFB5BE}" 
'    xml.Append("<PropDesc F_ANNOTATEDID=""")
'    xml.Append(annotationId)
'    xml.Append(Chr(34))

'    ' F_ENTRYDATE="2010-09-03T19:27:52.0000000-05:00" F_HEIGHT="0.5684211" 
'    xml.Append(" F_ENTRYDATE=""2010-09-03T19:27:52.0000000-05:00"" F_HEIGHT=""0.75""")

'    ' F_ID="{D94DE8C2-783E-4F30-9EE5-41C907FFB5BE}"
'    xml.Append(" F_ID=""")
'    xml.Append(annotationId)
'    xml.Append(Chr(34))

'    ' F_LEFT="0.4526316" F_MODIFYDATE="2010-09-03T19:27:47.0000000-05:00" F_PAGENUMBER="1" F_MULTIPAGETIFFPAGENUMBER="0" 
'    xml.Append(" F_LEFT=""0.4375"" F_MODIFYDATE=""2010-09-03T19:27:47.0000000-05:00"" F_PAGENUMBER=""1"" F_MULTIPAGETIFFPAGENUMBER=""0""")

'    ' F_NAME="-1-{D94DE8C2-783E-4F30-9EE5-41C907FFB5BE}" 
'    xml.Append(" F_NAME=""-1-")
'    xml.Append(annotationId)
'    xml.Append(Chr(34))

'    ' F_TOP="0.3894737" F_WIDTH="0.6842105" F_CLASSNAME="Highlight" .
'    xml.Append(" F_TOP=""0.3854166666666667"" F_WIDTH=""1.4583333333333333"" F_CLASSNAME=""Highlight""")

'    ' F_CLASSID="{5CF11942-018F-11D0-A87A-00A0246922A5}" 
'    xml.Append(" F_CLASSID=""{5CF11942-018F-11D0-A87A-00A0246922A5}""")

'    'xml.Append(""" F_ARROWHEAD_SIZE=""1"" F_CLASSID=""")
'    'xml.Append(annotationClassId)
'    'xml.Append(""" F_CLASSNAME=""Arrow"" F_ENTRYDATE=""2010-09-03T09:27:43.0000000-05:00"" F_HEIGHT=""0.5520833333333334"" F_ID=""")
'    'xml.Append(annotationId)
'    'xml.Append(""" F_LEFT=""1.3854166666666667"" F_LINE_BACKMODE=""2"" F_LINE_COLOR=""255"" F_LINE_END_X=""1.3854166666666667"" F_LINE_END_Y=""1.3958333333333333"" F_LINE_START_X=""1.3854166666666667"" F_LINE_START_Y=""0.84375"" F_LINE_STYLE=""0"" F_LINE_WIDTH=""3"" F_MODIFYDATE=""2010-09-03T09:27:43.0000000-05:00"" F_MULTIPAGETIFFPAGENUMBER=""0"" F_NAME=""-1-")
'    'xml.Append(annotationId)
'    ' xml.Append(""" F_PAGENUMBER=""1"" F_TOP=""0.84375"" F_WIDTH=""0.03125"">")

'    ' F_TEXT_BACKMODE="1" F_LINE_COLOR="65535" F_LINE_WIDTH="0" F_BRUSHCOLOR="65535">
'    xml.Append(" F_TEXT_BACKMODE=""1"" F_LINE_COLOR=""16711680"" F_LINE_WIDTH=""0"" F_BRUSHCOLOR=""65535"">")

'    ' <F_CUSTOM_BYTES /><F_POINTS /><F_TEXT /></PropDesc></FnAnno>
'    xml.Append("<F_CUSTOM_BYTES /><F_POINTS /><F_TEXT /></PropDesc></FnAnno>")
'    Dim result As String = xml.ToString()
'    Return result
'  End Function

'#Region "Web Services helpers"

'  Private Function ObtainWebServiceObject(ByVal wsUrl As String) As FNCEWS35ServiceWse
'    Dim wseService As New FNCEWS35ServiceWse() With {
'      .Url = wsUrl
'    }
'    Return wseService
'  End Function

'  Private Function ObtainSoapContext(ByVal wseService As FNCEWS35ServiceWse, ByVal userName As String, ByVal password As String) As SoapContext
'    ' Create a wse-enabled web service object to provide access to SOAP header
'    Dim soapContextObject As SoapContext = wseService.RequestSoapContext

'    ' Add security token to SOAP header with your username and password
'    Dim token As New UsernameToken(userName, password, PasswordOption.SendPlainText)
'    soapContextObject.Security.Tokens.Add(token)

'    ' Add default locale info to SOAP header
'    Dim defaultLocale = New Localization()
'    defaultLocale.Locale = "en-US"

'    Return soapContextObject
'  End Function

'  Private Function CreateAnnotation(ByVal wseService As FNCEWS35ServiceWse, ByVal annotatedObject As Object, ByVal annotatedContentElementIndex As Integer) As ObjectValue
'    ' Build the create action for an annotation
'    Dim verbCreate As New CreateAction() With {.classId = "Annotation"}

'    ' Assign the actions to the ChangeRequestType element
'    Dim elemChangeRequestType As New ChangeRequestType() With {.Action = New ActionType(1) {}}
'    elemChangeRequestType.Action(0) = DirectCast(verbCreate, ActionType)

'    ' Build a list of properties to set in the annotation
'    Dim elemInputProps As ModifiablePropertyType() = New ModifiablePropertyType(3) {}

'    ' Specify and set a string-valued property for the DescriptiveText property
'    Dim propDocumentTitle As New SingletonString() With {.Value = String.Empty, .propertyId = "DescriptiveText"}
'    elemInputProps(0) = propDocumentTitle

'    ' Specify and set a string-valued property for the AnnotatedObject property
'    Dim propAnnotatedObject As New SingletonObject() With {
'      .propertyId = "AnnotatedObject",
'      .Value = annotatedObject
'    }
'    elemInputProps(1) = propAnnotatedObject

'    ' Specify and set a string-valued property for the AnnotatedContentElement property
'    Dim propAnnotatedContentElement As New SingletonInteger32() With {
'      .propertyId = "AnnotatedContentElement",
'      .Value = annotatedContentElementIndex
'    }
'    elemInputProps(2) = propAnnotatedContentElement

'    ' Create array of ChangeRequestType elements and assign ChangeRequestType element to it
'    Dim elemChangeRequestTypeArray As ChangeRequestType() = New ChangeRequestType(0) {}
'    elemChangeRequestTypeArray(0) = elemChangeRequestType

'    ' Create ChangeResponseType element array
'    Dim response As ChangeResponseType()

'    ' Build ExecuteChangesRequest element and assign ChangeRequestType element array to it
'    Dim elemExecuteChangesRequest As New ExecuteChangesRequest() With {
'      .ChangeRequest = elemChangeRequestTypeArray,
'      .refresh = True,
'      .refreshSpecified = True
'    }

'    Dim objResponse As ObjectValue = Nothing
'    Try
'      ' Call ExecuteChanges operation to implement the doc checkout
'      response = wseService.ExecuteChanges(elemExecuteChangesRequest)
'    Catch ex As System.Net.WebException
'      Throw ex
'    Finally
'      ' The new document object should be returned, unless there is an error
'      If response Is Nothing OrElse response.Length < 1 Then
'        Console.WriteLine("A valid object was not returned from the ExecuteChanges operation")
'      Else
'        ' process response
'        objResponse = DirectCast(response(0), ChangeResponseType)
'      End If
'    End Try

'    Return objResponse

'  End Function

'  Private Function FetchInstance(ByVal wseService As FNCEWS35ServiceWse, ByVal objectStore As String, ByVal className As String, ByVal objectPath As String) As ObjectValue
'    ' Set a reference to the document
'    Dim objDocumentSpec As New ObjectSpecification() With {
'      .classId = className,
'      .path = objectPath,
'      .objectStore = objectStore
'    }

'    ' Create a property filter to get Containers property
'    'Dim elemPropFilter As New PropertyFilterType()
'    'elemPropFilter.maxRecursion = 1
'    'elemPropFilter.maxRecursionSpecified = True
'    'elemPropFilter.IncludeProperties = New FilterElementType(0) {}
'    'elemPropFilter.IncludeProperties(0) = New FilterElementType()
'    'elemPropFilter.IncludeProperties(0).Value = "Containers"

'    ' Create the request for GetObjects
'    Dim request As ObjectRequestType() = New ObjectRequestType(0) {}
'    request(0) = New ObjectRequestType()
'    request(0).SourceSpecification = objDocumentSpec
'    ' request(0).PropertyFilter = elemPropFilter

'    'Dim verbRequestObject As New ObjectRequestType() With {
'    '  .SourceSpecification = objDocument
'    '}

'    Dim response As ObjectResponseType()
'    Try
'      response = wseService.GetObjects(request)
'    Catch ex As System.Net.WebException
'      Throw ex
'    End Try

'    Dim objResponse As ObjectValue = Nothing
'    If TypeOf response(0) Is SingleObjectResponse Then
'      objResponse = DirectCast(response(0), SingleObjectResponse).[Object]
'    End If

'    Return objResponse
'  End Function

'  Private Function PopulateAnnotationContent(ByVal wseService As FNCEWS35ServiceWse, ByVal annotationObject As ObjectValue, ByVal filePath As String, ByVal fileMimeType As String) As ObjectValue

'    ' Create update action
'    Dim annotationUpdateAction As New UpdateAction()

'    ' Assign the action to the ChangeRequestType element
'    Dim elemChangeRequestTypeArray As ChangeRequestType() = New ChangeRequestType(0) {}
'    Dim elemChangeRequestType As New ChangeRequestType()
'    elemChangeRequestTypeArray(0) = elemChangeRequestType

'    ' Create ChangeResponseType element array 
'    Dim elemChangeResponseTypeArray As ChangeResponseType()

'    ' Build ExecuteChangesRequest element and assign ChangeRequestType element array to it
'    Dim elemExecuteChangesRequest As New ExecuteChangesRequest()
'    elemExecuteChangesRequest.ChangeRequest = elemChangeRequestTypeArray
'    elemExecuteChangesRequest.refresh = True
'    ' return a refreshed object
'    elemExecuteChangesRequest.refreshSpecified = True

'    elemChangeRequestType.Action = New ActionType(0) {}
'    elemChangeRequestType.Action(0) = DirectCast(annotationUpdateAction, ActionType)

'    ' Specify the target object (Reservation object) for the actions
'    elemChangeRequestType.TargetSpecification = New ObjectReference() With {
'      .classId = "Annotation",
'      .objectId = annotationObject.objectId,
'      .objectStore = annotationObject.objectStore
'    }
'    elemChangeRequestType.id = "1"

'    ' Assign ChangeRequestType element
'    elemChangeRequestTypeArray(0) = elemChangeRequestType

'    ' Create an object reference to dependently persistable ContentTransfer object
'    Dim lobjContentTransfer As DependentObjectType = Factory.NewDependentObjectType("ContentTransfer", _
'      DependentObjectTypeDependentAction.Insert, True)
'    lobjContentTransfer.Property = New CEWSI35.PropertyType(1) {}

'    ' Create reference to the object set of ContentTransfer objeCts returned by the Document.ContentElements property
'    Dim lobjContentElements As ListOfObject = Factory.NewListOfObject("ContentElements")
'    lobjContentElements.Value = New DependentObjectType(1) {}
'    lobjContentElements.Value(0) = lobjContentTransfer

'    ' Create DIME attachment and add to SOAP header
'    Dim lobjDimeAttachment As DimeAttachment = New DimeAttachment(fileMimeType, TypeFormat.MediaType, filePath)
'    wseService.RequestSoapContext.Attachments.Add(lobjDimeAttachment)

'    '' Create DIME attachment and add to SOAP header
'    'Dim elemDimeAttach As New Ecmg.Cts.Providers.CEWSI35.CEWSI35.DimeAttachment("text/xml", TypeFormat.MediaType, contentFile)
'    'wseService.RequestSoapContext.Attachments.Add(elemDimeAttach)

'    ' Set DimeContent element to DIME attachment ID
'    Dim lobjDimeContent As DIMEContent = New DIMEContent
'    lobjDimeContent.Attachment = New DIMEAttachmentReference
'    lobjDimeContent.Attachment.location = lobjDimeAttachment.Id

'    ' Create reference to Content pseudo-property
'    Dim lobjContentProperty As Ecmg.Cts.Providers.CEWSI35.CEWSI35.ContentData = New Ecmg.Cts.Providers.CEWSI35.CEWSI35.ContentData
'    lobjContentProperty.Value = CType(lobjDimeContent, ContentType)
'    lobjContentProperty.propertyId = "Content"

'    ' Assign Content property to ContentTransfer object 
'    lobjContentTransfer.Property(2) = lobjContentProperty

'    ' Create and assign ContentType string-valued property to ContentTransfer object
'    'Dim lobjContentTypeProperty As SingletonString = New SingletonString
'    Dim lobjContentTypeProperty As SingletonString = Factory.NewSingletonString("ContentType", fileMimeType)

'    '' Set MIME-type to XML
'    lobjContentTransfer.Property(0) = lobjContentTypeProperty

'    ' Create and assign RetrievalName string-valued property to ContentTransfer object
'    Dim lobjRetrievalNameProperty As SingletonString = Factory.NewSingletonString("RetrievalName", "annotation.xml")
'    lobjContentTransfer.Property(1) = lobjRetrievalNameProperty

'    '' Build a list of properties to set in the new doc
'    Dim elemInputProps As ModifiablePropertyType() = New ModifiablePropertyType(1) {}
'    '' Assign list of document properties to set in ChangeRequestType element
'    elemChangeRequestType.ActionProperties = elemInputProps

'    '' Build a list of properties to exclude on the new doc object that will be returned
'    Dim excludeProps As String() = New String(1) {}
'    excludeProps(0) = "Owner"
'    excludeProps(1) = "DateLastModified"

'    '' Assign the list of excluded properties to the ChangeRequestType element
'    elemChangeRequestType.RefreshFilter = New PropertyFilterType()
'    elemChangeRequestType.RefreshFilter.ExcludeProperties = excludeProps

'    Try
'      ' Call ExecuteChanges operation to implement the doc creation and checkin
'      elemChangeResponseTypeArray = wseService.ExecuteChanges(elemExecuteChangesRequest)
'    Catch ex As System.Net.WebException
'      Throw ex
'      '  Console.WriteLine("An exception occurred while creating a document: [" + ex.Message & "]")
'      '  Return
'    End Try

'    '' The new document object should be returned, unless there is an error
'    If elemChangeResponseTypeArray Is Nothing OrElse elemChangeResponseTypeArray.Length < 1 Then
'      Console.WriteLine("A valid object was not returned from the ExecuteChanges operation")
'      Return Nothing
'    End If

'  End Function
'#End Region
'  Public Sub GenerateAnnotationContent(ByVal annotationClassId As String, ByVal annotationId As String, ByVal filePath As String)
'    Dim annotationXml As String = GenerateAnnotationXML(annotationId, annotationClassId)
'    Dim fileWrite As New StreamWriter(filePath)
'    fileWrite.Write(annotationXml)
'    fileWrite.Close()
'  End Sub

'End Class

#End Region

'Used for String Builder
Imports System.Text

''' <summary>
''' A base implementation of EventSink class.
''' This class is an event sink for Visio events. It handles event
''' notification by implementing the IVisEventProc interface, which is
''' defined in the Visio type library. In order to be notified of events,
''' an instance of this class must be passed as the eventSink argument in
''' calls to the <see href="http://msdn.microsoft.com/en-us/library/ms367505%28v=office.12%29.aspx">AddAdvise</see> method.
''' This class demonstrates how to handle event notification for events
''' raised by Application, Document, and Page objects. This class also
''' demonstrates how to respond to QueryCancel events.
''' Recommended use is to create an instance of the class and override any event handlers you want
''' to use in your application.
''' </summary>
<System.Runtime.InteropServices.ComVisible(True)> Public Class EventSink
    Implements Visio.IVisEventProc
    Implements IDisposable

#Region "Variable and Property Declarations"

    ''' <summary>
    ''' Indicates whether or not the object has been flagged for disposal.
    ''' </summary>
    Protected _disposed As Boolean = False

    ''' <summary>
    ''' Tab character used to simplify debug printing.
    ''' </summary>
    Protected Const _tab As String = ControlChars.Tab

    ''' <summary>
    ''' A dictionary of eventDescriptions keyed by eventCode.
    ''' </summary>
    Protected _eventDescriptions As System.Collections.Specialized.StringDictionary

#End Region

#Region "Constructors"

    ''' <summary>
    ''' Initializes a new instance of the <see cref="EventSink">EventSink</see> class.
    ''' The constructor also initializes the event descriptions dictionary.
    ''' </summary>
    Public Sub New()
        initializeStrings()
    End Sub

#End Region

#Region "Initialization Methods"

    ''' <summary>
    ''' This method adds an event description to the
    ''' <see cref="_eventDescriptions">eventDescriptions</see> dictionary.
    ''' </summary>
    ''' <param name="eventCode">The <see href="http://msdn.microsoft.com/en-us/library/aa342177%28v=office.12%29.aspx">Event Code</see> of the event.</param>
    ''' <param name="description">Short description of the event specified by the <paramref name="eventCode">eventCode</paramref>.</param>
    Protected Sub addEventDescription(ByVal eventCode As Short, _
                                      ByVal description As String)

        Dim key As String
        key = Convert.ToString(eventCode, _
                               System.Globalization.CultureInfo.InvariantCulture)
        _eventDescriptions.Add(key, description)
    End Sub

    ''' <summary>
    ''' This method returns a short description for the given eventCode from
    ''' the <see cref="_eventDescriptions">eventDescriptions</see> dictionary.
    ''' </summary>
    ''' <param name="eventCode">The <see href="http://msdn.microsoft.com/en-us/library/aa342177%28v=office.12%29.aspx">Event Code</see> of the event.</param>
    ''' <returns>Short description of the event specified by the <paramref name="eventCode">eventCode</paramref></returns>
    Protected Function getEventDescription(ByVal eventCode As Short) As String
        Dim description As String
        Dim key As String

        key = Convert.ToString(eventCode, _
                               System.Globalization.CultureInfo.InvariantCulture)
        description = _eventDescriptions(key)

        If (description Is Nothing) Then
            description = "NoEventDescription"
        End If
        Return description
    End Function

    ''' <summary>
    ''' This method populates the <see cref="_eventDescriptions">eventDescriptions</see> dictionary
    ''' with a short description of each Visio <see href="http://msdn.microsoft.com/en-us/library/aa342177%28v=office.12%29.aspx">Event Code</see>.
    ''' </summary>
    Protected Sub initializeStrings()
        'Instantiate the dictionary
        _eventDescriptions = New System.Collections.Specialized.StringDictionary

        'Get all the eventCode values from the enumeration as an array
        Dim values As Short() = System.[Enum].GetValues(GetType(AllEvents))
        'Get all the eventCode names from the enumeration as an array
        Dim names As String() = System.[Enum].GetNames(GetType(AllEvents))

        'Loop through the arrays, adding event descriptions to the dictionary.
        For arrIndex = 0 To values.Count - 1
            addEventDescription( _
                values(arrIndex), _
                names(arrIndex))
        Next arrIndex
    End Sub

#End Region

#Region "VisEventProc Handling"

    ''' <summary>
    ''' This method is called by Visio when an event in the
    ''' EventList collection has been triggered. This method is an
    ''' implementation of IVisEventProc.VisEventProc method.
    ''' Uses a drill-down case-based approach to fire the appropriate
    ''' individual event handler.</summary>
    ''' <param name="eventCode">Event code of the event that fired.</param>
    ''' <param name="source">Reference to source of the event.</param>
    ''' <param name="eventId">Unique identifier of the event object that fired.</param>
    ''' <param name="eventSequenceNumber">Relative position of the event in the event list</param>
    ''' <param name="subject">Reference to the subject of the event.</param>
    ''' <param name="moreInformation">Additional information for the event.</param>
    ''' <returns>False to allow a QueryCancel operation or True to cancel a QueryCancel 
    ''' operation. The return value is ignored by Visio unless the event is a QueryCancel event.</returns>
    ''' <seealso cref="Visio.IVisEventProc"></seealso>
    Public Function VisEventProc(ByVal eventCode As Short, _
                                 ByVal source As Object, _
                                 ByVal eventId As Integer, _
                                 ByVal eventSequenceNumber As Integer, _
                                 ByVal subject As Object, _
                                 ByVal moreInformation As Object) As Object _
                                 Implements Visio.IVisEventProc.VisEventProc

        Dim messageBuilder As New StringBuilder
        Dim message As String = ""
        Dim name As String = ""
        Dim eventInformation As String = ""
        Dim returnValue As Object = True

        Dim subjectApplication As Visio.Application = Nothing
        Dim subjectDocument As Visio.Document
        Dim subjectPage As Visio.Page
        Dim subjectMaster As Visio.Master
        Dim subjectSelection As Visio.Selection
        Dim subjectShape As Visio.Shape
        Dim subjectCell As Visio.Cell
        Dim subjectConnects As Visio.Connects
        Dim subjectStyle As Visio.Style
        Dim subjectWindow As Visio.Window
        Dim subjectMouseEvent As Visio.MouseEvent
        Dim subjectKeyboardEvent As Visio.KeyboardEvent
        Dim subjectDataRecordset As Visio.DataRecordset
        Dim subjectDataRecordsetChangedEvent As Visio.DataRecordsetChangedEvent

        Try

            Select Case (eventCode)
                ' Document event codes
                Case _
                    AllEvents.BeforeDocumentClose, _
                    AllEvents.BeforeDocumentSave, _
                    AllEvents.BeforeDocumentSaveAs, _
                    AllEvents.DesignModeEntered, _
                    AllEvents.DocumentAdded, _
                    AllEvents.DocumentChanged, _
                    AllEvents.DocumentCloseCanceled, _
                    AllEvents.DocumentCreated, _
                    AllEvents.DocumentOpened, _
                    AllEvents.DocumentSaved, _
                    AllEvents.DocumentSavedAs, _
                    AllEvents.RunModeEntered, _
                    AllEvents.QueryCancelDocumentClose, _
                    AllEvents.AfterRemoveHiddenInformation

                    ' Subject object is a Document
                    '   Eventinfo may be non empty. 
                    '   (1) For DocumentChanged Event it may indicate what 
                    '   changed, e.g.  /pagereordered, etc. 
                    '   (2) For the save, saveas events the eventinfo is 
                    '   typically empty. However, starting with Visio
                    '   2000 SR1 it is the name of the recover file if 
                    '   save occured for autorecovery.  In general expect
                    '   non-empty eventinfo only for SaveAs.
                    '   (3) For RemoveHiddenInformation the eventinfo
                    '   includes the data that was removed. The various types 
                    '   are represented by the following strings: 
                    '   /visRHIPersonalInfo, /visRHIMasters, /visRHIStyles,
                    '   /visRHIDataRecordsets.
                    subjectDocument = DirectCast(subject, Visio.Document)
                    subjectApplication = subjectDocument.Application
                    name = subjectDocument.Name
                    'Handle Document Events as a Group
                    Select Case (eventCode)
                        ' Document event codes
                        Case AllEvents.BeforeDocumentClose
                            HandleBeforeDocumentClose(subjectDocument)
                        Case AllEvents.BeforeDocumentSave
                            HandleBeforeDocumentSave(subjectDocument)
                        Case AllEvents.BeforeDocumentSaveAs
                            HandleBeforeDocumentSaveAs(subjectDocument)
                        Case AllEvents.DesignModeEntered
                            HandleDesignModeEntered(subjectDocument)
                        Case AllEvents.DocumentAdded
                            HandleDocumentAdded(subjectDocument)
                        Case AllEvents.DocumentChanged
                            HandleDocumentChanged(subjectDocument)
                        Case AllEvents.DocumentCloseCanceled
                            HandleDocumentCloseCanceled(subjectDocument)
                        Case AllEvents.DocumentCreated
                            HandleDocumentCreated(subjectDocument)
                        Case AllEvents.DocumentOpened
                            HandleDocumentOpened(subjectDocument)
                        Case AllEvents.DocumentSaved
                            HandleDocumentSaved(subjectDocument)
                        Case AllEvents.DocumentSavedAs
                            HandleDocumentSavedAs(subjectDocument)
                        Case AllEvents.RunModeEntered
                            HandleRunModeEntered(subjectDocument)
                        Case AllEvents.QueryCancelDocumentClose
                            returnValue = HandleQueryCancelDocumentClose(subjectDocument)
                        Case AllEvents.AfterRemoveHiddenInformation
                            HandleAfterRemoveHiddenInformation(subjectDocument)
                    End Select

                    ' Page event codes
                Case _
                    AllEvents.BeforePageDelete, _
                    AllEvents.PageAdded, _
                    AllEvents.PageChanged, _
                    AllEvents.PageDeleteCanceled, _
                    AllEvents.QueryCancelPageDelete

                    ' Subject object is a Page
                    subjectPage = DirectCast(subject, Visio.Page)
                    subjectApplication = subjectPage.Application
                    name = subjectPage.Name
                    'Handle Page Events as a Group
                    Select Case (eventCode)
                        Case AllEvents.BeforePageDelete
                            HandleBeforePageDelete(subjectPage)
                        Case AllEvents.PageAdded
                            HandlePageAdded(subjectPage)
                        Case AllEvents.PageChanged
                            HandlePageChanged(subjectPage)
                        Case AllEvents.PageDeleteCanceled
                            HandlePageDeleteCanceled(subjectPage)
                        Case AllEvents.QueryCancelPageDelete
                            returnValue = HandleQueryCancelPageDelete(subjectPage)
                    End Select

                    ' Master event codes
                Case _
                    AllEvents.BeforeMasterDelete, _
                    AllEvents.MasterChanged, _
                    AllEvents.MasterDeleteCanceled, _
                    AllEvents.MasterAdded, _
                    AllEvents.QueryCancelMasterDelete

                    ' Subject object is a Master
                    subjectMaster = DirectCast(subject, Visio.Master)
                    subjectApplication = subjectMaster.Application
                    name = subjectMaster.Name
                    'Handle Master Events as a Group
                    Select Case (eventCode)
                        Case AllEvents.BeforeMasterDelete
                            HandleBeforeMasterDelete(subjectMaster)
                        Case AllEvents.MasterChanged
                            HandleMasterChanged(subjectMaster)
                        Case AllEvents.MasterDeleteCanceled
                            HandleMasterDeleteCanceled(subjectMaster)
                        Case AllEvents.MasterAdded
                            HandleMasterAdded(subjectMaster)
                        Case AllEvents.QueryCancelMasterDelete
                            returnValue = HandleQueryCancelMasterDelete(subjectMaster)
                    End Select

                    ' Selection event codes
                Case _
                    AllEvents.BeforeSelectionDelete, _
                    AllEvents.SelectionAdded, _
                    AllEvents.ConvertToGroupCanceled, _
                    AllEvents.SelectionDeleteCanceled, _
                    AllEvents.QueryCancelSelectionDelete, _
                    AllEvents.QueryCancelUngroup, _
                    AllEvents.QueryCancelConvertToGroup, _
                    AllEvents.UngroupCanceled, _
                    AllEvents.QueryCancelGroup, _
                    AllEvents.GroupCanceled

                    ' Subject object is a Selection
                    subjectSelection = DirectCast(subject, Visio.Selection)
                    subjectApplication = subjectSelection.Application
                    'Handle Selection Events as a Group
                    Select Case (eventCode)
                        ' Selection event codes
                        Case AllEvents.BeforeSelectionDelete
                            HandleBeforeSelectionDelete(subjectSelection)
                        Case AllEvents.SelectionAdded
                            HandleSelectionAdded(subjectSelection)
                        Case AllEvents.SelectionDeleteCanceled
                            HandleSelectionDeleteCanceled(subjectSelection)
                        Case AllEvents.ConvertToGroupCanceled
                            HandleConvertToGroupCanceled(subjectSelection)
                        Case AllEvents.QueryCancelUngroup
                            returnValue = HandleQueryCancelUngroup(subjectSelection)
                        Case AllEvents.QueryCancelConvertToGroup
                            returnValue = HandleQueryConvertToGroup(subjectSelection)
                        Case AllEvents.QueryCancelSelectionDelete
                            returnValue = HandleQueryCancelSelectionDelete(subjectSelection)
                        Case AllEvents.UngroupCanceled
                            HandleUngroupCanceled(subjectSelection)
                        Case AllEvents.QueryCancelGroup
                            returnValue = HandleQueryCancelGroup(subjectSelection)
                        Case AllEvents.GroupCanceled
                            HandleGroupCanceled(subjectSelection)
                    End Select


                    ' Shape event codes
                    '*visEvtShapeDataGraphicChanged, visEvtShapeLinkAdded, & visEvtShapeLinkDeleted
                    'These features are only available in the Pro Edition of Visio.
                Case _
                    AllEvents.BeforeShapeDelete, _
                    AllEvents.BeforeShapeTextEdit, _
                    AllEvents.ShapeAdded, _
                    AllEvents.ShapeChanged, _
                    AllEvents.ShapeExitedTextEdit, _
                    AllEvents.ShapeParentChanged, _
                    AllEvents.ShapesDeleted, _
                    AllEvents.TextChanged, _
                    AllEvents.ShapeDataGraphicChanged, _
                    AllEvents.ShapeLinkAdded, _
                    AllEvents.ShapeLinkDeleted

                    ' Subject object is a Shape
                    ' EventInfo is normally empty but for ShapeChanged events
                    ' it may indicate what changed, e.g. /data1, /name, etc.
                    subjectShape = DirectCast(subject, Visio.Shape)
                    subjectApplication = subjectShape.Application
                    name = subjectShape.Name
                    'Handle Shape Events as a Group
                    Select Case (eventCode)
                        Case AllEvents.BeforeShapeDelete
                            HandleBeforeShapeDelete(subjectShape)
                        Case AllEvents.BeforeShapeTextEdit
                            HandleBeforeShapeTextEdit(subjectShape)
                        Case AllEvents.ShapeAdded
                            HandleShapeAdded(subjectShape)
                        Case AllEvents.ShapeChanged
                            HandleShapeChanged(subjectShape, moreInformation)
                        Case AllEvents.ShapeExitedTextEdit
                            HandleShapeExitedTextEdit(subjectShape)
                        Case AllEvents.ShapeParentChanged
                            HandleShapeParentChanged(subjectShape)
                        Case AllEvents.ShapesDeleted
                            HandleShapesDeleted(subjectShape)
                        Case AllEvents.TextChanged
                            HandleTextChanged(subjectShape)
                            'Pro Edition Only
                        Case AllEvents.ShapeDataGraphicChanged
                            HandleShapeDataGraphicChanged(subjectShape)
                            'Pro Edition Only
                        Case AllEvents.ShapeLinkAdded, _
                             AllEvents.ShapeLinkDeleted

                            Dim values As String
                            values = subjectApplication.EventInfo(Visio.VisEventCodes.visEvtIdMostRecent)
                            Dim parsedString As String() = Nothing
                            Dim parsedResults As String() = Nothing
                            Dim dataRecordSetID As Long
                            Dim dataRowID As Long

                            parsedString = values.Split("/")

                            For Each s As String In parsedString
                                If s.Contains("DataRecordsetID") Then
                                    parsedResults = s.Split("=")
                                    dataRecordSetID = CType(Trim(parsedResults.Last), Long)
                                ElseIf s.Contains("DataRowID") Then
                                    parsedResults = s.Split("=")
                                    dataRowID = CType(Trim(parsedResults.Last), Long)
                                End If
                            Next s

                            Select Case eventCode
                                Case AllEvents.ShapeLinkAdded
                                    HandleShapeLinkAdded(subjectShape, _
                                                        dataRecordSetID, _
                                                        dataRowID)
                                Case AllEvents.ShapeLinkDeleted
                                    HandleShapeLinkDeleted(subjectShape, _
                                                        dataRecordSetID, _
                                                        dataRowID)
                            End Select
                    End Select





                    ' Cell event codes
                Case _
                    AllEvents.CellChanged, _
                    AllEvents.FormulaChanged

                    ' Subject object is a Cell
                    subjectCell = DirectCast(subject, Visio.Cell)
                    subjectShape = subjectCell.Shape
                    subjectApplication = subjectCell.Application
                    name = subjectShape.Name + "!" + subjectCell.Name
                    'Handle Cell Events as a Group
                    Select Case (eventCode)
                        Case AllEvents.CellChanged
                            HandleCellChanged(subjectCell)
                        Case AllEvents.FormulaChanged
                            HandleFormulaChanged(subjectCell)
                    End Select

                    ' Connects event codes
                Case _
                    AllEvents.ConnectionsAdded, _
                    AllEvents.ConnectionsDeleted

                    ' Subject object is a Connects collection
                    subjectConnects = DirectCast(subject, Visio.Connects)
                    subjectApplication = subjectConnects.Application
                    'Handle Connects Events as a Group
                    Select Case (eventCode)
                        Case AllEvents.ConnectionsAdded
                            HandleConnectionsAdded(subjectConnects)
                        Case AllEvents.ConnectionsDeleted
                            HandleConnectionsDeleted(subjectConnects)
                    End Select


                    ' Style event codes
                Case _
                    AllEvents.BeforeStyleDelete, _
                    AllEvents.StyleAdded, _
                    AllEvents.StyleChanged, _
                    AllEvents.StyleDeleteCanceled, _
                    AllEvents.QueryCancelStyleDelete

                    ' Subject object is a Style
                    subjectStyle = DirectCast(subject, Visio.Style)
                    subjectApplication = subjectStyle.Application
                    name = subjectStyle.Name
                    'Handle Style Events as a Group
                    Select Case (eventCode)
                        Case AllEvents.BeforeStyleDelete
                            HandleBeforeStyleDelete(subjectStyle)
                        Case AllEvents.StyleAdded
                            HandleStyleAdded(subjectStyle)
                        Case AllEvents.StyleChanged
                            HandleStyleChanged(subjectStyle)
                        Case AllEvents.StyleDeleteCanceled
                            HandleStyleDeleteCanceled(subjectStyle)
                        Case AllEvents.QueryCancelStyleDelete
                            returnValue = HandleQueryCancelStyleDelete(subjectStyle)
                    End Select

                    ' Window event codes
                Case _
                    AllEvents.BeforeWindowClosed, _
                    AllEvents.BeforeWindowPageTurn, _
                    AllEvents.WindowOpened, _
                    AllEvents.WindowChanged, _
                    AllEvents.WindowTurnedToPage, _
                    AllEvents.BeforeWindowSelDelete, _
                    AllEvents.WindowCloseCanceled, _
                    AllEvents.WindowActivated, _
                    AllEvents.SelectionChanged, _
                    AllEvents.ViewChanged, _
                    AllEvents.QueryCancelWindowClose

                    ' Subject object is a Window
                    subjectWindow = DirectCast(subject, Visio.Window)
                    subjectApplication = subjectWindow.Application
                    name = subjectWindow.Caption
                    'Handle Window Events as a Group
                    Select Case (eventCode)
                        Case AllEvents.BeforeWindowClosed
                            HandleBeforeWindowClosed(subjectWindow)
                        Case AllEvents.BeforeWindowPageTurn
                            HandleBeforeWindowPageTurn(subjectWindow)
                        Case AllEvents.WindowOpened
                            HandleWindowOpened(subjectWindow)
                        Case AllEvents.WindowChanged
                            HandleWindowChanged(subjectWindow)
                        Case AllEvents.WindowTurnedToPage
                            HandleWindowTurnedToPage(subjectWindow)
                        Case AllEvents.BeforeWindowSelDelete
                            HandleBeforeWindowSelDelete(subjectWindow)
                        Case AllEvents.WindowCloseCanceled
                            HandleWindowCloseCanceled(subjectWindow)
                        Case AllEvents.WindowActivated
                            HandleWindowActivated(subjectWindow)
                        Case AllEvents.SelectionChanged
                            HandleSelectionChanged(subjectWindow)
                        Case AllEvents.ViewChanged
                            HandleViewChanged(subjectWindow)
                        Case AllEvents.QueryCancelWindowClose
                            returnValue = HandleQueryCancelWindowClose(subjectWindow)
                    End Select

                    ' Application event codes
                Case _
                    AllEvents.AfterModal, _
                    AllEvents.AfterResume, _
                    AllEvents.AppActivated, _
                    AllEvents.AppDeactivated, _
                    AllEvents.AppObjActivated, _
                    AllEvents.AppObjDeactivated, _
                    AllEvents.BeforeModal, _
                    AllEvents.BeforeQuit, _
                    AllEvents.BeforeSuspend, _
                    AllEvents.EnterScope, _
                    AllEvents.ExitScope, _
                    AllEvents.MarkerEvent, _
                    AllEvents.MustFlushScopeBeginning, _
                    AllEvents.MustFlushScopeEnded, _
                    AllEvents.NoEventsPending, _
                    AllEvents.OnKeystrokeMessageForAddon, _
                    AllEvents.QueryCancelQuit, _
                    AllEvents.QueryCancelSuspend, _
                    AllEvents.QuitCanceled, _
                    AllEvents.SuspendCanceled, _
                    AllEvents.VisioIsIdle

                    ' Subject object is an Application
                    ' EventInfo is empty for most of these events.  However for
                    ' the Marker event, the EnterScope event and the ExitScope 
                    ' event eventinfo contains the context string. 
                    subjectApplication = DirectCast(subject, Visio.Application)
                    'Handle Application Events as a Group
                    Select Case (eventCode)
                        Case AllEvents.AfterModal
                            HandleAfterModal(subjectApplication)
                        Case AllEvents.AfterResume
                            HandleAfterResume(subjectApplication)
                        Case AllEvents.AppActivated
                            HandleAppActivated(subjectApplication)
                        Case AllEvents.AppDeactivated
                            HandleAppDeactivated(subjectApplication)
                        Case AllEvents.AppObjActivated
                            HandleAppObjActivated(subjectApplication)
                        Case AllEvents.AppObjDeactivated
                            HandleAppObjDeactivated(subjectApplication)
                        Case AllEvents.BeforeModal
                            HandleBeforeModal(subjectApplication)
                        Case AllEvents.BeforeQuit
                            HandleBeforeQuit(subjectApplication)
                        Case AllEvents.BeforeSuspend
                            HandleBeforeSuspend(subjectApplication)
                        Case AllEvents.EnterScope
                            eventInformation = subjectApplication.EventInfo(Visio.VisEventCodes.visEvtIdMostRecent)
                            HandleEnterScope(subjectApplication, moreInformation)
                        Case AllEvents.ExitScope
                            eventInformation = subjectApplication.EventInfo(Visio.VisEventCodes.visEvtIdMostRecent)
                            HandleExitScope(subjectApplication, moreInformation)
                        Case AllEvents.MarkerEvent
                            eventInformation = subjectApplication.EventInfo(Visio.VisEventCodes.visEvtIdMostRecent)
                            HandleMarkerEvent(subjectApplication, eventSequenceNumber, moreInformation)
                        Case AllEvents.MustFlushScopeBeginning
                            HandleMustFlushScopeBeginning(subjectApplication)
                        Case AllEvents.MustFlushScopeEnded
                            HandleMustFlushScopeEnded(subjectApplication)
                        Case AllEvents.NoEventsPending
                            HandleNoEventsPending(subjectApplication)
                        Case AllEvents.OnKeystrokeMessageForAddon
                            returnValue = HandleOnKeystrokeMessageForAddon(subjectApplication)
                        Case AllEvents.QueryCancelQuit
                            returnValue = HandleQueryCancelQuit(subjectApplication)
                        Case AllEvents.QueryCancelSuspend
                            returnValue = HandleQueryCancelSuspend(subjectApplication)
                        Case AllEvents.QuitCanceled
                            HandleQuitCanceled(subjectApplication)
                        Case AllEvents.SuspendCanceled
                            HandleSuspendCanceled(subjectApplication)
                        Case AllEvents.VisioIsIdle
                            HandleVisioIsIdle(subjectApplication)
                    End Select

                    ' Mouse Event 
                Case _
                    AllEvents.MouseDown, _
                    AllEvents.MouseMove, _
                    AllEvents.MouseUp

                    ' Subject is MouseEvent object
                    ' Note Mouse events can also be canceled. 
                    ' EventInfo may be non-empty for MouseMove events, and 
                    ' contains information about the DragState which is also
                    ' exposed as a property on the MouseEvent object.                        
                    subjectMouseEvent = DirectCast(subject, Visio.MouseEvent)
                    'Handle Mouse Events as a Group
                    Select Case (eventCode)
                        Case AllEvents.MouseDown
                            returnValue = HandleMouseDown(subjectMouseEvent)
                        Case AllEvents.MouseMove
                            returnValue = HandleMouseMove(subjectMouseEvent)
                        Case AllEvents.MouseUp
                            returnValue = HandleMouseUp(subjectMouseEvent)
                    End Select

                    ' Keyboard Event		
                Case _
                    AllEvents.KeyDown, _
                    AllEvents.KeyPress, _
                    AllEvents.KeyUp

                    ' Subject is KeyboardEvent object
                    ' Note KeyboardEvents can also be canceled. 
                    subjectKeyboardEvent = DirectCast(subject, Visio.KeyboardEvent)
                    'Handle Keyboard Events as a Group
                    Select Case (eventCode)
                        Case AllEvents.KeyDown
                            returnValue = HandleKeyDown(subjectKeyboardEvent)
                        Case AllEvents.KeyPress
                            returnValue = HandleKeyPress(subjectKeyboardEvent)
                        Case AllEvents.KeyUp
                            returnValue = HandleKeyUp(subjectKeyboardEvent)
                    End Select

                    ' Data Recordset Event codes
                    ' Data Recordset event with DataRecordset object
                Case _
                    AllEvents.DataRecordsetAdded, _
                    AllEvents.BeforeDataRecordsetDelete

                    ' Subject object is a DataRecordset
                    subjectDataRecordset = DirectCast(subject, Visio.DataRecordset)
                    'Handle DataRecordset Events as a Group
                    Select Case (eventCode)
                        Case AllEvents.DataRecordsetAdded
                            HandleDataRecordsetAdded(subjectDataRecordset)
                        Case AllEvents.BeforeDataRecordsetDelete
                            HandleBeforeDataRecordsetDelete(subjectDataRecordset)
                    End Select

                    ' Data Recordset events with DataRecordsetChangedEvent Object
                Case AllEvents.DataRecordsetChanged

                    ' Subject is DataRecordsetChangedEvent Object
                    subjectDataRecordsetChangedEvent = DirectCast(subject, Visio.DataRecordsetChangedEvent)
                    'Handle DataRecordSetChangedEvent as a Group
                    Select Case (eventCode)
                        Case AllEvents.DataRecordsetChanged
                            HandleDataRecordsetChanged(subjectDataRecordsetChangedEvent)
                    End Select
                Case Else
                    name = "Unknown"
                    subjectApplication = Nothing
                    HandleUnknownEvent(eventCode, subject, moreInformation)
            End Select

            ' Get a description for this event code
            messageBuilder.Append(getEventDescription(eventCode))

            ' Append the name of the subject object
            If (name.Length > 0) Then
                messageBuilder.Append(": " + name)
            End If



            ' Append event info when it has been set early... aka Marker event
            ' This can trigger other events which will throw us off. 
            If (Not eventInformation Is Nothing) Then
                messageBuilder.Append(_tab + eventInformation)
            End If
            ' Append event info when it is available
            If (Not subjectApplication Is Nothing) Then

                eventInformation = subjectApplication.EventInfo(Visio.VisEventCodes.visEvtIdMostRecent)

                If (Not eventInformation Is Nothing) Then
                    messageBuilder.Append(_tab + eventInformation)
                End If
            End If

            ' Append moreInformation when it is available
            If (Not moreInformation Is Nothing) Then
                messageBuilder.Append(_tab + moreInformation.ToString)
            End If

            ' Get the targetArgs string from the event object. 
            ' TargetArgs are added to the event object in the AddAdvise method
            Dim events As Visio.EventList
            Dim thisEvent As Visio.Event
            Dim sourceType As String
            Dim targetArgs As String

            sourceType = source.GetType().FullName

            If (sourceType.Contains(GetType(Visio.Application).FullName)) Then
                events = DirectCast(source, Visio.Application).EventList
            ElseIf (sourceType.Contains(GetType(Visio.Document).FullName)) Then
                events = DirectCast(source, Visio.Document).EventList
            ElseIf (sourceType.Contains(GetType(Visio.Page).FullName)) Then
                events = DirectCast(source, Visio.Page).EventList
            Else
                events = Nothing
            End If

            If (Not events Is Nothing) Then
                thisEvent = events.ItemFromID(eventId)
                targetArgs = thisEvent.TargetArgs
                ' Append targetArgs when it is available
                If (targetArgs.Length > 0) Then
                    messageBuilder.Append(" " + targetArgs)
                End If
            End If


            message = messageBuilder.ToString

            ' Write the event info to the output window
            logTheEvent(message)

        Catch err As System.Runtime.InteropServices.COMException
            System.Diagnostics.Debug.WriteLine(err.Message)
        End Try

        Return returnValue
    End Function

#End Region

#Region "Overridable Event Handlers"

#Region "Document Event Handlers"

    ''' <summary>
    ''' Handles the BeforeDocumentClose event.
    ''' Occurs before a document is closed.
    ''' Resets class-level variables and resets the MainUI
    ''' </summary>
    ''' <param name="doc">The document which is going to be closed.</param>
    Protected Overridable Sub HandleBeforeDocumentClose(ByVal doc As Visio.Document)
    End Sub
    ''' <summary>
    ''' Handles the before document save.
    ''' Occurs before a document is saved.
    ''' </summary>
    ''' <param name="Doc">The document being saved.</param>
    Protected Overridable Sub HandleBeforeDocumentSave(ByVal doc As Visio.Document)
    End Sub
    ''' <summary>
    ''' Handles the BeforeDocumentSaveAs Event
    ''' Occurs just before a document is saved by using the Save As command.
    ''' This is used to intercept the save event and process it using the addin's routine which 
    ''' makes sure the information is exported to the database properly.
    ''' </summary>
    ''' <param name="doc">The document that is going to be saved.</param>
    ''' <remarks>
    ''' The BeforeDocumentSaveAs event fires when a document is saved to either a 
    ''' native format (for example, VSD or VDX) or a non-native format (for example, HTM or BMP). 
    ''' It does not fire when a document is saved to DWG, DXF, and DGN formats. To save a 
    ''' document in a non-native format programmatically, you must use the Export method of 
    ''' the Page object. Note that when you call the SaveAs method, Microsoft Office Visio 
    ''' fires first the BeforeDocumentSaveAs event and then the DocumentSavedAs event. Calling 
    ''' the Export method, however, fires the BeforeDocumentSaveAs event but not the 
    ''' DocumentSavedAs event that follows it in response to the SaveAs method. The BeforeDocumentSaveAs event 
    ''' is one of a group of events for which the EventInfo property of the Application object contains extra information.
    ''' If the BeforeDocumentSaveAs event is fired because a save was initiated by a user or a 
    ''' program, the EventInfo property returns the following string:
    ''' "/saveasfile= &lt;filename&gt;"
    ''' If it fires because Visio is saving a copy of an open file (for autorecovery or to 
    ''' include as a mail attachment), the EventInfo property will return one of the 
    ''' following strings:
    ''' If the event is fired for autorecovery purposes, the name of a recovery file in 
    ''' this format: "/autosavefile=C:\TEMP\~$2VSO2FD.vsd"
    ''' If the event is fired because a document copy is being made to send as a mail 
    ''' attachment, the name of an attachment file in this format: 
    ''' "/mailfile=C:\TEMP\~$2VSO2FD.vsd"
    ''' </remarks>
    Protected Overridable Sub HandleBeforeDocumentSaveAs(ByVal doc As Visio.Document)
    End Sub
    ''' <summary>
    ''' Handles the DesignModeEntered event.
    ''' Occurs before a document enters design mode.
    ''' </summary>
    ''' <param name="doc">The document that is going to enter design mode.</param>
    Protected Overridable Sub HandleDesignModeEntered(ByVal doc As Visio.Document)
    End Sub
    ''' <summary>
    ''' Handles the DocumentAdded event.
    ''' Occurs when a document is added to a Visio Application instance's
    ''' Documents collection.
    ''' </summary>
    ''' <param name="doc">The documnet being added to the application's Documents collection.</param>
    Protected Overridable Sub HandleDocumentAdded(ByVal doc As Visio.Document)
    End Sub
    ''' <summary>
    ''' Handles the DocumentChanged event
    ''' Occurs after certain properties of a document are changed.
    ''' </summary>
    ''' <param name="doc">The document whose properties were changed.</param>
    ''' <remarks>The DocumentChanged event indicates that one of a document's properties, such as Author or Description, has changed.</remarks>
    Protected Overridable Sub HandleDocumentChanged(ByVal doc As Visio.Document)
    End Sub
    ''' <summary>
    ''' Handles the document close canceled.
    ''' Occurs after an event handler has returned True (cancel) to a QueryCancelDocumentClose event.
    ''' </summary>
    ''' <param name="doc">The document that was going to be closed.</param>
    Protected Overridable Sub HandleDocumentCloseCanceled(ByVal doc As Visio.Document)
    End Sub
    ''' <summary>
    ''' Handles the DocumentCreated event
    ''' Occurs after a document is created.
    ''' </summary>
    ''' <param name="doc">The document that was created.</param>
    ''' <remarks>You can add DocumentCreated events to the EventList collection of an 
    ''' Application object, Documents collection, or Document object. 
    ''' The first two are straightforward; if a document is opened or created in the 
    ''' scope of the Application object or its Documents collection, the DocumentCreated 
    ''' event occurs. However, adding a DocumentCreated event to the EventList 
    ''' collection of a Document object makes sense only if the event's action is 
    ''' visActCodeRunAddon. In this case, the event is persistable; it can be stored 
    ''' with the document. If the document that contains the persistent event is 
    ''' opened, its action is triggered. If a new document is based on or copied 
    ''' from the document that contains the persistent event, the DocumentCreated 
    ''' event is copied to the new document and its action is triggered. However, 
    ''' if the event's action is visActCodeAdvise, that event is not persistable and 
    ''' therefore is not stored with the document; hence, it is never triggered.</remarks>
    Protected Overridable Sub HandleDocumentCreated(ByVal doc As Visio.Document)
    End Sub
    ''' <summary>
    ''' Handles the DocumentOpened Event
    ''' Occurs when a document is opened.
    ''' This will test whether or not the addin should set the class-level document = 
    ''' the opened document . 
    ''' </summary>
    ''' <param name="doc">The document that was opened.</param>
    ''' <remarks>The DocumentOpened event is often added to the EventList collection of a 
    ''' Microsoft Visio template file (.vst). The event's action is triggered whenever an 
    ''' existing document is opened.</remarks>
    Protected Overridable Sub HandleDocumentOpened(ByVal doc As Visio.Document)
    End Sub
    ''' <summary>
    ''' Handles the DocumentSaved event.
    ''' Occurs after a document is saved.
    ''' </summary>
    ''' <param name="doc">The document that was saved.</param>
    Protected Overridable Sub HandleDocumentSaved(ByVal doc As Visio.Document)
    End Sub
    ''' <summary>
    ''' Handles the DocumentSavedAs event.
    ''' Occurs after a document is saved by using the Save As command.
    ''' </summary>
    ''' <param name="doc">The document which was saved.</param>
    ''' <remarks>The DocumentSavedAs event is one of a group of events for which the 
    ''' EventInfo property of the Application object contains extra information.
    ''' If the DocumentSavedAs event is fired because a save was initiated by a user 
    ''' or a program, the EventInfo property returns the following string:
    ''' "/saveasfile= &lt;filename&gt;"
    ''' If it fires because Microsoft Office Visio is saving a copy of an open file 
    ''' (for autorecovery or to include as a mail attachment), the EventInfo property returns one of the following strings:
    ''' If the event is fired for autorecovery purposes, the name of a recovery file 
    ''' in this format: "/autosavefile=drivename:\foldername\filename"
    ''' If the event is fired because a document copy is being made to send as a mail 
    ''' attachment, the name of an attachment file in this format: "/mailfile=drivename:\foldername\filename"</remarks>
    Protected Overridable Sub HandleDocumentSavedAs(ByVal doc As Visio.Document)
    End Sub
    ''' <summary>
    ''' Handles the RunModeEntered event.
    ''' Occurs after a document enters run mode.
    ''' </summary>
    ''' <param name="doc">The document which entered run-mode.</param>
    Protected Overridable Sub HandleRunModeEntered(ByVal doc As Visio.Document)
    End Sub
    ''' <summary>
    ''' Handles the QueryCancelDocumentClose event.
    ''' Occurs before the application closes a document in response to a user action 
    ''' in the interface. If any event handler returns True, the operation is canceled.
    ''' 'Visio will prompt to save if the document isn't saved and this query returns False
    ''' </summary>
    ''' <param name="doc">The document that is going to be closed.</param>
    ''' <returns>
    ''' True (cancel), the instance will fire 
    ''' DocumentCloseCanceled and not close the document.
    ''' False (don't cancel), the instance will fire 
    ''' BeforeDocumentClose and then close the document.
    ''' </returns>
    Protected Overridable Function HandleQueryCancelDocumentClose(ByVal doc As Visio.Document) As Boolean
        'We are not currently handling events, exit.
        Return False
    End Function
    ''' <summary>
    ''' Handles the AfterRemoveHiddenInformation event.
    ''' Occurs when hidden information is removed from the document.
    ''' </summary>
    ''' <param name="doc">The document from which hidden information has been removed.</param>
    ''' <remarks>The AfterRemoveHiddenInformation event is one of a group of events for 
    ''' which the EventInfo property of the Application object contains extra information. 
    ''' When the AfterRemoveHiddenInformation event is fired, the EventInfo property 
    ''' returns a string that contains information about which items were removed 
    ''' from the document, consisting of the sum of applicable constant values
    ''' from the VisRemoveHiddenInfoItems enumeration.</remarks>
    Protected Overridable Sub HandleAfterRemoveHiddenInformation(ByVal doc As Visio.Document)
    End Sub

#End Region

#Region "Page Event Handlers"

    ''' <summary>
    ''' Handles the BeforePageDelete event.
    ''' Occurs before a page is deleted.
    ''' </summary>
    ''' <param name="vsoPage">The page that is going to be deleted.</param>
    Protected Overridable Sub HandleBeforePageDelete(ByVal vsoPage As Visio.Page)
    End Sub
    ''' <summary>
    ''' Handles the PageAdded event.
    ''' Occurs after a new page is added to a document.
    ''' </summary>
    ''' <param name="vsoPage">The page that was added.</param>
    Protected Overridable Sub HandlePageAdded(ByVal vsoPage As Visio.Page)
    End Sub
    ''' <summary>
    ''' Handles the PageChanged event.
    ''' Occurs after the name of a page, the background page associated with a page, 
    ''' or the page type (foreground or background) changes.
    ''' </summary>
    ''' <param name="vsoPage">The page that changed.</param>
    Protected Overridable Sub HandlePageChanged(ByVal vsoPage As Visio.Page)
    End Sub
    ''' <summary>
    ''' Handles the PageDeleteCanceled event.
    ''' Occurs after an event handler has returned True (cancel) to a 
    ''' QueryCancelPageDelete event.
    ''' </summary>
    ''' <param name="vsoPage">The page that was going to be deleted.</param>
    Protected Overridable Sub HandlePageDeleteCanceled(ByVal vsoPage As Visio.Page)
    End Sub
    ''' <summary>
    ''' Handles the QueryCancelPageDelete event.
    ''' Occurs before the application deletes a page in response to a user action in the interface. 
    ''' If any event handler returns True, the operation is canceled.
    ''' Will only allow a page to be deleted if all shapes with valid Shape_Key values have been deleted.
    ''' </summary>
    ''' <param name="vsoPage">The page that is going to be deleted.</param>
    ''' <returns>
    ''' True (cancel), the instance will fire 
    ''' PageDeleteCanceled and not delete the page.
    ''' False (don't cancel) the instance will fire 
    ''' BeforePageDelete, and then delete the page.
    ''' </returns>
    Protected Overridable Function HandleQueryCancelPageDelete(ByVal vsoPage As Visio.Page) As Boolean
        Return False
    End Function

#End Region

#Region "Master Event Handlers"

    ''' <summary>
    ''' Handles the BeforeMasterDelete event.
    ''' Occurs before a master is deleted from a document.
    ''' </summary>
    ''' <param name="vsoMaster">The master that is going to be deleted.</param>
    Protected Overridable Sub HandleBeforeMasterDelete(ByVal vsoMaster As Visio.Master)
    End Sub
    ''' <summary>
    ''' Handles the MasterChanged event.
    ''' Occurs after properties of a master are changed and propagated to its instances.
    ''' </summary>
    ''' <param name="vsoMaster">The master whose properties changed.</param>
    Protected Overridable Sub HandleMasterChanged(ByVal vsoMaster As Visio.Master)
    End Sub
    ''' <summary>
    ''' Handles the MasterDeleteCanceled event.
    ''' Occurs after an event handler has returned True (cancel) to a 
    ''' QueryCancelMasterDelete event.
    ''' </summary>
    ''' <param name="vsoMaster">The master that was going to be deleted.</param>
    Protected Overridable Sub HandleMasterDeleteCanceled(ByVal vsoMaster As Visio.Master)
    End Sub
    ''' <summary>
    ''' Handles the MasterAdded event.
    ''' Occurs after a new master is added to a document.
    ''' </summary>
    ''' <param name="vsoMaster">The master that was added to the document.</param>
    Protected Overridable Sub HandleMasterAdded(ByVal vsoMaster As Visio.Master)
    End Sub
    ''' <summary>
    ''' Handles the QueryCancelMasterDelete event.
    ''' Occurs before the application deletes a master in response to a user action 
    ''' in the interface. If any event handler returns True, the operation is canceled.
    ''' </summary>
    ''' <param name="vsoMaster">The master that is going to be deleted.</param>
    ''' <returns>
    ''' True (cancel), the instance will fire 
    ''' MasterDeleteCanceled and not delete the master.
    ''' False (don't cancel), the instance will fire 
    ''' BeforeMasterDelete and then delete the master.
    ''' </returns>
    Protected Overridable Function HandleQueryCancelMasterDelete(ByVal vsoMaster As Visio.Master) As Boolean
        Return False
    End Function

#End Region

#Region "Selection Event Handlers"

    ''' <summary>
    ''' Handles the BeforeSelectionDelete event.
    ''' Occurs before selected objects are deleted.
    ''' </summary>
    ''' <param name="vsoSelection">The selected objects that are going to be deleted.</param>
    Protected Overridable Sub HandleBeforeSelectionDelete(ByVal vsoSelection As Visio.Selection)
    End Sub
    ''' <summary>
    ''' Handles the SelectionAdded event.
    ''' Occurs after one or more shapes are added to a document.
    ''' </summary>
    ''' <param name="vsoSelection">The selection of shapes that was added to the document.</param>
    ''' <remarks>A Shape object can serve as the source object for the SelectionAdded 
    ''' event if the shape's Type property is visTypeGroup(2) or visTypePage(1).
    ''' The SelectionAdded and ShapeAdded events are similar in that they both fire 
    ''' after shape(s) are created. They differ in how they behave when a single 
    ''' operation adds several shapes. Suppose a Paste operation creates three new 
    ''' shapes. The ShapeAdded event fires three times and acts on each of the three 
    ''' objects. The SelectionAdded event fires once, and it acts on a Selection object 
    ''' in which the three new shapes are selected.</remarks>
    Protected Overridable Sub HandleSelectionAdded(ByVal vsoSelection As Visio.Selection)
    End Sub
    ''' <summary>
    ''' Handles the SelectionDeleteCanceled event.
    ''' Occurs after an event handler has returned True (cancel) to a 
    ''' QueryCancelSelectionDelete event.
    ''' </summary>
    ''' <param name="vsoSelection">The selection of shapes that was going to be deleted.</param>
    Protected Overridable Sub HandleSelectionDeleteCanceled(ByVal vsoSelection As Visio.Selection)
    End Sub
    ''' <summary>
    ''' Handles the ConvertToGroupCanceled event.
    ''' Occurs after an event handler has returned True (cancel) to a 
    ''' QueryCancelConvertToGroup event.
    ''' </summary>
    ''' <param name="vsoSelection">The selection of shapes that was going to be grouped.</param>
    Protected Overridable Sub HandleConvertToGroupCanceled(ByVal vsoSelection As Visio.Selection)
    End Sub
    ''' <summary>
    ''' Handles the GroupCanceled event.
    ''' Occurs after an event handler has returned True (cancel) to a 
    ''' QueryCancelGroup event.
    ''' </summary>
    ''' <param name="vsoSelection">The selection of shapes that was going to be grouped.</param>
    Protected Overridable Sub HandleGroupCanceled(ByVal vsoSelection As Visio.Selection)
    End Sub
    ''' <summary>
    ''' Handles the UngroupCanceled event.
    ''' Occurs after an event handler has returned True (cancel) to a QueryCancelUngroup event.
    ''' </summary>
    ''' <param name="vsoSelection">The selection of shapes that was going to be ungrouped.</param>
    ''' <remarks>Resets the Application.AlertResponse property to its default.</remarks>
    Protected Overridable Sub HandleUngroupCanceled(ByVal vsoSelection As Visio.Selection)
    End Sub
    ''' <summary>
    ''' Handles the query cancel ungroup.
    ''' Occurs before the application ungroups a selection of shapes in response to a 
    ''' user action in the interface. If any event handler returns True, 
    ''' the operation is canceled.
    ''' Don't allow ungrouping of Asset shapes.
    ''' </summary>
    ''' <param name="vsoSelection">The selection of shapes that is going to be ungrouped.</param>
    ''' <returns>
    ''' True (cancel), the instance will fire 
    ''' UngroupCanceled and not ungroup the shapes.
    ''' False (don't cancel), the instance will fire 
    ''' ShapeParentChanged, BeforeSelectionDelete, and BeforeShapeDelete, 
    ''' and then ungroup the shapes.
    ''' </returns>
    ''' <remarks>Sets the Application.AlertResponse property = IDCancel(2) temporarily, since
    ''' Visio still brings up the "This will break the link to master" dialog box.</remarks>
    Protected Overridable Function HandleQueryCancelUngroup(ByVal vsoSelection As Visio.Selection) As Boolean
        Return False
    End Function
    ''' <summary>
    ''' Handles the query convert to group.
    ''' Occurs before the application converts a selection of shapes to a 
    ''' group in response to a user action in the interface. 
    ''' If any event handler returns True, the operation is canceled.
    ''' </summary>
    ''' <param name="vsoSelection">The selection of shapes that is going to be converted to a group.</param>
    ''' <returns>
    ''' True (cancel), the instance will fire 
    ''' ConvertToGroupCanceled and not convert the shapes.
    ''' False (don't cancel), the conversion will be performed.
    ''' </returns>
    ''' <remarks>In some cases, such as when a shape that has a ForeignType property 
    ''' of visTypeMetafile is converted to a group, the initial shape will be deleted 
    ''' and replaced with new shapes. In such cases, the Visio instance will 
    ''' subsequently fire BeforeSelectionDelete and BeforeShapeDelete events before 
    ''' converting the shapes.</remarks>
    Protected Overridable Function HandleQueryConvertToGroup(ByVal vsoSelection As Visio.Selection) As Boolean
        Return False
    End Function
    ''' <summary>
    ''' Handles the QueryCancelSelectionDelete event.
    ''' Occurs before the application deletes a selection of shapes in response to a user action in the interface. 
    ''' If any event handler returns True, the operation is canceled.
    ''' </summary>
    ''' <param name="vsoSelection">The selection of shapes which are slated for deletion.</param>
    ''' <returns>
    ''' True (cancel), the instance will fire 
    ''' SelectionDeleteCanceled and not delete the shapes.
    ''' False (don't cancel), the instance will fire 
    ''' BeforeSelectionDelete and BeforeShapeDelete and then delete the shapes.
    ''' </returns>
    Protected Overridable Function HandleQueryCancelSelectionDelete(ByVal vsoSelection As Visio.Selection) As Boolean
        Return False
    End Function
    ''' <summary>
    ''' Handles the QueryCancelGroup Event.
    ''' Occurs before the application groups a selection of shapes in response to a 
    ''' user action in the interface. 
    ''' If any event handler returns True, the operation is canceled. 
    ''' </summary>
    ''' <param name="vsoSelection">The selection to be grouped.</param>
    ''' <returns>
    ''' True (cancel), the instance fires 
    ''' GroupCanceled and not group the shapes.
    ''' If all handlers return False (do not cancel), the grouping is performed.
    ''' </returns>
    ''' <remarks>Returns True if the grouping would cause improper database syncing.</remarks>
    Protected Overridable Function HandleQueryCancelGroup(ByVal vsoSelection As Visio.Selection) As Boolean
        Return False
    End Function

#End Region

#Region "Shape Event Handlers"

    ''' <summary>
    ''' Handles the BeforeShapeDelete event.
    ''' Occurs before a shape is deleted.
    ''' </summary>
    ''' <param name="vsoShape">The shape that is going to be deleted.</param>
    ''' <remarks>A Shape object can serve as the source object for the BeforeShapeDelete 
    ''' event if the shape's Type property is visTypeGroup(2) or visTypePage(1).
    ''' The BeforeSelectionDelete and BeforeShapeDelete events are similar in that they 
    ''' both fire before shape(s) are deleted. They differ in how they behave when a 
    ''' single operation deletes several shapes. Suppose a Cut operation deletes three 
    ''' shapes. The BeforeShapeDelete event fires three times and acts on each of the 
    ''' three objects. The BeforeSelectionDelete event fires once, and it acts on a 
    ''' Selection object in which the three shapes that you want to delete are selected.</remarks>
    Protected Overridable Sub HandleBeforeShapeDelete(ByVal vsoShape As Visio.Shape)
    End Sub
    ''' <summary>
    ''' Handles the BeforeShapeTextEdit event..
    ''' Occurs before a shape is opened for text editing in the user interface.
    ''' </summary>
    ''' <param name="vsoShape">The shape that is going to be opened for text editing.</param>
    Protected Overridable Sub HandleBeforeShapeTextEdit(ByVal vsoShape As Visio.Shape)
    End Sub
    ''' <summary>
    ''' Handles the ShapeAdded Event for Visio, which is fired any time a shape is added to the document.
    ''' Occurs after one or more shapes are added to a document.
    ''' We are concerned with paste functions. If a shape is copy/pasted, we want to clear the Shape_Key fields.
    ''' </summary>
    ''' <param name="vsoShape">The shape that was added.</param>
    ''' <remarks>
    ''' The SelectionAdded and ShapeAdded events are similar in that they both fire 
    ''' after shape(s) are created. They differ in how they behave when a single operation 
    ''' adds several shapes. Suppose a Paste operation creates three new shapes. 
    ''' The ShapeAdded event fires three times and acts on each of the three objects. 
    ''' The SelectionAdded event fires once, and it acts on a Selection object in which the three new shapes are selected.
    ''' To determine if a ShapeAdded event was triggered by a new shape 
    ''' or group of shapes being added to the page, by:
    ''' 1. a set of existing shapes being grouped
    ''' 2. by a paste action
    ''' You can use the Application.IsInScope property. 
    ''' If IsInScope returns True when passed visCmdObjectGroup, 
    ''' the ShapeAdded event was triggered by a grouping action. 
    ''' If IsInScope returns True when passed visCmdUFEditPaste or visCmdEditPasteSpecial, 
    ''' the ShapeAdded event was triggered by a paste operation. 
    ''' If IsInScope returns False when passed all of these arguments,
    ''' the event must have been triggered by new shapes being added to the page.</remarks>
    Protected Overridable Sub HandleShapeAdded(ByVal vsoShape As Visio.Shape)
    End Sub
    ''' <summary>
    ''' Handles the ShapeChanged event.
    ''' Occurs after a property of a shape that is not stored in a cell is changed in a document.
    ''' </summary>
    ''' <param name="vsoShape">The shape whose property changed.</param>
    ''' <param name="moreInformation">A string containing information specific to the changed properties.</param>
    ''' <remarks>To determine which properties have changed when ShapeChanged fires, use the EventInfo property. 
    ''' The string returned by the EventInfo property contains a list of substrings that identify the properties that changed.
    ''' Changes to the following shape properties cause the ShapeChanged event to fire:
    ''' Name (the EventInfo property contains "/name")
    ''' Data1 (the EventInfo property contains "/data1")
    ''' Data2 (the EventInfo property contains "/data2")
    ''' Data3 (the EventInfo property contains "/data3")
    ''' UniqueID (the EventInfo property contains "/uniqueid")
    ''' If you are handling this event from a program that receives a notification 
    ''' over a connection that was created by using AddAdvise, the varMoreInfo argument 
    ''' to VisEventProc contains the string: "/doc=1 /page=1 /shape=Sheet.3"
    ''' </remarks>
    Protected Overridable Sub HandleShapeChanged(ByVal vsoShape As Visio.Shape, _
                                  ByVal moreInformation As String)
    End Sub
    ''' <summary>
    ''' Handles the ShapeExitedTextEdit event.
    ''' Occurs after a shape is no longer open for interactive text editing.
    ''' </summary>
    ''' <param name="vsoShape">The shape that was closed for text editing.</param>
    Protected Overridable Sub HandleShapeExitedTextEdit(ByVal vsoShape As Visio.Shape)
    End Sub
    ''' <summary>
    ''' Handles the ShapeParentChanged event.
    ''' Occurs after shapes are grouped or a group is ungrouped.
    ''' </summary>
    ''' <param name="vsoShape">The shape whose parent changed.</param>
    Protected Overridable Sub HandleShapeParentChanged(ByVal vsoShape As Visio.Shape)
    End Sub
    ''' <summary>
    ''' Handles the ShapesDeleted event.
    ''' Occurs when shapes are deleted.
    ''' </summary>
    ''' <param name="vsoShape">The shape that was deleted.</param>
    Protected Overridable Sub HandleShapesDeleted(ByVal vsoShape As Visio.Shape)
    End Sub
    ''' <summary>
    ''' Handles the TextChanged event.
    ''' Occurs after the text of a shape is changed in a document.
    ''' </summary>
    ''' <param name="vsoShape">The shape whose text changed.</param>
    ''' <remarks>
    ''' The TextChanged event is fired when the raw text of a shape changes, 
    ''' such as when the characters Microsoft Office Visio stores for the shape change. 
    ''' If a shape's characters change because a user is typing, the TextChanged event 
    ''' does not fire until the text editing session terminates.
    ''' When a field is added to or removed from a shape's text, its raw text changes; 
    ''' hence, a TextChanged event fires. However, no TextChanged event fires when the 
    ''' text in a field changes. For example, a shape has a text field that shows its 
    ''' width. A TextChanged event does not fire when the shape's width changes, because 
    ''' the raw text stored for the shape has not changed, even though the apparent 
    ''' (expanded) text of the shape does change. Use the CellChanged event for one of 
    ''' the cells in the Text Fields section to detect when the text in a text field 
    ''' changes. To access a shape's raw text, use the Text property. To access the 
    ''' text of a shape in which text fields have been expanded, use the Characters.Text 
    ''' property. You can determine the location and properties of text fields in a 
    ''' shape's text by using the Shape.Characters object.
    ''' </remarks>
    Protected Overridable Sub HandleTextChanged(ByVal vsoShape As Visio.Shape)
    End Sub

    'Pro Edition Only!
    ''' <summary>
    ''' Handles the ShapeDataGraphicChanged event.
    ''' Occurs after a data graphic is applied to or deleted from a shape.
    ''' </summary>
    ''' <param name="vsoShape">	The shape to which the data graphic was applied or from which it was deleted.</param>
    ''' <remarks>This Visio object or member is available only to licensed users of Microsoft Office Visio Professional 2007.
    ''' A data graphic is a Master object of type visTypeDataGraphic. When the same master that is already applied to a shape is 
    ''' re-applied to the shape, the ShapeDataGraphicChanged event does not fire, even if the master has been modified since it 
    ''' was applied originally. If, however, a different data graphic master is applied to the shape, the event does fire.</remarks>
    Protected Overridable Sub HandleShapeDataGraphicChanged(ByVal vsoShape As Visio.Shape)
    End Sub
    ''' <summary>
    ''' Handles the shape link added.
    ''' Occurs after a shape is linked to a data row.
    ''' </summary>
    ''' <param name="vsoShape">The shape that is linked to data.</param>
    ''' <param name="DataRecordsetID">The ID of the data recordset containing the data row linked to the shape.</param>
    ''' <param name="DataRowID">The ID of the data row linked to the shape.</param>
    ''' <remarks>This Visio object or member is available only to licensed users of Microsoft Office Visio Professional 2007.
    ''' The ShapeLinkAdded event is one of a group of events for which the EventInfo property of the Application object contains 
    ''' extra information. When the ShapeLinkAdded event is fired, the EventInfo property returns the following string:
    ''' /DataRecordsetID = n /DataRowID = m 
    ''' where n and m represent the IDs of the data recordset and data row, respectively, associated with the event.</remarks>
    Protected Overridable Sub HandleShapeLinkAdded(ByVal vsoShape As Visio.Shape, _
                                    ByVal DataRecordsetID As Long, _
                                    ByVal DataRowID As Long)
    End Sub
    ''' <summary>
    ''' Handles the shape link deleted.
    ''' Occurs after the link between a shape and a data row is deleted.
    ''' </summary>
    ''' <param name="vsoShape">The shape whose link to a data row was broken.</param>
    ''' <param name="DataRecordsetID">The ID of the data recordset containing the data row that was linked to the shape.</param>
    ''' <param name="DataRowID">The ID of the data row that was linked to the shape.</param>
    ''' <remarks>This Visio object or member is available only to licensed users of Microsoft Office Visio Professional 2007.
    ''' The ShapeLinkDeleted event is one of a group of events for which the EventInfo property of the Application object contains 
    ''' extra information. When the ShapeLinkDeleted event is fired, the EventInfo property returns the following string:
    ''' /DataRecordsetID = n /DataRowID = m 
    ''' where n and m represent the IDs of the data recordset and data row, respectively, associated with the event.</remarks>
    Protected Overridable Sub HandleShapeLinkDeleted(ByVal vsoShape As Visio.Shape, _
                                      ByVal DataRecordsetID As Long, _
                                      ByVal DataRowID As Long)
    End Sub

#End Region

#Region "Cell Event Handlers"

    ''' <summary>
    ''' Handles the CellChanged event.
    ''' Occurs after the value changes in a cell in a document.
    ''' </summary>
    ''' <param name="vsoCell">The cell whose value has changed.</param>
    Protected Overridable Sub HandleCellChanged(ByVal vsoCell As Visio.Cell)
    End Sub
    ''' <summary>
    ''' Handles the FormulaChanged event.
    ''' Occurs after a formula changes in a cell in the object that receives the event.
    ''' </summary>
    ''' <param name="vsoCell">The cell whose forumla has changed.</param>
    Protected Overridable Sub HandleFormulaChanged(ByVal vsoCell As Visio.Cell)
    End Sub

#End Region

#Region "Connects Event Handlers"

    ''' <summary>
    ''' Handles the ConnectionsAdded event.
    ''' Occurs after connections have been established between shapes.
    ''' </summary>
    ''' <param name="vsoConnects">The connections that were established.</param>
    Protected Overridable Sub HandleConnectionsAdded(ByVal vsoConnects As Visio.Connects)
    End Sub
    ''' <summary>
    ''' Handles the ConnectionsDeleted event.
    ''' Occurs after connections between shapes have been removed.
    ''' </summary>
    ''' <param name="vsoConnects">The connections that were removed.</param>
    Protected Overridable Sub HandleConnectionsDeleted(ByVal vsoConnects As Visio.Connects)
    End Sub

#End Region

#Region "Style Event Handlers"

    ''' <summary>
    ''' Handles the BeforeStyleDelete event.
    ''' Occurs before a style is deleted.
    ''' </summary>
    ''' <param name="vsoStyle">The style that is going to be deleted.</param>
    Protected Overridable Sub HandleBeforeStyleDelete(ByVal vsoStyle As Visio.Style)
    End Sub
    ''' <summary>
    ''' Handles the StyleAdded event.
    ''' Occurs after a new style is added to a document.
    ''' </summary>
    ''' <param name="vsoStyle">The style that was added to the document.</param>
    Protected Overridable Sub HandleStyleAdded(ByVal vsoStyle As Visio.Style)
    End Sub
    ''' <summary>
    ''' Handles the StyleChanged event.
    ''' Occurs after the name of a style is changed or a change to the style propagates to objects to which the style is applied.
    ''' </summary>
    ''' <param name="vsoStyle">	The style that changed.</param>
    Protected Overridable Sub HandleStyleChanged(ByVal vsoStyle As Visio.Style)
    End Sub
    ''' <summary>
    ''' Handles the StyleDeleteCanceled event.
    ''' Occurs after an event handler has returned True (cancel) to a QueryCancelStyleDelete event.
    ''' </summary>
    ''' <param name="vsoStyle">	The style that was going to be deleted.</param>
    Protected Overridable Sub HandleStyleDeleteCanceled(ByVal vsoStyle As Visio.Style)
    End Sub
    ''' <summary>
    ''' Handles the QueryCancelStyleDelete event.
    ''' Occurs before the application deletes a style in response to a user action in the interface. 
    ''' If any event handler returns True, the operation is canceled.
    ''' </summary>
    ''' <param name="vsoStyle">	The style that is going to be deleted.</param>
    ''' <returns>
    ''' True (cancel), the instance will fire 
    ''' StyleDeleteCanceled and not delete the style.
    ''' False (don't cancel), the instance will fire 
    ''' BeforeStyleDelete and then delete the style.
    ''' </returns>
    Protected Overridable Function HandleQueryCancelStyleDelete(ByVal vsoStyle As Visio.Style) As Boolean
    End Function

#End Region

#Region "Window Event Handlers"

    ''' <summary>
    ''' Handles the BeforeWindowClosed event.
    ''' Occurs before a window is closed.
    ''' </summary>
    ''' <param name="vsoWindow">The window that is going to be closed.</param>
    Protected Overridable Sub HandleBeforeWindowClosed(ByVal vsoWindow As Visio.Window)
    End Sub
    ''' <summary>
    ''' Handles the BeforeWindowPageTurn event.
    ''' Occurs before a window is about to show a different page.
    ''' </summary>
    ''' <param name="vsoWindow">The window that is going to show a different page.</param>
    Protected Overridable Sub HandleBeforeWindowPageTurn(ByVal vsoWindow As Visio.Window)
    End Sub
    ''' <summary>
    ''' Handles the SelectionChanged event.
    ''' Occurs after a set of shapes selected in a window changes.
    ''' </summary>
    ''' <param name="vsoWindow">The window in which the selection changed.</param>
    Protected Overridable Sub HandleSelectionChanged(ByVal vsoWindow As Visio.Window)
    End Sub
    ''' <summary>
    ''' Handles the BeforeWindowSelDelete event.
    ''' Occurs before the shapes in the selection of a window are deleted.
    ''' </summary>
    ''' <param name="vsoWindow">The window that contains the selection that is going to be deleted.</param>
    Protected Overridable Sub HandleBeforeWindowSelDelete(ByVal vsoWindow As Visio.Window)
    End Sub
    ''' <summary>
    ''' Handles the WindowOpened event.
    ''' Occurs after a window is opened.
    ''' </summary>
    ''' <param name="vsoWindow">The window that opened.</param>
    Protected Overridable Sub HandleWindowOpened(ByVal vsoWindow As Visio.Window)
    End Sub
    ''' <summary>
    ''' Handles the WindowChanged event.
    ''' Occurs when the size or position of a window changes.
    ''' </summary>
    ''' <param name="vsoWindow">The window whose size or position has changed.</param>
    Protected Overridable Sub HandleWindowChanged(ByVal vsoWindow As Visio.Window)
    End Sub
    ''' <summary>
    ''' Handles the WindowTurnedToPage event.
    ''' Occurs after a window shows a different page.
    ''' </summary>
    ''' <param name="vsoWindow">The window that shows a different page.</param>
    Protected Overridable Sub HandleWindowTurnedToPage(ByVal vsoWindow As Visio.Window)
    End Sub
    ''' <summary>
    ''' Handles the WindowCloseCanceled.
    ''' Occurs after an event handler has returned True (cancel) to a QueryCancelWindowClose event.
    ''' </summary>
    ''' <param name="vsoWindow">	The window that was going to be closed.</param>
    Protected Overridable Sub HandleWindowCloseCanceled(ByVal vsoWindow As Visio.Window)
    End Sub
    ''' <summary>
    ''' Handles the WindowActivated event.
    ''' Occurs after the active window changes in a Microsoft Office Visio instance.
    ''' </summary>
    ''' <param name="vsoWindow">The window that was activated.</param>
    Protected Overridable Sub HandleWindowActivated(ByVal vsoWindow As Visio.Window)
    End Sub
    ''' <summary>
    ''' Handles the ViewChanged event.
    ''' Occurs when the zoom level or scroll position of a drawing window changes.
    ''' </summary>
    ''' <param name="vsoWindow">The window whose zoom level or scroll position changed.</param>
    ''' <remarks>This event fires whenever the zoom level or scroll position of a Window object of the type visDrawing changes.</remarks>
    Protected Overridable Sub HandleViewChanged(ByVal vsoWindow As Visio.Window)
    End Sub
    ''' <summary>
    ''' Handles the QueryCancelWindowClose event.
    ''' Occurs before the application closes a window in response to a user action in the interface. 
    ''' If any event handler returns True, the operation is canceled.
    ''' </summary>
    ''' <param name="vsoWindow">The window that is going to be closed.</param>
    ''' <returns>
    ''' True (cancel), the instance will fire 
    ''' WindowCloseCanceled and not close the window.
    ''' False (don't cancel), the instance will fire 
    ''' BeforeWindowClosed and then close the window.
    ''' </returns>
    Protected Overridable Function HandleQueryCancelWindowClose(ByVal vsoWindow As Visio.Window) As Boolean
        Return False
    End Function

#End Region

#Region "Application Event Handlers"

    ''' <summary>
    ''' Handles the AfterModal event.
    ''' Occurs after the Microsoft Office Visio instance leaves a modal state.
    ''' </summary>
    ''' <param name="vsoApp">The instance that is no longer modal.</param>
    ''' <remarks>Visio becomes modal when it displays a dialog box. A modal instance of Visio does not handle Automation calls. 
    ''' The BeforeModal event indicates that the instance is about to become modal, and the AfterModal event indicates 
    ''' that the instance is no longer modal.</remarks>
    Protected Overridable Sub HandleAfterModal(ByVal vsoApp As Visio.Application)
    End Sub
    ''' <summary>
    ''' Handles the AfterResume event.
    ''' Occurs when the operating system resumes normal operation after having been suspended.
    ''' </summary>
    ''' <param name="vsoApp">The instance of Visio that resumes after the operating system resumes normal operation.</param>
    ''' <remarks>You can use the AfterResume event to reopen any network files that you may have closed in response to the BeforeSuspend event.
    ''' If your solution runs outside the Microsoft Office Visio process, you cannot be assured of receiving this event. 
    ''' For this reason, you should monitor window messages in your program.</remarks>
    Protected Overridable Sub HandleAfterResume(ByVal vsoApp As Visio.Application)
    End Sub
    ''' <summary>
    ''' Handles the AppActivated event.
    ''' Occurs after a Microsoft Office Visio instance becomes active.
    ''' </summary>
    ''' <param name="vsoApp">The instance of Visio that becomes the active application.</param>
    ''' <remarks>The AppActivated event indicates that an instance of Visio has become the active 
    ''' application on the Microsoft Windows desktop. The AppActivated event is different from the AppObjectActivated event, 
    ''' which occurs after an instance of Visio becomes active—the instance of Visio that is retrieved by the GetObject 
    ''' function in a Microsoft Visual Basic program.</remarks>
    Protected Overridable Sub HandleAppActivated(ByVal vsoApp As Visio.Application)
    End Sub
    ''' <summary>
    ''' Handles the app deactivated.
    ''' Occurs after a Microsoft Office Visio instance becomes inactive.
    ''' </summary>
    ''' <param name="vsoApp">The instance of Visio that is no longer the active application.</param>
    ''' <remarks>The AppDeactivated event indicates that an instance of Visio is no longer the active 
    ''' application on the Microsoft Windows desktop. The AppDeactivated event is different from the 
    ''' AppObjectDeactivated event, which occurs after an instance of Visio ceases to be the active 
    ''' instance—the instance of Visio that is retrieved by the GetObject function in a 
    ''' Microsoft Visual Basic program.</remarks>
    Protected Overridable Sub HandleAppDeactivated(ByVal vsoApp As Visio.Application)
    End Sub
    ''' <summary>
    ''' Handles the AppObjActivated event.
    ''' Occurs after a Microsoft Office Visio instance becomes active.
    ''' </summary>
    ''' <param name="vsoApp">The instance of Visio that has become the active instance.</param>
    ''' <remarks>The AppObjActivated event indicates that an instance of Visio has become active—the 
    ''' instance of Visio that is retrieved by the GetObject function in a Microsoft Visual Basic program. 
    ''' The AppObjActivated event is different from the AppActivated event, which occurs after an instance 
    ''' of Visio becomes the active application on the Microsoft Windows desktop.</remarks>
    Protected Overridable Sub HandleAppObjActivated(ByVal vsoApp As Visio.Application)
    End Sub
    ''' <summary>
    ''' Handles the AppObjDeactivated event.
    ''' Occurs after a Microsoft Office Visio instance becomes inactive.
    ''' </summary>
    ''' <param name="vsoApp">The instance of Visio that is no longer the active instance.</param>
    ''' <remarks>The AppObjDeactivated event indicates that the instance of Visio is no longer the 
    ''' active instance of Visio—the instance of Visio that is retrieved by the GetObject function in a 
    ''' Microsoft Visual Basic program. The AppObjDeactivated event is different from the AppDeactivated event, 
    ''' which occurs after an instance of Visio is no longer the active instance on the Microsoft Windows desktop.</remarks>
    Protected Overridable Sub HandleAppObjDeactivated(ByVal vsoApp As Visio.Application)
    End Sub
    ''' <summary>
    ''' Handles the BeforeModal event.
    ''' Occurs before a Microsoft Office Visio instance enters a modal state.
    ''' </summary>
    ''' <param name="vsoApp">The instance of Visio that is going to enter a modal state.</param>
    ''' <remarks>Visio becomes modal when it displays a dialog box. A modal instance of Visio does not handle 
    ''' Automation calls. The BeforeModal event indicates that an instance is about to become modal, and the 
    ''' AfterModal event indicates that the instance is no longer modal.</remarks>
    Protected Overridable Sub HandleBeforeModal(ByVal vsoApp As Visio.Application)
    End Sub
    ''' <summary>
    ''' Handles the BeforeQuit event.
    ''' Occurs before a Microsoft Office Visio instance terminates.
    ''' </summary>
    ''' <param name="vsoApp">The instance of Visio that is going to terminate</param>
    ''' <remarks>When programming with Microsoft Visual Basic, use the BeforeDocumentClose event instead 
    ''' of the BeforeQuit event. The code in a Visual Basic project of a Visio document never has the chance 
    ''' to respond to the BeforeQuit event because the project is a property of a document, and all documents 
    ''' are closed before the BeforeQuit event notification is sent.</remarks>
    Protected Overridable Sub HandleBeforeQuit(ByVal vsoApp As Visio.Application)
    End Sub
    ''' <summary>
    ''' Handles the BeforeSuspend event.
    ''' Occurs before the operating system enters a suspended state.
    ''' </summary>
    ''' <param name="vsoApp">The instance of Visio that is going to be suspended.</param>
    ''' <remarks>Client programs should close any open network files when this event is fired.
    ''' If your solution runs outside the Microsoft Office Visio process, you cannot be assured of 
    ''' receiving this event. For this reason, you should monitor window messages in your program.</remarks>
    Protected Overridable Sub HandleBeforeSuspend(ByVal vsoApp As Visio.Application)
    End Sub
    ''' <summary>
    ''' Handles the EnterScope event.
    ''' </summary>
    ''' <param name="vsoApp">The instance of Microsoft Office Visio that contains the scope.</param>
    ''' <param name="moreInformation">a string formatted as follows: 
    ''' [&lt;nScopeID&gt;;&lt;bErrOrCancelled&gt;;&lt;bstrDescription&gt;;&lt;nHwndContext&gt;], 
    ''' where nHwndContext is the window handle (HWND) of the window that is the context for the command. nHwndContext could be 0.
    ''' For EnterScope, bErrOrCancelled always equals zero.</param>
    ''' <remarks>The nScopeID value returned in the case of a Visio operation is the equivalent of the command-related 
    ''' constants that begin with visCmd*.
    ''' "nScopeID" = A language-independent number that describes the operation that just 
    ''' ended or the scope ID returned by the BeginUndoScope method.
    ''' "bstrDescription" =A textual description of the operation that changes in different 
    ''' language versions. Contains the user interface description of a Visio operation or the description 
    ''' passed to the BeginUndoScope method.</remarks>
    Protected Overridable Sub HandleEnterScope(ByVal vsoApp As Visio.Application, _
                                ByVal moreInformation As String)
        'varMoreInfo argument to VisEventProc contains a string formatted as follows: 
        '[<nScopeID>;<bErrOrCancelled>;<bstrDescription>;<nHwndContext>], 
        'where nHwndContext is the window handle (HWND) of the window that is the context for the command. nHwndContext could be 0. 
    End Sub
    ''' <summary>
    ''' Handles the ExitScope event.
    ''' Queued when an internal command ends, or when an Automation client exits a scope by using the EndUndoScope method.
    ''' </summary>
    ''' <param name="vsoApp">The instance of Microsoft Office Visio that contains the scope.</param>
    ''' <param name="moreInformation">a string formatted as follows: 
    ''' [&lt;nScopeID&gt;;&lt;bErrOrCancelled&gt;;&lt;bstrDescription&gt;;&lt;nHwndContext&gt;], 
    ''' where nHwndContext is the window handle (HWND) of the window that is the context for the command. nHwndContext could be 0.</param>
    ''' <remarks>The nScopeID value returned in the case of a Visio operation is the equivalent of the command 
    ''' related constants that begin with visCmd*.
    ''' For ExitScope, bErrOrCancelled is non-zero if the operation failed or was canceled.</remarks>
    Protected Overridable Sub HandleExitScope(ByVal vsoApp As Visio.Application, _
                               ByVal moreInformation As String)
    End Sub
    ''' <summary>
    ''' Handles the MarkerEvent event.
    ''' Caused by calling the QueueMarkerEvent method.
    ''' This allows me to create my own events which are fired within the Visio Engine.
    ''' </summary>
    ''' <param name="vsoApp">The app.</param>
    ''' <param name="eventSequenceNumber">The sequence number of the event.</param>
    ''' <remarks>Unlike other events that Visio fires, the MarkerEvent event is fired by a client program. 
    ''' A client program receives the MarkerEvent event only if the client program called the QueueMarkerEvent method.
    ''' By using the MarkerEvent event in conjunction with the QueueMarkerEvent method, a client program can queue an 
    ''' event to itself. The client program receives the MarkerEvent event after Visio fires all the events present in 
    ''' its event queue at the time of the QueueMarkerEvent call.
    ''' The MarkerEvent event passes both the context string that was passed by the QueueMarkerEvent method and the 
    ''' sequence number of the MarkerEvent event to the MarkerEvent event handler. Either of these values can be used 
    ''' to correlate QueueMarkerEvent calls with MarkerEvent events. In this way, a client program can distinguish events 
    ''' it caused from those it did not cause.
    ''' For example, a client program that changes the values of Visio cells may only want to respond to the CellChanged 
    ''' events that it did not cause. The client program can first call the QueueMarkerEvent method and pass a context 
    ''' string for later use to bracket the scope of its processing:
    ''' Then, in the MarkerEvent event handler, the client program could use the context string passed 
    ''' to the QueueMarkerEvent method to identify the CellChanged events that it caused: 
    ''' The EventInfo property returns ContextString as described above. The varMoreInfo argument to VisEventProc will be empty.</remarks>
    Protected Overridable Sub HandleMarkerEvent(ByVal vsoApp As Visio.Application, _
                                                       ByVal eventSequenceNumber As Integer, _
                                                       ByVal contextString As String)
    End Sub
    ''' <summary>
    ''' Handles the MustFlushScopeBeginning event.
    ''' Occurs before the Microsoft Office Visio instance is forced to flush its event queue.
    ''' </summary>
    ''' <param name="vsoApp">The instance of Visio that is forced to flush its event queue.</param>
    ''' <remarks>This event, along with the MustFlushScopeEnded event, can be used to identify whether an event is being 
    ''' fired because Visio is forced to flush its event queue.
    ''' Visio maintains a queue of pending events that it attempts to fire at discrete moments when it is able to process 
    ''' arbitrary requests (callbacks) from event handlers. Occasionally, Visio is forced to flush its event queue when it 
    ''' is not prepared to handle arbitrary requests. When this occurs, Visio first fires a MustFlushScopeBeginning event, 
    ''' and then it fires the events that are presently in its event queue. After firing all pending events, Visio fires 
    ''' the MustFlushScopeEnded event. After Visio has fired the MustFlushScopeBeginning event, client programs should 
    ''' not call Visio methods that have side effects until the MustFlushScopeEnded event is received. A client can 
    ''' perform arbitrary queries of Visio objects when Visio is between the MustFlushScopeBeginning event and MustFlushScopeEnded event, 
    ''' but operations that cause side effects may fail. Visio performs a forced flush of its event queue immediately prior to 
    ''' firing a "before" event such as BeforeDocumentClose or BeforeShapeDelete because queued events may apply to objects
    ''' that are about to close or be deleted. Using the BeforeDocumentClose event as an example, there can be queued events 
    ''' that apply to a shape object in the document that is being closed. So, before the document closes, Visio fires all the 
    ''' events in its event queue.
    ''' When a shape is deleted, events are fired in the following sequence:
    ''' 1. MustFlushScopeBeginning event 
    ''' Client should not call methods that have side effects.
    ''' 2. There are zero (0) or more events in the event queue.
    ''' 3. BeforeShapeDelete event
    ''' Shape is viable, but Visio is going to delete it.
    ''' 4. MustFlushScopeEnded event
    ''' Client can resume invoking methods that have side effects.
    ''' 5. ShapesDeleted event
    ''' Shape has been deleted.
    ''' 6. NoEventsPending event
    ''' No events remain to be fired.
    ''' An event is fired both before (BeforeShapeDeleted event) and after (ShapesDeleted event) the shape is deleted. 
    ''' If a program monitoring these events requires that additional shapes be deleted in response to the initial shape deletion, 
    ''' it should do so in the ShapesDeleted event handler, not the BeforeShapeDeleted event handler. The BeforeShapeDeleted event 
    ''' is inside the scope of the MustFlushScopeBeginning event and the MustFlushScopeEnded event, while the ShapesDeleted event is not.
    ''' The sequence number of a MustFlushScopeBeginning event may be higher than the sequence number of events the client sees after 
    ''' it has received the MustFlushScopeBeginning event because Visio assigns sequence numbers to events as they occur. Any events 
    ''' that were queued when the forced flush began have a lower sequence number than the MustFlushScopeBeginning event, even though 
    ''' the MustFlushScopeBeginning event fires first.</remarks>
    Protected Overridable Sub HandleMustFlushScopeBeginning(ByVal vsoApp As Visio.Application)
    End Sub
    ''' <summary>
    ''' Handles the MustFlushScopeEnded event.
    ''' Occurs after the Microsoft Office Visio instance is forced to flush its event queue.
    ''' </summary>
    ''' <param name="vsoApp">The instance of Visio that is forced to flush its event queue.</param>
    ''' <remarks>This event, along with the MustFlushScopeBeginning event, can be used to identify whether 
    ''' an event is being fired because Visio is forced to flush its event queue.</remarks>
    Protected Overridable Sub HandleMustFlushScopeEnded(ByVal vsoApp As Visio.Application)
    End Sub
    ''' <summary>
    ''' Handles the NoEventsPending event.
    ''' Occurs after the Microsoft Ofice Visio instance flushes its event queue.
    ''' </summary>
    ''' <param name="vsoApp">The instance of Visio that flushed its event queue.</param>
    ''' <remarks>Visio maintains a queue of events and fires them at discrete moments. Immediately 
    ''' after Visio fires the last event in its event queue, it fires a NoEventsPending event.
    ''' A client program can use the NoEventsPending event as a signal that Visio has completed a burst of activity. 
    ''' For example, a client program may want to react to changes in a shape's geometry. A single user action performed 
    ''' on the shape can generate several CellChanged events. The client program could record selected information for 
    ''' each CellChanged event and perform its processing after it receives the NoEventsPending event.
    ''' Visio fires the NoEventsPending event only if at least one of the events in the queue is being listened to. 
    ''' If no program is listening for any of the queued events, the NoEventsPending event does not fire. If your program 
    ''' is only listening to the NoEventsPending event, it does not fire unless another program is listening for some of the queued events.</remarks>
    Protected Overridable Sub HandleNoEventsPending(ByVal vsoApp As Visio.Application)
    End Sub
    ''' <summary>
    ''' Handles the OnKeystrokeMessageForAddon event.
    ''' Occurs when Microsoft Office Visio receives a keystroke message from Microsoft Windows 
    ''' that is targeted at an add-on window or child of an add-on window.
    ''' </summary>
    ''' <param name="vsoApp">The message that Visio receives.</param>
    ''' <returns>Returns True to indicate that the message was handled by the add-on. Otherwise, returns False.</returns>
    ''' <remarks>
    ''' The OnKeystrokeMessageForAddon event enables add-ons to intercept and process accelerator and keystroke 
    ''' messages directed at their own add-on windows and child windows of their add-on windows. Only add-on windows 
    ''' created using the Add method will source this event.
    ''' For this event to fire, the add-on window or one of its child windows must have keystroke focus and the 
    ''' Visio message loop must receive the keystroke message. This event does not fire if the message loop associated 
    ''' with an add-on is handling messages instead of Visio.
    ''' Visio fires the OnKeystrokeMessageForAddon event when it receives messages in the following range:
    ''' WM_KEYDOWN - 0x0100
    ''' WM_KEYUP - 0x0101
    ''' WM_CHAR - 0x0102
    ''' WM_DEADCHAR - 0x0103
    ''' WM_SYSKEYDOWN - 0x0104
    ''' WM_SYSKEYUP - 0x0105
    ''' WM_SYSCHAR - 0x0106
    ''' WM_SYSDEADCHAR - 0x0107
    ''' The MSGWrap object, passed to the event handler when the OnKeystrokeMessageForAddon event fires, wraps the 
    ''' Microsoft Windows MSG structure, which contains message data. See the MSGWrap object for more information, 
    ''' or refer to your Windows documentation.</remarks>
    Protected Overridable Function HandleOnKeystrokeMessageForAddon(ByVal vsoApp As Visio.Application) As Boolean
        Return False
    End Function
    ''' <summary>
    ''' Handles the QueryCancelQuit event.
    ''' Occurs before the application terminates in response to a user action in the interface. 
    ''' If any event handler returns True, the operation is canceled.
    ''' </summary>
    ''' <param name="vsoApp">The instance of Microsoft Office Visio that is going to terminate.</param>
    ''' <returns>
    ''' True (cancel), the instance will fire QuitCanceled and not terminate.
    ''' False (don't cancel), the instance will fire BeforeQuit and then terminate.
    ''' </returns>
    Protected Overridable Function HandleQueryCancelQuit(ByVal vsoApp As Visio.Application) As Boolean
        Return False
    End Function
    ''' <summary>
    ''' Handles the QueryCancelSuspend event.
    ''' Occurs before the operating system enters a suspended state. 
    ''' If any event handler returns True, the Microsoft Office Visio instance will deny the operating system's request.
    ''' </summary>
    ''' <param name="vsoApp">The instance of Visio that responds to the operating system request.</param>
    ''' <remarks>You will typically respond False and allow the operating system to enter a suspended state. 
    ''' If you have open network files, you can close them when you receive the BeforeSuspend event. 
    ''' If you have open network files that you cannot close, you can return True and Visio will deny the operating system's request.</remarks>
    ''' <returns>
    ''' True (cancel), the instance will fire SuspendCanceled and not enter a suspended state.
    ''' False (don't cancel), the instance will fire BeforeSuspend and then enter a suspended state.
    ''' </returns>
    Protected Overridable Function HandleQueryCancelSuspend(ByVal vsoApp As Visio.Application) As Boolean
        Return False
    End Function
    ''' <summary>
    ''' Handles the QuitCanceled event.
    ''' Occurs after an event handler has returned True (cancel) to a QueryCancelQuit event.
    ''' </summary>
    ''' <param name="vsoApp">The instance of Microsoft Office Visio that was going to be terminated.</param>
    Protected Overridable Sub HandleQuitCanceled(ByVal vsoApp As Visio.Application)
    End Sub
    ''' <summary>
    ''' Handles the SuspendCanceled event.
    ''' Occurs after an event handler has returned True (cancel) to a QueryCancelSuspend event.
    ''' </summary>
    ''' <param name="vsoApp">The instance of Microsoft Office Visio that was going to be suspended.</param>
    ''' <remarks>If your solution runs outside the Visio process, you cannot be assured of receiving this event. 
    ''' For this reason, you should monitor window messages in your program.</remarks>
    Protected Overridable Sub HandleSuspendCanceled(ByVal vsoApp As Visio.Application)
    End Sub
    ''' <summary>
    ''' Handles the VisioIsIdle event.
    ''' Occurs after the application empties its message queue.
    ''' </summary>
    ''' <param name="vsoApp">The instance of Microsoft Office Visio that emptied its message queue.</param>
    ''' <remarks>Visio continually processes messages in its message queue. When its message queue is empty:
    '''    1. Visio performs its own idle-time processing.
    '''    2. Visio tells Microsoft Visual Basic for Applications to perform its idle-time processing.
    '''    3. If the message queue is still empty, Visio fires the VisioIsIdle event.
    '''    4. If the message queue is still empty, Visio calls WaitMessage, which is a call to Microsoft Windows 
    '''       that doesn't return until a new message gets added to the Visio message queue.
    ''' A client program can use the VisioIsIdle event as a signal to perform its own background processing.
    ''' The VisioIsIdle event is not the equivalent of a standard timer event. Client programs that need to be 
    ''' called on a periodic basis should use standard timer techniques, because the duration in which Visio is 
    ''' idle (calls WaitMessage) is unpredictable. For client programs that are only monitoring Visio activity, 
    ''' however, the VisioIsIdle event can be sufficient, because until WaitMessage returns to Visio, there cannot 
    ''' have been any Visio activity since the VisioIsIdle event was last fired.</remarks>
    Protected Overridable Sub HandleVisioIsIdle(ByVal vsoApp As Visio.Application)
    End Sub

#End Region

#Region "Mouse Event Handlers"

    ''' <summary>
    ''' Handles the MouseDown event.
    ''' Occurs when a mouse button is clicked.
    ''' </summary>
    ''' <param name="vsoMouseEvent">The MouseEvent object which triggered the handler.</param>
    ''' <returns>True to cancel the mouse click. False to process the mouse click.</returns>
    ''' <remarks>Unlike some other Visio events, MouseDown does not have the prefix "Query," 
    ''' but it is nevertheless a query event. That is, you can cancel processing the message 
    ''' sent by MouseDown, either by setting CancelDefault to True, or, if you are using theVisEventProc
    ''' method to handle the event, by returning True.
    ''' Possible values for Button are shown in the following table, and are declared in VisKeyButtonFlags in the Visio type library. 
    ''' visMouseLeft 1
    ''' visMouseMiddle 16
    ''' visMouseRight 2
    ''' Possible values for KeyButtonState can be a combination of the values shown in the following table, 
    ''' which are declared in VisKeyButtonFlags in the Visio type library.
    '''  For example, if KeyButtonState returns 9, it indicates that the user clicked the left mouse button while pressing CTRL.
    ''' visKeyControl 8
    ''' visKeyShift 4
    ''' visMouseLeft 1
    ''' visMouseMiddle 16
    ''' visMouseRight 2</remarks>
    Protected Overridable Function HandleMouseDown(ByVal vsoMouseEvent As Visio.MouseEvent) As Boolean
    End Function
    ''' <summary>
    ''' Handles the MouseMove event.
    ''' Occurs when the mouse is moved.
    ''' </summary>
    ''' <param name="vsoMouseEvent">The MouseEvent object which triggered the handler.</param>
    ''' <returns>True to cancel the mouse move. False to process the mouse move message.</returns>
    ''' <remarks>Unlike some other Visio events, MouseMove does not have the prefix "Query," 
    ''' but it is nevertheless a query event. That is, you can cancel processing the message 
    ''' sent by MouseMove, either by setting CancelDefault to True, or, if you are using theVisEventProc
    ''' method to handle the event, by returning True.
    ''' Possible values for Button are shown in the following table, and are declared in VisKeyButtonFlags in the Visio type library. 
    ''' visMouseLeft 1
    ''' visMouseMiddle 16
    ''' visMouseRight 2
    ''' Possible values for KeyButtonState can be a combination of the values shown in the following table, 
    ''' which are declared in VisKeyButtonFlags in the Visio type library.
    '''  For example, if KeyButtonState returns 9, it indicates that the user clicked the left mouse button while pressing CTRL.
    ''' visKeyControl 8
    ''' visKeyShift 4
    ''' visMouseLeft 1
    ''' visMouseMiddle 16
    ''' visMouseRight 2</remarks>
    Protected Overridable Function HandleMouseMove(ByVal vsoMouseEvent As Visio.MouseEvent) As Boolean
    End Function
    ''' <summary>
    ''' Handles the MouseUp event.
    ''' Occurs when a mouse button is released.
    ''' </summary>
    ''' <param name="vsoMouseEvent">The MouseEvent object which triggered the handler.</param>
    ''' <returns>True to cancel the mouse click release. False to process the mouse click release message.</returns>
    ''' <remarks>Unlike some other Visio events, MouseUp does not have the prefix "Query," 
    ''' but it is nevertheless a query event. That is, you can cancel processing the message 
    ''' sent by MouseUp, either by setting CancelDefault to True, or, if you are using theVisEventProc
    ''' method to handle the event, by returning True.
    ''' Possible values for Button are shown in the following table, and are declared in VisKeyButtonFlags in the Visio type library. 
    ''' visMouseLeft 1
    ''' visMouseMiddle 16
    ''' visMouseRight 2
    ''' Possible values for KeyButtonState can be a combination of the values shown in the following table, 
    ''' which are declared in VisKeyButtonFlags in the Visio type library.
    '''  For example, if KeyButtonState returns 9, it indicates that the user clicked the left mouse button while pressing CTRL.
    ''' visKeyControl 8
    ''' visKeyShift 4
    ''' visMouseLeft 1
    ''' visMouseMiddle 16
    ''' visMouseRight 2</remarks>
    Protected Overridable Function HandleMouseUp(ByVal vsoMouseEvent As Visio.MouseEvent) As Boolean
    End Function

#End Region

#Region "Keyboard Event Handlers"

    ''' <summary>
    ''' Handles the KeyDown event.
    ''' Occurs when a keyboard key is pressed.
    ''' </summary>
    ''' <param name="vsoKeyboardEvent">The KeyboardEvent Object which triggered the event..</param>
    ''' <returns>True to cancel the key press. False to process the key press.</returns>
    ''' <remarks>Possible values for KeyCode are declared in KeyCodeConstants in the Microsoft Visual Basic for Applications (VBA) library.
    ''' Possible values for KeyButtonState can be a combination of the values shown in the following table, which are declared in 
    ''' VisKeyButtonFlags in the Visio type library. For example, if KeyButtonState returns 12, it indicates that the user held down both 
    ''' SHIFT and CTRL.
    ''' visKeyControl 8
    ''' visKeyShift 4 
    ''' visMouseLeft 1
    ''' visMouseMiddle 16
    ''' visMouseRight 2
    ''' If you set CancelDefault to True, Visio does not process the message received when the mouse button is clicked.
    ''' Unlike some other Visio events, KeyDown does not have the prefix "Query," but it is still a query event. 
    ''' That is, you can cancel processing the message sent by KeyDown, either by setting CancelDefault to True, or, if you are using 
    ''' theVisEventProc method to handle the event, by returning True. 
    ''' Pressing an accelererator key combination, for example, CTRL + C, does not fire the KeyDown event.</remarks>
    Protected Overridable Function HandleKeyDown(ByVal vsoKeyboardEvent As Visio.KeyboardEvent) As Boolean
    End Function
    ''' <summary>
    ''' Handles the KeyPress event.
    ''' Occurs when a keyboard key is pressed.
    ''' </summary>
    ''' <param name="vsoKeyboardEvent">The KeyboardEvent Object which triggered the event..</param>
    ''' <returns>True to cancel the key press. False to process the key press.</returns>
    ''' <remarks>Possible values for KeyAscii are the ASCII codes. To see a list of these codes, search for "ASCII character codes" on MSDN.
    ''' If you set CancelDefault to True, Visio does not process the message received when the mouse button is clicked.
    ''' Unlike some other Visio events, KeyDown does not have the prefix "Query," but it is still a query event. 
    ''' That is, you can cancel processing the message sent by KeyDown, either by setting CancelDefault to True, or, if you are using 
    ''' theVisEventProc method to handle the event, by returning True. 
    ''' Pressing an accelererator key combination, for example, CTRL + C, does not fire the KeyDown event.</remarks>
    Protected Overridable Function HandleKeyPress(ByVal vsoKeyboardEvent As Visio.KeyboardEvent) As Boolean
    End Function
    ''' <summary>
    ''' Handles the KeyUp event.
    ''' Occurs when a keyboard key is released.
    ''' </summary>
    ''' <param name="vsoKeyboardEvent">The KeyboardEvent Object which triggered the event..</param>
    ''' <returns>True to cancel the key release. False to process the key release.</returns>
    ''' <remarks>Possible values for KeyCode are declared in KeyCodeConstants in the Microsoft Visual Basic for Applications (VBA) library.
    ''' Possible values for KeyButtonState can be a combination of the values shown in the following table, which are declared in 
    ''' VisKeyButtonFlags in the Visio type library. For example, if KeyButtonState returns 12, it indicates that the user held down both 
    ''' SHIFT and CTRL.
    ''' visKeyControl 8
    ''' visKeyShift 4 
    ''' visMouseLeft 1
    ''' visMouseMiddle 16
    ''' visMouseRight 2
    ''' If you set CancelDefault to True, Visio does not process the message received when the mouse button is clicked.
    ''' Unlike some other Visio events, KeyDown does not have the prefix "Query," but it is still a query event. 
    ''' That is, you can cancel processing the message sent by KeyDown, either by setting CancelDefault to True, or, if you are using 
    ''' theVisEventProc method to handle the event, by returning True. 
    ''' Pressing an accelererator key combination, for example, CTRL + C, does not fire the KeyDown event.</remarks>
    Protected Overridable Function HandleKeyUp(ByVal vsoKeyboardEvent As Visio.KeyboardEvent) As Boolean
    End Function

#End Region

#Region "DataRecordSet Events"

    'Pro Edition Only
    ''' <summary>
    ''' Handles the DataRecordsetAdded event.
    ''' Occurs when a DataRecordset object is added to a DataRecordsets collection.
    ''' </summary>
    ''' <param name="vsoDataRecordset">The data recordset that was added.</param>
    ''' <remarks>This Visio object or member is available only to licensed users of Microsoft Office Visio Professional 2007.</remarks>
    Protected Overridable Sub HandleDataRecordsetAdded(ByVal vsoDataRecordset As Visio.DataRecordset)
    End Sub
    ''' <summary>
    ''' Handles the BeforeDataRecordsetDelete event.
    ''' Occurs before a DataRecordset object is deleted from the DataRecordsets collection.
    ''' </summary>
    ''' <param name="vsoDataRecordset">The data recordset that is going to be deleted.</param>
    ''' <remarks>This Visio object or member is available only to licensed users of Microsoft Office Visio Professional 2007.</remarks>
    Protected Overridable Sub HandleBeforeDataRecordsetDelete(ByVal vsoDataRecordset As Visio.DataRecordset)
    End Sub

#End Region

#Region "DataRecordsetChanged Event"

    ''' <summary>
    ''' Handles the DataRecordsetChanged event.
    ''' Occurs when a data recordset changes as a result of being refreshed.
    ''' </summary>
    ''' <param name="vsoDataRecordSetChangedEvent">The data recordset that changed.</param>
    ''' <remarks>This Visio object or member is available only to licensed users of Microsoft Office Visio Professional 2007.
    ''' When the DataRecordsetChanged event fires, the DataRecordsetChangedEvent object is passed to the IVisEventProc.VisEventProc 
    ''' method as the pSubjectObj parameter, which represents he subject of the event—the object to which the event occurs.</remarks>
    Protected Overridable Sub HandleDataRecordsetChanged(ByVal vsoDataRecordSetChangedEvent As Visio.DataRecordsetChangedEvent)
    End Sub

#End Region

    ''' <summary>
    ''' Handles the unknown event.
    ''' </summary>
    ''' <param name="eventCode">The event code.</param>
    ''' <param name="subject">The subject.</param>
    ''' <param name="moreInformation">The more information.</param>
    Protected Overridable Sub HandleUnknownEvent(ByVal eventCode As Integer, _
                                                 ByVal subject As Object, _
                                                 ByVal moreInformation As Object)
    End Sub

#End Region

#Region "Overridable Logging"
    ''' <summary>
    ''' Logs the event.
    ''' </summary>
    ''' <param name="eventMessage">The event message.</param>
    Protected Overridable Sub logTheEvent(ByVal eventMessage As String)
        ' Write the event info to the output window
        System.Diagnostics.Debug.WriteLine(eventMessage)
    End Sub
#End Region

#Region " IDisposable Support "

    'IDisposable
    ''' <summary>
    ''' Releases unmanaged and - optionally - managed resources
    ''' </summary>
    ''' <param name="disposing"><c>true</c> to release both managed and unmanaged resources; <c>false</c> to release only unmanaged resources.</param>
    Private Overloads Sub Dispose(ByVal disposing As Boolean)
        If Not Me._disposed Then
            If disposing Then
                ' TODO: put code to dispose managed resources
            End If
            ' TODO: put code to free unmanaged resources here
            _eventDescriptions = Nothing
        End If
        Me._disposed = True
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    ''' <summary>
    ''' Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
    ''' </summary>
    Public Overloads Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    ''' <summary>
    ''' Allows an <see cref="T:System.Object" /> to attempt to free resources and perform other cleanup operations before the <see cref="T:System.Object" /> is reclaimed by garbage collection.
    ''' </summary>
    Protected Overrides Sub Finalize()
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(False)
        MyBase.Finalize()
    End Sub

#End Region

End Class
''' <summary>
''' Enumerations to support the Visio Power Tools Assembly.
''' </summary>
Public Module PublicEnums

    ''' <summary>
    ''' Enumeration to track how to access the data of a <see href="http://msdn.microsoft.com/en-us/library/ms408988%28v=office.12%29.aspx">Visio.Section</see> object.
    ''' <seealso href="http://msdn.microsoft.com/en-us/library/bb902804%28v=office.12%29.aspx#Y2964">How to Access Different Section Types</seealso>
    ''' </summary>
    Public Enum visSectionTypes As Byte
        ''' <summary>
        ''' The section has named rows which can be added/deleted by the user.  
        ''' </summary>
        NamedRowsConstantCells = 0
        ''' <summary>
        ''' The section has unnamed rows, but have a constant number of rows and cells.
        ''' </summary>
        UnnamedConstantRows = 1
        ''' <summary>
        ''' The section has unnamed, non-constant rows, but each row has a constant number of cells.
        ''' </summary>
        UnnamedNonConstantRows = 2
        ''' <summary>
        ''' The section has unnamed, non-constant rows, cell constants can only be determined using RowType
        ''' </summary>
        UnnamedNonConstantRowsAndCells = 3
        ''' <summary>
        ''' The section is named, non-constant rows, cell constants can only be determined using RowType
        ''' </summary>
        NamedNonConstantRowsAndCells = 4
        ''' <summary>
        ''' The section can be named or unnamed, non-constant rows, cell constants can only be determined using RowType
        ''' </summary>
        NamedOrUnnamedNonConstantRowsAndCells = 5
        ''' <summary>
        ''' The section is not a valid Visio section. 
        ''' </summary>
        IsInvalid = 6
    End Enum

    ''' <summary>
    ''' Flags to extend the existing set flags and give functionality to choose the depth of the set
    ''' </summary>
    Public Enum visSetFlagsExtended As Byte
        ''' <summary>
        ''' The entire section will be cleared and replaced with the source section's data.
        ''' </summary>
        visSetReplaceAllExisting = 0
        ''' <summary>
        ''' The section will have any data from matching rows replaced by the source section's data. 
        ''' If the destination section does not have matching rows, the row will be ignored. 
        ''' </summary>
        visSetReplaceSelectiveAndIgnore = 1
        ''' <summary>
        ''' The section will have any data from matching rows replaced by the source section's data. 
        ''' If the destination section does not have matching rows, the rows will be added. 
        ''' </summary>
        visSetReplaceSelectiveAndAdd = 2
    End Enum

    ''' <summary>
    ''' Used to set Visio's AlertResponse property to automatically respond to prompts when needed.
    ''' </summary>
    Public Enum IDResponses As Short
        ''' <summary>
        ''' The Default Response for a given scenario.
        ''' </summary>
        IDDEFAULT = 0
        ''' <summary>
        ''' The "OK" Response
        ''' </summary>
        IDOK = 1
        ''' <summary>
        ''' The "CANCEL" Response
        ''' </summary>
        IDCANCEL = 2
        ''' <summary>
        ''' The "ABORT" Response
        ''' </summary>
        IDABORT = 3
        ''' <summary>
        ''' The "RETRY" Response
        ''' </summary>
        IDRETRY = 4
        ''' <summary>
        ''' The "IGNORE" Response
        ''' </summary>
        IDIGNORE = 5
        ''' <summary>
        ''' The "YES" Response
        ''' </summary>
        IDYES = 6
        ''' <summary>
        ''' The "NO" Response
        ''' </summary>
        IDNO = 7
    End Enum

    ''' <summary>
    ''' All of the predefined Visio Events Enumerated by Name
    ''' Created using the defualt enumeration located here: <see cref="Visio.VisEventCodes" />
    ''' </summary>
    Public Enum AllEvents As Short
        ''----------------------------------------------------------------
        '' While the events could be sperated based on subject (Document, Page, Application, etc.) This 
        '' leads to virtually unreadable documentation. As such they are listed alphabetically.
        '' You can see there subject seperation in the evetn sink class.s
        ''----------------------------------------------------------------

        ''' <summary>
        ''' AfterModal, visEvtApp+visEvtAfterModal,  &amp;H1040 (4160)
        ''' </summary>
        AfterModal = _
            CShort(Visio.VisEventCodes.visEvtApp) + _
            CShort(Visio.VisEventCodes.visEvtAfterModal)

        ''' <summary>
        ''' AfterRemoveHiddenInformation, visEvtRemoveHiddenInformation(), &amp;H000A (11)
        ''' </summary>
        AfterRemoveHiddenInformation = _
            CShort(Visio.VisEventCodes.visEvtRemoveHiddenInformation)

        ''' <summary>
        ''' AfterResume, visEvtCodeAfterResume, &amp;H00D1 (209)
        ''' </summary>
        AfterResume = _
            CShort(Visio.VisEventCodes.visEvtCodeAfterResume)

        ''' <summary>
        ''' AppActivated, visEvtApp+visEvtAppActivate, &amp;H1001 (4097)
        ''' </summary>
        AppActivated = _
            CShort(Visio.VisEventCodes.visEvtApp) + _
            CShort(Visio.VisEventCodes.visEvtAppActivate)

        ''' <summary>
        ''' AppDeactivated, visEvtApp+visEvtAppDeactivate, &amp;H1002 (4098)
        ''' </summary>
        AppDeactivated = _
            CShort(Visio.VisEventCodes.visEvtApp) + _
            CShort(Visio.VisEventCodes.visEvtAppDeactivate)

        ''' <summary>
        ''' AppObjActivated, visEvtApp+visEvtObjActivate, &amp;H1004 (4100)
        ''' </summary>
        AppObjActivated = _
            CShort(Visio.VisEventCodes.visEvtApp) + _
            CShort(Visio.VisEventCodes.visEvtObjActivate)

        ''' <summary>
        ''' AppObjDeactivated, visEvtApp+visEvtObjDeactivate, &amp;H1008 (4104)
        ''' </summary>
        AppObjDeactivated = _
            CShort(Visio.VisEventCodes.visEvtApp) + _
            CShort(Visio.VisEventCodes.visEvtObjDeactivate)

        ''' <summary>
        ''' BeforeDataRecordsetDelete(), visEvtDel(+visEvtDataRecordset), &amp;H4020 (16416)
        ''' </summary>
        BeforeDataRecordsetDelete = _
            CShort(Visio.VisEventCodes.visEvtDataRecordset) + _
            CShort(Visio.VisEventCodes.visEvtDel)

        ''' <summary>
        ''' BeforeDocumentClose, visEvtDel+visEvtDoc, &amp;H4002 (16386)
        ''' </summary>
        BeforeDocumentClose = _
            CShort(Visio.VisEventCodes.visEvtDoc) + _
            CShort(Visio.VisEventCodes.visEvtDel)

        ''' <summary>
        ''' BeforeDocumentSave, visEvtCodeBefDocSave(), &amp;0007 (7)
        ''' </summary>
        BeforeDocumentSave = _
            CShort(Visio.VisEventCodes.visEvtCodeBefDocSave)

        ''' <summary>
        ''' BeforeDocumentSaveAs, visEvtCodeBefDocSaveAs(), &amp;H0008 (8) 
        ''' </summary>
        BeforeDocumentSaveAs = _
            CShort(Visio.VisEventCodes.visEvtCodeBefDocSaveAs)

        ''' <summary>
        ''' BeforeMasterDelete() ,visEvtDel(+visEvtMaster), &amp;H4008 (16392)
        ''' </summary>
        BeforeMasterDelete = _
            CShort(Visio.VisEventCodes.visEvtMaster) + _
            CShort(Visio.VisEventCodes.visEvtDel)

        ''' <summary>
        ''' BeforeModal(), visEvtApp(+visEvtBeforeModal), &amp;H1020 (4128)
        ''' </summary>
        BeforeModal = _
            CShort(Visio.VisEventCodes.visEvtApp) + _
            CShort(Visio.VisEventCodes.visEvtBeforeModal)

        ''' <summary>
        ''' BeforePageDelete, visEvtDel(+visEvtPage), &amp;H4010 (16400)
        ''' </summary>
        BeforePageDelete = _
            CShort(Visio.VisEventCodes.visEvtPage) + _
            CShort(Visio.VisEventCodes.visEvtDel)

        ''' <summary>
        ''' BeforeQuit(), visEvtApp(+visEvtBeforeQuit), &amp;H1010 (4112)
        ''' </summary>
        BeforeQuit = _
            CShort(Visio.VisEventCodes.visEvtApp) + _
            CShort(Visio.VisEventCodes.visEvtBeforeQuit)

        ''' <summary>
        ''' BeforeSelectionDelete(), visEvtCodeBefSelDel(), &amp;H0385 (901)
        ''' </summary>
        BeforeSelectionDelete = _
            CShort(Visio.VisEventCodes.visEvtCodeBefSelDel)

        ''' <summary>
        ''' BeforeShapeDelete(), visEvtDel(+visEvtShape), &amp;H4040 (16448)
        ''' </summary>
        BeforeShapeDelete = _
            CShort(Visio.VisEventCodes.visEvtShape) + _
            CShort(Visio.VisEventCodes.visEvtDel)

        ''' <summary>
        ''' BeforeShapeTextEdit, visEvtCodeShapeBeforeTextEdit(), &amp;H0323 (803)
        ''' </summary>
        BeforeShapeTextEdit = _
            CShort(Visio.VisEventCodes.visEvtCodeShapeBeforeTextEdit)

        ''' <summary>
        ''' BeforeStyleDelete(), visEvtDel(+visEvtStyle), &amp;H4004 (16388)
        ''' </summary>
        BeforeStyleDelete = _
        CShort(Visio.VisEventCodes.visEvtStyle) + _
        CShort(Visio.VisEventCodes.visEvtDel)

        ''' <summary>
        ''' BeforeSuspend, visEvtCodeBeforeSuspend(), &amp;H00D0(208)
        ''' </summary>
        BeforeSuspend = _
            CShort(Visio.VisEventCodes.visEvtCodeBeforeSuspend)

        ''' <summary>
        ''' BeforeWindowClosed(), visEvtDel(+visEvtWindow), &amp;H4001 (16385)
        ''' </summary>
        BeforeWindowClosed = _
            CShort(Visio.VisEventCodes.visEvtWindow) + _
            CShort(Visio.VisEventCodes.visEvtDel)

        ''' <summary>
        ''' BeforeWindowPageTurn(), visEvtCodeBefWinPageTurn(), &amp;H02BF (703)
        ''' </summary>
        BeforeWindowPageTurn = _
            CShort(Visio.VisEventCodes.visEvtCodeBefWinPageTurn)

        ''' <summary>
        ''' BeforeWindowSelDelete(), visEvtCodeBefWinSelDel(), &amp;H02BE (702)
        ''' </summary>
        BeforeWindowSelDelete = _
            CShort(Visio.VisEventCodes.visEvtCodeBefWinSelDel)

        ''' <summary>
        ''' CellChanged(), visEvtMod(+visEvtCell), &amp;H2800 (10240)
        ''' </summary>
        CellChanged = _
            CShort(Visio.VisEventCodes.visEvtCell) + _
            CShort(Visio.VisEventCodes.visEvtMod)

        ''' <summary>
        ''' ConnectionsAdded, g_visEvtAdd(+visEvtConnect), &amp;H8100 (33024)
        ''' </summary>
        ConnectionsAdded = _
            CShort(Visio.VisEventCodes.visEvtConnect) + _
            visEvtAdd

        ''' <summary>
        ''' ConnectionsDeleted(), visEvtDel(+visEvtConnect), &amp;H4100 (16640)
        ''' </summary>
        ConnectionsDeleted = _
            CShort(Visio.VisEventCodes.visEvtConnect) + _
            CShort(Visio.VisEventCodes.visEvtDel)

        ''' <summary>
        ''' ConvertToGroupCanceled,visEvtCodeCancelConvertToGroup(), &amp;H038C (908)
        ''' </summary>
        ConvertToGroupCanceled = _
            CShort(Visio.VisEventCodes.visEvtCodeCancelConvertToGroup)

        ''' <summary>
        ''' DataRecordsetAdded(), g_visEvtAdd(+visEvtDataRecordset), &amp;H8020 (32800)
        ''' </summary>
        DataRecordsetAdded = _
            CShort(Visio.VisEventCodes.visEvtDataRecordset) + _
            visEvtAdd

        ''' <summary>
        ''' DataRecordsetChanged(), visEvtMod(+VisEvtDataRecordset), &amp;H2020 (8224)
        ''' </summary>
        DataRecordsetChanged = _
            CShort(Visio.VisEventCodes.visEvtMod) + _
            CShort(Visio.VisEventCodes.visEvtDataRecordset)

        ''' <summary>
        ''' DesignModeEntered, visEvtCodeDocDesign(), &amp;H0006 (6)
        ''' </summary>
        DesignModeEntered = _
            CShort(Visio.VisEventCodes.visEvtCodeDocDesign)

        ''' <summary>
        ''' DocumentAdded, g_visEvtAdd(+visEvtDoc),  &amp;H8002 (32770)
        ''' </summary>
        DocumentAdded = _
            CShort(Visio.VisEventCodes.visEvtDoc) + _
            visEvtAdd

        ''' <summary>
        ''' DocumentChanged(), visEvtMod(+visEvtDoc), &amp;H2002 (8194)
        ''' </summary>
        DocumentChanged = _
            CShort(Visio.VisEventCodes.visEvtDoc) + _
            CShort(Visio.VisEventCodes.visEvtMod)

        ''' <summary>
        ''' DocumentCloseCanceled(), visEvtCodeCancelDocClose(), &amp;H0010 (10)
        ''' </summary>
        DocumentCloseCanceled = _
            CShort(Visio.VisEventCodes.visEvtCodeCancelDocClose)

        ''' <summary>
        ''' DocumentCreated(), visEvtCodeDocCreate(), &amp;H0001 (1)
        ''' </summary>
        DocumentCreated = _
            CShort(Visio.VisEventCodes.visEvtCodeDocCreate)

        ''' <summary>
        ''' DocumentOpened(), visEvtCodeDocOpen(), &amp;H0002 (2)
        ''' </summary>
        DocumentOpened = _
           CShort(Visio.VisEventCodes.visEvtCodeDocOpen)

        ''' <summary>
        ''' DocumentSaved(), visEvtCodeDocSave(), &amp;H0003 (3)
        ''' </summary>
        DocumentSaved = _
            CShort(Visio.VisEventCodes.visEvtCodeDocSave)

        ''' <summary>
        ''' DocumentSavedAs(), visEvtCodeDocSaveAs(), &amp;H0004 (4)
        ''' </summary>
        DocumentSavedAs = _
            CShort(Visio.VisEventCodes.visEvtCodeDocSaveAs)

        ''' <summary>
        ''' EnterScope(), visEvtCodeEnterScope(), &amp;H00CA (202)
        ''' </summary>
        EnterScope = _
            CShort(Visio.VisEventCodes.visEvtCodeEnterScope)

        ''' <summary>
        ''' ExitScope(), visEvtCodeExitScope(), &amp;H00CB (203)
        ''' </summary>
        ExitScope = _
            CShort(Visio.VisEventCodes.visEvtCodeExitScope)

        ''' <summary>
        ''' FormulaChanged, visEvtMod(+visEvtFormula), &amp;H3000 (12288)
        ''' </summary>
        FormulaChanged = _
            CShort(Visio.VisEventCodes.visEvtFormula) + _
            CShort(Visio.VisEventCodes.visEvtMod)

        ''' <summary>
        ''' GroupCanceled, visEvtCodeCancelSelGroup(), &amp;H038E (910)
        ''' </summary>
        GroupCanceled = _
            CShort(Visio.VisEventCodes.visEvtCodeCancelSelGroup)

        ''' <summary>
        ''' KeyDown(), visEvtCodeKeyDown(), &amp;H02C8 (712)
        ''' </summary>
        KeyDown = _
            CShort(Visio.VisEventCodes.visEvtCodeKeyDown)

        ''' <summary>
        ''' KeyPress(), visEvtCodeKeyPress(), &amp;H02C9 (713)
        ''' </summary>
        KeyPress = _
            CShort(Visio.VisEventCodes.visEvtCodeKeyPress)

        ''' <summary>
        ''' KeyUp(), visEvtCodeKeyUp(), &amp;H02CA (714)
        ''' </summary>
        KeyUp = _
            CShort(Visio.VisEventCodes.visEvtCodeKeyUp)

        ''' <summary>
        ''' MarkerEvent(), visEvtApp(+visEvtMarker), &amp;H1100 (4352) 
        ''' </summary>
        MarkerEvent = _
            CShort(Visio.VisEventCodes.visEvtApp) + _
            CShort(Visio.VisEventCodes.visEvtMarker)

        ''' <summary>
        ''' MasterAdded, g_visEvtAdd(+visEvtMaster), &amp;H8008 (32776)
        ''' </summary>
        MasterAdded = _
            CShort(Visio.VisEventCodes.visEvtMaster) + _
            visEvtAdd

        ''' <summary>
        ''' MasterChanged, visEvtMod(+visEvtMaster), &amp;H2008 (8200)
        ''' </summary>
        MasterChanged = _
            CShort(Visio.VisEventCodes.visEvtMaster) + _
            CShort(Visio.VisEventCodes.visEvtMod)

        ''' <summary>
        ''' MasterDeleteCanceled, visEvtCodeCancelMasterDel(), &amp;H0191 (401)
        ''' </summary>
        MasterDeleteCanceled = _
            CShort(Visio.VisEventCodes.visEvtCodeCancelMasterDel)

        ''' <summary>
        ''' MouseDown(), visEvtCodeMouseDown(), &amp;H02C5 (709)
        ''' </summary>
        MouseDown = _
            CShort(Visio.VisEventCodes.visEvtCodeMouseDown)

        ''' <summary>
        ''' MouseMove(), visEvtCodeMouseMove(), &amp;H02C6 (710)
        ''' </summary>
        MouseMove = _
            CShort(Visio.VisEventCodes.visEvtCodeMouseMove)

        ''' <summary>
        ''' MouseUp(), visEvtCodeMouseUp(), &amp;H02C7 (711)
        ''' </summary>
        MouseUp = _
            CShort(Visio.VisEventCodes.visEvtCodeMouseUp)

        ''' <summary>
        ''' MustFlushScopeBeginning(), visEvtCodeBefForcedFlush(), &amp;H00C8 (200)
        ''' </summary>
        MustFlushScopeBeginning = _
            CShort(Visio.VisEventCodes.visEvtCodeBefForcedFlush)

        ''' <summary>
        ''' MustFlushScopeEnded(), visEvtCodeAfterForcedFlush(), &amp;H00C9 (201)
        ''' </summary>
        MustFlushScopeEnded = _
            CShort(Visio.VisEventCodes.visEvtCodeAfterForcedFlush)

        ''' <summary>
        ''' NoEventsPending(), visEvtApp(+visEvtNonePending), &amp;H1200 (4608)
        ''' </summary>
        NoEventsPending = _
            CShort(Visio.VisEventCodes.visEvtApp) + _
            CShort(Visio.VisEventCodes.visEvtNonePending)

        ''' <summary>
        ''' OnKeystrokeMessageForAddon(), visEvtCodeWinOnAddonKeyMSG(), &amp;H02C4 (708)
        ''' </summary>
        OnKeystrokeMessageForAddon = _
            CShort(Visio.VisEventCodes.visEvtCodeWinOnAddonKeyMSG)

        ''' <summary>
        ''' PageAdded, g_visEvtAdd(+visEvtPage), &amp;H8010 (32784)
        ''' </summary>
        PageAdded = _
            CShort(Visio.VisEventCodes.visEvtPage) + _
            visEvtAdd

        ''' <summary>
        ''' PageChanged(), visEvtMod(+visEvtPage), &amp;H2010 (8208)
        ''' </summary>
        PageChanged = _
            CShort(Visio.VisEventCodes.visEvtPage) + _
            CShort(Visio.VisEventCodes.visEvtMod)

        ''' <summary>
        ''' PageDeleteCanceled(), visEvtCodeCancelPageDel(), &amp;H01F5 (501)
        ''' </summary>
        PageDeleteCanceled = _
            CShort(Visio.VisEventCodes.visEvtCodeCancelPageDel)

        ''' <summary>
        ''' QueryCancelConvertToGroup, visEvtCodeQueryCancelConvertToGroup(), &amp;H038B (907)
        ''' </summary>
        QueryCancelConvertToGroup = _
            CShort(Visio.VisEventCodes.visEvtCodeQueryCancelConvertToGroup)

        ''' <summary>
        ''' QueryCancelDocumentClose(), visEvtCodeQueryCancelDocClose(), &amp;H0009 (9)
        ''' </summary>
        QueryCancelDocumentClose = _
            CShort(Visio.VisEventCodes.visEvtCodeQueryCancelDocClose)

        ''' <summary>
        ''' QueryCancelGroup(), visEvtCodeQueryCancelSelGroup(), &amp;H038D (909)
        ''' </summary>
        QueryCancelGroup = _
            CShort(Visio.VisEventCodes.visEvtCodeQueryCancelSelGroup)

        ''' <summary>
        ''' QueryCancelMasterDelete, visEvtCodeQueryCancelMasterDel(), &amp;H0190 (400)
        ''' </summary>
        QueryCancelMasterDelete = _
            CShort(Visio.VisEventCodes.visEvtCodeQueryCancelMasterDel)

        ''' <summary>
        ''' QueryCancelPageDelete(), visEvtCodeQueryCancelPageDel(), &amp;H01F4 (500)
        ''' </summary>
        QueryCancelPageDelete = _
            CShort(Visio.VisEventCodes.visEvtCodeQueryCancelPageDel)

        ''' <summary>
        ''' QueryCancelQuit(), visEvtCodeQueryCancelQuit(), &amp;H00CC (204)
        ''' </summary>
        QueryCancelQuit = _
            CShort(Visio.VisEventCodes.visEvtCodeQueryCancelQuit)

        ''' <summary>
        ''' QueryCancelSelectionDelete, visEvtCodeQueryCancelSelDel(), &amp;H0387 (903)
        ''' </summary>
        QueryCancelSelectionDelete = _
            CShort(Visio.VisEventCodes.visEvtCodeQueryCancelSelDel)

        ''' <summary>
        ''' QueryCancelStyleDelete(), visEvtCodeQueryCancelStyleDel(), &amp;H012C (300)
        ''' </summary>
        QueryCancelStyleDelete = _
            CShort(Visio.VisEventCodes.visEvtCodeQueryCancelStyleDel)

        ''' <summary>
        ''' QueryCancelSuspend(), visEvtCodeQueryCancelSuspend(), &amp;H00CE(206)
        ''' </summary>
        QueryCancelSuspend = _
            CShort(Visio.VisEventCodes.visEvtCodeQueryCancelSuspend)

        ''' <summary>
        ''' QueryCancelUngroup, visEvtCodeQueryCancelUngroup(), &amp;H0389 (905)
        ''' </summary>
        QueryCancelUngroup = _
            CShort(Visio.VisEventCodes.visEvtCodeQueryCancelUngroup)

        ''' <summary>
        ''' QueryCancelWindowClose, visEvtCodeQueryCancelWinClose(), &amp;H02C2 (706)
        ''' </summary>
        QueryCancelWindowClose = _
            CShort(Visio.VisEventCodes.visEvtCodeQueryCancelWinClose)

        ''' <summary>
        ''' QuitCanceled(), visEvtCodeCancelQuit(), &amp;H00CD (205)
        ''' </summary>
        QuitCanceled = _
            CShort(Visio.VisEventCodes.visEvtCodeCancelQuit)

        ''' <summary>
        ''' RunModeEntered(), visEvtCodeDocRunning(), &amp;H0005 (5)
        ''' </summary>
        RunModeEntered = _
            CShort(Visio.VisEventCodes.visEvtCodeDocRunning)

        ''' <summary>
        ''' SelectionAdded(), visEvtCodeSelAdded(), &amp;H0386 (902)
        ''' </summary>
        SelectionAdded = _
            CShort(Visio.VisEventCodes.visEvtCodeSelAdded)

        ''' <summary>
        ''' SelectionChanged(), visEvtCodeWinSelChange(), &amp;H02BD (701)
        ''' </summary>
        SelectionChanged = _
            CShort(CShort(Visio.VisEventCodes.visEvtCodeWinSelChange))

        ''' <summary>
        ''' SelectionDeleteCanceled(), visEvtCodeCancelSelDel(), &amp;H0388(904)
        ''' </summary>
        SelectionDeleteCanceled = _
            CShort(Visio.VisEventCodes.visEvtCodeCancelSelDel)

        ''' <summary>
        ''' ShapeAdded, g_visEvtAdd(+visEvtShape), &amp;H8040 (32832)
        ''' </summary>
        ShapeAdded = _
            CShort(Visio.VisEventCodes.visEvtShape) + _
            visEvtAdd

        ''' <summary>
        ''' ShapeChanged(), visEvtMod(+visEvtShape), &amp;H2040 (8256)
        ''' </summary>
        ShapeChanged = _
            CShort(Visio.VisEventCodes.visEvtShape) + _
            CShort(Visio.VisEventCodes.visEvtMod)

        ''' <summary>
        ''' ShapeDataGraphicChanged(), visEvtShapeDataGraphicChanged(), &amp;H0327 (807)
        ''' </summary>
        ''' <remarks>Available in Professional Edition Only</remarks>
        ShapeDataGraphicChanged = _
            CShort(Visio.VisEventCodes.visEvtShapeDataGraphicChanged)

        ''' <summary>
        ''' ShapeExitedTextEdit(), visEvtCodeShapeExitTextEdit(), &amp;H0324 (804)
        ''' </summary>
        ShapeExitedTextEdit = _
            CShort(Visio.VisEventCodes.visEvtCodeShapeExitTextEdit)

        ''' <summary>
        ''' ShapeLinkAdded(), visEvtShapeLinkAdded(), &amp;H0325 (805)
        ''' </summary>
        ShapeLinkAdded = _
            CShort(Visio.VisEventCodes.visEvtShapeLinkAdded)

        ''' <summary>
        ''' ShapeLinkDeleted(), visEvtShapeLinkDeleted(), &amp;H0326 (806)
        ''' </summary>
        ShapeLinkDeleted = _
            CShort(Visio.VisEventCodes.visEvtShapeLinkDeleted)

        ''' <summary>
        ''' ShapeParentChanged(), visEvtCodeShapeParentChange(), &amp;H0322 (802)
        ''' </summary>
        ShapeParentChanged = _
            CShort(Visio.VisEventCodes.visEvtCodeShapeParentChange)

        ''' <summary>
        ''' ShapesDeleted(), visEvtCodeShapeDelete(), &amp;H0321 (801)
        ''' </summary>
        ShapesDeleted = _
            CShort(Visio.VisEventCodes.visEvtCodeShapeDelete)

        ''' <summary>
        ''' StyleAdded(), g_visEvtAdd(+visEvtStyle), &amp;H8004 (32772)
        ''' </summary>
        StyleAdded = _
            CShort(Visio.VisEventCodes.visEvtStyle) + _
            visEvtAdd

        ''' <summary>
        ''' StyleChanged(), visEvtMod(+visEvtStyle), &amp;H2004 (8196)
        ''' </summary>
        StyleChanged = _
            CShort(Visio.VisEventCodes.visEvtStyle) + _
            CShort(Visio.VisEventCodes.visEvtMod)

        ''' <summary>
        ''' StyleDeleteCanceled(), visEvtCodeCancelStyleDel(), &amp;H012D (301)
        ''' </summary>
        StyleDeleteCanceled = _
            CShort(Visio.VisEventCodes.visEvtCodeCancelStyleDel)

        ''' <summary>
        ''' SuspendCanceled(), visEvtCodeCancelSuspend(), &amp;H00CF(207)
        ''' </summary>
        SuspendCanceled = _
            CShort(Visio.VisEventCodes.visEvtCodeCancelSuspend)

        ''' <summary>
        ''' TextChanged, visEvtMod(+visEvtText), &amp;H2080 (8320)
        ''' </summary>
        TextChanged = _
            CShort(Visio.VisEventCodes.visEvtText) + _
            CShort(Visio.VisEventCodes.visEvtMod)

        ''' <summary>
        ''' UngroupCanceled, visEvtCodeCancelUngroup(),  &amp;H038A (906)
        ''' </summary>
        UngroupCanceled = _
            CShort(Visio.VisEventCodes.visEvtCodeCancelUngroup)

        ''' <summary>
        ''' ViewChanged, visEvtCodeViewChanged(), &amp;H02C1 (705)
        ''' </summary>
        ViewChanged = _
            CShort(Visio.VisEventCodes.visEvtCodeViewChanged)

        ''' <summary>
        ''' VisioIsIdle(), visEvtApp(+visEvtIdle), &amp;H1400 (5120)
        ''' </summary>
        VisioIsIdle = _
            CShort(Visio.VisEventCodes.visEvtApp) + _
            CShort(Visio.VisEventCodes.visEvtIdle)

        ''' <summary>
        ''' WindowActivated(), visEvtApp(+visEvtWinActivate), &amp;H1080 (4224)
        ''' </summary>
        WindowActivated = _
            CShort(Visio.VisEventCodes.visEvtApp) + _
            CShort(Visio.VisEventCodes.visEvtWinActivate)

        ''' <summary>
        ''' WindowChanged(), visEvtMod(+visEvtWindow), &amp;H2001 (8193)
        ''' </summary>
        WindowChanged = _
            CShort(Visio.VisEventCodes.visEvtWindow) + _
            CShort(Visio.VisEventCodes.visEvtMod)

        ''' <summary>
        ''' WindowCloseCanceled(), visEvtCodeCancelWinClose(), &amp;H02C3 (707)
        ''' </summary>
        WindowCloseCanceled = _
            CShort(Visio.VisEventCodes.visEvtCodeCancelWinClose)

        ''' <summary>
        ''' WindowOpened(), g_visEvtAdd(+visEvtWindow), &amp;H8001 (32769)
        ''' </summary>
        WindowOpened = _
            CShort(Visio.VisEventCodes.visEvtWindow) + _
            visEvtAdd

        ''' <summary>
        ''' WindowTurnedToPage(), visEvtCodeWinPageTurn(), &amp;H02C0 (704)
        ''' </summary>
        WindowTurnedToPage = _
            CShort(Visio.VisEventCodes.visEvtCodeWinPageTurn)

    End Enum

    ''' <summary>
    ''' Arguments used when Copying shape data from one shape to another.
    ''' </summary>
    Public Enum visCopyShapeDataArgs As Integer
        ''' <summary>
        ''' Adds any rows from the source shape which aren't already in the destination shape. 
        ''' </summary>
        addIfNonExisting = 2
        ''' <summary>
        ''' Deletes any rows from the destination shape which aren't already in the source shape. 
        ''' </summary>
        deleteIfNoMatch = 4
        ''' <summary>
        ''' Protects any formula in the destination shape which references "ThePage!" or "TheDoc!" from being overwritten.
        ''' </summary>
        protectReferences = 8
    End Enum

    ''' <summary>
    ''' Arguments used when searching and replacing text in shapeData.
    ''' </summary>
    Public Enum visSearchAndReplaceArgs As Integer
        ''' <summary>
        ''' Normal Search.
        ''' </summary>
        None = 0
        ''' <summary>
        ''' Match case while searching.
        ''' </summary>
        MatchCase = 1
        ''' <summary>
        ''' Only match whole words, not text inside other words.
        ''' </summary>
        WholeWordsOnly = 2
        ''' <summary>
        ''' The "findText" will be treated as a regular expression and matched as a pattern. 
        ''' All other searchArgs will be ignored.
        ''' </summary>
        RegEx = 3
    End Enum

End Module

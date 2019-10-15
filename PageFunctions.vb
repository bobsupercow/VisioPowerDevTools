Imports System.Text
Imports System.Runtime.CompilerServices

''' <summary>
''' Functions and Methods used to process data related to 
''' <see href="http://msdn.microsoft.com/en-us/library/ms408981%28v=office.12%29.aspx">Visio.Page</see> objects.
''' </summary>
Public Module PageFunctions

    ''' <summary>
    ''' <see cref=" Visio.VisPaperSizes"/> 
    ''' Copies a <see href="http://msdn.microsoft.com/en-us/library/ms408981%28v=office.12%29.aspx">Visio.Page</see>. 
    ''' Additionally, copies the information from the User, Action, and Properties sections.
    ''' Optionally, copies all shapes from all layers.
    ''' </summary>
    ''' <param name="vsoPage">The <see href="http://msdn.microsoft.com/en-us/library/ms408981%28v=office.12%29.aspx">Visio.Page</see> to copy.</param>
    ''' <param name="deepCopy">If set to <c>True</c>, the function performs a deep copy of all shapes on all layers regardless of the layer's lock status.</param>
    ''' <returns>
    ''' The copied page. 
    ''' </returns>
    Public Function CopyPage(ByVal vsoPage As Visio.Page, _
                             ByVal deepCopy As Boolean) _
                             As Visio.Page

        Const maxCharsToCopy As Integer = 23

        Try
            'Test if the page is Null
            If vsoPage IsNot Nothing Then
                Dim activeWindow As Visio.Window
                'Windows Collections is 1 indexed.
                'Loop through all of the windows in the application, activate the one containing the source page. 
                For currWindow = 1 To vsoPage.Application.Windows.Count
                    If vsoPage.Application.Windows(currWindow).Document Is vsoPage.Document Then
                        vsoPage.Application.Windows(currWindow).Activate()
                    End If
                Next


                'The page is NOT Null, continue.
                activeWindow = vsoPage.Application.ActiveWindow

                Dim layerLockStatusArray() As Short = Nothing
                'Read the page's layers' lock status into an array.
                'Test if deep copy. 
                If deepCopy = True Then
                    'We want a deep copy, so we will unlock each layer to ensure that 
                    'shapes on all layers are copied. In order to make sure the page isn't changed, 
                    'we will read the layer lock status into an array so it can be restored later. 
                    'Create an array to hold the lock status of the page. 
                    ReDim layerLockStatusArray(vsoPage.Layers.Count)
                    'Layers are 1 indexed. 
                    'Loop through each layer on the page.
                    For layerIndex = 1 To vsoPage.Layers.Count
                        'Read the lock status into the array.
                        'Uses layerIndex - 1 because the array is 0-based, but Layers is 1-based.
                        layerLockStatusArray(layerIndex - 1) = _
                            vsoPage.Layers(layerIndex).CellsC(Visio.VisCellIndices.visLayerLock). _
                            ResultInt(Visio.VisUnitCodes.visNumber, Visio.VisRoundFlags.visTruncate)
                        'Unlock the layer
                        vsoPage.Layers(layerIndex).CellsC(Visio.VisCellIndices.visLayerLock).FormulaForceU = False
                    Next layerIndex
                End If
                'End Test if deep copy. 


                Dim allShapes As Visio.Selection = Nothing
                'Select all the shapes
                'Test if the page has any shapes.
                If vsoPage.Shapes.Count > 0 Then
                    'The page has shapes, select them. 
                    activeWindow.SelectAll()
                    allShapes = activeWindow.Selection
                    'Copy the selection
                    allShapes.Copy(Visio.VisCutCopyPasteCodes.visCopyPasteNoTranslate)
                End If
                'End Test if the page has any shapes.
                'If there are no shapes to select, allShapes = Nothing

                ' Create the new page    
                Dim newPage As Visio.Page = vsoPage.Document.Pages.Add
                Dim vsoNewPageSheet As Visio.Shape = newPage.PageSheet

                'Set the page's index.
                'Test if the page to be copied is a background page.
                If CBool(vsoPage.Background) = False Then
                    'It is not a background page, set the index = to the current index + 1
                    'This makes sure the page is copied next to its source.
                    newPage.Index = CShort(vsoPage.Index + 1)
                End If
                'End Test if the page to be copied is a background page.

                ' Create a proper name for the new page    
                'Limit the length of the page's name
                'We want to reserve enough space for the other naming work done below. 
                Dim charsToCopy As Integer = Len(vsoPage.Name)
                If (charsToCopy > maxCharsToCopy) Then
                    charsToCopy = maxCharsToCopy
                End If

                'Set the copy counter = 1 since we are assuming this is the first copy. 
                Dim copyCounter As Integer = 1

                Dim newNameBuilder As New StringBuilder(maxCharsToCopy + 10)
                'Set the new page's name as a "copy" of the old page
                newNameBuilder.Append(vsoPage.Name.Substring(0, charsToCopy))
                newNameBuilder.Append(" Copy (")
                newNameBuilder.Append((copyCounter.ToString))
                newNameBuilder.Append(")")
                'Test to see if the name is already in use.
TestName:
                'Pages are 1 indexed.
                For vsoPageIndex = 1 To vsoPage.Document.Pages.Count
                    'Test if the name is already in use.
                    If vsoPage.Document.Pages(vsoPageIndex).Name = newNameBuilder.ToString Then
                        'The name is already in use, increment the copy number and retest.
                        newNameBuilder.Remove(newNameBuilder.Length - (copyCounter.ToString.Length + 1), copyCounter.ToString.Length + 1)
                        copyCounter = copyCounter + 1
                        newNameBuilder.Append(copyCounter.ToString + ")")
                        GoTo TestName
                    End If
                    'End Test if the name is already in use.
                Next vsoPageIndex
                'Set the page's name.
                newPage.Name = newNameBuilder.ToString


                'Create the layers on the new page
                'Test if deep copy. 
                If deepCopy Then
                    Dim vsoRow As Visio.Row
                    'Loop though each layer in the source page.
                    For layerIndex = 1 To vsoPage.Layers.Count
                        'Add a new layer to the new page with the same name as the corresponding row on the current page
                        newPage.Layers.Add(vsoPage.Layers(layerIndex).Name)
                        vsoRow = vsoPage.PageSheet.Section(Visio.VisSectionIndices.visSectionLayer).Row(vsoPage.Layers(layerIndex).Row)
                        'Loop through each cell in the layer row, matching flags from the source page to the new page. 
                        For vsoCellIndex = 0 To vsoRow.Count - 1
                            'Set the flag on the new page = the corresponding flag on the old page
                            newPage.Layers(layerIndex).CellsC(vsoCellIndex).FormulaForce = vsoPage.Layers(layerIndex).CellsC(vsoCellIndex).Formula
                        Next vsoCellIndex
                    Next layerIndex
                    vsoRow = Nothing
                End If
                'End Test if deep copy. 

                'Paste the selection to the new page.
                If allShapes IsNot Nothing Then
                    newPage.Paste(Visio.VisCutCopyPasteCodes.visCopyPasteNoTranslate)
                    'Cleanup the selection
                    activeWindow.DeselectAll()
                End If

                'Restore the lock status of the pages.
                'Test if deep copy. 
                If deepCopy Then
                    'Loop through each layer in the source page. 
                    For layerIndex = 1 To vsoPage.Layers.Count
                        'Restore the layer lock status from the old page to the old page
                        vsoPage.Layers(layerIndex).CellsC(Visio.VisCellIndices.visLayerLock).FormulaForce = layerLockStatusArray(layerIndex - 1)
                        'Restore the layer lock status from the old page to the new page
                        newPage.Layers(layerIndex).CellsC(Visio.VisCellIndices.visLayerLock).FormulaForce = layerLockStatusArray(layerIndex - 1)
                    Next layerIndex
                End If
                'End Test if deep copy. 

                Dim vsoSourcePageSheet As Visio.Shape = vsoPage.PageSheet
                'Create custom cell formulas on the new page that aren't already there.
                Dim sectionsToSet(3) As Visio.VisSectionIndices
                sectionsToSet(0) = Visio.VisSectionIndices.visSectionProp
                sectionsToSet(1) = Visio.VisSectionIndices.visSectionUser
                sectionsToSet(2) = Visio.VisSectionIndices.visSectionAction

                Dim setFlag As Visio.VisGetSetArgs = _
                    Visio.VisGetSetArgs.visSetFormulas + _
                    Visio.VisGetSetArgs.visSetBlastGuards + _
                    Visio.VisGetSetArgs.visSetUniversalSyntax

                Dim setFlags(3) As Visio.VisGetSetArgs
                Dim setFlagsExtended(3) As visSetFlagsExtended
                For i = 0 To setFlags.Count - 1
                    setFlags(i) = setFlag
                    setFlagsExtended(i) = visSetFlagsExtended.visSetReplaceSelectiveAndAdd
                Next
                ShapeFunctions.SetShapeSection(vsoNewPageSheet, _
                                               ShapeFunctions.GetShapeSection(vsoSourcePageSheet, Visio.VisGetSetArgs.visGetFormulasU, sectionsToSet), _
                                               sectionsToSet, _
                                               setFlags, _
                                               setFlagsExtended)

                Return newPage
            Else
                'The page is Null, prompt the user with an error message.
                MsgBox("The Page to be copied must be in the active Visio Window.", , _
                       "Error. The requested operation cannot be performed.")
                Return Nothing
            End If
            'End Test if the page is Null

        Catch ex As Exception
            MsgBox("An error has occured while trying to copy the page: Source " & ex.Source & Environment.NewLine & _
                   "Error Line: " & Erl() & Environment.NewLine & _
                   "Error: (" & Err.Number & ") " & ex.Message, vbCritical)
            Return Nothing
        Finally
            'Nothing the free here.
            'Clear the clipboard
            System.Windows.Forms.Clipboard.Clear()
        End Try
    End Function

    ''' <summary>
    ''' Copies the page's shapes to another page.
    ''' </summary>
    ''' <param name="vsoPage">The page whose shapes we want to copy.</param>
    ''' <param name="destPage">The page to copy them to.</param>
    ''' <returns></returns>
    ''' 
    <Extension()> _
    Public Function CopyPageShapes(ByVal vsoPage As Visio.Page, _
                                   ByVal destPage As Visio.Page) _
                                   As Visio.Page

        Try
            'Test if the page is Null
            If vsoPage IsNot Nothing AndAlso destPage IsNot Nothing Then
                Dim activeWindow As Visio.Window
                'Windows Collections is 1 indexed.
                'Loop through all of the windows in the application, activate the one containing the source page. 
                For currWindow = 1 To vsoPage.Application.Windows.Count
                    If vsoPage.Application.Windows(currWindow).Document Is vsoPage.Document Then
                        vsoPage.Application.Windows(currWindow).Activate()
                    End If
                Next

                'The page is NOT Null, continue.
                activeWindow = vsoPage.Application.ActiveWindow

                Dim allShapes As Visio.Selection = Nothing
                'Select all the shapes
                'Test if the page has any shapes.
                If vsoPage.Shapes.Count > 0 Then
                    'The page has shapes, select them. 
                    activeWindow.SelectAll()
                    allShapes = activeWindow.Selection
                    'Copy the selection
                    allShapes.Copy(Visio.VisCutCopyPasteCodes.visCopyPasteNoTranslate)
                End If
                'End Test if the page has any shapes.

                Dim newPage As Visio.Page = Nothing
                'If there are no shapes to select, allShapes = Nothing
                'Paste the selection to the new page.
                If allShapes IsNot Nothing Then
                    ' Create the new page    
                    newPage = destPage
                    newPage.Paste(Visio.VisCutCopyPasteCodes.visCopyPasteNoTranslate)
                    'Cleanup the selection
                    activeWindow.DeselectAll()
                End If

                Return newPage
            Else
                'One of the pages is null, throw an exception.
                Dim ex As New ArgumentNullException("One or more of the parameters is invalid.")
                Throw ex
            End If
            'End Test if the page is Null

        Catch ex As Exception
            Dim outerEx As New CopyPageShapeException(ex)
            Throw outerEx
        Finally
            'Nothing the free here.
            'Clear the clipboard
            System.Windows.Forms.Clipboard.Clear()
        End Try
    End Function

    ''' <summary>
    ''' Adds the proper layers, removing any existing layers.
    ''' </summary>
    ''' <param name="vsoPage">The page which is having its layers replaced.</param>
    ''' <param name="layerNameArray">An array containing the names of the new layers.</param>
    <Extension()> _
    Public Sub ReplaceLayers(ByVal vsoPage As Visio.Page, _
                             ByVal layerNameArray() As String)
        Dim vsoPageSheet As Visio.Shape = vsoPage.PageSheet
        If Not vsoPageSheet.SectionExists(Visio.VisSectionIndices.visSectionLayer, Visio.VisExistsFlags.visExistsAnywhere) Then
            vsoPageSheet.AddSection(Visio.VisSectionIndices.visSectionLayer)
        End If

        'Delete all existing layers
        Dim layerCount As Integer = vsoPageSheet.Section(Visio.VisSectionIndices.visSectionLayer).Count
        For i = 0 To vsoPageSheet.Section(Visio.VisSectionIndices.visSectionLayer).Count - 1
            vsoPageSheet.DeleteRow(Visio.VisSectionIndices.visSectionLayer, 0)
        Next


        'Add each Layer
        For i = 0 To layerNameArray.Count - 1
            If vsoPageSheet.ContainingMaster IsNot Nothing Then
                vsoPageSheet.ContainingMaster.Layers.Add(layerNameArray(i))
            ElseIf vsoPageSheet.ContainingPage IsNot Nothing Then
                vsoPageSheet.ContainingPage.Layers.Add(layerNameArray(i))
            End If
        Next
    End Sub

    ''' <summary>
    ''' Tests for orphaned shapes (Shapes without masters)
    ''' </summary>
    ''' <param name="vsoPage">The vso page to test.</param>
    ''' <returns>
    ''' An array containing the shape.ID values for each 2-D orphaned shape on the page.
    ''' </returns>
    <Extension()> _
    Public Function TestForOrphaned2DShapes(ByVal vsoPage As Visio.Page) As Integer()
        Dim shapeList As New List(Of Integer)
        Try
            For Each vsoShape As Visio.Shape In vsoPage.Shapes
                If vsoShape.Master Is Nothing Then
                    If Not vsoShape.OneD Then
                        shapeList.Add(vsoShape.ID)
                    End If
                End If
            Next
            Return shapeList.ToArray
        Catch ex As Exception
            Return shapeList.ToArray
        End Try
    End Function
End Module

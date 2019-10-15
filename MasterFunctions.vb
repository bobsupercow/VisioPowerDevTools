Imports System.Runtime.CompilerServices

''' <summary>
''' Functions and Methods used to process data related to Visio.Master objects.
''' </summary>
Public Module MasterFunctions


    ''' <summary>
    ''' Adds the proper layers, removing any existing layers.
    ''' </summary>
    ''' <param name="vsoMaster">The master to add layers to.</param>
    ''' <param name="layerNameArray">An array containing the names of the new layers.</param>
    <Extension()> _
    Public Sub ReplaceLayers(ByVal vsoMaster As Visio.Master, _
                             ByVal layerNameArray() As String)
        Dim vsoPageSheet As Visio.Shape = vsoMaster.PageSheet
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
    ''' Replaces a master in the document stencil with a matching one from an external stencil.
    ''' </summary>
    ''' <param name="masterName">Name of the master to replace.</param>
    ''' <param name="vsoMasters">The masters collection which contains the master to be replaced.</param>
    ''' <param name="stencilFullPath">The full path of the stencil which contains the new version of the master.</param>
    ''' <param name="addIfNonExisting">Indicates whether to add a master from the external stencil even if it doesn't already exist in the masters collection.</param>
    ''' <param name="ignoreCase">Indicates whether or not the case of <paramref name=" masterName">masterName</paramref> matters when searching for the master.</param>
    ''' <returns><c>Nothing</c> if there was any type of error; else <c>the new master in the document stencil</c></returns>
    ''' <remarks>Assumes that the master being replaced has no shape instances currently in the document, else they will be orphaned.</remarks>
    <Extension()> _
    Public Function ReplaceMasterFromExternalStencil(ByVal vsoMasters As Visio.Masters, _
                                                     ByVal masterName As String, _
                                                     ByVal stencilFullPath As String, _
                                                     ByVal addIfNonExisting As Boolean, _
                                                     ByVal ignoreCase As Boolean) As Visio.Master

        Try
            'Get the old master in the document stencil, then delete it.
            Dim oldMaster As Visio.Master = Nothing
            If ignoreCase Then
                For Each vsoMaster As Visio.Master In vsoMasters
                    If (String.Equals(vsoMaster.Name, masterName, StringComparison.CurrentCultureIgnoreCase)) OrElse _
                        (String.Equals(vsoMaster.NameU, masterName, StringComparison.CurrentCultureIgnoreCase)) Then
                        oldMaster = vsoMaster
                        Exit For
                    End If
                Next
            Else
                Try
                    oldMaster = vsoMasters(masterName)
                Catch ex As System.Runtime.InteropServices.COMException
                    'Object name not found
                    If ex.ErrorCode = -2032465660 Then
                        'This is expected if no master is found with this name, simply continue execution here.
                    Else
                        'This is not expected, throw it up the chain.
                        Dim outerEx As New ReplaceMasterException("Master shape not found in destination stencil.", ex)
                        Throw ex
                    End If
                    'Throw anything else up the chain.
                End Try
            End If


            'Test whether or not we found a master by that name in the old stencil.
            If oldMaster IsNot Nothing Then
                'Master found, delete it.
                oldMaster.Delete()
            Else
                'Master not found, determine whether to add to the stencil anyway
                If addIfNonExisting = False Then
                    'The caller does not want to add the master if it doesn't already exist.
                    Dim ex As New ReplaceMasterException("Master shape not found in destination stencil.")
                    Throw ex
                End If
            End If

            'We will only reach this point in the code if either:
            '1. The master was found in the document stencil and we are ready to replace it.
            '2. The master was not found in the document stencil, but the caller wants to add it to the document stencil.


            'We want to open up the stencil to get the most recent revision of the master.
            'Get the stencil, opening it if it isn't open already.
            Dim externalStencil As Visio.Document = Nothing
            Dim externalStencilWasNothing As Boolean = False
            'See if the stencil is already open.
            For docIndex = 1 To vsoMasters.Application.Documents.Count - 0
                'If the filename matches the stencil, then set the variable
                If String.Equals(vsoMasters.Application.Documents(docIndex).FullName, stencilFullPath, StringComparison.CurrentCultureIgnoreCase) Then
                    externalStencil = vsoMasters.Application.Documents(docIndex)
                    Exit For
                End If
            Next docIndex
            'Open the stencil if it isn't already opened.
            If externalStencil Is Nothing Then
                externalStencilWasNothing = True
                externalStencil = vsoMasters.Application.Documents.OpenEx(stencilFullPath, _
                                                                  Visio.VisOpenSaveArgs.visOpenRW + Visio.VisOpenSaveArgs.visOpenDocked)
            End If

            'Get the new master from the stencil
            Dim newMaster As Visio.Master = Nothing
            If ignoreCase Then
                For Each vsoMaster As Visio.Master In externalStencil.Masters
                    If (String.Equals(vsoMaster.Name, masterName, StringComparison.CurrentCultureIgnoreCase)) OrElse _
                        (String.Equals(vsoMaster.NameU, masterName, StringComparison.CurrentCultureIgnoreCase)) Then
                        newMaster = vsoMaster
                        Exit For
                    End If
                Next
            Else
                Try
                    newMaster = externalStencil.Masters(masterName)
                Catch ex As System.Runtime.InteropServices.COMException
                    'Object name not found
                    If ex.ErrorCode = -2032465660 Then
                        'This is expected if no master is found with this name, simply continue execution here.
                    Else
                        'This is not expected, throw it up the chain.
                        Dim outerEx As New ReplaceMasterException("Master shape not found in source stencil.", ex)
                        Throw ex
                    End If
                    'Throw anything else up the chain.
                End Try
            End If


            'Make sure the master exists
            If newMaster Is Nothing Then
                'The master doesn't exist in the source stencil.
                Dim ex As New ReplaceMasterException("Master shape not found in source stencil.")
                Throw ex
            End If
            'Add the new master to the masters collection.
            vsoMasters.Drop(newMaster, 0, 0)
            'Close the stencil if it wasn't already open, we are done with it now.
            If externalStencilWasNothing Then
                externalStencil.Close()
            End If
            'Reset the variable
            newMaster = vsoMasters(masterName)
            'Return the new master in the document stencil.
            Return newMaster

        Catch ex As Exception
            'This is not expected, throw it up the chain.
            Dim outerEx As New ReplaceMasterException(ex)
            Throw outerEx
        End Try
    End Function

    ''' <summary>
    ''' Replaces a master shape and its instances with a new shape not in the document stencil.
    ''' </summary>
    ''' <param name="oldMaster">The master being replaced.</param>
    ''' <param name="newMaster">The new master.</param>
    ''' <returns>The new master.</returns>
    ''' <remarks>
    ''' Retains the following information for the instances.
    ''' The values of any common shape properties, its connections, and its size.
    ''' </remarks>
    <Extension()> _
    Public Function ReplaceMasterAndInstances(ByVal oldMaster As Visio.Master, _
                                              ByVal newMaster As Visio.Master) As Visio.Master
        Dim shapeX As Double
        Dim shapeY As Double
        Dim shapeHeight As Double
        Dim shapeWidth As Double
        Dim newShape As Visio.Shape
        Dim oldShape As Visio.Shape
        Dim shapesByPage As New List(Of List(Of String))
        Dim newMasterInDoc As Visio.Master
        Dim vsoDoc As Visio.Document = oldMaster.Document

        For Each vsoPage As Visio.Page In vsoDoc.Pages
            Dim shapesToReplace As New List(Of String)
            'Determine which instances need replaced on the page.
            If False = CBool(vsoPage.Background) Then
                For Each vsoShape As Visio.Shape In vsoPage.Shapes
                    If vsoShape.Master IsNot Nothing Then
                        If vsoShape.Master Is oldMaster Then
                            shapesToReplace.Add(vsoShape.UniqueID(Visio.VisUniqueIDArgs.visGetOrMakeGUID))
                        End If
                    End If
                Next
            End If
            'Add the list to the list for the current page.
            shapesByPage.Add(shapesToReplace)
        Next
        Dim oldNameU = oldMaster.NameU
        Dim oldName = oldMaster.Name
        'Delete the old master.
        oldMaster.Delete()
        'Add the newMaster
        newMasterInDoc = vsoDoc.Masters.Drop(newMaster, 0, 0)
        newMasterInDoc.NameU = oldNameU
        newMasterInDoc.Name = oldName

        For i = 0 To shapesByPage.Count - 1
            For j = 0 To shapesByPage(i).Count - 1
                oldShape = vsoDoc.Pages(i + 1).Shapes(shapesByPage(i)(j))

                'Save the values of the shape. 
                shapeX = oldShape.CellsSRC(Visio.VisSectionIndices.visSectionObject, _
                                           Visio.VisRowIndices.visRowXFormOut, _
                                           Visio.VisCellIndices.visXFormPinX).ResultIU
                shapeY = oldShape.CellsSRC(Visio.VisSectionIndices.visSectionObject, _
                                           Visio.VisRowIndices.visRowXFormOut, _
                                           Visio.VisCellIndices.visXFormPinY).ResultIU
                shapeHeight = oldShape.CellsSRC(Visio.VisSectionIndices.visSectionObject, _
                                                Visio.VisRowIndices.visRowXFormOut, _
                                                Visio.VisCellIndices.visXFormHeight).ResultIU
                shapeWidth = oldShape.CellsSRC(Visio.VisSectionIndices.visSectionObject, _
                                               Visio.VisRowIndices.visRowXFormOut, _
                                               Visio.VisCellIndices.visXFormWidth).ResultIU
                'Drop a new version of the shape.
                newShape = vsoDoc.Pages(i + 1).Drop(newMasterInDoc, shapeX, shapeY)
                newShape.CellsSRC(Visio.VisSectionIndices.visSectionObject, _
                                  Visio.VisRowIndices.visRowXFormOut, _
                                  Visio.VisCellIndices.visXFormHeight).FormulaU = shapeHeight
                newShape.CellsSRC(Visio.VisSectionIndices.visSectionObject, _
                                  Visio.VisRowIndices.visRowXFormOut, _
                                  Visio.VisCellIndices.visXFormWidth).FormulaU = shapeWidth

                'Copy the shape data
                VisioPowerDevTools.CopyShapeData(oldShape, _
                                                 newShape, _
                                                 Visio.VisGetSetArgs.visSetBlastGuards + Visio.VisGetSetArgs.visSetUniversalSyntax, _
                                                 VisioPowerDevTools.visCopyShapeDataArgs.protectReferences)


                MoveConnections(oldShape, newShape)


                'Delete the oldShape
                oldShape.Delete()
            Next
        Next

        Return newMasterInDoc
    End Function

End Module

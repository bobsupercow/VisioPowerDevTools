Imports System.Runtime.CompilerServices
Imports System.Text.RegularExpressions


''' <summary>
''' Functions and Methods used to process data related to 
''' <see href="http://msdn.microsoft.com/en-us/library/ms408994%28v=office.12%29.aspx">Visio.Shape</see> objects.
''' </summary>
Public Module ShapeFunctions

#Region "Clone Sections"

#Region "NamedConstantRows Sections"

    ''' <summary>
    ''' Gets the data from the  <see href="http://msdn.microsoft.com/en-us/library/ms408988%28v=office.12%29.aspx">Visio.Section</see>
    ''' (<paramref name="vsoSectionIndex"/>) of the shape (<paramref name="vsoShape"/>) as an <see cref="ArrayList"/>
    ''' of <see cref="ArrayList"/> objects. This function supports any section that has named rows. 
    ''' <seealso href="http://msdn.microsoft.com/en-us/library/bb902804%28v=office.12%29.aspx#Y2964">Supported Sections</seealso>
    ''' </summary>
    ''' <param name="vsoShape">The <see href="http://msdn.microsoft.com/en-us/library/ms408994%28v=office.12%29.aspx">Visio.Shape</see> to get data from.</param>
    ''' <param name="getFlag">Indicates whether to get values instead of formulas as well as the format of the results.</param>
    ''' <param name="vsoSectionIndex">The <see href="http://msdn.microsoft.com/en-us/library/ms408988%28v=office.12%29.aspx">Visio.Section</see> to get data from.</param>
    ''' <returns>
    ''' <c>Nothing</c> if there is no data; otherwise
    ''' an <see cref="ArrayList"/> which contains the data for each row in the following format.
    ''' <para>list(0) = An <see cref="Array"/> which contains the cell indicies the data has been gathered for.</para>
    ''' <para>list(1) = The name of the first row as a <see cref="String"/>.</para>
    ''' <para>list(2) = An <see cref="Array"/> which contains the cell values or formulas for each cell index in list(0) for the row in list(1)</para>
    ''' <para>list(3) = The name of the second row as a <see cref="String"/>.</para>
    ''' <para>list(4) = An <see cref="Array"/> which contains the cell values or formulas for each cell index in list(0) for the row in list(3)</para>
    ''' </returns>
    ''' <remarks>
    ''' Supports only sections which have named rows. 
    ''' <para>
    ''' Returns the cell values as strings if <paramref name="getVals"/> = <c>True</c>.
    ''' </para>
    ''' <para>
    ''' Returns the cell formulas as strings if <paramref name="getVals"/> = <c>False</c>.
    ''' </para>
    ''' </remarks>
    Private Function GetSectionNamedRowsConstantCells(ByVal vsoShape As Visio.Shape, _
                                                      ByVal getFlag As Visio.VisGetSetArgs, _
                                                      ByVal vsoSectionIndex As Visio.VisSectionIndices, _
                                                      ByVal vsoUnitsNamesOrCodes() As Object) _
                                                      As List(Of Object)
        Try
            'Test whether the given section even exists.
            If vsoShape.SectionExists(vsoSectionIndex, Visio.VisExistsFlags.visExistsAnywhere) Then
                'It exists, Test whether the section contains any rows.

                Dim rowCount As Short = vsoShape.RowCount(vsoSectionIndex)
                If rowCount > 0 Then
                    'There ARE rows present, return the data
                    Dim returnList As New List(Of Object)
                    'Add the sectionIndex to the first element of the outer list. 
                    returnList.Add(vsoSectionIndex)
                    'Get the cell indicies for the section.
                    'Uses visRowFirst for all Named Sections, visTagDefault is ignored in this case.
                    Dim visCellIndiciesArray() As Visio.VisCellIndices = _
                        GetCellIndicies(vsoSectionIndex, Visio.VisRowIndices.visRowFirst, Visio.VisRowTags.visTagDefault)

                    'Add the Cell Indicies Array to the second element of the outer list.
                    returnList.Add(visCellIndiciesArray)

                    'Create the SRCStream
                    Dim vsoSRCStreamArray(((visCellIndiciesArray.Count * 3) * rowCount) - 1) As Short
                    Dim vsoReturnArray((visCellIndiciesArray.Count * rowCount) - 1) As Object
                    Dim vsoRowNamesArray(rowCount - 1) As String
                    Dim currentFormula As Integer = 0
                    'Loop through each Row in the Section.
                    For rowIndex = 0 To vsoShape.RowCount(vsoSectionIndex) - 1
                        'Add RowName to the rowName Array
                        vsoRowNamesArray(rowIndex) = vsoShape.Section(vsoSectionIndex).Row(rowIndex).NameU

                        For i = 0 To visCellIndiciesArray.Count - 1
                            vsoSRCStreamArray(currentFormula) = vsoSectionIndex
                            currentFormula += 1
                            vsoSRCStreamArray(currentFormula) = rowIndex
                            currentFormula += 1
                            vsoSRCStreamArray(currentFormula) = visCellIndiciesArray(i)
                            currentFormula += 1
                        Next
                    Next rowIndex

                    'Add RowNamesArray to the return List
                    returnList.Add(vsoRowNamesArray)
                    'Test if the caller wants values or formulas returned.
                    Select Case getFlag
                        Case _
                            Visio.VisGetSetArgs.visGetFloats, _
                            Visio.VisGetSetArgs.visGetRoundedInts, _
                            Visio.VisGetSetArgs.visGetStrings, _
                            Visio.VisGetSetArgs.visGetTruncatedInts

                            'Results
                            'Execute the StreamGet
                            vsoShape.GetResults(vsoSRCStreamArray, _
                                                getFlag, _
                                                vsoUnitsNamesOrCodes, _
                                                vsoReturnArray)
                            'Add the results Array to the returnList
                            returnList.Add(vsoReturnArray)

                        Case Visio.VisGetSetArgs.visGetFormulas
                            'Formulas
                            'Execute the StreamGet
                            vsoShape.GetFormulas(vsoSRCStreamArray, _
                                                 vsoReturnArray)
                            'Add the Formula Array to the returnList
                            returnList.Add(vsoReturnArray)

                        Case Visio.VisGetSetArgs.visGetFormulasU
                            'FormulasU
                            'Execute the StreamGet
                            vsoShape.GetFormulasU(vsoSRCStreamArray, _
                                                  vsoReturnArray)
                            'Add the Formula Array to the returnList
                            returnList.Add(vsoReturnArray)

                        Case Else
                            'Invalid arguments, Return Nothing
                            Return Nothing

                    End Select
                    'Return the returnList
                    Return returnList
                Else
                    'There are no rows, return nothing. 
                    Return Nothing
                End If
                'End Test whether the section contains any rows.
            Else
                'It does not exist, return nothing.
                Return Nothing
            End If
            'End Test whether the given section even exists.

            'Catch ex As Exception
            'CODE EXCEPTION HANDLER HERE
            'Return nothing in the event of an error.
            'Return Nothing
        Finally
            'Nothing to free here. 
        End Try
    End Function

    ''' <summary>
    ''' Sets the data from the  <see href="http://msdn.microsoft.com/en-us/library/ms408988%28v=office.12%29.aspx">Visio.Section</see>
    ''' (<paramref name="vsoSectionIndex"/>) of the shape (<paramref name="vsoShape"/>) from an <see cref="ArrayList"/>
    ''' of <see cref="ArrayList"/> objects. This function supports any section that has named rows. 
    ''' <seealso href="http://msdn.microsoft.com/en-us/library/bb902804%28v=office.12%29.aspx#Y2964">Supported Sections</seealso>
    ''' </summary>
    ''' <param name="vsoShape">The <see href="http://msdn.microsoft.com/en-us/library/ms408994%28v=office.12%29.aspx">Visio.Shape</see> to set data for.</param>
    ''' <param name="returnedList">
    ''' An <see cref="ArrayList"/> which contains the data for each row in the following format.
    ''' <para>list(0) = An <see cref="Array"/> which contains the cell indicies the data has been gathered for.</para>
    ''' <para>list(1) = The name of the first row as a <see cref="String"/>.</para>
    ''' <para>list(2) = An <see cref="Array"/> which contains the cell values or formulas for each cell index in list(0) for the row in list(1)</para>
    ''' <para>list(3) = The name of the second row as a <see cref="String"/>.</para>
    ''' <para>list(4) = An <see cref="Array"/> which contains the cell values or formulas for each cell index in list(0) for the row in list(3)</para>
    ''' </param> 
    ''' <param name="setFlagsExtended">Indicates whether or not to add rows to the shape if they don't already exist.</param>
    ''' <returns>
    ''' <c>True</c> if set successfully; otherwise <c>False</c>
    ''' </returns>
    ''' <remarks>
    ''' Supports only sections which have named rows. 
    ''' </remarks>
    Private Function SetSectionNamedRowsConstantCells(ByVal vsoShape As Visio.Shape, _
                                                      ByVal returnedList As List(Of Object), _
                                                      ByVal setFlags As Integer, _
                                                      ByVal setFlagsExtended As visSetFlagsExtended, _
                                                      ByVal vsoUnitsNamesOrCodes() As Object) _
                                                      As Boolean

        Try
            If returnedList Is Nothing Then
                Return False
            End If

            'Get the Section Index from the list.
            Dim vsoSectionIndex As Visio.VisSectionIndices = returnedList(0)
            'Get the Cell Indicies from the list.
            Dim visCellIndiciesArray() As Visio.VisCellIndices = returnedList(1)
            'Get the Row Names from the list.
            Dim vsoRowNamesArray() As String = returnedList(2)
            'Get the FormulaArray from the list.
            Dim vsoReturnedArray() As Object = returnedList(3)


            'Create the SRCStream
            Dim currentStreamIndex As Integer = 0
            Dim vsoSRCStreamArray(((visCellIndiciesArray.Count * vsoRowNamesArray.Count) * 3) - 1) As Short
            Dim currentRow As Integer
            Dim cellName As String
            Dim currentFormula As Integer
            'In order to get the correct shapesheet row, we must see if it already exists.
            'Determine the appropriate add action based on the extended flags. 
            Select Case setFlagsExtended
                Case visSetFlagsExtended.visSetReplaceAllExisting
                    'Shapes that are instances of masters automatically have any row contained in the master
                    'recreated in the instance when a section that has been cleared is recreated. 
                    'As such, there is not a way to truly replace all of that kind of shape. 
                    'Test if the shape is a master shape. 
                    If vsoShape.Master Is Nothing Then
                        'The shape is not inherited continue.
                        'Since we have already cleared the section, we know we need to add everything.
                        'Loop through the rowNames
                        For i = 0 To vsoRowNamesArray.Count - 1
                            'It doesn't exist, add the row and set the name. 
                            currentRow = vsoShape.AddNamedRow(vsoSectionIndex, _
                                                              vsoRowNamesArray(i), _
                                                              Visio.VisRowTags.visTagDefault)

                            For j = 0 To visCellIndiciesArray.Count - 1
                                vsoSRCStreamArray(currentStreamIndex) = vsoSectionIndex
                                currentStreamIndex += 1
                                vsoSRCStreamArray(currentStreamIndex) = currentRow
                                currentStreamIndex += 1
                                vsoSRCStreamArray(currentStreamIndex) = visCellIndiciesArray(j)
                                currentStreamIndex += 1
                            Next
                        Next
                    Else
                        Return False
                        'The shape is inherited from a master.
                    End If
                Case visSetFlagsExtended.visSetReplaceSelectiveAndAdd
                    'Loop through the rowNames
                    For i = 0 To vsoRowNamesArray.Count - 1
                        'Get the cell name, with it's prefix string.
                        cellName = GetSectionPrefix(vsoSectionIndex) & vsoRowNamesArray(i)
                        'Test whether a row exists with that name.
                        If vsoShape.CellExistsU(cellName, Visio.VisExistsFlags.visExistsAnywhere) Then
                            'It exists, set the row and continue.
                            currentRow = vsoShape.CellsU(cellName).Row
                        Else
                            'It doesn't exist, add the row and set the name. 
                            currentRow = vsoShape.AddNamedRow(vsoSectionIndex, _
                                                              vsoRowNamesArray(i), _
                                                              Visio.VisRowTags.visTagDefault)
                        End If

                        For j = 0 To visCellIndiciesArray.Count - 1
                            vsoSRCStreamArray(currentStreamIndex) = vsoSectionIndex
                            currentStreamIndex += 1
                            vsoSRCStreamArray(currentStreamIndex) = currentRow
                            currentStreamIndex += 1
                            vsoSRCStreamArray(currentStreamIndex) = visCellIndiciesArray(j)
                            currentStreamIndex += 1
                        Next

                    Next
                Case visSetFlagsExtended.visSetReplaceSelectiveAndIgnore
                    'We need to make sure the formula array range is also ignored.
                    Dim tempList As List(Of Object) = vsoReturnedArray.ToList()
                    currentFormula = 0
                    'Loop through the rowNames
                    For i = 0 To vsoRowNamesArray.Count - 1
                        'Get the cell name, with it's prefix string.
                        cellName = GetSectionPrefix(vsoSectionIndex) & vsoRowNamesArray(i)
                        'Test whether a row exists with that name.
                        If vsoShape.CellExistsU(cellName, Visio.VisExistsFlags.visExistsAnywhere) Then
                            'It exists, set the row and continue.
                            currentRow = vsoShape.CellsU(cellName).Row
                            For j = 0 To visCellIndiciesArray.Count - 1
                                vsoSRCStreamArray(currentStreamIndex) = vsoSectionIndex
                                currentStreamIndex += 1
                                vsoSRCStreamArray(currentStreamIndex) = currentRow
                                currentStreamIndex += 1
                                vsoSRCStreamArray(currentStreamIndex) = visCellIndiciesArray(j)
                                currentStreamIndex += 1
                            Next
                            currentFormula += visCellIndiciesArray.Count
                        Else
                            'It doesn't exit, ignore it.
                            'We need to make sure the formula array range is also ignored.
                            tempList.RemoveRange(currentFormula, visCellIndiciesArray.Count)
                        End If
                    Next
                    'We need to make sure the formula array range is also ignored.
                    vsoReturnedArray = tempList.ToArray
            End Select
            'End Determine the appropriate add action based on the extended flags. 


            'Execute the SRCStream based on the set flags.
            Select Case setFlags
                Case _
                    Visio.VisGetSetArgs.visSetBlastGuards, _
                    Visio.VisGetSetArgs.visSetTestCircular, _
                    Visio.VisGetSetArgs.visSetBlastGuards + Visio.VisGetSetArgs.visSetTestCircular
                    'SetResults
                    vsoShape.SetResults(vsoSRCStreamArray, _
                                        vsoUnitsNamesOrCodes, _
                                        vsoReturnedArray, _
                                        setFlags)
                Case _
                    Visio.VisGetSetArgs.visSetFormulas, _
                    Visio.VisGetSetArgs.visSetFormulas + Visio.VisGetSetArgs.visSetUniversalSyntax, _
                    Visio.VisGetSetArgs.visSetFormulas + Visio.VisGetSetArgs.visSetTestCircular, _
                    Visio.VisGetSetArgs.visSetFormulas + Visio.VisGetSetArgs.visSetBlastGuards, _
                    Visio.VisGetSetArgs.visSetFormulas + Visio.VisGetSetArgs.visSetUniversalSyntax + Visio.VisGetSetArgs.visSetTestCircular, _
                    Visio.VisGetSetArgs.visSetFormulas + Visio.VisGetSetArgs.visSetUniversalSyntax + Visio.VisGetSetArgs.visSetBlastGuards, _
                    Visio.VisGetSetArgs.visSetFormulas + Visio.VisGetSetArgs.visSetUniversalSyntax + Visio.VisGetSetArgs.visSetTestCircular + Visio.VisGetSetArgs.visSetTestCircular
                    'Set Formulas
                    vsoShape.SetFormulas(vsoSRCStreamArray, _
                                         vsoReturnedArray, _
                                         setFlags - Visio.VisGetSetArgs.visSetFormulas)
                Case Else
                    'Invalid Flags
                    Return False
            End Select
            ' End Execute the SRCStream based on the set flags.

            'Success
            Return True

            'Catch ex As Exception
            'CODE EXCEPTION HANDLER HERE
            'Return False
        Finally
            'Nothing to free here.
        End Try
    End Function

#End Region

#Region "UnnamedConstantRows Sections"

    Private Function GetSectionUnnamedConstantRows(ByVal vsoShape As Visio.Shape, _
                                                   ByVal getFlag As Visio.VisGetSetArgs, _
                                                   ByVal vsoSectionIndex As Visio.VisSectionIndices, _
                                                   ByVal vsoUnitsNamesOrCodes() As Object) _
                                                   As List(Of Object)

        'Get the data from a shape, optionally getting all cells
        Try
            'Test whether the given section even exists.
            If vsoShape.SectionExists(vsoSectionIndex, Visio.VisExistsFlags.visExistsAnywhere) Then
                'It exists, Test whether the section contains any rows.
                If vsoShape.RowCount(vsoSectionIndex) > 0 Then
                    'There ARE rows present, return the data
                    Dim returnList As New List(Of Object)
                    Dim visCellIndiciesArraysList As New List(Of Visio.VisCellIndices())
                    'Get the row indicies for the section.
                    Dim visRowIndiciesArray() As Visio.VisRowIndices = GetRowIndicies(vsoSectionIndex)
                    Dim vsoExistingRows As New List(Of Visio.VisRowIndices)
                    'Set the formulaCount = 0 
                    Dim formulaCount As Integer = 0
                    'Get the cell indicies for each present row into an array of arrays.
                    For i = 0 To visRowIndiciesArray.Count - 1
                        'Only add rows that already exist. 
                        If vsoShape.RowExists(vsoSectionIndex, visRowIndiciesArray(i), Visio.VisExistsFlags.visExistsAnywhere) Then
                            vsoExistingRows.Add(visRowIndiciesArray(i))
                            visCellIndiciesArraysList.Add(GetCellIndicies(vsoSectionIndex, visRowIndiciesArray(i), Visio.VisRowTags.visTagDefault))
                            formulaCount += visCellIndiciesArraysList.Last.Length
                        End If
                    Next
                    'Create the SRCStream
                    Dim vsoSRCStreamArray((formulaCount * 3) - 1) As Short
                    Dim vsoReturnArray(formulaCount - 1) As Object
                    Dim currentFormula As Integer = 0
                    'Loop through each rowIndex
                    For i = 0 To vsoExistingRows.Count - 1
                        'Get the Cell Indicies for that row.
                        For j = 0 To visCellIndiciesArraysList(i).Count - 1
                            vsoSRCStreamArray(currentFormula) = vsoSectionIndex
                            currentFormula += 1
                            vsoSRCStreamArray(currentFormula) = vsoExistingRows(i)
                            currentFormula += 1
                            vsoSRCStreamArray(currentFormula) = visCellIndiciesArraysList(i)(j)
                            currentFormula += 1
                        Next
                    Next

                    'Add the stream to the list so we don't have to recreate it during the set routine.
                    returnList.Add(vsoSRCStreamArray)
                    'Test if the caller wants values or formulas returned.
                    Select Case getFlag
                        Case _
                            Visio.VisGetSetArgs.visGetFloats, _
                            Visio.VisGetSetArgs.visGetRoundedInts, _
                            Visio.VisGetSetArgs.visGetStrings, _
                            Visio.VisGetSetArgs.visGetTruncatedInts

                            'Results
                            'Execute the StreamGet
                            vsoShape.GetResults(vsoSRCStreamArray, _
                                                getFlag, _
                                                vsoUnitsNamesOrCodes, _
                                                vsoReturnArray)
                            'Add the results Array to the returnList
                            returnList.Add(vsoReturnArray)

                        Case Visio.VisGetSetArgs.visGetFormulas
                            'Formulas
                            'Execute the StreamGet
                            vsoShape.GetFormulas(vsoSRCStreamArray, _
                                                 vsoReturnArray)
                            'Add the Formula Array to the returnList
                            returnList.Add(vsoReturnArray)

                        Case Visio.VisGetSetArgs.visGetFormulasU
                            'FormulasU
                            'Execute the StreamGet
                            vsoShape.GetFormulasU(vsoSRCStreamArray, _
                                                  vsoReturnArray)
                            'Add the Formula Array to the returnList
                            returnList.Add(vsoReturnArray)

                        Case Else
                            'Invalid arguments, Return Nothing
                            Return Nothing

                    End Select
                    'Return the returnList
                    Return returnList
                Else
                    'There are no rows, return nothing. 
                    Return Nothing
                End If
                'End Test whether the section contains any rows.
            Else
                'It does not exist, return nothing.
                Return Nothing
            End If
            'End Test whether the given section even exists.

            'Catch ex As Exception
            'CODE EXCEPTION HANDLER HERE
            'Return nothing in the event of an error.
            'Return Nothing
        Finally
            'Nothing to free here. 
        End Try
    End Function

    Private Function SetSectionUnnamedConstantRows(ByVal vsoShape As Visio.Shape, _
                                                   ByVal returnedList As List(Of Object), _
                                                   ByVal setFlags As Integer, _
                                                   ByVal setFlagsExtended As visSetFlagsExtended, _
                                                   ByVal vsoUnitsNamesOrCodes() As Object) _
                                                   As Boolean

        Try
            If returnedList Is Nothing Then
                Return False
            End If


            'Get the stream from the list.
            Dim vsoSRCStream() As Short = returnedList(0)
            'Get the formulas from the list.
            Dim vsoFormulaArray() As Object = returnedList(1)
            'Get the sectionIndex from the SRC Stream
            Dim vsoSectionIndex As Visio.VisSectionIndices = vsoSRCStream(0)



            'Determine the appropriate add action based on the extended flags. 
            Select Case setFlagsExtended
                Case _
                    visSetFlagsExtended.visSetReplaceAllExisting, _
                    visSetFlagsExtended.visSetReplaceSelectiveAndAdd

                    'Loop through the SRC Stream one cellSet at a time.
                    For i = 1 To vsoSRCStream.Count - 1 Step 3
                        'Add any rows that don't already exist. 
                        If Not vsoShape.RowExists(vsoSectionIndex, vsoSRCStream(i), Visio.VisExistsFlags.visExistsAnywhere) Then
                            vsoShape.AddRow(vsoSectionIndex, vsoSRCStream(i), Visio.VisRowTags.visTagDefault)
                        End If
                    Next
                Case Else
                    'Both SetResults & SetFormulas blast & ignore missing cells by defauls, 
                    'no need to modify to stream or the shapesheet in this case.
            End Select
            'End Determine the appropriate add action based on the extended flags. 

            'Execute the SRCStream based on the set flags.
            Select Case setFlags
                Case _
                    Visio.VisGetSetArgs.visSetBlastGuards, _
                    Visio.VisGetSetArgs.visSetTestCircular, _
                    Visio.VisGetSetArgs.visSetBlastGuards + Visio.VisGetSetArgs.visSetTestCircular
                    'Set Results
                    vsoShape.SetResults(vsoSRCStream, _
                                        vsoUnitsNamesOrCodes, _
                                        vsoFormulaArray, _
                                        setFlags)
                Case _
                    Visio.VisGetSetArgs.visSetFormulas, _
                    Visio.VisGetSetArgs.visSetFormulas + Visio.VisGetSetArgs.visSetUniversalSyntax, _
                    Visio.VisGetSetArgs.visSetFormulas + Visio.VisGetSetArgs.visSetTestCircular, _
                    Visio.VisGetSetArgs.visSetFormulas + Visio.VisGetSetArgs.visSetBlastGuards, _
                    Visio.VisGetSetArgs.visSetFormulas + Visio.VisGetSetArgs.visSetUniversalSyntax + Visio.VisGetSetArgs.visSetTestCircular, _
                    Visio.VisGetSetArgs.visSetFormulas + Visio.VisGetSetArgs.visSetUniversalSyntax + Visio.VisGetSetArgs.visSetBlastGuards, _
                    Visio.VisGetSetArgs.visSetFormulas + Visio.VisGetSetArgs.visSetUniversalSyntax + Visio.VisGetSetArgs.visSetTestCircular + Visio.VisGetSetArgs.visSetTestCircular
                    'Set Formulas
                    vsoShape.SetFormulas(vsoSRCStream, _
                                         vsoFormulaArray, _
                                         setFlags - Visio.VisGetSetArgs.visSetFormulas)

                Case Else
                    'Invalid SetFlags
                    Return False
            End Select
            'End Execute the SRCStream based on the set flags.
            'Success
            Return True


            'Catch ex As Exception
            'CODE EXCEPTION HANDLER HERE
            'Return False
        Finally
            'Nothing to free here.
        End Try
    End Function

#End Region

#Region "UnnamedNonConstantRows Sections"

    Private Function GetSectionUnnamedNonConstantRows(ByVal vsoShape As Visio.Shape, _
                                                      ByVal getFlag As Visio.VisGetSetArgs, _
                                                      ByVal vsoSectionIndex As Visio.VisSectionIndices, _
                                                      ByVal vsoUnitsNamesOrCodes() As Object) _
                                                      As List(Of Object)

        Try
            'Test whether the given section even exists.
            If vsoShape.SectionExists(vsoSectionIndex, Visio.VisExistsFlags.visExistsAnywhere) Then
                'It exists, Test whether the section contains any rows.
                If vsoShape.RowCount(vsoSectionIndex) > 0 Then
                    'There ARE rows present, return the data
                    Dim returnList As New List(Of Object)
                    'Get the row indicies for the section.
                    Dim visRowIndiciesArray() As Visio.VisRowIndices = GetRowIndicies(vsoSectionIndex)
                    'Get the cell indicies for the row's present in the section.

                    'Since all the rows in this section type have the same cell indicies., this is safe.
                    Dim visCellIndiciesArray() As Visio.VisCellIndices = _
                        GetCellIndicies(vsoSectionIndex, visRowIndiciesArray(0), Visio.VisRowTags.visTagDefault)

                    Dim formulaCount As Integer = visCellIndiciesArray.Count * vsoShape.RowCount(vsoSectionIndex)
                    Dim currentFormula As Integer = 0
                    'Create the SRCStream
                    Dim vsoSRCStreamArray((formulaCount * 3) - 1) As Short
                    Dim vsoReturnArray(formulaCount - 1) As Object
                    'Loop through each rowIndex
                    For rowIndex = 0 To vsoShape.RowCount(vsoSectionIndex) - 1
                        'Get the Cell Indicies for that row.
                        For cellIndex = 0 To visCellIndiciesArray.Count - 1
                            vsoSRCStreamArray(currentFormula) = vsoSectionIndex
                            currentFormula += 1
                            vsoSRCStreamArray(currentFormula) = rowIndex
                            currentFormula += 1
                            vsoSRCStreamArray(currentFormula) = visCellIndiciesArray(cellIndex)
                            currentFormula += 1
                        Next
                    Next

                    'Add the stream to the list so we don't have to recreate it during the set routine.
                    returnList.Add(vsoSRCStreamArray)
                    'Test if the caller wants values or formulas returned.
                    Select Case getFlag
                        Case _
                            Visio.VisGetSetArgs.visGetFloats, _
                            Visio.VisGetSetArgs.visGetRoundedInts, _
                            Visio.VisGetSetArgs.visGetStrings, _
                            Visio.VisGetSetArgs.visGetTruncatedInts

                            'Results
                            'Execute the StreamGet
                            vsoShape.GetResults(vsoSRCStreamArray, _
                                                getFlag, _
                                                vsoUnitsNamesOrCodes, _
                                                vsoReturnArray)
                            'Add the results Array to the returnList
                            returnList.Add(vsoReturnArray)

                        Case Visio.VisGetSetArgs.visGetFormulas
                            'Formulas
                            'Execute the StreamGet
                            vsoShape.GetFormulas(vsoSRCStreamArray, _
                                                 vsoReturnArray)
                            'Add the Formula Array to the returnList
                            returnList.Add(vsoReturnArray)

                        Case Visio.VisGetSetArgs.visGetFormulasU
                            'FormulasU
                            'Execute the StreamGet
                            vsoShape.GetFormulasU(vsoSRCStreamArray, _
                                                  vsoReturnArray)
                            'Add the Formula Array to the returnList
                            returnList.Add(vsoReturnArray)

                        Case Else
                            'Invalid arguments, Return Nothing
                            Return Nothing

                    End Select
                    'Return the returnList
                    Return returnList
                Else
                    'There are no rows, return nothing. 
                    Return Nothing
                End If
                'End Test whether the section contains any rows.
            Else
                'It does not exist, return nothing.
                Return Nothing
            End If
            'End Test whether the given section even exists.

            'Catch ex As Exception
            'CODE EXCEPTION HANDLER HERE
            'Return nothing in the event of an error.
            'Return Nothing
        Finally
            'Nothing to free here. 
        End Try
    End Function

    Private Function SetSectionUnnamedNonConstantRows(ByVal vsoShape As Visio.Shape, _
                                                      ByVal returnList As List(Of Object), _
                                                      ByVal setFlags As Visio.VisGetSetArgs, _
                                                      ByVal setFlagsExtended As visSetFlagsExtended, _
                                                      ByVal vsoUnitsNamesOrCodes() As Object) _
                                                      As Boolean

        Try
            If returnList Is Nothing Then
                Return False
            End If

            'Get the stream from the list.
            Dim vsoSRCStreamArray() As Short = returnList(0)
            'Get the formulas from the list.
            Dim vsoReturnedArray() As Object = returnList(1)
            'Get the sectionIndex from the SRC Array
            Dim vsoSectionIndex As Visio.VisSectionIndices = vsoSRCStreamArray(0)

            'Determine the appropriate add action based on the extended flags. 
            Select Case setFlagsExtended
                Case _
                    visSetFlagsExtended.visSetReplaceAllExisting, _
                    visSetFlagsExtended.visSetReplaceSelectiveAndAdd

                    'Loop through the SRC Stream one cellSet at a time.
                    For i = 1 To vsoSRCStreamArray.Count - 1 Step 3
                        'Add any rows that don't already exist. 
                        If Not vsoShape.RowExists(vsoSectionIndex, vsoSRCStreamArray(i), Visio.VisExistsFlags.visExistsAnywhere) Then
                            vsoShape.AddRow(vsoSectionIndex, vsoSRCStreamArray(i), Visio.VisRowTags.visTagDefault)
                        End If
                    Next
                Case Else
                    'Both SetResults & SetFormulas blast & ignore missing cells by defauls, 
                    'no need to modify to stream or the shapesheet in this case.
            End Select
            'End Determine the appropriate add action based on the extended flags. 



            'Execute the SRCStream based on the set flags.
            Select Case setFlags
                Case _
                    Visio.VisGetSetArgs.visSetBlastGuards, _
                    Visio.VisGetSetArgs.visSetTestCircular, _
                    Visio.VisGetSetArgs.visSetBlastGuards + Visio.VisGetSetArgs.visSetTestCircular
                    'Set Results
                    vsoShape.SetResults(vsoSRCStreamArray, _
                                        vsoUnitsNamesOrCodes, _
                                        vsoReturnedArray, _
                                        setFlags)
                Case _
                    Visio.VisGetSetArgs.visSetFormulas, _
                    Visio.VisGetSetArgs.visSetFormulas + Visio.VisGetSetArgs.visSetUniversalSyntax, _
                    Visio.VisGetSetArgs.visSetFormulas + Visio.VisGetSetArgs.visSetTestCircular, _
                    Visio.VisGetSetArgs.visSetFormulas + Visio.VisGetSetArgs.visSetBlastGuards, _
                    Visio.VisGetSetArgs.visSetFormulas + Visio.VisGetSetArgs.visSetUniversalSyntax + Visio.VisGetSetArgs.visSetTestCircular, _
                    Visio.VisGetSetArgs.visSetFormulas + Visio.VisGetSetArgs.visSetUniversalSyntax + Visio.VisGetSetArgs.visSetBlastGuards, _
                    Visio.VisGetSetArgs.visSetFormulas + Visio.VisGetSetArgs.visSetUniversalSyntax + Visio.VisGetSetArgs.visSetTestCircular + Visio.VisGetSetArgs.visSetTestCircular
                    'Set Formulas
                    vsoShape.SetFormulas(vsoSRCStreamArray, _
                                         vsoReturnedArray, _
                                         setFlags - Visio.VisGetSetArgs.visSetFormulas)

                Case Else
                    'Invalid SetFlags
                    Return False
            End Select
            'End Execute the SRCStream based on the set flags.
            'Success
            Return True

            'Catch ex As Exception
            'CODE EXCEPTION HANDLER HERE
            'Return False
        Finally
            'Nothing to free here. 
        End Try
    End Function

#End Region

#Region "UnnamedNonConstantRowsAndCells Sections"

    Private Function GetSectionUnnamedNonConstantRowsAndCells(ByVal vsoShape As Visio.Shape, _
                                                              ByVal getFlag As Visio.VisGetSetArgs, _
                                                              ByVal vsoSectionIndex As Visio.VisSectionIndices, _
                                                              ByVal vsoUnitsNamesOrCodes() As Object) _
                                                              As List(Of Object)

        Try
            'Test whether the given section even exists.
            If vsoShape.SectionExists(vsoSectionIndex, Visio.VisExistsFlags.visExistsAnywhere) Then
                Dim rowCount As Integer = vsoShape.RowCount(vsoSectionIndex)

                'It exists, Test whether the section contains any rows.
                If rowCount > 0 Then
                    'There ARE rows present, return the data
                    Dim returnList As New List(Of Object)
                    'Add the sectionIndex to the first element in the list.
                    returnList.Add(vsoSectionIndex)
                    'Get the cell indicies for the row's present in the section.
                    Dim visRowTagArray(rowCount - 1) As Visio.VisRowTags
                    Dim visCellIndiciesArraysList As New List(Of Visio.VisCellIndices())


                    'Set the formulaCount = 0 
                    Dim formulaCount As Integer = 0
                    'Since the rows may have different cells, we need to use the RowTag attribute when getting the indicies. 
                    'Loop through each Row in the Section.
                    For rowIndex = 0 To rowCount - 1
                        'Add RowTag to the rowTag array. 
                        visRowTagArray(rowIndex) = vsoShape.RowType(vsoSectionIndex, rowIndex)
                        'Add the CellIndicies for the row to the list as an array.
                        visCellIndiciesArraysList.Add(GetCellIndicies(vsoSectionIndex, rowIndex, vsoShape.RowType(vsoSectionIndex, rowIndex)))
                        'Increment the formula count appropriately
                        formulaCount += visCellIndiciesArraysList.Last.Count
                    Next rowIndex

                    'Add the RowTagArray to the list.
                    returnList.Add(visRowTagArray)
                    'Add the cellIndiciesArraysList to the list.
                    returnList.Add(visCellIndiciesArraysList)

                    'Create the SRCStream
                    Dim vsoSRCStreamArray((formulaCount * 3) - 1) As Short
                    Dim vsoReturnArray(formulaCount - 1) As Object
                    Dim currentFormula As Integer = 0
                    'Loop through each Row one more time. 
                    For rowIndex = 0 To rowCount - 1
                        'Get the Cell Indicies for that row.
                        For cellIndex = 0 To visCellIndiciesArraysList(rowIndex).Count - 1
                            vsoSRCStreamArray(currentFormula) = vsoSectionIndex
                            currentFormula += 1
                            vsoSRCStreamArray(currentFormula) = rowIndex
                            currentFormula += 1
                            vsoSRCStreamArray(currentFormula) = visCellIndiciesArraysList(rowIndex)(cellIndex)
                            currentFormula += 1
                        Next
                    Next

                    'Test if the caller wants values or formulas returned.
                    Select Case getFlag
                        Case _
                            Visio.VisGetSetArgs.visGetFloats, _
                            Visio.VisGetSetArgs.visGetRoundedInts, _
                            Visio.VisGetSetArgs.visGetStrings, _
                            Visio.VisGetSetArgs.visGetTruncatedInts

                            'Results
                            'Execute the StreamGet
                            vsoShape.GetResults(vsoSRCStreamArray, _
                                                getFlag, _
                                                vsoUnitsNamesOrCodes, _
                                                vsoReturnArray)
                            'Add the results Array to the returnList
                            returnList.Add(vsoReturnArray)

                        Case Visio.VisGetSetArgs.visGetFormulas
                            'Formulas
                            'Execute the StreamGet
                            vsoShape.GetFormulas(vsoSRCStreamArray, _
                                                 vsoReturnArray)
                            'Add the Formula Array to the returnList
                            returnList.Add(vsoReturnArray)

                        Case Visio.VisGetSetArgs.visGetFormulasU
                            'FormulasU
                            'Execute the StreamGet
                            vsoShape.GetFormulasU(vsoSRCStreamArray, _
                                                  vsoReturnArray)
                            'Add the Formula Array to the returnList
                            returnList.Add(vsoReturnArray)

                        Case Else
                            'Invalid arguments, Return Nothing
                            Return Nothing

                    End Select
                    'Return the returnList
                    Return returnList
                Else
                    'There are no rows, return nothing. 
                    Return Nothing
                End If
                'End Test whether the section contains any rows.
            Else
                'It does not exist, return nothing.
                Return Nothing
            End If
            'End Test whether the given section even exists.

            'Catch ex As Exception
            'CODE EXCEPTION HANDLER HERE
            'Return nothing in the event of an error.
            'Return Nothing
        Finally
            'Nothing to free here. 
        End Try
    End Function

    Private Function SetSectionUnnamedNonConstantRowsAndCells(ByVal vsoShape As Visio.Shape, _
                                                              ByVal returnList As List(Of Object), _
                                                              ByVal setFlags As Visio.VisGetSetArgs, _
                                                              ByVal setFlagsExtended As visSetFlagsExtended, _
                                                              ByVal vsoUnitsNamesOrCodes() As Object) _
                                                              As Boolean

        Try
            If returnList Is Nothing Then
                Return False
            End If


            'Get the sectionIndex from the list
            Dim vsoSectionIndex As Visio.VisSectionIndices = returnList(0)
            'Get the rowTagsArray from the list.
            Dim visRowTagArray() As Visio.VisRowTags = returnList(1)
            'Get the CellIndiciesArraysList from the list.
            Dim visCellIndiciesArraysList As List(Of Visio.VisCellIndices()) = returnList(2)
            'Get the formulas from the list.
            Dim vsoReturnedArray() As Object = returnList(3)




            'Determine the appropriate add action based on the extended flags. 
            Select Case setFlagsExtended
                Case _
                    visSetFlagsExtended.visSetReplaceAllExisting, _
                    visSetFlagsExtended.visSetReplaceSelectiveAndAdd

                    For i = 0 To visRowTagArray.Count - 1
                        'Loop through the rowTags Array, changing the types of any rows that are incorrect. 
                        If Not visRowTagArray(i) = vsoShape.RowType(vsoSectionIndex, i) Then
                            'The rowTypes are a mismatch. Change it.
                            vsoShape.RowType(vsoSectionIndex, i) = visRowTagArray(i)
                        End If
                    Next
                Case Else
                    'Both SetResults & SetFormulas blast & ignore missing cells by defauls, 
                    'no need to modify to stream or the shapesheet in this case.
            End Select
            'End Determine the appropriate add action based on the extended flags. 


            Dim vsoSRCStreamArray(0) As Short

            'Execute the SRCStream based on the set flags.
            Select Case setFlags
                Case _
                    Visio.VisGetSetArgs.visSetBlastGuards, _
                    Visio.VisGetSetArgs.visSetTestCircular, _
                    Visio.VisGetSetArgs.visSetBlastGuards + Visio.VisGetSetArgs.visSetTestCircular
                    'Set Results
                    vsoShape.SetResults(vsoSRCStreamArray, _
                                        vsoUnitsNamesOrCodes, _
                                        vsoReturnedArray, _
                                        setFlags)
                Case _
                    Visio.VisGetSetArgs.visSetFormulas, _
                    Visio.VisGetSetArgs.visSetFormulas + Visio.VisGetSetArgs.visSetUniversalSyntax, _
                    Visio.VisGetSetArgs.visSetFormulas + Visio.VisGetSetArgs.visSetTestCircular, _
                    Visio.VisGetSetArgs.visSetFormulas + Visio.VisGetSetArgs.visSetBlastGuards, _
                    Visio.VisGetSetArgs.visSetFormulas + Visio.VisGetSetArgs.visSetUniversalSyntax + Visio.VisGetSetArgs.visSetTestCircular, _
                    Visio.VisGetSetArgs.visSetFormulas + Visio.VisGetSetArgs.visSetUniversalSyntax + Visio.VisGetSetArgs.visSetBlastGuards, _
                    Visio.VisGetSetArgs.visSetFormulas + Visio.VisGetSetArgs.visSetUniversalSyntax + Visio.VisGetSetArgs.visSetTestCircular + Visio.VisGetSetArgs.visSetTestCircular
                    'Set Formulas
                    vsoShape.SetFormulas(vsoSRCStreamArray, _
                                         vsoReturnedArray, _
                                         setFlags - Visio.VisGetSetArgs.visSetFormulas)

                Case Else
                    'Invalid SetFlags
                    Return False
            End Select
            'End Execute the SRCStream based on the set flags.
            'Success
            Return True

            'Catch ex As Exception
            'CODE EXCEPTION HANDLER HERE
            'Return False
        Finally
            'Nothing to free here. 
        End Try
    End Function

#End Region

#Region "Get Functions and Overloads"

    ''' <summary>
    ''' Gets the data from the  <see href="http://msdn.microsoft.com/en-us/library/ms408988%28v=office.12%29.aspx">Visio.Section</see>
    ''' (<paramref name="vsoSectionIndex"/>) of the shape (<paramref name="vsoShape"/>) as an <see cref="ArrayList"/>
    ''' of <see cref="ArrayList"/> objects. This function supports any section.
    ''' </summary>
    ''' <param name="vsoShape">The <see href="http://msdn.microsoft.com/en-us/library/ms408994%28v=office.12%29.aspx">Visio.Shape</see> to get data from.</param>
    ''' <param name="getFlag">Indicates whether to get values instead of formulas.</param>
    ''' <param name="vsoSectionIndex">The <see href="http://msdn.microsoft.com/en-us/library/ms408988%28v=office.12%29.aspx">Visio.Section</see> to get data from.</param>
    ''' <returns>
    ''' <c>Nothing</c> if there is no data in the section; otherwise
    ''' an <see cref="ArrayList"/> of <see cref="ArrayList"/> objects which contains the data for each row.
    ''' </returns>
    ''' <remarks>
    ''' If there is only one row in the section, the outer ArrayList.Count = 1. 
    ''' <para>
    ''' Returns the cell values as strings if <paramref name="getVals"/> = <c>True</c>.
    ''' </para>
    ''' <para>
    ''' Returns the cell formulas as strings if <paramref name="getVals"/> = <c>False</c>.
    ''' </para>
    ''' </remarks>
    Public Function GetShapeSection(ByVal vsoShape As Visio.Shape, _
                                    ByVal getFlag As Visio.VisGetSetArgs, _
                                    ByVal vsoSectionIndex As Visio.VisSectionIndices, _
                                    ByVal vsoUnitsNamesOrCodes() As Object) _
                                    As List(Of Object)
        'Drill down to the appropriate method to handle this section. 
        Select Case GetSectionType(vsoSectionIndex)
            Case visSectionTypes.NamedRowsConstantCells
                Return GetSectionNamedRowsConstantCells(vsoShape, getFlag, vsoSectionIndex, vsoUnitsNamesOrCodes)
            Case visSectionTypes.UnnamedConstantRows
                Return GetSectionUnnamedConstantRows(vsoShape, getFlag, vsoSectionIndex, vsoUnitsNamesOrCodes)
            Case visSectionTypes.UnnamedNonConstantRows
                Return GetSectionUnnamedNonConstantRows(vsoShape, getFlag, vsoSectionIndex, vsoUnitsNamesOrCodes)
            Case visSectionTypes.UnnamedNonConstantRowsAndCells
                Return GetSectionUnnamedNonConstantRowsAndCells(vsoShape, getFlag, vsoSectionIndex, vsoUnitsNamesOrCodes)
            Case visSectionTypes.NamedOrUnnamedNonConstantRowsAndCells
                'Return GetSectionNamedOrUnnamedNonConstantRowsAndCells(vsoShape, getFlag, vsoSectionIndex, vsoUnitsNamesOrCodes)
                Return Nothing
            Case visSectionTypes.IsInvalid
                Return Nothing
            Case Else
                Return Nothing
        End Select
    End Function

    ''' <summary>
    ''' Gets the data from multiple  <see href="http://msdn.microsoft.com/en-us/library/ms408988%28v=office.12%29.aspx">Visio.Sections</see>
    ''' (<paramref name="vsoSectionIndicies"/>) of the shape (<paramref name="vsoShape"/>) as a triple nested group of <see cref="ArrayList"/>
    ''' objects. Edge->Outer->Inner = Sections->Rows->Cells
    ''' </summary>
    ''' <param name="vsoShape">The <see href="http://msdn.microsoft.com/en-us/library/ms408994%28v=office.12%29.aspx">Visio.Shape</see> to get data from.</param>
    ''' <param name="getFlag">Indicates whether to get values instead of formulas.</param>
    ''' <param name="vsoSectionIndicies">The <see cref="Array"/> of Section Indicies from which to get data.</param>
    ''' <returns>
    ''' A triple nested group of <see cref="ArrayList"/>
    ''' objects. Edge->Outer->Inner = Sections->Rows->Cells
    ''' </returns>
    ''' <remarks>
    ''' If there is only one row in the section, the outer ArrayList.Count = 1. 
    ''' <para>
    ''' Returns the cell values as strings if <paramref name="getVals"/> = <c>True</c>.
    ''' </para>
    ''' <para>
    ''' Returns the cell formulas as strings if <paramref name="getVals"/> = <c>False</c>.
    ''' </para>
    ''' </remarks>
    Public Function GetShapeSection(ByVal vsoShape As Visio.Shape, _
                                    ByVal getFlag As Visio.VisGetSetArgs, _
                                    ByVal vsoSectionIndicies() As Visio.VisSectionIndices, _
                                    ByVal vsoUnitsNamesOrCodes() As Object) _
                                    As List(Of Object)
        'Drill down to the appropriate method to handle this section. 
        Dim vsoSectionIndex As Short = Nothing
        Dim vsoSectionsDataList As List(Of Object) = Nothing
        Try
            vsoSectionsDataList = New List(Of Object)
            For i = 0 To vsoSectionIndicies.Count - 1
                vsoSectionIndex = vsoSectionIndicies(i)
                vsoSectionsDataList.Add(GetShapeSection(vsoShape, getFlag, vsoSectionIndex, vsoUnitsNamesOrCodes))
            Next i
            Return vsoSectionsDataList
        Finally
            vsoSectionIndex = Nothing
        End Try
    End Function

    Public Function GetShapeSection(ByVal vsoShape As Visio.Shape, _
                                    ByVal getFlag As Visio.VisGetSetArgs, _
                                    ByVal vsoSectionIndex As Visio.VisSectionIndices) _
                                    As List(Of Object)
        Dim vsoUnitsNamesOrCodes() As Object = {Visio.VisUnitCodes.visNoCast}
        Return GetShapeSection(vsoShape, getFlag, vsoSectionIndex, vsoUnitsNamesOrCodes)
    End Function

    Public Function GetShapeSection(ByVal vsoShape As Visio.Shape, _
                                     ByVal getFlag As Visio.VisGetSetArgs, _
                                     ByVal vsoSectionIndicies() As Visio.VisSectionIndices) _
                                     As List(Of Object)
        Dim vsoUnitsNamesOrCodes() As Object = {Visio.VisUnitCodes.visNoCast}
        Return GetShapeSection(vsoShape, getFlag, vsoSectionIndicies, vsoUnitsNamesOrCodes)
    End Function

#End Region

#Region "Set Functions and Overloads"

    ''' <summary>
    ''' Sets the data for the  <see href="http://msdn.microsoft.com/en-us/library/ms408988%28v=office.12%29.aspx">Visio.Section</see>
    ''' (<paramref name="vsoSectionIndex"/>) of the shape (<paramref name="vsoShape"/>) from an <see cref="ArrayList"/>
    ''' of <see cref="ArrayList"/> objects (<paramref name="vsoRowsArray"/>). This function supports any section.
    ''' </summary>
    ''' <param name="vsoShape">The <see href="http://msdn.microsoft.com/en-us/library/ms408994%28v=office.12%29.aspx">Visio.Shape</see> to set the data for.</param>
    ''' <param name="vsoRowsList">An <see cref="ArrayList"/> of <see cref="ArrayList"/> objects, which holds the data for each row (outer) and cell (inner) in the section.</param>
    ''' <param name="vsoSectionIndex">The <see href="http://msdn.microsoft.com/en-us/library/ms408988%28v=office.12%29.aspx">Visio.Section</see> to set data for.</param>
    ''' <returns>
    ''' <c>True</c> if the section's data was set successfully; otherwise <c>False</c>
    ''' </returns>
    ''' <remarks>
    ''' Returns <c>False</c> if the specified section is not valid. 
    ''' </remarks>
    Public Function SetShapeSection(ByVal vsoShape As Visio.Shape, _
                                     ByVal vsoRowsList As List(Of Object), _
                                     ByVal vsoSectionIndex As Visio.VisSectionIndices, _
                                     ByVal setFlags As Visio.VisGetSetArgs, _
                                     ByVal setFlagsExtended As visSetFlagsExtended, _
                                     ByVal vsoUnitsNamesOrCodes() As Object) _
                                     As Boolean


        Dim sectionType As visSectionTypes = GetSectionType(vsoSectionIndex)

        'Determine whether the section is valid or not. 
        Select Case sectionType

            Case _
                visSectionTypes.NamedOrUnnamedNonConstantRowsAndCells, _
                visSectionTypes.NamedRowsConstantCells, _
                visSectionTypes.UnnamedConstantRows, _
                visSectionTypes.UnnamedNonConstantRows, _
                visSectionTypes.UnnamedNonConstantRowsAndCells
                'Valid Indexes


                'Determine the appropriate section add action based on the extended flags. 
                Select Case setFlagsExtended
                    Case visSetFlagsExtended.visSetReplaceAllExisting
                        'Test if the section exists or not. 
                        If Not vsoShape.SectionExists(vsoSectionIndex, Visio.VisExistsFlags.visExistsAnywhere) Then
                            'It does not, add it.
                            vsoShape.AddSection(vsoSectionIndex)
                        Else
                            'It exists, clear it out completely. 
                            'Shapes that are instances of masters automatically have any row contained in the master
                            'recreated in the instance when a section that has been cleared is recreated. 
                            'As such, there is not a way to truly replace all of that kind of shape. 
                            'Test if the shape is a master shape. 
                            If vsoShape.Master Is Nothing Then
                                'The shape is NOT inherited from a master, continue.
                                vsoShape.DeleteSection(vsoSectionIndex)
                                vsoShape.AddSection(vsoSectionIndex)
                                'Delete Section doesn't actually delete everything if the master's section is populated, 
                                'it must be deleted one row at a time. 
                                For i = 0 To vsoShape.RowCount(vsoSectionIndex)
                                    vsoShape.DeleteRow(vsoSectionIndex, Visio.VisRowIndices.visRowLast)
                                Next
                            Else
                                'The shape is inherited from a master.
                                Return False
                            End If
                            'End Test if the shape is a master shape. 
                        End If
                        'End Test if the section exists or not.
                    Case visSetFlagsExtended.visSetReplaceSelectiveAndAdd
                        'Test if the section exists or not. 
                        If Not vsoShape.SectionExists(vsoSectionIndex, Visio.VisExistsFlags.visExistsAnywhere) Then
                            'It does not, add it.
                            vsoShape.AddSection(vsoSectionIndex)
                        End If
                        'End Test if the section exists or not.
                    Case visSetFlagsExtended.visSetReplaceSelectiveAndIgnore
                        'Test if the section exists or not. 
                        If Not vsoShape.SectionExists(vsoSectionIndex, Visio.VisExistsFlags.visExistsAnywhere) Then
                            'It does not, exit.
                            Return False
                        End If
                        'End Test if the section exists or not.
                End Select
                'End Determine the appropriate section add action based on the extended flags. 


                'Drill down to the appropriate method to handle this section. 
                Select Case sectionType
                    Case visSectionTypes.NamedRowsConstantCells
                        Return SetSectionNamedRowsConstantCells(vsoShape, vsoRowsList, setFlags, setFlagsExtended, vsoUnitsNamesOrCodes)
                    Case visSectionTypes.UnnamedConstantRows
                        Return SetSectionUnnamedConstantRows(vsoShape, vsoRowsList, setFlags, setFlagsExtended, vsoUnitsNamesOrCodes)
                    Case visSectionTypes.UnnamedNonConstantRows
                        Return SetSectionUnnamedNonConstantRows(vsoShape, vsoRowsList, setFlags, setFlagsExtended, vsoUnitsNamesOrCodes)
                    Case visSectionTypes.UnnamedNonConstantRowsAndCells
                        'Return SetSectionUnnamedNonConstantRowsAndCells(vsoShape, vsoRowsList, setFlags, setFlagsExtended, vsoUnitsNamesOrCodes)
                        Return True
                    Case visSectionTypes.NamedOrUnnamedNonConstantRowsAndCells
                        Return True
                End Select
                'End Drill down to the appropriate method to handle this section.
            Case Else
                'Includeds visSectionTypes.Invalid
                Return False
        End Select
        'End Determine whether the section is valid or not.

    End Function

    ''' <summary>
    ''' Sets the data for the <see href="http://msdn.microsoft.com/en-us/library/ms408988%28v=office.12%29.aspx">Visio.Section</see>
    ''' (<paramref name="vsoSectionIndex"/>) of the shape (<paramref name="vsoShape"/>)  multiple  
    ''' <see href="http://msdn.microsoft.com/en-us/library/ms408988%28v=office.12%29.aspx">Visio.Sections</see>
    ''' (<paramref name="vsoSectionIndicies"/>) of the shape (<paramref name="vsoShape"/>) as a triple nested group of <see cref="ArrayList"/>
    ''' objects. Edge->Outer->Inner = Sections->Rows->Cells (<paramref name="vsoSectionArray"/>). This function supports any section.
    ''' </summary>
    ''' <param name="vsoShape">The <see href="http://msdn.microsoft.com/en-us/library/ms408994%28v=office.12%29.aspx">Visio.Shape</see> to set the data for.</param>
    ''' <param name="vsoSectionIndicies">The <see cref="Array"/> of Section Indicies from which to get data.</param>
    ''' <param name="vsoSectionRowsList">A triple nested group of <see cref="ArrayList"/> objects. Edge->Outer->Inner = Sections->Rows->Cells</param>
    ''' <returns>
    ''' An array of <see cref="Boolean"/> values an element corresponding to each index specified in (<paramref name="vsoSectionIndicies"/>).
    ''' <c>True</c> if the section's data was set successfully; otherwise <c>False</c>
    ''' </returns>
    ''' <remarks>
    ''' Returns <c>False</c> if the specified section is not valid. 
    ''' </remarks>
    Public Function SetShapeSection(ByVal vsoShape As Visio.Shape, _
                                    ByVal vsoSectionRowsList As List(Of Object), _
                                    ByVal vsoSectionIndicies() As Visio.VisSectionIndices, _
                                    ByVal setFlags() As Visio.VisGetSetArgs, _
                                    ByVal setFlagsExtended() As visSetFlagsExtended, _
                                    ByVal vsoUnitsNamesOrCodes() As Object) _
                                    As Boolean()

        Dim successArray(vsoSectionIndicies.Count) As Boolean
        Dim currFlag As Integer = 0
        Dim currAdd As Boolean = False
        For i = 0 To setFlags.Count - 1
            successArray(i) = SetShapeSection(vsoShape, _
                                              vsoSectionRowsList(i), _
                                              vsoSectionIndicies(i), _
                                              setFlags(i), _
                                              setFlagsExtended(i), _
                                              vsoUnitsNamesOrCodes)
            currFlag = i
        Next
        For i = currFlag To vsoSectionIndicies.Count - 1
            successArray(i) = SetShapeSection(vsoShape, _
                                              vsoSectionRowsList(i), _
                                              vsoSectionIndicies(i), _
                                              setFlags(currFlag), _
                                              setFlagsExtended(i), _
                                              vsoUnitsNamesOrCodes)
        Next
        Return successArray
    End Function

    Public Function SetShapeSection(ByVal vsoShape As Visio.Shape, _
                                    ByVal vsoSectionRowsList As List(Of Object), _
                                    ByVal vsoSectionIndicies() As Visio.VisSectionIndices, _
                                    ByVal setFlags() As Visio.VisGetSetArgs, _
                                    ByVal setFlagsExtended() As visSetFlagsExtended) _
                                    As Boolean()

        Dim vsoUnitsNamesOrCodes() As Object = {Visio.VisUnitCodes.visNoCast}
        Return SetShapeSection(vsoShape, vsoSectionRowsList, vsoSectionIndicies, setFlags, setFlagsExtended, vsoUnitsNamesOrCodes)

    End Function

    Public Function SetShapeSection(ByVal vsoShape As Visio.Shape, _
                                     ByVal vsoRowsList As List(Of Object), _
                                     ByVal vsoSectionIndex As Visio.VisSectionIndices, _
                                     ByVal setFlags As Visio.VisGetSetArgs, _
                                     ByVal setFlagsExtended As visSetFlagsExtended) _
                                     As Boolean

        Dim vsoUnitsNamesOrCodes() As Object = {Visio.VisUnitCodes.visNoCast}
        Return SetShapeSection(vsoShape, vsoRowsList, vsoSectionIndex, setFlags, setFlagsExtended, vsoUnitsNamesOrCodes)

    End Function
#End Region

#Region "Section Types and Indicies"

    ''' <summary>
    ''' Gets the <see cref="visSectionTypes">Section Type</see> of the specified <see href="http://msdn.microsoft.com/en-us/library/ms408988%28v=office.12%29.aspx">Visio.Section</see> (<paramref name="vsoSectionIndicies"/>).
    ''' </summary>
    ''' <param name="vsoSectionIndex">Index of the <see href="http://msdn.microsoft.com/en-us/library/ms408988%28v=office.12%29.aspx">Visio.Section</see> to get the type for.</param>
    ''' <returns>
    ''' The <see cref="visSectionTypes">Section Type</see> of the specified <see href="http://msdn.microsoft.com/en-us/library/ms408988%28v=office.12%29.aspx">Visio.Section</see> (<paramref name="vsoSectionIndicies"/>).
    ''' </returns>
    Public Function GetSectionType(ByVal vsoSectionIndex As Visio.VisSectionIndices) As visSectionTypes
        Select Case vsoSectionIndex
            Case _
                Visio.VisSectionIndices.visSectionAction, _
                Visio.VisSectionIndices.visSectionHyperlink, _
                Visio.VisSectionIndices.visSectionProp, _
                Visio.VisSectionIndices.visSectionSmartTag, _
                Visio.VisSectionIndices.visSectionUser
                Return visSectionTypes.NamedRowsConstantCells
            Case _
                Visio.VisSectionIndices.visSectionObject
                Return visSectionTypes.UnnamedConstantRows
            Case _
                Visio.VisSectionIndices.visSectionAnnotation, _
                Visio.VisSectionIndices.visSectionCharacter, _
                Visio.VisSectionIndices.visSectionLayer, _
                Visio.VisSectionIndices.visSectionParagraph, _
                Visio.VisSectionIndices.visSectionReviewer, _
                Visio.VisSectionIndices.visSectionScratch, _
                Visio.VisSectionIndices.visSectionTab, _
                Visio.VisSectionIndices.visSectionTextField
                Return visSectionTypes.UnnamedNonConstantRows
            Case _
                Visio.VisSectionIndices.visSectionTab
                Return visSectionTypes.UnnamedNonConstantRowsAndCells
            Case _
                Visio.VisSectionIndices.visSectionControls
                Return visSectionTypes.NamedNonConstantRowsAndCells
            Case _
                Visio.VisSectionIndices.visSectionConnectionPts
                Return visSectionTypes.NamedOrUnnamedNonConstantRowsAndCells
            Case _
                Visio.VisSectionIndices.visSectionInval, _
                Visio.VisSectionIndices.visSectionNone
                Return visSectionTypes.IsInvalid
            Case Else
                'Any section indexed between visSectionFirstComponent and visSectionLastComponent belongs to the Geometry Section.
                If (Visio.VisSectionIndices.visSectionFirstComponent <= vsoSectionIndex) AndAlso _
                    (vsoSectionIndex <= Visio.VisSectionIndices.visSectionLastComponent) Then
                    'Belongs to Geometry Section which can have a variety of Rows. 
                    Return visSectionTypes.UnnamedNonConstantRowsAndCells
                Else
                    'The section is not a valid geometry section, it is not a valid section.
                    Return visSectionTypes.IsInvalid
                End If
        End Select
    End Function

    ''' <summary>
    ''' Gets the known row indicies for a given section. (<paramref name="vsoSectionIndex"/>.
    ''' </summary>
    ''' <param name="vsoSectionIndex">Index of the <see href="http://msdn.microsoft.com/en-us/library/ms408988%28v=office.12%29.aspx">Visio.Section</see> to get the subsections for.</param>
    ''' <returns>An <see cref="ArrayList"> containing the known row indicies for a given section.</see></returns>
    Public Function GetRowIndicies(ByVal vsoSectionIndex As Visio.VisSectionIndices) As Visio.VisRowIndices()
        Dim vsoRowIndicies() As Visio.VisRowIndices
        'Initialize the array as empty. 
        'Else the array should be empty.
        vsoRowIndicies = New Visio.VisRowIndices() {}

        'Add all of the known rows for a given section. 
        Select Case vsoSectionIndex
            Case Visio.VisSectionIndices.visSectionAction
                vsoRowIndicies = New Visio.VisRowIndices() { _
                    Visio.VisRowIndices.visRowAction}
            Case Visio.VisSectionIndices.visSectionAnnotation
                vsoRowIndicies = New Visio.VisRowIndices() { _
                    Visio.VisRowIndices.visRowAnnotation}
            Case Visio.VisSectionIndices.visSectionCharacter
                vsoRowIndicies = New Visio.VisRowIndices() { _
                    Visio.VisRowIndices.visRowCharacter}
            Case Visio.VisSectionIndices.visSectionConnectionPts
                vsoRowIndicies = New Visio.VisRowIndices() { _
                    Visio.VisRowIndices.visRowConnectionPts}
            Case Visio.VisSectionIndices.visSectionControls
                vsoRowIndicies = New Visio.VisRowIndices() { _
                    Visio.VisRowIndices.visRowControl}
            Case Visio.VisSectionIndices.visSectionHyperlink
                vsoRowIndicies = New Visio.VisRowIndices() { _
                    Visio.VisRowIndices.visRowHyperlink}
            Case Visio.VisSectionIndices.visSectionLayer
                vsoRowIndicies = New Visio.VisRowIndices() { _
                    Visio.VisRowIndices.visRowLayer}
            Case Visio.VisSectionIndices.visSectionObject
                vsoRowIndicies = New Visio.VisRowIndices() { _
                    Visio.VisRowIndices.visRowAlign, _
                    Visio.VisRowIndices.visRowDoc, _
                    Visio.VisRowIndices.visRowEvent, _
                    Visio.VisRowIndices.visRowFill, _
                    Visio.VisRowIndices.visRowForeign, _
                    Visio.VisRowIndices.visRowGroup, _
                    Visio.VisRowIndices.visRowHelpCopyright, _
                    Visio.VisRowIndices.visRowImage, _
                    Visio.VisRowIndices.visRowLayerMem, _
                    Visio.VisRowIndices.visRowLine, _
                    Visio.VisRowIndices.visRowLock, _
                    Visio.VisRowIndices.visRowMisc, _
                    Visio.VisRowIndices.visRowPageLayout, _
                    Visio.VisRowIndices.visRowPage, _
                    Visio.VisRowIndices.visRowPrintProperties, _
                    Visio.VisRowIndices.visRowRulerGrid, _
                    Visio.VisRowIndices.visRowShapeLayout, _
                    Visio.VisRowIndices.visRowStyle, _
                    Visio.VisRowIndices.visRowTextXForm, _
                    Visio.VisRowIndices.visRowText, _
                    Visio.VisRowIndices.visRowXForm1D, _
                    Visio.VisRowIndices.visRowXFormOut}
            Case Visio.VisSectionIndices.visSectionParagraph
                vsoRowIndicies = New Visio.VisRowIndices() { _
                    Visio.VisRowIndices.visRowParagraph}
            Case Visio.VisSectionIndices.visSectionProp
                vsoRowIndicies = New Visio.VisRowIndices() { _
                    Visio.VisRowIndices.visRowProp}
            Case Visio.VisSectionIndices.visSectionReviewer
                vsoRowIndicies = New Visio.VisRowIndices() { _
                    Visio.VisRowIndices.visRowReviewer}
            Case Visio.VisSectionIndices.visSectionScratch
                vsoRowIndicies = New Visio.VisRowIndices() { _
                    Visio.VisRowIndices.visRowScratch}
            Case Visio.VisSectionIndices.visSectionTab
                vsoRowIndicies = New Visio.VisRowIndices() { _
                    Visio.VisRowIndices.visRowTab}
            Case Visio.VisSectionIndices.visSectionTextField
                vsoRowIndicies = New Visio.VisRowIndices() { _
                    Visio.VisRowIndices.visRowField}
            Case Visio.VisSectionIndices.visSectionSmartTag
                vsoRowIndicies = New Visio.VisRowIndices() { _
                    Visio.VisRowIndices.visRowSmartTag}
            Case Visio.VisSectionIndices.visSectionUser
                vsoRowIndicies = New Visio.VisRowIndices() { _
                    Visio.VisRowIndices.visRowUser}
            Case Else
                If (Visio.VisSectionIndices.visSectionFirstComponent <= vsoSectionIndex) AndAlso _
                    (vsoSectionIndex <= Visio.VisSectionIndices.visSectionLastComponent) Then
                    'Belongs to Geometry Section which can have a variety of Rows.
                    vsoRowIndicies = New Visio.VisRowIndices() { _
                        Visio.VisRowIndices.visRowComponent, _
                        Visio.VisRowIndices.visRowVertex}
                End If
        End Select
        Return vsoRowIndicies
    End Function

    ''' <summary>
    ''' Gets the known cell indicies for a given Section/Row/Row Tag combination. 
    ''' </summary>
    ''' <param name="vsoSectionIndex">Index of the <see href="http://msdn.microsoft.com/en-us/library/ms408988%28v=office.12%29.aspx">Visio.Section</see> containing the row.</param>
    ''' <param name="vsoRowIndex">Index of the <see href="http://msdn.microsoft.com/en-us/library/ms408986%28v=office.12%29.aspx">Visio.Row</see> containing the cells.</param>
    ''' <param name="vsoRowTag">Constant specifying an identifying tag used to specify row type.</param>
    ''' <returns>
    ''' An <see cref="Array"> containing the cell indicies for the row.</see>
    ''' </returns>
    ''' <remarks>
    ''' The <paramref name="vsoRowIndex"/> parameter is only used when necessary to determine which cells will be present in the cells. 
    ''' (Sections: Controls, Connection Points, Tab, and Geometry) It is ignored for all other sections.
    ''' Returns an empty array if there are no cells or the caller hasn't specified a valid section/row combination.
    ''' </remarks>
    Public Function GetCellIndicies(ByVal vsoSectionIndex As Visio.VisSectionIndices, _
                                    ByVal vsoRowIndex As Visio.VisRowIndices, _
                                    ByVal vsoRowTag As Visio.VisRowTags) _
                                    As Visio.VisCellIndices()
        Dim vsoCellsArray() As Visio.VisCellIndices
        'Initialize the array as empty. 
        'Else the array should be empty.
        vsoCellsArray = New Visio.VisCellIndices() {}


        'Add all of the known cells for the given row.
        Select Case vsoSectionIndex
            Case Visio.VisSectionIndices.visSectionAction
                Select Case vsoRowIndex
                    Case Visio.VisRowIndices.visRowAction
                        'First row in the Action Section.
                        vsoCellsArray = New Visio.VisCellIndices() { _
                            Visio.VisCellIndices.visActionAction, _
                            Visio.VisCellIndices.visActionBeginGroup, _
                            Visio.VisCellIndices.visActionButtonFace, _
                            Visio.VisCellIndices.visActionChecked, _
                            Visio.VisCellIndices.visActionDisabled, _
                            Visio.VisCellIndices.visActionInvisible, _
                            Visio.VisCellIndices.visActionMenu, _
                            Visio.VisCellIndices.visActionReadOnly, _
                            Visio.VisCellIndices.visActionSortKey, _
                            Visio.VisCellIndices.visActionTagName}
                        'Omitted because it isn't documented in the SDK.
                        'Visio.VisCellIndices.visActionHelp, _
                        'Visio.VisCellIndices.visActionPrompt, _
                End Select
            Case Visio.VisSectionIndices.visSectionAnnotation
                Select Case vsoRowIndex
                    Case Visio.VisRowIndices.visRowAnnotation
                        'First row of the Annotation Section
                        vsoCellsArray = New Visio.VisCellIndices() { _
                            Visio.VisCellIndices.visAnnotationComment, _
                            Visio.VisCellIndices.visAnnotationDate, _
                            Visio.VisCellIndices.visAnnotationLangID, _
                            Visio.VisCellIndices.visAnnotationMarkerIndex, _
                            Visio.VisCellIndices.visAnnotationReviewerID, _
                            Visio.VisCellIndices.visAnnotationX, _
                            Visio.VisCellIndices.visAnnotationY}
                End Select
            Case Visio.VisSectionIndices.visSectionCharacter
                Select Case vsoRowIndex
                    Case Visio.VisRowIndices.visRowCharacter
                        'First row in the Character Section.
                        vsoCellsArray = New Visio.VisCellIndices() { _
                            Visio.VisCellIndices.visCharacterAsianFont, _
                            Visio.VisCellIndices.visCharacterCase, _
                            Visio.VisCellIndices.visCharacterColor, _
                            Visio.VisCellIndices.visCharacterColorTrans, _
                            Visio.VisCellIndices.visCharacterComplexScriptFont, _
                            Visio.VisCellIndices.visCharacterComplexScriptSize, _
                            Visio.VisCellIndices.visCharacterDblUnderline, _
                            Visio.VisCellIndices.visCharacterDoubleStrikethrough, _
                            Visio.VisCellIndices.visCharacterFont, _
                            Visio.VisCellIndices.visCharacterFontScale, _
                            Visio.VisCellIndices.visCharacterLangID, _
                            Visio.VisCellIndices.visCharacterLetterspace, _
                            Visio.VisCellIndices.visCharacterLocale, _
                            Visio.VisCellIndices.visCharacterLocalizeFont, _
                            Visio.VisCellIndices.visCharacterOverline, _
                            Visio.VisCellIndices.visCharacterPos, _
                            Visio.VisCellIndices.visCharacterSize, _
                            Visio.VisCellIndices.visCharacterStrikethru, _
                            Visio.VisCellIndices.visCharacterStyle}
                        'Omitted since cell is not implemented in Visio 2007
                        'Visio.VisCellIndices.visCharacterRTLText, _
                        'Visio.VisCellIndices.visCharacterUseVertical, _
                        'Omitted because it isn't documented in the SDK.
                        'Visio.VisCellIndices.visCharacterPerpendicular, _
                End Select
            Case Visio.VisSectionIndices.visSectionConnectionPts
                Select Case vsoRowIndex
                    Case Visio.VisRowIndices.visRowConnectionPts
                        'First row of the Connection Points Section
                        Select Case vsoRowTag
                            Case Visio.VisRowTags.visTagCnnctPt
                                'The row type of a row in a visSectionConnectionPts section that has unnamed rows.
                                vsoCellsArray = New Visio.VisCellIndices() { _
                                    Visio.VisCellIndices.visCnnctX, _
                                    Visio.VisCellIndices.visCnnctY, _
                                    Visio.VisCellIndices.visCnnctDirX, _
                                    Visio.VisCellIndices.visCnnctDirY, _
                                    Visio.VisCellIndices.visCnnctType}
                            Case Visio.VisRowTags.visTagCnnctPtABCD
                                'The row type of an extended row in a visSectionConnectionPts section that has unnamed rows. Seldom used.
                                vsoCellsArray = New Visio.VisCellIndices() { _
                                    Visio.VisCellIndices.visCnnctX, _
                                    Visio.VisCellIndices.visCnnctY, _
                                    Visio.VisCellIndices.visCnnctA, _
                                    Visio.VisCellIndices.visCnnctB, _
                                    Visio.VisCellIndices.visCnnctC, _
                                    Visio.VisCellIndices.visCnnctD}
                            Case Visio.VisRowTags.visTagCnnctNamed
                                'The row type of a row in a visSectionConnectionPts section that has named rows.
                                vsoCellsArray = New Visio.VisCellIndices() { _
                                    Visio.VisCellIndices.visCnnctX, _
                                    Visio.VisCellIndices.visCnnctY, _
                                    Visio.VisCellIndices.visCnnctDirX, _
                                    Visio.VisCellIndices.visCnnctDirY, _
                                    Visio.VisCellIndices.visCnnctType}
                            Case Visio.VisRowTags.visTagCnnctNamedABCD
                                'The row type of an extended row in a visSectionConnectionPts section that has named rows. Seldom used.
                                vsoCellsArray = New Visio.VisCellIndices() { _
                                    Visio.VisCellIndices.visCnnctX, _
                                    Visio.VisCellIndices.visCnnctY, _
                                    Visio.VisCellIndices.visCnnctA, _
                                    Visio.VisCellIndices.visCnnctB, _
                                    Visio.VisCellIndices.visCnnctC, _
                                    Visio.VisCellIndices.visCnnctD}
                                'Omitted because it is the same index as visCnnctD
                                'Visio.VisCellIndices.visCnnctAutoGen, _
                        End Select
                End Select
            Case Visio.VisSectionIndices.visSectionControls
                Select Case vsoRowIndex
                    Case Visio.VisRowIndices.visRowControl
                        'First row of the Controls Section
                        Select Case vsoRowTag
                            Case Visio.VisRowTags.visTagCtlPt
                                'Control Point WITHOUT the Tooltip cell.
                                vsoCellsArray = New Visio.VisCellIndices() { _
                                    Visio.VisCellIndices.visCtlGlue, _
                                    Visio.VisCellIndices.visCtlX, _
                                    Visio.VisCellIndices.visCtlXCon, _
                                    Visio.VisCellIndices.visCtlXDyn, _
                                    Visio.VisCellIndices.visCtlY, _
                                    Visio.VisCellIndices.visCtlYCon, _
                                    Visio.VisCellIndices.visCtlYDyn}
                            Case Visio.VisRowTags.visTagCtlPtTip
                                'Control Point WITH the Tooltip cell.
                                vsoCellsArray = New Visio.VisCellIndices() { _
                                    Visio.VisCellIndices.visCtlGlue, _
                                    Visio.VisCellIndices.visCtlTip, _
                                    Visio.VisCellIndices.visCtlX, _
                                    Visio.VisCellIndices.visCtlXCon, _
                                    Visio.VisCellIndices.visCtlXDyn, _
                                    Visio.VisCellIndices.visCtlY, _
                                    Visio.VisCellIndices.visCtlYCon, _
                                    Visio.VisCellIndices.visCtlYDyn}
                                'Omitted because it isn't documented in the SDK.
                                'Visio.VisCellIndices.visCtlType, _
                        End Select
                End Select
            Case Visio.VisSectionIndices.visSectionHyperlink
                Select Case vsoRowIndex
                    Case Visio.VisRowIndices.visRowHyperlink
                        'First row of the Hyperlinks Section
                        vsoCellsArray = New Visio.VisCellIndices() { _
                            Visio.VisCellIndices.visHLinkAddress, _
                            Visio.VisCellIndices.visHLinkDefault, _
                            Visio.VisCellIndices.visHLinkDescription, _
                            Visio.VisCellIndices.visHLinkExtraInfo, _
                            Visio.VisCellIndices.visHLinkFrame, _
                            Visio.VisCellIndices.visHLinkInvisible, _
                            Visio.VisCellIndices.visHLinkNewWin, _
                            Visio.VisCellIndices.visHLinkSortKey, _
                            Visio.VisCellIndices.visHLinkSubAddress}
                End Select
            Case Visio.VisSectionIndices.visSectionLayer
                Select Case vsoRowIndex
                    Case Visio.VisRowIndices.visRowLayer
                        'First row of the Layers Section
                        vsoCellsArray = New Visio.VisCellIndices() { _
                            Visio.VisCellIndices.visLayerActive, _
                            Visio.VisCellIndices.visLayerColorTrans, _
                            Visio.VisCellIndices.visLayerColor, _
                            Visio.VisCellIndices.visLayerGlue, _
                            Visio.VisCellIndices.visLayerLock, _
                            Visio.VisCellIndices.visLayerNameUniv, _
                            Visio.VisCellIndices.visLayerName, _
                            Visio.VisCellIndices.visLayerPrint, _
                            Visio.VisCellIndices.visLayerSnap, _
                            Visio.VisCellIndices.visLayerStatus, _
                            Visio.VisCellIndices.visLayerVisible}
                End Select
            Case Visio.VisSectionIndices.visSectionObject
                Select Case vsoRowIndex
                    Case Visio.VisRowIndices.visRowXForm1D
                        '1-D Endpoints SubSection
                        vsoCellsArray = New Visio.VisCellIndices() { _
                            Visio.VisCellIndices.vis1DBeginX, _
                            Visio.VisCellIndices.vis1DBeginY, _
                            Visio.VisCellIndices.vis1DEndX, _
                            Visio.VisCellIndices.vis1DEndY}
                    Case Visio.VisRowIndices.visRowAlign
                        'Alignment SubSection
                        vsoCellsArray = New Visio.VisCellIndices() { _
                            Visio.VisCellIndices.visAlignBottom, _
                            Visio.VisCellIndices.visAlignCenter, _
                            Visio.VisCellIndices.visAlignLeft, _
                            Visio.VisCellIndices.visAlignMiddle, _
                            Visio.VisCellIndices.visAlignRight, _
                            Visio.VisCellIndices.visAlignTop}
                    Case Visio.VisRowIndices.visRowDoc
                        'Document Properties Subsection
                        vsoCellsArray = New Visio.VisCellIndices() { _
                            Visio.VisCellIndices.visDocAddMarkup, _
                            Visio.VisCellIndices.visDocLangID, _
                            Visio.VisCellIndices.visDocLockPreview, _
                            Visio.VisCellIndices.visDocOutputFormat, _
                            Visio.VisCellIndices.visDocPreviewQuality, _
                            Visio.VisCellIndices.visDocPreviewScope, _
                            Visio.VisCellIndices.visDocViewMarkup}
                    Case Visio.VisRowIndices.visRowEvent
                        'Events SubSection
                        vsoCellsArray = New Visio.VisCellIndices() { _
                            Visio.VisCellIndices.visEvtCellDblClick, _
                            Visio.VisCellIndices.visEvtCellDrop, _
                            Visio.VisCellIndices.visEvtCellMultiDrop, _
                            Visio.VisCellIndices.visEvtCellTheText, _
                            Visio.VisCellIndices.visEvtCellXFMod}
                    Case Visio.VisRowIndices.visRowFill
                        'Fill SubSection
                        vsoCellsArray = New Visio.VisCellIndices() { _
                            Visio.VisCellIndices.visFillBkgnd, _
                            Visio.VisCellIndices.visFillBkgndTrans, _
                            Visio.VisCellIndices.visFillForegnd, _
                            Visio.VisCellIndices.visFillForegndTrans, _
                            Visio.VisCellIndices.visFillPattern, _
                            Visio.VisCellIndices.visFillShdwBkgnd, _
                            Visio.VisCellIndices.visFillShdwBkgndTrans, _
                            Visio.VisCellIndices.visFillShdwForegnd, _
                            Visio.VisCellIndices.visFillShdwForegndTrans, _
                            Visio.VisCellIndices.visFillShdwObliqueAngle, _
                            Visio.VisCellIndices.visFillShdwOffsetX, _
                            Visio.VisCellIndices.visFillShdwOffsetY, _
                            Visio.VisCellIndices.visFillShdwPattern, _
                            Visio.VisCellIndices.visFillShdwScaleFactor}
                    Case Visio.VisRowIndices.visRowForeign
                        'Foreign Image Info SubSection
                        vsoCellsArray = New Visio.VisCellIndices() { _
                            Visio.VisCellIndices.visFrgnImgHeight, _
                            Visio.VisCellIndices.visFrgnImgOffsetX, _
                            Visio.VisCellIndices.visFrgnImgOffsetY, _
                            Visio.VisCellIndices.visFrgnImgWidth}
                    Case Visio.VisRowIndices.visRowGroup
                        'Group Properties SubSection
                        vsoCellsArray = New Visio.VisCellIndices() { _
                            Visio.VisCellIndices.visGroupDisplayMode, _
                            Visio.VisCellIndices.visGroupDontMoveChildren, _
                            Visio.VisCellIndices.visGroupIsDropTarget, _
                            Visio.VisCellIndices.visGroupIsSnapTarget, _
                            Visio.VisCellIndices.visGroupIsTextEditTarget, _
                            Visio.VisCellIndices.visGroupSelectMode}
                    Case Visio.VisRowIndices.visRowHelpCopyright
                        'Help and Copyright SubSection
                        vsoCellsArray = New Visio.VisCellIndices() { _
                            Visio.VisCellIndices.visObjHelp, _
                            Visio.VisCellIndices.visCopyright}
                    Case Visio.VisRowIndices.visRowImage
                        'Image Properties SubSection
                        vsoCellsArray = New Visio.VisCellIndices() { _
                            Visio.VisCellIndices.visImageBlur, _
                            Visio.VisCellIndices.visImageBrightness, _
                            Visio.VisCellIndices.visImageContrast, _
                            Visio.VisCellIndices.visImageDenoise, _
                            Visio.VisCellIndices.visImageGamma, _
                            Visio.VisCellIndices.visImageSharpen, _
                            Visio.VisCellIndices.visImageTransparency}
                    Case Visio.VisRowIndices.visRowLayerMem
                        'Layer Membership SubSection
                        vsoCellsArray = New Visio.VisCellIndices() { _
                            Visio.VisCellIndices.visLayerMember}
                    Case Visio.VisRowIndices.visRowLine
                        'Line Format SubSection
                        vsoCellsArray = New Visio.VisCellIndices() { _
                            Visio.VisCellIndices.visLineBeginArrow, _
                            Visio.VisCellIndices.visLineBeginArrowSize, _
                            Visio.VisCellIndices.visLineColor, _
                            Visio.VisCellIndices.visLineColorTrans, _
                            Visio.VisCellIndices.visLineEndArrow, _
                            Visio.VisCellIndices.visLineEndArrowSize, _
                            Visio.VisCellIndices.visLineEndCap, _
                            Visio.VisCellIndices.visLinePattern, _
                            Visio.VisCellIndices.visLineRounding, _
                            Visio.VisCellIndices.visLineWeight}
                        'Omitted because it isn't documented in the SDK.
                        'Visio.VisCellIndices.visLineArrowSize, _
                    Case Visio.VisRowIndices.visRowLock
                        'Protection SubSection
                        vsoCellsArray = New Visio.VisCellIndices() { _
                            Visio.VisCellIndices.visLockAspect, _
                            Visio.VisCellIndices.visLockBegin, _
                            Visio.VisCellIndices.visLockCalcWH, _
                            Visio.VisCellIndices.visLockCrop, _
                            Visio.VisCellIndices.visLockCustProp, _
                            Visio.VisCellIndices.visLockDelete, _
                            Visio.VisCellIndices.visLockEnd, _
                            Visio.VisCellIndices.visLockFormat, _
                            Visio.VisCellIndices.visLockFromGroupFormat, _
                            Visio.VisCellIndices.visLockGroup, _
                            Visio.VisCellIndices.visLockHeight, _
                            Visio.VisCellIndices.visLockMoveX, _
                            Visio.VisCellIndices.visLockMoveY, _
                            Visio.VisCellIndices.visLockRotate, _
                            Visio.VisCellIndices.visLockSelect, _
                            Visio.VisCellIndices.visLockTextEdit, _
                            Visio.VisCellIndices.visLockThemeColors, _
                            Visio.VisCellIndices.visLockThemeEffects, _
                            Visio.VisCellIndices.visLockVtxEdit, _
                            Visio.VisCellIndices.visLockWidth}
                    Case Visio.VisRowIndices.visRowMisc
                        'Miscellaneous & Glue Info SubSection
                        vsoCellsArray = New Visio.VisCellIndices() { _
                            Visio.VisCellIndices.visBegTrigger, _
                            Visio.VisCellIndices.visComment, _
                            Visio.VisCellIndices.visCopyright, _
                            Visio.VisCellIndices.visDropSource, _
                            Visio.VisCellIndices.visDynFeedback, _
                            Visio.VisCellIndices.visEndTrigger, _
                            Visio.VisCellIndices.visGlueType, _
                            Visio.VisCellIndices.visHideText, _
                            Visio.VisCellIndices.visLOFlags, _
                            Visio.VisCellIndices.visNoAlignBox, _
                            Visio.VisCellIndices.visNoCtlHandles, _
                            Visio.VisCellIndices.visNoLiveDynamics, _
                            Visio.VisCellIndices.visNonPrinting, _
                            Visio.VisCellIndices.visNoObjHandles, _
                            Visio.VisCellIndices.visObjCalendar, _
                            Visio.VisCellIndices.visObjDropOnPageScale, _
                            Visio.VisCellIndices.visObjHelp, _
                            Visio.VisCellIndices.visObjKeywords, _
                            Visio.VisCellIndices.visObjLangID, _
                            Visio.VisCellIndices.visObjLocalizeMerge, _
                            Visio.VisCellIndices.visObjTheme, _
                            Visio.VisCellIndices.visUpdateAlignBox, _
                            Visio.VisCellIndices.visWalkPref}
                    Case Visio.VisRowIndices.visRowPageLayout
                        'Page Layout SubSection
                        vsoCellsArray = New Visio.VisCellIndices() { _
                            Visio.VisCellIndices.visPLOAvenueSizeX, _
                            Visio.VisCellIndices.visPLOAvenueSizeY, _
                            Visio.VisCellIndices.visPLOBlockSizeX, _
                            Visio.VisCellIndices.visPLOBlockSizeY, _
                            Visio.VisCellIndices.visPLOCtrlAsInput, _
                            Visio.VisCellIndices.visPLODynamicsOff, _
                            Visio.VisCellIndices.visPLOEnableGrid, _
                            Visio.VisCellIndices.visPLOJumpCode, _
                            Visio.VisCellIndices.visPLOJumpDirX, _
                            Visio.VisCellIndices.visPLOJumpDirY, _
                            Visio.VisCellIndices.visPLOJumpFactorX, _
                            Visio.VisCellIndices.visPLOJumpFactorY, _
                            Visio.VisCellIndices.visPLOJumpStyle, _
                            Visio.VisCellIndices.visPLOLineAdjustFrom, _
                            Visio.VisCellIndices.visPLOLineAdjustTo, _
                            Visio.VisCellIndices.visPLOLineRouteExt, _
                            Visio.VisCellIndices.visPLOLineToLineX, _
                            Visio.VisCellIndices.visPLOLineToLineY, _
                            Visio.VisCellIndices.visPLOLineToNodeX, _
                            Visio.VisCellIndices.visPLOLineToNodeY, _
                            Visio.VisCellIndices.visPLOPlaceDepth, _
                            Visio.VisCellIndices.visPLOPlaceFlip, _
                            Visio.VisCellIndices.visPLOPlaceStyle, _
                            Visio.VisCellIndices.visPLOPlowCode, _
                            Visio.VisCellIndices.visPLOResizePage, _
                            Visio.VisCellIndices.visPLORouteStyle, _
                            Visio.VisCellIndices.visPLOSplit}
                    Case Visio.VisRowIndices.visRowPage
                        'Page SubSection
                        vsoCellsArray = New Visio.VisCellIndices() { _
                            Visio.VisCellIndices.visFillShdwType, _
                            Visio.VisCellIndices.visPageDrawingScale, _
                            Visio.VisCellIndices.visPageDrawScaleType, _
                            Visio.VisCellIndices.visPageDrawSizeType, _
                            Visio.VisCellIndices.visPageHeight, _
                            Visio.VisCellIndices.visPageInhibitSnap, _
                            Visio.VisCellIndices.visPageScale, _
                            Visio.VisCellIndices.visPageShdwObliqueAngle, _
                            Visio.VisCellIndices.visPageShdwOffsetX, _
                            Visio.VisCellIndices.visPageShdwOffsetY, _
                            Visio.VisCellIndices.visPageShdwScaleFactor, _
                            Visio.VisCellIndices.visPageShdwType, _
                            Visio.VisCellIndices.visPageUIVisibility, _
                            Visio.VisCellIndices.visPageWidth}
                    Case Visio.VisRowIndices.visRowPrintProperties
                        'Print Properties SubSection
                        vsoCellsArray = New Visio.VisCellIndices() { _
                            Visio.VisCellIndices.visPrintPropertiesBottomMargin, _
                            Visio.VisCellIndices.visPrintPropertiesCenterX, _
                            Visio.VisCellIndices.visPrintPropertiesCenterY, _
                            Visio.VisCellIndices.visPrintPropertiesLeftMargin, _
                            Visio.VisCellIndices.visPrintPropertiesOnPage, _
                            Visio.VisCellIndices.visPrintPropertiesPageOrientation, _
                            Visio.VisCellIndices.visPrintPropertiesPagesX, _
                            Visio.VisCellIndices.visPrintPropertiesPagesY, _
                            Visio.VisCellIndices.visPrintPropertiesPaperKind, _
                            Visio.VisCellIndices.visPrintPropertiesPaperSource, _
                            Visio.VisCellIndices.visPrintPropertiesPrintGrid, _
                            Visio.VisCellIndices.visPrintPropertiesRightMargin, _
                            Visio.VisCellIndices.visPrintPropertiesScaleX, _
                            Visio.VisCellIndices.visPrintPropertiesScaleY, _
                            Visio.VisCellIndices.visPrintPropertiesTopMargin}
                    Case Visio.VisRowIndices.visRowRulerGrid
                        'Ruler and Grid SubSection
                        vsoCellsArray = New Visio.VisCellIndices() { _
                            Visio.VisCellIndices.visXGridDensity, _
                            Visio.VisCellIndices.visXGridOrigin, _
                            Visio.VisCellIndices.visXGridSpacing, _
                            Visio.VisCellIndices.visXRulerDensity, _
                            Visio.VisCellIndices.visXRulerOrigin, _
                            Visio.VisCellIndices.visYGridDensity, _
                            Visio.VisCellIndices.visYGridOrigin, _
                            Visio.VisCellIndices.visYGridSpacing, _
                            Visio.VisCellIndices.visYRulerDensity, _
                            Visio.VisCellIndices.visYRulerOrigin}
                    Case Visio.VisRowIndices.visRowShapeLayout
                        'Shape Layout SubSection
                        vsoCellsArray = New Visio.VisCellIndices() { _
                            Visio.VisCellIndices.visSLOConFixedCode, _
                            Visio.VisCellIndices.visSLOFixedCode, _
                            Visio.VisCellIndices.visSLOJumpCode, _
                            Visio.VisCellIndices.visSLOJumpDirX, _
                            Visio.VisCellIndices.visSLOJumpDirY, _
                            Visio.VisCellIndices.visSLOJumpStyle, _
                            Visio.VisCellIndices.visSLOLineRouteExt, _
                            Visio.VisCellIndices.visSLOPermeablePlace, _
                            Visio.VisCellIndices.visSLOPermX, _
                            Visio.VisCellIndices.visSLOPermY, _
                            Visio.VisCellIndices.visSLOPlaceFlip, _
                            Visio.VisCellIndices.visSLOPlaceStyle, _
                            Visio.VisCellIndices.visSLOPlowCode, _
                            Visio.VisCellIndices.visSLORouteStyle, _
                            Visio.VisCellIndices.visSLOSplit, _
                            Visio.VisCellIndices.visSLOSplittable}
                    Case Visio.VisRowIndices.visRowStyle
                        'Style SubSection
                        vsoCellsArray = New Visio.VisCellIndices() { _
                            Visio.VisCellIndices.visStyleHidden, _
                            Visio.VisCellIndices.visStyleIncludesFill, _
                            Visio.VisCellIndices.visStyleIncludesLine, _
                            Visio.VisCellIndices.visStyleIncludesText}
                    Case Visio.VisRowIndices.visRowTextXForm
                        'Text Transform SubSection
                        vsoCellsArray = New Visio.VisCellIndices() { _
                            Visio.VisCellIndices.visXFormAngle, _
                            Visio.VisCellIndices.visXFormHeight, _
                            Visio.VisCellIndices.visXFormLocPinX, _
                            Visio.VisCellIndices.visXFormLocPinY, _
                            Visio.VisCellIndices.visXFormPinX, _
                            Visio.VisCellIndices.visXFormPinY, _
                            Visio.VisCellIndices.visXFormWidth}
                    Case Visio.VisRowIndices.visRowText
                        'Text Block Format SubSection
                        vsoCellsArray = New Visio.VisCellIndices() { _
                            Visio.VisCellIndices.visTxtBlkBkgnd, _
                            Visio.VisCellIndices.visTxtBlkBkgndTrans, _
                            Visio.VisCellIndices.visTxtBlkBottomMargin, _
                            Visio.VisCellIndices.visTxtBlkDefaultTabStop, _
                            Visio.VisCellIndices.visTxtBlkDirection, _
                            Visio.VisCellIndices.visTxtBlkLeftMargin, _
                            Visio.VisCellIndices.visTxtBlkRightMargin, _
                            Visio.VisCellIndices.visTxtBlkTopMargin, _
                            Visio.VisCellIndices.visTxtBlkVerticalAlign}
                    Case Visio.VisRowIndices.visRowXFormOut
                        'Shape Transform SubSection
                        vsoCellsArray = New Visio.VisCellIndices() { _
                            Visio.VisCellIndices.visXFormAngle, _
                            Visio.VisCellIndices.visXFormFlipX, _
                            Visio.VisCellIndices.visXFormFlipY, _
                            Visio.VisCellIndices.visXFormHeight, _
                            Visio.VisCellIndices.visXFormLocPinX, _
                            Visio.VisCellIndices.visXFormLocPinY, _
                            Visio.VisCellIndices.visXFormPinX, _
                            Visio.VisCellIndices.visXFormPinY, _
                            Visio.VisCellIndices.visXFormResizeMode, _
                            Visio.VisCellIndices.visXFormWidth}
                End Select
            Case Visio.VisSectionIndices.visSectionParagraph
                Select Case vsoRowIndex
                    Case Visio.VisRowIndices.visRowParagraph
                        'First and only row of Paragraph Section
                        vsoCellsArray = New Visio.VisCellIndices() { _
                            Visio.VisCellIndices.visBulletFontSize, _
                            Visio.VisCellIndices.visBulletFont, _
                            Visio.VisCellIndices.visBulletIndex, _
                            Visio.VisCellIndices.visBulletString, _
                            Visio.VisCellIndices.visFlags, _
                            Visio.VisCellIndices.visHorzAlign, _
                            Visio.VisCellIndices.visIndentFirst, _
                            Visio.VisCellIndices.visIndentLeft, _
                            Visio.VisCellIndices.visIndentRight, _
                            Visio.VisCellIndices.visLocalizeBulletFont, _
                            Visio.VisCellIndices.visSpaceAfter, _
                            Visio.VisCellIndices.visSpaceBefore, _
                            Visio.VisCellIndices.visSpaceLine, _
                            Visio.VisCellIndices.visTextPosAfterBullet}
                End Select
            Case Visio.VisSectionIndices.visSectionProp
                Select Case vsoRowIndex
                    Case Visio.VisRowIndices.visRowProp
                        'First row of the Custom Properties Section
                        vsoCellsArray = New Visio.VisCellIndices() { _
                            Visio.VisCellIndices.visCustPropsAsk, _
                            Visio.VisCellIndices.visCustPropsCalendar, _
                            Visio.VisCellIndices.visCustPropsFormat, _
                            Visio.VisCellIndices.visCustPropsInvis, _
                            Visio.VisCellIndices.visCustPropsLabel, _
                            Visio.VisCellIndices.visCustPropsLangID, _
                            Visio.VisCellIndices.visCustPropsPrompt, _
                            Visio.VisCellIndices.visCustPropsSortKey, _
                            Visio.VisCellIndices.visCustPropsType, _
                            Visio.VisCellIndices.visCustPropsValue}
                End Select
            Case Visio.VisSectionIndices.visSectionReviewer
                Select Case vsoRowIndex
                    Case Visio.VisRowIndices.visRowReviewer
                        'First row of the Reviewer Section
                        vsoCellsArray = New Visio.VisCellIndices() { _
                            Visio.VisCellIndices.visReviewerColor, _
                            Visio.VisCellIndices.visReviewerCurrentIndex, _
                            Visio.VisCellIndices.visReviewerInitials, _
                            Visio.VisCellIndices.visReviewerName, _
                            Visio.VisCellIndices.visReviewerReviewerID}
                End Select
            Case Visio.VisSectionIndices.visSectionScratch
                Select Case vsoRowIndex
                    Case Visio.VisRowIndices.visRowScratch
                        'First row of the Scratch Section
                        vsoCellsArray = New Visio.VisCellIndices() { _
                            Visio.VisCellIndices.visScratchA, _
                            Visio.VisCellIndices.visScratchB, _
                            Visio.VisCellIndices.visScratchC, _
                            Visio.VisCellIndices.visScratchD, _
                            Visio.VisCellIndices.visScratchX, _
                            Visio.VisCellIndices.visScratchY}
                End Select
            Case Visio.VisSectionIndices.visSectionTextField
                Select Case vsoRowIndex
                    Case Visio.VisRowIndices.visRowField
                        'First row of the Text Field Section
                        vsoCellsArray = New Visio.VisCellIndices() { _
                            Visio.VisCellIndices.visFieldCalendar, _
                            Visio.VisCellIndices.visFieldCell, _
                            Visio.VisCellIndices.visFieldEditMode, _
                            Visio.VisCellIndices.visFieldFormat, _
                            Visio.VisCellIndices.visFieldObjectKind, _
                            Visio.VisCellIndices.visFieldType, _
                            Visio.VisCellIndices.visFieldUICategory, _
                            Visio.VisCellIndices.visFieldUICode, _
                            Visio.VisCellIndices.visFieldUIFormat}
                End Select
            Case Visio.VisSectionIndices.visSectionSmartTag
                Select Case vsoRowIndex
                    Case Visio.VisRowIndices.visRowSmartTag
                        'First row of the Smart Tag Section
                        vsoCellsArray = New Visio.VisCellIndices() { _
                            Visio.VisCellIndices.visSmartTagButtonFace, _
                            Visio.VisCellIndices.visSmartTagDescription, _
                            Visio.VisCellIndices.visSmartTagDisabled, _
                            Visio.VisCellIndices.visSmartTagDisplayMode, _
                            Visio.VisCellIndices.visSmartTagName, _
                            Visio.VisCellIndices.visSmartTagXJustify, _
                            Visio.VisCellIndices.visSmartTagX, _
                            Visio.VisCellIndices.visSmartTagYJustify, _
                            Visio.VisCellIndices.visSmartTagY}
                End Select
            Case Visio.VisSectionIndices.visSectionTab
                Select Case vsoRowIndex
                    Case Visio.VisRowIndices.visRowTab
                        'First row of Tab Section
                        vsoCellsArray = New Visio.VisCellIndices() { _
                            Visio.VisCellIndices.visTabAlign, _
                            Visio.VisCellIndices.visTabPos, _
                            Visio.VisCellIndices.visTabStopCount}
                End Select
            Case Visio.VisSectionIndices.visSectionUser
                Select Case vsoRowIndex
                    Case Visio.VisRowIndices.visRowUser
                        'First row of the User Section
                        vsoCellsArray = New Visio.VisCellIndices() { _
                            Visio.VisCellIndices.visUserPrompt, _
                            Visio.VisCellIndices.visUserValue}
                End Select
            Case Else
                'Test if the section is one of the possible geometry sections.
                If (Visio.VisSectionIndices.visSectionFirstComponent <= vsoSectionIndex) AndAlso _
                    (vsoSectionIndex <= Visio.VisSectionIndices.visSectionLastComponent) Then
                    'Belongs to Geometry Section which can have a variety of Rows. 
                    Select Case vsoRowIndex
                        Case Visio.VisRowIndices.visRowComponent
                            'Component Row of any Geometry Section
                            vsoCellsArray = New Visio.VisCellIndices() { _
                                Visio.VisCellIndices.visCompNoFill, _
                                Visio.VisCellIndices.visCompNoLine, _
                                Visio.VisCellIndices.visCompNoShow, _
                                Visio.VisCellIndices.visCompNoSnap, _
                                Visio.VisCellIndices.visCompPath}
                        Case Visio.VisRowIndices.visRowVertex
                            Select Case vsoRowTag
                                Case Visio.VisRowTags.visTagArcTo
                                    vsoCellsArray = New Visio.VisCellIndices() { _
                                        Visio.VisCellIndices.visX, _
                                        Visio.VisCellIndices.visY, _
                                        Visio.VisCellIndices.visBow}
                                Case Visio.VisRowTags.visTagEllipse
                                    vsoCellsArray = New Visio.VisCellIndices() { _
                                        Visio.VisCellIndices.visEllipseCenterX, _
                                        Visio.VisCellIndices.visEllipseCenterY, _
                                        Visio.VisCellIndices.visEllipseMajorX, _
                                        Visio.VisCellIndices.visEllipseMajorY, _
                                        Visio.VisCellIndices.visEllipseMinorX, _
                                        Visio.VisCellIndices.visEllipseMinorY}
                                Case Visio.VisRowTags.visTagEllipticalArcTo
                                    vsoCellsArray = New Visio.VisCellIndices() { _
                                        Visio.VisCellIndices.visX, _
                                        Visio.VisCellIndices.visY, _
                                        Visio.VisCellIndices.visControlX, _
                                        Visio.VisCellIndices.visControlY, _
                                        Visio.VisCellIndices.visEccentricityAngle, _
                                        Visio.VisCellIndices.visAspectRatio}
                                Case Visio.VisRowTags.visTagInfiniteLine
                                    vsoCellsArray = New Visio.VisCellIndices() { _
                                        Visio.VisCellIndices.visInfiniteLineX1, _
                                        Visio.VisCellIndices.visInfiniteLineY1, _
                                        Visio.VisCellIndices.visInfiniteLineX2, _
                                        Visio.VisCellIndices.visInfiniteLineY2}
                                Case Visio.VisRowTags.visTagLineTo
                                    vsoCellsArray = New Visio.VisCellIndices() { _
                                        Visio.VisCellIndices.visX, _
                                        Visio.VisCellIndices.visY}
                                Case Visio.VisRowTags.visTagMoveTo
                                    vsoCellsArray = New Visio.VisCellIndices() { _
                                        Visio.VisCellIndices.visX, _
                                        Visio.VisCellIndices.visY}
                                Case Visio.VisRowTags.visTagNURBSTo
                                    vsoCellsArray = New Visio.VisCellIndices() { _
                                        Visio.VisCellIndices.visX, _
                                        Visio.VisCellIndices.visY, _
                                        Visio.VisCellIndices.visNURBSKnot, _
                                        Visio.VisCellIndices.visNURBSWeight, _
                                        Visio.VisCellIndices.visNURBSKnotPrev, _
                                        Visio.VisCellIndices.visNURBSWeightPrev, _
                                        Visio.VisCellIndices.visNURBSData}
                                Case Visio.VisRowTags.visTagPolylineTo
                                    vsoCellsArray = New Visio.VisCellIndices() { _
                                        Visio.VisCellIndices.visX, _
                                        Visio.VisCellIndices.visY, _
                                        Visio.VisCellIndices.visPolylineData}
                                Case Visio.VisRowTags.visTagSplineBeg
                                    vsoCellsArray = New Visio.VisCellIndices() { _
                                        Visio.VisCellIndices.visX, _
                                        Visio.VisCellIndices.visY, _
                                        Visio.VisCellIndices.visSplineKnot, _
                                        Visio.VisCellIndices.visSplineKnot2, _
                                        Visio.VisCellIndices.visSplineKnot3, _
                                        Visio.VisCellIndices.visSplineDegree}
                                Case Visio.VisRowTags.visTagSplineSpan
                                    vsoCellsArray = New Visio.VisCellIndices() { _
                                        Visio.VisCellIndices.visX, _
                                        Visio.VisCellIndices.visY, _
                                        Visio.VisCellIndices.visSplineKnot}
                            End Select
                            'Omitted because it isn't documented in the SDK
                            'Case Visio.VisRowTags.visTagBase
                    End Select
                End If
                'Else there are no cells to return, so return the empty array.
        End Select

        'Return the array.
        Return vsoCellsArray
    End Function

    ''' <summary>
    ''' Gets the section prefix.
    ''' </summary>
    ''' <param name="vsoSectionIndex">Index of the vso section.</param>
    ''' <returns></returns>
    Public Function GetSectionPrefix(ByVal vsoSectionIndex As Visio.VisSectionIndices) As String
        Select Case vsoSectionIndex
            Case Visio.VisSectionIndices.visSectionAction
                Return "Actions."
            Case Visio.VisSectionIndices.visSectionConnectionPts
                Return "Connections."
            Case Visio.VisSectionIndices.visSectionControls
                Return "Controls."
            Case Visio.VisSectionIndices.visSectionProp
                Return "Prop."
            Case Visio.VisSectionIndices.visSectionHyperlink
                Return "Hyperlink."
            Case Visio.VisSectionIndices.visSectionUser
                Return "User."
            Case Else
                Return Nothing
        End Select
    End Function

    ''' <summary>
    ''' Tests whether or not it is a valid option to add a particular section to a shape depending on the shape's type.
    ''' </summary>
    ''' <param name="vsoShape">The shape to test.</param>
    ''' <param name="vsoSection">The section to add.</param>
    ''' <returns>
    ''' <c>true</c>if it is valid to add the section to a shape of type <paramref name=" vsoShape"></paramref>.Type;otherwise <c>false</c>
    ''' </returns>
    <Extension()> _
    Public Function SectionIsValidToAdd(ByVal vsoShape As Visio.Shape, _
                                        ByVal vsoSection As Visio.VisSectionIndices) As Boolean
        Select Case vsoShape.Type
            Case Visio.VisShapeTypes.visTypeDoc
                Select Case vsoSection
                    Case _
                        Visio.VisSectionIndices.visSectionHyperlink, _
                        Visio.VisSectionIndices.visSectionProp, _
                        Visio.VisSectionIndices.visSectionReviewer, _
                        Visio.VisSectionIndices.visSectionScratch, _
                        Visio.VisSectionIndices.visSectionUser
                        Return True
                    Case Else
                        Return False
                End Select
            Case Visio.VisShapeTypes.visTypePage
                Select Case vsoSection
                    Case _
                        Visio.VisSectionIndices.visSectionAction, _
                        Visio.VisSectionIndices.visSectionAnnotation, _
                        Visio.VisSectionIndices.visSectionHyperlink, _
                        Visio.VisSectionIndices.visSectionLayer, _
                        Visio.VisSectionIndices.visSectionProp, _
                        Visio.VisSectionIndices.visSectionScratch, _
                        Visio.VisSectionIndices.visSectionSmartTag, _
                        Visio.VisSectionIndices.visSectionUser
                        Return True
                    Case Else
                        Return False
                End Select
            Case Visio.VisShapeTypes.visTypeBitmap, _
                Visio.VisShapeTypes.visTypeForeignObject, _
                Visio.VisShapeTypes.visTypeGroup, _
                Visio.VisShapeTypes.visTypeGuide, _
                Visio.VisShapeTypes.visTypeShape
                Select Case vsoSection
                    Case _
                        Visio.VisSectionIndices.visSectionAction, _
                        Visio.VisSectionIndices.visSectionCharacter, _
                        Visio.VisSectionIndices.visSectionConnectionPts, _
                        Visio.VisSectionIndices.visSectionControls, _
                        Visio.VisSectionIndices.visSectionHyperlink, _
                        Visio.VisSectionIndices.visSectionMember, _
                        Visio.VisSectionIndices.visSectionParagraph, _
                        Visio.VisSectionIndices.visSectionProp, _
                        Visio.VisSectionIndices.visSectionTab, _
                        Visio.VisSectionIndices.visSectionTextField, _
                        Visio.VisSectionIndices.visSectionScratch, _
                        Visio.VisSectionIndices.visSectionSmartTag, _
                        Visio.VisSectionIndices.visSectionUser
                        Return True
                    Case Else
                        Return False
                End Select
            Case Else
                Return False
        End Select
    End Function
#End Region

#End Region

    ''' <summary>
    ''' Copies the shape data from one shape to another. 
    ''' </summary>
    ''' <param name="sourceShape">The source shape.</param>
    ''' <param name="destShape">The destination shape.</param>
    ''' <param name="getSetArgs">Arguments which determine how formulas are set: <see cref="Visio.VisGetSetArgs">visGetSetArgs</see></param>
    ''' <param name="copyArgs">Arguments which determine how data is copied: <see cref="visCopyShapeDataArgs">visCopyShapeDataArgs</see></param>
    <Extension()> _
    Public Sub CopyShapeData(ByVal sourceShape As Visio.Shape, _
                             ByVal destShape As Visio.Shape, _
                             ByVal getSetArgs As Short, _
                             ByVal copyArgs As Short)

        If sourceShape.SectionExists(Visio.VisSectionIndices.visSectionProp, Visio.VisExistsFlags.visExistsAnywhere) AndAlso _
        sourceShape.Section(Visio.VisSectionIndices.visSectionProp).Count > 0 Then
            'Process Bitwise Flags
            Dim blastGuards As Boolean = Not ((getSetArgs And Visio.VisGetSetArgs.visSetBlastGuards) = 0)
            Dim testCircular As Boolean = Not ((getSetArgs And Visio.VisGetSetArgs.visSetTestCircular) = 0)
            Dim universalSyntax As Boolean = Not ((getSetArgs And Visio.VisGetSetArgs.visSetUniversalSyntax) = 0)

            Dim addIfNonExisting As Boolean = Not ((copyArgs And VisioPowerDevTools.visCopyShapeDataArgs.addIfNonExisting) = 0)
            Dim protectReferences As Boolean = Not ((copyArgs And VisioPowerDevTools.visCopyShapeDataArgs.protectReferences) = 0)
            Dim deleteIfNoMatch As Boolean = Not ((copyArgs And VisioPowerDevTools.visCopyShapeDataArgs.deleteIfNoMatch) = 0)

            Dim sourceSection As Visio.Section = sourceShape.Section(Visio.VisSectionIndices.visSectionProp)
            Dim sourceRow As Visio.Row
            Dim destRow As Visio.Row
            Dim vsoRowName As String
            Try
                If addIfNonExisting And protectReferences Then
                    If blastGuards And universalSyntax Then
                        'Add the custom properties section if it doesn't already exist. 
                        If Not destShape.SectionExists(Visio.VisSectionIndices.visSectionProp, Visio.VisExistsFlags.visExistsAnywhere) Then
                            destShape.AddSection(Visio.VisSectionIndices.visSectionProp)
                        End If

                        For i = 0 To sourceSection.Count - 1
                            sourceRow = sourceSection.Row(i)
                            vsoRowName = sourceRow.NameU
                            If Not destShape.CellExistsU("Prop." & vsoRowName, Visio.VisExistsFlags.visExistsAnywhere) Then
                                destShape.AddNamedRow(Visio.VisSectionIndices.visSectionProp, vsoRowName, Visio.VisRowTags.visTagDefault)
                            End If
                            destRow = destShape.CellsU("Prop." & vsoRowName).ContainingRow

                            For j = 0 To sourceRow.Count - 1
                                If destRow(j).FormulaU.Contains("TheDoc!") OrElse destRow(j).FormulaU.Contains("ThePage!") Then
                                    Continue For
                                Else
                                    destRow.CellU(j).FormulaForceU = sourceRow.CellU(j).FormulaU
                                End If
                            Next
                        Next
                    ElseIf blastGuards Then
                        If Not destShape.SectionExists(Visio.VisSectionIndices.visSectionProp, Visio.VisExistsFlags.visExistsAnywhere) Then
                            destShape.AddSection(Visio.VisSectionIndices.visSectionProp)
                        End If

                        For i = 0 To sourceSection.Count - 1
                            sourceRow = sourceSection.Row(i)
                            vsoRowName = sourceRow.Name
                            If Not destShape.CellExistsU("Prop." & vsoRowName, Visio.VisExistsFlags.visExistsAnywhere) Then
                                destShape.AddNamedRow(Visio.VisSectionIndices.visSectionProp, vsoRowName, Visio.VisRowTags.visTagDefault)
                            End If
                            destRow = destShape.Cells("Prop." & vsoRowName).ContainingRow
                            For j = 0 To sourceRow.Count - 1
                                If destRow(j).Formula.Contains("TheDoc!") OrElse destRow(j).Formula.Contains("ThePage!") Then
                                    Continue For
                                Else
                                    destRow.CellU(j).FormulaForce = sourceRow.CellU(j).Formula
                                End If
                            Next
                        Next
                    ElseIf universalSyntax Then
                        If Not destShape.SectionExists(Visio.VisSectionIndices.visSectionProp, Visio.VisExistsFlags.visExistsAnywhere) Then
                            destShape.AddSection(Visio.VisSectionIndices.visSectionProp)
                        End If

                        For i = 0 To sourceSection.Count - 1
                            sourceRow = sourceSection.Row(i)
                            vsoRowName = sourceRow.NameU
                            If Not destShape.CellExistsU("Prop." & vsoRowName, Visio.VisExistsFlags.visExistsAnywhere) Then
                                destShape.AddNamedRow(Visio.VisSectionIndices.visSectionProp, vsoRowName, Visio.VisRowTags.visTagDefault)
                            End If
                            destRow = destShape.CellsU("Prop." & vsoRowName).ContainingRow
                            For j = 0 To sourceRow.Count - 1
                                If destRow(j).FormulaU.Contains("TheDoc!") OrElse destRow(j).FormulaU.Contains("ThePage!") Then
                                    Continue For
                                Else
                                    destRow.CellU(j).FormulaU = sourceRow.CellU(j).FormulaU
                                End If
                            Next
                        Next
                    Else
                        If Not destShape.SectionExists(Visio.VisSectionIndices.visSectionProp, Visio.VisExistsFlags.visExistsAnywhere) Then
                            destShape.AddSection(Visio.VisSectionIndices.visSectionProp)
                        End If

                        For i = 0 To sourceSection.Count - 1
                            sourceRow = sourceSection.Row(i)
                            vsoRowName = sourceRow.Name
                            If Not destShape.CellExistsU("Prop." & vsoRowName, Visio.VisExistsFlags.visExistsAnywhere) Then
                                destShape.AddNamedRow(Visio.VisSectionIndices.visSectionProp, vsoRowName, Visio.VisRowTags.visTagDefault)
                            End If
                            destRow = destShape.Cells("Prop." & vsoRowName).ContainingRow
                            For j = 0 To sourceRow.Count - 1
                                If destRow(j).Formula.Contains("TheDoc!") OrElse destRow(j).Formula.Contains("ThePage!") Then
                                    Continue For
                                Else
                                    destRow.CellU(j).Formula = sourceRow.CellU(j).Formula
                                End If
                            Next
                        Next
                    End If
                ElseIf addIfNonExisting Then
                    If blastGuards And universalSyntax Then
                        'Add the custom properties section if it doesn't already exist. 
                        If Not destShape.SectionExists(Visio.VisSectionIndices.visSectionProp, Visio.VisExistsFlags.visExistsAnywhere) Then
                            destShape.AddSection(Visio.VisSectionIndices.visSectionProp)
                        End If

                        For i = 0 To sourceSection.Count - 1
                            sourceRow = sourceSection.Row(i)
                            vsoRowName = sourceRow.NameU
                            If Not destShape.CellExistsU("Prop." & vsoRowName, Visio.VisExistsFlags.visExistsAnywhere) Then
                                destShape.AddNamedRow(Visio.VisSectionIndices.visSectionProp, vsoRowName, Visio.VisRowTags.visTagDefault)
                            End If
                            destRow = destShape.CellsU("Prop." & vsoRowName).ContainingRow

                            For j = 0 To sourceRow.Count - 1
                                destRow.Cell(j).FormulaForceU = sourceRow.Cell(j).FormulaU
                            Next
                        Next
                    ElseIf blastGuards Then
                        If Not destShape.SectionExists(Visio.VisSectionIndices.visSectionProp, Visio.VisExistsFlags.visExistsAnywhere) Then
                            destShape.AddSection(Visio.VisSectionIndices.visSectionProp)
                        End If

                        For i = 0 To sourceSection.Count - 1
                            sourceRow = sourceSection.Row(i)
                            vsoRowName = sourceRow.Name
                            If Not destShape.CellExistsU("Prop." & vsoRowName, Visio.VisExistsFlags.visExistsAnywhere) Then
                                destShape.AddNamedRow(Visio.VisSectionIndices.visSectionProp, vsoRowName, Visio.VisRowTags.visTagDefault)
                            End If
                            destRow = destShape.Cells("Prop." & vsoRowName).ContainingRow
                            For j = 0 To sourceRow.Count - 1
                                destRow.CellU(j).FormulaForce = sourceRow.CellU(j).Formula
                            Next
                        Next
                    ElseIf universalSyntax Then
                        If Not destShape.SectionExists(Visio.VisSectionIndices.visSectionProp, Visio.VisExistsFlags.visExistsAnywhere) Then
                            destShape.AddSection(Visio.VisSectionIndices.visSectionProp)
                        End If

                        For i = 0 To sourceSection.Count - 1
                            sourceRow = sourceSection.Row(i)
                            vsoRowName = sourceRow.NameU
                            If Not destShape.CellExistsU("Prop." & vsoRowName, Visio.VisExistsFlags.visExistsAnywhere) Then
                                destShape.AddNamedRow(Visio.VisSectionIndices.visSectionProp, vsoRowName, Visio.VisRowTags.visTagDefault)
                            End If
                            destRow = destShape.CellsU("Prop." & vsoRowName).ContainingRow
                            For j = 0 To sourceRow.Count - 1
                                destRow.CellU(j).FormulaU = sourceRow.CellU(j).FormulaU
                            Next
                        Next
                    Else
                        If Not destShape.SectionExists(Visio.VisSectionIndices.visSectionProp, Visio.VisExistsFlags.visExistsAnywhere) Then
                            destShape.AddSection(Visio.VisSectionIndices.visSectionProp)
                        End If

                        For i = 0 To sourceSection.Count - 1
                            sourceRow = sourceSection.Row(i)
                            vsoRowName = sourceRow.Name
                            If Not destShape.CellExistsU("Prop." & vsoRowName, Visio.VisExistsFlags.visExistsAnywhere) Then
                                destShape.AddNamedRow(Visio.VisSectionIndices.visSectionProp, vsoRowName, Visio.VisRowTags.visTagDefault)
                            End If
                            destRow = destShape.Cells("Prop." & vsoRowName).ContainingRow
                            For j = 0 To sourceRow.Count - 1
                                destRow.CellU(j).Formula = sourceRow.CellU(j).Formula
                            Next
                        Next
                    End If
                ElseIf protectReferences Then
                    'If the destination shape doesn't have the properties section, and we aren't adding it, exit. 
                    If Not destShape.SectionExists(Visio.VisSectionIndices.visSectionProp, Visio.VisExistsFlags.visExistsAnywhere) Then
                        Exit Sub
                    End If
                    If blastGuards And universalSyntax Then
                        For i = 0 To sourceSection.Count - 1
                            sourceRow = sourceSection.Row(i)
                            vsoRowName = sourceRow.NameU
                            If destShape.CellExistsU("Prop." & vsoRowName, Visio.VisExistsFlags.visExistsAnywhere) Then
                                destRow = destShape.CellsU("Prop." & vsoRowName).ContainingRow
                                For j = 0 To sourceRow.Count - 1
                                    If destRow(j).FormulaU.Contains("TheDoc!") OrElse destRow(j).FormulaU.Contains("ThePage!") Then
                                        Continue For
                                    Else
                                        destRow.CellU(j).FormulaForceU = sourceRow.CellU(j).FormulaU
                                    End If
                                Next
                            End If
                        Next
                    ElseIf blastGuards Then
                        For i = 0 To sourceSection.Count - 1
                            sourceRow = sourceSection.Row(i)
                            vsoRowName = sourceRow.Name
                            If destShape.CellExistsU("Prop." & vsoRowName, Visio.VisExistsFlags.visExistsAnywhere) Then
                                destRow = destShape.Cells("Prop." & vsoRowName).ContainingRow
                                For j = 0 To sourceRow.Count - 1
                                    If destRow(j).Formula.Contains("TheDoc!") OrElse destRow(j).Formula.Contains("ThePage!") Then
                                        Continue For
                                    Else
                                        destRow.CellU(j).FormulaForce = sourceRow.CellU(j).Formula
                                    End If
                                Next
                            End If
                        Next
                    ElseIf universalSyntax Then
                        For i = 0 To sourceSection.Count - 1
                            sourceRow = sourceSection.Row(i)
                            vsoRowName = sourceRow.NameU
                            If destShape.CellExistsU("Prop." & vsoRowName, Visio.VisExistsFlags.visExistsAnywhere) Then
                                destRow = destShape.CellsU("Prop." & vsoRowName).ContainingRow
                                For j = 0 To sourceRow.Count - 1
                                    If destRow(j).FormulaU.Contains("TheDoc!") OrElse destRow(j).FormulaU.Contains("ThePage!") Then
                                        Continue For
                                    Else
                                        destRow.CellU(j).FormulaU = sourceRow.CellU(j).FormulaU
                                    End If
                                Next
                            End If
                        Next
                    Else
                        For i = 0 To sourceSection.Count - 1
                            sourceRow = sourceSection.Row(i)
                            vsoRowName = sourceRow.Name
                            If destShape.CellExistsU("Prop." & vsoRowName, Visio.VisExistsFlags.visExistsAnywhere) Then
                                destRow = destShape.Cells("Prop." & vsoRowName).ContainingRow
                                For j = 0 To sourceRow.Count - 1
                                    If destRow(j).Formula.Contains("TheDoc!") OrElse destRow(j).Formula.Contains("ThePage!") Then
                                        Continue For
                                    Else
                                        destRow.CellU(j).Formula = sourceRow.CellU(j).Formula
                                    End If
                                Next
                            End If
                        Next
                    End If
                Else
                    'If the destination shape doesn't have the properties section, and we aren't adding it, exit. 
                    If Not destShape.SectionExists(Visio.VisSectionIndices.visSectionProp, Visio.VisExistsFlags.visExistsAnywhere) Then
                        Exit Sub
                    End If
                    If blastGuards And universalSyntax Then
                        For i = 0 To sourceSection.Count - 1
                            sourceRow = sourceSection.Row(i)
                            vsoRowName = sourceRow.NameU
                            If destShape.CellExistsU("Prop." & vsoRowName, Visio.VisExistsFlags.visExistsAnywhere) Then
                                destRow = destShape.CellsU("Prop." & vsoRowName).ContainingRow
                                For j = 0 To sourceRow.Count - 1
                                    destRow.CellU(j).FormulaForceU = sourceRow.CellU(j).FormulaU
                                Next
                            End If
                        Next
                    ElseIf blastGuards Then
                        For i = 0 To sourceSection.Count - 1
                            sourceRow = sourceSection.Row(i)
                            vsoRowName = sourceRow.Name
                            If destShape.CellExists("Prop." & vsoRowName, Visio.VisExistsFlags.visExistsAnywhere) Then
                                destRow = destShape.Cells("Prop." & vsoRowName).ContainingRow
                                For j = 0 To sourceRow.Count - 1
                                    destRow.Cell(j).FormulaForce = sourceRow.Cell(j).Formula
                                Next
                            End If
                        Next
                    ElseIf universalSyntax Then
                        For i = 0 To sourceSection.Count - 1
                            sourceRow = sourceSection.Row(i)
                            vsoRowName = sourceRow.NameU
                            If destShape.CellExistsU("Prop." & vsoRowName, Visio.VisExistsFlags.visExistsAnywhere) Then
                                destRow = destShape.CellsU("Prop." & vsoRowName).ContainingRow
                                For j = 0 To sourceRow.Count - 1
                                    destRow.CellU(j).FormulaU = sourceRow.CellU(j).FormulaU
                                Next
                            End If
                        Next
                    Else
                        For i = 0 To sourceSection.Count - 1
                            sourceRow = sourceSection.Row(i)
                            vsoRowName = sourceRow.Name
                            If destShape.CellExists("Prop." & vsoRowName, Visio.VisExistsFlags.visExistsAnywhere) Then
                                destRow = destShape.Cells("Prop." & vsoRowName).ContainingRow
                                For j = 0 To sourceRow.Count - 1
                                    destRow.Cell(j).FormulaForce = sourceRow.Cell(j).Formula
                                Next
                            End If
                        Next
                    End If
                End If

                If deleteIfNoMatch Then
                    'Using a Do loop and manual counter manipulation allows me to 
                    'delete rows within the loop instead of having to maintain a seperate list or array 
                    'to post-process.

                    'Set counters up before the loop. 
                    'Do Until > vsoSection.Count - 1 reEvaluates count each loop, which won't work. 
                    Dim destSection As Visio.Section = destShape.Section(Visio.VisSectionIndices.visSectionProp)
                    Dim upperLimit As Integer = destSection.Count - 1
                    Dim vsoRowIndex As Integer = 0
                    'Loop through each row in the section.
                    Do Until vsoRowIndex > upperLimit
                        'See if the destShape Row exists in the source shape.
                        If sourceShape.CellExistsU("Prop." & destSection.Row(vsoRowIndex).NameU, Visio.VisExistsFlags.visExistsAnywhere) Then
                            'There is a match, we don't want to delete it.
                            'Move to the next row.
                            vsoRowIndex += 1
                        Else
                            'There is NO match, we need to delete that shapesheet row. 
                            destShape.DeleteRow(destSection.Index, CShort(vsoRowIndex))
                            'We reduced the row count with this action
                            upperLimit -= 1
                            'Since we deleted the current row, there is no need to increment the row index. 
                        End If
                    Loop
                End If
            Catch ex As Exception
                Throw
            End Try
        End If


    End Sub

    ''' <summary>
    ''' Moves the connections from one shape to another.
    ''' </summary>
    ''' <param name="oldShape">The old shape.</param>
    ''' <param name="newShape">The new shape.</param>
    ''' <remarks>
    ''' Per VisioSDK: this method glues to The pin of a 2-D shape and does the following:
    ''' The shape being glued from must be routable (ObjType includes visLOFlagsRoutable ) 
    ''' or have a dynamic glue type (GlueType includes visGlueTypeWalking ), 
    ''' and does not prohibit dynamic glue (GlueType does not include visGlueTypeNoWalking ). 
    ''' Gluing to PinX creates dynamic glue with a horizontal walking preference and gluing to 
    ''' PinY creates dynamic glue with a vertical walking preference.
    ''' </remarks>
    <Extension()> _
    Public Sub MoveConnections(ByVal oldShape As Visio.Shape, _
                               ByVal newShape As Visio.Shape)
        'Need the page to reference shapeIDs.
        Dim vsoPage As Visio.Page = oldShape.ContainingPage

        'Process the connections.
        Dim newShapeCell As Visio.Cell = Nothing
        Dim shapeSectionIndex As Visio.VisSectionIndices
        Dim shapeRowIndex As Visio.VisRowIndices
        Dim shapeCellIndex As Visio.VisCellIndices

        'These connections are from the shape to a connector (e.g. Guide), so GlueTo does not remove any connections from the oldShape.
        For Each vsoConnect As Visio.Connect In oldShape.Connects
            'Connects are 1-Indexed... stupid.
            'Get the SRC information of the cell in the oldShape which the connector is connected to.
            shapeSectionIndex = vsoConnect.FromCell.Section
            shapeRowIndex = vsoConnect.FromCell.Row
            shapeCellIndex = vsoConnect.FromCell.Column
            newShapeCell = newShape.CellsSRC(shapeSectionIndex, shapeRowIndex, shapeCellIndex)

            newShapeCell.GlueTo(vsoConnect.ToCell)
        Next

        'From connects are made to the shape from a connector, so GlueTo removes the connection from the collection.
        'Since the collection gets polled at each iteration, we can't process in-loop since the counter will get screwed up.
        'Save the connection data to local variables, using GUIDs, since IDs can restructure after GlueTo
        Dim fromConnects As New List(Of String)
        For Each vsoConnect As Visio.Connect In oldShape.FromConnects
            fromConnects.Add(vsoConnect.FromSheet.UniqueID(Visio.VisUniqueIDArgs.visGetOrMakeGUID))
        Next

        'FromConnects collections is all of the connections where the ToSheet Is the shape.
        For j = 0 To fromConnects.Count - 1
            'Get the GUID of the connector itself.
            Dim id As String = fromConnects(j)
            'Use Connects(1) to process the topmost connection on the oldShape.
            Dim vsoConnect As Visio.Connect = oldShape.FromConnects(1)

            'Get the SRC information of the cell in the oldShape which the connector is connected to.
            shapeSectionIndex = vsoConnect.ToCell.Section
            shapeRowIndex = vsoConnect.ToCell.Row
            shapeCellIndex = vsoConnect.ToCell.Column
            newShapeCell = newShape.CellsSRC(shapeSectionIndex, shapeRowIndex, shapeCellIndex)

            'This effectively removes the topmost connection from the collection and slides up one.
            'So we can still use Connects(1)
            vsoConnect.FromCell.GlueTo(newShapeCell)
        Next
    End Sub

#Region "Search And Replace"

#Region "GetUniqueShapeDataFields"

    ''' <summary>
    ''' Gets an array of all the unique shapeData fields contained within all of the shapes in a document. 
    ''' </summary>
    ''' <param name="vsoDoc">The  Visio document to search.</param>
    ''' <param name="fullyQualifiedCellName">A boolean used to indicate whether or not to include the cellname prefix.</param>
    ''' <returns>
    ''' An array of all the unique shapeData fields contained within all of the shapes in a document. 
    ''' </returns>
    ''' <remarks>
    ''' Does not search background pages.
    ''' </remarks>
    Public Function GetUniqueShapeDataFields(ByVal vsoDoc As Visio.Document, _
                                             ByVal fullyQualifiedCellName As Boolean) _
                                             As String()
        Dim allFields As New List(Of String)

        If fullyQualifiedCellName Then
            For Each vsoPage As Visio.Page In vsoDoc.Pages
                If False = CType(vsoPage.Background, Boolean) Then
                    For Each vsoShape As Visio.Shape In vsoPage.Shapes
                        For i = 0 To vsoShape.RowCount(Visio.VisSectionIndices.visSectionProp) - 1
                            allFields.Add("Prop." & vsoShape.Section(Visio.VisSectionIndices.visSectionProp).Row(i).NameU)
                        Next
                    Next
                End If
            Next
        Else
            For Each vsoPage As Visio.Page In vsoDoc.Pages
                If False = CType(vsoPage.Background, Boolean) Then
                    For Each vsoShape As Visio.Shape In vsoPage.Shapes
                        For i = 0 To vsoShape.RowCount(Visio.VisSectionIndices.visSectionProp) - 1
                            allFields.Add(vsoShape.Section(Visio.VisSectionIndices.visSectionProp).Row(i).NameU)
                        Next
                    Next
                End If
            Next
        End If

        Return allFields.Distinct.ToArray

        'Alternative method uses Contains and checks the list before adding rather than Distinct afterwords.
        'This might be faster on smaller datasets, but should only be by a few microseconds.
        'Scaling up, Distinct will be faster.
        'Dim uniqueFields As New List(Of String)
        'Dim nameU As String = ""
        'For Each vsoPage As Visio.Page In vsoDoc.Pages
        '   If False = CType(vsoPage.Background, Boolean) Then
        '   For Each vsoShape As Visio.Shape In vsoPage.Shapes
        '       For i = 0 To vsoShape.RowCount(Visio.VisSectionIndices.visSectionProp) - 1
        '           nameU = vsoShape.Section(Visio.VisSectionIndices.visSectionProp).Row(i).NameU
        '           If False = uniqueFields.Contains(nameU) Then
        '               uniqueFields.Add(nameU)
        '           End If
        '       Next
        '   Next
        '   End If
        'Next
        'Return uniqueFields.ToArray
    End Function

    ''' <summary>
    ''' Gets an array of all the unique shapeData fields contained within all of the shapes in a page. 
    ''' </summary>
    ''' <param name="vsoPage">The  Visio page to search.</param>
    ''' <param name="fullyQualifiedCellName">A boolean used to indicate whether or not to include the cellname prefix.</param>
    ''' <returns>
    ''' An array of all the unique shapeData fields contained within all of the shapes in a page. 
    ''' </returns>
    Public Function GetUniqueShapeDataFields(ByVal vsoPage As Visio.Page, _
                                             ByVal fullyQualifiedCellName As Boolean) _
                                             As String()
        Dim allFields As New List(Of String)
        If fullyQualifiedCellName Then
            For Each vsoShape As Visio.Shape In vsoPage.Shapes
                For i = 0 To vsoShape.RowCount(Visio.VisSectionIndices.visSectionProp) - 1
                    allFields.Add("Prop." & vsoShape.Section(Visio.VisSectionIndices.visSectionProp).Row(i).NameU)
                Next
            Next
        Else
            For Each vsoShape As Visio.Shape In vsoPage.Shapes
                For i = 0 To vsoShape.RowCount(Visio.VisSectionIndices.visSectionProp) - 1
                    allFields.Add(vsoShape.Section(Visio.VisSectionIndices.visSectionProp).Row(i).NameU)
                Next
            Next
        End If

        Return allFields.Distinct.ToArray
    End Function

    ''' <summary>
    ''' Gets an array of all the unique shapeData fields contained within all of the shapes in a selection. 
    ''' </summary>
    ''' <param name="vsoSelection">The  Visio selection to search.</param>
    ''' <param name="fullyQualifiedCellName">A boolean used to indicate whether or not to include the cellname prefix.</param>
    ''' <returns>
    ''' An array of all the unique shapeData fields contained within all of the shapes in a selection. 
    ''' </returns>
    Public Function GetUniqueShapeDataFields(ByVal vsoSelection As Visio.Selection, _
                                             ByVal fullyQualifiedCellName As Boolean) _
                                             As String()
        Dim allFields As New List(Of String)
        If fullyQualifiedCellName Then
            For Each vsoShape As Visio.Shape In vsoSelection
                For i = 0 To vsoShape.RowCount(Visio.VisSectionIndices.visSectionProp) - 1
                    allFields.Add("Prop." & vsoShape.Section(Visio.VisSectionIndices.visSectionProp).Row(i).NameU)
                Next
            Next
        Else
            For Each vsoShape As Visio.Shape In vsoSelection
                For i = 0 To vsoShape.RowCount(Visio.VisSectionIndices.visSectionProp) - 1
                    allFields.Add(vsoShape.Section(Visio.VisSectionIndices.visSectionProp).Row(i).NameU)
                Next
            Next
        End If

        Return allFields.Distinct.ToArray
    End Function

#End Region

#Region "ReplaceTextInShapeData"

    'To replace a piece of Text with a blank string.
    'String.Replace(original, findText, Nothing)
    'Regex.Replace(original, findText, "")
    'ReplaceCaseInsensitive(original, findText, "")


    ''' <summary>
    ''' Replaces one string with another string in any field in the ShapeData section of all the shapes in a document.
    ''' </summary>
    ''' <param name="findText">The old text to replace..</param>
    ''' <param name="replaceText">The new text to add.</param>
    ''' <param name="searchAndReplaceArgs">Arguments used to specify how to search. <seealso cref="visSearchAndReplaceArgs">visSearchAndReplaceArgs</seealso></param>
    ''' <param name="vsoDoc">The Visio document to search.</param>
    ''' <returns>
    ''' The number of replacements performed.
    ''' </returns>
    ''' <remarks>
    ''' To replace a piece of text with a blank string. <paramref name="replaceText">replaceText</paramref>
    ''' should be a blank string: <c>""</c>. 
    ''' </remarks>
    <Extension()> _
    Public Function ReplaceTextInShapeData(ByVal vsoDoc As Visio.Document, _
                                           ByVal findText As String, _
                                           ByVal replaceText As String, _
                                           ByVal searchAndReplaceArgs As Short) _
                                           As Integer
        Dim vsoCell As Visio.Cell
        Dim counter As Integer = 0
        Dim newFormula As String = ""
        'Process Bitwise Flags
        Dim matchCase As Boolean = Not ((searchAndReplaceArgs And VisioPowerDevTools.visSearchAndReplaceArgs.MatchCase) = 0)
        Dim wholeWordsOnly As Boolean = Not ((searchAndReplaceArgs And VisioPowerDevTools.visSearchAndReplaceArgs.WholeWordsOnly) = 0)
        Dim regExSearch As Boolean = Not ((searchAndReplaceArgs And VisioPowerDevTools.visSearchAndReplaceArgs.RegEx) = 0)

        'RegExSearch cannot be used in combination with matchCase and wholeWords Only. 
        If regExSearch Then
            For Each vsoPage As Visio.Page In vsoDoc.Pages
                If False = CType(vsoPage.Background, Boolean) Then
                    For Each vsoShape As Visio.Shape In vsoPage.Shapes
                        For i = 0 To vsoShape.RowCount(Visio.VisSectionIndices.visSectionProp) - 1
                            vsoCell = vsoShape.CellsSRC(Visio.VisSectionIndices.visSectionProp, i, Visio.VisCellIndices.visCustPropsValue)
                            'Treat the findText as a regEx pattern. 
                            newFormula = Regex.Replace(vsoCell.FormulaU, findText, replaceText)
                            If Not String.Equals(vsoCell.FormulaU, newFormula) Then
                                vsoCell.FormulaU = newFormula
                                counter += 1
                            End If
                        Next
                    Next
                End If
            Next
            Return counter
        End If



        'Uses different search methods depending on what is needed. 
        'E.g., don't use RegEx if not required. 
        If matchCase AndAlso wholeWordsOnly Then
            For Each vsoPage As Visio.Page In vsoDoc.Pages
                If False = CType(vsoPage.Background, Boolean) Then
                    For Each vsoShape As Visio.Shape In vsoPage.Shapes
                        For i = 0 To vsoShape.RowCount(Visio.VisSectionIndices.visSectionProp) - 1
                            vsoCell = vsoShape.CellsSRC(Visio.VisSectionIndices.visSectionProp, i, Visio.VisCellIndices.visCustPropsValue)
                            '\bword\b performs a "whole words only" search using a regular expression. 
                            newFormula = Regex.Replace(vsoCell.FormulaU, "\b" & Regex.Escape(findText) & "\b", replaceText)
                            If Not String.Equals(vsoCell.FormulaU, newFormula) Then
                                vsoCell.FormulaU = newFormula
                                counter += 1
                            End If
                        Next
                    Next
                End If
            Next
        ElseIf matchCase Then
            For Each vsoPage As Visio.Page In vsoDoc.Pages
                If False = CType(vsoPage.Background, Boolean) Then
                    For Each vsoShape As Visio.Shape In vsoPage.Shapes
                        For i = 0 To vsoShape.RowCount(Visio.VisSectionIndices.visSectionProp) - 1
                            vsoCell = vsoShape.CellsSRC(Visio.VisSectionIndices.visSectionProp, i, Visio.VisCellIndices.visCustPropsValue)
                            newFormula = vsoCell.FormulaU.Replace(findText, replaceText)
                            If Not String.Equals(vsoCell.FormulaU, newFormula) Then
                                vsoCell.FormulaU = newFormula
                                counter += 1
                            End If
                        Next
                    Next
                End If
            Next
        ElseIf wholeWordsOnly Then
            For Each vsoPage As Visio.Page In vsoDoc.Pages
                If False = CType(vsoPage.Background, Boolean) Then
                    For Each vsoShape As Visio.Shape In vsoPage.Shapes
                        For i = 0 To vsoShape.RowCount(Visio.VisSectionIndices.visSectionProp) - 1
                            vsoCell = vsoShape.CellsSRC(Visio.VisSectionIndices.visSectionProp, i, Visio.VisCellIndices.visCustPropsValue)
                            '\bword\b performs a "whole words only" search using a regular expression. 
                            newFormula = Regex.Replace(vsoCell.FormulaU, "\b" & Regex.Escape(findText) & "\b", replaceText, RegexOptions.IgnoreCase)
                            If Not String.Equals(vsoCell.FormulaU, newFormula) Then
                                vsoCell.FormulaU = newFormula
                                counter += 1
                            End If
                        Next
                    Next
                End If
            Next
        Else
            For Each vsoPage As Visio.Page In vsoDoc.Pages
                If False = CType(vsoPage.Background, Boolean) Then
                    For Each vsoShape As Visio.Shape In vsoPage.Shapes
                        For i = 0 To vsoShape.RowCount(Visio.VisSectionIndices.visSectionProp) - 1
                            vsoCell = vsoShape.CellsSRC(Visio.VisSectionIndices.visSectionProp, i, Visio.VisCellIndices.visCustPropsValue)
                            newFormula = ReplaceCaseInsensitive(vsoCell.FormulaU, findText, replaceText)
                            If Not String.Equals(vsoCell.FormulaU, newFormula) Then
                                vsoCell.FormulaU = newFormula
                                counter += 1
                            End If
                        Next
                    Next
                End If
            Next
        End If

        Return counter
    End Function

    ''' <summary>
    ''' Replaces one string with another string in specified fields in the ShapeData section all the shapes in a document.
    ''' </summary>
    ''' <param name="findText">The old text to replace..</param>
    ''' <param name="replaceText">The new text to add.</param>
    ''' <param name="searchAndReplaceArgs">Arguments used to specify how to search. <seealso cref="visSearchAndReplaceArgs">visSearchAndReplaceArgs</seealso></param>
    ''' <param name="vsoDoc">The Visio document to search.</param>
    ''' <param name="fieldsToSearch">An array of fully qualified cell names representing the ShapeData fields to search.</param>
    ''' <returns>
    ''' The number of replacements performed.
    ''' </returns>
    <Extension()> _
    Public Function ReplaceTextInShapeData(ByVal vsoDoc As Visio.Document, _
                                           ByVal findText As String, _
                                           ByVal replaceText As String, _
                                           ByVal searchAndReplaceArgs As Short, _
                                           ByVal fieldsToSearch As String()) _
                                           As Integer

        Dim vsoCell As Visio.Cell
        Dim counter As Integer = 0
        Dim newFormula As String = ""
        'Process Bitwise Flags
        Dim matchCase As Boolean = Not ((searchAndReplaceArgs And VisioPowerDevTools.visSearchAndReplaceArgs.MatchCase) = 0)
        Dim wholeWordsOnly As Boolean = Not ((searchAndReplaceArgs And VisioPowerDevTools.visSearchAndReplaceArgs.WholeWordsOnly) = 0)
        Dim regExSearch As Boolean = Not ((searchAndReplaceArgs And VisioPowerDevTools.visSearchAndReplaceArgs.RegEx) = 0)

        'RegExSearch cannot be used in combination with matchCase and wholeWords Only. 
        If regExSearch Then
            For Each vsoPage As Visio.Page In vsoDoc.Pages
                If False = CType(vsoPage.Background, Boolean) Then
                    For Each vsoShape As Visio.Shape In vsoPage.Shapes
                        For i = 0 To fieldsToSearch.Count - 1
                            If vsoShape.CellExistsU(fieldsToSearch(i), Visio.VisExistsFlags.visExistsAnywhere) Then
                                vsoCell = vsoShape.CellsU(fieldsToSearch(i))
                                'Treat the findText as a regEx pattern. 
                                newFormula = Regex.Replace(vsoCell.FormulaU, findText, replaceText)
                                If Not String.Equals(vsoCell.FormulaU, newFormula) Then
                                    vsoCell.FormulaU = newFormula
                                    counter += 1
                                End If
                            End If
                        Next
                    Next
                End If
            Next
            Return counter
        End If



        'Uses different search methods depending on what is needed. 
        'E.g., don't use RegEx if not required. 
        If matchCase AndAlso wholeWordsOnly Then
            For Each vsoPage As Visio.Page In vsoDoc.Pages
                If False = CType(vsoPage.Background, Boolean) Then
                    For Each vsoShape As Visio.Shape In vsoPage.Shapes
                        For i = 0 To fieldsToSearch.Count - 1
                            If vsoShape.CellExistsU(fieldsToSearch(i), Visio.VisExistsFlags.visExistsAnywhere) Then
                                vsoCell = vsoShape.CellsU(fieldsToSearch(i))
                                '\bword\b performs a "whole words only" search using a regular expression. 
                                newFormula = Regex.Replace(vsoCell.FormulaU, "\b" & Regex.Escape(findText) & "\b", replaceText)
                                If Not String.Equals(vsoCell.FormulaU, newFormula) Then
                                    vsoCell.FormulaU = newFormula
                                    counter += 1
                                End If
                            End If
                        Next
                    Next
                End If
            Next
        ElseIf matchCase Then
            For Each vsoPage As Visio.Page In vsoDoc.Pages
                If False = CType(vsoPage.Background, Boolean) Then
                    For Each vsoShape As Visio.Shape In vsoPage.Shapes
                        For i = 0 To fieldsToSearch.Count - 1
                            If vsoShape.CellExistsU(fieldsToSearch(i), Visio.VisExistsFlags.visExistsAnywhere) Then
                                vsoCell = vsoShape.CellsU(fieldsToSearch(i))
                                newFormula = vsoCell.FormulaU.Replace(findText, replaceText)
                                If Not String.Equals(vsoCell.FormulaU, newFormula) Then
                                    vsoCell.FormulaU = newFormula
                                    counter += 1
                                End If
                            End If
                        Next
                    Next
                End If
            Next
        ElseIf wholeWordsOnly Then
            For Each vsoPage As Visio.Page In vsoDoc.Pages
                If False = CType(vsoPage.Background, Boolean) Then
                    For Each vsoShape As Visio.Shape In vsoPage.Shapes
                        For i = 0 To fieldsToSearch.Count - 1
                            If vsoShape.CellExistsU(fieldsToSearch(i), Visio.VisExistsFlags.visExistsAnywhere) Then
                                vsoCell = vsoShape.CellsU(fieldsToSearch(i))
                                '\bword\b performs a "whole words only" search using a regular expression. 
                                newFormula = Regex.Replace(vsoCell.FormulaU, "\b" & Regex.Escape(findText) & "\b", replaceText, RegexOptions.IgnoreCase)
                                If Not String.Equals(vsoCell.FormulaU, newFormula) Then
                                    vsoCell.FormulaU = newFormula
                                    counter += 1
                                End If
                            End If
                        Next
                    Next
                End If
            Next
        Else
            For Each vsoPage As Visio.Page In vsoDoc.Pages
                If False = CType(vsoPage.Background, Boolean) Then
                    For Each vsoShape As Visio.Shape In vsoPage.Shapes
                        For i = 0 To fieldsToSearch.Count - 1
                            If vsoShape.CellExistsU(fieldsToSearch(i), Visio.VisExistsFlags.visExistsAnywhere) Then
                                vsoCell = vsoShape.CellsU(fieldsToSearch(i))
                                newFormula = ReplaceCaseInsensitive(vsoCell.FormulaU, findText, replaceText)
                                If Not String.Equals(vsoCell.FormulaU, newFormula) Then
                                    vsoCell.FormulaU = newFormula
                                    counter += 1
                                End If
                            End If
                        Next
                    Next
                End If
            Next
        End If

        Return counter

    End Function

    ''' <summary>
    ''' Replaces one string with another string in any field in the ShapeData section all the shapes in a page.
    ''' </summary>
    ''' <param name="findText">The old text to replace..</param>
    ''' <param name="replaceText">The new text to add.</param>
    ''' <param name="searchAndReplaceArgs">Arguments used to specify how to search. <seealso cref="visSearchAndReplaceArgs">visSearchAndReplaceArgs</seealso></param>
    ''' <param name="vsoPage">The Visio page to search.</param>
    ''' <returns>
    ''' The number of replacements performed.
    ''' </returns>
    <Extension()> _
    Public Function ReplaceTextInShapeData(ByVal vsoPage As Visio.Page, _
                                           ByVal findText As String, _
                                           ByVal replaceText As String, _
                                           ByVal searchAndReplaceArgs As Short) _
                                           As Integer
        Dim vsoCell As Visio.Cell
        Dim counter As Integer = 0
        Dim newFormula As String = ""
        'Process Bitwise Flags
        Dim matchCase As Boolean = Not ((searchAndReplaceArgs And VisioPowerDevTools.visSearchAndReplaceArgs.MatchCase) = 0)
        Dim wholeWordsOnly As Boolean = Not ((searchAndReplaceArgs And VisioPowerDevTools.visSearchAndReplaceArgs.WholeWordsOnly) = 0)
        Dim regExSearch As Boolean = Not ((searchAndReplaceArgs And VisioPowerDevTools.visSearchAndReplaceArgs.RegEx) = 0)

        'RegExSearch cannot be used in combination with matchCase and wholeWords Only. 
        If regExSearch Then
            For Each vsoShape As Visio.Shape In vsoPage.Shapes
                For i = 0 To vsoShape.RowCount(Visio.VisSectionIndices.visSectionProp) - 1
                    vsoCell = vsoShape.CellsSRC(Visio.VisSectionIndices.visSectionProp, i, Visio.VisCellIndices.visCustPropsValue)
                    'Treat the findText as a regEx pattern. 
                    newFormula = Regex.Replace(vsoCell.FormulaU, findText, replaceText)
                    If Not String.Equals(vsoCell.FormulaU, newFormula) Then
                        vsoCell.FormulaU = newFormula
                        counter += 1
                    End If
                Next
            Next
            Return counter
        End If

        'Uses different search methods depending on what is needed. 
        'E.g., don't use RegEx if not required. 
        If matchCase AndAlso wholeWordsOnly Then
            For Each vsoShape As Visio.Shape In vsoPage.Shapes
                For i = 0 To vsoShape.RowCount(Visio.VisSectionIndices.visSectionProp) - 1
                    vsoCell = vsoShape.CellsSRC(Visio.VisSectionIndices.visSectionProp, i, Visio.VisCellIndices.visCustPropsValue)
                    '\bword\b performs a "whole words only" search using a regular expression. 
                    newFormula = Regex.Replace(vsoCell.FormulaU, "\b" & Regex.Escape(findText) & "\b", replaceText)
                    If Not String.Equals(vsoCell.FormulaU, newFormula) Then
                        vsoCell.FormulaU = newFormula
                        counter += 1
                    End If
                Next
            Next
        ElseIf matchCase Then
            For Each vsoShape As Visio.Shape In vsoPage.Shapes
                For i = 0 To vsoShape.RowCount(Visio.VisSectionIndices.visSectionProp) - 1
                    vsoCell = vsoShape.CellsSRC(Visio.VisSectionIndices.visSectionProp, i, Visio.VisCellIndices.visCustPropsValue)
                    newFormula = vsoCell.FormulaU.Replace(findText, replaceText)
                    If Not String.Equals(vsoCell.FormulaU, newFormula) Then
                        vsoCell.FormulaU = newFormula
                        counter += 1
                    End If
                Next
            Next
        ElseIf wholeWordsOnly Then
            For Each vsoShape As Visio.Shape In vsoPage.Shapes
                For i = 0 To vsoShape.RowCount(Visio.VisSectionIndices.visSectionProp) - 1
                    vsoCell = vsoShape.CellsSRC(Visio.VisSectionIndices.visSectionProp, i, Visio.VisCellIndices.visCustPropsValue)
                    '\bword\b performs a "whole words only" search using a regular expression. 
                    newFormula = Regex.Replace(vsoCell.FormulaU, "\b" & Regex.Escape(findText) & "\b", replaceText, RegexOptions.IgnoreCase)
                    If Not String.Equals(vsoCell.FormulaU, newFormula) Then
                        vsoCell.FormulaU = newFormula
                        counter += 1
                    End If
                Next
            Next
        Else
            For Each vsoShape As Visio.Shape In vsoPage.Shapes
                For i = 0 To vsoShape.RowCount(Visio.VisSectionIndices.visSectionProp) - 1
                    vsoCell = vsoShape.CellsSRC(Visio.VisSectionIndices.visSectionProp, i, Visio.VisCellIndices.visCustPropsValue)
                    newFormula = ReplaceCaseInsensitive(vsoCell.FormulaU, findText, replaceText)
                    If Not String.Equals(vsoCell.FormulaU, newFormula) Then
                        vsoCell.FormulaU = newFormula
                        counter += 1
                    End If
                Next
            Next
        End If

        Return counter
    End Function

    ''' <summary>
    ''' Replaces one string with another string in specified fields in the ShapeData section all the shapes in a page.
    ''' </summary>
    ''' <param name="findText">The old text to replace..</param>
    ''' <param name="replaceText">The new text to add.</param>
    ''' <param name="searchAndReplaceArgs">Arguments used to specify how to search. <seealso cref="visSearchAndReplaceArgs">visSearchAndReplaceArgs</seealso></param>
    ''' <param name="vsoPage">The Visio page to search.</param>
    ''' <param name="fieldsToSearch">An array of fully qualified cell names representing the ShapeData fields to search.</param>
    ''' <returns>
    ''' The number of replacements performed.
    ''' </returns>
    <Extension()> _
    Public Function ReplaceTextInShapeData(ByVal vsoPage As Visio.Page, _
                                           ByVal findText As String, _
                                           ByVal replaceText As String, _
                                           ByVal searchAndReplaceArgs As Short, _
                                           ByVal fieldsToSearch As String()) _
                                           As Integer

        Dim vsoCell As Visio.Cell
        Dim counter As Integer = 0
        Dim newFormula As String = ""
        'Process Bitwise Flags
        Dim matchCase As Boolean = Not ((searchAndReplaceArgs And VisioPowerDevTools.visSearchAndReplaceArgs.MatchCase) = 0)
        Dim wholeWordsOnly As Boolean = Not ((searchAndReplaceArgs And VisioPowerDevTools.visSearchAndReplaceArgs.WholeWordsOnly) = 0)
        Dim regExSearch As Boolean = Not ((searchAndReplaceArgs And VisioPowerDevTools.visSearchAndReplaceArgs.RegEx) = 0)

        'RegExSearch cannot be used in combination with matchCase and wholeWords Only. 
        If regExSearch Then
            For Each vsoShape As Visio.Shape In vsoPage.Shapes
                For i = 0 To fieldsToSearch.Count - 1
                    If vsoShape.CellExistsU(fieldsToSearch(i), Visio.VisExistsFlags.visExistsAnywhere) Then
                        vsoCell = vsoShape.CellsU(fieldsToSearch(i))
                        'Treat the findText as a regEx pattern. 
                        newFormula = Regex.Replace(vsoCell.FormulaU, findText, replaceText)
                        If Not String.Equals(vsoCell.FormulaU, newFormula) Then
                            vsoCell.FormulaU = newFormula
                            counter += 1
                        End If
                    End If
                Next
            Next
            Return counter
        End If


        'Uses different search methods depending on what is needed. 
        'E.g., don't use RegEx if not required. 
        If matchCase AndAlso wholeWordsOnly Then
            For Each vsoShape As Visio.Shape In vsoPage.Shapes
                For i = 0 To fieldsToSearch.Count - 1
                    If vsoShape.CellExistsU(fieldsToSearch(i), Visio.VisExistsFlags.visExistsAnywhere) Then
                        vsoCell = vsoShape.CellsU(fieldsToSearch(i))
                        '\bword\b performs a "whole words only" search using a regular expression. 
                        newFormula = Regex.Replace(vsoCell.FormulaU, "\b" & Regex.Escape(findText) & "\b", replaceText)
                        If Not String.Equals(vsoCell.FormulaU, newFormula) Then
                            vsoCell.FormulaU = newFormula
                            counter += 1
                        End If
                    End If
                Next
            Next
        ElseIf matchCase Then
            For Each vsoShape As Visio.Shape In vsoPage.Shapes
                For i = 0 To fieldsToSearch.Count - 1
                    If vsoShape.CellExistsU(fieldsToSearch(i), Visio.VisExistsFlags.visExistsAnywhere) Then
                        vsoCell = vsoShape.CellsU(fieldsToSearch(i))
                        newFormula = vsoCell.FormulaU.Replace(findText, replaceText)
                        If Not String.Equals(vsoCell.FormulaU, newFormula) Then
                            vsoCell.FormulaU = newFormula
                            counter += 1
                        End If
                    End If
                Next
            Next
        ElseIf wholeWordsOnly Then
            For Each vsoShape As Visio.Shape In vsoPage.Shapes
                For i = 0 To fieldsToSearch.Count - 1
                    If vsoShape.CellExistsU(fieldsToSearch(i), Visio.VisExistsFlags.visExistsAnywhere) Then
                        vsoCell = vsoShape.CellsU(fieldsToSearch(i))
                        '\bword\b performs a "whole words only" search using a regular expression. 
                        newFormula = Regex.Replace(vsoCell.FormulaU, "\b" & Regex.Escape(findText) & "\b", replaceText, RegexOptions.IgnoreCase)
                        If Not String.Equals(vsoCell.FormulaU, newFormula) Then
                            vsoCell.FormulaU = newFormula
                            counter += 1
                        End If
                    End If
                Next
            Next
        Else
            For Each vsoShape As Visio.Shape In vsoPage.Shapes
                For i = 0 To fieldsToSearch.Count - 1
                    If vsoShape.CellExistsU(fieldsToSearch(i), Visio.VisExistsFlags.visExistsAnywhere) Then
                        vsoCell = vsoShape.CellsU(fieldsToSearch(i))
                        newFormula = ReplaceCaseInsensitive(vsoCell.FormulaU, findText, replaceText)
                        If Not String.Equals(vsoCell.FormulaU, newFormula) Then
                            vsoCell.FormulaU = newFormula
                            counter += 1
                        End If
                    End If
                Next
            Next
        End If

        Return counter
    End Function

    ''' <summary>
    ''' Replaces one string with another string in any field in the ShapeData section all the shapes in a selection.
    ''' </summary>
    ''' <param name="findText">The old text to replace..</param>
    ''' <param name="replaceText">The new text to add.</param>
    ''' <param name="searchAndReplaceArgs">Arguments used to specify how to search. <seealso cref="visSearchAndReplaceArgs">visSearchAndReplaceArgs</seealso></param>
    ''' <param name="vsoSelection">The Visio selection to search.</param>
    ''' <returns>
    ''' The number of replacements performed.
    ''' </returns>
    <Extension()> _
    Public Function ReplaceTextInShapeData(ByVal vsoSelection As Visio.Selection, _
                                           ByVal findText As String, _
                                           ByVal replaceText As String, _
                                           ByVal searchAndReplaceArgs As Short) _
                                           As Integer
        Dim vsoCell As Visio.Cell
        Dim counter As Integer = 0
        Dim newFormula As String = ""
        'Process Bitwise Flags
        Dim matchCase As Boolean = Not ((searchAndReplaceArgs And VisioPowerDevTools.visSearchAndReplaceArgs.MatchCase) = 0)
        Dim wholeWordsOnly As Boolean = Not ((searchAndReplaceArgs And VisioPowerDevTools.visSearchAndReplaceArgs.WholeWordsOnly) = 0)
        Dim regExSearch As Boolean = Not ((searchAndReplaceArgs And VisioPowerDevTools.visSearchAndReplaceArgs.RegEx) = 0)

        'RegExSearch cannot be used in combination with matchCase and wholeWords Only. 
        If regExSearch Then
            For Each vsoShape As Visio.Shape In vsoSelection
                For i = 0 To vsoShape.RowCount(Visio.VisSectionIndices.visSectionProp) - 1
                    vsoCell = vsoShape.CellsSRC(Visio.VisSectionIndices.visSectionProp, i, Visio.VisCellIndices.visCustPropsValue)
                    'Treat the findText as a regEx pattern. 
                    newFormula = Regex.Replace(vsoCell.FormulaU, findText, replaceText)
                    If Not String.Equals(vsoCell.FormulaU, newFormula) Then
                        vsoCell.FormulaU = newFormula
                        counter += 1
                    End If
                Next
            Next
            Return counter
        End If

        'Uses different search methods depending on what is needed. 
        'E.g., don't use RegEx if not required. 
        If matchCase AndAlso wholeWordsOnly Then
            For Each vsoShape As Visio.Shape In vsoSelection
                For i = 0 To vsoShape.RowCount(Visio.VisSectionIndices.visSectionProp) - 1
                    vsoCell = vsoShape.CellsSRC(Visio.VisSectionIndices.visSectionProp, i, Visio.VisCellIndices.visCustPropsValue)
                    '\bword\b performs a "whole words only" search using a regular expression. 
                    newFormula = Regex.Replace(vsoCell.FormulaU, "\b" & Regex.Escape(findText) & "\b", replaceText)
                    If Not String.Equals(vsoCell.FormulaU, newFormula) Then
                        vsoCell.FormulaU = newFormula
                        counter += 1
                    End If
                Next
            Next
        ElseIf matchCase Then
            For Each vsoShape As Visio.Shape In vsoSelection
                For i = 0 To vsoShape.RowCount(Visio.VisSectionIndices.visSectionProp) - 1
                    vsoCell = vsoShape.CellsSRC(Visio.VisSectionIndices.visSectionProp, i, Visio.VisCellIndices.visCustPropsValue)
                    newFormula = vsoCell.FormulaU.Replace(findText, replaceText)
                    If Not String.Equals(vsoCell.FormulaU, newFormula) Then
                        vsoCell.FormulaU = newFormula
                        counter += 1
                    End If
                Next
            Next
        ElseIf wholeWordsOnly Then
            For Each vsoShape As Visio.Shape In vsoSelection
                For i = 0 To vsoShape.RowCount(Visio.VisSectionIndices.visSectionProp) - 1
                    vsoCell = vsoShape.CellsSRC(Visio.VisSectionIndices.visSectionProp, i, Visio.VisCellIndices.visCustPropsValue)
                    '\bword\b performs a "whole words only" search using a regular expression. 
                    newFormula = Regex.Replace(vsoCell.FormulaU, "\b" & Regex.Escape(findText) & "\b", replaceText, RegexOptions.IgnoreCase)
                    If Not String.Equals(vsoCell.FormulaU, newFormula) Then
                        vsoCell.FormulaU = newFormula
                        counter += 1
                    End If
                Next
            Next
        Else
            For Each vsoShape As Visio.Shape In vsoSelection
                For i = 0 To vsoShape.RowCount(Visio.VisSectionIndices.visSectionProp) - 1
                    vsoCell = vsoShape.CellsSRC(Visio.VisSectionIndices.visSectionProp, i, Visio.VisCellIndices.visCustPropsValue)
                    newFormula = ReplaceCaseInsensitive(vsoCell.FormulaU, findText, replaceText)
                    If Not String.Equals(vsoCell.FormulaU, newFormula) Then
                        vsoCell.FormulaU = newFormula
                        counter += 1
                    End If
                Next
            Next
        End If

        Return counter

    End Function

    ''' <summary>
    ''' Replaces one string with another string in specified fields in the ShapeData section all the shapes in a selection.
    ''' </summary>
    ''' <param name="findText">The old text to replace..</param>
    ''' <param name="replaceText">The new text to add.</param>
    ''' <param name="searchAndReplaceArgs">Arguments used to specify how to search. <seealso cref="visSearchAndReplaceArgs">visSearchAndReplaceArgs</seealso></param>
    ''' <param name="vsoSelection">The Visio selection to search.</param>
    ''' <param name="fieldsToSearch">An array of fully qualified cell names representing the ShapeData fields to search.</param>
    ''' <returns>
    ''' The number of replacements performed.
    ''' </returns>
    <Extension()> _
    Public Function ReplaceTextInShapeData(ByVal vsoSelection As Visio.Selection, _
                                           ByVal findText As String, _
                                           ByVal replaceText As String, _
                                           ByVal searchAndReplaceArgs As Short, _
                                           ByVal fieldsToSearch As String()) _
                                           As Integer

        Dim vsoCell As Visio.Cell
        Dim counter As Integer = 0
        Dim newFormula As String = ""
        'Process Bitwise Flags
        Dim matchCase As Boolean = Not ((searchAndReplaceArgs And VisioPowerDevTools.visSearchAndReplaceArgs.MatchCase) = 0)
        Dim wholeWordsOnly As Boolean = Not ((searchAndReplaceArgs And VisioPowerDevTools.visSearchAndReplaceArgs.WholeWordsOnly) = 0)
        Dim regExSearch As Boolean = Not ((searchAndReplaceArgs And VisioPowerDevTools.visSearchAndReplaceArgs.RegEx) = 0)

        'RegExSearch cannot be used in combination with matchCase and wholeWords Only. 
        If regExSearch Then
            For Each vsoShape As Visio.Shape In vsoSelection
                For i = 0 To fieldsToSearch.Count - 1
                    If vsoShape.CellExistsU(fieldsToSearch(i), Visio.VisExistsFlags.visExistsAnywhere) Then
                        vsoCell = vsoShape.CellsU(fieldsToSearch(i))
                        'Treat the findText as a regEx pattern. 
                        newFormula = Regex.Replace(vsoCell.FormulaU, findText, replaceText)
                        If Not String.Equals(vsoCell.FormulaU, newFormula) Then
                            vsoCell.FormulaU = newFormula
                            counter += 1
                        End If
                    End If
                Next
            Next
            Return counter
        End If
        'Uses different search methods depending on what is needed. 
        'E.g., don't use RegEx if not required. 
        If matchCase AndAlso wholeWordsOnly Then
            For Each vsoShape As Visio.Shape In vsoSelection
                For i = 0 To fieldsToSearch.Count - 1
                    If vsoShape.CellExistsU(fieldsToSearch(i), Visio.VisExistsFlags.visExistsAnywhere) Then
                        vsoCell = vsoShape.CellsU(fieldsToSearch(i))
                        '\bword\b performs a "whole words only" search using a regular expression. 
                        newFormula = Regex.Replace(vsoCell.FormulaU, "\b" & Regex.Escape(findText) & "\b", replaceText)
                        If Not String.Equals(vsoCell.FormulaU, newFormula) Then
                            vsoCell.FormulaU = newFormula
                            counter += 1
                        End If
                    End If
                Next
            Next
        ElseIf matchCase Then
            For Each vsoShape As Visio.Shape In vsoSelection
                For i = 0 To fieldsToSearch.Count - 1
                    If vsoShape.CellExistsU(fieldsToSearch(i), Visio.VisExistsFlags.visExistsAnywhere) Then
                        vsoCell = vsoShape.CellsU(fieldsToSearch(i))
                        newFormula = vsoCell.FormulaU.Replace(findText, replaceText)
                        If Not String.Equals(vsoCell.FormulaU, newFormula) Then
                            vsoCell.FormulaU = newFormula
                            counter += 1
                        End If
                    End If
                Next
            Next
        ElseIf wholeWordsOnly Then
            For Each vsoShape As Visio.Shape In vsoSelection
                For i = 0 To fieldsToSearch.Count - 1
                    If vsoShape.CellExistsU(fieldsToSearch(i), Visio.VisExistsFlags.visExistsAnywhere) Then
                        vsoCell = vsoShape.CellsU(fieldsToSearch(i))
                        '\bword\b performs a "whole words only" search using a regular expression. 
                        newFormula = Regex.Replace(vsoCell.FormulaU, "\b" & Regex.Escape(findText) & "\b", replaceText, RegexOptions.IgnoreCase)
                        If Not String.Equals(vsoCell.FormulaU, newFormula) Then
                            vsoCell.FormulaU = newFormula
                            counter += 1
                        End If
                    End If
                Next
            Next
        Else
            For Each vsoShape As Visio.Shape In vsoSelection
                For i = 0 To fieldsToSearch.Count - 1
                    If vsoShape.CellExistsU(fieldsToSearch(i), Visio.VisExistsFlags.visExistsAnywhere) Then
                        vsoCell = vsoShape.CellsU(fieldsToSearch(i))
                        newFormula = ReplaceCaseInsensitive(vsoCell.FormulaU, findText, replaceText)
                        If Not String.Equals(vsoCell.FormulaU, newFormula) Then
                            vsoCell.FormulaU = newFormula
                            counter += 1
                        End If
                    End If
                Next
            Next
        End If

        Return counter
    End Function

#End Region

#Region "ClearFieldsInShapeData"

    ''' <summary>
    ''' Clears the formula of all the fields in the ShapeData section of all the shapes in a document.
    ''' </summary>
    ''' <param name="vsoDoc">The Visio document to search.</param>
    ''' <returns>
    ''' The number of fields cleared.
    ''' </returns>
    ''' <remarks>
    ''' Does not clear fields of shapes on background pages. 
    ''' </remarks>
    <Extension()> _
    Public Function ClearFieldsInShapeData(ByVal vsoDoc As Visio.Document) As Integer
        Dim vsoCell As Visio.Cell
        Dim counter As Integer = 0

        'Reset the formula to NoFormula for each shapeData field. 
        For Each vsoPage As Visio.Page In vsoDoc.Pages
            If False = CType(vsoPage.Background, Boolean) Then
                For Each vsoShape As Visio.Shape In vsoPage.Shapes
                    For i = 0 To vsoShape.RowCount(Visio.VisSectionIndices.visSectionProp) - 1
                        vsoCell = vsoShape.CellsSRC(Visio.VisSectionIndices.visSectionProp, i, Visio.VisCellIndices.visCustPropsValue)
                        vsoCell.FormulaU = VisioPowerDevTools.visFormulaNoFormula
                        counter += 1
                    Next
                Next
            End If
        Next
        Return counter
    End Function

    ''' <summary>
    ''' Clears the formula in specified fields in the ShapeData section all the shapes in a document.
    ''' </summary>
    ''' <param name="vsoDoc">The Visio document to search.</param>
    ''' <param name="fieldsToSearch">An array of fully qualified cell names representing the ShapeData fields to search.</param>
    ''' <returns>
    ''' The number of fields cleared.
    ''' </returns>
    <Extension()> _
    Public Function ClearFieldsInShapeData(ByVal vsoDoc As Visio.Document, _
                                           ByVal fieldsToSearch As String()) _
                                           As Integer

        Dim vsoCell As Visio.Cell
        Dim counter As Integer = 0

        'Reset the formula to NoFormula for each shapeData field. 
        For Each vsoPage As Visio.Page In vsoDoc.Pages
            If False = CType(vsoPage.Background, Boolean) Then
                For Each vsoShape As Visio.Shape In vsoPage.Shapes
                    For i = 0 To fieldsToSearch.Count - 1
                        If vsoShape.CellExistsU(fieldsToSearch(i), Visio.VisExistsFlags.visExistsAnywhere) Then
                            vsoCell = vsoShape.CellsU(fieldsToSearch(i))
                            vsoCell.FormulaU = VisioPowerDevTools.visFormulaNoFormula
                            counter += 1
                        End If
                    Next
                Next
            End If
        Next
        Return counter
    End Function

    ''' <summary>
    ''' Clears the formula of all the fields in the ShapeData section of all the shapes in a page.
    ''' </summary>
    ''' <param name="vsoPage">The Visio page to search.</param>
    ''' <returns>
    ''' The number of fields cleared.
    ''' </returns>
    <Extension()> _
    Public Function ClearFieldsInShapeData(ByVal vsoPage As Visio.Page) As Integer
        Dim vsoCell As Visio.Cell
        Dim counter As Integer = 0

        'Reset the formula to NoFormula for each shapeData field. 
        For Each vsoShape As Visio.Shape In vsoPage.Shapes
            For i = 0 To vsoShape.RowCount(Visio.VisSectionIndices.visSectionProp) - 1
                vsoCell = vsoShape.CellsSRC(Visio.VisSectionIndices.visSectionProp, i, Visio.VisCellIndices.visCustPropsValue)
                vsoCell.FormulaU = VisioPowerDevTools.visFormulaNoFormula
                counter += 1
            Next
        Next
        Return counter
    End Function

    ''' <summary>
    ''' Clears the formula in specified fields in the ShapeData section all the shapes in a page.
    ''' </summary>
    ''' <param name="vsoPage">The Visio page to search.</param>
    ''' <param name="fieldsToSearch">An array of fully qualified cell names representing the ShapeData fields to search.</param>
    ''' <returns>
    ''' The number of fields cleared.
    ''' </returns>
    <Extension()> _
    Public Function ClearFieldsInShapeData(ByVal vsoPage As Visio.Page, _
                                           ByVal fieldsToSearch As String()) _
                                           As Integer

        Dim vsoCell As Visio.Cell
        Dim counter As Integer = 0

        'Reset the formula to NoFormula for each shapeData field. 
        For Each vsoShape As Visio.Shape In vsoPage.Shapes
            For i = 0 To fieldsToSearch.Count - 1
                If vsoShape.CellExistsU(fieldsToSearch(i), Visio.VisExistsFlags.visExistsAnywhere) Then
                    vsoCell = vsoShape.CellsU(fieldsToSearch(i))
                    vsoCell.FormulaU = VisioPowerDevTools.visFormulaNoFormula
                    counter += 1
                End If
            Next
        Next
        Return counter
    End Function

    ''' <summary>
    ''' Clears the formula of all the fields in the ShapeData section of all the shapes in a selection.
    ''' </summary>
    ''' <param name="vsoSelection">The Visio selection to search.</param>
    ''' <returns>
    ''' The number of fields cleared.
    ''' </returns>
    <Extension()> _
    Public Function ClearFieldsInShapeData(ByVal vsoSelection As Visio.Selection) As Integer
        Dim vsoCell As Visio.Cell
        Dim counter As Integer = 0

        'Reset the formula to NoFormula for each shapeData field. 
        For Each vsoShape As Visio.Shape In vsoSelection
            For i = 0 To vsoShape.RowCount(Visio.VisSectionIndices.visSectionProp) - 1
                vsoCell = vsoShape.CellsSRC(Visio.VisSectionIndices.visSectionProp, i, Visio.VisCellIndices.visCustPropsValue)
                vsoCell.FormulaU = VisioPowerDevTools.visFormulaNoFormula
                counter += 1
            Next
        Next
        Return counter
    End Function

    ''' <summary>
    ''' Clears the formula in specified fields in the ShapeData section all the shapes in a selection.
    ''' </summary>
    ''' <param name="vsoSelection">The Visio selection to search.</param>
    ''' <param name="fieldsToSearch">An array of fully qualified cell names representing the ShapeData fields to search.</param>
    ''' <returns>
    ''' The number of fields cleared.
    ''' </returns>
    <Extension()> _
    Public Function ClearFieldsInShapeData(ByVal vsoSelection As Visio.Selection, _
                                           ByVal fieldsToSearch As String()) _
                                           As Integer

        Dim vsoCell As Visio.Cell
        Dim counter As Integer = 0

        'Reset the formula to NoFormula for each shapeData field. 
        For Each vsoShape As Visio.Shape In vsoSelection
            For i = 0 To fieldsToSearch.Count - 1
                If vsoShape.CellExistsU(fieldsToSearch(i), Visio.VisExistsFlags.visExistsAnywhere) Then
                    vsoCell = vsoShape.CellsU(fieldsToSearch(i))
                    vsoCell.FormulaU = VisioPowerDevTools.visFormulaNoFormula
                    counter += 1
                End If
            Next
        Next
        Return counter
    End Function

#End Region

#End Region

End Module

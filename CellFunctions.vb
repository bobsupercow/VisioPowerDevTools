Imports System.Text
Imports System.Runtime.CompilerServices

''' <summary>
''' Functions and Methods used to process data related to 
''' <see href="http://msdn.microsoft.com/en-us/library/ms368327%28v=office.12%29.aspx">Visio.Cell</see> objects.
''' </summary>
Public Module CellFunctions




    ''' <summary>
    ''' Copies the cell from one shape to another.
    ''' Optionally copies the entire row's contents.
    ''' </summary>
    ''' <param name="vsoSourceCell">The source cell.</param>
    ''' <param name="vsoDestShape">The destination shape.</param>
    ''' <param name="copyRow">Indicates whether or not to also copy the contents of the cell's parent row.</param>
    ''' <returns>
    ''' A reference to the newly copied cell. 
    ''' </returns>
    ''' <remarks>
    ''' Copies by Formula, not by Result. Maintains locale settings.
    ''' </remarks>
    ''' <exception cref="Exceptions.UnsupportedSectionException">
    ''' Thrown if <paramref name="vsoDestShape" >vsoDestShape</paramref> does not support the section to which <paramref name="vsoSourceCell">vsoSourceCell</paramref> belongs.
    ''' </exception>
    <Extension()> _
    Public Function CopyCellToShape(ByVal vsoSourceCell As Visio.Cell, _
                                    ByVal vsoDestShape As Visio.Shape, _
                                    ByVal copyRow As Boolean) As Visio.Cell
        Try
            'Get all needed variables.
            'Dim vsoDoc As Visio.Document = vsoSourceCell.Document
            Dim vsoSourceShape As Visio.Shape = vsoSourceCell.Shape
            Dim vsoSectionIndex As Visio.VisSectionIndices = vsoSourceCell.Section
            Dim vsoSourceRowIndex As Short = vsoSourceCell.Row
            Dim vsoRowTag As Visio.VisRowTags = vsoSourceCell.Shape.RowType(vsoSectionIndex, vsoSourceRowIndex)
            Dim vsoRowName As String = Nothing
            Dim vsoRowNameU As String = Nothing
            Dim vsoColumnIndex As Short = vsoSourceCell.Column
            Dim vsoCellName As String = vsoSourceCell.Name
            Dim vsoDestCell As Visio.Cell
            Dim vsoDestRowIndex As Short
            Dim vsoRowPrefix As String = VisioPowerDevTools.ShapeFunctions.GetSectionPrefix(vsoSectionIndex)


            'Make sure you're using a valid section based on shape-type
            If Not vsoDestShape.SectionIsValidToAdd(vsoSectionIndex) Then
                'Section is not supported...Throw appropriate exception.
                Dim ex As New UnsupportedSectionException( _
                    "You cannot create section " & vsoSectionIndex.ToString & _
                    " in " & vsoDestShape.Name & _
                    " because it is type: " & vsoDestShape.Type)
                Throw ex
            End If


            'Create section if it doesn't already exist. 
            If Not vsoDestShape.SectionExists(vsoSectionIndex, Visio.VisExistsFlags.visExistsAnywhere) Then
                vsoDestShape.AddSection(vsoSectionIndex)
            End If


            'Since we are moving to a shape, not a section, we already know that the type of row and cell are valid for the section.
            'Find out if the row is named or not.
            Dim sectionType As VisioPowerDevTools.visSectionTypes = VisioPowerDevTools.ShapeFunctions.GetSectionType(vsoSectionIndex)
            If _
                sectionType = visSectionTypes.NamedNonConstantRowsAndCells OrElse _
                sectionType = visSectionTypes.NamedOrUnnamedNonConstantRowsAndCells OrElse _
                sectionType = visSectionTypes.NamedRowsConstantCells Then
                'Row is Named, set variables.
                vsoRowName = vsoSourceCell.RowName
                vsoRowNameU = vsoSourceCell.RowNameU
                'Create named row if it doesn't already exist.
                If Not vsoDestShape.CellExistsU(vsoRowPrefix & vsoRowNameU, Visio.VisExistsFlags.visExistsAnywhere) Then
                    vsoDestRowIndex = vsoDestShape.AddNamedRow(vsoSectionIndex, vsoRowNameU, vsoRowTag)
                    'Keep any changes to locale name intact by setting "Name" in addition to "NameU"
                    vsoDestShape.Section(vsoSectionIndex).Row(vsoDestRowIndex).Name = vsoRowName
                Else
                    vsoDestRowIndex = vsoDestShape.Cells(vsoRowPrefix & vsoRowName).Row
                End If
            Else
                'Row is not named, just index.
                'Create row if it doesn't already exist.
                If Not vsoDestShape.RowExists(vsoSectionIndex, vsoSourceRowIndex, Visio.VisExistsFlags.visExistsAnywhere) Then
                    vsoDestRowIndex = vsoDestShape.AddRow(vsoSectionIndex, vsoSourceRowIndex, vsoRowTag)
                Else
                    vsoDestRowIndex = vsoSourceRowIndex
                End If
            End If

            'Test if the user wants to copy all of the row information, or just the cell info.
            If copyRow Then
                'Copy all of the cell formulas in the row.
                'Replace references to the source shape with references to the dest shape. 
                For i = 0 To vsoSourceCell.Shape.RowsCellCount(vsoSectionIndex, vsoSourceRowIndex) - 1
                    vsoDestShape.CellsSRC(vsoSectionIndex, vsoDestRowIndex, i).FormulaForceU = _
                        vsoSourceShape.CellsSRC(vsoSectionIndex, vsoSourceRowIndex, i).FormulaU.Replace(vsoSourceShape.Name, vsoDestShape.Name)
                    'Maintain locale formulas
                    vsoDestShape.CellsSRC(vsoSectionIndex, vsoDestRowIndex, i).FormulaForce = _
                        vsoSourceShape.CellsSRC(vsoSectionIndex, vsoSourceRowIndex, i).Formula.Replace(vsoSourceShape.Name, vsoDestShape.Name)
                Next
            Else
                vsoDestShape.CellsSRC(vsoSectionIndex, vsoDestRowIndex, vsoColumnIndex).FormulaForceU = _
                    vsoSourceCell.FormulaU
                'Maintain locale formulas
                vsoDestShape.CellsSRC(vsoSectionIndex, vsoDestRowIndex, vsoColumnIndex).FormulaForce = _
                    vsoSourceCell.Formula
            End If

            vsoDestCell = vsoDestShape.CellsSRC(vsoSectionIndex, vsoDestRowIndex, vsoColumnIndex)
            Return vsoDestCell

        Catch ex As Exception
            'Rethrow any other exceptions.
            Throw
        End Try
    End Function

#Region "Sort List"

    ''' <overloads>
    ''' Alphabetizes a list in a <see href="http://msdn.microsoft.com/en-us/library/ms368327%28v=office.12%29.aspx">Visio.Cell</see> (<paramref name="vsoCell">vsoCell</paramref>), 
    ''' and properly resets any cell formula that is referencing the list by indexing it.
    ''' </overloads>
    ''' <summary>
    ''' Alphabetizes a list in a <see href="http://msdn.microsoft.com/en-us/library/ms368327%28v=office.12%29.aspx">Visio.Cell</see> (<paramref name="vsoCell">vsoCell</paramref>), 
    ''' and properly resets any cell formula that is referencing the list by indexing it.
    ''' </summary>
    ''' <param name="vsoCell">The <see href="http://msdn.microsoft.com/en-us/library/ms368327%28v=office.12%29.aspx">Visio.Cell</see> containing the list.</param>
    ''' <param name="delimiter">The delimiting character of the list.</param>
    ''' <param name="alwaysOnTop">A string which should always be placed on top of the list.</param>
    ''' <param name="sortListArgs">Arguments used to determine various sort options. <seealso cref="SortListArgs">SortListArgs</seealso>.</param>
    ''' <remarks>
    ''' Assumes that it has been passed a valid <paramref name="vsoCell">cell</paramref> containing a delimited list.
    ''' If <paramref name="alwaysOnTop">alwaysOnTop</paramref> contains a valid string, any duplicates of that string will be replaced, 
    ''' regardless of the value of <paramref name="removeDuplicates">removeDuplicates</paramref>.
    ''' </remarks>
    Public Function SortList(ByVal vsoCell As Visio.Cell, _
                             ByVal delimiter As Char, _
                             ByVal alwaysOnTop As String, _
                             ByVal sortListArgs As Integer) As Boolean

        Try
            Dim reverseAlpha As Boolean = Not ((sortListArgs And CellFunctions.SortListArgs.reverseAlpha) = 0)
            Dim removeDuplicates As Boolean = Not ((sortListArgs And CellFunctions.SortListArgs.removeDuplicates) = 0)
            Dim removeUnused As Boolean = Not ((sortListArgs And CellFunctions.SortListArgs.removeUnused) = 0)


            'Get the raw list as a string so we can parse it.
            Dim listAsString As String = VisioFormulaToString(vsoCell.Formula)

            'Get the array of cells dependent on the list cell of the shape. 
            Dim cellsWhichIndexTheList As List(Of Visio.Cell) = GetIndexedCellReferences(vsoCell)
            'An array which will hold the index which the referencing-cell is using. 
            Dim indiciesUsedByReferences(cellsWhichIndexTheList.Count) As Integer
            'The dependent Cell
            Dim vsoDependentCell As Visio.Cell
            'The extracted Index from a shape represented as a string. 
            Dim myIntsAsString As String
            'The new index for the shape to use in the lookup cell.
            Dim newIndex As Integer
            'Regex which only returns numbers
            Dim numbersRegex As RegularExpressions.Regex = New RegularExpressions.Regex( _
                "[0-9][0-9]?        # All nums", _
                RegularExpressions.RegexOptions.IgnoreCase _
                Or RegularExpressions.RegexOptions.Multiline _
                Or RegularExpressions.RegexOptions.IgnorePatternWhitespace _
                Or RegularExpressions.RegexOptions.Compiled)
            'Loop through each of the dependent cells. 
            For i = 0 To cellsWhichIndexTheList.Count - 1
                'Get the cell.
                vsoDependentCell = cellsWhichIndexTheList(i)

                ' Loop through the match collection to retrieve all 
                ' matches and positions.
                Dim mc As RegularExpressions.MatchCollection = numbersRegex.Matches(vsoDependentCell.Formula)
                'Get the first match, ignore the rest.
                myIntsAsString = numbersRegex.Match(vsoDependentCell.Formula).Value

                indiciesUsedByReferences(i) = CInt(myIntsAsString)
            Next

            'Parse the list and get the individual items.
            Dim origSplitList As List(Of String) = listAsString.Split(";").ToList
            Dim keys As List(Of String)
            'If we are removing unused items, this must be done differently. 
            If removeUnused Then
                'Add an entry for each of the unique indicies referenced by other cells. 
                keys = New List(Of String)
                For Each i As Integer In indiciesUsedByReferences.Distinct
                    keys.Add(origSplitList(i))
                Next
            Else
                'If we aren't removing unused items, stick with the original.
                keys = origSplitList
            End If

            If removeDuplicates Then
                keys = keys.Distinct.ToList()
            End If

            ' Sort the keys using the specified method.
            If reverseAlpha = True Then
                keys.Sort(Function(a, b) a.CompareTo(b) * -1)
            Else
                keys.Sort(Function(a, b) a.CompareTo(b))
            End If
            'Move specified string to top of list if directed.
            If alwaysOnTop.Count > 0 Then
                If keys.Contains(alwaysOnTop) Then
                    keys.RemoveAll(Function(x) x.Equals(alwaysOnTop))
                    keys.Insert(0, alwaysOnTop)
                End If
            End If


            'A stringBuilder used to create the new list.
            Dim newListBuilder As StringBuilder = New StringBuilder

            ' Loop over the sorted keys, excluding the final key.
            For i = 0 To keys.Count - 2
                'Add the value to the new list.
                newListBuilder.Append(keys(i))
                'Add the seperator. 
                newListBuilder.Append(";", 1)
            Next
            'Add the final value to the list, but don't add a seperator.
            newListBuilder.Append(keys(keys.Count - 1))
            'Set the cell to use the new formula.
            vsoCell.Formula = StringToVisioFormula(newListBuilder.ToString)

            Dim oldIndex As Integer
            Dim formulaBuilder As New StringBuilder(50)
            'Loop through each of the dependent cells. 
            For i = 0 To cellsWhichIndexTheList.Count - 1
                'Get the cell.
                vsoDependentCell = cellsWhichIndexTheList(i)

                'Cast the string of integers to a useful index value.
                oldIndex = indiciesUsedByReferences(i)

                'Find the new index of the value.
                newIndex = keys.IndexOf(origSplitList(oldIndex))
                'Set the cell's formula, replacing our old index with a new one. 
                formulaBuilder.Append("INDEX(")
                formulaBuilder.Append(newIndex.ToString)
                formulaBuilder.Append(",")
                formulaBuilder.Append(vsoCell.Name)
                formulaBuilder.Append(")")
                vsoDependentCell.FormulaForceU = formulaBuilder.ToString

                'Clear the stringBuilder for reuse
                formulaBuilder.Length = 0
            Next

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

#Region "Overloaded SortList Functions"

    ''' <summary>
    ''' Overloads the <see cref="SortList">SortList</see> method for flexibility. 
    ''' </summary>
    ''' <param name="vsoCell">The <see href="http://msdn.microsoft.com/en-us/library/ms368327%28v=office.12%29.aspx">Visio.Cell</see> containing the list.</param>
    ''' <remarks>
    ''' Calling this method is equivalent to calling <see cref="SortList"/> and passing the following parameters.
    ''' <list type="bullet">
    ''' <item>
    ''' <term>vsoCell</term><description><c><paramref name="vsoCell"/></c></description>
    ''' </item>
    ''' <item>
    ''' <term>delimiter</term><description><c>;</c></description>
    ''' </item>
    ''' <item>
    ''' <term>alwaysOnTop</term><description><c>""</c></description>
    ''' </item>
    ''' <item>
    ''' <term>sortListArgs</term><description><c><see cref="SortListArgs.removeDuplicates">removeDuplicate</see> + <see cref="SortListArgs.removeUnused">removeUnused</see></c></description>
    ''' </item>
    ''' </list>
    ''' </remarks>
    Public Sub SortList(ByVal vsoCell As Visio.Cell)
        SortList(vsoCell, defaultDelimiter, defaultAlwaysOnTop, defaultArgs)
    End Sub

    ''' <summary>
    ''' Overloads the <see cref="SortList">SortList</see> method for flexibility. 
    ''' </summary>
    ''' <param name="vsoCell">The <see href="http://msdn.microsoft.com/en-us/library/ms368327%28v=office.12%29.aspx">Visio.Cell</see> containing the list.</param>
    ''' <param name="delimiter">The delimiting character of the list.</param>
    ''' <remarks>
    ''' Calling this method is equivalent to calling <see cref="SortList"/> and passing the following parameters.
    ''' <list type="bullet">
    ''' <item>
    ''' <term>vsoCell</term><description><c><paramref name="vsoCell"/></c></description>
    ''' </item>
    ''' <item>
    ''' <term>delimiter</term><description><c><paramref name="delimiter"/></c></description>
    ''' </item>
    ''' <item>
    ''' <term>alwaysOnTop</term><description><c>""</c></description>
    ''' </item>
    ''' <item>
    ''' <term>sortListArgs</term><description><c><see cref="SortListArgs.removeDuplicates">removeDuplicate</see> + <see cref="SortListArgs.removeUnused">removeUnused</see></c></description>
    ''' </item>
    ''' </list>
    ''' </remarks>
    Public Sub SortList(ByVal vsoCell As Visio.Cell, ByVal delimiter As Char)
        SortList(vsoCell, delimiter, defaultAlwaysOnTop, defaultArgs)
    End Sub
    ''' <summary>
    ''' Overloads the <see cref="SortList">SortList</see> method for flexibility. 
    ''' </summary>
    ''' <param name="vsoCell">The <see href="http://msdn.microsoft.com/en-us/library/ms368327%28v=office.12%29.aspx">Visio.Cell</see> containing the list.</param>
    ''' <param name="delimiter">The delimiting character of the list.</param>
    ''' <param name="alwaysOnTop">A string which should always be placed on top of the list.</param>
    ''' <remarks>
    ''' Calling this method is equivalent to calling <see cref="SortList"/> and passing the following parameters.
    ''' <list type="bullet">
    ''' <item>
    ''' <term>vsoCell</term><description><c><paramref name="vsoCell"/></c></description>
    ''' </item>
    ''' <item>
    ''' <term>delimiter</term><description><c><paramref name="delimiter"/></c></description>
    ''' </item>
    ''' <item>
    ''' <term>alwaysOnTop</term><description><c><paramref name="alwaysOnTop"/></c></description>
    ''' </item>
    ''' <item>
    ''' <term>sortListArgs</term><description><c><see cref="SortListArgs.removeDuplicates">removeDuplicate</see> + <see cref="SortListArgs.removeUnused">removeUnused</see></c></description>
    ''' </item>
    ''' </list>
    ''' </remarks>
    Public Sub SortList(ByVal vsoCell As Visio.Cell, ByVal delimiter As Char, ByVal alwaysOnTop As String)
        SortList(vsoCell, delimiter, alwaysOnTop, defaultArgs)
    End Sub
    ''' <summary>
    ''' Overloads the <see cref="SortList">SortList</see> method for flexibility. 
    ''' </summary>
    ''' <param name="vsoCell">The <see href="http://msdn.microsoft.com/en-us/library/ms368327%28v=office.12%29.aspx">Visio.Cell</see> containing the list.</param>
    ''' <param name="delimiter">The delimiting character of the list.</param>
    ''' <param name="sortListArgs">Arguments used to determine various sort options. <seealso cref="SortListArgs">SortListArgs</seealso>.</param>
    ''' <remarks>
    ''' Calling this method is equivalent to calling <see cref="SortList"/> and passing the following parameters.
    ''' <list type="bullet">
    ''' <item>
    ''' <term>vsoCell</term><description><c><paramref name="vsoCell"/></c></description>
    ''' </item>
    ''' <item>
    ''' <term>delimiter</term><description><c><paramref name="delimiter"/></c></description>
    ''' </item>
    ''' <item>
    ''' <term>alwaysOnTop</term><description><c>""</c></description>
    ''' </item>
    ''' <item>
    ''' <term>sortListArgs</term><description><c><see cref="SortListArgs.removeDuplicates">removeDuplicate</see> + <see cref="SortListArgs.removeUnused">removeUnused</see></c></description>
    ''' </item>
    ''' </list>
    ''' </remarks>
    Public Sub SortList(ByVal vsoCell As Visio.Cell, ByVal delimiter As Char, ByVal sortListArgs As Integer)
        SortList(vsoCell, delimiter, defaultAlwaysOnTop, sortListArgs)
    End Sub
    ''' <summary>
    ''' Overloads the <see cref="SortList">SortList</see> method for flexibility. 
    ''' </summary>
    ''' <param name="vsoCell">The <see href="http://msdn.microsoft.com/en-us/library/ms368327%28v=office.12%29.aspx">Visio.Cell</see> containing the list.</param>
    ''' <param name="alwaysOnTop">A string which should always be placed on top of the list.</param>
    ''' <remarks>
    ''' Calling this method is equivalent to calling <see cref="SortList"/> and passing the following parameters.
    ''' <list type="bullet">
    ''' <item>
    ''' <term>vsoCell</term><description><c><paramref name="vsoCell"/></c></description>
    ''' </item>
    ''' <item>
    ''' <term>delimiter</term><description><c>;</c></description>
    ''' </item>
    ''' <item>
    ''' <term>alwaysOnTop</term><description><c><paramref name="alwaysOnTop"/></c></description>
    ''' </item>
    ''' <item>
    ''' <term>sortListArgs</term><description><c><see cref="SortListArgs.removeDuplicates">removeDuplicate</see> + <see cref="SortListArgs.removeUnused">removeUnused</see></c></description>
    ''' </item>
    ''' </list>
    ''' </remarks>
    Public Sub SortList(ByVal vsoCell As Visio.Cell, ByVal alwaysOnTop As String)
        SortList(vsoCell, defaultDelimiter, alwaysOnTop, defaultArgs)
    End Sub
    ''' <summary>
    ''' Overloads the <see cref="SortList">SortList</see> method for flexibility. 
    ''' </summary>
    ''' <param name="vsoCell">The <see href="http://msdn.microsoft.com/en-us/library/ms368327%28v=office.12%29.aspx">Visio.Cell</see> containing the list.</param>
    ''' <param name="alwaysOnTop">A string which should always be placed on top of the list.</param>
    ''' <param name="sortListArgs">Arguments used to determine various sort options. <seealso cref="SortListArgs">SortListArgs</seealso>.</param>
    ''' <remarks>
    ''' Calling this method is equivalent to calling <see cref="SortList"/> and passing the following parameters.
    ''' <list type="bullet">
    ''' <item>
    ''' <term>vsoCell</term><description><c><paramref name="vsoCell"/></c></description>
    ''' </item>
    ''' <item>
    ''' <term>delimiter</term><description><c>;</c></description>
    ''' </item>
    ''' <item>
    ''' <term>alwaysOnTop</term><description><c><paramref name="alwaysOnTop"/></c></description>
    ''' </item>
    ''' <item>
    ''' <term>sortListArgs</term><description><c><paramref name="sortListArgs"/></c></description>
    ''' </item>
    ''' </list>
    ''' </remarks>
    Public Sub SortList(ByVal vsoCell As Visio.Cell, ByVal alwaysOnTop As String, ByVal sortListArgs As Integer)
        SortList(vsoCell, defaultDelimiter, alwaysOnTop, sortListArgs)
    End Sub
    ''' <summary>
    ''' Overloads the <see cref="SortList">SortList</see> method for flexibility. 
    ''' </summary>
    ''' <param name="vsoCell">The <see href="http://msdn.microsoft.com/en-us/library/ms368327%28v=office.12%29.aspx">Visio.Cell</see> containing the list.</param>
    ''' <param name="sortListArgs">Arguments used to determine various sort options. <seealso cref="SortListArgs">SortListArgs</seealso>.</param>
    ''' <remarks>
    ''' Calling this method is equivalent to calling <see cref="SortList"/> and passing the following parameters.
    ''' <list type="bullet">
    ''' <item>
    ''' <term>vsoCell</term><description><c><paramref name="vsoCell"/></c></description>
    ''' </item>
    ''' <item>
    ''' <term>delimiter</term><description><c>;</c></description>
    ''' </item>
    ''' <item>
    ''' <term>alwaysOnTop</term><description><c>""</c></description>
    ''' </item>
    ''' <item>
    ''' <term>sortListArgs</term><description><c><paramref name="sortListArgs"/></c></description>
    ''' </item>
    ''' </list>
    ''' </remarks>
    Public Sub SortList(ByVal vsoCell As Visio.Cell, ByVal sortListArgs As Integer)
        SortList(vsoCell, defaultDelimiter, defaultAlwaysOnTop, sortListArgs)
    End Sub
#End Region

#Region "Sort List Support"

    ''' <summary>
    ''' Bitwise arguments for the <see cref="SortList">SortList</see> method. 
    ''' </summary>
    Public Enum SortListArgs As Integer
        ''' <summary>
        ''' Sort in reverse alphabetical order.
        ''' </summary>
        reverseAlpha = 1
        ''' <summary>
        ''' Remove duplicate entries from the list.
        ''' </summary>
        removeDuplicates = 2
        ''' <summary>
        ''' Remove unused entries from the list. 
        ''' </summary>
        removeUnused = 4
    End Enum
    ''' <summary>
    ''' The default delimiter parameter used by <see cref="SortList">SortList</see>
    ''' </summary>
    Private Const defaultDelimiter As Char = ";"
    ''' <summary>
    ''' The default alwaysOnTop parameter used by <see cref="SortList">SortList</see>
    ''' </summary>
    Private Const defaultAlwaysOnTop As String = ""
    ''' <summary>
    ''' The default arguments by <see cref="SortList">SortList</see>
    ''' </summary>
    Private Const defaultArgs As Integer = SortListArgs.removeDuplicates + SortListArgs.removeUnused

    ''' <summary>
    ''' Gets all of the cells which directly index the <paramref name="vsoCell"></paramref>.
    ''' Designed for use when the <paramref name="vsoCell"></paramref> is a list. 
    ''' </summary>
    ''' <param name="vsoCell">The <see href="http://msdn.microsoft.com/en-us/library/ms368327%28v=office.12%29.aspx">Visio.Cell</see> to get references for.</param>
    ''' <returns>
    ''' A <see cref="List">List</see> containing all of the cells which directly index <paramref name="vsoCell"/>.
    ''' </returns>
    Public Function GetIndexedCellReferences(ByVal vsoCell As Visio.Cell) As List(Of Visio.Cell)
        Dim returnList As New List(Of Visio.Cell)

        'Get the array of cells dependent on the list cell of the shape. 
        Try
            'Don't check for dependents if formulaU = No Formula, it will cause death and destruction.
            If Not vsoCell.FormulaU = "" Then
                For Each dependentCell As Visio.Cell In vsoCell.Dependents
                    If dependentCell.FormulaU.Contains("INDEX(") And dependentCell.FormulaU.Contains(vsoCell.Name) Then
                        returnList.Add(dependentCell)
                        'We don't need to get the dependencies for a cell which is indexing.
                        Continue For
                    Else
                        For Each nestedDependentCell As Visio.Cell In GetIndexedCellReferences(dependentCell)
                            returnList.Add(nestedDependentCell)
                        Next
                    End If
                Next
            End If
        Catch ex As System.Runtime.InteropServices.COMException
            'Unexpected End of File, occurs when vsoCell.Dependents is an empty array of pointers, since ForEach is trying to enumerate it.
            If ex.ErrorCode = -2032466967 Then
                Return returnList
            Else
                'IDK what it is, throw it back up.
                Throw
            End If

        End Try

        Return returnList
    End Function

    ''' <summary>
    ''' Formats the list for use in a Visio Cell used as either a fixed or variable list.
    ''' </summary>
    ''' <param name="list">The list.</param>
    ''' <param name="alwaysOnTop">A string which should always be on top of the list.</param>
    ''' <param name="reverseAlpha">if set to <c>true</c> [Sort in Reverse Alphabetical Order].</param>
    ''' <param name="removeDuplicates">if set to <c>true</c> [Remove Duplicate Entries].</param>
    ''' <returns></returns>
    Public Function FormatList(ByVal list As String, _
                               ByVal usedList As String, _
                               ByVal alwaysOnTop As String, _
                               ByVal reverseAlpha As Boolean, _
                               ByVal removeDuplicates As Boolean, _
                               ByVal removeUnused As Boolean) _
                               As String

        Try

            'Parse the list and get the individual items.
            Dim origSplitList As List(Of String) = list.Split(";").ToList
            Dim usedSplitList As List(Of String) = usedList.Split(";").ToList
            Dim keys As List(Of String)
            'If we are removing unused items, this must be done differently. 
            If removeUnused Then
                'Add an entry for each of the unique indicies referenced by other cells. 
                keys = New List(Of String)
                For i = 0 To usedSplitList.Count - 1
                    keys.Add(usedSplitList(i))
                Next
            Else
                'If we aren't removing unused items, stick with the original.
                keys = origSplitList
            End If

            If removeDuplicates Then
                keys = keys.Distinct.ToList()
            End If

            ' Sort the keys using the specified method.
            If reverseAlpha = True Then
                keys.Sort(Function(a, b) a.CompareTo(b) * -1)
            Else
                keys.Sort(Function(a, b) a.CompareTo(b))
            End If
            'Move specified string to top of list if directed.
            If alwaysOnTop.Count > 0 Then
                If keys.Contains(alwaysOnTop) Then
                    keys.RemoveAll(Function(x) x.Equals(alwaysOnTop))
                    keys.Insert(0, alwaysOnTop)
                End If
            End If


            'A stringBuilder used to create the new list.
            Dim newListBuilder As StringBuilder = New StringBuilder

            ' Loop over the sorted keys, excluding the final key.
            For i = 0 To keys.Count - 2
                'Add the value to the new list.
                newListBuilder.Append(keys(i))
                'Add the seperator. 
                newListBuilder.Append(";", 1)
            Next
            'Add the final value to the list, but don't add a seperator.
            newListBuilder.Append(keys(keys.Count - 1))

            Return newListBuilder.ToString

        Catch ex As Exception
            Return Nothing
        End Try
    End Function
#End Region

#End Region

End Module

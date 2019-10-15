Imports System.Text
Imports System.Runtime.CompilerServices

''' <summary>
''' Functions and Methods used to process data related to 
''' Visio.Document objects.
''' </summary>
Public Module DocumentFunctions

    ''' <summary>
    ''' Tests for orphaned shapes (Shapes without masters)
    ''' </summary>
    ''' <param name="vsoDoc">The Visio Document to Test.</param>
    ''' <returns>
    ''' A list which contains <c>n</c> elements where <c>n</c> = the number of pages in the document. 
    ''' Each element is an arroy of integers, containing the shape.ID values for each 2-D orphaned shape on the page.
    ''' The lists can be returned empty, in which case there are no orphaned shapes of that type.
    ''' </returns>
    <Extension()> _
    Public Function TestForOrphaned2DShapes(ByVal vsoDoc As Visio.Document) As List(Of Integer())
        Dim returnList As New List(Of Integer())
        Try
            For Each vsoPage As Visio.Page In vsoDoc.Pages
                returnList.Add(vsoPage.TestForOrphaned2DShapes)
            Next
            Return returnList
        Catch ex As Exception
            Return returnList
        End Try
    End Function

End Module

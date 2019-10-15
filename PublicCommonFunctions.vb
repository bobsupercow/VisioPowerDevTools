''' <summary>
''' Functions and Mehtods commonly used for Visio application development.
''' </summary>
''' <remarks></remarks>
Public Module PublicCommonFunctions
#Region "Formula Functions"

    ''' <summary>
    ''' Converts a string to a properly formatted string that can be directly used as a Visio formula by adding the 
    ''' appropriate amount of double quotes.
    ''' </summary>
    ''' <param name="str">The string to convert.</param>
    ''' <returns>A properly formatted string which can be directly used as a Visio formula.</returns>
    Public Function StringToVisioFormula(ByVal str As String) As String
        Return """" & str & """"
    End Function

    ''' <summary>
    ''' Converts a Visio formula to a normal string without extra double quotes.
    ''' </summary>
    ''' <param name="vsoFormula">The string to convert.</param>
    ''' <returns>The Visio formula without extra double quotes.</returns>
    Public Function VisioFormulaToString(ByVal vsoFormula As String) As String
        'Take of the first and last characters since they include extra quotations.
        Return vsoFormula.Substring(1, vsoFormula.Count - 2)
    End Function

    ''' <summary>
    ''' Returns a new string in which all occurrences of a specified string in the current instance are replaced with another specified string.
    ''' CaseInsensitive
    ''' </summary>
    ''' <param name="original">The original string.</param>
    ''' <param name="findText">The string to be replaced.</param>
    ''' <param name="replaceText">The string to replace all occurrences of <paramref name="findText">findText</paramref>.</param>
    ''' <returns>
    ''' A string that is equivalent to the current string except that all instances of oldValue are replaced with newValue.
    ''' </returns>
    ''' <remarks>
    ''' All credit goes to Huisheng Chen
    ''' Found here: http://www.codeproject.com/Articles/10890/Fastest-C-Case-Insenstive-String-Replace
    ''' Date: 03/15/12
    ''' Modified for VB: 03/16/12
    ''' Faster than all other methods in almost every case for CaseInsensitive replace.
    ''' </remarks>
    Public Function ReplaceCaseInsensitive(ByVal original As String, _
                                           ByVal findText As String, _
                                           ByVal replaceText As String) _
                                           As String

        'Used to keep track of how many characters have been filled in in the Char array.
        Dim count As Integer = 0

        Dim position0 As Integer = 0
        Dim position1 As Integer = 0

        'Case insensitive. 
        'This is the one place that can cause a bottleneck which can be slower than other methods in certain cases.
        'However, those cases are rare.
        Dim upperString As String = original.ToUpper()
        Dim upperPattern As String = findText.ToUpper()
        'Determine how much bigger the new string will be if the replacement occurs.
        Dim inc As Integer = (original.Length \ findText.Length) * (replaceText.Length - findText.Length)
        'Make a new Char Array long enough to store the new string if needed.
        Dim chars As Char() = New Char(original.Length + (Math.Max(0, inc) - 1)) {}
        'Determine if/where the first instance of findText is found in the original string.
        position1 = upperString.IndexOf(upperPattern, position0)
        'Since there may be more than one instance of the findText, we must check for it more than once.
        'If IndexOf(string, startingIndex) fails to find the string, it returns -1
        While (position1 <> -1)
            'Write everything from the last known non-matching index to the next instance of the findText to the Char array.
            For i = position0 To position1 - 1
                chars(count) = original(i)
                count += 1
            Next
            'Add the replacement text.
            For i = 0 To replaceText.Length - 1
                chars(count) = replaceText(i)
                count += 1
            Next
            'position0 = how much of the original string we have processed.
            position0 = position1 + findText.Length
            'position1 = -1 if there are no more matches or 'n' if another instance was found at index 'n'.
            position1 = upperString.IndexOf(upperPattern, position0)
        End While

        'If position0 was never reset, then no match was ever found. ;)
        If position0 = 0 Then
            Return original
        Else
            'Add the remaining unprocessed text from the original string to the char array.
            For i = position0 To original.Length - 1
                chars(count) = original(i)
                count += 1
            Next

            Return New String(chars, 0, count)
        End If
    End Function


#End Region
End Module

''' <summary>
''' Global Variables to support the Visio Power Tools Assembly.
''' </summary>
Public Module Globals
    ''' <summary>
    ''' Required since VB.Net doesn't handle integer overflow gracefully with Shorts
    ''' </summary>
    Public Const visEvtAdd As Short = -32768 '<-- this is the magic

    Public Const visFormulaNoFormula As String = ""
    Public Const visFormulaEmptyString As String = """"
End Module

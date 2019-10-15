Imports System
Imports System.IO
Imports System.Runtime.Serialization
Imports Microsoft.VisualBasic
Imports System.Security.Permissions

Module Exceptions

    ''' <summary>
    ''' Exception to throw when the user attempts to modify a visio object or section which is unsupported by the called method.
    ''' </summary>
    Public Class UnsupportedSectionException
        Inherits System.ApplicationException

        Private Const defaultMessage As String = ""

#Region "Public Constructors used by class instantiators."
        Public Sub New()
            MyBase.New(defaultMessage)
        End Sub ' New

        Public Sub New(ByVal auxMessage As String)
            MyBase.New(String.Format("{0} - {1}", _
                defaultMessage, auxMessage))
        End Sub ' New

        Public Sub New(ByVal inner As Exception)
            MyBase.New(defaultMessage, inner)
        End Sub ' New

        Public Sub New(ByVal auxMessage As String, ByVal inner As Exception)
            MyBase.New(String.Format("{0} - {1}", _
                defaultMessage, auxMessage), inner)
        End Sub ' New
#End Region

#Region "Public Constructors used for deserialization."
        Protected Sub New(ByVal info As SerializationInfo, _
                          ByVal context As StreamingContext)
            MyBase.New(info, context)
        End Sub ' New
#End Region

    End Class

    ''' <summary>
    ''' Exception to throw when the user attempts to modify a visio object or section which is unsupported by the called method.
    ''' </summary>
    Public Class CopyPageShapeException
        Inherits System.ApplicationException

        Private Const defaultMessage As String = "An error has occured while trying to copy the page"

#Region "Public Constructors used by class instantiators."
        Public Sub New()
            MyBase.New(defaultMessage)
        End Sub ' New

        Public Sub New(ByVal auxMessage As String)
            MyBase.New(String.Format("{0} - {1}", _
                defaultMessage, auxMessage))
        End Sub ' New

        Public Sub New(ByVal inner As Exception)
            MyBase.New(defaultMessage, inner)
        End Sub ' New

        Public Sub New(ByVal auxMessage As String, ByVal inner As Exception)
            MyBase.New(String.Format("{0} - {1}", _
                defaultMessage, auxMessage), inner)
        End Sub ' New
#End Region

#Region "Public Constructors used for deserialization."
        Protected Sub New(ByVal info As SerializationInfo, _
                          ByVal context As StreamingContext)
            MyBase.New(info, context)
        End Sub ' New
#End Region

    End Class

    ''' <summary>
    ''' Exception to throw when the user attempts to replace a master shape that...method.
    ''' </summary>
    Public Class ReplaceMasterException
        Inherits System.ApplicationException

        Private Const defaultMessage As String = "An error has occured while trying to replace a master shape from a stencil."

#Region "Public Constructors used by class instantiators."
        Public Sub New()
            MyBase.New(defaultMessage)
        End Sub ' New

        Public Sub New(ByVal auxMessage As String)
            MyBase.New(String.Format("{0} - {1}", _
                defaultMessage, auxMessage))
        End Sub ' New

        Public Sub New(ByVal inner As Exception)
            MyBase.New(defaultMessage, inner)
        End Sub ' New

        Public Sub New(ByVal auxMessage As String, ByVal inner As Exception)
            MyBase.New(String.Format("{0} - {1}", _
                defaultMessage, auxMessage), inner)
        End Sub ' New
#End Region

#Region "Public Constructors used for deserialization."
        Protected Sub New(ByVal info As SerializationInfo, _
                          ByVal context As StreamingContext)
            MyBase.New(info, context)
        End Sub ' New
#End Region
    End Class

End Module

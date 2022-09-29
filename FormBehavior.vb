Imports Microsoft.VisualBasic

Public Class FormBehavior
    Private _Caption As String
    Private _Code As String

    Public Property Code() As String
        Get
            Return _Code
        End Get
        Set(ByVal value As String)
            _Code = value
        End Set
    End Property

    Public Property Caption() As String
        Get
            Return _Caption
        End Get
        Set(ByVal value As String)
            _Caption = value
        End Set
    End Property
End Class

Imports Microsoft.VisualBasic

Public Class View
    Private _Name As String
    Private _Columns As Collection


    Public Property Name() As String
        Get
            Return _Name
        End Get
        Set(ByVal value As String)
            _Name = value
        End Set
    End Property

    Public Property Columns() As Collection
        Get
            Return _Columns
        End Get
        Set(ByVal value As Collection)
            _Columns = value
        End Set
    End Property
End Class

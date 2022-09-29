Imports Microsoft.VisualBasic

Public Class DatagridClass
    Private _Index As Integer
    Public Property Index() As Integer
        Get
            Return _Index
        End Get
        Set(ByVal value As Integer)
            _Index = value
        End Set
    End Property
End Class

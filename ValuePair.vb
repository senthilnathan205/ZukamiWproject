Imports Microsoft.VisualBasic

Public Class ValuePair
    Private _value As Object
    Private _caption As String
    Private _value2 As Object
    Private _value3 As Object
    Private _value4 As Object
    Private _value5 As Object
    Private _value6 As Object
    Private _value7 As Object
    Private _mainobject As Object

    Public Property Value6() As Object
        Get
            Return _value6
        End Get
        Set(ByVal value As Object)
            _value6 = value
        End Set
    End Property

    Public Property Value7() As Object
        Get
            Return _value7
        End Get
        Set(ByVal value As Object)
            _value7 = value
        End Set
    End Property

    Public Property MainObject() As Object
        Get
            Return _mainobject
        End Get
        Set(ByVal value As Object)
            _mainobject = value
        End Set
    End Property


    Public Property Value4() As Object
        Get
            Return _value4
        End Get
        Set(ByVal value As Object)
            _value4 = value
        End Set
    End Property

    Public Property Value5() As Object
        Get
            Return _value5
        End Get
        Set(ByVal value As Object)
            _value5 = value
        End Set
    End Property

    Public Property Value3() As Object
        Get
            Return _value3
        End Get
        Set(ByVal value As Object)
            _value3 = value
        End Set
    End Property

    Public Property Value2() As Object
        Get
            Return _value2
        End Get
        Set(ByVal value As Object)
            _value2 = value
        End Set
    End Property

    Public Property Value() As Object
        Get
            Return _value
        End Get
        Set(ByVal value As Object)
            _value = value
        End Set
    End Property

    Public Property Caption() As String
        Get
            Return _caption
        End Get
        Set(ByVal value As String)
            _caption = value
        End Set
    End Property

    Public Overrides Function ToString() As String
        Return _caption
    End Function



End Class

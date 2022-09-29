Imports Microsoft.VisualBasic

Public Class ButtonProperties
    Private _Name As String
    Private _Caption As String
    Private _JavascriptAttr As String
    Private _CSSClass As String
    Private _Hidden As Boolean
    Private _CustomButton As Boolean = False

    Public Property CustomButton() As Boolean
        Get
            Return _CustomButton
        End Get
        Set(ByVal value As Boolean)
            _CustomButton = value
        End Set
    End Property

    Public Property Hidden() As Boolean
        Get
            Return _Hidden
        End Get
        Set(ByVal value As Boolean)
            _Hidden = value
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

    Public Property JavascriptAttr() As String
        Get
            Return _JavascriptAttr
        End Get
        Set(ByVal value As String)
            _JavascriptAttr = value
        End Set
    End Property

    Public Property CSSClass() As String
        Get
            Return _CSSClass
        End Get
        Set(ByVal value As String)
            _CSSClass = value
        End Set
    End Property

    Public Property Name() As String
        Get
            Return _Name
        End Get
        Set(ByVal value As String)
            _Name = value
        End Set
    End Property
End Class

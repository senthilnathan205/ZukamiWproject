Imports Microsoft.VisualBasic

Public Class CustomImageLink
    Private _CustomImageLinkURL As String
    Private _customImageLinkToolTip As String
    Private _CustomImageLinkOnclickJScript As String
    Private _CustomImageLinkType As String
    Private _UseMapping As Boolean
    Private _ImageMap As Collection
    Private _MappedColumn As String
    Private _IDColumn As String
    Private _VisibleGUIDColumn As String

    Public Property VisibleGUIDColumn() As String
        Get
            Return _VisibleGUIDColumn
        End Get
        Set(ByVal value As String)
            _VisibleGUIDColumn = value
        End Set
    End Property


    Public Property MappedColumn() As String
        Get
            Return _MappedColumn
        End Get
        Set(ByVal value As String)
            _MappedColumn = value
        End Set
    End Property

    Public Property IDColumn() As String
        Get
            Return _IDColumn
        End Get
        Set(ByVal value As String)
            _IDColumn = value
        End Set
    End Property

    Public Property UseMapping() As Boolean
        Get
            Return _UseMapping
        End Get
        Set(ByVal value As Boolean)
            _UseMapping = value
        End Set
    End Property

    Public Property ImageMap() As Collection
        Get
            Return _ImageMap
        End Get
        Set(ByVal value As Collection)
            _ImageMap = value
        End Set
    End Property

    Public Property CustomImageLinkType() As String
        Get
            Return _CustomImageLinkType
        End Get
        Set(ByVal value As String)
            _CustomImageLinkType = value
        End Set
    End Property
    Public Property CustomImageLinkOnclickJScript() As String
        Get
            Return _CustomImageLinkOnclickJScript
        End Get
        Set(ByVal value As String)
            _CustomImageLinkOnclickJScript = value
        End Set
    End Property

    Public Property CustomImageLinkURL() As String
        Get
            Return _CustomImageLinkURL
        End Get
        Set(ByVal value As String)
            _CustomImageLinkURL = value
        End Set
    End Property

    Public Property CustomImageLinkToolTip() As String
        Get
            Return _customImageLinkToolTip
        End Get
        Set(ByVal value As String)
            _customImageLinkToolTip = value
        End Set
    End Property
End Class

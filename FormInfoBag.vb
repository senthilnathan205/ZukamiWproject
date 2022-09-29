Imports Microsoft.VisualBasic

Public Class FormInfoBag
    Private _appID As Guid
    Private _hfAppID As String
    Private _viewID As Guid
    Private _listName As String
    Private _customRowtemplate As String
    Private _aggregate As String
    Private _Listdescription As String
    Private _actions As String
    Private _isinlineview As Boolean
    Private _idfield As String
    Private _sourceType As Integer
    Private _hfsourceType As String
    Private _hfSourceDSN As String
    Private _hfSourceData As String
    Private _hfOrdering As String
    Private _hfGrouping As String
    Private _hfFieldList As String
    Private _hfTableName As String
    Private _hfMainTable As String
    Private _CountryColl As Collection
    Private _DGColumns As New Collection
    Private _biglookups As New Collection

    Public Property BigLookups() As Collection
        Get
            Return _biglookups
        End Get
        Set(value As Collection)
            _biglookups = value
        End Set
    End Property

    Public Property ViewID() As Guid
        Get
            Return _viewID
        End Get
        Set(value As Guid)
            _viewID = value
        End Set
    End Property

    Public Property DGColumns() As Collection
        Get
            Return _DGColumns
        End Get
        Set(value As Collection)
            _DGColumns = value
        End Set
    End Property

    Public Property CountryColl() As Collection
        Get
            Return _CountryColl
        End Get
        Set(value As Collection)
            _CountryColl = value
        End Set
    End Property


    Public Property hfMainTable() As String
        Get
            Return _hfMainTable
        End Get
        Set(value As String)
            _hfMainTable = value
        End Set
    End Property

    Public Property hfTableName() As String
        Get
            Return _hfTableName
        End Get
        Set(value As String)
            _hfTableName = value
        End Set
    End Property

    Public Property hfFieldList() As String
        Get
            Return _hfFieldList
        End Get
        Set(value As String)
            _hfFieldList = value
        End Set
    End Property

    Public Property hfGrouping() As String
        Get
            Return _hfGrouping
        End Get
        Set(value As String)
            _hfGrouping = value
        End Set
    End Property

    Public Property hfOrdering() As String
        Get
            Return _hfOrdering
        End Get
        Set(value As String)
            _hfOrdering = value
        End Set
    End Property

    Public Property hfSourceData() As String
        Get
            Return _hfSourceData
        End Get
        Set(value As String)
            _hfSourceData = value
        End Set
    End Property

    Public Property hfSourceDSN() As String
        Get
            Return _hfSourceDSN
        End Get
        Set(value As String)
            _hfSourceDSN = value
        End Set
    End Property

    Public Property hfSourceType() As String
        Get
            Return _hfsourceType
        End Get
        Set(value As String)
            _hfsourceType = value
        End Set
    End Property

    Public Property SourceType() As Integer
        Get
            Return _sourceType
        End Get
        Set(value As Integer)
            _sourceType = value
        End Set
    End Property

    Public Property IDField() As String
        Get
            Return _idfield
        End Get
        Set(value As String)
            _idfield = value
        End Set
    End Property

    Public Property IsInlineView() As Boolean
        Get
            Return _isinlineview
        End Get
        Set(value As Boolean)
            _isinlineview = value
        End Set
    End Property

    Public Property Actions() As String
        Get
            Return _actions
        End Get
        Set(value As String)
            _actions = value
        End Set
    End Property

    Public Property ListDescription() As String
        Get
            Return _Listdescription
        End Get
        Set(value As String)
            _Listdescription = value
        End Set
    End Property

    Public Property Aggregate() As String
        Get
            Return _aggregate
        End Get
        Set(value As String)
            _aggregate = value
        End Set
    End Property

    Public Property CustomRowTemplate() As String
        Get
            Return _customRowtemplate
        End Get
        Set(value As String)
            _customRowtemplate = value
        End Set
    End Property

    Public Property ListName() As String
        Get
            Return _listName
        End Get
        Set(value As String)
            _listName = value
        End Set
    End Property

    Public Property AppID() As Guid
        Get
            Return _appID
        End Get
        Set(value As Guid)
            _appID = value
        End Set
    End Property

    Public Property hfAppID() As String
        Get
            Return _hfAppID
        End Get
        Set(value As String)
            _hfAppID = value
        End Set
    End Property

End Class

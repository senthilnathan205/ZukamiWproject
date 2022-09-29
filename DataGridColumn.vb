Imports Microsoft.VisualBasic

Public Class DataGridColumn
    Private _BoundColumn As String
    Private _BoundColumnFormula As String
    Private _BoundDataColumnName As String
    Private _ColumnCaption As String
    Private _Width As Integer
    Private _FType As GlobalFunctions.FIELDTYPES
    Private _Lookup As Collection = Nothing
    Private _arguments As String
    Private _showfilter As Boolean
    Private _formatting As String
    Private _formatting2 As String = ""
    Private _fldAttributes As String
    Private _Attribute As Object = Nothing
    Private _popupEdit As Boolean = False
    Private _UpdateSQL As String = ""
    'Field attributes for 


    Public Class TIFFAttributes
        Private _Resize As Boolean = False
        Private _DisplayAsLink As Boolean = False
        Private _Width As String = ""
        Private _Height As String = ""
        Private _Javascript As String = ""

        Public Property Javascript() As String
            Get
                Return _Javascript
            End Get
            Set(value As String)
                _Javascript = value
            End Set
        End Property

        Public Property Resize() As Boolean
            Get
                Return _Resize
            End Get
            Set(ByVal value As Boolean)
                _Resize = value
            End Set
        End Property
        Public Property DisplayAsLink() As Boolean
            Get
                Return _DisplayAsLink
            End Get
            Set(ByVal value As Boolean)
                _DisplayAsLink = value
            End Set
        End Property

        Public Property Width() As String
            Get
                Return _Width
            End Get
            Set(ByVal value As String)
                _Width = value
            End Set
        End Property

        Public Property Height() As String
            Get
                Return _Height
            End Get
            Set(ByVal value As String)
                _Height = value
            End Set
        End Property
    End Class

    Public Property BoundColumnFormula() As String
        Get
            Return _BoundColumnFormula
        End Get
        Set(ByVal value As String)
            _BoundColumnFormula = value
        End Set
    End Property

    Public ReadOnly Property Attribute() As Object
        Get
            Return _Attribute
        End Get
    End Property

    Public Property Formatting2() As String
        Get
            Return _formatting2
        End Get
        Set(ByVal value As String)
            _formatting2 = value
        End Set
    End Property

    Public Property Formatting() As String
        Get
            Return _formatting
        End Get
        Set(ByVal value As String)
            _formatting = value
        End Set
    End Property

    Public Property PopupEdit() As Boolean
        Get
            Return _popupEdit
        End Get
        Set(ByVal value As Boolean)
            _popupEdit = value
        End Set
    End Property

    Public Property FldAttributes() As String
        Get
            Return _fldAttributes
        End Get
        Set(ByVal value As String)
            _fldAttributes = value
            If Len(_fldAttributes) > 0 Then
                Select Case _FType
                    Case GlobalFunctions.FIELDTYPES.FT_TIFFVIEWER, GlobalFunctions.FIELDTYPES.FT_CAMERA, GlobalFunctions.FIELDTYPES.FT_SIGNATURE
                        Dim _attr As New TIFFAttributes
                        GlobalFunctions.LoadTIFFViewerColAttr(_fldAttributes, _attr.DisplayAsLink, _attr.Resize, _attr.Width, _attr.Height, _attr.javascript)
                        _Attribute = _attr
                End Select
            End If
        End Set
    End Property

    Public Property BoundDataColumnName() As String
        Get
            Return _BoundDataColumnName
        End Get
        Set(ByVal value As String)
            _BoundDataColumnName = value
        End Set
    End Property

    Public Property ShowFilter() As Boolean
        Get
            Return _showfilter
        End Get
        Set(ByVal value As Boolean)
            _showfilter = value
        End Set
    End Property

    Public Property Arguments() As String
        Get
            Return _arguments
        End Get
        Set(ByVal value As String)
            _arguments = value
        End Set
    End Property

    Public Property UpdateSQL() As String
        Get
            Return _UpdateSQL
        End Get
        Set(ByVal value As String)
            _UpdateSQL = value
        End Set
    End Property

    Public Property FType() As GlobalFunctions.FIELDTYPES
        Get
            Return _FType
        End Get
        Set(ByVal value As GlobalFunctions.FIELDTYPES)
            _FType = value
        End Set
    End Property

    Public Property Lookup() As Collection
        Get
            Return _Lookup
        End Get
        Set(ByVal value As Collection)
            _Lookup = value
        End Set
    End Property

    Public Property Width() As Integer
        Get
            Return _Width
        End Get
        Set(ByVal value As Integer)
            _Width = value
        End Set
    End Property

    Public Property ColumnCaption() As String
        Get
            Return _ColumnCaption
        End Get
        Set(ByVal value As String)
            _ColumnCaption = value
        End Set
    End Property

    Public Property BoundColumn() As String
        Get
            Return _BoundColumn
        End Get
        Set(ByVal value As String)
            _BoundColumn = value
        End Set
    End Property
End Class

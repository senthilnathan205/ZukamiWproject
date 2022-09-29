Imports Microsoft.VisualBasic

Public Class ZField
    Private _BoundField As String
    Private _FieldName As String
    Private _FieldCaption As String
    Private _FieldControl As Object
    Private _FieldControl2 As Object
    Private _FieldControl3 As Object
    Private _FieldControl4 As Object
    Private _FieldLabelControl As Object
    Private _FieldLabelDescription As Object = Nothing
    Private _IsHighlighted As Boolean = False
    Private _IsCompulsory As Boolean
    Private _AllowDuplicates As Boolean
    Private _Arguments As String
    Private _FieldGUID As Guid
    Private _UpdateDB As Boolean
    Private _DoPostback As Boolean
    Private _FieldType As GlobalFunctions.FIELDTYPES
    Private _DefaultValue As String
    Private _InputMask As String
    Private _maxChars As Integer
    Private _TopLimit As String
    Private _BottomLimit As String
    Private _JavascriptAttr As String
    Private _Attributes As Collection
    Private _ExcludeFromInsert As Boolean
    Private _ExcludeFromUpdate As Boolean
    Private _Enabled As Boolean
    Private _ReadOnlyControl As Object
    Private _CSS As String
    Private _LabelCSS As String
    Private _FocusCSS As String
    Private _ReadonlyCSS As String
    Private _FormColumn As Integer
    Private _FieldSearchTarget As String
    Private _ExternalBindSource As String
    Private _IsHidden As Boolean
    Private _Width As Integer
    Private _ButtonList As String
    Private _ParentMobileLabel As TableRow = Nothing
    Private _ParentRow As Object = Nothing
    Private _EditModeControlsContainer As Object = Nothing

    Public Property EditModeControlsContainer As Object
        Get
            Return _EditModeControlsContainer
        End Get
        Set(value As Object)
            _EditModeControlsContainer = value
        End Set
    End Property
    Public Property ParentRow As Object
        Get
            Return _ParentRow
        End Get
        Set(value As Object)
            _ParentRow = value
        End Set
    End Property

    Public Property ParentMobileLabel As TableRow
        Get
            Return _ParentMobileLabel
        End Get
        Set(value As TableRow)
            _ParentMobileLabel = value
        End Set
    End Property

    Public Property ButtonList() As String
        Get
            Return _ButtonList
        End Get
        Set(value As String)
            _ButtonList = value
        End Set
    End Property

    Public Property Width() As Integer
        Get
            Return _Width
        End Get
        Set(value As Integer)
            _Width = value
        End Set
    End Property

    Public Property IsHidden() As Boolean
        Get
            Return _IsHidden
        End Get
        Set(value As Boolean)
            _IsHidden = value
        End Set
    End Property

    Public Property ExternalBindSource() As String
        Get
            Return _ExternalBindSource
        End Get
        Set(value As String)
            _ExternalBindSource = value
        End Set
    End Property

    Public Property FieldSearchTarget() As String
        Get
            Return _FieldSearchTarget
        End Get
        Set(ByVal value As String)
            _FieldSearchTarget = value
        End Set
    End Property

    Public Property FormColumn() As Integer
        Get
            Return _FormColumn
        End Get
        Set(ByVal value As Integer)
            _FormColumn = value
        End Set
    End Property

    Public Property FieldLabelDescription() As Object
        Get
            Return _FieldLabelDescription
        End Get
        Set(ByVal value As Object)
            _FieldLabelDescription = value
        End Set
    End Property

    Public Property IsHighlighted() As Boolean
        Get
            Return _IsHighlighted
        End Get
        Set(ByVal value As Boolean)
            _IsHighlighted = value
        End Set
    End Property

    Public Property FocusCSS() As String
        Get
            Return _FocusCSS
        End Get
        Set(ByVal value As String)
            _FocusCSS = value
        End Set
    End Property


    Public Property ReadonlyCSS() As String
        Get
            Return _ReadonlyCSS
        End Get
        Set(ByVal value As String)
            _ReadonlyCSS = value
        End Set
    End Property

    Public Property FieldLabelControl() As Object
        Get
            Return _FieldLabelControl
        End Get
        Set(ByVal value As Object)
            _FieldLabelControl = value
        End Set
    End Property

    Public Property CSS() As String
        Get
            Return _CSS
        End Get
        Set(ByVal value As String)
            _CSS = value
        End Set
    End Property

    Public Property LabelCSS() As String
        Get
            Return _LabelCSS
        End Get
        Set(ByVal value As String)
            _LabelCSS = value
        End Set
    End Property

    Public Property ReadOnlyControl() As Object
        Get
            Return _ReadOnlyControl
        End Get
        Set(ByVal value As Object)
            _ReadOnlyControl = value
        End Set
    End Property

    Public Property Enabled() As Boolean
        Get
            Return _Enabled
        End Get
        Set(ByVal value As Boolean)
            _Enabled = value
        End Set
    End Property

    Public Property ExcludeFromInsert() As Boolean
        Get
            Return _ExcludeFromInsert
        End Get
        Set(ByVal value As Boolean)
            _ExcludeFromInsert = value
        End Set
    End Property

    Public Property ExcludeFromUpdate() As Boolean
        Get
            Return _ExcludeFromUpdate
        End Get
        Set(ByVal value As Boolean)
            _ExcludeFromUpdate = value
        End Set
    End Property

    Public Property AllowDuplicates() As Boolean
        Get
            Return _AllowDuplicates
        End Get
        Set(ByVal value As Boolean)
            _AllowDuplicates = value
        End Set
    End Property

    Public Property JavascriptAttr() As String
        Get
            Return _JavascriptAttr
        End Get
        Set(ByVal value As String)
            _JavascriptAttr = value
            If Len(_JavascriptAttr) > 0 Then
                _Attributes = GlobalFunctions.LoadJavascript(_JavascriptAttr)
            End If
        End Set
    End Property

    Public ReadOnly Property Attributes() As Collection
        Get
            Return _Attributes
        End Get
    End Property

    Public Property FieldName() As String
        Get
            Return _FieldName
        End Get
        Set(ByVal value As String)
            _FieldName = value
        End Set
    End Property

    Public Property TopLimit() As String
        Get
            Return _TopLimit
        End Get
        Set(ByVal value As String)
            _TopLimit = value
        End Set
    End Property

    Public Property BottomLimit() As String
        Get
            Return _BottomLimit
        End Get
        Set(ByVal value As String)
            _BottomLimit = value
        End Set
    End Property

    Public Property MaxChars() As Integer
        Get
            Return _maxChars
        End Get
        Set(ByVal value As Integer)
            _maxChars = value
        End Set
    End Property

    Public Property InputMask() As String
        Get
            Return _InputMask
        End Get
        Set(ByVal value As String)
            _InputMask = value
        End Set
    End Property

    Public Property DefaultValue() As String
        Get
            Return _DefaultValue
        End Get
        Set(ByVal value As String)
            _DefaultValue = value
        End Set
    End Property

    Public Property DoPostback() As Boolean
        Get
            Return _DoPostback
        End Get
        Set(ByVal value As Boolean)
            _DoPostback = value
        End Set
    End Property

    Public Property Arguments() As String
        Get
            Return _Arguments
        End Get
        Set(ByVal value As String)
            _Arguments = value
        End Set
    End Property

    Public Property FieldGUID() As Guid
        Get
            Return _FieldGUID
        End Get
        Set(ByVal value As Guid)
            _FieldGUID = value
        End Set
    End Property

    Public Property FieldType() As GlobalFunctions.FIELDTYPES
        Get
            Return _FieldType
        End Get
        Set(ByVal value As GlobalFunctions.FIELDTYPES)
            _FieldType = value
            UpdateDB = GlobalFunctions.IsDBMapped(value)
        End Set
    End Property

    Public Property UpdateDB() As Boolean
        Get
            Return _UpdateDB
        End Get
        Set(ByVal value As Boolean)
            _UpdateDB = value
        End Set
    End Property

    Public Property BoundField() As String
        Get
            Return _BoundField
        End Get
        Set(ByVal value As String)
            _BoundField = value
        End Set
    End Property

    Public Property FieldCaption() As String
        Get
            Return _FieldCaption
        End Get
        Set(ByVal value As String)
            _FieldCaption = value
        End Set
    End Property

    Public Property IsCompulsory() As Boolean
        Get
            Return _IsCompulsory
        End Get
        Set(ByVal value As Boolean)
            _IsCompulsory = value
        End Set
    End Property

    Public Property FieldControl() As Object
        Get
            Return _FieldControl
        End Get
        Set(ByVal value As Object)
            _FieldControl = value
        End Set
    End Property

    Public Property FieldControl2() As Object
        Get
            Return _FieldControl2
        End Get
        Set(ByVal value As Object)
            _FieldControl2 = value
        End Set
    End Property

    Public Property FieldControl3() As Object
        Get
            Return _FieldControl3
        End Get
        Set(ByVal value As Object)
            _FieldControl3 = value
        End Set
    End Property

    Public Property FieldControl4() As Object
        Get
            Return _FieldControl4
        End Get
        Set(ByVal value As Object)
            _FieldControl4 = value
        End Set
    End Property

End Class

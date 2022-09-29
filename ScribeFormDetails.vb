Imports Microsoft.VisualBasic

Public Class ScribeFormDetails
    Private _Caption As String
    Private _isPreviewMode As Boolean
    Private _isReadOnlyMode As Boolean
    Private _GUID As Guid
    Private _isSubForm As Boolean
    Private _isEditMode As Boolean
    Private _isPostBack As Boolean
    Private _recordID As String
    Private _parentRecordID As String

    Public Property ParentRecordID() As String
        Get
            Return _parentRecordID
        End Get
        Set(value As String)
            _parentRecordID = value
        End Set
    End Property

    Public Property RecordID() As String
        Get
            Return _recordID
        End Get
        Set(value As String)
            _recordID = value
        End Set
    End Property

    Public Property IsPostBack() As Boolean
        Get
            Return _isPostBack
        End Get
        Set(ByVal value As Boolean)
            _isPostBack = value
        End Set
    End Property

    Public Property IsReadonlyMode() As Boolean
        Get
            Return _isReadOnlyMode
        End Get
        Set(ByVal value As Boolean)
            _isReadOnlyMode = value
        End Set
    End Property

    Public Property GUID() As Guid
        Get
            Return _GUID
        End Get
        Set(ByVal value As Guid)
            _GUID = value
        End Set
    End Property

    Public Property IsPreviewMode() As Boolean
        Get
            Return _ispreviewmode
        End Get
        Set(ByVal value As Boolean)
            _ispreviewmode = value
        End Set
    End Property

    Public Property IsEditMode() As Boolean
        Get
            Return _iseditmode
        End Get
        Set(ByVal value As Boolean)
            _isEditmode = value
        End Set
    End Property

    Public Property IsSubform() As Boolean
        Get
            Return _issubform
        End Get
        Set(ByVal value As Boolean)
            _issubform = value
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

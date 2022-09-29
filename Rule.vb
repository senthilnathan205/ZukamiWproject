Imports Microsoft.VisualBasic

Public Class Rule
    Public Enum RuleOperator
        OP_LESS
        OP_MORE
        OP_LESSEQUAL
        OP_MOREEQUAL
        OP_EQUAL
        OP_CONTAINS
        OP_NOTEQUALS
    End Enum

    Private _Fieldname As String = ""
    Private _Type As RuleOperator
    Private _Fieldvalue As String = ""


    Public Property Type() As RuleOperator
        Get
            Return _Type
        End Get
        Set(ByVal value As RuleOperator)
            _Type = value
        End Set
    End Property

    Public Property FieldName() As String
        Get
            Return _Fieldname
        End Get
        Set(ByVal value As String)
            _Fieldname = value
        End Set
    End Property

    Public Property FieldValue() As String
        Get
            Return _Fieldvalue
        End Get
        Set(ByVal value As String)
            _Fieldvalue = value
        End Set
    End Property

End Class

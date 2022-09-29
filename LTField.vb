Imports Microsoft.VisualBasic

Public Class LTField
    Private _fieldcaption As String = ""
    Private _fieldname As String = ""
    Private _fieldid As String = ""
    Private _values As String = ""

    Public Property FieldCaption() As String
        Get
            Return _fieldcaption
        End Get
        Set(value As String)
            _fieldcaption = value
        End Set
    End Property

    Public Property FieldName() As String
        Get
            Return _fieldname
        End Get
        Set(value As String)
            _fieldname = value
        End Set
    End Property

    Public Property FieldID() As String
        Get
            Return _fieldid
        End Get
        Set(value As String)
            _fieldid = value
        End Set
    End Property

    Public Property Values() As String
        Get
            Return _values
        End Get
        Set(value As String)
            _values = value
        End Set
    End Property

    Public Sub Deserialize(Input As String)
        Dim arrMsg() As String = Split(Input, "/;/")
        If UBound(arrMsg) = 3 Then
            _fieldname = arrMsg(0)
            _fieldcaption = arrMsg(1)
            _fieldid = arrMsg(2)
            _values = arrMsg(3)
        End If
    End Sub

    Public Function Serialize() As String
        Return _fieldname & "/;/" & _fieldcaption & "/;/" & _fieldid & "/;/" & _values
    End Function

End Class

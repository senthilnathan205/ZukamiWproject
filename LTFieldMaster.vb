Imports Microsoft.VisualBasic

Public Class LTFieldMaster
    Private _coll As New Collection

    Public ReadOnly Property Items() As Collection
        Get
            Return _coll
        End Get
    End Property

    Public Sub New(ByRef Deserialized As String)
        Dim arrSplits() As String = Split(Deserialized, "|;|")
        _coll.Clear()
        Dim i As Integer
        For i = 0 To UBound(arrSplits)
            Dim strMy As String = Trim(arrSplits(i))
            If Len(strMy) > 0 Then
                Dim lt As New LTField
                lt.Deserialize(strMy)
                _coll.Add(lt, lt.FieldID.ToString)
            End If
        Next
    End Sub

    Public Function Serialize() As String
        Dim i As Integer
        Dim _strmain As String = ""
        For i = 1 To _coll.Count
            Dim lt As LTField = _coll.Item(i)
            If Len(_strmain) > 0 Then _strmain += "|;|"
            _strmain += lt.serialize
        Next i
        Return _strmain
    End Function

End Class

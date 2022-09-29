Public Class JSON
    Private _namecoll As New Collection
    Private _valuecoll As New Collection
    Public Sub AddKey(Name As String, Value As String)
        _namecoll.Add(Name)
        _valuecoll.Add(Value)
    End Sub

    Public Sub AddCollectionKey(Name As String, Value As Collection)
        _namecoll.Add(Name)
        _valuecoll.Add(Value)
    End Sub

    Public Sub AddSingleCollectionKey(Name As String, ByRef JSONObj As JSON)
        _namecoll.Add(Name)
        _valuecoll.Add(JSONObj)
    End Sub

    Private Function EscChars(Input As String) As String
        Dim _input As String = Replace(Input, "\", "\\\\")
        _input = Replace(_input, """", "\\""")
        _input = Replace(_input, vbCrLf, "\\n")
        _input = Replace(_input, vbCr, "\\n")
        _input = Replace(_input, vbLf, "\\n")
        Return _input
    End Function

    Public Function GetJSON() As String
        Dim _start As String = "{"

        Dim i As Integer
        Dim _Cont As String = ""
        For i = 1 To _namecoll.Count
            If Len(_Cont) > 0 Then _Cont += ", "

            If TypeOf _valuecoll.Item(i) Is Collection Then
                _Cont += """" & EscChars(_namecoll.Item(i)) & """ :  ["

                Dim _Cont2 As String = ""
                Dim j As Integer
                Dim _tempcoll As Collection = _valuecoll.Item(i)

                If _tempcoll.Count > 0 Then
                    If TypeOf _tempcoll.Item(1) Is JSON Then
                        For j = 1 To _tempcoll.Count
                            If Len(_Cont2) > 0 Then _Cont2 += ","
                            _Cont2 += CType(_tempcoll.Item(j), JSON).GetJSON()
                        Next j
                    ElseIf TypeOf _tempcoll.Item(1) Is String Then
                        For j = 1 To _tempcoll.Count
                            If Len(_Cont2) > 0 Then _Cont2 += ","
                            _Cont2 += """" & EscChars(CStr(_tempcoll.Item(j))) & """"
                        Next j
                    End If
                End If

                _Cont += _Cont2
                _Cont += "]"
            ElseIf TypeOf _valuecoll.Item(i) Is JSON Then
                _Cont += """" & EscChars(_namecoll.Item(i)) & """ : "
                _Cont += CType(_valuecoll.Item(i), JSON).GetJSON
            Else
                _Cont += """" & EscChars(_namecoll.Item(i)) & """ :   """ & EscChars(_valuecoll.Item(i)) & """"
            End If
        Next
        _start += _Cont
        _start += "}"
        Return _start
    End Function

End Class

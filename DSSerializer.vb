Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Linq
Imports System.Web
Imports System.Xml.Serialization
Imports DevExpress.XtraReports.Native
Imports System.Xml
Imports System.Xml.Schema

Public Class DSSerializer
    Implements IDataSerializer
    Public Const Name As String = "MyDataSerializer"
    Private Shared _set As New Microsoft.VisualBasic.Collection

    Public Shared Sub ApplyDSet(ByRef DSet As DataSet)
        Dim _Settings As ZukamiLib.ZukamiSettings = GetZukamiSettings()

        If _set.Contains(_Settings.CurrentUserGUID.ToString) Then
            _set.Remove(_Settings.CurrentUserGUID.ToString)
        End If

        _set.Add(DSet, _Settings.CurrentUserGUID.ToString)

    End Sub


    Public Shared Function GenerateDSet() As DataSet
        'Try
        Dim _Settings As ZukamiLib.ZukamiSettings = GetZukamiSettings()
        Dim cu As String = _Settings.CurrentUserGUID.ToString
        If _set.Contains(cu) Then
            Dim _myset As DataSet = _set.Item(cu)
            Return _myset
        End If





        'Return _set.Item(_Settings.CurrentUserGUID.ToString)
        'Catch ex As Exception
        '    Return Nothing
        'End Try

    End Function

    Private Function IDataSerializer_CanSerialize(data As Object, extensionProvider As Object) As Boolean Implements IDataSerializer.CanSerialize
        Return TypeOf data Is DataSet
    End Function

    Private Function IDataSerializer_Serialize(data As Object, extensionProvider As Object) As String Implements IDataSerializer.Serialize
        Dim ds = DirectCast(data, DataSet)
        Return ds.DataSetName
    End Function

    Private Function IDataSerializer_CanDeserialize(value As String, typeName As String, extensionProvider As Object) As Boolean Implements IDataSerializer.CanDeserialize
        If StrComp(typeName, "system.data.dataset", vbTextCompare) = 0 Then
            Return True
        Else
            Return False
        End If
        Return True
    End Function

    Private Function IDataSerializer_Deserialize(value As String, typeName As String, extensionProvider As Object) As Object Implements IDataSerializer.Deserialize

        Return GenerateDSet()


    End Function
End Class

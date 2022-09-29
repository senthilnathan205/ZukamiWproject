Imports Microsoft.VisualBasic
Imports System.Xml
Imports System.Data

Public Class FormBehaviors
    Private _behaviors As New Collection

    Public Sub LoadBehaviors(ByVal FormBehaviors As String)
        _behaviors.Clear()
        If Len(FormBehaviors) = 0 Then Exit Sub
        Dim _system As New XmlDocument()
        _system.LoadXml(FormBehaviors)

    End Sub

    Public Function GetDataset() As dataset
        Dim _table As New DataSet
        _table.Tables.Add(New DataTable)
        _table.Tables(0).Columns.Add(New DataColumn("BehaviorID", GetType(Guid)))
        _table.Tables(0).Columns.Add(New DataColumn("Caption", GetType(String)))
        _table.Tables(0).Columns.Add(New DataColumn("Behavior", GetType(String)))

        Return _table
    End Function

    Public Function SaveBehaviors() As String

    End Function

End Class

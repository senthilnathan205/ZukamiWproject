Imports Microsoft.VisualBasic
Imports System.Data

Public Class SubmissionDetails
    Private _RecordID As Guid = Guid.Empty

    Public Property RecordID() As Guid
        Get
            Return _RecordID
        End Get
        Set(ByVal value As Guid)
            _RecordID = value
        End Set
    End Property

    Public Function ActiveNodes() As Collection
        Dim _coll As New Collection
        If RecordID = Guid.Empty Then Return _coll
        Dim _web As New ZukamiLib.WebSession(GetZukamiSettings)
        Dim _Set As DataSet
        _web.OpenConnection()
        _Set = _web.ActiveNodeIDs_GetByRecordID(RecordID)
        _web.CloseConnection()
        If _Set Is Nothing Then
            Return _coll
        End If

        Dim _counter As Integer
        For _counter = 0 To _Set.Tables(0).Rows.Count - 1
            Dim _NodeID As String = GlobalFunctions.FormatData(_Set.Tables(0).Rows(_counter).Item("NodeID"))
            If Len(_NodeID) > 0 Then
                If _coll.Contains(_NodeID) = False Then
                    _coll.Add(_NodeID, _NodeID)
                End If
            End If
        Next _counter
        Return _coll
    End Function

    Public Function WorkflowStage() As String
        If RecordID = Guid.Empty Then Return ""

        Dim _web As New ZukamiLib.WebSession(GetZukamiSettings)
        Dim _Set As DataSet
        _web.OpenConnection()
        _Set = _web.WorkflowInstance_GetByRecordID(RecordID)
        _web.CloseConnection()
        If _Set Is Nothing Then
            Return ""
        End If

        Return GlobalFunctions.FormatData(_Set.Tables(0).Rows(0).Item("StageName"))
    End Function


End Class

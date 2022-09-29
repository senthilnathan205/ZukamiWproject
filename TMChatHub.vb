Imports Microsoft.AspNet.SignalR
Imports Newtonsoft.Json

Public Class TMChatHub
    Inherits Hub

    'Take note that the MS CLR has a 2GB limit, so collection objects can only reach up to 2GB big
    'For bigger objects, more reading here : http://blogs.msdn.com/b/joshwil/archive/2005/08/10/450202.aspx

    Private Shared OnlineUsers As New Collection        'This tracks the users that are online
    Private Shared OnlineStatus As New Collection       'This tracks the current page every user is accessing

    Public Sub Hello()
        Clients.All.Hello()
    End Sub

    Public Sub declareViewing(userID As String, PageID As String)
        If OnlineStatus.Contains(PageID) = True Then
            OnlineStatus.Remove(PageID)
        End If

        OnlineStatus.Add(PageID, "U" & userID)
    End Sub

    Public Sub sendToUser(userID As String, message As String)
        'Try
        Clients.Group("U" & userID).broadcastMessage("Edmund", message)   'We send to a group because we're using single-user groups to send direct messages
        ' Catch ex As Exception

        'End Try
    End Sub

    Public Function sendDirectMessage(GrpMembers() As Integer, userID As String, message As String, SenderName As String) As Integer
        'Dim _chatID As Integer = SaveChatMessage(GrpMembers, userID, TMBL.mgData.CHATCHANNELS.CMODE_USER, message, SenderName)
        'Clients.Group("U" & userID).broadcastMessage(CreateOriginTag(userID, SenderName, _chatID), message)   'We send to a group because we're using single-user groups to send direct messages
        'Return _chatID
    End Function

    Public Sub sendToUserArray(userArray() As Integer, message As String)
        Dim i As Integer
        For i = 0 To UBound(userArray)
            sendToUser(userArray(i), message)
        Next
    End Sub


    Public Sub sendToUsers(userIDs As String, message As String)
        Dim arrUserIDs() As String = Split(userIDs, ";")
        Dim i As Integer
        For i = 0 To UBound(arrUserIDs)
            Dim _uid As Integer = arrUserIDs(i)
            If IsNumeric(_uid) Then
                sendToUser(_uid, message)
            End If
        Next
    End Sub

    Public Sub sendToWholeWorld(Name As String, message As String)
        Clients.All.broadcastmessage(Name, message)
    End Sub

    'Joins the current user to the room specified by RoomName
    Public Function joinUserChannel(UserID As String) As System.Threading.Tasks.Task
        Return Groups.Add(Context.ConnectionId, "U" & UserID)
    End Function

    'Joins the current user to the room specified by RoomName
    Public Function joinRoom(UserID As String, GroupID As String) As System.Threading.Tasks.Task
        Return Groups.Add(Context.ConnectionId, "G" & GroupID)
    End Function

    Public Function LeaveRoom(GroupID As String) As System.Threading.Tasks.Task
        'we need to create code here to remove user from a group
        Return Groups.Remove(Context.ConnectionId, "G" & GroupID)
    End Function

    Public Sub sendGrpMsg(GroupID As String, Message As String)
        Clients.Group("G" & GroupID).broadcastMessage("Edmund", Message)
    End Sub

    Public Function sendOtherMsg(GrpMembers() As Integer, GroupID As String, Message As String, SenderName As String) As Integer
        ''Save the chat into table
        'Dim _msgID As Integer = SaveChatMessage(GrpMembers, GroupID, TMBL.mgData.CHATCHANNELS.CMODE_GROUP, Message, SenderName)
        'Clients.OthersInGroup("G" & GroupID).broadcastMessage(CreateOriginTag(GroupID, SenderName, _msgID), Message)
        'Return _msgID
    End Function

    Public Sub indicateTyping(GroupID As String, SenderName As String, IsTyping As Boolean)
        Dim _tag As String = "cmd_it"
        If IsTyping = False Then
            _tag = "cmd_nt"
        End If
        Clients.OthersInGroup("G" & GroupID).broadcastMessage(CreateOriginTag(GroupID, SenderName, -1), _tag & ";" & SenderName & ";1")
    End Sub

    'Public Sub postDeleteRequest(MsgID As String, PageType As String, GroupID As String, SenderName As String)
    '    Dim _tag As String = "cmd_dp"

    '    Dim _do As New TMBL.mgData
    '    Dim _res As TMBL.RetObject = _do.OpenConn(WebconfigSettings.ConnectionString)
    '    If _res.Success = True Then
    '        _do.DeleteMessage(MsgID)
    '    End If
    '    _do.CloseConn()
    '    _do = Nothing

    '    Clients.OthersInGroup(PageType & GroupID).broadcastMessage(CreateOriginTag(GroupID, SenderName, -1), _tag & ";" & SenderName & ";" & MsgID & ";" & GroupID & ";" & PageType)
    'End Sub

    'Private Function GetUnfocusedOtherUser(GrpMembers() As Integer, UserID As String) As String
    '    Dim _key As String = "U" & UserID
    '    Dim _hasfocus As Boolean = True
    '    Dim _unfocused As String = ""
    '    Dim _userpgID As String = Generic.GetCModeChar(CHATMODE.CMODE_USER) & UserID

    '    If UBound(GrpMembers) >= 0 Then
    '        If OnlineStatus.Contains(_key) Then
    '            Dim _pg As String = OnlineStatus.Item(_key)
    '            If StrComp(_pg, _userpgID, CompareMethod.Text) <> 0 Then
    '                _hasfocus = False
    '            End If
    '        Else
    '            _hasfocus = False
    '        End If
    '        If _hasfocus = False Then
    '            If Len(_unfocused) > 0 Then _unfocused += ";"
    '            _unfocused = _unfocused & GrpMembers(0)
    '        End If
    '    End If
    '    Return _unfocused
    'End Function

    'Private Function GetUnfocusedUsersList(GrpMembers() As Integer, GroupID As String) As String
    '    Dim i As Integer
    '    Dim _grouppgID As String = Generic.GetCModeChar(CHATMODE.CMODE_GROUP) & GroupID
    '    Dim _unfocused As String = ""
    '    For i = 0 To UBound(GrpMembers)
    '        If OnlineStatus Is Nothing = False Then
    '            Dim _key As String = "U" & GrpMembers(i)
    '            Dim _hasfocus As Boolean = True
    '            If OnlineStatus.Contains(_key) Then
    '                Dim _pg As String = OnlineStatus.Item(_key)
    '                If StrComp(_pg, _grouppgID, CompareMethod.Text) <> 0 Then
    '                    _hasfocus = False
    '                End If
    '            Else
    '                _hasfocus = False
    '            End If
    '            If _hasfocus = False Then
    '                If Len(_unfocused) > 0 Then _unfocused += ";"
    '                _unfocused = _unfocused & GrpMembers(i)
    '            End If
    '        End If
    '    Next
    '    Return _unfocused
    'End Function

    'Private Function SaveChatMessage(GrpMembers() As Integer, ObjectID As String, CMode As TMBL.mgData.CHATCHANNELS, Message As String, SenderInfo As String) As Integer
    '    SaveChatMessage = -1
    '    Dim _do As New TMBL.mgData
    '    Dim _res As TMBL.RetObject = _do.OpenConn(WebconfigSettings.ConnectionString)
    '    If _res.Success = True Then
    '        Dim arrSI() As String = Split(SenderInfo, ";")
    '        If UBound(arrSI) = 1 Then
    '            'we check the users in this group , and then check their status. We create notifications for those that are not around
    '            'The unfocusedusers variable contains a string of users that are currently unfocused. Eg: 2,4,8
    '            Dim _unfocusedusers As String = ""
    '            If CMode = TMBL.mgData.CHATCHANNELS.CMODE_GROUP Then
    '                _unfocusedusers = GetUnfocusedUsersList(GrpMembers, ObjectID)
    '            ElseIf CMode = TMBL.mgData.CHATCHANNELS.CMODE_USER Then
    '                _unfocusedusers = GetUnfocusedOtherUser(GrpMembers, ObjectID)
    '            End If
    '            Dim _retval As TMBL.RetObject = _do.CreateMessage(arrSI(1), arrSI(0), Message, ObjectID, CMode, _unfocusedusers)
    '            SaveChatMessage = CInt(_retval.RetObj1)
    '        End If
    '        _do.CloseConn()
    '    End If
    '    _do = Nothing

    'End Function

    Private Function CreateOriginTag(ObjectID As String, Sender As String, MsgID As Integer) As String
        Return ObjectID & ";" & Sender & ";" & MsgID
    End Function

    'Public Overrides Function OnDisconnected(stopCalled As Boolean) As Threading.Tasks.Task
    '    Dim _cke As Microsoft.AspNet.SignalR.Cookie = Context.Request.Cookies.Item("mgly")
    '    Dim _mguser As TMBL.mgUser = CheckSignalRCookie(_cke)
    '    If _mguser Is Nothing Then Return Nothing

    '    If OnlineUsers Is Nothing = False Then
    '        'remove from onlineusers collection
    '        If OnlineUsers.Contains(_mguser.ID.ToString) = True Then
    '            OnlineUsers.Remove(_mguser.ID.ToString)
    '        End If
    '    End If

    '    'we also want to signal to other users that this guy has signed out
    '    GetOffBoard(_mguser.ID.ToString)


    '    Return MyBase.OnDisconnected(stopCalled)
    'End Function

    'Private Function CheckSignalRCookie(ByRef mgck As Microsoft.AspNet.SignalR.Cookie) As TMBL.mgUser
    '    Dim mgUsr As TMBL.mgUser
    '    If mgck Is Nothing Then Return Nothing
    '    mgUsr = New TMBL.mgUser

    '    'Decode from base64
    '    Dim decodedBytes As Byte()
    '    decodedBytes = Convert.FromBase64String(mgck.Value)
    '    Dim decodedText As String
    '    decodedText = Encoding.UTF8.GetString(decodedBytes)

    '    Dim _scv As SessionCookieValues = JsonConvert.DeserializeObject(Of SessionCookieValues)(decodedText)

    '    mgUsr.ID = _scv.ID
    '    mgUsr.Lang = _scv.Lang
    '    mgUsr.Country = _scv.Country
    '    mgUsr.LiveStatus = _scv.LiveStatus
    '    Return mgUsr
    'End Function

    Public Overrides Function OnConnected() As System.Threading.Tasks.Task

        'Context.User.Identity.Name always returns empty because we're not using the Microsoft authentication functions
        'Dim _cke As Microsoft.AspNet.SignalR.Cookie = Context.Request.Cookies.Item("mgly")
        'Dim _mguser As TMBL.mgUser = CheckSignalRCookie(_cke)

        'If _mguser Is Nothing = False Then
        '    'Create single user group
        '    Dim _clientid As Integer = _mguser.ID

        '    'add to the onlineusers collection
        '    If OnlineUsers.Contains(_clientid.ToString) = False Then
        '        OnlineUsers.Add(_clientid.ToString, _clientid.ToString)
        '    End If

        '    Dim _connid As String = Context.ConnectionId
        '    Groups.Add(_connid, "U" & _clientid)

        '    'Automatically join all previous joined groups and change the user's status to online, 
        '    'and grab list of all active users that has this user in their contact lists
        '    'and send a push to all online contacts
        '    GetOnBoard(_clientid)
        '    Return MyBase.OnConnected
        'Else
        '    Return Nothing
        'End If

        Return MyBase.OnConnected

    End Function

    'Private Sub GetOffBoard(UserID As Integer)
    '    Dim _do As New TMBL.mgData
    '    Dim _res As TMBL.RetObject
    '    Generic.PrintIfError(_do.OpenConn(WebconfigSettings.ConnectionString))
    '    _res = _do.GetOffChat(UserID)
    '    Generic.PrintIfError(_do.CloseConn())
    '    If _res.Success = True Then
    '        Dim _contacts As Collection = _res.RetObj1
    '        'Send a push to all online contacts
    '        For Each ctct As TMBL.mgContact In _contacts
    '            sendToUser(ctct.ID, "cmd;" & UserID.ToString & ";0")
    '        Next
    '        _contacts.Clear()
    '        _contacts = Nothing
    '    Else
    '        Generic.PrintIfError(_res)
    '    End If
    'End Sub

    'Private Sub GetOnBoard(UserID As Integer)
    '    Dim _do As New TMBL.mgData
    '    Dim _res As TMBL.RetObject
    '    Generic.PrintIfError(_do.OpenConn(WebconfigSettings.ConnectionString))
    '    _res = _do.GetOnChat(UserID)
    '    Generic.PrintIfError(_do.CloseConn())
    '    If _res.Success = True Then
    '        Dim _groups As Collection = _res.RetObj1
    '        Dim _contacts As Collection = _res.RetObj2
    '        'Now we rejoin all the groups
    '        For Each grp As TMBL.mgGroup In _groups
    '            joinRoom(UserID, grp.ID.ToString)
    '        Next

    '        'Send a push to all online contacts
    '        For Each ctct As TMBL.mgContact In _contacts
    '            sendToUser(ctct.ID, "cmd;" & UserID.ToString & ";1")
    '        Next

    '        _groups.Clear()
    '        _groups = Nothing
    '    Else
    '        Generic.PrintIfError(_res)
    '    End If
    'End Sub
End Class


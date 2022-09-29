Imports System.Data
Imports Microsoft.VisualBasic
Imports NLog

Public Class LoginTicket
    Private Shared logger As Logger = LogManager.GetCurrentClassLogger()

    Public _ticket As String = ""
    Public _creationTime As DateTime?
    Public _username As String = ""
    Public Sub New(r As DataRow)
        logger.Debug("start")
        If r IsNot Nothing Then
            _ticket = GlobalFunctions.FormatData(r.Item("LoginTicket"))
            _username = GlobalFunctions.FormatData(r.Item("Username"))
            If Not IsDBNull(r.Item("LoginTicketCreationTime")) Then
                'creationTime = r.Item("LoginTicketCreationTime")
                'DateTime.TryParse(r.Item("LoginTicketCreationTime"), creationTime)
                _creationTime = CType(r.Item("LoginTicketCreationTime"), DateTime)
            End If
            logger.Debug("Username: " + GlobalFunctions.FormatData(r.Item("username")) + ", ticket: " + _ticket + ", creationTime: " + GlobalFunctions.FormatDateTime(_creationTime))
        Else
            logger.Debug("data row is nothing, cannot init")
        End If
    End Sub
    Public Sub New(username As String)
        _ticket = Guid.NewGuid().ToString()
        _creationTime = DateTime.Now
        _username = username
    End Sub
    Public Shared Sub TakeLoginActions(username As String)
        logger.Debug("start")
        Try
            Dim loginTicketFromDB As New LoginTicket(GetLoginTicketRowFromDB(username))
            If Not loginTicketFromDB.HasTicket() OrElse loginTicketFromDB.TicketExpired() Then
                Dim t As New LoginTicket(username)
                t.IssueNewLoginTicket(HttpContext.Current)
            Else
                Dim loginTicketCookie As HttpCookie = GetLoginTicketCookie()
                If loginTicketCookie Is Nothing OrElse loginTicketCookie.Value.ToLower() <> loginTicketFromDB._ticket.ToLower() Then
                    ' do nothing
                Else
                    loginTicketFromDB.RefreshCreationTimeToDB()
                    loginTicketFromDB.SaveToCookie(HttpContext.Current)
                End If
            End If
            logger.Debug("done")
        Catch ex As Exception
            logger.Error(ex)
        End Try
    End Sub
    Public Shared Function IsMultipleLogin(username As String) As Boolean
        logger.Debug("start")
        Dim isMultLogin As Boolean = False
        Try
            Dim loginTicketFromDB As New LoginTicket(GetLoginTicketRowFromDB(username))
            If Not loginTicketFromDB.HasTicket() OrElse loginTicketFromDB.TicketExpired() Then
                'Dim t As New LoginTicket(username)
                't.IssueNewLoginTicket(HttpContext.Current)
                isMultLogin = False
            Else
                Dim loginTicketCookie As HttpCookie = GetLoginTicketCookie()
                If loginTicketCookie Is Nothing OrElse loginTicketCookie.Value.ToLower() <> loginTicketFromDB._ticket.ToLower() Then
                    isMultLogin = True
                Else
                    'loginTicketFromDB.RefreshCreationTimeToDB()
                    'loginTicketFromDB.SaveToCookie(HttpContext.Current)
                    isMultLogin = False
                End If
            End If
        Catch ex As Exception
            logger.Error(ex)
        End Try
        logger.Debug("isMultLogin: " + isMultLogin.ToString())
        Return isMultLogin
    End Function
    Public Shared Function GetLoginTicketRowFromDB(username As String) As DataRow
        Dim r As DataRow = Nothing
        Dim ticket As String = ""
        Dim sql As String = String.Format("select * from [users] where username = N'{0}' ", username.Replace("'", "''"))
        Dim dt As DataTable = GlobalFunctions.GetDataTableBySql(sql)
        If dt.Rows.Count > 0 Then
            r = dt.Rows(0)
        End If
        Return r
    End Function
    Public Shared Function GetLoginTicketRowFromDB(userid As Guid) As DataRow
        Dim r As DataRow = Nothing
        Dim ticket As String = ""
        Dim sql As String = String.Format("select * from [users] where UserID = N'{0}' ", userid.ToString())
        Dim dt As DataTable = GlobalFunctions.GetDataTableBySql(sql)
        If dt.Rows.Count > 0 Then
            r = dt.Rows(0)
        End If
        Return r
    End Function
    Public Shared Function GetLoginTicketCookie() As HttpCookie
        logger.Debug("start")
        Dim theCookie As HttpCookie = HttpContext.Current.Request.Cookies("LoginTicket")
        If theCookie Is Nothing Then
            logger.Debug("LoginTicket not found from cookie")
        Else
            logger.Debug("LoginTicket value: " + theCookie.Value)
        End If
        Return theCookie
    End Function
    Public Sub IssueNewLoginTicket(content As HttpContext)
        logger.Debug("start")
        SaveToDB()
        SaveToCookie(content)
        logger.Debug("done")
    End Sub
    Public Sub SaveToDB()
        logger.Debug("start")
        Dim sql As String = String.Format(
            " update [users] set [LoginTicket] = '{0}', [LoginTicketCreationTime] = {1} where [username] = N'{2}' ",
            _ticket.Replace("'", "''"), GetDbFormatDateTimeString(_creationTime), _username.Replace("'", "''"))
        GlobalFunctions.RunSql(sql)
    End Sub
    Public Function GetDbFormatDateTimeString(d As DateTime?) As String
        If d.HasValue Then
            Return "'" + d.Value.ToString("yyyy-MM-dd HH:mm:ss") + "'"
        Else
            Return "null"
        End If
    End Function
    Public Sub RemoveLoginTicketFromDB()
        logger.Debug("start")
        Dim sql As String = String.Format(
            " update [users] set [LoginTicket] = null, [LoginTicketCreationTime] = null where [username] = N'{0}' ",
            _username.Replace("'", "''"))
        GlobalFunctions.RunSql(sql)
    End Sub
    Public Sub RemoveLoginTicketFromCookie(content As HttpContext)
        logger.Debug("start")
        GlobalFunctions.RemoveCookie("LoginTicket")
        'Dim ck As New HttpCookie("LoginTicket", _ticket)
        'ck.Expires = DateTime.Now.AddDays(-1)
        'content.Response.Cookies.Remove("LoginTicket")
        'content.Response.Cookies.Add(ck)
    End Sub

    Public Sub ClearUpLoginTicket(content As HttpContext)
        logger.Debug("start")
        Dim cookie As HttpCookie = GetLoginTicketCookie()
        If cookie IsNot Nothing Then
            If _ticket = cookie.Value Then
                RemoveLoginTicketFromDB()
            Else
                logger.Debug("cannot clearup ticket in db as the ticket in cookie " + cookie.Value + " is not match with the db ticket " + _ticket)
            End If
        End If

        RemoveLoginTicketFromCookie(content)
    End Sub

    Public Sub RefreshCreationTimeToDB()
        logger.Debug("start")
        _creationTime = DateTime.Now
        SaveToDB()
    End Sub
    Public Sub SaveToCookie(content As HttpContext)
        logger.Debug("start")
        Dim ck As New HttpCookie("LoginTicket", _ticket)
        ck.Expires = New DateTime(Year(Now) + 20, 12, 31) ' +20 years so it's like never expires
        content.Response.Cookies.Remove("LoginTicket")
        content.Response.Cookies.Add(ck)
    End Sub
    Public Function HasTicket() As Boolean
        Dim has As Boolean = False
        If Not String.IsNullOrWhiteSpace(_ticket) Then
            has = True
        End If
        logger.Debug("ticket: " + _ticket + ", has: " + has.ToString())
        Return has
    End Function

    Public Function TicketExpired() As Boolean
        Dim expired As Boolean = True
        Dim cookieExpiryMinutes As Integer = GlobalFunctions.FormatInteger(GlobalFunctions.FromConfig("InactivityTimeout", "20"))
        If _creationTime.HasValue Then
            Dim expiryTime As DateTime = _creationTime.Value.AddMinutes(cookieExpiryMinutes)
            Dim nowTime As DateTime = DateTime.Now
            If DateTime.Compare(nowTime, expiryTime) <= 0 Then
                expired = False
            End If
            logger.Debug("expiryTime: " + GlobalFunctions.FormatDateTime(expiryTime) +
                         ", nowTime: " + GlobalFunctions.FormatDateTime(nowTime) +
                         ", expired: " + expired.ToString())
        Else
            logger.Debug("creationTime has no value, cannot compare")
        End If
        logger.Debug("expired: " + expired.ToString())
        Return expired
    End Function
End Class
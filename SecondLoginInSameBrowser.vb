Imports Microsoft.VisualBasic
Imports NLog

Public Class SecondLoginInSameBrowser
    Private Shared logger As Logger = LogManager.GetCurrentClassLogger()
    Public Shared Function IsSecondLogin(username As String) As Boolean
        Dim is2nd As Boolean = False
        Dim currentUserHash As String = GlobalFunctions.OneWayHashWithHardCodedSalt(username)
        Dim userCookie As HttpCookie = HttpContext.Current.Request.Cookies("CurrentLoggedInUser")
        If userCookie IsNot Nothing AndAlso userCookie.Value <> currentUserHash Then
            is2nd = True
        Else
            is2nd = False
        End If
        Return is2nd
    End Function
    Public Shared Sub TakeLoginActions(username As String)
        logger.Debug("start")
        Dim currentUserHash As String = GlobalFunctions.OneWayHashWithHardCodedSalt(username)
        HttpContext.Current.Response.Cookies.Remove("CurrentLoggedInUser")
        Dim ck As HttpCookie = New HttpCookie("CurrentLoggedInUser", currentUserHash)
        ' ck.Expires = New DateTime(Year(Now) + 20, 12, 31) ' +20 years so it's like never expires
        ck.Expires = DateAdd(DateInterval.Minute, GlobalFunctions.FormatInteger(GlobalFunctions.FromConfig("InactivityTimeout", "20")), Now)
        HttpContext.Current.Response.Cookies.Add(ck)
        logger.Debug("done")
    End Sub

End Class

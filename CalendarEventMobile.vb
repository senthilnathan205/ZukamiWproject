Imports System.Data
Imports Microsoft.VisualBasic

Public Class CalendarEventMobile
    Public Day As DateTime
    Public Time As String
    Public Title As String
    Public Link As String

    Public Sub New(id As String, _appID As String, listId As String, _day As String, fromTo As String, caption As String)
        Day = _day
        Time = fromTo
        Title = caption
        Link = "FillFormMobile.aspx?RO=1&FT=1&ListID=" & listId & "&ID=" & id & "&a=" & _appID
    End Sub
End Class

Imports Microsoft.VisualBasic
Imports System.Web.HttpContext
Imports NLog
Imports System.Threading

Public Class NavigatorMobile

    Private Shared logger As Logger = LogManager.GetCurrentClassLogger()

    Public Shared Sub CaptureMileStone(ByRef MyPage As Page, Optional ByVal MileStonePage As String = "")
        Try
            If MyPage.IsPostBack = False Then
                If MyPage.Request.Url Is Nothing = False Then
                    Dim _currentURL As String = MyPage.Request.Url.PathAndQuery
                    If MileStonePage = "" Then
                        MyPage.Session("PageMilestone") = _currentURL
                    Else
                        MyPage.Session("PageMilestone" & MileStonePage) = _currentURL
                    End If

                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Public Shared Sub RevertToPreviousMileStone(ByRef MyPage As Page, ByVal DefaultLink As String, Optional ByVal MileStonePage As String = "")

        If MileStonePage = "" Then
            If Len(MyPage.Session("PageMilestone")) > 0 Then
                MyPage.Response.Redirect(MyPage.Session("PageMilestone"))
            Else
                MyPage.Response.Redirect(DefaultLink)
            End If
        Else
            If Len(MyPage.Session("PageMilestone" & CStr(MileStonePage))) > 0 Then
                MyPage.Response.Redirect(MyPage.Session("PageMilestone" & CStr(MileStonePage)))
            Else
                MyPage.Response.Redirect(DefaultLink)
            End If
        End If



    End Sub

    Public Shared Function GetJScriptURL(ByVal TargetLocation As String, Optional ByVal Param1 As Object = Nothing, Optional ByVal Param2 As Object = Nothing, Optional ByVal Param3 As Object = Nothing, Optional ByVal Param4 As Object = Nothing, Optional ByVal Param5 As Object = Nothing, Optional ByVal Param6 As Object = Nothing, Optional ByVal Param7 As Object = Nothing, Optional ByVal Param8 As Object = Nothing, Optional ByVal Param9 As Object = Nothing) As String
        Select Case TargetLocation
            Case "LaunchMobileFilters"
                Return "mobileviewsettings.aspx?AppID=" & CStr(Param1) & "&ID=" & CStr(Param2)
            Case "ERD"
                Return "ERD.aspx?AppID=" & CStr(Param1) & "&ID=" & CStr(Param2)
            Case "CodeBehindView"
                Return "codebackend2.aspx?AppID=" & CStr(Param1) & "&ID=" & CStr(Param2)
            Case "LaunchMobileSorting"
                Return "mobileorderingsettings.aspx?AppID=" & CStr(Param1) & "&ID=" & CStr(Param2)
            Case "Mobility"
                Return "MobilitySettings.aspx?AppID=" & CStr(Param1) & "&ID=" & CStr(Param2)
            Case "CodeBehind"
                Return "CodeBackend.aspx?AppID=" & CStr(Param1) & "&ID=" & CStr(Param2)
            Case "ReportViewer"
                Return "ReportViewer.aspx?a=" & CStr(Param1) & "&FT=1&ViewID=" & CStr(Param2)
            Case "MapViewer"
                Return "MapList.aspx?a=" & CStr(Param1) & "&FT=1&ViewID=" & CStr(Param2)
            Case "MapViewerWithRoute"
                Return "MapList.aspx?RouteID=" & CStr(Param3) & "&a=" & CStr(Param1) & "&FT=1&ViewID=" & CStr(Param2)
            Case "DeclareReusable"
                Return "newreusable.aspx?appid=" & CStr(Param1) & "&ID=" & CStr(Param2)
            Case "DesignReport"
                Return "reportdesigner.aspx?appid=" & CStr(Param1) & "&ID=" & CStr(Param2)
            Case "Jump"
                Return "jumpMobile.aspx?WFlowID=" & Param1 & "&ID=" & Param2 & "&a=" & Param3 & "&ID2=" & Param4
            Case "EditFormDevice"
                Return "EForm.aspx?DM=" & Param4 & "&AppID=" & Param1 & "&ID=" & Param2 & IIf(Len(Param3) > 0, "&Tab=" & Param3, "") & IIf(Len(Param5) > 0, "&Del=" & Param5, "")
            Case "CloudSettings"
                Return "cloudsettings.aspx?appid=" & CStr(Param1)
            Case "RunAddPage"
                Return "runaddpage.aspx?appid=" & CStr(Param1)
            Case "NavList"
                Return "navlist.aspx?appid=" & CStr(Param1)
            Case "SecMatrix"
                Return "secmatrix.aspx?appid=" & CStr(Param1)
            Case "Default"
                Return GlobalFunctions.GrabPostLoginAppPage(Param1)
            Case "AboutResource"
                Return "resourceinfo.aspx?appid=" & CStr(Param1) & "&Type=" & CStr(Param2) & "&ID=" & CStr(Param3)
            Case "GenerateDoc"
                Return "generatedocumentation.aspx?appid=" & CStr(Param1) & "&ID=" & CStr(Param2)
            Case "GeneratePDF"
                Return "generatePDF.aspx?appid=" & CStr(Param1) & "&ID=" & CStr(Param2)
            Case "GlobalCode"
                Return "globalcode.aspx?appid=" & CStr(Param1)
            Case "ChangeHistory"
                Return "AuditTrailViewer.aspx?a=" & CStr(Param1) & "&SQ=" & Param2 & "&ReturnURL=" & Param3
            Case "CustomRedirect"
                Return Param1
            Case "ImageMap"
                Return "viewsettings.aspx?AppID=" & CStr(Param1) & "&ID=" & CStr(Param2) & "&Tab=" & Param3
            Case "TreeView"
                Return "viewsettings.aspx?AppID=" & CStr(Param1) & "&ID=" & CStr(Param2) & "&Tab=" & Param3
            Case "AppInfo"
                Return "appinfo.aspx?ID=" & Param1
            Case "ExportApp"
                Return "exportapp.aspx?ID=" & Param1
            Case "SendForReview"
                Dim strParam1 As String = Param1
                'If GlobalFunctions.FormatBoolean(Param2) = True Then
                Return "sendforreviewMobile.aspx?ID=" & strParam1 & "&InstanceID=" & Param3 & "&NodeID=" & Param4 & "&ID2=" & Param5 & "&a=" & Param6
            ' Else
            ' Return "sendforreview.aspx?ID=" & strParam1 & "&a=" & Param6
            ' End If
            Case "SendForInput"
                Dim strParam1 As String = Param1
                '   If GlobalFunctions.FormatBoolean(Param2) = True Then
                Return "sendforreviewMobile.aspx?SFI=1&ID=" & strParam1 & "&InstanceID=" & Param3 & "&NodeID=" & Param4 & "&ID2=" & Param5 & "&a=" & Param6
            '   Else
            '   Return "sendforreview.aspx?SFI=1&ID=" & strParam1 & "&a=" & Param6
            '   End If
            Case "CreateDuplicateView"
                Return "createduplicateview.aspx?a=" & Param1 & "&ID=" & Param2
            Case "CorrectionDone"
                Dim strParam1 As String = Param1
                Dim strParam2 As String = Param2
                Dim strParam3 As String = Param3
                Return "correctionmadeMobile.aspx?ID=" & strParam1 & "&OriginatorID=" & strParam2 & "&InstanceID=" & strParam3 & "&a=" & Param4
            Case "InputDone"
                Dim strParam1 As String = Param1
                Dim strParam2 As String = Param2
                Dim strParam3 As String = Param3
                Dim strParam5 As String = Param5
                Return "correctionmadeMobile.aspx?SFI=1&ID=" & strParam1 & "&OriginatorID=" & strParam2 & "&InstanceID=" & strParam3 & "&a=" & Param4 & "&TaskID=" & strParam5
            Case "Correction"
                Dim strParam1 As String = Param1
                Return "actionMobile.aspx?Action=Correction&ID=" & strParam1
            Case "UserLists"
                Return "applists.aspx?Type=0&AppID=" & Param1
            Case "Ordering"
                Return "viewordering.aspx?AppID=" & Param1 & "&ID=" & Param2
            Case "Lists"
                Return "listtypes.aspx?AppID=" & Param1
            Case "NavLists"
                Return "navlist.aspx?AppID=" & Param1
            Case "DatasourceConns"
                Return "datasourceconnections.aspx?AppID=" & Param1
            Case "ExternalConns"
                Return "externalconns.aspx?AppID=" & Param1
            Case "Actions"
                Return "viewactions.aspx?AppID=" & Param1 & "&ID=" & Param2
            Case "Webservices"
                Return "webservices.aspx?AppID=" & Param1
            Case "Libraries"
                Return "libraries.aspx?AppID=" & Param1
            Case "SetPermissions"
                Return "setpermissions.aspx?AppID=" & Param1 & "&ID=" & Param2
            Case "SubstitutionStrings"
                Return "substitutionstrings.aspx?AppID=" & Param1
            Case "EditBackgroundJob"
                Return "ScheduledBehavior.aspx?AppID=" & Param1 & "&ID=" & Param2
            Case "NewBackJob"
                Return "ScheduledBehavior.aspx?AppID=" & Param1
            Case "ImportAsCSVDM"
                Return "importascsv.aspx?dm=1&a=" & Param1 & "&ViewID=" & Param2
            Case "ManageUsers"
                Return "manageusers.aspx?AppID=" & Param1
            Case "ViewBranches"
                Return "viewbranches.aspx?AppID=" & Param1 & "&ID=" & Param2
            Case "AppBranches"
                Return "appbranches.aspx?AppID=" & Param1
            Case "FormBranches"
                Return "formbranches.aspx?AppID=" & Param1 & "&ID=" & Param2
            Case "NewSearchForm"
                Return "searchproperties.aspx?AppID=" & Param1
            Case "SearchConfigPage"
                Return "Searchconfigpage.aspx?AppID=" & Param1 & "&ID=" & Param2
            Case "SearchConfigPagePublish"
                Return "Searchconfigpage.aspx?AppID=" & Param1 & "&ID=" & Param2 & "&Publish=" & Param3
            Case "LaunchCompositeView"
                Return "compositeview.aspx?ID=" & Param1 & "&a=" & Param2
            Case "CompositePageContent"
                Return "compositepagecontent.aspx?AppID=" & Param1 & "&ID=" & Param2
            Case "EditCompositeViewPage"
                Return "compositepage.aspx?AppID=" & Param1 & "&ID=" & Param2
            Case "EditCompositePageConfig"
                Return "compositepageconfig.aspx?AppID=" & Param1 & "&ID=" & Param2
            Case "NewCompositePage"
                Return "compositepage.aspx?AppID=" & Param1
            Case "Settings"
                Return "settings.aspx?a=" & Param1
            Case "Dashboard"
                Return "dashboard.aspx?a=" & Param1
            Case "AppTabs"
                Return "AppTabs.aspx?AppID=" & Param1
            Case "PageSettings"
                Return "pagesettings.aspx?AppID=" & Param1 & "&ID=" & Param2 & "&Type=" & Param3
            Case "Autoarchiving"
                Return "Archiving.aspx?AppID=" & Param1 & "&ID=" & Param2
            Case "ResourceFiles"
                Return "resourcefiles.aspx?AppID=" & Param1
            Case "StopPreview"
                Return "StopPreview.aspx?AppID=" & Param1
            Case "EditFormField"
                Return "EditFormField.aspx?ListItemID=" & Param1 & "&ID=" & Param2 & "&AppID=" & Param3
            Case "EditTheme"
                Return "newtheme.aspx?AppID=" & Param1 & "&ID=" & Param2
            Case "RunApp"
                Return "login.aspx?AppID=" & Param1 & IIf(Len(Param2) > 0, "&PMode=" & Param2, "")
            Case "ChangeLang"
                Return "changelang.aspx?AppID=" & Param1 & "&ID=" & Param2
            Case "ImportForm"
                Return "importform.aspx?ID=" & Param1
            Case "ExportForm"
                Return "exportform.aspx?AppID=" & Param1 & "&ID=" & Param2
            Case "Aggregate"
                Return "aggregatedesigner.aspx?AppID=" & Param1 & "&ID=" & Param2
            Case "ExportView"
                Return "exportview.aspx?AppID=" & Param1 & "&ViewID=" & Param2
            Case "EditViewStyles"
                Return "viewstyles.aspx?AppID=" & Param1 & "&ID=" & Param2
            Case "CheckWorkflow"
                Return "checkworkflow.aspx?ID=" & Param1 & "&AppID=" & Param2 & "&ListID=" & Param3
            Case "EditFormConfigs"
                Return "FormConfigs.aspx?ListID=" & Param1 & "&AppID=" & Param2
            Case "EditMobileFormConfigs"
                Return "FormDeviceConfigs.aspx?ListID=" & Param1 & "&AppID=" & Param2
            Case "EditFormPageConfig"
                Return "EFormConfigPage.aspx?AppID=" & Param1 & "&ID=" & Param2
            Case "EditFormPageConfigPublish"
                Return "EFormConfigPage.aspx?AppID=" & Param1 & "&ID=" & Param2 & "&Publish=" & Param3
            Case "LaunchViewSettingsPublish"
                Return "viewconfigpage.aspx?AppID=" & CStr(Param1) & "&ID=" & CStr(Param2) & "&Publish=" & CStr(Param3)
            Case "MyApplications"
                Return "myapplications.aspx"
            Case "Application2"
                If Len(Param2) = 0 Then
                    Return "CreateApp.aspx?ID=" & Param1
                Else
                    Return "CreateApp.aspx?ID=" & Param1 & "&Tab=" & Param2
                End If
            Case "Applications"
                Return "applications.aspx"
            Case "EditAppConfigPage"
                Return "AppConfigPage.aspx?ID=" & Param1
            Case "EditApplication"
                If Len(Param2) = 0 Then
                    Return "CreateApp.aspx?ID=" & Param1
                Else
                    Return "CreateApp.aspx?ID=" & Param1 & "&Tab=" & Param2
                End If
            Case "SearchFormSettings"
                Return "SearchFormSettings.aspx?AppID=" & CStr(Param1) & "&ID=" & CStr(Param2)
            Case "LaunchViewSettings"
                Return "viewconfigpage.aspx?AppID=" & CStr(Param1) & "&ID=" & CStr(Param2)
            Case "ConvertViewToMobile"
                Return "viewconfigpage.aspx?AppID=" & CStr(Param1) & "&ID=" & CStr(Param2) & "&ConvToggle=1"
            Case "EditView"
                Return "viewsettings.aspx?AppID=" & CStr(Param1) & "&ID=" & CStr(Param2) & "&Tab=" & Param3
            Case "EditForm"
                Return "EForm.aspx?AppID=" & Param1 & "&ID=" & Param2 & IIf(Len(Param3) > 0, "&Tab=" & Param3, "")
            Case "NewView"
                Return "viewsettings.aspx?AppID=" & CStr(Param1)
            Case "NewForm"
                Return "eform.aspx?AppID=" & CStr(Param1)
            Case "PrintFriendly"
                Return "printfriendly.aspx?FT=1&ListID=" & CStr(Param1) & "&ID=" & CStr(Param2) & "&a=" & Param3
            Case "CreateApp"
                Return "createapp.aspx"
            Case "FillForm2"
                Return "fillformMobile.aspx?ListID=" & CStr(Param1) & "&ID=" & CStr(Param2) & IIf(Len(Param3) > 0, "&Origin=" & System.Web.HttpContext.Current.Server.UrlEncode(CStr(Param3)), "") & "&a=" & Param4 & IIf(Len(Param5) > 0, "&InboxID=" & Param5, "") & IIf(Len(Param6) > 0, "&Actions=" & Param6, "") & IIf(Len(Param7) > 0, "&NoFRM=" & Param7, "") & IIf(Len(Param8) > 0, "&OPID=" & Param8, "") & GetQSRemains(Param9)
            Case "CustomFillForm2"
                Return CStr(Param1) & "?ListID=" & CStr(Param2) & "&ID=" & CStr(Param3)
            Case "FillForm2RO"
                Return "fillformMobile.aspx?FT=1&RO=1&ListID=" & CStr(Param1) & "&ID=" & CStr(Param2) & "&a=" & CStr(Param3) & GetQSRemains(Param9)
            Case "SearchPage"
                Return "searchform.aspx?a=" & Param1 & "&Category=" & Param2 & "&Workflow=" & Param3 & "&ListID=" & Param4 & "&ST=" & Param5 & "&TargetViewID=" & Param6 & IIf(System.Web.HttpContext.Current.Request.QueryString("NoFrm") = "1", "&NoFrm=1", "")
            Case "Search"
                Return "search.aspx?Category=aecafa58-395F-42F5-99ad-f3fe3d26affe&Workflow=" & Param1 & "&ListID=" & Param2 & "&a=" & Param3
            Case "SearchForm"
                Return "searchform.aspx?Category=aecafa58-395F-42F5-99ad-f3fe3d26affe&Workflow=" & Param1 & "&ListID=" & Param2 & "&a=" & Param3 & GetQSRemains(Param4) & IIf(System.Web.HttpContext.Current.Request.QueryString("NoFrm") = "1", "&NoFrm=1", "")
            Case "SearchView"
                Return "searchform.aspx?Category=aecafa58-395F-42F5-99ad-f3fe3d26affe&Workflow=&TargetViewID=" & Param1 & "&ListID=" & Param2 & "&a=" & Param3
            Case "LaunchView"
                Return "FormListMobile.aspx?ViewID=" & Param1 & "&a=" & Param2
            Case "LaunchCalView"
                Return "CalList.aspx?ViewID=" & Param1 & "&a=" & Param2
            Case "LaunchScheduleView"
                Return "SchedulerList.aspx?ViewID=" & Param1 & "&a=" & Param2
            Case "LaunchImageMapView"
                Return "ImageMapList.aspx?ViewID=" & Param1 & "&a=" & Param2
            Case "LaunchChartView"
                Return "ChartList.aspx?ViewID=" & Param1 & "&a=" & Param2
            Case "LaunchTreeView"
                Return "TreeviewList.aspx?ViewID=" & Param1 & "&a=" & Param2
            Case "FillForm"
                Return "fillformMobile.aspx?FT=1&ListID=" & Param1 & "&a=" & Param2 & GetQSRemains(Param9)
            Case "CustomFillForm"
                Return Param1 & "?FT=1&ListID=" & Param2
            'Case "FormList"
            '    Return "FormListMobile.aspx?ViewID=" & Param1 & "&a=" & Param2 & GetQSRemains(Param9)
            Case "FormList"
                Return "applicationmobile.aspx?a=" & Param2
            Case "FillForm2"
                Return "fillformMobile.aspx?FT=1&ListID=" & Param1 & "&ID=" & Param2 & "&a=" & Param3
            Case "FillFormEdit"
                Return "fillformMobile.aspx?FT=1&ListID=" & Param1 & "&ID=" & Param2 & "&Origin=" & System.Web.HttpContext.Current.Server.UrlEncode(Param3) & "&a=" & Param4 & IIf(Len(Param5) > 0, "#JumpHere", "")
            Case "FillFormAmend"
                Return "fillformMobile.aspx?Amend=1&FT=1&ListID=" & Param1 & "&ID=" & Param2 & "&Origin=" & System.Web.HttpContext.Current.Server.UrlEncode(Param3) & "&a=" & Param4 & IIf(Len(Param5) > 0, "#JumpHere", "")
            Case "FillFormEditAction"
                Return "fillformMobile.aspx?FT=1&ListID=" & Param1 & "&ID=" & Param2 & "&Origin=" & System.Web.HttpContext.Current.Server.UrlEncode(Param3) & "&a=" & Param4 & "&InboxID=" & Param6 & "&Actions=" & System.Web.HttpContext.Current.Server.UrlEncode(Param7) & IIf(Len(Param5) > 0, "#JumpHere", "")
            Case "Application"
                Return "application.aspx?ID=" & Param1
            Case "MainHandler"
                Return "~\weberrorMobile.aspx?err=Unhandled"
            Case "Reassign"
                Dim strParam1 As String = Param1
                Return "reassignMobile.aspx?a=" & Param2 & "&ID=" & strParam1
            Case "SubmissionView"
                Dim strParam1 As String = Param1
                Return "SubmissionMobile.aspx?FT=1&ID=" & strParam1 & "&a=" & Param2 & "&acid=" & Param9
            Case "Inbox"
                Return "Default.aspx?a=" & Param1
            Case "MySubmissions"
                Return "start.aspx?a=" & Param1
            Case "DesignWorkflow"
                Dim strParam1 As String = Param1
                Return "designworkflow.aspx?ID=" & strParam1
            Case "Approve"
                Dim strParam1 As String = Param1
                Return "actionMobile.aspx?a=" & Param2 & "&Action=Approve&ID=" & strParam1
            Case "StandardAction"
                Dim strParam1 As String = Param1
                Return "actionMobile.aspx?a=" & Param2 & "&Action=" & Param3 & "&ID=" & strParam1
            Case "Correction"
                Dim strParam1 As String = Param1
                Return "actionMobile.aspx?Action=Correction&ID=" & strParam1
            Case "CustomAction"
                Dim strParam1 As String = Param1
                Dim strParam2 As String = Param2
                Return "actionMobile.aspx?a=" & Param3 & "&Action=Custom&CustomName=" & strParam2 & "&ID=" & strParam1
            Case "Reject"
                Dim strParam1 As String = Param1
                Return "actionMobile.aspx?a=" & Param2 & "&Action=Reject&ID=" & strParam1
            Case "WorkflowStatus"
                Dim strParam1 As String = Param1
                Dim strParam2 As String = Param2
                Dim strParam3 As String = Param3
                Dim strParam4 As String = ""
                If Len(Param4) > 0 Then
                    strParam4 = "&Typ=" & Param4
                End If
                If Len(strParam3) > 0 Then
                    Return "workflowstatusMobile.aspx?a=" & Param5 & "&noframe=1&ID=" & strParam1 & "&TaskID=" & strParam2 & strParam4
                Else
                    Return "workflowstatusMobile.aspx?a=" & Param5 & "&ID=" & strParam1 & "&TaskID=" & strParam2 & strParam4
                End If
            Case "DebugMessagesSN"
                Dim strParam1 As String = Param1
                Return "debugmessages.aspx?SN=" & strParam1
            Case "ActionScreen"
                Dim strParam1 As String = Param1
                Return "actionscreenMobile.aspx?FT=1&ID=" & strParam1 & "&a=" & Param2
            Case "CancelSubmission"
                Dim strParam1 As String = Param1
                Return "SubmissionMobile.aspx?FT=1&CS=1&ID=" & strParam1 & "&a=" & Param2
            Case "Workflows"
                Return "workflows.aspx"
            Case "ListData"
                Dim strParam1 As String = Param1
                Return "listdata.aspx?ListID=" & strParam1
            Case "ListConfig"
                Dim strParam1 As String = Param1
                Return "list.aspx?ID=" & strParam1
            Case "FillListEdit"
                Dim strParam1 As String = Param1
                Dim strParam2 As String = Param2
                Dim strParam3 As String = Param3
                Dim strParam4 As String = Param4
                If Len(strParam4) > 0 Then
                    strParam4 = "&Src=" & strParam4
                End If
                If Len(strParam3) > 0 Then
                    Return "FillList.aspx?noframe=1&ListID=" & strParam1 & "&ID=" & strParam2 & strParam4
                Else
                    Return "FillList.aspx?ListID=" & strParam1 & "&ID=" & strParam2 & strParam4
                End If
            Case "PostMessage"
                Dim strParam1 As String = Param1
                Dim strParam2 As String = Param2
                Dim strParam3 As String = Param3
                If Len(strParam2) > 0 Then
                    Return "postmessageMobile.aspx?a=" & Param4 & "&FS=" & strParam3 & "&ReplyTo=" & strParam2 & "&ID=" & strParam1
                Else
                    Return "postmessageMobile.aspx?a=" & Param4 & "&FS=" & strParam3 & "&ID=" & strParam1
                End If

            Case Else
                Return ""
        End Select
    End Function

    Public Shared Function GetQSRemains(QS As String) As String
        Dim ExcludeTags As String = "id,origin,listid,a,parentrec,new,ft,inboxid,actions,nofrm,viewid,opid,ro"
        Dim _qs As String = ""

        Dim arrTags() As String = Split(ExcludeTags, ",")
        Dim exc As New Collection
        Dim j As Integer
        For j = 0 To UBound(arrTags)
            If exc.Contains(LCase(arrTags(j))) = False Then
                exc.Add(LCase(arrTags(j)), LCase(arrTags(j)))
            End If
        Next




        Dim arrMains() As String = Split(QS, "&")
        Dim i As Integer
        For i = 0 To UBound(arrMains)
            Dim arrpair() As String = Split(arrMains(i), "=")
            If UBound(arrpair) = 1 Then
                Dim ftag As String = LCase(arrpair(0))
                Dim fval As String = arrpair(1)

                If exc.Contains(ftag) = False Then
                    'is not part of exclude list,so we add to querystring
                    If Len(_qs) > 0 Then _qs += "&"
                    _qs += arrpair(0) & "=" & fval
                End If
            End If
        Next
        If Len(_qs) = 0 Then
            Return ""
        Else
            Return "&" & _qs
        End If


    End Function

    Public Shared Sub Navigate(ByVal TargetLocation As String, Optional ByVal Param1 As Object = Nothing, Optional ByVal Param2 As Object = Nothing, Optional ByVal Param3 As Object = Nothing, Optional ByVal Param4 As Object = Nothing, Optional ByVal Param5 As Object = Nothing, Optional ByVal Param6 As Object = Nothing, Optional ByVal Param7 As Object = Nothing, Optional ByVal Param8 As Object = Nothing, Optional ByVal param9 As Object = Nothing, Optional ByVal param10 As Object = Nothing)
        Try
            logger.Debug("TargetLocation: " + GlobalFunctions.FormatData(TargetLocation) + ", Param1: " + GlobalFunctions.FormatData(Param1))
            'logger.Debug("Response Cookies: " + Newtonsoft.Json.JsonConvert.SerializeObject(Current.Response.Cookies))
            If (GlobalFunctions.FromConfig("Debug_StopResponseRedirect") = "true") Then
                'Current.Response.End()
                Exit Sub
            End If

            Select Case TargetLocation
                Case "MobilePages"
                    Current.Response.Redirect("mobilepages.aspx")
                Case "Libraries"
                    Current.Response.Redirect("libraries.aspx?AppID=" & Param1)
                Case "NewLibrary"
                    Current.Response.Redirect("library.aspx?AppID=" & Param1)
                Case "NewMobilePage"
                    Current.Response.Redirect("mobilepage.aspx")
                Case "ReusablePieces"
                    Current.Response.Redirect("ReusablePieces.aspx")
                Case "ReqGathering"
                    Current.Response.Redirect("EditReqGathering.aspx?ListItemID=" & Param1 & "&ID=" & Param2 & "&AppID=" & Param3)
                Case "EditAppWizard"
                    Current.Response.Redirect("CreateApp.aspx?ID=" & Param1 & "&Tab=LookNFeel")
                Case "WorkflowStgs"
                    Current.Response.Redirect("workflowgeneral.aspx?ID=" & Param1)
                Case "Navlistitem"
                    Current.Response.Redirect("navlistitem.aspx?AppID=" & Param1)
                Case "Searchform"
                    Current.Response.Redirect("searchform.aspx?a=" & Param1 & "&Category=" & Param2 & "&Workflow=" & Param3 & "&ListID=" & Param4 & "&ST=" & Param5)
                Case "NewSyncProfile"
                    Current.Response.Redirect("syncprofile.aspx")
                Case "SyncProfiles"
                    Current.Response.Redirect("runsync.aspx")
                Case "Plugins"
                    Current.Response.Redirect("plugins.aspx")
                Case "NewPlugin"
                    Current.Response.Redirect("plugin.aspx")
                Case "CustomAction"
                    Dim strParam1 As String = Param1
                    Dim strParam2 As String = Param2
                    Current.Response.Redirect("actionMobile.aspx?a=" & Param3 & "&Action=Custom&CustomName=" & strParam2 & "&ID=" & strParam1)
                Case "ChangeHistory"
                    Current.Response.Redirect("AuditTrailViewer.aspx?a=" & CStr(Param1) & "&ListID=" & Param2 & "&SQ=" & Param3 & "&ReturnURL=" & Param4)
                Case "AppInfo"
                    Current.Response.Redirect("appinfo.aspx?ID=" & CStr(Param1))
                Case "AppShop"
                    Current.Response.Redirect("appshopitem.aspx?ID=" & CStr(Param1))
                Case "submission"
                    Current.Response.Redirect("SubmissionMobile.aspx?ID=" & CStr(Param1) & "&a=" & Param2 & "&FT=1")
                Case "ImportAsCSV"
                    Current.Response.Redirect("importascsv.aspx?a=" & Param1 & "&ViewID=" & Param2 & "&acid=" & Param3)
                Case "AppLists"
                    Current.Response.Redirect("applists.aspx?AppID=" & Param1 & "&Type=" & Param2)
                Case "SelfSettingsApp"
                    Current.Response.Redirect("selfsettings.aspx?a=" & Param1 & IIf(Len(Param2) > 0, "&mode=", "") & CStr(Param2))
                Case "SelfSettings"
                    Current.Response.Redirect("selfsettings.aspx")
                Case "ViewActionButton"
                    Current.Response.Redirect("ViewActionButton.aspx?AppID=" & Param1 & "&ViewID=" & Param2 & "&ID=" & Param3)
                Case "ViewActions"
                    Current.Response.Redirect("ViewActions.aspx?AppID=" & Param1 & "&ID=" & Param2)
                Case "NewDatasourceConnection"
                    Current.Response.Redirect("Datasourceconnection.aspx?AppID=" & Param1)
                Case "DatasourceConnections"
                    Current.Response.Redirect("Datasourceconnections.aspx?AppID=" & Param1)
                Case "SetPermissions"
                    Current.Response.Redirect("setpermissions.aspx?AppID=" & Param1 & "&ID=" & Param2)
                Case "NewWebservice"
                    Current.Response.Redirect("webservice.aspx?AppID=" & Param1)
                Case "Webservices"
                    Current.Response.Redirect("webservices.aspx?AppID=" & Param1)
                Case "ExternalConns"
                    Current.Response.Redirect("externalconns.aspx?AppID=" & Param1)
                Case "NewExternalConn"
                    Current.Response.Redirect("externalconn.aspx?AppID=" & Param1)
                Case "SubstitutionString"
                    Current.Response.Redirect("substitutionstring.aspx?AppID=" & Param1)
                Case "SubstitutionStrings"
                    Current.Response.Redirect("substitutionstrings.aspx?AppID=" & Param1)
                Case "Default"
                    Dim _page As String = GlobalFunctions.GrabPostLoginAppPage(Param1)
                    Current.Response.Redirect(_page)
                Case "ManageUsers"
                    Current.Response.Redirect("manageusers.aspx?AppID=" & Param1)
                Case "AddAppUser"
                    Current.Response.Redirect("AddUser.aspx?AppID=" & Param1)
                Case "AddAppDev"
                    Current.Response.Redirect("AddUser.aspx?Type=Dev&AppID=" & Param1)
                Case "EditCompositePageConfig"
                    Current.Response.Redirect("compositepageconfig.aspx?AppID=" & Param1 & "&ID=" & Param2)
                Case "AppTab"
                    Current.Response.Redirect("apptab.aspx?AppID=" & Param1)
                Case "AppTabs"
                    Current.Response.Redirect("apptabs.aspx?AppID=" & Param1)
                Case "Archival"
                    Current.Response.Redirect("archival.aspx?AppID=" & Param1 & "&ID=" & Param2)
                Case "ArchivalRules"
                    Current.Response.Redirect("archiving.aspx?AppID=" & Param1 & "&ID=" & Param2)
                Case "GlobalImage"
                    Current.Response.Redirect("globalimage.aspx")
                Case "ResourceFile"
                    Current.Response.Redirect("resourcefile.aspx?AppID=" & Param1)
                Case "AppLogin"
                    Current.Response.Redirect("login.aspx?AppID=" & Param1 & IIf(Len(Param2) > 0, "&PMode=" & Param2, "") & IIf(Len(GlobalFunctions.FormatData(Param3)) > 0, "&msg=" & HttpUtility.UrlEncode(GlobalFunctions.FormatData(Param3)), ""))
                Case "AppDefault"
                    Current.Response.Redirect("Default.aspx?a=" & Param1)
                Case "ResourceFiles"
                    Current.Response.Redirect("resourcefiles.aspx?AppID=" & Param1)
                Case "GlobalImages"
                    Current.Response.Redirect("globalimages.aspx")
                Case "Themes"
                    Current.Response.Redirect("themes.aspx")
                Case "EditTheme"
                    Current.Response.Redirect("newtheme.aspx?AppID=" & Param1 & "&ID=" & Param2)
                Case "NewTheme"
                    Current.Response.Redirect("newtheme.aspx?AppID=" & Param1)
                Case "ChangeLang"
                    Current.Response.Redirect("changelang.aspx?AppID=" & Param1 & "&ID=" & Param2)
                Case "AggregateDesign"
                    Current.Response.Redirect("AggregateDesigner.aspx?AppID=" & Param1 & "&ID=" & Param2)
                Case "EditAppConfigPage"
                    Current.Response.Redirect("AppConfigPage.aspx?ID=" & Param1)
                Case "LaunchViewSettings"
                    Current.Response.Redirect("viewconfigpage.aspx?AppID=" & CStr(Param1) & "&ID=" & CStr(Param2) & "&acid=" & Param3)
                Case "AppPage"
                    Current.Response.Redirect("application.aspx?ID=" & Param1)
                Case "AppTables"
                    Current.Response.Redirect("apptables.aspx")
                Case "EditFormPageConfig"
                    Current.Response.Redirect("EFormConfigPage.aspx?AppID=" & Param1 & "&ID=" & Param2)
                Case "EditFormConfigs"
                    If Len(Param3) = 0 Or StrComp(Param3, "IM", vbTextCompare) = 0 Then
                        Current.Response.Redirect("FormConfigs.aspx?ListID=" & Param1 & "&AppID=" & Param2)
                    Else
                        Current.Response.Redirect("FormDeviceConfigs.aspx?ListID=" & Param1 & "&AppID=" & Param2)
                    End If
                Case "EditMobileFormConfigs"
                    Current.Response.Redirect("FormDeviceConfigs.aspx?ListID=" & Param1 & "&AppID=" & Param2)
                Case "EditFormAdvSettings"
                    Current.Response.Redirect("EditFormBehavior.aspx?ListID=" & Param1 & "&AppID=" & Param2)
                Case "EditMobileFormAdvSettings"
                    Current.Response.Redirect("EditMobileFormBehavior.aspx?ListID=" & Param1 & "&AppID=" & Param2)
                Case "EditMobileFormAdvSettings2"
                    Current.Response.Redirect("EditMobileFormBehavior.aspx?ListID=" & Param1 & "&AppID=" & Param2 & "&ID=" & Param3 & Param4)
                Case "EditFormField"
                    Current.Response.Redirect("EditFormField.aspx?ListItemID=" & Param1 & "&ID=" & Param2 & "&AppID=" & Param3)
                Case "CalList"
                    Current.Response.Redirect("CalList.aspx?ViewID=" & Param1)
                Case "AskToGenView"
                    Current.Response.Redirect("asktogenview.aspx?ID=" & CStr(Param1))
                Case "NewFillForm"
                    Current.Response.Redirect("fillformMobile.aspx?ListID=" & CStr(Param1) & "&FT=1&a=" & Param2 & IIf(Len(Param3) > 0, "&", "") & Param3)
                Case "NewFillFormWithDefaults"
                    Current.Response.Redirect("fillformMobile.aspx?Def=" & Param3 & "&ListID=" & CStr(Param1) & "&FT=1&a=" & Param2)
                Case "FillForm"
                    Current.Response.Redirect("fillformMobile.aspx?ListID=" & CStr(Param1) & "&ParentRec=" & CStr(Param2) & IIf(Len(Param3) > 0, "&Origin=" & CStr(Param3), "") & "&a=" & Param4 & IIf(Len(Param5) > 0, "&InboxID=" & Param5, "") & IIf(Len(Param6) > 0, "&Actions=" & Param6, "") & IIf(Len(Param7) > 0, "&NoFRM=" & Param7, "") & IIf(Len(Param8) > 0, "&New=" & Param8, "") & GetQSRemains(param10))
                Case "FillFormEdit"
                    Current.Response.Redirect("fillformMobile.aspx?FT=1&ListID=" & CStr(Param1) & "&ID=" & CStr(Param2) & "&a=" & Param3 & GetQSRemains(param9))
                Case "FillFormRO"
                    Current.Response.Redirect("fillformMobile.aspx?RO=1&FT=1&ListID=" & CStr(Param1) & "&ID=" & CStr(Param2) & "&a=" & Param3 & "&NoFRM=" & Param4 & GetQSRemains(param9))
                'Case "FormList"
                '    Current.Response.Redirect("FormListMobile.aspx?ViewID=" & CStr(Param1) & "&Sort=" & Param2 & "&Dir=" & Param3)
                Case "FormList"
                    Current.Response.Redirect("applicationmobile.aspx?a=" & CStr(Param1))

                Case "LaunchView"
                    Current.Response.Redirect("FormListMobile.aspx?ViewID=" & CStr(Param1) & "&a=" & Param2 & "&acid=" & Param3)
                Case "FillForm2"
                    Current.Response.Redirect("fillformMobile.aspx?ListID=" & CStr(Param1) & "&ID=" & CStr(Param2) & IIf(Len(Param3) > 0, "&Origin=" & CStr(Param3), "") & "&a=" & Param4 & IIf(Len(Param5) > 0, "&InboxID=" & Param5, "") & IIf(Len(Param6) > 0, "&Actions=" & Param6, "") & IIf(Len(Param7) > 0, "&NoFRM=" & Param7, "") & IIf(Len(Param8) > 0, "&OPID=" & Param8, "") & GetQSRemains(param9))
                Case "FillForm5"
                    Current.Response.Redirect("fillformMobile.aspx?ListID=" & CStr(Param1) & "&ParentRec=" & CStr(Param2) & IIf(Len(Param3) > 0, "&Origin=" & CStr(Param3), "") & "&a=" & Param4 & IIf(Len(Param5) > 0, "&InboxID=" & Param5, "") & IIf(Len(Param6) > 0, "&Actions=" & Param6, "") & IIf(Len(Param7) > 0, "&NoFRM=" & Param7, "") & IIf(Len(Param8) > 0, "&OPID=" & Param8, "") & GetQSRemains(param9))
                Case "CustomFillForm2"
                    Current.Response.Redirect(CStr(Param1) & "?ListID=" & CStr(Param2) & "&ID=" & CStr(Param3))
                Case "NewView"
                    Current.Response.Redirect("viewsettings.aspx?AppID=" & CStr(Param1))
                Case "EditView"
                    Current.Response.Redirect("viewsettings.aspx?AppID=" & CStr(Param1) & "&ID=" & CStr(Param2) & "&Tab=" & Param3)
                Case "submitfile"
                    Dim strParam1 As String = Param1
                    Dim strParam2 As String = Param2
                    Current.Response.Redirect("submitfile.aspx?listitemid=" & CStr(strParam1) & "&wid=" & CStr(strParam2))
                Case "SubmissionSuccess"
                    Current.Response.Redirect("submissionsuccess.aspx?ID=" & CStr(Param1))
                Case "CheckWorkflow"
                    Current.Response.Redirect("checkworkflow.aspx?ID=" & CStr(Param1) & "&Latest=1&a=" & CStr(Param2) & "&AppID=" & CStr(Param2) & "&ListID=" & CStr(Param3))
                Case "AppChangePassword"
                    Current.Response.Redirect("changepassword.aspx?a=" & CStr(Param1) & IIf(Len(Param2) > 0, "&mode=", "") & CStr(Param2))
                Case "ChangePassword"
                    Current.Response.Redirect("changepassword.aspx")
                Case "CustomizeActionsWindow"
                    Dim strParam1 As String = Param1
                    Current.Response.Redirect("CustomizeActionsWindow.aspx?ID=" & strParam1)
                Case "DLoad"
                    Dim strParam1 As String = Param1
                    Dim strParam2 As String = Param2
                    Current.Response.Redirect("DLoad.aspx?ID=" & strParam1 & "&Mode=" & strParam2)
                Case "ActionScreen"
                    Dim strParam1 As String = Param1
                    Current.Response.Redirect("actionscreenMobile.aspx?FT=1&ID=" & strParam1 & "&a=" & Param2)
                Case "ActionScreenFull"
                    Dim strParam1 As String = Param1
                    Current.Response.Redirect("actionscreenMobile.aspx?FT=1&ID=" & strParam1 & "&a=" & Param2 & "&ID2=" & Param3)
                Case "MFPInstances"
                    Current.Response.Redirect("mfpinstances.aspx")
                Case "OutOfOffices"
                    Current.Response.Redirect("outofoffices.aspx")
                Case "AppOutOfOffices"
                    Current.Response.Redirect("outofoffices2.aspx?a=" & Param1)
                Case "Applications"
                    Current.Response.Redirect("applications.aspx")
                Case "NewApplication"
                    Current.Response.Redirect("CreateApp.aspx")
                Case "Application"

                    If Len(Param2) = 0 Then
                        Current.Response.Redirect("CreateApp.aspx?ID=" & Param1)
                    Else
                        Current.Response.Redirect("CreateApp.aspx?ID=" & Param1 & "&Tab=" & Param2)
                    End If
                Case "OutOfOffice"
                    Current.Response.Redirect("outofoffice.aspx")
                Case "OutOfOffice2"
                    Current.Response.Redirect("outofoffice2.aspx?a=" & Param1)
                Case "ActionScreenAuto"
                    Dim strParam1 As String = Param1
                    Dim strParam2 As String = Param2
                    Dim strParam3 As String = Param3
                    Current.Response.Redirect("actionscreen.aspx?FT=1&ID=" & strParam1 & "&ID2=" & strParam2 & IIf(Len(strParam3) > 0, "&Mode=" & strParam3, ""))
                Case "ActionScreenSubAuto"
                    Dim strParam1 As String = Param1
                    Dim strParam2 As String = Param2
                    Dim strParam3 As String = Param3
                    Dim strParam4 As String = Param4
                    Dim strParam5 As String = Param5
                    Current.Response.Redirect("actionscreen.aspx?FT=1&ID=" & strParam1 & "&ID2=" & strParam2 & IIf(Len(strParam3) > 0, "&Mode=" & strParam3, "") & IIf(Len(strParam4) > 0, "&ListID=" & strParam4, "") & IIf(Len(strParam5) > 0, "&ParentRec=" & strParam5, ""))
                Case "CustomSource"
                    Current.Response.Redirect(Param1)
                Case "NewForm"
                    Current.Response.Redirect("eform.aspx?AppID=" & Param1)
                Case "Search"
                    Dim strParam1 As String = Param1
                    Current.Response.Redirect("Search.aspx")
                Case "DebugMessagesSN"
                    Dim strParam1 As String = Param1
                    Current.Response.Redirect("DebugMessages.aspx?SN=" & strParam1)
                Case "CustomizeActionsWindow"
                    Dim strParam1 As String = Param1
                    Current.Response.Redirect("CustomizeActionsWindow.aspx?ID=" & strParam1)
                Case "Home"
                    Current.Response.Redirect("Default.aspx")
                Case "MyApps"
                    If Len(WebconfigSettings.HomeURL) > 0 Then
                        Current.Response.Redirect(WebconfigSettings.HomeURL)
                    Else
                        Current.Response.Redirect("myapplications.aspx")
                    End If
                Case "EditForm"
                    Current.Response.Redirect("EForm.aspx?DM=" & Param4 & "&AppID=" & Param1 & "&ID=" & Param2 & IIf(Len(Param3) > 0, "&Tab=" & Param3, ""))
                Case "TaskInbox"
                    Dim strParam1 As String = Param1
                    Dim strParam2 As String = Param2

                    If Len(strParam1) > 0 Or Len(strParam2) > 0 Then
                        Current.Response.Redirect("Default.aspx?a=" & Param3 & "&CategoryID=" & strParam1 & "&Status=" & strParam2)
                    Else
                        Current.Response.Redirect("Default.aspx?a=" & Param3)
                    End If
                Case "GlobalInbox"
                    Dim strParam1 As String = Param1
                    Dim strParam2 As String = Param2

                    If Len(strParam1) > 0 Or Len(strParam2) > 0 Then
                        Current.Response.Redirect("globalinbox.aspx?Status=" & strParam2)
                    Else
                        Current.Response.Redirect("globalinbox.aspx")
                    End If
                Case "SubmissionSearch"
                    Dim strParam1 As String = Param1
                    Dim strParam2 As String = Param2
                    Dim strParam3 As String = Param3

                    If Len(strParam3) > 0 Then
                        Current.Response.Redirect("search.aspx?SP=1")
                    Else
                        If Len(strParam1) > 0 Or Len(strParam2) > 0 Then
                            If GlobalFunctions.IsGUID(strParam1) Or GlobalFunctions.IsGUID(strParam2) Then
                                Current.Response.Redirect("search.aspx?Category=" & strParam1 & "&Workflow=" & strParam2 & "&ListID=" & strParam2)
                            Else
                                Current.Response.Redirect("search.aspx?Category=" & strParam1 & "&Workflow=" & strParam2)
                            End If
                        Else
                            Current.Response.Redirect("search.aspx")
                        End If
                    End If
                Case "DebugMessages"
                    Dim strParam1 As String = Param1
                    Dim strParam2 As String = Param2

                    If Len(strParam1) > 0 Or Len(strParam2) > 0 Then
                        Current.Response.Redirect("debugmessages.aspx?Category=" & strParam1 & "&Status=" & strParam2)
                    Else
                        Current.Response.Redirect("debugmessages.aspx")
                    End If
                Case "MySubmissions"
                    Dim strParam1 As String = Param1
                    Dim strParam2 As String = Param2

                    If Len(strParam1) > 0 Or Len(strParam2) > 0 Then
                        Current.Response.Redirect("start.aspx?Category=" & strParam1 & "&Status=" & strParam2)
                    Else
                        Current.Response.Redirect("start.aspx")
                    End If
                Case "Login"
                    Dim urlToGo As String = "login.aspx" & IIf(Len(GlobalFunctions.FormatData(Param1)) > 0, "?msg=" & HttpUtility.UrlEncode(GlobalFunctions.FormatData(Param1)), "")
                    Current.Response.Redirect(urlToGo)
                Case "Users"
                    Current.Response.Redirect("users.aspx")
                Case "Departments"
                    Current.Response.Redirect("departments.aspx")
                Case "NewDepartment"
                    Current.Response.Redirect("department.aspx")
                Case "PublicHolidays"
                    Current.Response.Redirect("PublicHolidays.aspx")
                Case "NewPublicHoliday"
                    Current.Response.Redirect("PublicHoliday.aspx")
                Case "NewUser"
                    Current.Response.Redirect("user.aspx")
                Case "UserEdit"
                    Dim strParam1 As String = Param1
                    Current.Response.Redirect("user.aspx?ID=" & strParam1)
                Case "Settings"
                    Current.Response.Redirect("settings.aspx")
                Case "GroupEdit"
                    Dim strParam1 As String = Param1
                    Current.Response.Redirect("role.aspx?ID=" & strParam1)
                Case "Groups"
                    Current.Response.Redirect("roles.aspx")
                Case "NewRole"
                    Current.Response.Redirect("role.aspx")
                Case "NoticeEdit"
                    Dim strParam1 As String = Param1
                    Current.Response.Redirect("Notice.aspx?ID=" & strParam1)
                Case "Notices"
                    Current.Response.Redirect("Notices.aspx")
                Case "NewNotice"
                    Current.Response.Redirect("Notice.aspx")
                Case "NewListFromWorkflow"
                    Dim strParam1 As String = Param1
                    Current.Response.Redirect("list.aspx?Type=WFlow&WID=" & strParam1)
                Case "Lists"
                    Current.Response.Redirect("lists.aspx")
                Case "EditList"
                    Dim strParam1 As String = Param1
                    Current.Response.Redirect("list.aspx?ID=" & strParam1)
                Case "NewWorkflowList"
                    Dim strParam1 As String = Param1
                    Current.Response.Redirect("list.aspx?WID=" & strParam1)
                Case "EditListWFlow"
                    Dim strParam1 As String = Param1
                    Dim strParam2 As String = Param2
                    Current.Response.Redirect("list.aspx?Type=WFlow&ID=" & strParam1 & "&WID=" & strParam2)
                Case "EmailTemplates"
                    Current.Response.Redirect("emailtemplates.aspx")
                Case "ArrowDetailsEdit"
                    Dim strParam1 As String = Param1
                    Dim strParam2 As String = Param2
                    Current.Response.Redirect("arrowdetails.aspx?WorkflowID=" & strParam1 & "&ArrowID=" & strParam2)
                Case "ArrowDetailsNew"
                    Dim strParam1 As String = Param1
                    Current.Response.Redirect("arrowdetails.aspx?WorkflowID=" & strParam1)
                Case "WorkflowVersionEdit"
                    Dim strParam1 As String = Param1
                    Current.Response.Redirect("workflowversion.aspx?Latest=1&ID=" & strParam1)
                Case "DesignWorkflow"
                    Dim strParam1 As String = Param1
                    Dim strParam2 As String = Param2
                    If Len(strParam2) > 0 Then
                        strParam2 = "&I=" + strParam2
                    End If
                    Current.Response.Redirect("designworkflow.aspx?ID=" & strParam1 & strParam2)
                Case "DesignWorkflowVersionCreated"
                    Dim strParam1 As String = Param1
                    Current.Response.Redirect("designworkflow.aspx?ID=" & strParam1 & "&NV=1")
                Case "Workflows"
                    Dim strParam1 As String = Param1
                    Current.Response.Redirect("workflows.aspx?CategoryID=" & strParam1)
                Case "WorkflowProperties"
                    Dim strParam1 As String = Param1
                    Current.Response.Redirect("workflow.aspx?ID=" & strParam1)
                Case "WorkflowsAllCat"
                    Current.Response.Redirect("workflows.aspx")
                Case "ConfigureSharing"
                    Dim strParam1 As String = Param1
                    Dim strParam2 As Integer = Param2
                    Current.Response.Redirect("configuresharing.aspx?ID=" & strParam1 & "&Type=" & strParam2)
                Case "WorkflowCategories"
                    Current.Response.Redirect("workflowcategories.aspx")
                Case "NewWorkflowCategory"
                    Current.Response.Redirect("workflowcategory.aspx")
                Case "EditWorkflowCategory"
                    Dim strParam1 As String = Param1
                    Current.Response.Redirect("workflowcategory.aspx?ID=" & strParam1)
                Case "NewWorkflowCategoryFromWorkflow"
                    Current.Response.Redirect("workflowcategory.aspx?FW=1")
                Case "NewWorkflow"
                    Current.Response.Redirect("workflow.aspx?type=documents")
                Case "ListData"
                    Dim strParam1 As String = Param1
                    Current.Response.Redirect("listdata.aspx?ListID=" & strParam1)
                Case "ExportWorkflows"
                    Dim strParam1 As String = Param1
                    Current.Response.Redirect("exportworkflow.aspx?ID=" & strParam1)
                Case "ListDataNoFrame"
                    Dim strParam1 As String = Param1
                    Dim strParam2 As String = Param2
                    Current.Response.Redirect("listdata.aspx?ListID=" & strParam1 & "&NewID=" & strParam2)
                Case "NewListEntry"
                    Dim strParam1 As String = Param1
                    Current.Response.Redirect("FillList.aspx?ListID=" & strParam1)
                Case "NewEmailTemplate"
                    Current.Response.Redirect("EmailTemplate.aspx")
                Case "NewListEntryNoFrame"
                    Dim strParam1 As String = Param1
                    Current.Response.Redirect("FillList.aspx?noframe=1&ListID=" & strParam1)
                Case "SendNotification"
                    Dim strParam1 As String = Param1
                    Current.Response.Redirect("ConfigureSendNotification.aspx?WorkflowID=" & strParam1)
                Case "ApprovalConfig"
                    Dim strParam1 As String = Param1
                    Current.Response.Redirect("ConfigureApproval.aspx?WorkflowID=" & strParam1)
                Case "ApprovalConfigEdit"
                    Dim strParam1 As String = Param1
                    Dim strParam2 As String = Param2
                    Current.Response.Redirect("ConfigureApproval.aspx?WorkflowID=" & strParam1 & "&ID=" & strParam2)
                Case "RunCodeConfig"
                    Dim strParam1 As String = Param1
                    Current.Response.Redirect("ConfigureRunCode.aspx?WorkflowID=" & strParam1)
                Case "SaveToFT8ConfigEdit"
                    Dim strParam1 As String = Param1
                    Dim strParam2 As String = Param2
                    Current.Response.Redirect("ConfigureFT8.aspx?WorkflowID=" & strParam1 & "&ID=" & strParam2)
                Case "SaveAttachedFiles"
                    Dim strParam1 As String = Param1
                    Current.Response.Redirect("ConfigureStorage.aspx?WorkflowID=" & strParam1)
                Case "RunCodeConfigEdit"
                    Dim strParam1 As String = Param1
                    Dim strParam2 As String = Param2
                    Current.Response.Redirect("ConfigureRunCode.aspx?WorkflowID=" & strParam1 & "&ID=" & strParam2)
                Case "SaveToFileConfigEdit"
                    Dim strParam1 As String = Param1
                    Dim strParam2 As String = Param2
                    Current.Response.Redirect("ConfigureStorage.aspx?WorkflowID=" & strParam1 & "&ID=" & strParam2)
                Case "InstanceViewer"
                    Dim strParam1 As String = Param1
                    Current.Response.Redirect("SubmissionMobile.aspx?FT=1&noframe=1&ID=" & strParam1)
                Case "SendEmail"
                    Dim strParam1 As String = Param1
                    Current.Response.Redirect("ConfigureSendEmail.aspx?WorkflowID=" & strParam1)
                Case "SendEmailEdit"
                    Dim strParam1 As String = Param1
                    Dim strParam2 As String = Param2
                    Current.Response.Redirect("ConfigureSendEmail.aspx?WorkflowID=" & strParam1 & "&ID=" & strParam2)
                Case "SubmissionView"
                    Dim strParam1 As String = Param1
                    Current.Response.Redirect("SubmissionMobile.aspx?FT=1&ID=" & strParam1 & "&a=" & Param2)
                Case "SubmissionViewPC"
                    Dim strParam1 As String = Param1
                    Current.Response.Redirect("SubmissionMobile.aspx?FT=1&PC=1&ID=" & strParam1 & "&a=" & Param2)
                Case "OutgoingMails"
                    Current.Response.Redirect("OutgoingEmails.aspx")
                Case "SharpdeskUsers"
                    Current.Response.Redirect("SharpdeskUsers.aspx")
                Case "QFUpload"
                    Current.Server.Transfer("QFUpload.aspx", True)
                Case "CustomRedirect"
                    Dim strParam1 As String = Param1
                    Current.Response.Redirect(strParam1)
                Case "ViewChart"
                    Dim strParam1 As String = Param1
                    Current.Response.Redirect("Viewchart.aspx?" & strParam1)
                Case "MasterViewChart"
                    Dim strParam1 As String = Param1
                    Current.Response.Redirect("MasterViewchart.aspx?" & strParam1)
                Case "Logout"
                    GlobalFunctions.ZukamiLogout()
                    Dim appId As String = HttpContext.Current.Request.QueryString("AppID")
                    If String.IsNullOrWhiteSpace(appId) Then
                        appId = HttpContext.Current.Request.QueryString("a")
                    End If
                    If String.IsNullOrWhiteSpace(appId) Then
                        Navigate("Login", Param1)
                    Else
                        Navigate("AppLogin", appId, "", Param1)
                    End If
            End Select
        Catch ex As ThreadAbortException
            logger.Debug("System.Threading.ThreadAbortException: due to url redirect")
        Catch ex As ApplicationException
            logger.Debug(ex)
            If "Logout,Login,AppLogin".Contains(TargetLocation) Then
                RedirectLocationToPage("Login.aspx")
            Else
                logger.Debug("not sure what to do")
            End If
        Catch ex As Exception
            logger.Error(ex, "failed to redirect")
            Throw
        End Try
    End Sub

    Private Shared Sub RedirectLocationToPage(pageUrl As String)
        logger.Debug("start")
        HttpContext.Current.Response.Status = "302 Found"
        HttpContext.Current.Response.StatusCode = 302
        HttpContext.Current.Response.RedirectLocation = pageUrl
        HttpContext.Current.Response.End()
    End Sub

    Public Shared Sub RaiseWebError(ByVal WebError As String, Optional ByVal Tag As String = "")
        Dim urlParams As String = "err=" & WebError & "&Tag1=" & System.Web.HttpContext.Current.Server.UrlEncode(Tag) & "&AppID=" & GlobalFunctions.GetAppID()
        Dim errorPage As String = "weberrorMobile.aspx"
        If Len(GlobalFunctions.FromConfig("CustomErrorPageMobile")) > 0 Then
            errorPage = GlobalFunctions.FromConfig("CustomErrorPageMobile")
        End If
        Dim delimiter As String = "?"
        If errorPage.Contains("?") Then
            delimiter = "&"
        End If
        Current.Response.Redirect(errorPage & delimiter & urlParams)
    End Sub
End Class

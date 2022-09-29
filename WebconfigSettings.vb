Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.OleDb
Imports System.Configuration.ConfigurationManager
Imports System.DirectoryServices.AccountManagement
Imports NLog

Public Module WebconfigSettings
    Private logger As Logger = LogManager.GetCurrentClassLogger()
    Public strcustombtn As String = ""
    Public ConnectionStringft As String = AppSettings.Get("ConnectionStringft")
    Public ConnectionString As String = GlobalFunctions.GetDecryptedConnectionString() ' IIf(GlobalFunctions.FormatBoolean(AppSettings.Get("EncryptSQLConnection"), False) = True, GlobalFunctions.SimpleCrypt(GlobalFunctions.DecryptBase64(AppSettings.Get("ConnectionString"))), AppSettings.Get("ConnectionString"))
    Public CloudConnectionString As String = AppSettings.Get("CloudConnectionString")
    Public DoctrixConnectionString As String = AppSettings.Get("DoctrixConnectionString")
    Public UploadPath As String = AppSettings.Get("UploadPath")
    Public AppTitle As String = AppSettings.Get("AppTitle")
    Public Queue As String = AppSettings.Get("QueueLocation")
    Public ProxyMode As String = AppSettings.Get("ProxyMode")
    Public Authentication As String = AppSettings.Get("Authentication")
    Public UseAnonymous As Boolean = AppSettings.Get("UseAnonymous")
    Public InactivityTimeout As Integer = AppSettings.Get("InactivityTimeout")
    Public ListTablePrefix As String = AppSettings.Get("ListTablePrefix")
    Public DisableSeekInput As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("DisableSeekInput"), False)
    Public DisableRelayBubble As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("DisableRelayBubble"), False)
    Public UnicodeFriendly As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("UnicodeFriendly"), False)
    Public UseSSOIfAvailable As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("UseSSOIfAvailable"), False)
    Public EnableMobileMode As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("EnableMobileMode"), False)
    Public WorkflowAuditTrailAscending As String = AppSettings.Get("Workflow_AuditTrailASC")
    Public ReassignCaption As String = AppSettings.Get("WorkflowButton_ReassignCaption")
    Public RequestReviewCaption As String = AppSettings.Get("WorkflowButton_RequestReviewCaption")
    Public RequestReviewCaption2 As String = AppSettings.Get("WorkflowButton_RequestReviewCaptionSingleline")
    Public FormSubmissionRemarksCaption As String = AppSettings.Get("Form_SubmissionRemarksCaption")
    Public Portal_OutofofficeDelegateeCaption As String = AppSettings.Get("Portal_OutofofficeDelegateeCaption")
    Public PostAMessage_SendToSubmitterOnly As String = AppSettings.Get("PostAMessage_SendToSubmitterOnly")
    Public PostAMessage_IncludeSubmitter As String = AppSettings.Get("PostAMessage_IncludeSubmitter")
    Public WorkflowHistory_ApproveMsg As String = AppSettings.Get("WorkflowHistory_ApproveMsg")
    Public WorkflowHistory_RejectMsg As String = AppSettings.Get("WorkflowHistory_RejectMsg")
    Public WorkflowHistory_ActionMsg As String = AppSettings.Get("WorkflowHistory_ActionMsg")
    Public ADLDAPPath As String = AppSettings.Get("ADLDAPPath")
    Public ADLDAPPathFilter As String = AppSettings.Get("ADLDAPPathFilter")
    Public ADLDAPPathUsername As String = AppSettings.Get("ADLDAPPathUsername")
    Public ADLDAPPathPassword As String = AppSettings.Get("ADLDAPPathPassword")
    Public ADLDAPPathAuthType As String = AppSettings.Get("ADLDAPPathAuthType")
    Public ADLDAPPath2 As String = AppSettings.Get("ADLDAPPath2")
    Public ADLDAPPath2Filter As String = AppSettings.Get("ADLDAPPath2Filter")
    Public ADLDAPPath2Username As String = AppSettings.Get("ADLDAPPath2Username")
    Public ADLDAPPath2Password As String = AppSettings.Get("ADLDAPPath2Password")
    Public ADLDAPPath2AuthType As String = AppSettings.Get("ADLDAPPath2AuthType")
    Public ADLDAPPath3 As String = AppSettings.Get("ADLDAPPath3")
    Public ADLDAPPath3Filter As String = AppSettings.Get("ADLDAPPath3Filter")
    Public ADLDAPPath3Username As String = AppSettings.Get("ADLDAPPath3Username")
    Public ADLDAPPath3Password As String = AppSettings.Get("ADLDAPPath3Password")
    Public ADLDAPPath3AuthType As String = AppSettings.Get("ADLDAPPath3AuthType")
    Public ADLDAPPath4 As String = AppSettings.Get("ADLDAPPath4")
    Public ADLDAPPath4Filter As String = AppSettings.Get("ADLDAPPath4Filter")
    Public ADLDAPPath4Username As String = AppSettings.Get("ADLDAPPath4Username")
    Public ADLDAPPath4Password As String = AppSettings.Get("ADLDAPPath4Password")
    Public ADLDAPPath4AuthType As String = AppSettings.Get("ADLDAPPath4AuthType")
    Public ADLDAPPath5 As String = AppSettings.Get("ADLDAPPath5")
    Public ADLDAPPath5Filter As String = AppSettings.Get("ADLDAPPath5Filter")
    Public ADLDAPPath5Username As String = AppSettings.Get("ADLDAPPath5Username")
    Public ADLDAPPath5Password As String = AppSettings.Get("ADLDAPPath5Password")
    Public ADLDAPPath5AuthType As String = AppSettings.Get("ADLDAPPath5AuthType")
    Public ADLDAPPath6 As String = AppSettings.Get("ADLDAPPath6")
    Public ADLDAPPath6Filter As String = AppSettings.Get("ADLDAPPath6Filter")
    Public ADLDAPPath6Username As String = AppSettings.Get("ADLDAPPath6Username")
    Public ADLDAPPath6Password As String = AppSettings.Get("ADLDAPPath6Password")
    Public ADLDAPPath6AuthType As String = AppSettings.Get("ADLDAPPath6AuthType")
    Public ADLDAPPath7 As String = AppSettings.Get("ADLDAPPath7")
    Public ADLDAPPath7Filter As String = AppSettings.Get("ADLDAPPath7Filter")
    Public ADLDAPPath7Username As String = AppSettings.Get("ADLDAPPath7Username")
    Public ADLDAPPath7Password As String = AppSettings.Get("ADLDAPPath7Password")
    Public ADLDAPPath7AuthType As String = AppSettings.Get("ADLDAPPath7AuthType")
    Public ADLDAPPath8 As String = AppSettings.Get("ADLDAPPath8")
    Public ADLDAPPath8Filter As String = AppSettings.Get("ADLDAPPath8Filter")
    Public ADLDAPPath8Username As String = AppSettings.Get("ADLDAPPath8Username")
    Public ADLDAPPath8Password As String = AppSettings.Get("ADLDAPPath8Password")
    Public ADLDAPPath8AuthType As String = AppSettings.Get("ADLDAPPath8AuthType")
    Public ADLDAPPath9 As String = AppSettings.Get("ADLDAPPath9")
    Public ADLDAPPath9Filter As String = AppSettings.Get("ADLDAPPath9Filter")
    Public ADLDAPPath9Username As String = AppSettings.Get("ADLDAPPath9Username")
    Public ADLDAPPath9Password As String = AppSettings.Get("ADLDAPPath9Password")
    Public ADLDAPPath9AuthType As String = AppSettings.Get("ADLDAPPath9AuthType")
    Public ADLDAPPath10 As String = AppSettings.Get("ADLDAPPath10")
    Public ADLDAPPath10Filter As String = AppSettings.Get("ADLDAPPath10Filter")
    Public ADLDAPPath10Username As String = AppSettings.Get("ADLDAPPath10Username")
    Public ADLDAPPath10Password As String = AppSettings.Get("ADLDAPPath10Password")
    Public ADLDAPPath10AuthType As String = AppSettings.Get("ADLDAPPath10AuthType")
    Public ADLDAPPath11 As String = AppSettings.Get("ADLDAPPath11")
    Public ADLDAPPath11Filter As String = AppSettings.Get("ADLDAPPath11Filter")
    Public ADLDAPPath11Username As String = AppSettings.Get("ADLDAPPath11Username")
    Public ADLDAPPath11Password As String = AppSettings.Get("ADLDAPPath11Password")
    Public ADLDAPPath11AuthType As String = AppSettings.Get("ADLDAPPath11AuthType")
    Public ADLDAPPath12 As String = AppSettings.Get("ADLDAPPath12")
    Public ADLDAPPath12Filter As String = AppSettings.Get("ADLDAPPath12Filter")
    Public ADLDAPPath12Username As String = AppSettings.Get("ADLDAPPath12Username")
    Public ADLDAPPath12Password As String = AppSettings.Get("ADLDAPPath12Password")
    Public ADLDAPPath12AuthType As String = AppSettings.Get("ADLDAPPath12AuthType")
    Public ADLDAPPath13 As String = AppSettings.Get("ADLDAPPath13")
    Public ADLDAPPath13Filter As String = AppSettings.Get("ADLDAPPath13Filter")
    Public ADLDAPPath13Username As String = AppSettings.Get("ADLDAPPath13Username")
    Public ADLDAPPath13Password As String = AppSettings.Get("ADLDAPPath13Password")
    Public ADLDAPPath13AuthType As String = AppSettings.Get("ADLDAPPath13AuthType")
    Public ADLDAPPath14 As String = AppSettings.Get("ADLDAPPath14")
    Public ADLDAPPath14Filter As String = AppSettings.Get("ADLDAPPath14Filter")
    Public ADLDAPPath14Username As String = AppSettings.Get("ADLDAPPath14Username")
    Public ADLDAPPath14Password As String = AppSettings.Get("ADLDAPPath14Password")
    Public ADLDAPPath14AuthType As String = AppSettings.Get("ADLDAPPath14AuthType")
    Public ADLDAPPath15 As String = AppSettings.Get("ADLDAPPath15")
    Public ADLDAPPath15Filter As String = AppSettings.Get("ADLDAPPath15Filter")
    Public ADLDAPPath15Username As String = AppSettings.Get("ADLDAPPath15Username")
    Public ADLDAPPath15Password As String = AppSettings.Get("ADLDAPPath15Password")
    Public ADLDAPPath15AuthType As String = AppSettings.Get("ADLDAPPath15AuthType")
    Public ADLDAPPath16 As String = AppSettings.Get("ADLDAPPath16")
    Public ADLDAPPath16Filter As String = AppSettings.Get("ADLDAPPath16Filter")
    Public ADLDAPPath16Username As String = AppSettings.Get("ADLDAPPath16Username")
    Public ADLDAPPath16Password As String = AppSettings.Get("ADLDAPPath16Password")
    Public ADLDAPPath16AuthType As String = AppSettings.Get("ADLDAPPath16AuthType")
    Public ADLDAPPath17 As String = AppSettings.Get("ADLDAPPath17")
    Public ADLDAPPath17Filter As String = AppSettings.Get("ADLDAPPath17Filter")
    Public ADLDAPPath17Username As String = AppSettings.Get("ADLDAPPath17Username")
    Public ADLDAPPath17Password As String = AppSettings.Get("ADLDAPPath17Password")
    Public ADLDAPPath17AuthType As String = AppSettings.Get("ADLDAPPath17AuthType")
    Public ADLDAPPath18 As String = AppSettings.Get("ADLDAPPath18")
    Public ADLDAPPath18Filter As String = AppSettings.Get("ADLDAPPath18Filter")
    Public ADLDAPPath18Username As String = AppSettings.Get("ADLDAPPath18Username")
    Public ADLDAPPath18Password As String = AppSettings.Get("ADLDAPPath18Password")
    Public ADLDAPPath18AuthType As String = AppSettings.Get("ADLDAPPath18AuthType")
    Public ADLDAPPath19 As String = AppSettings.Get("ADLDAPPath19")
    Public ADLDAPPath19Filter As String = AppSettings.Get("ADLDAPPath19Filter")
    Public ADLDAPPath19Username As String = AppSettings.Get("ADLDAPPath19Username")
    Public ADLDAPPath19Password As String = AppSettings.Get("ADLDAPPath19Password")
    Public ADLDAPPath19AuthType As String = AppSettings.Get("ADLDAPPath19AuthType")
    Public ADLDAPPath20 As String = AppSettings.Get("ADLDAPPath20")
    Public ADLDAPPath20Filter As String = AppSettings.Get("ADLDAPPath20Filter")
    Public ADLDAPPath20Username As String = AppSettings.Get("ADLDAPPath20Username")
    Public ADLDAPPath20Password As String = AppSettings.Get("ADLDAPPath20Password")
    Public ADLDAPPath20AuthType As String = AppSettings.Get("ADLDAPPath20AuthType")
    Public CopyrightMessage_Show As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("CopyrightMessage_Show"), True)
    Public AllowNotificationDelete As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("AllowNotificationDelete"), True)
    Public Sidebar_Show As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("Sidebar_Show"), True)
    Public AdminMode As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("AdminMode"), True)
    Public ShowSubmissionNumber As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("ShowSubmissionNumber"), True)
    Public ActionPanelBottom As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("ActionPanelBottom"), True)
    Public Sync_IncludeDepts As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("Sync_IncludeDepts"), True)
    Public Sync_IncludeGroups As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("Sync_IncludeGroups"), True)
    Public PDFTIFFConverter As String = AppSettings.Get("PDFTIFFConverter")
    Public ZukamiTemp As String = AppSettings.Get("ZukamiTemp")
    Public ImagemagickConverter As String = AppSettings.Get("ImagemagickConverter")
    Public HistoryCaptions As String = AppSettings.Get("HistoryCaptions")
    Public DropdownFindButtonCaption As String = AppSettings.Get("DropdownFindButtonCaption")
    Public DropdownCheckButtonCaption As String = AppSettings.Get("DropdownCheckButtonCaption")
    Public DropdownEmptyValueCaption As String = AppSettings.Get("DropdownEmptyValueCaption")
    Public DropdownNoMatchCaption As String = AppSettings.Get("DropdownNoMatchCaption")
    Public Syncdownrole As String = AppSettings.Get("Syncdownrole")
    Public EnforceDomain As String = AppSettings.Get("EnforceDomain")
    Public AutoDomain As String = AppSettings.Get("AutoDomain")
    Public HomeURL As String = AppSettings.Get("HomeURL")
    Public WorkHours As String = AppSettings.Get("WorkHours")
    Public Sync_DefaultPassword As String = AppSettings.Get("Sync_DefaultPassword")
    Public RichTextDisable As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("RichTextDisable"), False)
    Public DTPTimeIntervals As String = AppSettings.Get("DTPTimeIntervals")
    Public MyAppsTabCaption As String = AppSettings.Get("MyAppsTabCaption")
    Public DOCX2TIFFResolution As Integer = AppSettings.Get("DOCX2TIFFResolution")
    Public MyTasksColumns As String = AppSettings.Get("MyTasksColumns")
    Public MySubmissionsColumns As String = AppSettings.Get("MySubmissionsColumns")
    Public MyWorkflowHistoryColumns As String = AppSettings.Get("MyWorkflowHistoryColumns")
    Public ViewPageSize As Integer = AppSettings.Get("ViewPageSize")
    Public FileUploadMaxsize As Integer = AppSettings.Get("FileUploadMaxsize")
    Public ShowPrintFriendlyButton As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("ShowPrintFriendlyButton"), True)
    Public ShowErrorsAsJSPopup As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("ShowErrorsAsJSPopup"), False)
    Public ShowWorkflowActionButtonsInEditMode As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("ShowWorkflowActionButtonsInEditMode"), False)
    Public JumpOnEdit As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("JumpOnEdit"), False)
    Public DatefilterStart As String = AppSettings.Get("DatefilterStart")
    Public AllowExport As String = AppSettings.Get("AllowExport")
    Public DatefilterEndYearsForward As String = AppSettings.Get("DatefilterEndYearsForward")
    Public ShowWorkflowRemarksBoxByDefault As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("ShowWorkflowRemarksBoxByDefault"), False)
    Public IncludeSectionHeaderShadows As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("IncludeSectionHeaderShadows"), True)
    Public ShowPostAMessage As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("ShowPostAMessage"), True)
    Public EnableRequirementsGathering As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("EnableRequirementsGathering"), False)
    Public IgnoreExistingTablesOnImport As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("IgnoreExistingTablesOnImport"), False)
    Public AllowAD As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("AllowAD"), True)
    Public UseCaptcha As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("UseCaptcha"), False)
    Public ExportsData As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("ExportsData"), True)
    Public SimplifiedHTML As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("SimplifiedHTML"), False)
    Public FormBorderShow As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("FormBorderShow"), True)
    Public FormFooterRGB As String = GlobalFunctions.FormatData(AppSettings.Get("FormFooterRGB"))
    Public AdditionalDLLs As String = GlobalFunctions.FormatData(AppSettings.Get("AdditionalDLLs"))
    Public CaptchaHintMsg As String = GlobalFunctions.FormatData(AppSettings.Get("CaptchaHintMsg"))
    Public Annotatemode As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("AnnotateMode"), False)
    Public SPRINGMode As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("SPRINGMode"), False)
    Public KeepConnAlive As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("KeepDBConnectionAlive"), False)
    Public PerfDebugEnable As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("PerfDebugEnable"), True)
    Public DebugOutput As String = GlobalFunctions.FormatData(AppSettings.Get("DebugOutput"))
    Public FormRedirection As String = GlobalFunctions.FormatData(AppSettings.Get("FormRedirection"))
    Public BasePath As String = AppSettings.Get("BasePath")
    Public RequestforAmendmentMsg As String = AppSettings.Get("RequestforAmendmentMsg")
    Public SPRINGLogoutRedirect As String = AppSettings.Get("SPRINGLogoutRedirect")
    Public MobileModeFieldCaptionSize As String = AppSettings.Get("MobileModeFieldCaptionSize")
    Public BlynkImagePath As String = AppSettings.Get("BlynkImagePath")
    Public BlynkImageURL As String = AppSettings.Get("BlynkImageURL")
    Public BlynkFileDloadURL As String = AppSettings.Get("BlynkFileDloadURL")
    Public DropdownlistMax As String = AppSettings.Get("DropdownlistMax")
    Public ScriptLanguage As String = AppSettings.Get("ScriptLanguage")
    Public GoogleMapsKey As String = AppSettings.Get("GoogleMapsKey")
    Public DisplayDateFormat As String = AppSettings.Get("DisplayDateFormat")
    Public DisplayTimeFormat As String = AppSettings.Get("DisplayTimeFormat")
    Public AllowAmend As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("AllowAmend"))
    Public AllowReSubmit As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("AllowResubmit"))
    Public UseCSSFormHeader As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("UseCSSFormHeader"), False)
    Public DisableDeveloperLogin As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("DisableDeveloperLogin"), False)
    Public ForgotPassFunctionality As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("ForgotPassFunctionality"), False)
    Public DisableSignature As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("DisableSignature"), False)
    Public DisableDigitalSignature As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("DisableDigitalSignature"), False)
    Public PrimaryIDIsAlwaysParent As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("PrimaryIDIsAlwaysParent"), False)
    Public Pbkdf2Iterations As Integer = AppSettings.Get("Pbkdf2Iterations")
    Public ADSyncWebLogin As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("ADSyncWebLogin"), False)
    Public WEBFTAPI As String = AppSettings.Get("WEBFTAPI")
    Public WEBFTtempPath As String = AppSettings.Get("WEBFTtempPath")
    Public WEBFTUD As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("WEBFTUD"), False)
    Public AllowDMSLogin As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("AllowDMSLogin"), False)
    Public CurrentWinNTName As String = ""
    Public LastLoginResult As String = ""
    Public strCustomizebtn As String = ""
    Public strCustomizebtnname As String = ""
    Public AsposeLic As String = AppSettings.Get("AsposeLic")
    Public lognpath As String = AppSettings.Get("lognpath")
    Public Themeid As String = AppSettings.Get("Themeid")
    Public Hidesidemenuothers As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("Hidesidemenuothers"), False)
    Public ShowDetailedErrors As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("ShowDetailedErrors"), False)
    Public ShowSignatureinHistory As Boolean = GlobalFunctions.FormatBoolean(AppSettings.Get("ShowSignatureinHistory"), False)


    Public Function GetZukamiSettings() As ZukamiLib.ZukamiSettings
        'logger.Debug("start")
        Dim _settings As ZukamiLib.ZukamiSettings = Nothing
        Dim lngcounter As Integer
        For lngcounter = 0 To System.Web.HttpContext.Current.Request.Cookies.Count - 1
            If StrComp(System.Web.HttpContext.Current.Request.Cookies(lngcounter).Name, FormsAuthentication.FormsCookieName, CompareMethod.Text) = 0 Then
                Dim oriCookie As HttpCookie = System.Web.HttpContext.Current.Request.Cookies(lngcounter)
                Try
                    Dim oriFormTicket As FormsAuthenticationTicket = FormsAuthentication.Decrypt(oriCookie.Value)
                    _settings = ZukamiLib.ZukamiSettings.Liquify(oriFormTicket.UserData)
                Catch ex As Exception
                    logger.Error(ex, "failed to get data from cookie")
                End Try
                Exit For
            End If
        Next lngcounter

        If _settings Is Nothing Then
            'check Winauthlogin
            _settings = CheckWinAuthOkay()
        End If
        'If _settings Is Nothing Then
        '    logger.Debug("_settings Is Nothing")
        'Else
        '    'logger.Debug("_settings init done")
        'End If
        Return _settings
    End Function

    Private Function CheckWinAuthOkay() As ZukamiLib.ZukamiSettings
        'logger.Debug("start")
        CheckWinAuthOkay = Nothing
        CurrentWinNTName = ""
        LastLoginResult = ""
        If UseSSOIfAvailable = True Then
            Dim _identityname As String = System.Web.HttpContext.Current.User.Identity.Name
            CurrentWinNTName = _identityname
            If Len(_identityname) > 0 Then
                'we use the authenticated identity to look up the matching user
                Try


                    Dim _settings As ZukamiLib.ZukamiSettings = CreateDefaultZukamiSettings()
                    Dim _web As New ZukamiLib.WebSession(_settings)
                    Dim _result As ZukamiLib.WebSession.AUTH_RETURNCODES
                    Dim _Dataset As DataSet
                    _web.OpenConnection()
                    _Dataset = _web.LoginZukami(_identityname, _result)
                    Select Case _result
                        Case ZukamiLib.WebSession.AUTH_RETURNCODES.AUTH_GRANTED
                            _settings.CurrentFullName = GlobalFunctions.FormatData(_Dataset.Tables(0).Rows(0).Item("Fullname"))
                            _settings.CurrentUserGUID = GlobalFunctions.GetGUID(_Dataset.Tables(0).Rows(0).Item("UserID"))
                            _settings.Culture = GlobalFunctions.FormatData(_Dataset.Tables(0).Rows(0).Item("Locality"))
                            _settings.UICulture = GlobalFunctions.FormatData(_Dataset.Tables(0).Rows(0).Item("Language"))

                            _settings.PreviewMode = False
                            GlobalFunctions.CalculatePermissions(_web, _settings.CurrentUserGUID, _settings)

                            Dim _cookie As HttpCookie = GlobalFunctions.ZukamiLogin(_Dataset.Tables(0).Rows(0).Item("Username"), _settings.Solidify, False)
                            'Try
                            '    System.Web.HttpContext.Current.Response.Cookies.Add(_cookie)
                            'Catch ex As Exception
                            'End Try
                            GlobalFunctions.IssueFormsAuthenticationCookieAndSetSessionId(_cookie)
                            CheckWinAuthOkay = _settings
                            LastLoginResult = "Granted"
                        Case Else
                            LastLoginResult = "Unsuccessful"
                    End Select
                    _web.CloseConnection()
                    _web = Nothing
                Catch ex As Exception
                    LastLoginResult = "SSO Exception Error:" & ex.ToString
                End Try
            End If
        Else
            LastLoginResult = "SSO is not enabled. To enable SSO, please set the [UseSSOIfAvailable] key in web.config to true"
        End If

        'logger.Debug("LastLoginResult: " + LastLoginResult)
    End Function


    Public Sub CreateKookie()
        'logger.Debug("start, Authentication: " + Authentication)
        If StrComp(Authentication, "Windows", CompareMethod.Text) = 0 Then
            Dim _found As Boolean = False
            For lngcounter = 0 To System.Web.HttpContext.Current.Request.Cookies.Count - 1
                If StrComp(System.Web.HttpContext.Current.Request.Cookies(lngcounter).Name, FormsAuthentication.FormsCookieName, CompareMethod.Text) = 0 Then
                    'logger.Debug("login cookie found")
                    _found = True
                    Exit For
                End If
            Next lngcounter

            If _found = False Then
                'logger.Debug("login cookie not found")
                Dim _settings As ZukamiLib.ZukamiSettings = CreateDefaultZukamiSettings()
                Dim _web As New ZukamiLib.WebSession(_settings)
                _web.OpenConnection()

                'here we are going to assume a single AD 


                Dim _uname As String = System.Web.HttpContext.Current.User.Identity.Name
                Dim arrSplits() As String = Split(_uname, "\")
                If UBound(arrSplits) > 0 Then
                    _uname = arrSplits(UBound(arrSplits))
                End If

                Dim _wadname As String = GlobalFunctions.GetWinADName()
                Dim _dataset As DataSet = _web.LoginZukami(_wadname, Nothing)
                If _dataset.Tables(0).Rows.Count = 0 Then
                    _web.Users_Insert(_uname, "password", _wadname, _uname, "", "", "", "", "", "", "en-US", "en-US", False, False, Guid.Empty, Guid.Empty, 0, "AD;", "", "", "", "", "", Pbkdf2Iterations, Guid.Empty)
                    _dataset = _web.LoginZukami(_wadname, Nothing)
                End If



                If _dataset.Tables(0).Rows.Count > 0 Then
                    _settings.FramelessMode = False
                    _settings.CurrentFullName = GlobalFunctions.FormatData(_dataset.Tables(0).Rows(0).Item("Fullname"))
                    _settings.CurrentUserGUID = GlobalFunctions.GetGUID(_dataset.Tables(0).Rows(0).Item("UserID"))
                    _settings.Culture = GlobalFunctions.FormatData(_dataset.Tables(0).Rows(0).Item("Locality"))
                    _settings.UICulture = GlobalFunctions.FormatData(_dataset.Tables(0).Rows(0).Item("Language"))
                    Dim _username As String = GlobalFunctions.FormatData(_dataset.Tables(0).Rows(0).Item("Username"))
                    GlobalFunctions.CalculatePermissions(_web, _settings.CurrentUserGUID, _settings)
                    _web.CloseConnection()
                    _web = Nothing
                    Dim _cookie As HttpCookie = GlobalFunctions.ZukamiLogin(_username, _settings.Solidify, False)
                    'System.Web.HttpContext.Current.Response.Cookies.Add(_cookie)
                    GlobalFunctions.IssueFormsAuthenticationCookieAndSetSessionId(_cookie)
                Else
                    'if no matching user is found, we need to register this as an error
                    _web.CloseConnection()
                    _web = Nothing
                End If
            End If
        End If
    End Sub


    'Private Function WinAuthOkay() As Boolean
    '    WinAuthOkay = False
    '    If UseSSOIfAvailable = True Then
    '        Dim _identityname As String = System.Web.HttpContext.Current.User.Identity.Name
    '        If Len(_identityname) > 0 Then
    '            'we use the authenticated identity to look up the matching user
    '            Dim _settings As ZukamiLib.ZukamiSettings = CreateDefaultZukamiSettings()
    '            Dim _web As New ZukamiLib.WebSession(_settings)
    '            Dim _result As ZukamiLib.WebSession.AUTH_RETURNCODES
    '            Dim _Dataset As DataSet
    '            _web.OpenConnection()
    '            _Dataset = _web.LoginZukami(_identityname, _result)
    '            Select Case _result
    '                Case ZukamiLib.WebSession.AUTH_RETURNCODES.AUTH_GRANTED
    '                    _settings.CurrentFullName = GlobalFunctions.FormatData(_Dataset.Tables(0).Rows(0).Item("Fullname"))
    '                    _settings.CurrentUserGUID = GlobalFunctions.GetGUID(_Dataset.Tables(0).Rows(0).Item("UserID"))
    '                    _settings.Culture = GlobalFunctions.FormatData(_Dataset.Tables(0).Rows(0).Item("Locality"))
    '                    _settings.UICulture = GlobalFunctions.FormatData(_Dataset.Tables(0).Rows(0).Item("Language"))
    '                    If System.Web.HttpContext.Current.Request.QueryString("PMode") = "1" Then
    '                        _settings.PreviewMode = True
    '                    Else
    '                        _settings.PreviewMode = False
    '                    End If
    '                    GlobalFunctions.CalculatePermissions(_web, _settings.CurrentUserGUID, _settings)


    '                    If System.Web.HttpContext.Current.Request.QueryString("PMode") = "1" Then
    '                        _settings.PreviewMode = True
    '                        GlobalFunctions.SwitchToLogin(_Dataset.Tables(0).Rows(0).Item("Username"), _settings.Solidify, True)
    '                    Else
    '                        _settings.PreviewMode = False
    '                        Dim _cookie As HttpCookie = GlobalFunctions.ZukamiLogin(_Dataset.Tables(0).Rows(0).Item("Username"), _settings.Solidify, False)
    '                        System.Web.HttpContext.Current.Response.Cookies.Add(_cookie)
    '                    End If


    '                    Try
    '                        System.Web.HttpContext.Current.Response.Cookies.Remove("CurrentApp")
    '                    Catch ex As Exception

    '                    End Try


    '                    _web.CloseConnection()
    '                    _web = Nothing



    '                Case ZukamiLib.WebSession.AUTH_RETURNCODES.AUTH_DENIED
    '                    'We automatically sync down new accounts

    '                    'else
    '                Case ZukamiLib.WebSession.AUTH_RETURNCODES.AUTH_LOCKED
    '                    'txtUsername.Text = ""
    '                    'lblWarning.Text = "Your account has been locked. Please contact your Administrator."
    '                    'lblWarning.Visible = True
    '                Case ZukamiLib.WebSession.AUTH_RETURNCODES.AUTH_DISABLED
    '                    'txtUsername.Text = ""
    '                    'lblWarning.Text = "Your account has been disabled. Please contact your Administrator."
    '                    'lblWarning.Visible = True
    '                Case ZukamiLib.WebSession.AUTH_RETURNCODES.AUTH_EXPIRED
    '                    'txtUsername.Text = ""
    '                    'lblWarning.Text = "Your account has expired. Please contact your Administrator."
    '                    'lblWarning.Visible = True

    '            End Select
    '            _web.CloseConnection()
    '            _web = Nothing
    '        End If
    '    End If
    'End Function


    Public Function CreateDefaultZukamiSettings() As ZukamiLib.ZukamiSettings
        Dim _settings As New ZukamiLib.ZukamiSettings
        If StrComp(Authentication, "Windows", CompareMethod.Text) = 0 Then
            _settings.AuthenticationMode = ZukamiLib.ZukamiSettings.AUTHENTICATION_MODES.AUTHMODE_AD
        Else
            _settings.AuthenticationMode = ZukamiLib.ZukamiSettings.AUTHENTICATION_MODES.AUTHMODE_ZUKAMI
        End If
        _settings.PrimaryConnectionString = ConnectionString
        _settings.UploadPath = UploadPath
        _settings.CurrentFullName = ""
        _settings.Culture = "en-US"
        _settings.UICulture = "en"
        _settings.Queue = Queue
        _settings.FramelessMode = False
        _settings.SessionId = Left(Guid.NewGuid().ToString(), 5)
        Return _settings
    End Function
End Module



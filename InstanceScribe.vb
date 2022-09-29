Imports System.Data
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Drawing.Drawing2D
Imports System.Drawing.Text
Imports System.Reflection
Imports System.Security
Imports Progress.Open4GL.Proxy
Imports Winnovative.WnvHtmlConvert
Imports NLog

Public Class InstanceScribe
    Public Class FormAttributes
        Private _FormCaptionLabel As Object = Nothing
        Private _FormDescriptionLabel As Object = Nothing
        Private _FormTable As Table = Nothing
        Private _FormTableBorder As Table = Nothing
        Private _AppID As Guid = Nothing
        Private _FormID As String = ""
        Private _FormParentID As String = ""
        Private _FormSessionParentRecID As String = ""
        Private _FormSessionRecID As String = ""
        Private _WorkflowIDHolder As Object = Nothing
        Private _FormObject As Object = Nothing

        Public Property FormSessionRecID() As String
            Get
                Return _FormSessionRecID
            End Get
            Set(ByVal value As String)
                _FormSessionRecID = value
            End Set
        End Property

        Public Property FormSessionParentRecID() As String
            Get
                Return _FormSessionParentRecID
            End Get
            Set(ByVal value As String)
                _FormSessionParentRecID = value
            End Set
        End Property

        Public Property FormObject() As Object
            Get
                Return _FormObject
            End Get
            Set(ByVal value As Object)
                _FormObject = value
            End Set
        End Property

        Public Property WorkflowIDholder() As Object
            Get
                Return _WorkflowIDHolder
            End Get
            Set(ByVal value As Object)
                _WorkflowIDHolder = value
            End Set
        End Property

        Public Property FormID() As String
            Get
                Return _FormID
            End Get
            Set(ByVal value As String)
                _FormID = value
            End Set
        End Property

        Public Property FormParentID() As String
            Get
                Return _FormParentID
            End Get
            Set(ByVal value As String)
                _FormParentID = value
            End Set
        End Property

        Public Property AppID() As Guid
            Get
                Return _AppID
            End Get
            Set(ByVal value As Guid)
                _AppID = value
            End Set
        End Property

        Public Property FormTableBorder() As Table
            Get
                Return _FormTableBorder
            End Get
            Set(ByVal value As Table)
                _FormTableBorder = value
            End Set
        End Property

        Public Property FormTable() As Table
            Get
                Return _FormTable
            End Get
            Set(ByVal value As Table)
                _FormTable = value
            End Set
        End Property

        Public Property FormCaptionLabel() As Object
            Get
                Return _FormCaptionLabel
            End Get
            Set(ByVal value As Object)
                _FormCaptionLabel = value
            End Set
        End Property

        Public Property FormDescriptionLabel() As Object
            Get
                Return _FormDescriptionLabel
            End Get
            Set(ByVal value As Object)
                _FormDescriptionLabel = value
            End Set
        End Property
    End Class

    Public Class DPair
        Private _Caption As String = ""
        Private _Value As String = ""
        Public Sub New(Caption As String, Value As String)
            _Caption = Caption
            _Value = Value
        End Sub

        Public Property Caption() As String
            Get
                Return _Caption
            End Get
            Set(value As String)
                _Caption = value
            End Set
        End Property

        Public Property Value() As String
            Get
                Return _Value
            End Get
            Set(value As String)
                _Value = value
            End Set
        End Property
    End Class



    Public Class DBQuery
        Private _sqltable As String
        Private _whereclause As String = ""
        Private _customSQL As String = ""
        Private _lastError As String = ""
        Private _Internalconnsession As ZukamiLib.WebSession = Nothing
        Private _fieldColl As New Collection
        Private _valuecoll As New Collection

        Public Function InitiateSession() As ZukamiLib.WebSession
            If _Internalconnsession Is Nothing Then
                _Internalconnsession = New ZukamiLib.WebSession(GetZukamiSettings)
                _Internalconnsession.OpenConnection()
            Else
                If _Internalconnsession.IsDatabaseOpen() = False Then
                    _Internalconnsession.OpenConnection()
                End If
            End If
            Return _Internalconnsession
        End Function

        Public Function GetSession() As ZukamiLib.WebSession
            Return InitiateSession()
        End Function

        Public Sub TerminateSession()
            If _Internalconnsession Is Nothing = False Then
                _Internalconnsession.CloseConnection()
                _Internalconnsession = Nothing
            End If
        End Sub


        Public ReadOnly Property LastError() As String
            Get
                Return _lastError
            End Get
        End Property

        Public Sub SetInternalConn(ByRef conn As ZukamiLib.WebSession)
            _Internalconnsession = conn
        End Sub

        Public Sub New(ByVal SQLTable As String)
            _sqltable = SQLTable
            _customSQL = ""
        End Sub

        Public Sub AddCustomSQL(ByVal CustomSQL As String)
            _customSQL = CustomSQL
        End Sub

        Public Sub AddInsertPair(Fieldname As String, Fieldvalue As String)
            _fieldColl.Add(Fieldname)
            _valuecoll.Add(Fieldvalue)
        End Sub


        Public Sub AddFilter(ByVal FieldName As String, ByVal OperatorSymbol As String, ByVal Value As String)
            If Len(_whereclause) > 0 Then _whereclause += " AND "
            _whereclause += "[" & FieldName & "] " & OperatorSymbol & " N'" & GlobalFunctions.FormatSQLData(Value) & "'"
        End Sub

        Public Sub RunInsert()
            Dim _webObj As ZukamiLib.WebSession = Nothing
            If KeepConnAlive = True Then
                _webObj = GetSession()
            Else
                _webObj = New ZukamiLib.WebSession(GetZukamiSettings)
                _webObj.OpenConnection()
            End If



            Dim _sql As String = "INSERT INTO [" & _sqltable & "]"
            Dim j As Integer
            Dim _fpiece As String = ""
            Dim _vpiece As String = ""
            For j = 1 To _fieldColl.Count
                Dim _fname As String = _fieldColl.Item(j)
                Dim _vname As String = _valuecoll.Item(j)
                If Len(_fpiece) > 0 Then _fpiece += ","
                If Len(_vpiece) > 0 Then _vpiece += ","
                _fpiece += "[" + _fname + "]"
                _vpiece += "'" + GlobalFunctions.FormatSQLData(_vname) + "'"

            Next j
            _sql += "(" & _fpiece & ") VALUES (" & _vpiece & ")"


            _webObj.CustomSQLCommand(_sql)

            _webObj.CustomSQLExecute()

            _lastError = _webObj.LastError

            If KeepConnAlive = False Then
                _webObj.CloseConnection()
                _webObj = Nothing
            End If
        End Sub

        Public Function Run() As DataSet
            Dim _webObj As ZukamiLib.WebSession = Nothing
            If KeepConnAlive = True Then
                _webObj = GetSession()
            Else
                _webObj = New ZukamiLib.WebSession(GetZukamiSettings)
                _webObj.OpenConnection()
            End If


            If Len(_customSQL) > 0 Then
                _webObj.CustomSQLCommand(_customSQL)
            Else
                Dim _sql As String = "SELECT * FROM [" & _sqltable & "]"
                If Len(_whereclause) > 0 Then _sql += " WHERE " & _whereclause
                _webObj.CustomSQLCommand(_sql)
            End If

            Run = _webObj.CustomSQLExecuteReturn()

            _lastError = _webObj.LastError

            If KeepConnAlive = False Then
                _webObj.CloseConnection()
                _webObj = Nothing
            End If


        End Function
    End Class

    Public Class ExternalDataQuery
        Private _sqltable As String
        Private _whereclause As String = ""
        Private _customSQL As String = ""
        Private _lastError As String = ""
        Private _fieldColl As New Collection
        Private _valuecoll As New Collection
        Private _allfieldColl As New Collection
        Private _allvaluecoll As New Collection
        Private _funcType As String = ""
        Private _initdone As Boolean = False
        Private _funcname As String = ""

        Public Sub New(ExternalDatasourceName As String, FunctionName As String, Type As String)
            _initdone = True
            _funcname = FunctionName
            _sqltable = "QF_" & FunctionName
            _funcType = Type
            _customSQL = ""
        End Sub

        Public ReadOnly Property LastError() As String
            Get
                Return _lastError
            End Get
        End Property

        Public Sub AddInputArgument(Fieldname As String, FieldValue As String)
            _fieldColl.Add(Fieldname)
            _valuecoll.Add(FieldValue)

            If _allfieldColl.Contains(Fieldname) = False Then
                _allfieldColl.Add(Fieldname)
                _allvaluecoll.Add(FieldValue)
            End If

        End Sub
        Public Sub AddFilterArgument(Fieldname As String, FieldValue As String)
            If Len(_whereclause) > 0 Then _whereclause += " AND "
            _whereclause += "[" & Fieldname & "] = N'" & GlobalFunctions.FormatSQLData(FieldValue) & "'"

            If _allfieldColl.Contains(Fieldname) = False Then
                _allfieldColl.Add(Fieldname)
                _allvaluecoll.Add(FieldValue)
            End If
        End Sub


        Public Function RunActual() As String
            Dim _webObj As ZukamiLib.WebSession = Nothing
            _webObj = New ZukamiLib.WebSession(GetZukamiSettings)
            _webObj.OpenConnection()

            Dim _newid As String = Guid.NewGuid.ToString
            Dim _input As String = "1|" & _funcname & "|"

            Dim i As Integer
            For i = 1 To _allfieldColl.Count
                Dim _f As String = _allfieldColl.Item(i)
                Dim _v As String = _allvaluecoll.Item(i)
                _input += _v & "|"
            Next

            Dim _sql As String = "INSERT INTO [qf_as400job] ([ID],[Input],[Output],[Message],[Status]) VALUES('" & _newid & "','" & _input & "','','','Pending')"
            _webObj.CustomSQLCommand(_sql)
            _webObj.CustomSQLExecute()

            Dim _getvalue As String = ""
            Dim _timeelapsed As Integer = 0
            Do While _getvalue = "" And _timeelapsed < 15
                System.Threading.Thread.Sleep(1000)
                _webObj.CustomSQLCommand("SELECT [Output] FROM [qf_as400job] WHERE ID='" & _newid & "'")
                Dim _set As DataSet = _webObj.CustomSQLExecuteReturn()
                If _set.Tables(0).Rows.Count > 0 Then
                    _getvalue = GlobalFunctions.FormatData(_set.Tables(0).Rows(0).Item("Output"))
                End If
                _set.Dispose()
                _set = Nothing
                _timeelapsed += 1
            Loop
            If Len(_getvalue) > 0 Then
                'result was found, so we rebuild it into dataset
                RunActual = _getvalue
            End If


            _webObj.CloseConnection()
            _webObj = Nothing

        End Function

        Public Function Run() As DataSet
            'If WebconfigSettings.ProxyMode = "" And StrComp(_funcname, "SEARCH_CUST", vbTextCompare) = 0 Then
            '    Return RunActual()
            'End If

            _lastError = ""
            If _initdone = False Then
                _lastError = "scribe.OpenSession() was not called"
                Return Nothing
            End If

            Dim _webObj As ZukamiLib.WebSession = Nothing
            _webObj = New ZukamiLib.WebSession(GetZukamiSettings)
            _webObj.OpenConnection()

            Dim _sql As String = ""
            Select Case LCase(_funcType)
                Case "get"
                    _sql = "SELECT * FROM [" & _sqltable & "]"
                    If Len(_whereclause) > 0 Then _sql += " WHERE " & _whereclause
                Case "ins"
                    _sql = "INSERT INTO [" & _sqltable & "]"
                    Dim j As Integer
                    Dim _fpiece As String = ""
                    Dim _vpiece As String = ""
                    For j = 1 To _fieldColl.Count
                        Dim _fname As String = _fieldColl.Item(j)
                        Dim _vname As String = _valuecoll.Item(j)
                        If Len(_fpiece) > 0 Then _fpiece += ","
                        If Len(_vpiece) > 0 Then _vpiece += ","
                        _fpiece += "[" + _fname + "]"
                        _vpiece += "'" + GlobalFunctions.FormatSQLData(_vname) + "'"
                    Next j
                    _sql += "(" & _fpiece & ") VALUES (" & _vpiece & ")"
                Case "upd"
                    _sql = "UPDATE [" & _sqltable & "] SET "
                    Dim j As Integer
                    Dim _fpiece As String = ""
                    Dim _vpiece As String = ""
                    Dim _fieldlist As String = ""
                    For j = 1 To _fieldColl.Count
                        Dim _fname As String = _fieldColl.Item(j)
                        Dim _vname As String = _valuecoll.Item(j)
                        If Len(_fieldlist) > 0 Then _fieldlist += ","
                        _fieldlist += "[" + _fname + "]=" & "'" + GlobalFunctions.FormatSQLData(_vname) + "'"
                    Next j
                    _sql += _fieldlist
                    If Len(_whereclause) > 0 Then _sql += " WHERE " & _whereclause
                Case "del"
                    _sql = "DELETE FROM [" & _sqltable & "]"
                    If Len(_whereclause) > 0 Then _sql += " WHERE " & _whereclause
            End Select

            _webObj.CustomSQLCommand(_sql)
            Run = _webObj.CustomSQLExecuteReturn()

            If Len(_webObj.LastError) > 0 Then
                _lastError = "SQL is :" & _sql & " and error follows below:" & vbCrLf & _webObj.LastError
            End If


            _webObj.CloseConnection()
            _webObj = Nothing



        End Function
    End Class

    Public Shared logger As Logger = LogManager.GetCurrentClassLogger()
    Private _fieldList As DataSet = Nothing
    Private _ListTableName As String = ""
    Private _HTML As String = ""
    Private _DataBag As DataSet = Nothing
    Private _FormDetails As ScribeFormDetails = Nothing
    Private _zfields As Collection = Nothing
    Private _TableDataBags As Collection = Nothing
    Private _SubmissionDetails As New SubmissionDetails
    Private _FormButtonsList As Control = Nothing
    Private _FormAttributes As FormAttributes = Nothing
    Private _filterargs As Collection = Nothing
    Private _GlobalImage As Bitmap = Nothing
    Private _GlobalGraphics As Graphics = Nothing
    Private _InternalConnSession As ZukamiLib.WebSession = Nothing

    Public ReadOnly Property IsCurrentRequestMobileVersion() As Boolean
        Get
            Return GlobalFunctions.IsCurrentRequestMobileVersion()
        End Get
    End Property
    Public Property HTML() As String
        Get
            Return _HTML
        End Get
        Set(value As String)
            _HTML = value
        End Set
    End Property

    Public Function InitiateSession() As ZukamiLib.WebSession
        If _InternalConnSession Is Nothing Then
            _InternalConnSession = New ZukamiLib.WebSession(GetZukamiSettings)
            _InternalConnSession.OpenConnection()
        Else
            If _InternalConnSession.IsDatabaseOpen() = False Then
                _InternalConnSession.OpenConnection()
            End If
        End If
        Return _InternalConnSession
    End Function

    Public Function GetSession() As ZukamiLib.WebSession
        Return InitiateSession()
    End Function
    Public Sub UserPasswordHistory_Insert(ByVal username As String, ByVal password As String, Optional ByVal isNew As Boolean = False)
        GlobalFunctions.UserPasswordHistory_Insert(username, password, isNew)

    End Sub
    Public Sub TerminateSession()
        If _InternalConnSession Is Nothing = False Then
            _InternalConnSession.CloseConnection()
            _InternalConnSession = Nothing
        End If
    End Sub


    Public Property FormAttributeObject() As FormAttributes
        Get
            Return _FormAttributes
        End Get
        Set(ByVal value As FormAttributes)
            _FormAttributes = value
        End Set
    End Property

    Public Property FormButtonsList() As Control
        Get
            logger.Debug("_FormButtonsList: " + GlobalFunctions.FormatData(_FormButtonsList))
            Return _FormButtonsList
        End Get
        Set(ByVal value As Control)
            _FormButtonsList = value
        End Set
    End Property

    Public Function OneWayHash(ByVal Data As String) As String
        Return GlobalFunctions.HashPassword(Data)
    End Function

    Public ReadOnly Property ConnectionStringTemplate() As String
        Get
            Return WebconfigSettings.CloudConnectionString
        End Get
    End Property

    Public ReadOnly Property SessionSettings() As ZukamiLib.ZukamiSettings
        Get
            Return GetZukamiSettings()
        End Get
    End Property

    '  Public Function Cloudware_CheckAccountExists(ByVal AccountName As String) As Boolean
    '      Dim _zed As New ZED.ZED
    '      Cloudware_CheckAccountExists = _zed.Cloudware_AccountExists(AccountName)
    '      _zed.Dispose()
    '      _zed = Nothing
    '  End Function

    '  Public Function Cloudware_ResetPassword(ByVal AccountName As String, ByVal NewPassword As String) As Boolean
    '      Dim _Zed As New ZED.ZED
    '      Cloudware_ResetPassword = _Zed.Cloudware_ResetPassword(AccountName, NewPassword)
    '      _Zed.Dispose()
    '      _Zed = Nothing
    '  End Function

    Public Sub SetAllAnnotation(ByVal Show As Boolean)
        Dim _webObj As ZukamiLib.WebSession = Nothing
        If KeepConnAlive = True Then
            _webObj = GetSession()
        Else
            _webObj = New ZukamiLib.WebSession(GetZukamiSettings)
            _webObj.OpenConnection()
        End If

        If _FormAttributes Is Nothing = False Then
            Dim _formID As String = _FormAttributes.FormID

            _webObj.CustomSQLCommand("UPDATE [Annotation] SET Invisible=" & IIf(Show = True, "0", "1") & " WHERE RecordID='" & PrimaryID.ToString & "'")
            _webObj.CustomClearParameters()
            _webObj.CustomSQLExecute()
        End If

        If KeepConnAlive = False Then
            _webObj.CloseConnection()
            _webObj = Nothing
        End If

    End Sub

    Public Function MSWord_CreateNew() As Object
        GemBox.Document.ComponentInfo.SetLicense("DMPX-J9AT-EL54-2YBC")
        Dim _doc As New GemBox.Document.DocumentModel
        Return _doc
    End Function

    'Public Function PushDocumentToDocLibrary(ByVal URL As String, ByVal Domain As String, ByVal Username As String, ByVal Password As String, ByVal PathOfDocument As String, ByVal DocLibraryname As String) As Object
    '    Dim item As New Sharepoint.SPFunction
    '    Return item.PushDocumentToDocLibrary(URL, Domain, Username, Password, PathOfDocument, DocLibraryname)
    'End Function

    Public Function MSWord_AddParagraph(ByRef DocObject As Object, ByVal Text As String) As Boolean
        Dim _doc As GemBox.Document.DocumentModel = DocObject

        Dim arrFull() As GemBox.Document.Inline
        Dim arrSplits() As String = Split(Text, "\n")
        Dim _counter As Integer
        Dim _current As Integer = 0
        For _counter = 0 To UBound(arrSplits)
            Dim _item As New GemBox.Document.Run(_doc, arrSplits(_counter))
            ReDim Preserve arrFull(_current)
            arrFull(_current) = _item
            _current += 1

            Dim _br As New GemBox.Document.SpecialCharacter(_doc, GemBox.Document.SpecialCharacterType.LineBreak)
            ReDim Preserve arrFull(_current)
            arrFull(_current) = _br
            _current += 1

        Next _counter

        Dim _internal As GemBox.Document.Paragraph = New GemBox.Document.Paragraph(_doc, arrFull)
        _doc.Sections.Add(New GemBox.Document.Section(_doc, _internal))
    End Function

    Public Function MSWord_SaveDoc(ByRef DocObject As Object, ByVal OutputPath As String) As Boolean
        Try
            System.IO.File.Delete(OutputPath)
        Catch ex As Exception
        End Try
        Dim _doc As GemBox.Document.DocumentModel = DocObject
        _doc.Save(OutputPath, GemBox.Document.SaveOptions.DocxDefault)
        _doc = Nothing
    End Function

    Public Function MSWord_Create(ByVal TextToWrite As String, ByVal OutputPath As String) As Boolean
        Try
            System.IO.File.Delete(OutputPath)
        Catch ex As Exception
        End Try

        GemBox.Document.ComponentInfo.SetLicense("DMPX-J9AT-EL54-2YBC")
        Dim _doc As New GemBox.Document.DocumentModel

        Dim _internal As GemBox.Document.Paragraph = New GemBox.Document.Paragraph(_doc, TextToWrite)
        _doc.Sections.Add(New GemBox.Document.Section(_doc, _internal))
        _doc.Save(OutputPath, GemBox.Document.SaveOptions.DocxDefault)
        _doc = Nothing
    End Function

    Public Function MSWord_GrabText(ByVal InputPath As String) As String
        GemBox.Document.ComponentInfo.SetLicense("DMPX-J9AT-EL54-2YBC")
        Dim _doc As GemBox.Document.DocumentModel
        _doc = GemBox.Document.DocumentModel.Load(InputPath, GemBox.Document.LoadOptions.DocxDefault)

        Dim sb As New StringBuilder()

        For Each paragraph As GemBox.Document.Paragraph In _doc.GetChildElements(True, GemBox.Document.ElementType.Paragraph)
            For Each run As GemBox.Document.Run In paragraph.GetChildElements(True, GemBox.Document.ElementType.Run)
                Dim isBold As Boolean = run.CharacterFormat.Bold
                Dim text As String = run.Text

                sb.AppendFormat("{0}{1}{2}", If(isBold, "<b>", ""), text, If(isBold, "</b>", ""))
            Next
            sb.AppendLine()
        Next

        Return sb.ToString

    End Function

    Public Sub SendMessage()

    End Sub

    Public Function GetUploadPath() As String
        Return UploadPath.TrimEnd("\")
    End Function

    Public Function CreateNewUploadFolder(Optional ByRef FolderGuid As String = "") As String
        FolderGuid = Guid.NewGuid.ToString
        Dim _path As String = UploadPath.TrimEnd("\") & "\" & FolderGuid
        System.IO.Directory.CreateDirectory(_path)
        Return _path
    End Function

    Public Function GetFullCountryName(ByVal CountryCode As String) As String
        Return GlobalFunctions.GetFullCountryName(CountryCode)
    End Function

    Public Function TableRowExists(TableField As String, FieldName As String, Value As Object) As Boolean
        Dim dt As DataTable = Data(TableField)
        If dt Is Nothing Then Return False
        Dim i As Integer
        For i = 0 To dt.Rows.Count - 1
            Try
                If dt.Rows(i).Item(FieldName) = Value Then
                    Return True
                End If
            Catch ex As Exception
            End Try
        Next i
        Return False
    End Function

    Public Function CreateCV(ByVal ConfidentialTemplate As String, ByVal NonConfidentialTemplate As String, ByVal TargetPath As String, ByRef drow As DataRow, ByRef educationrow As DataSet, ByRef languagesrow As DataSet, ByRef careerRow As DataSet, ByVal isconfidential As Boolean) As String
        GemBox.Document.ComponentInfo.SetLicense("DMPX-J9AT-EL54-2YBC")
        Dim document As GemBox.Document.DocumentModel = Nothing

        Dim _template As String = ""
        If isconfidential = True Then
            _template = ConfidentialTemplate  '"d:\vbtest\CVTemplate.docx"
        Else
            _template = NonConfidentialTemplate '"d:\vbtest\CVTemplate2.docx"
        End If
        Dim sourceDocument As GemBox.Document.DocumentModel = GemBox.Document.DocumentModel.Load(_template, GemBox.Document.DocxLoadOptions.DocxDefault)


        Dim _Name As String = GlobalFunctions.FormatData(drow.Item("Salutation")) & " " & GlobalFunctions.FormatData(drow.Item("Name"))
        Dim _Designation As String = GlobalFunctions.FormatData(drow.Item("Designation"))
        Dim _Country As String = GetFullCountryName(GlobalFunctions.FormatData(drow.Item("Nationality")))
        Dim _DOB As String = GlobalFunctions.FormatData(drow.Item("Date of Birth"))
        Dim _Age As String = GlobalFunctions.FormatData(drow.Item("Age"))
        Dim _marital As String = GlobalFunctions.FormatData(drow.Item("Marital Status"))

        Dim _Education As String = ""
        If educationrow Is Nothing = False Then
            Dim i As Integer
            For i = 0 To educationrow.Tables(0).Rows.Count - 1
                Dim _u1 As String = GlobalFunctions.FormatData(educationrow.Tables(0).Rows(i).Item("University Name"))
                Dim _u2 As String = GlobalFunctions.FormatData(educationrow.Tables(0).Rows(i).Item("Faculty"))
                Dim _u3 As String = GlobalFunctions.FormatData(educationrow.Tables(0).Rows(i).Item("Graduation Year"))
                If Len(_Education) > 0 Then _Education += "\n"
                _Education += _u1 & "," & _u2 & "," & _u3
            Next
        End If

        Dim _career As String = ""
        If careerRow Is Nothing = False Then
            Dim i As Integer
            For i = 0 To careerRow.Tables(0).Rows.Count - 1
                Dim _u1 As String = GlobalFunctions.FormatData(careerRow.Tables(0).Rows(i).Item("Period of Service"))
                Dim _u2 As String = GlobalFunctions.FormatData(careerRow.Tables(0).Rows(i).Item("Designation"))
                Dim _u3 As String = GlobalFunctions.FormatData(careerRow.Tables(0).Rows(i).Item("Organization"))
                If Len(_career) > 0 Then _career += "\n"
                _career += _u1 & "," & _u2 & "," & _u3
            Next
        End If

        Dim _Appts As String = ""
        Dim _Memberships As String = GlobalFunctions.FormatData(drow.Item("Memberships"))
        Dim _Awards As String = GlobalFunctions.FormatData(drow.Item("Awards / Decorations"))
        Dim _Publications As String = GlobalFunctions.FormatData(drow.Item("Publications"))


        Dim _languages As String = ""
        If languagesrow Is Nothing = False Then
            Dim i As Integer
            For i = 0 To languagesrow.Tables(0).Rows.Count - 1
                Dim _u1 As String = GlobalFunctions.FormatData(languagesrow.Tables(0).Rows(i).Item("Language"))
                If Len(_languages) > 0 Then _languages += ","
                _languages += _u1
            Next
        End If

        Dim _Interests As String = GlobalFunctions.FormatData(drow.Item("Delegate's Hobby"))
        Dim _Remarks As String = GlobalFunctions.FormatData(drow.Item("Remarks"))
        Dim _ConfidentialRemarks As String = GlobalFunctions.FormatData(drow.Item("Confidential"))
        Dim _Photo As String = GlobalFunctions.FormatData(drow.Item("Photo"))
        If Len(_Photo) > 0 Then _Photo = GetFullUploadedFilepath(_Photo)

        If Len(_Photo) > 0 Then
            Dim _pic As New GemBox.Document.Picture(sourceDocument, _Photo, 80, 100)
            Dim _para As New GemBox.Document.Paragraph(sourceDocument, _pic)
            _para.ParagraphFormat.Alignment = GemBox.Document.HorizontalAlignment.Right
            sourceDocument.Sections.Item(0).Blocks.Insert(0, _para)

        End If





        For Each paragraph As GemBox.Document.Paragraph In sourceDocument.GetChildElements(True, GemBox.Document.ElementType.Paragraph)
            For Each run As GemBox.Document.Run In paragraph.GetChildElements(True, GemBox.Document.ElementType.Run)
                Dim isHeader As Boolean = False
                Dim text As String = run.Text
                If text = "[$Name]" Then
                    isHeader = True
                    run.Text = _Name
                    sourceDocument.MailMerge.Execute(paragraph)
                ElseIf text = "[$Designation]" Then
                    isHeader = True
                    run.Text = _Designation
                    sourceDocument.MailMerge.Execute(paragraph)
                ElseIf text = "[$Country]" Then
                    isHeader = True
                    run.Text = _Country
                    sourceDocument.MailMerge.Execute(paragraph)
                ElseIf text = "[$DOB]" Then
                    isHeader = True
                    run.Text = _DOB
                    sourceDocument.MailMerge.Execute(paragraph)
                ElseIf text = "[$Age]" Then
                    isHeader = True
                    run.Text = _Age
                    sourceDocument.MailMerge.Execute(paragraph)
                ElseIf text = "[$Marital]" Then
                    isHeader = True
                    run.Text = _marital
                    sourceDocument.MailMerge.Execute(paragraph)
                ElseIf text = "[$Education]" Then
                    isHeader = True
                    run.Text = _Education
                    sourceDocument.MailMerge.Execute(paragraph)
                ElseIf text = "[$Career]" Then
                    isHeader = True
                    run.Text = _career
                    sourceDocument.MailMerge.Execute(paragraph)
                ElseIf text = "[$Appointments]" Then
                    isHeader = True
                    run.Text = _Appts
                    sourceDocument.MailMerge.Execute(paragraph)
                ElseIf text = "[$Memberships]" Then
                    isHeader = True
                    run.Text = _Memberships
                    sourceDocument.MailMerge.Execute(paragraph)
                ElseIf text = "[$Awards]" Then
                    isHeader = True
                    run.Text = _Awards
                    sourceDocument.MailMerge.Execute(paragraph)
                ElseIf text = "[$Publications]" Then
                    isHeader = True
                    run.Text = _Publications
                    sourceDocument.MailMerge.Execute(paragraph)
                ElseIf text = "[$Languages]" Then
                    isHeader = True
                    run.Text = _languages
                    sourceDocument.MailMerge.Execute(paragraph)
                ElseIf text = "[$Interests]" Then
                    isHeader = True
                    run.Text = _Interests
                    sourceDocument.MailMerge.Execute(paragraph)
                ElseIf text = "[$Remarks]" Then
                    isHeader = True
                    run.Text = _Remarks
                    sourceDocument.MailMerge.Execute(paragraph)
                ElseIf text = "[$ConfidentialRemarks]" Then
                    isHeader = True
                    run.Text = _ConfidentialRemarks
                    sourceDocument.MailMerge.Execute(paragraph)
                End If

            Next
        Next

        sourceDocument.Save(TargetPath, GemBox.Document.DocxSaveOptions.DocxDefault)
        Return TargetPath
    End Function

    Public Sub SetSubFormButtonVisibility(FieldName As String, ShowAddButton As Boolean, ShowDeleteButton As Boolean)
        Try
            Control(FieldName).ShowTableEntryRow = ShowAddButton
            Control(FieldName).DataGridObject.AllowDelete = ShowDeleteButton
            Control(FieldName).RefreshGrid()
        Catch ex As Exception

        End Try

    End Sub

    Public Function CallWebService(ByVal URL As String, ByVal ServiceName As String, ByVal MethodName As String, ByVal ArgsArray() As String) As Object
        Dim invoker As New WebServiceInvoker2(New Uri(URL))
        Dim service As String = ServiceName
        Dim method As String = MethodName

        Return invoker.InvokeMethod(Of String)(service, method, ArgsArray)
    End Function

    Public Function InvokeWebServiceCall(ByVal WebServiceID As String, ServiceName As String, ByVal MethodName As String, ByRef ArgsArray() As Object, Optional DLLList As String = "") As Object
        logger.Trace("start")
        Dim _webObj As New ZukamiLib.WebSession(GetZukamiSettings)
        _webObj.OpenConnection()
        Dim _set As DataSet = _webObj.Webservices_GetRecord(New Guid(WebServiceID))
        _webObj.CloseConnection()
        _webObj = Nothing
        logger.Trace("db query done - got WebService by ID")

        Dim _URL As String = ""
        Dim _password As String = ""
        Dim _username As String = ""
        Dim _useproxy As Boolean = False
        Dim _proxyaddr As String = ""
        Dim _proxyport As Integer = -1
        Dim _domain As String = ""
        If _set.Tables(0).Rows.Count > 0 Then
            _URL = GlobalFunctions.FormatData(_set.Tables(0).Rows(0).Item("URL"))
            _useproxy = GlobalFunctions.FormatBoolean(_set.Tables(0).Rows(0).Item("UseProxy"))
            _proxyaddr = GlobalFunctions.FormatData(_set.Tables(0).Rows(0).Item("ProxyAddress"))
            _proxyport = GlobalFunctions.FormatInteger(_set.Tables(0).Rows(0).Item("ProxyPort"), -1)
            _username = GlobalFunctions.FormatData(_set.Tables(0).Rows(0).Item("Username"))
            _password = GlobalFunctions.FormatData(_set.Tables(0).Rows(0).Item("Password"))
            _domain = GlobalFunctions.FormatData(_set.Tables(0).Rows(0).Item("Domain"))
            If Len(_URL) = 0 Then Return ""
        Else
            Return ""
        End If

        logger.Trace("before do actual calling of web service")
        Dim result As Object = WebSvcCls.CallingWebService(_URL, ServiceName, MethodName, ArgsArray, DLLList, _username, _password, _domain, _proxyaddr, _proxyport)
        logger.Trace("done")
        Return result
    End Function

    Public Sub PopulateRoleTable(RolesTable As String, RoleColumn As String, UserID As String)
        If Me.QueryString("FT") = "1" Then
            Dim _dt As DataTable = Data(RolesTable)
            _dt.Rows.Clear()

            Dim _webobj As New ZukamiLib.WebSession(GetZukamiSettings)
            _webobj.OpenConnection()
            Dim _ug As DataSet = _webobj.UsersGroups_Get(New Guid(CStr(Data(UserID))))
            Dim i As Integer = 0
            For i = 0 To _ug.Tables(0).Rows.Count - 1
                Dim _GroupID As String = GlobalFunctions.FormatData(_ug.Tables(0).Rows(i).Item("GroupID"))
                Dim _dr As DataRow = _dt.NewRow()
                _dr.Item("ID") = Guid.NewGuid
                _dr.Item(RoleColumn) = New Guid(_GroupID)
                _dt.Rows.Add(_dr)
            Next
            _webobj.CloseConnection()
            _webobj = Nothing

            Data(RolesTable) = _dt
        End If


    End Sub

    Public Sub UpdateRoleTable(RolesTable As String, RoleColumn As String, UserID As String)
        Dim _webobj As New ZukamiLib.WebSession(GetZukamiSettings)
        _webobj.OpenConnection()
        _webobj.GroupMembership_Clear(New Guid(CStr(Data(UserID))))

        Dim _dt As DataTable = Data(RolesTable)
        For i = 0 To _dt.Rows.Count - 1
            Dim _roleID As Guid = GlobalFunctions.GetGUID(_dt.Rows(i).Item(RoleColumn))
            _webobj.GroupMembership_Insert(New Guid(CStr(Data(UserID))), _roleID)
        Next

        _webobj.CloseConnection()
        _webobj = Nothing

    End Sub


    Public ReadOnly Property ControlID(fieldName As String) As String
        Get
            If _zfields.Contains(fieldName) = False Then Return Nothing
            Dim _temp As ZField = _zfields.Item(fieldName)
            Try
                Return CType(_temp.FieldControl, Object).ClientID
            Catch ex As Exception
                Return ""
            End Try
        End Get
    End Property


    'Public Function FT8_PushDocument(ByVal FilePath As String, ByVal FileNameSPC As String, ByVal Extension As String, ByVal Folderpath As String, ByVal Username As String) As String
    '    Try
    '        Dim test As New FT8Service.CheckinServiceSoapClient
    '        FT8_PushDocument = test.fnFilesCheckin(FileNameSPC, Extension, System.IO.File.ReadAllBytes(FilePath), Folderpath, Username)
    '    Catch ex As Exception
    '        Return ex.ToString
    '    End Try
    'End Function

    'Public Function FT_PushProperties(ByVal Databasename As String, ByRef arrFields() As String, ByRef arrValues() As String, ByVal FilePath As String) As String
    '    Try
    '        Dim test As New FT8Service.CheckinServiceSoapClient
    '        FT_PushProperties = test.FnStoringProfValues(Databasename, arrFields, arrValues, FilePath)
    '    Catch ex As Exception
    '        Return ex.ToString
    '    End Try

    'End Function
    Private Sub InsertAllSubformRecs(ByRef webobj As ZukamiLib.WebSession, formID As String, newrecordid As String, parentrecordid As String)
        Dim _set As DataSet = webobj.Lists_GetRecord(New Guid(formID), Nothing)
        If _set.Tables(0).Rows.Count = 0 Then GoTo errEnd
        Dim _tbs As String = GlobalFunctions.FormatData(_set.Tables(0).Rows(0).Item("TableBindSource"))

        Dim _sql As String = "INSERT INTO [" & _tbs & "]([ID],[ParentID],"
        Dim _set2 As DataSet = webobj.ListItems_Get(New Guid(formID))
        Dim i As Integer = 0
        Dim _flist As String = ""
        For i = 0 To _set2.Tables(0).Rows.Count - 1
            Dim _fname As String = GlobalFunctions.FormatData(_set2.Tables(0).Rows(i).Item("FieldBindSource"))
            Dim _listitemid As String = GlobalFunctions.FormatData(_set2.Tables(0).Rows(i).Item("ListItemID"))
            Dim _ftype As GlobalFunctions.FIELDTYPES = GlobalFunctions.FormatData(_set2.Tables(0).Rows(i).Item("FieldType"))
            Select Case _ftype
                Case GlobalFunctions.FIELDTYPES.FT_HEADER, GlobalFunctions.FIELDTYPES.FT_LABEL
                Case Else
                    If Len(_flist) > 0 Then _flist += ","
                    _flist += "[" & _fname & "]"
            End Select


        Next
        _sql += _flist & ") SELECT NEWID(),'" & newrecordid & "'," & _flist & " FROM [" & _tbs & "] WHERE [ParentID]='" & parentrecordid & "'"
        webobj.CustomSQLCommand(_sql)
        webobj.CustomSQLExecute()
errEnd:
    End Sub

    Public Function DuplicateRecord(FormID As String, RecordID As String) As String
        Dim _webobj As New ZukamiLib.WebSession(GetZukamiSettings)
        _webobj.OpenConnection()

        Dim _newrecordid As String = Guid.NewGuid.ToString

        Dim _set As DataSet = _webobj.Lists_GetRecord(New Guid(FormID), Nothing)
        If _set.Tables(0).Rows.Count = 0 Then GoTo errEnd
        Dim _tbs As String = GlobalFunctions.FormatData(_set.Tables(0).Rows(0).Item("TableBindSource"))

        Dim _sql As String = "INSERT INTO [" & _tbs & "]([ID],"
        Dim _set2 As DataSet = _webobj.ListItems_Get(New Guid(FormID))
        Dim i As Integer = 0
        Dim _flist As String = ""
        For i = 0 To _set2.Tables(0).Rows.Count - 1
            Dim _fname As String = GlobalFunctions.FormatData(_set2.Tables(0).Rows(i).Item("FieldBindSource"))
            Dim _listitemid As String = GlobalFunctions.FormatData(_set2.Tables(0).Rows(i).Item("ListItemID"))
            Dim _ftype As GlobalFunctions.FIELDTYPES = GlobalFunctions.FormatData(_set2.Tables(0).Rows(i).Item("FieldType"))
            Select Case _ftype
                Case GlobalFunctions.FIELDTYPES.FT_TABLE
                    InsertAllSubformRecs(_webobj, _listitemid, _newrecordid.ToString, RecordID)
                Case GlobalFunctions.FIELDTYPES.FT_HEADER, GlobalFunctions.FIELDTYPES.FT_LABEL
                Case Else
                    If Len(_flist) > 0 Then _flist += ","
                    _flist += "[" & _fname & "]"
            End Select


        Next
        _sql += _flist & ") SELECT '" & _newrecordid & "'," & _flist & " FROM [" & _tbs & "] WHERE [ID]='" & RecordID & "'"
        _webobj.CustomSQLCommand(_sql)
        _webobj.CustomSQLExecute()

        DuplicateRecord = _newrecordid
errEnd:
        _webobj.CloseConnection()
        _webobj = Nothing

    End Function

    Public Function CreateAdminNote(ByVal ANTemplate As String, ByVal TargetPath As String, ByRef dRow As DataRow) As String
        GemBox.Document.ComponentInfo.SetLicense("DMPX-J9AT-EL54-2YBC")
        Dim document As GemBox.Document.DocumentModel = Nothing

        Dim _template As String = ANTemplate
        Dim sourceDocument As GemBox.Document.DocumentModel = GemBox.Document.DocumentModel.Load(_template, GemBox.Document.DocxLoadOptions.DocxDefault)


        Dim _CallByName As String = GlobalFunctions.FormatData(dRow.Item("Callbyname"))
        Dim _CallOnName As String = GlobalFunctions.FormatData(dRow.Item("Callonname"))
        Dim _Country As String = GlobalFunctions.GetFullCountryName(GlobalFunctions.FormatData(dRow.Item("Country / Organisation")))
        Dim _Designation As String = GlobalFunctions.FormatData(dRow.Item("Designation"))
        Dim _Date As String = GlobalFunctions.FormatData(dRow.Item("Date"))
        Dim _Venue As String = GlobalFunctions.FormatData(dRow.Item("Venue"))
        Dim _dcd As String = GlobalFunctions.FormatData(dRow.Item("Dress Code Description"))
        Dim _so As String = GlobalFunctions.FormatData(dRow.Item("StaffingOfficer"))
        Dim _motdel As String = GlobalFunctions.FormatData(dRow.Item("MOTDelegation"))
        Dim _visdel As String = GlobalFunctions.FormatData(dRow.Item("VisitingDelegation"))



        For Each paragraph As GemBox.Document.Paragraph In sourceDocument.GetChildElements(True, GemBox.Document.ElementType.Paragraph)
            For Each run As GemBox.Document.Run In paragraph.GetChildElements(True, GemBox.Document.ElementType.Run)
                Dim isHeader As Boolean = False
                Dim text As String = run.Text
                If text = "[$CallOn]" Then
                    isHeader = True
                    run.Text = _CallOnName
                    sourceDocument.MailMerge.Execute(paragraph)
                ElseIf text = "[$CallBy]" Then
                    isHeader = True
                    run.Text = _CallByName
                    sourceDocument.MailMerge.Execute(paragraph)
                ElseIf text = "[$Country]" Then
                    isHeader = True
                    run.Text = _Country
                    sourceDocument.MailMerge.Execute(paragraph)
                ElseIf text = "[$Designation]" Then
                    isHeader = True
                    run.Text = _Designation
                    sourceDocument.MailMerge.Execute(paragraph)
                ElseIf text = "[$Date]" Then
                    isHeader = True
                    run.Text = _Date
                    sourceDocument.MailMerge.Execute(paragraph)
                ElseIf text = "[$Venue]" Then
                    isHeader = True
                    run.Text = _Venue
                    sourceDocument.MailMerge.Execute(paragraph)
                ElseIf text = "[$DressCode]" Then
                    isHeader = True
                    run.Text = _dcd
                    sourceDocument.MailMerge.Execute(paragraph)
                ElseIf text = "[$StaffingOfficer]" Then
                    isHeader = True
                    run.Text = _so
                    sourceDocument.MailMerge.Execute(paragraph)
                ElseIf text = "[$MOTDelegation]" Then

                    isHeader = True
                    run.Text = _motdel
                    sourceDocument.MailMerge.Execute(paragraph)
                ElseIf text = "[$VisitingDelegation]" Then
                    isHeader = True
                    run.Text = _visdel
                    sourceDocument.MailMerge.Execute(paragraph)
                End If
            Next
        Next



        sourceDocument.Save(TargetPath, GemBox.Document.DocxSaveOptions.DocxDefault)
        Return TargetPath
    End Function

    Public Function ExternalFuncCall(ExternalPluginName As String, FuncName As String, InParam() As String) As DataSet

    End Function


    Public ReadOnly Property Arguments(ByVal ArgumentName As String) As String
        Get
            If _filterargs Is Nothing Then Return ""
            If _filterargs.Contains(ArgumentName) Then
                Return _filterargs.Item(ArgumentName)
            Else
                Return ""
            End If
        End Get
    End Property

    Public Function CopyandPaste(TargetFormID As String, SourceFormID As String, Optional ByRef FieldNameColl As Collection = Nothing) As Boolean
        Dim _webObj As New ZukamiLib.WebSession(GetZukamiSettings)
        _webObj.OpenConnection()
        CopyandPaste = GlobalFunctions.CopyAndPaste(_webObj, TargetFormID, SourceFormID, FieldNameColl)
        _webObj.CloseConnection()
        _webObj = Nothing


    End Function

    Public Function NewDBQuery(ByVal TableName As String) As DBQuery
        Return New DBQuery(TableName)
    End Function

    Public Function NewExternalCommand(ExternalDatasourceName As String, FunctionName As String, Operation As String) As ExternalDataQuery
        Return New ExternalDataQuery(ExternalDatasourceName, FunctionName, Operation)
    End Function

    Public ReadOnly Property QueryString(ByVal Variable As String) As String
        Get
            Return System.Web.HttpContext.Current.Request.QueryString(Variable)
        End Get
    End Property

    Public Function GenerateAutoID(ByVal FieldID As String) As String
        Dim strFinalID As String = ""


        Dim _webObj As ZukamiLib.WebSession = Nothing
        If KeepConnAlive = True Then
            _webObj = GetSession()
        Else
            _webObj = New ZukamiLib.WebSession(GetZukamiSettings)
            _webObj.OpenConnection()
        End If


        Dim _set As DataSet = _webObj.ListItems_GetRecord(New Guid(FieldID))
        If _set.Tables(0).Rows.Count > 0 Then
            Dim _fa As String = GlobalFunctions.FormatData(_set.Tables(0).Rows(0).Item("FieldArguments"))
            strFinalID = GenerateNewAutoID(_webObj, _fa, FieldID)
        End If

        If KeepConnAlive = False Then
            _webObj.CloseConnection()
            _webObj = Nothing
        End If

        Return strFinalID
    End Function

    Public Function HighlightRows(FieldName As String, ByRef Fields() As String, ByRef Operators() As String, Values() As String, CSSClass As String) As Collection
        Dim _set As DataTable = Data(FieldName)

    End Function

    Public Function HighlightField(FieldName As String, BgColor As String, FgColor As String) As String
        Dim _Ctrl As Control = Control(FieldName)
        Dim i As Integer
        For i = 0 To _Ctrl.Parent.Parent.Controls.Count - 1
            Try
                CType(_Ctrl.Parent.Parent.Controls(i), Object).Style.remove("color")
                CType(_Ctrl.Parent.Parent.Controls(i), Object).Style.remove("background-color")
            Catch ex As Exception

            End Try

            CType(_Ctrl.Parent.Parent.Controls(i), Object).Style.Add("background-color", BgColor)
            CType(_Ctrl.Parent.Parent.Controls(i), Object).Style.Add("color", FgColor)

            Dim td As Control = _Ctrl.Parent.Parent.Controls(i)
            'now we attempt to set the label colors inside the field too
            Dim j As Integer
            For j = 0 To td.Controls.Count - 1
                Dim _ctrl2 As Object = td.Controls(j)
                If StrComp(TypeName(_ctrl2), "Label", vbTextCompare) = 0 Then
                    Try
                        CType(_ctrl2, Object).Style.remove("color")
                        CType(_ctrl2, Object).Style.remove("background-color")
                    Catch ex As Exception

                    End Try
                    Try
                        CType(_ctrl2, Object).Style.add("color", FgColor)
                        CType(_ctrl2, Object).Style.add("background-color", BgColor)
                    Catch ex As Exception

                    End Try

                End If

            Next j


        Next i



    End Function

    Public Function GenerateRunningNumber(ByVal ID As String) As Long
        GenerateRunningNumber = -1

        Dim _webObject As ZukamiLib.WebSession = Nothing
        If KeepConnAlive = True Then
            _webObject = GetSession()
        Else
            _webObject = New ZukamiLib.WebSession(GetZukamiSettings)
            _webObject.OpenConnection()
        End If

        Dim _Set As DataSet = _webObject.Autonumbers_Get(ID)
        If _Set.Tables(0).Rows.Count > 0 Then
            Dim _idnumber As Integer = GlobalFunctions.FormatInteger(_Set.Tables(0).Rows(0).Item(0))
            GenerateRunningNumber = _idnumber
        End If
        If KeepConnAlive = False Then
            _webObject.CloseConnection()
            _webObject = Nothing
        End If

    End Function

    Public Sub SetFormCaption(ByVal Caption As String)
        If _FormAttributes Is Nothing = False Then
            If _FormAttributes.FormCaptionLabel Is Nothing = False Then
                CType(_FormAttributes.FormCaptionLabel, Label).Text = Caption
            End If
        End If
    End Sub

    Public Sub SetWorkflowID(ByVal WorkflowID As String)
        If _FormAttributes Is Nothing = False Then
            If _FormAttributes.WorkflowIDholder Is Nothing = False Then
                CType(_FormAttributes.WorkflowIDholder, HiddenField).Value = WorkflowID
            End If
        End If
    End Sub

    Public Function IsHoliday(ByVal DateCheck As DateTime, ByRef Hols As Collection) As Boolean
        Dim _counter As Integer
        For _counter = 1 To Hols.Count
            Dim _dayr As String = Hols.Item(_counter)
            If StrComp(Left(_dayr, Len("every ")), "Every ", CompareMethod.Text) = 0 Then
                'its an every
                Dim _replaced As String = Replace(_dayr, "Every ", "")
                If IsDate(_replaced) = True Then
                    Dim _ddate As DateTime = CDate(_replaced)
                    If Day(DateCheck) = Day(_ddate) And Month(DateCheck) = Month(_ddate) Then
                        Return True
                    End If
                End If
            Else
                'Absolute date
                If IsDate(_dayr) Then
                    Dim _absdate As DateTime = CDate(_dayr)
                    If Day(DateCheck) = Day(_absdate) And Month(DateCheck) = Month(_absdate) And Year(DateCheck) = Year(_absdate) Then
                        Return True
                    End If
                End If
            End If

        Next _counter
        Return False
    End Function

    Public Function IsCountedDay(ByVal DateCheck As DateTime, ByRef DayRanges As Collection) As Boolean
        Dim _counter As Integer
        For _counter = 1 To DayRanges.Count
            Dim _dayr As Long = DayRanges.Item(_counter)
            If DateCheck.DayOfWeek = _dayr Then
                Return True
            End If
        Next _counter
        Return False
    End Function


    Public Function AutoResizeFileUploadImage(ByVal newWidth As Integer, ByVal newHeight As Integer, ByVal FieldName As String) As Boolean
        If Len(Data(FieldName)) > 0 Then
            Dim _resizedfilename As String = Guid.NewGuid.ToString
            Dim _temp As String = WebconfigSettings.ZukamiTemp.TrimEnd("\") & "\" & Guid.NewGuid.ToString & System.IO.Path.GetExtension(Data(FieldName))
            If ResizeImage(newWidth, newHeight, GetFullUploadedFilepath(Data(FieldName)), _temp) = True Then
                System.IO.File.Copy(_temp, GetFullUploadedFilepath(Data(FieldName)), True)
                Data(FieldName) = Data(FieldName)
                Try
                    System.IO.File.Delete(_temp)
                Catch ex As Exception
                End Try
            End If
        End If
    End Function

    Public ReadOnly Property LastWorkflowActionButtonClicked() As String
        Get
            If _FormAttributes Is Nothing = False Then
                Return _FormAttributes.FormObject.LastWorkflowActionButtonClicked
            Else
                Return ""
            End If
        End Get
    End Property


    Public Function CreateESignature(ByVal UserID As Guid) As String
        Dim _lstring As String = ""
        For _counter = 1 To _zfields.Count
            Dim _temp As ZField = _zfields.Item(_counter)
            Try
                Select Case _temp.FieldType
                    Case GlobalFunctions.FIELDTYPES.FT_SHORTTEXT
                        _lstring += CStr(CType(_temp.FieldControl, TextBox).Text) & ","
                    Case GlobalFunctions.FIELDTYPES.FT_LONGTEXT
                        _lstring += CStr(CType(_temp.FieldControl, TextBox).Text) & ","
                    Case GlobalFunctions.FIELDTYPES.FT_HTML
                        _lstring += CStr(_temp.FieldControl.value) & ","
                    Case GlobalFunctions.FIELDTYPES.FT_COUNTRY
                        _lstring += CStr(CType(_temp.FieldControl, DropDownList).SelectedValue) & ","
                    Case GlobalFunctions.FIELDTYPES.FT_CURRENCY, GlobalFunctions.FIELDTYPES.FT_FLOAT
                        _lstring += GlobalFunctions.FormatDouble(CType(_temp.FieldControl, TextBox).Text) & ","
                    Case GlobalFunctions.FIELDTYPES.FT_INT
                        _lstring += GlobalFunctions.FormatInteger(CType(_temp.FieldControl, TextBox).Text) & ","
                    Case GlobalFunctions.FIELDTYPES.FT_CHECKLIST
                        _lstring += CStr(GlobalFunctions.GetCheckListValue(CType(_temp.FieldControl, Collection))) & ","
                    Case GlobalFunctions.FIELDTYPES.FT_RADIO
                        _lstring += CStr(GlobalFunctions.GetRadioListValue(CType(_temp.FieldControl, Collection))) & ","
                    Case GlobalFunctions.FIELDTYPES.FT_DATE, GlobalFunctions.FIELDTYPES.FT_DATETIME
                        If _temp.FieldControl.isempty Then
                            _lstring += "" & ","
                        Else
                            _lstring += GlobalFunctions.GetDateTime(CType(_temp.FieldControl, Object).value) & ","
                        End If
                    Case GlobalFunctions.FIELDTYPES.FT_DROPDOWN
                        _lstring += CType(_temp.FieldControl, Object).text & ","
                    Case GlobalFunctions.FIELDTYPES.FT_YESNO
                        _lstring += GlobalFunctions.FormatBoolean(CType(_temp.FieldControl, DropDownList).SelectedValue) & ","
                    Case GlobalFunctions.FIELDTYPES.FT_AUTOID, GlobalFunctions.FIELDTYPES.FT_LABEL, GlobalFunctions.FIELDTYPES.FT_DBLABEL
                        _lstring += CStr(CType(_temp.FieldControl, Label).Text) & ","
                    Case GlobalFunctions.FIELDTYPES.FT_FILE
                        _lstring += CStr(CType(_temp.FieldControl, Object).Getinternalpath) & ","
                    Case GlobalFunctions.FIELDTYPES.FT_HIDDENFIELD
                        _lstring += CStr(CType(_temp.FieldControl, HiddenField).Value) & ","
                    Case GlobalFunctions.FIELDTYPES.FT_USER
                        _lstring += CStr(CType(_temp.FieldControl, Object).text) & ","
                    Case GlobalFunctions.FIELDTYPES.FT_IMAGE, GlobalFunctions.FIELDTYPES.FT_TIFFVIEWER, GlobalFunctions.FIELDTYPES.FT_SIGNATURE, GlobalFunctions.FIELDTYPES.FT_CAMERA
                        _lstring += CStr(CType(_temp.FieldControl, Object).GetInternalPath) & ","
                    Case GlobalFunctions.FIELDTYPES.FT_TABLE
                End Select
            Catch ex As Exception
            End Try
        Next _counter



        Dim _sig As String = _lstring & "," & UserID.ToString
        Return OneWayHash(_sig)
    End Function

    Public Sub RotateImage(FullFilePath As String, TargetPath As String)
        'get the path to the image




        'create an image object from the image in that path
        Dim img As System.Drawing.Image = System.Drawing.Image.FromFile(FullFilePath)

        'rotate the image
        img.RotateFlip(RotateFlipType.Rotate90FlipNone)

        'save the image out to the file
        img.Save(TargetPath)

        'release image file
        img.Dispose()
        img = Nothing

    End Sub

    Public Sub FlipImage(FullFilePath As String, TargetPath As String, FlipHorizontal As Boolean)
        'get the path to the image
        'create an image object from the image in that path
        Dim img As System.Drawing.Image = System.Drawing.Image.FromFile(FullFilePath)

        'rotate the image
        If FlipHorizontal = True Then
            img.RotateFlip(RotateFlipType.RotateNoneFlipX)
        Else
            img.RotateFlip(RotateFlipType.RotateNoneFlipY)
        End If


        'save the image out to the file
        img.Save(TargetPath)

        'release image file
        img.Dispose()
        img = Nothing

    End Sub

    Public Function ResizeImage(ByVal newWidth As Integer, ByVal newHeight As Integer, ByVal FromFilePath As String, ByVal ToFilePath As String) As Boolean
        Using image__1 = Image.FromFile(FromFilePath)
            If image__1.Width <> newWidth Or image__1.Height <> newHeight Then
                Using thumbnailBitmap = New Bitmap(newWidth, newHeight)
                    Using thumbnailGraph = Graphics.FromImage(thumbnailBitmap)
                        thumbnailGraph.CompositingQuality = CompositingQuality.HighQuality
                        thumbnailGraph.SmoothingMode = SmoothingMode.HighQuality
                        thumbnailGraph.InterpolationMode = InterpolationMode.HighQualityBicubic

                        Dim imageRectangle = New Rectangle(0, 0, newWidth, newHeight)
                        thumbnailGraph.DrawImage(image__1, imageRectangle)

                        thumbnailBitmap.Save(ToFilePath, image__1.RawFormat)
                        thumbnailGraph.Dispose()
                        thumbnailBitmap.Dispose()
                        image__1.Dispose()
                        Return True
                    End Using

                End Using
            End If
            image__1.Dispose()
        End Using
        Return False
    End Function


    Public Function CalcDateDiff(ByVal Denomination As String, ByVal StartDate As DateTime, ByVal EndDate As DateTime, ByVal DayRange As String, ByVal TimeRange As String, ByVal ExclusionList As String) As Long
        Dim _time As Long = DateDiff(Denomination, StartDate, EndDate)
        Dim HolColl As New Collection
        Dim DayRangeColl As New Collection
        Dim TimeRangeColl As New Collection

        Dim _arrDays() As String = Split(ExclusionList, ",")
        Dim _Counter As Integer
        For _Counter = 0 To UBound(_arrDays)
            Dim _Day As String = Trim(_arrDays(_Counter))
            If Len(_Day) > 0 Then
                HolColl.Add(_Day)
            End If
        Next _Counter

        Dim _TimeRanges() As String = Split(TimeRange, ",")
        Dim _Counter3 As Integer
        For _Counter3 = 0 To UBound(_TimeRanges)
            Dim _range As String = Trim(_TimeRanges(_Counter3))
            If Len(_range) > 0 Then
                TimeRangeColl.Add(_range)
            End If
        Next _Counter3

        Dim dr() As String = Split(DayRange, ";")
        For _Counter = 0 To UBound(dr)
            Dim desc As String = Trim(dr(_Counter))
            Dim daytoinclude As Long = -1
            Select Case LCase(desc)
                Case "mon"
                    daytoinclude = DayOfWeek.Monday
                Case "tue"
                    daytoinclude = DayOfWeek.Tuesday
                Case "wed"
                    daytoinclude = DayOfWeek.Wednesday
                Case "thu"
                    daytoinclude = DayOfWeek.Thursday
                Case "fri"
                    daytoinclude = DayOfWeek.Friday
                Case "sat"
                    daytoinclude = DayOfWeek.Saturday
                Case "sun"
                    daytoinclude = DayOfWeek.Sunday
            End Select
            If daytoinclude <> -1 Then
                DayRangeColl.Add(daytoinclude)
            End If
        Next _Counter


        If DayRangeColl.Count = 0 Then
            DayRangeColl.Add(DayOfWeek.Monday)
            DayRangeColl.Add(DayOfWeek.Tuesday)
            DayRangeColl.Add(DayOfWeek.Wednesday)
            DayRangeColl.Add(DayOfWeek.Thursday)
            DayRangeColl.Add(DayOfWeek.Friday)
            DayRangeColl.Add(DayOfWeek.Saturday)
            DayRangeColl.Add(DayOfWeek.Sunday)
        End If

        If StrComp(Denomination, "d", CompareMethod.Text) = 0 Then
            'Days, just deduct number of weekends
            Dim _temp As DateTime = StartDate
            Do While _temp <= EndDate
                If IsHoliday(_temp, HolColl) Or IsCountedDay(_temp, DayRangeColl) = False Then
                    _time -= 1
                End If
                _temp = _temp.AddDays(1)
            Loop
            Return _time
        ElseIf StrComp(Denomination, "h", CompareMethod.Text) = 0 Then
            'Hours
            Dim _temp As DateTime = StartDate
            Dim _hours As Integer = 0
            Do While _temp < EndDate
                If IsHoliday(_temp, HolColl) = False And IsCountedDay(_temp, DayRangeColl) = True Then
                    If IsInHourRange(_temp, TimeRangeColl) Then
                        _hours += 1
                    End If
                End If
                _temp = _temp.AddHours(1)
            Loop
            Return _hours
        Else
            Return _time
        End If
    End Function

    Public Function DoPlagiarismCheck(ByVal Text1 As String, ByVal Text2 As String) As Long
        Dim s1 As String = Text1
        Dim s2 As String = Text2

        Dim _finallen As Double = Math.Max(Len(s1), Len(s2))
        Dim _discreps As Double = LevenshteinDistance.Compute(s1, s2) - 1
        Return ((((_finallen - _discreps) / _finallen) * 100))
    End Function

    Public Function IsInHourRange(ByVal DateValue As DateTime, ByRef TimeRangeColl As Collection) As Boolean
        If TimeRangeColl.Count = 0 Then Return True
        Dim _counter As Integer
        For _counter = 1 To TimeRangeColl.Count
            Dim _range As String = TimeRangeColl.Item(_counter)
            Dim arrsplits() As String = Split(_range, "-")
            If UBound(arrsplits) = 1 Then
                If IsDate(arrsplits(0)) And IsDate(arrsplits(1)) Then

                    If DateValue.Hour >= CDate(arrsplits(0)).Hour And DateValue.Hour <= CDate(arrsplits(1)).Hour Then
                        Return True
                    End If
                End If
            End If
        Next _counter
        Return False
    End Function

    Public Sub SetFilterArgs(ByRef FA As Collection)
        _filterargs = FA
    End Sub


    Public Function GetParentToken() As String
        Return "ParentSession_" & GetUniqueParentFormID() & "_" & CurrentForm.ParentRecordID
    End Function

    Public Function GetUniqueParentFormID() As String
        If _FormAttributes Is Nothing = False Then
            If Len(_FormAttributes.FormParentID) > 0 Then
                Return _FormAttributes.FormParentID
            Else
                Return _FormAttributes.FormID
            End If
        Else
            Return ""
        End If

    End Function

    Public Sub SetFormDescription(ByVal Description As String)
        If _FormAttributes Is Nothing = False Then
            If _FormAttributes.FormDescriptionLabel Is Nothing = False Then
                CType(_FormAttributes.FormDescriptionLabel, Label).Text = Description
                If Len(Description) > 0 Then
                    CType(_FormAttributes.FormDescriptionLabel, Label).Visible = True
                End If

            End If
        End If
    End Sub

    Public Function CreateBlynkImageLink(ByVal FileTag As String) As String
        If Len(FileTag) = 0 Then Return ""
        Dim arrfields() As String = Split(FileTag, ";")
        Dim fileguid As String = arrfields(0)
        Dim filename As String = arrfields(1)
        Dim target As String = BlynkImagePath
        Dim targetfilefolder As String = ""
        targetfilefolder = System.IO.Path.Combine(target, fileguid)
        Try
            System.IO.Directory.CreateDirectory(targetfilefolder)
        Catch ex As Exception
        End Try

        System.IO.File.Copy(GetFullUploadedFilepath(FileTag), System.IO.Path.Combine(targetfilefolder, filename), True)
        Return Replace(BlynkImageURL.TrimEnd("/") & "/" & fileguid & "/" & filename, " ", "%20", , , CompareMethod.Text)
    End Function

    Public Function InvokeDLL(ByVal DLLPath As String, ByVal NamespaceAndClassName As String, ByVal Methodname As String, ByRef ArgumentArray() As Object, Optional ByRef ErrorMsg As String = "") As Object

        Dim _fullpath As String = WebconfigSettings.BasePath.TrimEnd("\") & "\bin\" & DLLPath

        Dim a As Assembly = Nothing
        Try
            a = Assembly.LoadFile(_fullpath)
            Dim classType As Type = a.GetType(NamespaceAndClassName)
            Dim obj As Object = Activator.CreateInstance(classType)

            Dim mi As MethodInfo = classType.GetMethod(Methodname)
            InvokeDLL = mi.Invoke(obj, ArgumentArray)
        Catch ex As Exception
            ErrorMsg = ex.ToString
        End Try


    End Function

    Public Function PDF_AddImage(SourcePDF As String, ImagePath As String, OutputPath As String, PageNumber As Integer, ImgWidth As Integer, ImgHeight As Integer, AbsPositionX As Integer, AbsPositionY As Integer) As String
        Try
            Using pdfStream As System.IO.Stream = New System.IO.FileStream(SourcePDF, System.IO.FileMode.Open)
                Using imageStream As System.IO.Stream = New System.IO.FileStream(ImagePath, System.IO.FileMode.Open)
                    Using newpdfStream As System.IO.Stream = New System.IO.FileStream(OutputPath, System.IO.FileMode.Create, System.IO.FileAccess.ReadWrite)
                        Dim pdfReader As New iTextSharp.text.pdf.PdfReader(pdfStream)
                        Dim pdfStamper As New iTextSharp.text.pdf.PdfStamper(pdfReader, newpdfStream)
                        Dim pdfContentByte As iTextSharp.text.pdf.PdfContentByte = pdfStamper.GetOverContent(PageNumber)

                        Dim image As iTextSharp.text.Image = iTextSharp.text.Image.GetInstance(imageStream)
                        image.SetAbsolutePosition(AbsPositionX, AbsPositionY)
                        image.ScaleToFit(ImgWidth, ImgHeight)
                        pdfContentByte.AddImage(image)
                        pdfStamper.Close()
                    End Using
                End Using
            End Using
            Return ""
        Catch ex As Exception
            Return ex.ToString
        End Try

    End Function




    Public Sub SetFormHeaderCSS(ByVal CSSStyle As String)
        If _FormAttributes Is Nothing = False Then
            If _FormAttributes.FormTable Is Nothing = False Then
                _FormAttributes.FormTable.CssClass = CSSStyle
            End If
            If _FormAttributes.FormTableBorder Is Nothing = False Then
                _FormAttributes.FormTableBorder.CssClass = CSSStyle
            End If
        End If
    End Sub

    Public Sub SetFieldLabelCaption(ByVal FieldName As String, ByVal Caption As String)
        Dim _zfield As ZField = GetZField(FieldName)
        If _zfield Is Nothing = False Then
            _zfield.FieldCaption = Caption
            RefreshFieldLabel(_zfield)
        Else
            logger.Warn("_zfield is nothing, failed")
        End If
    End Sub

    Private Sub RefreshFieldLabel(ByRef ZFieldObj As ZField)
        If ZFieldObj.FieldLabelControl Is Nothing = False Then
            If TypeOf ZFieldObj.FieldLabelControl Is Label Then
                CType(ZFieldObj.FieldLabelControl, Label).Text = IIf(ZFieldObj.IsHighlighted, "<a name='JumpHere'></a><img border='' src='images\icoFingerAnim.gif' />", "") & IIf(ZFieldObj.IsCompulsory, "* ", "") & ZFieldObj.FieldCaption
            End If
        End If
    End Sub

    Private Sub SetFieldLabelCaptionTemp(ByVal FieldName As String, ByVal Caption As String)
        Dim _zfield As ZField = GetZField(FieldName)
        If _zfield Is Nothing = False Then
            If _zfield.FieldLabelControl Is Nothing = False Then
                If TypeOf _zfield.FieldLabelControl Is Label Then
                    CType(_zfield.FieldLabelControl, Label).Text = Caption
                End If
            End If
        End If

    End Sub

    Public Sub FixFileUploads()
        Dim _scriptkey539167 As String = "UnhideAttachment"
        Dim _script539167 As String = "var table =document.all('ctl00_ContentPlaceHolder1_FormEngine1_MainTable'); " &
            "var rows = table.getElementsByTagName('tr'); " &
            "for (var i=0; i<rows.length;i++){ " &
            "   var cels = rows[i].getElementsByTagName('td'); " &
            "   for (var j=0; j<cels.length;j++){ " &
            "       var ObjInput=cels[j].innerHTML; " &
            "       if (ObjInput.includes('fileupload_normal')){" &
            "           rows[i].style.display='none'; " &
            "       }" &
            "   }" &
            "}"
        Dim _addScriptTags539167 As Boolean = True
        RegisterStartupScript(_scriptkey539167, _script539167, _addScriptTags539167)
    End Sub

    Public Sub SetFieldLabelDescription(ByVal FieldName As String, ByVal Description As String)
        Dim _zfield As ZField = GetZField(FieldName)
        If _zfield Is Nothing = False Then
            If _zfield.FieldLabelDescription Is Nothing = False Then
                CType(_zfield.FieldLabelDescription, Label).Text = "<br/>" & Description
            End If
        End If
    End Sub

    Public Sub SetFieldLabelCSS(ByVal FieldName As String, ByVal CSSStyle As String)
        Dim _zfield As ZField = GetZField(FieldName)
        If _zfield Is Nothing = False Then
            Dim _td As TableCell = Nothing
            Try
                _td = CType(_zfield.FieldLabelControl, Control).Parent
            Catch ex As Exception
            End Try
            If _td Is Nothing = False Then
                Dim _tr As TableRow = _td.Parent
                If _tr Is Nothing = False Then
                    Dim _counter As Integer
                    For _counter = 0 To _tr.Cells.Count - 1
                        _tr.Cells.Item(_counter).CssClass = CSSStyle
                    Next _counter
                End If
            End If
        End If
    End Sub

    Public Sub SetFieldSectionCSS(ByVal FieldName As String, ByVal CSSStyle As String)
        Dim _zfield As ZField = GetZField(FieldName)
        If _zfield Is Nothing = False Then
            Dim _td As TableCell = Nothing
            Try
                _td = CType(_zfield.FieldLabelControl, Control).Parent
            Catch ex As Exception
            End Try
            If _td Is Nothing = False Then
                Dim _tr As TableRow = _td.Parent
                If _tr Is Nothing = False Then
                    Dim _counter As Integer
                    For _counter = 0 To _tr.Cells.Count - 1
                        _tr.Cells.Item(_counter).CssClass = CSSStyle
                    Next _counter
                End If


                'We set the before and after
                Dim tbl As Table = _tr.Parent
                Dim i As Integer
                For i = 0 To tbl.Rows.Count - 1
                    Try
                        If tbl.Rows(i) Is _tr Then
                            Dim _trtoset As TableRow = tbl.Rows(i + 1)
                            For _counter = 0 To _trtoset.Cells.Count - 1
                                _trtoset.Cells.Item(_counter).CssClass = CSSStyle
                            Next _counter

                            _trtoset = tbl.Rows(i - 1)
                            For _counter = 0 To _trtoset.Cells.Count - 1
                                _trtoset.Cells.Item(_counter).CssClass = CSSStyle
                            Next _counter
                            Exit For

                        End If
                    Catch ex As Exception

                    End Try


                Next i


            End If
        End If
    End Sub

    Public Sub AssociateTable(ByVal ParentTable As String, ByVal ChildTable As String, ByVal ParentID As String)
        Dim _obj As Object = Control(ParentTable)
        _obj.PassTable(Control(ChildTable), 0, ParentID)

    End Sub

    Public Sub HighlightField(ByVal FieldName As String, ByVal Highlight As Boolean)
        Dim _zfield As ZField = GetZField(FieldName)
        If _zfield Is Nothing = False Then
            _zfield.IsHighlighted = Highlight
            RefreshFieldLabel(_zfield)
        End If
    End Sub

    Public Function GetTableAddRowControl(ByVal TableName As String, ByVal ControlName As String) As Object

        Dim _counter As Integer
        Dim _formname As String = ""
        Dim _captionfield As String = ""
        Dim _valuefield As String = ""
        For _counter = 1 To _zfields.Count
            Dim _temp As ZField = _zfields.Item(_counter)
            If StrComp(TableName, _temp.FieldName, CompareMethod.Text) = 0 Then
                Try
                    Dim _objTabularGrid As Object = _temp.FieldControl
                    Return _objTabularGrid.getAddRowControl(ControlName)
                Catch ex As Exception
                End Try
            End If
        Next _counter
        Return Nothing

    End Function

    Public Sub PopulateDropdown(Fieldname As String, ByRef DropdownArr() As String)
        If _zfields.Contains(Fieldname) = False Then Exit Sub
        Dim _temp As ZField = _zfields.Item(Fieldname)
        Dim _ddlist As DropDownList = Nothing
        Try
            _ddlist = CType(_temp.FieldControl, Object).DropDownBox
        Catch ex As Exception
            Exit Sub
        End Try

        Dim _Val As String = ""
        _Val = _ddlist.SelectedValue
        _ddlist.Items.Clear()
        Dim i As Integer
        For i = 0 To UBound(DropdownArr)
            Dim _dp As String = DropdownArr(i)
            Dim arrsplits() As String = Split(_dp, "||")
            If UBound(arrsplits) = 1 Then
                Dim _li As New ListItem(arrsplits(0), arrsplits(1))
                _ddlist.Items.Add(_li)
            End If
        Next
        Try
            _ddlist.SelectedValue = _Val
        Catch ex As Exception

        End Try


    End Sub


    Public Sub ShowErrorPopup(fulltag As String, Optional Animate As Boolean = False, Optional DivID As String = "")
        'Dim _script As String = "$(""<div></div>"").attr('style','" & Replace(Style, "'", "\'") & "').appendTo('body');"
        Dim _script As String = "$(""" & fulltag & """).appendTo('body');"
        If Len(DivID) > 0 And Animate = True Then
            _script += "$(""#" & DivID & """).slideDown();"
        End If
        RegisterStartupScript("ErrorKey", _script, True)
    End Sub


    Public Sub FilterTableDropdownBySQL(ByVal TableName As String, ByVal DropdownFieldName As String, ByVal SQL As String, Optional ByVal Datasource As String = "")

        Dim dsourceweb As ZukamiLib.WebSession = Nothing
        Dim webObj As ZukamiLib.WebSession = Nothing
        If KeepConnAlive = True Then
            webObj = GetSession()
        Else
            webObj = New ZukamiLib.WebSession(GetZukamiSettings)
            webObj.OpenConnection()
        End If



        If Len(Datasource) > 0 Then
            Dim _settings As ZukamiLib.ZukamiSettings = GlobalFunctions.GetDatasourceConnectionString(webObj, CurrentAppID, Datasource)
            dsourceweb = New ZukamiLib.WebSession(_settings)
            dsourceweb.OpenOLEDBConnection(_settings.PrimaryConnectionString)
        End If

        Dim _DDList As Object = Nothing
        Dim _counter As Integer
        Dim _formname As String = ""
        Dim _captionfield As String = ""
        Dim _valuefield As String = ""
        For _counter = 1 To _zfields.Count
            Dim _temp As ZField = _zfields.Item(_counter)
            If StrComp(TableName, _temp.FieldName, CompareMethod.Text) = 0 Then
                Try
                    Dim _objTabularGrid As Object = _temp.FieldControl
                    _DDList = _objTabularGrid.GetAddRowControl(DropdownFieldName)
                    Exit For
                Catch ex As Exception
                    _DDList = Nothing
                End Try
            End If
        Next _counter


        If _DDList Is Nothing Then GoTo errEnd

        _DDList = _DDList.dropdownbox


        Dim _selectedvalue As String = ""
        Dim _listitem As ListItem = _DDList.SelectedItem
        If _listitem Is Nothing = False Then
            _selectedvalue = _listitem.Value

        End If

        Try
            Dim _resultset As DataSet = Nothing
            Dim _sql As String = ""
            _DDList.Items.Clear()

            If Len(Datasource) > 0 Then
                Try
                    _sql = SQL
                    dsourceweb.CustomOLEDBSQLCommand(_sql)
                    dsourceweb.CustomOLEDBClearParameters()
                    _resultset = dsourceweb.CustomOLEDBSQLExecuteReturn
                Catch ex As Exception
                End Try
            Else
                _sql = SQL
                webObj.CustomSQLCommand(_sql)
                webObj.CustomClearParameters()
                _resultset = webObj.CustomSQLExecuteReturn
            End If




            Dim _selectedfound As Boolean = False
            If _resultset.Tables(0).Rows.Count > 0 Then
                'now we load the content into the drop down list

                Dim _counter3 As Integer
                For _counter3 = 0 To _resultset.Tables(0).Rows.Count - 1
                    Dim _row As DataRow = _resultset.Tables(0).Rows(_counter3)
                    Dim _value As String = GlobalFunctions.FormatData(_row.Item(1))

                    Dim _lv As New ListItem(GlobalFunctions.TruncateIt(GlobalFunctions.FormatData(_row.Item(0))), _value)


                    If StrComp(_value, _selectedvalue, CompareMethod.Text) = 0 Then
                        _selectedfound = True
                        _lv.Selected = True
                    End If
                    _DDList.Items.Add(_lv)
                Next _counter3
            End If

            Dim _lv2 As New ListItem("--Please select a value--", "")
            If _selectedfound = False Then
                Try
                    _lv2.Selected = True
                Catch ex As Exception
                End Try
            End If
            _DDList.Items.Add(_lv2)


        Catch ex As Exception
        End Try


errEnd:
        If Len(Datasource) > 0 Then
            dsourceweb.CloseOLEDBConnection()
        End If
        If KeepConnAlive = False Then
            webObj.CloseConnection()
            webObj = Nothing
        End If


    End Sub

    Public Sub Blynk_DuplicateApp(ByVal AppID As String)
        'we duplicate appconfig
        Dim _webObj As ZukamiLib.WebSession = Nothing
        _webObj = New ZukamiLib.WebSession(GetZukamiSettings)
        _webObj.OpenConnection()

        Dim _newappID As Guid = Guid.NewGuid

        'Duplicate appconfig
        _webObj.CustomSQLCommand("INSERT INTO AppConfig(ID,)")
        _webObj.CustomClearParameters()
        _webObj.CustomSQLExecute()

        _webObj.CustomSQLCommand("INSERT INTO Catalog(ID,)")
        _webObj.CustomClearParameters()
        _webObj.CustomSQLExecute()

        _webObj.CustomSQLCommand("INSERT INTO Category(ID,)")
        _webObj.CustomClearParameters()
        _webObj.CustomSQLExecute()

        _webObj.CustomSQLCommand("INSERT INTO MobileHomeTile(ID,)")
        _webObj.CustomClearParameters()
        _webObj.CustomSQLExecute()

        _webObj.CustomSQLCommand("INSERT INTO Product(ID,)")
        _webObj.CustomClearParameters()
        _webObj.CustomSQLExecute()

        _webObj.CustomSQLCommand("INSERT INTO ProductCategories(ID,)")
        _webObj.CustomClearParameters()
        _webObj.CustomSQLExecute()

        _webObj.CustomSQLCommand("INSERT INTO ProductImage(ID,)")
        _webObj.CustomClearParameters()
        _webObj.CustomSQLExecute()

        _webObj.CustomSQLCommand("INSERT INTO Outlet(ID,)")
        _webObj.CustomClearParameters()
        _webObj.CustomSQLExecute()

        _webObj.CustomSQLCommand("INSERT INTO Salesman(ID,)")
        _webObj.CustomClearParameters()
        _webObj.CustomSQLExecute()


        _webObj.CloseConnection()
        _webObj = Nothing
    End Sub

    Public Sub FilterDropdownBySQL(ByVal FieldName As String, ByVal SQL As String, Optional ByVal Datasource As String = "")
        Dim dsourceweb As ZukamiLib.WebSession = Nothing
        Dim _webObj As ZukamiLib.WebSession = Nothing
        If KeepConnAlive = True Then
            _webObj = GetSession()
        Else
            _webObj = New ZukamiLib.WebSession(GetZukamiSettings)
            _webObj.OpenConnection()
        End If

        If Len(Datasource) > 0 Then
            Dim _settings As ZukamiLib.ZukamiSettings = GlobalFunctions.GetDatasourceConnectionString(_webObj, CurrentAppID, Datasource)
            dsourceweb = New ZukamiLib.WebSession(_settings)
            dsourceweb.OpenOLEDBConnection(_settings.PrimaryConnectionString)
        End If



        Dim _DDList As DropDownList = Nothing
        Dim _counter As Integer
        Dim _formname As String = ""
        Dim _captionfield As String = ""
        Dim _valuefield As String = ""
        For _counter = 1 To _zfields.Count
            Dim _temp As ZField = _zfields.Item(_counter)
            If StrComp(FieldName, _temp.FieldName, CompareMethod.Text) = 0 Then
                Try
                    _DDList = _temp.FieldControl.dropdownbox
                    Exit For
                Catch ex As Exception
                    _DDList = Nothing
                End Try
            End If
        Next _counter
        If _DDList Is Nothing Then GoTo errEnd

        Dim _selectedvalue As String = ""
        Dim _listitem As ListItem = _DDList.SelectedItem
        If _listitem Is Nothing = False Then
            _selectedvalue = _listitem.Value
        End If

        Try
            Dim _resultset As DataSet = Nothing
            Dim _sql As String = ""
            _DDList.Items.Clear()

            If Len(Datasource) > 0 Then
                Try
                    _sql = SQL
                    dsourceweb.CustomOLEDBSQLCommand(_sql)
                    dsourceweb.CustomOLEDBClearParameters()
                    _resultset = dsourceweb.CustomOLEDBSQLExecuteReturn
                Catch ex As Exception
                End Try
            Else
                _sql = SQL
                _webObj.CustomSQLCommand(_sql)
                _webObj.CustomClearParameters()
                _resultset = _webObj.CustomSQLExecuteReturn
            End If




            Dim _selectedfound As Boolean = False
            If _resultset.Tables(0).Rows.Count > 0 Then
                'now we load the content into the drop down list

                Dim _counter3 As Integer
                For _counter3 = 0 To _resultset.Tables(0).Rows.Count - 1
                    Dim _row As DataRow = _resultset.Tables(0).Rows(_counter3)
                    Dim _value As String = GlobalFunctions.FormatData(_row.Item(1))

                    Dim _lv As New ListItem(GlobalFunctions.TruncateIt(GlobalFunctions.FormatData(_row.Item(0))), _value)


                    If StrComp(_value, _selectedvalue, CompareMethod.Text) = 0 Then
                        _selectedfound = True
                        _lv.Selected = True
                    End If

                    _DDList.Items.Add(_lv)
                Next _counter3
            End If

            Dim _lv2 As New ListItem("--Please select a value--", "")
            If _selectedfound = False Then
                Try
                    _lv2.Selected = True
                Catch ex As Exception
                End Try
            End If
            _DDList.Items.Add(_lv2)


        Catch ex As Exception
        End Try


errEnd:
        If Len(Datasource) > 0 Then
            dsourceweb.CloseOLEDBConnection()
        End If
        If KeepConnAlive = False Then
            _webObj.CloseConnection()
            _webObj = Nothing
        End If



    End Sub

    Private Function GetAllTabs(GroupName As String, ActiveTabGUID As String) As String
        Dim tabstring As String = ""
        For i = 1 To _zfields.Count
            Dim _zfield As ZField = _zfields.Item(i)
            If _zfield.FieldType = GlobalFunctions.FIELDTYPES.FT_HEADER Then
                Dim _arg As String = _zfield.Arguments
                Dim _istab As Boolean = False
                Dim _groupName As String = ""
                Dim _parentsectionGUID As String = ""
                GlobalFunctions.ExtractSectionArguments(_arg, _istab, _groupName, _parentsectionGUID)
                If _istab = True Then
                    If StrComp(_groupName, GroupName, vbTextCompare) = 0 Then

                        Dim _istabvisible As Boolean = True
                        Dim _tabvisible As String = Me.Session("TV_" & CurrentUser.UserID.ToString & "_" & PrimaryID.ToString & "_" & _zfield.FieldName)


                        If Len(_tabvisible) > 0 Then
                            _istabvisible = GlobalFunctions.FormatBoolean(_tabvisible)
                        End If

                        If _istabvisible = True Then


                            If StrComp(_zfield.FieldGUID.ToString, ActiveTabGUID, vbTextCompare) = 0 Then
                                'this is the active tab
                                If Len(tabstring) > 0 Then tabstring += "&nbsp;&nbsp;|&nbsp;&nbsp;"
                                tabstring += "<span class=""formfieldtabcaption"">" & _zfield.FieldCaption & "</span>"
                            Else
                                'this is the non active tab
                                If Len(tabstring) > 0 Then tabstring += "&nbsp;&nbsp;|&nbsp;&nbsp;"
                                tabstring += "<a href=""#"" class=""formfieldtablink"" onclick=""__doPostBack('ctl00_UpdatePanel1$DefaultButtonClick','R4CHGTAB!_!" & GroupName & "!_!" & _zfield.FieldGUID.ToString & "');"">" & _zfield.FieldCaption & "</a>"
                            End If
                        End If
                    End If
                End If
            End If
        Next i


        Return tabstring
    End Function

    Public Sub Tabify()
        'run through all sections, and try to find the 'master' sections. 
        Dim _ea As String = EventArgument()
        If Len(_ea) > 0 Then
            Dim arrSplits() As String = Split(_ea, "!_!")
            If UBound(arrSplits) >= 2 Then
                If StrComp(arrSplits(0), "R4CHGTAB", vbTextCompare) = 0 Then
                    Dim _groupname As String = arrSplits(1)
                    Dim _activetabguid As String = arrSplits(2)
                    If GlobalFunctions.IsGUID(_activetabguid) And Len(_groupname) > 0 Then
                        Me.Session(CurrentUser.UserID.ToString & "_" & PrimaryID.ToString & "_" & _groupname) = _activetabguid
                    End If
                End If
            End If
        End If



        Dim _allgroups As New Collection
        Dim i As Integer = 0
        For i = 1 To _zfields.Count
            Dim _zfield As ZField = _zfields.Item(i)
            If _zfield.FieldType = GlobalFunctions.FIELDTYPES.FT_HEADER Then
                Dim _arg As String = _zfield.Arguments
                Dim _istab As Boolean = False
                Dim _groupName As String = ""
                Dim _parentsectionGUID As String = ""
                GlobalFunctions.ExtractSectionArguments(_arg, _istab, _groupName, _parentsectionGUID)



                If _istab = True Then
                    If Len(_groupName) > 0 Then
                        If _allgroups.Contains(_groupName) = False Then _allgroups.Add(_groupName, _groupName)

                        Dim _activetab As String = Me.Session(CurrentUser.UserID.ToString & "_" & PrimaryID.ToString & "_" & _groupName)
                        If Len(_activetab) = 0 Then
                            Me.Session(CurrentUser.UserID.ToString & "_" & PrimaryID.ToString & "_" & _groupName) = _zfield.FieldGUID.ToString
                            _activetab = Me.Session(CurrentUser.UserID.ToString & "_" & PrimaryID.ToString & "_" & _groupName)
                        End If

                        If StrComp(_activetab, _zfield.FieldGUID.ToString, vbTextCompare) = 0 Then
                            'this is the active tab, we show
                            SectionVisible(_zfield.BoundField) = True


                            SetFieldSectionCSS(_zfield.BoundField, "formfieldtabseparator_cell")
                            'we also draw all the other child tab names here

                            SetFieldLabelCaptionTemp(_zfield.BoundField, GetAllTabs(_groupName, _activetab))

                        Else
                            SectionVisible(_zfield.BoundField) = False
                        End If
                    End If
                End If

            End If
        Next i
    End Sub

    Public Function ReincrementColumn(TableFieldName As String, ColumnName As String) As Boolean
        Dim dt As DataTable = Data(TableFieldName)
        Dim i As Integer
        For i = 0 To dt.Rows.Count - 1
            dt.Rows(i).Item(ColumnName) = CStr(i + 1)
        Next i
        Data(TableFieldName) = dt
    End Function

    Public Function RScript(script As String) As String

    End Function

    Public Function RemoveColumn(FieldName As String, ColumnName As String) As String
        Dim _tct As Object = Control(FieldName)
        If _tct Is Nothing = False Then
            If StrComp(TypeName(_tct), "usercontrols_tabularlist_ascx", vbTextCompare) = 0 Then
                _tct.removecolumn(ColumnName)
            End If
        End If
    End Function

    Public Function GenerateEventURL(ByVal EventID As Guid, ByVal EventName As String, ByVal EventSubject As String, ByVal EventDescription As String, ByVal EventLocation As String, ByVal EventStart As DateTime, ByVal EventEnd As DateTime, ByVal StartTimeZone As String, ByVal EndTimeZone As String, Optional ByVal LinkCaption As String = "") As String
        Dim _web As ZukamiLib.WebSession = Nothing
        If KeepConnAlive = True Then
            _web = GetSession()
        Else
            _web = New ZukamiLib.WebSession(GetZukamiSettings)
            _web.OpenConnection()
        End If

        _web.Events_Insert(EventID, EventName, EventSubject, EventDescription, EventLocation, EventStart, EventEnd, StartTimeZone, EndTimeZone, LinkCaption)

        If KeepConnAlive = False Then
            _web.CloseConnection()
            _web = Nothing
        End If
        Return CreateLink("GenICS.aspx?EventID=" & EventID.ToString, LinkCaption)
    End Function

    Public Function CreateLink(ByVal LinkURL As String, ByVal LinkText As String) As String
        Return "<a href=""" & LinkURL & """ class=""createics_link"">" & LinkText & "</a>"
    End Function


    Public Sub SetCompulsory(ByVal Fieldname As String, ByVal IsCompulsory As Boolean)
        'Here we set whether compulsory or not
        Dim _zfield As ZField = GetZField(Fieldname)
        If _zfield Is Nothing = False Then
            _zfield.IsCompulsory = IsCompulsory
            RefreshFieldLabel(_zfield)
        End If
    End Sub

    'Public Function DuplicateImage(ByVal FileTag As String) As String
    '    Dim _sourcef As String = Me.GetFullUploadedFilepath(FileTag)
    '    Dim _fname As String = System.IO.Path.GetFileName(_sourcef)

    '    Dim _newid As String = Guid.NewGuid.ToString
    '    Dim _newtarget As String = _newid & ";" & _fname

    '    Try
    '        System.IO.Directory.CreateDirectory(WebconfigSettings.UploadPath.TrimEnd("\") & "\" & _newid)
    '        Dim _targetf As String = Me.GetFullUploadedFilepath(_newtarget)

    '        System.IO.File.Copy(_sourcef, _targetf)
    '        Return _targetf
    '    Catch ex As Exception

    '    End Try
    '    Return ""
    'End Function

    Public Sub OpenImage(ByVal FileTag As String)
        If System.IO.File.Exists(Me.GetFullUploadedFilepath(FileTag)) Then
            _GlobalImage = New Bitmap(Me.GetFullUploadedFilepath(FileTag))
            _GlobalGraphics = Graphics.FromImage(_GlobalImage)
            _GlobalGraphics.SmoothingMode = SmoothingMode.AntiAlias
        End If
    End Sub

    Public Sub CloseImage()
        _GlobalGraphics.Dispose()
        _GlobalImage.Dispose()
        _GlobalGraphics = Nothing
        _GlobalImage = Nothing
    End Sub

    Public Sub WriteImageToImage(ByVal x As Integer, ByVal y As Integer, ByVal ImagePath As String)
        Dim _gimage As Image = Image.FromFile(ImagePath)
        _GlobalGraphics.DrawImage(_gimage, x, y)
    End Sub

    Public Function WriteTextToImage(ByVal x As Integer, ByVal y As Integer, ByVal TextToWrite As String, ByVal Font As String, ByVal FontSize As Integer) As String
        Try
            _GlobalGraphics.DrawString(TextToWrite, New Font(Font, FontSize, FontStyle.Regular), SystemBrushes.WindowText, New Point(x, y))
            Return ""
        Catch ex As Exception
            Return ex.ToString
        End Try
    End Function

    Public Sub RegisterStartupScript(ScriptKey As String, Script As String, AddScriptTags As Boolean)
        ScriptManager.RegisterStartupScript(CType(System.Web.HttpContext.Current.Handler, Page), System.Web.HttpContext.Current.Handler.GetType, ScriptKey, Script, AddScriptTags)
    End Sub

    Public Function SaveImage() As String
        Dim _newID As String = Guid.NewGuid.ToString
        Dim _newfolder As String = WebconfigSettings.UploadPath.TrimEnd("\") & "\" & _newID
        Dim _newfpath As String = _newfolder & "\modifiedimage.png"
        System.IO.Directory.CreateDirectory(_newfolder)

        _GlobalImage.Save(_newfpath, System.Drawing.Imaging.ImageFormat.Png)

        Return _newID & ";modifiedimage.png"
    End Function

    Public Function CNTP_TestDLL() As String
        Dim outDT As New StrongTypesNS.tmpclm100DataTable
        Dim outDTuwm100 As New StrongTypesNS.tmpuwm100DataTable
        Dim dr As StrongTypesNS.tmpclm100Row = outDT.Newtmpclm100Row()
        Return dr.ToString

    End Function

    Private Function GenerateNewAutoID(ByRef WebObj As ZukamiLib.WebSession, ByVal Arguments As String, ByVal FieldID As String) As String
        Dim _format As String = Arguments
        Dim _ID As String = FieldID
        Dim _FinalID As String = _format

        'insert
        Dim _Set As DataSet = WebObj.Autonumbers_Get(_ID.ToString)
        If _Set.Tables(0).Rows.Count > 0 Then
            Dim _idnumber As Integer = GlobalFunctions.FormatInteger(_Set.Tables(0).Rows(0).Item(0))


            _FinalID = Replace(_FinalID, "[%d]", CStr(Day(Now)), , , CompareMethod.Text)
            _FinalID = Replace(_FinalID, "[%mm]", CStr(Month(Now)), , , CompareMethod.Text)
            _FinalID = Replace(_FinalID, "[%yy]", Right(CStr(Year(Now)), 2), , , CompareMethod.Text)
            _FinalID = Replace(_FinalID, "[%yyyy]", CStr(Year(Now)), , , CompareMethod.Text)

            Dim intPctg As Integer = InStr(_FinalID, "[", CompareMethod.Text)
            Dim intPctgEnd As Integer
            If intPctg > 0 Then
                intPctgEnd = InStr(intPctg, _FinalID, "]", CompareMethod.Text)
                Dim strTag As String = Mid(_FinalID, intPctg, intPctgEnd - intPctg + 1)
                Dim intDigits As Integer = GlobalFunctions.FormatInteger(Replace(Mid(_FinalID, intPctg + 1, intPctgEnd - intPctg - 1), "%", ""), 1)
                Dim strDigitTag As String = ""
                Dim _counter7 As Integer
                For _counter7 = 1 To intDigits
                    strDigitTag += "0"
                Next _counter7
                _FinalID = Replace(_FinalID, strTag, Format(_idnumber, strDigitTag))
            End If
        End If
        Return _FinalID
    End Function

    Public ReadOnly Property QueryString() As String
        Get
            Return System.Web.HttpContext.Current.Request.QueryString.ToString
        End Get
    End Property

    Public ReadOnly Property Form(ByVal Variable As String) As String
        Get
            Return System.Web.HttpContext.Current.Request.Form(Variable)
        End Get
    End Property

    Public ReadOnly Property Form() As Object
        Get
            Return System.Web.HttpContext.Current.Request.Form
        End Get
    End Property

    Public ReadOnly Property EventTarget() As String
        Get
            Return System.Web.HttpContext.Current.Request.Form("__EVENTTARGET")
        End Get
    End Property

    Public ReadOnly Property EventArgument() As String
        Get
            Return System.Web.HttpContext.Current.Request.Form("__EVENTARGUMENT")
        End Get
    End Property

    Public Sub Redirect(ByVal URL As String)
        System.Web.HttpContext.Current.Response.Redirect(URL)
    End Sub

    Public ReadOnly Property CurrentUser() As UserDetails
        Get
            Dim _webObject As ZukamiLib.WebSession = Nothing
            If KeepConnAlive = True Then
                _webObject = GetSession()
            Else
                _webObject = New ZukamiLib.WebSession(GetZukamiSettings)
                _webObject.OpenConnection()
            End If

            Dim _currentuserguid As Guid = GetZukamiSettings.CurrentUserGUID
            Dim _userSet As DataSet = _webObject.Users_GetRecord(_currentuserguid)
            Dim _roleSet As DataSet = _webObject.Groups_GetByUser(_currentuserguid)

            If KeepConnAlive = False Then
                _webObject.CloseConnection()
                _webObject = Nothing
            End If

            If _userSet.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim _user As New UserDetails(_userSet.Tables(0).Rows(0), _roleSet)
                Return _user
            End If
        End Get
    End Property

    Public Sub CreateEmail(ByVal ToList As String, ByVal CCList As String, ByVal BCCList As String, ByVal Subject As String, ByVal Content As String, Optional ByVal Attachments As String = "")
        Dim _webObject As ZukamiLib.WebSession = Nothing
        If KeepConnAlive = True Then
            _webObject = GetSession()
        Else
            _webObject = New ZukamiLib.WebSession(GetZukamiSettings)
            _webObject.OpenConnection()
        End If

        GlobalFunctions.SendEmail(_webObject, ToList, CCList, BCCList, Subject, Content, Attachments)
        If KeepConnAlive = False Then
            _webObject.CloseConnection()
            _webObject = Nothing
        End If
    End Sub

    Public Sub DestroyReminder(ByVal JobID As Guid)
        Dim _webObject As ZukamiLib.WebSession = Nothing
        If KeepConnAlive = True Then
            _webObject = GetSession()
        Else
            _webObject = New ZukamiLib.WebSession(GetZukamiSettings)
            _webObject.OpenConnection()
        End If


        _webObject.CustomSQLCommand("DELETE FROM [JobTimers] WHERE JobID='" & JobID.ToString & "'")
        _webObject.CustomClearParameters()
        _webObject.CustomSQLExecute()
        If KeepConnAlive = False Then
            _webObject.CloseConnection()
            _webObject = Nothing
        End If
    End Sub

    Public Sub CreateEmailReminder(ByVal JobID As Guid, ByVal ReminderTime As DateTime, ByVal ToList As String, ByVal CCList As String, ByVal BCCList As String, ByVal Subject As String, ByVal Content As String, Optional ByVal Recurrence As String = "")
        Dim _webObject As ZukamiLib.WebSession = Nothing
        If KeepConnAlive = True Then
            _webObject = GetSession()
        Else
            _webObject = New ZukamiLib.WebSession(GetZukamiSettings)
            _webObject.OpenConnection()
        End If

        Dim _settings As String = ToList & GlobalFunctions.JSCRIPTSEPARATOR & CCList & GlobalFunctions.JSCRIPTSEPARATOR & BCCList & GlobalFunctions.JSCRIPTSEPARATOR & Subject
        _webObject.JobTimers_Set(ReminderTime, ZukamiLib.WebSession.JOBTIMERTYPE.JOBTIMER_EMAIL, Guid.Empty, Guid.Empty, JobID, False, Guid.Empty, Guid.Empty, Guid.Empty, Guid.Empty, Guid.Empty, _settings, Content, Recurrence, 0)

        If KeepConnAlive = False Then
            _webObject.CloseConnection()
            _webObject = Nothing
        End If
    End Sub

    Public Sub CreateSMSReminder(ByVal JobID As Guid, ByVal ReminderTime As DateTime, ByVal PhoneNumber As String, ByVal Message As String, Optional ByVal Recurrence As String = "")
        Dim _webObject As ZukamiLib.WebSession = Nothing
        If KeepConnAlive = True Then
            _webObject = GetSession()
        Else
            _webObject = New ZukamiLib.WebSession(GetZukamiSettings)
            _webObject.OpenConnection()
        End If

        _webObject.JobTimers_Set(ReminderTime, ZukamiLib.WebSession.JOBTIMERTYPE.JOBTIMER_SMS, Guid.Empty, Guid.Empty, JobID, False, Guid.Empty, Guid.Empty, Guid.Empty, Guid.Empty, Guid.Empty, PhoneNumber, Message, Recurrence, 0)

        If KeepConnAlive = False Then
            _webObject.CloseConnection()
            _webObject = Nothing
        End If
    End Sub

    Public Function URLEncode(ByVal URL As String) As String
        Return System.Web.HttpContext.Current.Server.UrlEncode(URL)
    End Function

    Public Sub RefreshTable(ByVal Fieldname As String)
        Dim _counter As Integer
        For _counter = 1 To _zfields.Count
            Dim _temp As ZField = _zfields.Item(_counter)
            Try
                If StrComp(Fieldname, _temp.FieldName, CompareMethod.Text) = 0 Then
                    Select Case _temp.FieldType
                        Case GlobalFunctions.FIELDTYPES.FT_TABLE
                            CType(_temp.FieldControl, Object).refresh()
                            Data(Fieldname) = CType(_temp.FieldControl, Object).datasource
                    End Select
                    Exit Sub
                End If
            Catch ex As Exception
            End Try
        Next _counter
    End Sub

    Public Property SubmissionDetails() As SubmissionDetails
        Get
            Return _SubmissionDetails
        End Get
        Set(ByVal value As SubmissionDetails)
            _SubmissionDetails = value
        End Set
    End Property

    Public Property DataBag() As DataSet
        Get
            Return _DataBag
        End Get
        Set(ByVal value As DataSet)
            _DataBag = value
        End Set
    End Property

    Public Sub SetTableRowClickBehavior(TableFieldName As String, AllowRowClick As Boolean, Optional CustomOnClickCode As String = "", Optional FixedQueryString As String = "", Optional NavigateURL As String = "")
        Dim tbcontrol As Object = Control(TableFieldName)
        If tbcontrol Is Nothing = False Then
            CType(tbcontrol, Object).AllowRowClick = AllowRowClick
            CType(tbcontrol, Object).CustomOnClickCode = CustomOnClickCode
            CType(tbcontrol, Object).FixedQuerystring = FixedQueryString
            CType(tbcontrol, Object).NavigateURL = NavigateURL
            CType(tbcontrol, Object).refreshgrid

        End If
    End Sub

    Public Function CreateFormField(ByRef FormGUID As Guid, ByVal FieldName As String, ByVal FieldCaption As String, ByVal HelperText As String, ByVal FieldType As Integer, Optional ByVal Compulsory As Boolean = False, Optional ByVal Enabled As Boolean = True, Optional ByVal DefaultValue As String = "", Optional ByVal FieldArguments As String = "", Optional ByVal AllowDuplicates As Boolean = True, Optional ByVal CSS As String = "", Optional ByVal LabelCSS As String = "", Optional ByVal FocusCSS As String = "", Optional ByVal ReadonlyCSS As String = "", Optional ByVal FieldOrder As Integer = 0, Optional ByVal CreateViewColumn As Boolean = False) As Boolean
        Dim webObj As ZukamiLib.WebSession = Nothing
        Dim retguid As String = ""
        webObj = New ZukamiLib.WebSession(GetZukamiSettings)
        webObj.OpenConnection()

        CreateFormField = GlobalFunctions.CreateFormField(webObj, FormGUID, FieldName, FieldCaption, HelperText, FieldType, Compulsory, Enabled, DefaultValue, FieldArguments, AllowDuplicates, CSS, LabelCSS, FocusCSS, ReadonlyCSS, FieldOrder, retguid)

        If CreateViewColumn = True Then
            If GlobalFunctions.IsGUID(retguid) = True Then
                Dim _set As DataSet = webObj.Lists_GetRecord(FormGUID, Nothing)
                Dim _tbs As String = GlobalFunctions.FormatData(_set.Tables(0).Rows(0).Item("TableBindSource"))

                Dim _viewrecord As DataSet = webObj.ListItems_GetRecord(FormGUID)
                If _viewrecord.Tables(0).Rows.Count > 0 Then
                    Dim _fargs As String = GlobalFunctions.FormatData(_viewrecord.Tables(0).Rows(0).Item("FieldArguments"))
                    Dim _tvdata As String = ""
                    GlobalFunctions.ExtractTableArgs(_fargs, _tvdata, False, False)
                    If GlobalFunctions.IsGUID(_tvdata) = True Then
                        'search for listitem where listitemid = request.querystring("ID")
                        Dim forder As Integer = 0
                        If FieldOrder = -1 Then
                            webObj.CustomSQLCommand("SELECT Max(ColOrder) FROM ViewColumns WHERE ViewID='" & _tvdata & "'")
                            webObj.CustomClearParameters()
                            Dim _vcset As DataSet = webObj.CustomSQLExecuteReturn
                            If _vcset.Tables(0).Rows.Count > 0 Then
                                forder = GlobalFunctions.FormatInteger(_vcset.Tables(0).Rows(0).Item(0)) + 1
                            End If
                            _vcset.Dispose()
                            _vcset = Nothing
                        End If

                        webObj.ViewColumns_Insert(Guid.NewGuid, FormGUID, New Guid(retguid), FieldCaption, GlobalFunctions.GetGUID(_tvdata), forder, 0, FieldName, _tbs, False, True, FieldType, "", "", "")
                    End If
                End If


            End If
        End If



        webObj.CloseConnection()
        webObj = Nothing

    End Function

    Public Function ConvertDocxToPDF(ByVal Sourcepath As String, ByVal TargetPath As String) As String
        Return GlobalFunctions.ConvertDocxToPDF(Sourcepath, TargetPath)
    End Function

    Public Function GenerateReport(ReportID As Guid, OutputPath As String, ExportFormat As String, Optional ByRef QS As Collection = Nothing) As String
        Dim _webobj As New ZukamiLib.WebSession(GetZukamiSettings)
        _webobj.OpenConnection()
        Dim _Viewset As DataSet = _webobj.Views_GetRecord(ReportID)
        If _Viewset.Tables(0).Rows.Count > 0 Then
            Dim _lc As String = GlobalFunctions.FormatData(_Viewset.Tables(0).Rows(0).Item("LayoutContent"))
            Dim test As New DevExpress.XtraReports.UI.XtraReport()
            If Len(_lc) > 0 Then
                Dim arrBytes() As Byte = Convert.FromBase64String(_lc)
                test.LoadLayout(New System.IO.MemoryStream(arrBytes))

                Dim _viewid As Guid = ReportID
                Dim _sql As String = ""
                Dim _filtercoll As New Collection
                Dim _set As DataSet = GlobalFunctions.loadViewDatasource(_viewid, _filtercoll, _sql)
                test.Extensions(DevExpress.XtraReports.Native.SerializationService.Guid) = DSSerializer.Name

                DSSerializer.ApplyDSet(_set)
                test.DataSource = DSSerializer.GenerateDSet
                test.DataMember = _set.Tables(0).TableName



                LoadParams(test, QS)
                test.CreateDocument()

                Select Case (LCase(ExportFormat))
                    Case "png"
                        test.ExportToImage(OutputPath)
                    Case "pdf"
                        test.ExportToPdf(OutputPath)
                    Case "csv"
                        test.ExportToCsv(OutputPath)
                    Case "html"
                        test.ExportToHtml(OutputPath)
                    Case "mht"
                        test.ExportToMht(OutputPath)
                    Case "rtf"
                        test.ExportToRtf(OutputPath)
                    Case "txt"
                        test.ExportToText(OutputPath)
                    Case "xls"
                        test.ExportToXls(OutputPath)
                    Case "xlsx"
                        test.ExportToXlsx(OutputPath)
                End Select

            End If




        End If



        _webobj.CloseConnection()
        _webobj = Nothing

    End Function

    Private Shared Sub LoadParams(ByRef rep As DevExpress.XtraReports.UI.XtraReport, ByRef QS As Collection)
        If QS Is Nothing Then Exit Sub
        Dim i As Integer = 0
        Dim _mustrequest As Boolean = False
        For i = 0 To rep.Parameters.Count - 1
            Dim _param As DevExpress.XtraReports.Parameters.Parameter = rep.Parameters.Item(i)

            If QS.Contains(_param.Name) = True Then
                _param.Value = QS.Item(_param.Name)
                _param.Visible = False

            Else
                _mustrequest = True
                _param.Visible = False
                '_param.Visible = True
            End If

        Next

        rep.RequestParameters = False


    End Sub

    Public Function Resource(TagName As String, Optional CheckLanguage As Boolean = True) As String
        If CheckLanguage = True Then
            Return GetSubstitutionTag(TagName & "." & CurrentUser.Language, TagName)
        Else
            Return GetSubstitutionTag(TagName)
        End If
    End Function

    Private Function GetSubstitutionTag(AttemptTagName1 As String, Optional AttemptTagName2 As String = "") As String
        If Len(AttemptTagName1) = 0 Then Return ""
        Dim _settings As ZukamiLib.ZukamiSettings = GetZukamiSettings()
        Dim _webObj As New ZukamiLib.WebSession(_settings)
        _webObj.OpenConnection()
        Dim _MainDataset As DataSet = GetSubString(_webObj, AttemptTagName1)
        If _MainDataset.Tables(0).Rows.Count = 0 Then
            If Len(AttemptTagName2) > 0 Then
                _MainDataset = GetSubString(_webObj, AttemptTagName2)
            End If
        End If

        If _MainDataset.Tables(0).Rows.Count = 0 Then
            GetSubstitutionTag = ""
        Else
            Dim _myval As String = GlobalFunctions.FormatData(_MainDataset.Tables(0).Rows(0).Item("Value"))
            If Len(_myval) > 3 Then
                Dim _key As String = LCase(Mid(_myval, 1, 3))
                Dim _data As String = Mid(_myval, 4, Len(_myval) - 3)

                If StrComp(_key, "ss,", vbTextCompare) = 0 Then
                    GetSubstitutionTag = _data
                End If
            End If
        End If


        _webObj.CloseConnection()
        _webObj = Nothing
    End Function

    Private Function GetSubString(ByRef WebObj As ZukamiLib.WebSession, SSTag As String) As DataSet
        Dim _Sql As String = "SELECT [StringName],[Value] FROM [SubstitutionStrings] WHERE [AppID]='" & CurrentAppID.ToString & "' AND [StringName]='" & GlobalFunctions.FormatSQLData(SSTag) & "'"
        WebObj.CustomSQLCommand(_Sql)
        GetSubString = WebObj.CustomSQLExecuteReturn()
    End Function

    Public Function GenerateCompositeViewPDFWithOptions(ByRef CompositeViewID As Guid, ByVal OutputPath As String, PageSize As String, PageOrientationPortrait As Boolean, TopMargin As Integer, LeftMargin As Integer, BottomMargin As Integer, RightMargin As Integer, PageNumberFormat As String, ShowPageNumber As Boolean, Optional ByRef QS As Collection = Nothing) As String
        Try
            Dim pdfConverter As New PdfConverter()
            pdfConverter.LicenseKey = "zuX/7v/u//7u+eD+7v3/4P/84Pf39/c="

            Dim _pdfpagesize As PdfPageSize
            Select Case LCase(PageSize)
                Case "a0"
                    _pdfpagesize = PdfPageSize.A0
                Case "a1"
                    _pdfpagesize = PdfPageSize.A1
                Case "a2"
                    _pdfpagesize = PdfPageSize.A2
                Case "a3"
                    _pdfpagesize = PdfPageSize.A3
                Case "a5"
                    _pdfpagesize = PdfPageSize.A5
                Case "a6"
                    _pdfpagesize = PdfPageSize.A6
                Case "a7"
                    _pdfpagesize = PdfPageSize.A7
                Case "a8"
                    _pdfpagesize = PdfPageSize.A8
                Case "a9"
                    _pdfpagesize = PdfPageSize.A9
                Case "a10"
                    _pdfpagesize = PdfPageSize.A10
                Case "b0"
                    _pdfpagesize = PdfPageSize.B0
                Case "b1"
                    _pdfpagesize = PdfPageSize.B1
                Case "b2"
                    _pdfpagesize = PdfPageSize.B2
                Case "b3"
                    _pdfpagesize = PdfPageSize.B3
                Case "b4"
                    _pdfpagesize = PdfPageSize.B4
                Case "b5"
                    _pdfpagesize = PdfPageSize.B5
                Case "letter"
                    _pdfpagesize = PdfPageSize.Letter
                Case "legal"
                    _pdfpagesize = PdfPageSize.Legal
                Case Else
                    _pdfpagesize = PdfPageSize.A4
            End Select


            ' call the converter and get a Document object from URL
            pdfConverter.PdfDocumentOptions.PdfPageSize = _pdfpagesize
            pdfConverter.PdfDocumentOptions.PdfPageOrientation = IIf(PageOrientationPortrait = True, PDFPageOrientation.Portrait, PDFPageOrientation.Landscape)
            pdfConverter.PdfDocumentOptions.BottomMargin = BottomMargin
            pdfConverter.PdfDocumentOptions.TopMargin = TopMargin
            pdfConverter.PdfDocumentOptions.LeftMargin = LeftMargin
            pdfConverter.PdfDocumentOptions.RightMargin = RightMargin

            pdfConverter.PdfFooterOptions.PageNumberingFormatString = PageNumberFormat
            pdfConverter.PdfFooterOptions.ShowPageNumber = ShowPageNumber
            pdfConverter.PdfFooterOptions.FooterTextFontName = "Calibri"
            pdfConverter.PdfDocumentOptions.ShowFooter = True
            pdfConverter.PdfDocumentOptions.ShowHeader = False
            pdfConverter.PdfHeaderOptions.DrawHeaderLine = False
            pdfConverter.PdfFooterOptions.DrawFooterLine = False
            pdfConverter.PdfFooterOptions.FooterTextYLocation = 42
            pdfConverter.PdfDocumentOptions.GenerateSelectablePdf = True
            pdfConverter.PdfDocumentOptions.EmbedFonts = True



            Dim baseUrl As String = "file:///" + WebconfigSettings.BasePath.TrimEnd("\") + "\"



            Dim _sw As New System.IO.StringWriter

            Dim _HTML As String = GenerateCompositeViewHTML(CompositeViewID, QS)


            Dim pdfBytes As Byte() = pdfConverter.GetPdfBytesFromHtmlString(_HTML, baseUrl)
            System.IO.File.WriteAllBytes(OutputPath, pdfBytes)
            Return ""
        Catch ex As Exception
            Return ex.ToString
        End Try

    End Function


    'The Generate PDF function only generates PDFs without the standard substitution tags ([$xxxx])
    Public Function GenerateCompositeViewPDF(ByRef CompositeViewID As Guid, ByVal OutputPath As String, Optional ByRef QS As Collection = Nothing) As String
        Try
            Dim pdfConverter As New PdfConverter()
            pdfConverter.LicenseKey = "zuX/7v/u//7u+eD+7v3/4P/84Pf39/c="

            ' call the converter and get a Document object from URL
            pdfConverter.PdfDocumentOptions.PdfPageSize = PdfPageSize.A4
            pdfConverter.PdfDocumentOptions.PdfPageOrientation = PDFPageOrientation.Portrait
            pdfConverter.PdfDocumentOptions.BottomMargin = 29
            pdfConverter.PdfDocumentOptions.TopMargin = 29
            pdfConverter.PdfDocumentOptions.LeftMargin = 29
            pdfConverter.PdfDocumentOptions.RightMargin = 29

            pdfConverter.PdfFooterOptions.PageNumberingFormatString = "Page &p; of &P;"
            pdfConverter.PdfFooterOptions.ShowPageNumber = True
            pdfConverter.PdfFooterOptions.FooterTextFontName = "Calibri"
            pdfConverter.PdfDocumentOptions.ShowFooter = True
            pdfConverter.PdfDocumentOptions.ShowHeader = False
            pdfConverter.PdfHeaderOptions.DrawHeaderLine = False
            pdfConverter.PdfFooterOptions.DrawFooterLine = False
            pdfConverter.PdfFooterOptions.FooterTextYLocation = 42
            pdfConverter.PdfDocumentOptions.GenerateSelectablePdf = True
            pdfConverter.PdfDocumentOptions.EmbedFonts = True



            Dim baseUrl As String = "file:///" + WebconfigSettings.BasePath.TrimEnd("\") + "\"



            Dim _sw As New System.IO.StringWriter

            Dim _HTML As String = GenerateCompositeViewHTML(CompositeViewID, QS)


            Dim pdfBytes As Byte() = pdfConverter.GetPdfBytesFromHtmlString(_HTML, baseUrl)
            System.IO.File.WriteAllBytes(OutputPath, pdfBytes)
            Return ""
        Catch ex As Exception
            Return ex.ToString
        End Try

    End Function

    Public Function GenerateCompositeViewPDF2(ByRef CompositeViewID As Guid, ByVal OutputPath As String, Optional ByRef QS As Collection = Nothing) As String
        Try
            Dim pdfConverter As New PdfConverter()
            pdfConverter.LicenseKey = "zuX/7v/u//7u+eD+7v3/4P/84Pf39/c="

            ' call the converter and get a Document object from URL
            pdfConverter.PdfDocumentOptions.PdfPageSize = PdfPageSize.A4
            pdfConverter.PdfDocumentOptions.PdfPageOrientation = PDFPageOrientation.Portrait
            pdfConverter.PdfDocumentOptions.BottomMargin = 29
            pdfConverter.PdfDocumentOptions.TopMargin = 29
            pdfConverter.PdfDocumentOptions.LeftMargin = 29
            pdfConverter.PdfDocumentOptions.RightMargin = 29

            pdfConverter.PdfFooterOptions.PageNumberingFormatString = "Page &p; of &P;"
            pdfConverter.PdfFooterOptions.ShowPageNumber = True
            pdfConverter.PdfFooterOptions.FooterTextFontName = "Calibri"
            pdfConverter.PdfDocumentOptions.ShowFooter = True
            pdfConverter.PdfDocumentOptions.ShowHeader = False
            pdfConverter.PdfHeaderOptions.DrawHeaderLine = False
            pdfConverter.PdfFooterOptions.DrawFooterLine = False
            pdfConverter.PdfFooterOptions.FooterTextYLocation = 42
            pdfConverter.PdfDocumentOptions.GenerateSelectablePdf = True
            pdfConverter.PdfDocumentOptions.EmbedFonts = True



            Dim baseUrl As String = "file:///" + WebconfigSettings.BasePath.TrimEnd("\") + "\"



            Dim _sw As New System.IO.StringWriter

            Dim _HTML As String = GenerateCompositeViewHTML2(CompositeViewID, QS)


            Dim pdfBytes As Byte() = pdfConverter.GetPdfBytesFromHtmlString(_HTML, baseUrl)
            System.IO.File.WriteAllBytes(OutputPath, pdfBytes)
            Return ""
        Catch ex As Exception
            Return ex.ToString
        End Try

    End Function

    Public Function GenerateCompositeViewPDF2WithOptions(ByRef CompositeViewID As Guid, ByVal OutputPath As String, PageSize As String, PageOrientationPortrait As Boolean, TopMargin As Integer, LeftMargin As Integer, BottomMargin As Integer, RightMargin As Integer, PageNumberFormat As String, ShowPageNumber As Boolean, Optional ByRef QS As Collection = Nothing) As String
        Try
            Dim pdfConverter As New PdfConverter()
            pdfConverter.LicenseKey = "zuX/7v/u//7u+eD+7v3/4P/84Pf39/c="

            Dim _pdfpagesize As PdfPageSize
            Select Case LCase(PageSize)
                Case "a0"
                    _pdfpagesize = PdfPageSize.A0
                Case "a1"
                    _pdfpagesize = PdfPageSize.A1
                Case "a2"
                    _pdfpagesize = PdfPageSize.A2
                Case "a3"
                    _pdfpagesize = PdfPageSize.A3
                Case "a5"
                    _pdfpagesize = PdfPageSize.A5
                Case "a6"
                    _pdfpagesize = PdfPageSize.A6
                Case "a7"
                    _pdfpagesize = PdfPageSize.A7
                Case "a8"
                    _pdfpagesize = PdfPageSize.A8
                Case "a9"
                    _pdfpagesize = PdfPageSize.A9
                Case "a10"
                    _pdfpagesize = PdfPageSize.A10
                Case "b0"
                    _pdfpagesize = PdfPageSize.B0
                Case "b1"
                    _pdfpagesize = PdfPageSize.B1
                Case "b2"
                    _pdfpagesize = PdfPageSize.B2
                Case "b3"
                    _pdfpagesize = PdfPageSize.B3
                Case "b4"
                    _pdfpagesize = PdfPageSize.B4
                Case "b5"
                    _pdfpagesize = PdfPageSize.B5
                Case "letter"
                    _pdfpagesize = PdfPageSize.Letter
                Case "legal"
                    _pdfpagesize = PdfPageSize.Legal
                Case Else
                    _pdfpagesize = PdfPageSize.A4
            End Select


            ' call the converter and get a Document object from URL
            pdfConverter.PdfDocumentOptions.PdfPageSize = _pdfpagesize
            pdfConverter.PdfDocumentOptions.PdfPageOrientation = IIf(PageOrientationPortrait = True, PDFPageOrientation.Portrait, PDFPageOrientation.Landscape)
            pdfConverter.PdfDocumentOptions.BottomMargin = BottomMargin
            pdfConverter.PdfDocumentOptions.TopMargin = TopMargin
            pdfConverter.PdfDocumentOptions.LeftMargin = LeftMargin
            pdfConverter.PdfDocumentOptions.RightMargin = RightMargin

            pdfConverter.PdfFooterOptions.PageNumberingFormatString = PageNumberFormat
            pdfConverter.PdfFooterOptions.ShowPageNumber = ShowPageNumber
            pdfConverter.PdfFooterOptions.FooterTextFontName = "Calibri"
            pdfConverter.PdfDocumentOptions.ShowFooter = True
            pdfConverter.PdfDocumentOptions.ShowHeader = False
            pdfConverter.PdfHeaderOptions.DrawHeaderLine = False
            pdfConverter.PdfFooterOptions.DrawFooterLine = False
            pdfConverter.PdfFooterOptions.FooterTextYLocation = 42
            pdfConverter.PdfDocumentOptions.GenerateSelectablePdf = True
            pdfConverter.PdfDocumentOptions.EmbedFonts = True




            Dim baseUrl As String = "file:///" + WebconfigSettings.BasePath.TrimEnd("\") + "\"



            Dim _sw As New System.IO.StringWriter

            Dim _HTML As String = GenerateCompositeViewHTML2(CompositeViewID, QS)


            Dim pdfBytes As Byte() = pdfConverter.GetPdfBytesFromHtmlString(_HTML, baseUrl)
            System.IO.File.WriteAllBytes(OutputPath, pdfBytes)
            Return ""
        Catch ex As Exception
            Return ex.ToString
        End Try

    End Function


    Public Function GeneratePDFFromHTML(ByRef HTML As String, ByVal HeaderHTML As String, ByVal FooterHTML As String, ByVal OutputPath As String, Optional ByVal Margin As Long = 29, Optional ByVal ShowPageNumber As Boolean = True, Optional ByVal PageNumberingFormat As String = "Page &p; of &P;") As String
        Try
            Dim pdfConverter As New PdfConverter()
            pdfConverter.LicenseKey = "zuX/7v/u//7u+eD+7v3/4P/84Pf39/c="

            ' call the converter and get a Document object from URL
            pdfConverter.PdfDocumentOptions.PdfPageSize = PdfPageSize.A4
            pdfConverter.PdfDocumentOptions.PdfPageOrientation = PDFPageOrientation.Portrait
            pdfConverter.PdfDocumentOptions.BottomMargin = Margin
            pdfConverter.PdfDocumentOptions.TopMargin = Margin
            pdfConverter.PdfDocumentOptions.LeftMargin = Margin
            pdfConverter.PdfDocumentOptions.RightMargin = Margin

            pdfConverter.PdfFooterOptions.PageNumberingFormatString = PageNumberingFormat
            pdfConverter.PdfFooterOptions.ShowPageNumber = ShowPageNumber
            pdfConverter.PdfFooterOptions.FooterTextFontName = "Calibri"
            pdfConverter.PdfDocumentOptions.ShowFooter = IIf(Len(FooterHTML) > 0, True, False)
            pdfConverter.PdfDocumentOptions.ShowHeader = IIf(Len(HeaderHTML) > 0, True, False)

            Dim baseUrl As String = "file:///" + WebconfigSettings.BasePath.TrimEnd("\") + "\"

            If Len(HeaderHTML) > 0 Then
                Dim _hdr As New HtmlToPdfArea(HeaderHTML, baseUrl)
                pdfConverter.PdfHeaderOptions.HtmlToPdfArea = _hdr
            End If

            If Len(FooterHTML) > 0 Then
                Dim _ftr As New HtmlToPdfArea(FooterHTML, baseUrl)
                pdfConverter.PdfFooterOptions.HtmlToPdfArea = _ftr
            End If


            pdfConverter.PdfHeaderOptions.DrawHeaderLine = False
            pdfConverter.PdfFooterOptions.DrawFooterLine = False
            pdfConverter.PdfFooterOptions.FooterTextYLocation = 42
            pdfConverter.PdfDocumentOptions.GenerateSelectablePdf = True
            pdfConverter.PdfDocumentOptions.EmbedFonts = True



            Dim _sw As New System.IO.StringWriter

            Dim pdfBytes As Byte() = pdfConverter.GetPdfBytesFromHtmlString(HTML, baseUrl)
            System.IO.File.WriteAllBytes(OutputPath, pdfBytes)
            Return ""
        Catch ex As Exception
            Return ex.ToString
        End Try

    End Function

    Public Function GeneratePDFFromHTMLWithOptions(ByRef HTML As String, ByVal HeaderHTML As String, ByVal FooterHTML As String, ByVal OutputPath As String, PageSize As String, PageOrientationPortrait As Boolean, TopMargin As Integer, LeftMargin As Integer, BottomMargin As Integer, RightMargin As Integer, PageNumberFormat As String, ShowPageNumber As Boolean, Optional ByRef QS As Collection = Nothing) As String
        Try
            Dim pdfConverter As New PdfConverter()
            pdfConverter.LicenseKey = "zuX/7v/u//7u+eD+7v3/4P/84Pf39/c="

            Dim _pdfpagesize As PdfPageSize
            Select Case LCase(PageSize)
                Case "a0"
                    _pdfpagesize = PdfPageSize.A0
                Case "a1"
                    _pdfpagesize = PdfPageSize.A1
                Case "a2"
                    _pdfpagesize = PdfPageSize.A2
                Case "a3"
                    _pdfpagesize = PdfPageSize.A3
                Case "a5"
                    _pdfpagesize = PdfPageSize.A5
                Case "a6"
                    _pdfpagesize = PdfPageSize.A6
                Case "a7"
                    _pdfpagesize = PdfPageSize.A7
                Case "a8"
                    _pdfpagesize = PdfPageSize.A8
                Case "a9"
                    _pdfpagesize = PdfPageSize.A9
                Case "a10"
                    _pdfpagesize = PdfPageSize.A10
                Case "b0"
                    _pdfpagesize = PdfPageSize.B0
                Case "b1"
                    _pdfpagesize = PdfPageSize.B1
                Case "b2"
                    _pdfpagesize = PdfPageSize.B2
                Case "b3"
                    _pdfpagesize = PdfPageSize.B3
                Case "b4"
                    _pdfpagesize = PdfPageSize.B4
                Case "b5"
                    _pdfpagesize = PdfPageSize.B5
                Case "letter"
                    _pdfpagesize = PdfPageSize.Letter
                Case "legal"
                    _pdfpagesize = PdfPageSize.Legal
                Case Else
                    _pdfpagesize = PdfPageSize.A4
            End Select


            ' call the converter and get a Document object from URL
            pdfConverter.PdfDocumentOptions.PdfPageSize = _pdfpagesize
            pdfConverter.PdfDocumentOptions.PdfPageOrientation = IIf(PageOrientationPortrait = True, PDFPageOrientation.Portrait, PDFPageOrientation.Landscape)
            pdfConverter.PdfDocumentOptions.BottomMargin = BottomMargin
            pdfConverter.PdfDocumentOptions.TopMargin = TopMargin
            pdfConverter.PdfDocumentOptions.LeftMargin = LeftMargin
            pdfConverter.PdfDocumentOptions.RightMargin = RightMargin

            pdfConverter.PdfFooterOptions.PageNumberingFormatString = PageNumberFormat
            pdfConverter.PdfFooterOptions.ShowPageNumber = ShowPageNumber
            pdfConverter.PdfFooterOptions.FooterTextFontName = "Calibri"
            pdfConverter.PdfDocumentOptions.ShowFooter = True
            pdfConverter.PdfDocumentOptions.ShowHeader = False
            pdfConverter.PdfHeaderOptions.DrawHeaderLine = False
            pdfConverter.PdfFooterOptions.DrawFooterLine = False
            pdfConverter.PdfFooterOptions.FooterTextYLocation = 42
            pdfConverter.PdfDocumentOptions.GenerateSelectablePdf = True
            pdfConverter.PdfDocumentOptions.EmbedFonts = True


            Dim baseUrl As String = "file:///" + WebconfigSettings.BasePath.TrimEnd("\") + "\"

            If Len(HeaderHTML) > 0 Then
                Dim _hdr As New HtmlToPdfArea(HeaderHTML, baseUrl)
                pdfConverter.PdfHeaderOptions.HtmlToPdfArea = _hdr
            End If

            If Len(FooterHTML) > 0 Then
                Dim _ftr As New HtmlToPdfArea(FooterHTML, baseUrl)
                pdfConverter.PdfFooterOptions.HtmlToPdfArea = _ftr
            End If




            Dim _sw As New System.IO.StringWriter

            Dim pdfBytes As Byte() = pdfConverter.GetPdfBytesFromHtmlString(HTML, baseUrl)
            System.IO.File.WriteAllBytes(OutputPath, pdfBytes)
            Return ""
        Catch ex As Exception
            Return ex.ToString
        End Try

    End Function


    Public Function GenerateCompositeViewHTML(ByRef CompositeViewID As Guid, ByRef QS As Collection) As String
        Dim _settings As ZukamiLib.ZukamiSettings = GetZukamiSettings()
        Dim _webObj As New ZukamiLib.WebSession(_settings)
        _webObj.OpenConnection()
        Dim _MainDataset As DataSet = _webObj.CompositePage_GetRecord(CompositeViewID)
        _webObj.CloseConnection()
        _webObj = Nothing

        If _MainDataset.Tables(0).Rows.Count = 0 Then
            Return ""
        Else
            Dim _Header As String = GlobalFunctions.FormatData(_MainDataset.Tables(0).Rows(0).Item("PageHeader"))
            Dim _Footer As String = GlobalFunctions.FormatData(_MainDataset.Tables(0).Rows(0).Item("PageFooter"))
            Dim _Content As String = GlobalFunctions.FormatData(_MainDataset.Tables(0).Rows(0).Item("PageContent"))

            Dim _AppID As Guid = GlobalFunctions.GetGUID(_MainDataset.Tables(0).Rows(0).Item("AppID"))
            Dim _webobj2 As New ZukamiLib.WebSession(_settings)
            _webobj2.OpenConnection()
            GenerateCompositeViewHTML = ApplyTheme(_webobj2, _AppID, _Header, _Content, _Footer, QS)
            _webobj2.CloseConnection()
            _webobj2 = Nothing
        End If
    End Function

    Public Function GenerateCompositeViewHTML2(ByRef CompositeViewID As Guid, ByRef QS As Collection) As String
        Dim _settings As ZukamiLib.ZukamiSettings = GetZukamiSettings()
        Dim _webObj As New ZukamiLib.WebSession(_settings)
        _webObj.OpenConnection()
        Dim _MainDataset As DataSet = _webObj.CompositePage_GetRecord(CompositeViewID)
        _webObj.CloseConnection()
        _webObj = Nothing

        If _MainDataset.Tables(0).Rows.Count = 0 Then
            Return ""
        Else
            Dim _Header As String = GlobalFunctions.FormatData(_MainDataset.Tables(0).Rows(0).Item("PageHeader"))
            Dim _Footer As String = GlobalFunctions.FormatData(_MainDataset.Tables(0).Rows(0).Item("PageFooter"))
            Dim _Content As String = GlobalFunctions.FormatData(_MainDataset.Tables(0).Rows(0).Item("PageContent"))

            Dim _AppID As Guid = GlobalFunctions.GetGUID(_MainDataset.Tables(0).Rows(0).Item("AppID"))
            Dim _webobj2 As New ZukamiLib.WebSession(_settings)
            _webobj2.OpenConnection()
            GenerateCompositeViewHTML2 = ApplyTheme2(_webobj2, _AppID, _Header, _Content, _Footer, QS)
            _webobj2.CloseConnection()
            _webobj2 = Nothing
        End If
    End Function

    Private Shared Function ApplyTheme2(ByRef WebObj As ZukamiLib.WebSession, ByRef AppID As Guid, ByVal Header As String, ByVal Content As String, ByVal Footer As String, ByRef QS As Collection) As String
        Dim plcMain As New System.Web.UI.WebControls.PlaceHolder
        Dim plcMaster As New System.Web.UI.WebControls.PlaceHolder


        '   GlobalFunctions.SubstituteTags("CV", WebObj, "CV", WebObj, plcMain, Content, AppID)



        ''Do for footer
        If Len(Footer) > 0 Then
            plcMain.Controls.Add(New LiteralControl(Footer))
        End If


        Dim sb As New System.Text.StringBuilder
        Dim _tw As New System.IO.StringWriter(sb)
        Dim _hw As New System.Web.UI.HtmlTextWriter(_tw)
        plcMain.RenderControl(_hw)
        Dim _body As String = sb.ToString

        Dim sb2 As New System.Text.StringBuilder
        Dim _tw2 As New System.IO.StringWriter(sb2)
        Dim _hw2 As New System.Web.UI.HtmlTextWriter(_tw2)
        plcMaster.RenderControl(_hw2)

        Dim custhdr As String = "<html><head id=""ctl00_Head1""></head>"
        Dim custhdrend As String = "</html>"
        'Return RenderIt(WebObj, custhdr + sb2.ToString & _body + custhdrend)

        Dim _index As Integer = 1
        Dim _start As Integer = 1
        Dim _end As Integer = 0
        If QS Is Nothing = False Then
            Do While _start > 0
                _index = 1
                _start = InStr(_index, Content, "[$querystring:", CompareMethod.Text)
                If _start > 0 Then
                    _index = _start
                    _end = InStr(_index, Content, "]", CompareMethod.Text)
                    Dim _Tagname As String = Mid(Content, _start + Len("[$querystring:"), _end - (_start + Len("[$querystring:")))
                    Dim _Tagvalue As String = ""
                    If QS.Contains(_Tagname) = True Then
                        _Tagvalue = QS.Item(_Tagname)
                    End If
                    Content = Replace(Content, "[$querystring:" & _Tagname & "]", _Tagvalue,,, CompareMethod.Text)
                End If
            Loop
        End If




        Return RenderIt(WebObj, custhdr + sb2.ToString & Content + custhdrend)



    End Function

    Private Shared Function ApplyTheme(ByRef WebObj As ZukamiLib.WebSession, ByRef AppID As Guid, ByVal Header As String, ByVal Content As String, ByVal Footer As String, ByRef QS As Collection) As String
        Dim plcMain As New System.Web.UI.WebControls.PlaceHolder
        Dim plcMaster As New System.Web.UI.WebControls.PlaceHolder


        '   GlobalFunctions.SubstituteTags("CV", WebObj, "CV", WebObj, plcMain, Content, AppID)



        ''Do for footer
        If Len(Footer) > 0 Then
            plcMain.Controls.Add(New LiteralControl(Footer))
        End If


        Dim sb As New System.Text.StringBuilder
        Dim _tw As New System.IO.StringWriter(sb)
        Dim _hw As New System.Web.UI.HtmlTextWriter(_tw)
        plcMain.RenderControl(_hw)
        Dim _body As String = sb.ToString

        Dim sb2 As New System.Text.StringBuilder
        Dim _tw2 As New System.IO.StringWriter(sb2)
        Dim _hw2 As New System.Web.UI.HtmlTextWriter(_tw2)
        plcMaster.RenderControl(_hw2)

        Dim custhdr As String = "<html><head id=""ctl00_Head1""></head>"
        Dim custhdrend As String = "</html>"
        Return RenderIt(WebObj, custhdr + sb2.ToString & _body + custhdrend)

        '  Return RenderIt(WebObj, custhdr + sb2.ToString & Content + custhdrend)



    End Function

    Private Shared Function RenderIt(ByRef WebObj As ZukamiLib.WebSession, ByVal Data As String) As String
        Return GlobalFunctions.ResolveRowTemplate(WebObj, Data)
    End Function



    Public Function ConvertDocxToTIFF(ByVal SourcePath As String, ByVal TargetPath As String) As String
        Return GlobalFunctions.ConvertDocxToTIFF(SourcePath, TargetPath)
    End Function

    Public Sub DeleteFormField(ByVal FormGUID As Guid, ByVal fieldname As String)
        Dim webobj As ZukamiLib.WebSession = Nothing
        webobj = New ZukamiLib.WebSession(GetZukamiSettings)
        webobj.OpenConnection()

        Dim _listset As DataSet = webobj.ListItems_GetByName(fieldname, FormGUID)
        If _listset.Tables(0).Rows.Count > 0 Then
            Dim _listItemID As Guid = GlobalFunctions.GetGUID(_listset.Tables(0).Rows(0).Item("ListItemID"))
            ' webobj.ListItems_Delete(_listItemID, FormGUID)
            GlobalFunctions.ListItems_Delete(_listItemID, FormGUID, webobj)

            'now we delete from the main view as well
            webobj.ViewColumns_DeleteBySourceColID(_listItemID)

            'we also delete any attached javascript code
            GlobalFunctions.DeleteDynamicCalcFunction(webobj, FormGUID, _listItemID)
        End If

        webobj.CloseConnection()
        webobj = Nothing


    End Sub

    Public Function RunDynamicSQL(ByVal SQL As String) As String
        Dim _sql As String = SQL
        Try
            RunDynamicSQL = ""
            Dim webObj As ZukamiLib.WebSession = Nothing
            If KeepConnAlive = True Then
                webObj = GetSession()
            Else
                webObj = New ZukamiLib.WebSession(GetZukamiSettings)
                webObj.OpenConnection()
            End If



            Dim _counter As Integer
            If _zfields Is Nothing = False Then
                For _counter = 1 To _zfields.Count
                    Dim _temp As ZField = _zfields.Item(_counter)
                    Dim _fieldname As String = "[$" & _temp.FieldName & "]"

                    Dim obj As Object = Data(_temp.FieldName)
                    If Not TypeOf obj Is DataTable Then
                        Try
                            _sql = Replace(_sql, _fieldname, GlobalFunctions.FormatSQLData(obj), , , CompareMethod.Text)
                        Catch ex As Exception

                        End Try

                    End If
                Next _counter
            End If
            _sql = Replace(_sql, "[$ID]", PrimaryID, , , CompareMethod.Text)


            webObj.CustomSQLCommand(_sql)
            webObj.CustomClearParameters()
            webObj.CustomSQLExecute()
            If Len(webObj.LastError) > 0 Then
                RunDynamicSQL = webObj.LastError
                logger.Error(webObj.LastError + ", sql: " + _sql)
            End If
errEnd:
            If KeepConnAlive = False Then
                webObj.CloseConnection()
                webObj = Nothing
            End If

        Catch ex As Exception
            RunDynamicSQL = ex.ToString
            logger.Error(ex, "_sql: " + _sql)
        End Try

    End Function


    Public Function RunDynamicSQLReturn(ByVal SQL As String) As DataSet
        Dim _sql As String = SQL
        Try
            RunDynamicSQLReturn = Nothing
            Dim webObj As ZukamiLib.WebSession = Nothing
            If KeepConnAlive = True Then
                webObj = GetSession()
            Else
                webObj = New ZukamiLib.WebSession(GetZukamiSettings)
                webObj.OpenConnection()
            End If


            If _zfields Is Nothing = False Then


                Dim _counter As Integer
                For _counter = 1 To _zfields.Count
                    Dim _temp As ZField = _zfields.Item(_counter)
                    Dim _fieldname As String = "[$" & _temp.FieldName & "]"

                    Dim obj As Object = Data(_temp.FieldName)
                    If Not TypeOf obj Is DataTable Then
                        Try
                            _sql = Replace(_sql, _fieldname, GlobalFunctions.FormatSQLData(obj), , , CompareMethod.Text)
                        Catch ex As Exception

                        End Try

                    End If
                Next _counter
            End If
            _sql = Replace(_sql, "[$ID]", PrimaryID, , , CompareMethod.Text)


            webObj.CustomSQLCommand(_sql)
            webObj.CustomClearParameters()
            RunDynamicSQLReturn = webObj.CustomSQLExecuteReturn
            If Len(webObj.LastError) > 0 Then
                RunDynamicSQLReturn = Nothing
                logger.Error(webObj.LastError + ", sql: " + _sql)
            End If
errEnd:
            If KeepConnAlive = False Then
                webObj.CloseConnection()
                webObj = Nothing
            End If

        Catch ex As Exception
            RunDynamicSQLReturn = Nothing
            logger.Error(ex, "_sql: " + _sql)
        End Try

    End Function


    Public Function RunDynamicSQL(ByVal AppID As Guid, ByVal Datasource As String, ByVal SQL As String) As String
        Dim _sql As String = SQL
        Try
            Dim webobj As ZukamiLib.WebSession = Nothing
            RunDynamicSQL = ""

            Dim _tempweb As New ZukamiLib.WebSession(GetZukamiSettings)
            _tempweb.OpenConnection()
            Dim _zuksettings As ZukamiLib.ZukamiSettings = GlobalFunctions.GetDatasourceConnectionString(_tempweb, AppID, Datasource)
            _tempweb.CloseConnection()
            _tempweb = Nothing

            webobj = New ZukamiLib.WebSession(_zuksettings)
            webobj.OpenOLEDBConnection(_zuksettings.PrimaryConnectionString)
            If Len(webobj.LastError) > 0 Then
                Return webobj.LastError
            End If

            Dim _counter As Integer
            If _zfields Is Nothing = False Then
                For _counter = 1 To _zfields.Count
                    Dim _temp As ZField = _zfields.Item(_counter)
                    Dim _fieldname As String = "[$" & _temp.FieldName & "]"

                    Dim obj As Object = Data(_temp.FieldName)
                    If Not TypeOf obj Is DataTable Then
                        Try
                            _sql = Replace(_sql, _fieldname, GlobalFunctions.FormatSQLData(obj), , , CompareMethod.Text)
                        Catch ex As Exception

                        End Try

                    End If
                Next _counter
            End If
            _sql = Replace(_sql, "[$ID]", PrimaryID, , , CompareMethod.Text)

            webobj.CustomOLEDBSQLCommand(_sql)
            webobj.CustomOLEDBClearParameters()
            webobj.CustomOLEDBSQLExecute()
            If Len(webobj.LastError) > 0 Then
                RunDynamicSQL = webobj.LastError
                logger.Error(webobj.LastError + ", sql: " + _sql)
            End If
errEnd:
            webobj.CloseOLEDBConnection()
        Catch ex As Exception
            RunDynamicSQL = ex.ToString
            logger.Error(ex, "_sql: " + _sql)
        End Try
    End Function

    Public Function RunDynamicSQLReturn(ByVal AppID As Guid, ByVal Datasource As String, ByVal SQL As String) As DataSet
        Dim _sql As String = SQL
        Try
            logger.Debug("AppID: " + AppID.ToString() + ", Datasource: " + Datasource + ", SQL: " + SQL)
            Dim webobj As ZukamiLib.WebSession = Nothing
            RunDynamicSQLReturn = Nothing

            Dim _tempweb As New ZukamiLib.WebSession(GetZukamiSettings)
            _tempweb.OpenConnection()
            Dim _zuksettings As ZukamiLib.ZukamiSettings = GlobalFunctions.GetDatasourceConnectionString(_tempweb, AppID, Datasource)
            _tempweb.CloseConnection()
            _tempweb = Nothing

            webobj = New ZukamiLib.WebSession(_zuksettings)
            webobj.OpenOLEDBConnection(_zuksettings.PrimaryConnectionString)
            If Len(webobj.LastError) > 0 Then
                logger.Error(webobj.LastError + ", sql: " + _sql)
                Return Nothing
            End If

            Dim _counter As Integer
            If _zfields Is Nothing = False Then
                For _counter = 1 To _zfields.Count
                    Dim _temp As ZField = _zfields.Item(_counter)
                    Dim _fieldname As String = "[$" & _temp.FieldName & "]"

                    Dim obj As Object = Data(_temp.FieldName)
                    If Not TypeOf obj Is DataTable Then
                        Try
                            _sql = Replace(_sql, _fieldname, GlobalFunctions.FormatSQLData(obj), , , CompareMethod.Text)
                        Catch ex As Exception
                            logger.Error(ex, "sql: " + _sql)
                        End Try

                    End If
                Next _counter
            End If

            If (_sql.ToUpper().Contains("[$ID]")) Then
                _sql = Replace(_sql, "[$ID]", PrimaryID, , , CompareMethod.Text)
            End If

            RunDynamicSQLReturn = webobj.RunOLEDBSQL(_sql)
            If Len(webobj.LastError) > 0 Then
                RunDynamicSQLReturn = Nothing
                logger.Error(webobj.LastError + ", sql: " + _sql)
            End If
errEnd:
            webobj.CloseOLEDBConnection()
        Catch ex As Exception
            RunDynamicSQLReturn = Nothing
            logger.Error(ex, ", sql: " + _sql)
        End Try
    End Function

    Public Function IsChecked(ByRef Data As Object, ByVal CheckedItems As String) As Boolean
        If Not TypeOf Data Is Collection Then Return False
        Dim _itemstocompare As New Collection
        Dim arrItems() As String = Split(CheckedItems, ";")
        Dim _i As Integer
        For _i = 0 To UBound(arrItems)
            Dim _strItem As String = Trim(arrItems(_i))
            If Len(_strItem) > 0 Then
                _itemstocompare.Add(_strItem, _strItem)
            End If
        Next _i

        Dim _fullfilled As Integer = 0
        Dim _coll As Collection = Data
        Dim _counter As Integer
        For _counter = 1 To _coll.Count
            Dim _rb As CheckBox = _coll.Item(_counter)
            If _itemstocompare.Contains(_rb.Text) = True Then
                If _rb.Checked = True Then
                    _fullfilled += 1
                End If
            End If
        Next
        If _fullfilled = _itemstocompare.Count Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function GetRBValue(ByRef Data As Object) As String
        If Not TypeOf Data Is Collection Then Return ""
        Dim _coll As Collection = Data
        Dim _counter As Integer
        For _counter = 1 To _coll.Count
            Dim _rb As RadioButton = _coll.Item(_counter)
            If _rb.Checked = True Then
                Return _rb.Text
            End If
        Next
        Return ""
    End Function

    Public Function GetInt(ByVal Data As Object) As Integer
        Return GlobalFunctions.FormatInteger(Data)
    End Function

    Public Function GetBool(ByVal Data As Object) As Boolean
        Return GlobalFunctions.FormatBoolean(Data)
    End Function

    Public Function GetDbl(ByVal Data As Object) As Double
        Return GlobalFunctions.FormatDouble(Data)
    End Function

    Public Function GetDate(ByVal Data As Object) As Date
        Return GlobalFunctions.GetDateTime(Data)
    End Function

    Public Function GetMoney(ByVal Data As Object) As String
        Return GlobalFunctions.FormatMoney(Data)
    End Function


    Public Function ImportFromCSV(ByVal Filepath As String, ByVal TableName As String, ByRef Assignments As Collection) As String
        Dim arrStrings() As String
        Try
            arrStrings = System.IO.File.ReadAllLines(Filepath)
        Catch ex As Exception
            Return ex.ToString
        End Try

        Dim _errorLines As String = ""
        Dim _counter As Integer
        Dim _sql As String = ""
        Dim _valuelist As String = ""
        Dim _fieldlist As String = ""

        Dim _webObj As ZukamiLib.WebSession = Nothing
        If KeepConnAlive = True Then
            _webObj = GetSession()
        Else
            _webObj = New ZukamiLib.WebSession(GetZukamiSettings)
            _webObj.OpenConnection()
        End If

        Try

            For _counter = 0 To UBound(arrStrings)
                If Len(arrStrings(_counter)) > 0 Then
                    _valuelist = ""
                    _fieldlist = ""
                    Dim arrFieldPairs() As String = Split(arrStrings(_counter), ",")
                    _sql = "INSERT INTO " & TableName
                    Dim _counter2 As Integer
                    For _counter2 = 1 To Assignments.Count
                        Dim strAssgt As String = Assignments.Item(_counter2)
                        Dim _firstEqual As Long = InStr(1, strAssgt, "=", vbTextCompare)
                        If _firstEqual > 0 Then
                            Dim _fieldname As String = Mid(strAssgt, 1, _firstEqual - 1)
                            Dim _data As String = Mid(strAssgt, _firstEqual + 1, Len(strAssgt) - _firstEqual)
                            If Len(_fieldlist) > 0 Then _fieldlist += ","
                            _fieldlist += _fieldname

                            'parse value list first
                            _data = SubstituteValueList(_data, arrFieldPairs)
                            If Len(_valuelist) > 0 Then _valuelist += ","
                            _valuelist += _data
                        Else
                            'skip processing
                            If Len(_errorLines) > 0 Then _errorLines += vbCrLf
                            _errorLines += arrStrings(_counter)
                            GoTo errNext
                        End If
                    Next _counter2

                    _sql += "(" & _fieldlist & ") VALUES(" & _valuelist & ")"
                    'now we try to execute this
                    _webObj.CustomSQLCommand(_sql)
                    _webObj.CustomSQLExecute()
                    If Len(_webObj.LastError) > 0 Then
                        If Len(_errorLines) > 0 Then _errorLines += vbCrLf
                        _errorLines += arrStrings(_counter) & ",Error:" & _webObj.LastError & ",SQL:" & _sql
                    End If
                End If
errNext:
            Next _counter
        Catch ex As Exception
            Return ex.ToString
        End Try

        If KeepConnAlive = False Then
            _webObj.CloseConnection()
            _webObj = Nothing
        End If


        Return _errorLines
    End Function

    Private Function SubstituteValueList(ByVal Value As String, ByRef FieldPairs() As String) As String
        If Left(Value, 2) = "[$" And Right(Value, 1) = "]" Then
            Dim index As String = Mid(Value, 3, Len(Value) - 3)
            If IsNumeric(index) = False Then Return Value
            Dim _numindex As Integer = CLng(index)
            If _numindex > UBound(FieldPairs) Then Return Value
            Return FieldPairs(_numindex)
        Else
            Return Value
        End If
    End Function

    Public Function GetText(ByVal Data As Object) As String
        Return GlobalFunctions.FormatData(Data)
    End Function

    Public Function GrabLink(ByVal LinkType As String, Optional ByVal SpecialTag As String = "") As String
        Dim _req As HttpRequest = System.Web.HttpContext.Current.Request
        Dim _inboxitemid As String = ""
        Dim _taskid As String = ""
        If Len(_req.QueryString("InboxID")) > 0 Then
            'from form
            _taskid = _req.QueryString("InboxID")
        Else
            'from actionscreen
            _taskid = _req.QueryString("ID")
        End If

        If Len(_req.QueryString("ID2")) > 0 Then
            'from actionscreen
            _inboxitemid = _req.QueryString("ID2")
        Else
            'from form
            Dim webObj As ZukamiLib.WebSession = Nothing
            If KeepConnAlive = True Then
                webObj = GetSession()
            Else
                webObj = New ZukamiLib.WebSession(GetZukamiSettings)
                webObj.OpenConnection()
            End If

            webObj.CustomSQLCommand("SELECT * FROM Inboxitems WHERE TaskID='" & _taskid & "'")
            webObj.CustomClearParameters()
            Dim _set As DataSet = webObj.CustomSQLExecuteReturn
            If _set.Tables(0).Rows.Count > 0 Then
                _inboxitemid = GlobalFunctions.FormatData(_set.Tables(0).Rows(0).Item("ItemID"))
            End If
            If KeepConnAlive = False Then
                webObj.CloseConnection()
                webObj = Nothing
            End If

            _set.Dispose()

        End If



        Dim _MainDataset As DataSet = Nothing
        If GlobalFunctions.HasAdminGroup = True Then
            _MainDataset = GlobalFunctions.Task_GetByTaskID(_taskid)
        Else
            _MainDataset = GlobalFunctions.Task_GetRecord(_taskid)
        End If
        Dim InstanceID As Guid = Guid.Empty
        Dim _InstanceNodeID As Guid = Guid.Empty
        Dim _WorkflowID As Guid = Guid.Empty

        If _MainDataset.Tables(0).Rows.Count > 0 Then
            InstanceID = GlobalFunctions.GetGUID(_MainDataset.Tables(0).Rows(0).Item("InstanceID"))
            _InstanceNodeID = GlobalFunctions.GetGUID(_MainDataset.Tables(0).Rows(0).Item("InstanceNodeID"))
            _WorkflowID = GlobalFunctions.GetGUID(_MainDataset.Tables(0).Rows(0).Item("WorkflowID"))
        End If

        Select Case LCase(LinkType)
            Case "customaction"
                Return "action.aspx?a=" & _req.QueryString("a") & "&Action=Approve&ID=" & _taskid
            Case "seekinput"
                Return "sendforreview.aspx?SFI=1&ID=" & _taskid & "&InstanceID=" & InstanceID.ToString & "&NodeID=" & _InstanceNodeID.ToString & "&ID2=" & _inboxitemid.ToString & "&a=" & _req.QueryString("a")
            Case "reassign"
                Return "reassign.aspx?a=" & _req.QueryString("a") & "&ID=" & _taskid
            Case "rework"
                Return "sendforreview.aspx?ID=" & _taskid & "&InstanceID=" & InstanceID.ToString & "&NodeID=" & _InstanceNodeID.ToString & "&ID2=" & _inboxitemid.ToString & "&a=" & _req.QueryString("a")
            Case "jump"
                Return "jump.aspx?WFlowID=" & _WorkflowID.ToString & "&ID=" & _taskid & "&a=" & _req.QueryString("a") & "&ID2=" & _inboxitemid.ToString
            Case "jumpstraight"
                Return "jump.aspx?SpecialTag=" & SpecialTag & "&Window=No&WFlowID=" & _WorkflowID.ToString & "&ID=" & _taskid & "&a=" & _req.QueryString("a") & "&ID2=" & _inboxitemid.ToString
            Case "postamessage"
                Return "postmessage.aspx?a=" & _req.QueryString("a") & "&FS=&ID=" & _taskid
            Case Else
                Return ""
        End Select
    End Function



    Public Sub AddToDropdown(ByVal FieldName As String, ByVal Caption As String, ByVal Value As String, Optional ByVal InsertLocation As Integer = -1)
        Try

            Dim _DDList As DropDownList = Nothing
            Dim _args As String = ""
            Dim _counter As Integer
            Dim _formname As String = ""
            Dim _captionfield As String = ""
            Dim _valuefield As String = ""
            Dim _GroupMatch As String = ""
            For _counter = 1 To _zfields.Count
                Dim _temp As ZField = _zfields.Item(_counter)
                If StrComp(FieldName, _temp.FieldName, CompareMethod.Text) = 0 Then
                    Try
                        _DDList = _temp.FieldControl.dropdownbox
                        _GroupMatch = _temp.FieldControl.GroupMatch
                        _temp.FieldControl.GroupMatch = ""
                        _args = _temp.Arguments
                        Exit For
                    Catch ex As Exception
                        _DDList = Nothing
                    End Try
                End If
            Next _counter
            If _DDList Is Nothing Then GoTo errEnd

            Dim _selectedvalue As String = ""
            Dim _listitem As ListItem = _DDList.SelectedItem
            If _listitem Is Nothing = False Then
                _selectedvalue = _listitem.Value
            End If


            Dim _li As New System.Web.UI.WebControls.ListItem(Caption, Value)
            If InsertLocation = -1 Then
                _DDList.Items.Add(_li)
            Else
                _DDList.Items.Insert(InsertLocation, _li)
            End If



        Catch ex As Exception
        End Try
errEnd:

    End Sub

    Public Sub RemoveFromDropdown(ByVal FieldName As String, ByVal Value As String)
        Try

            Dim _DDList As DropDownList = Nothing
            Dim _args As String = ""
            Dim _counter As Integer
            Dim _formname As String = ""
            Dim _captionfield As String = ""
            Dim _valuefield As String = ""
            Dim _GroupMatch As String = ""
            For _counter = 1 To _zfields.Count
                Dim _temp As ZField = _zfields.Item(_counter)
                If StrComp(FieldName, _temp.FieldName, CompareMethod.Text) = 0 Then
                    Try
                        _DDList = _temp.FieldControl.dropdownbox
                        _GroupMatch = _temp.FieldControl.GroupMatch
                        _temp.FieldControl.GroupMatch = ""
                        _args = _temp.Arguments
                        Exit For
                    Catch ex As Exception
                        _DDList = Nothing
                    End Try
                End If
            Next _counter
            If _DDList Is Nothing Then GoTo errEnd

            Dim _selectedvalue As String = ""
            Dim _listitem As ListItem = _DDList.SelectedItem
            If _listitem Is Nothing = False Then
                _selectedvalue = _listitem.Value
            End If


            Dim _counter2 As Integer
            For _counter2 = 0 To _DDList.Items.Count - 1
                Dim _li As System.Web.UI.WebControls.ListItem = _DDList.Items(_counter2)
                If StrComp(_li.Value, Value, CompareMethod.Text) = 0 Then
                    _DDList.Items.Remove(_li)
                    Exit For
                End If
            Next _counter2

            'select the item back again
            For _counter2 = 0 To _DDList.Items.Count - 1
                Dim _li As System.Web.UI.WebControls.ListItem = _DDList.Items(_counter2)
                If StrComp(_li.Value, _selectedvalue, CompareMethod.Text) = 0 Then
                    _li.Selected = True
                    Exit For
                End If
            Next _counter2

        Catch ex As Exception
        End Try
errEnd:

    End Sub


    Public Sub FilterDropdown(ByVal FieldName As String, Optional ByVal FilterClause As String = "", Optional ByVal SortClause As String = "")
        Dim webObj As ZukamiLib.WebSession = Nothing
        Try
            logger.Trace("FieldName: " + FieldName + ", FilterClause: " + FilterClause + ", SortClause: " + SortClause)
            If KeepConnAlive = True Then
                webObj = GetSession()
            Else
                webObj = New ZukamiLib.WebSession(GetZukamiSettings)
                webObj.OpenConnection()
            End If

            Dim _DDList As DropDownList = Nothing
            Dim _args As String = ""
            Dim _counter As Integer
            Dim _formname As String = ""
            Dim _captionfield As String = ""
            Dim _valuefield As String = ""
            Dim _GroupMatch As String = ""
            For _counter = 1 To _zfields.Count
                Dim _temp As ZField = _zfields.Item(_counter)
                If StrComp(FieldName, _temp.FieldName, CompareMethod.Text) = 0 Then
                    Try
                        _DDList = _temp.FieldControl.dropdownbox
                        _GroupMatch = _temp.FieldControl.GroupMatch
                        _temp.FieldControl.GroupMatch = ""
                        _args = _temp.Arguments
                        Exit For
                    Catch ex As Exception
                        _DDList = Nothing
                        logger.Error(ex)
                    End Try
                End If
            Next _counter
            If _DDList Is Nothing Then GoTo errEnd
            logger.Trace("step 1 passed")

            Dim _selectedvalue As String = ""
            Dim _listitem As ListItem = _DDList.SelectedItem
            If _listitem Is Nothing = False Then
                _selectedvalue = _listitem.Value
            End If

            GlobalFunctions.GetDynamicArguments(_args, _formname, _captionfield, _valuefield, "")

            Dim _FormSet As DataSet = webObj.Forms_GetByFormName(_formname)
            If _FormSet.Tables(0).Rows.Count = 0 Then GoTo errEnd
            logger.Trace("step 2 passed")
            Dim _tbs As String = GlobalFunctions.FormatData(_FormSet.Tables(0).Rows(0).Item("TableBindSource"))

            Dim _sql As String = "SELECT [" & _captionfield & "],[" & _valuefield & "] FROM [" & _tbs & "]"
            Dim _wclause As String = ""
            If Len(_GroupMatch) > 0 Then
                _wclause = "[" & _captionfield & "] LIKE N'%" & GlobalFunctions.FormatSQLData(_GroupMatch) & "%'"
            End If
            If Len(FilterClause) > 0 Then
                If Len(_wclause) > 0 Then _wclause += " AND "
                _wclause += FilterClause
            End If
            If Len(_wclause) > 0 Then
                _sql += " WHERE " & _wclause
            End If
            If Len(SortClause) > 0 Then
                _sql += " ORDER BY " & SortClause

            End If
            webObj.CustomSQLCommand(_sql)
            webObj.CustomClearParameters()

            _DDList.Items.Clear()
            'Dim _resultset As DataSet = webObj.CustomSQLExecuteReturn
            Dim _resultset As New DataSet
            If Not GlobalFunctions.OracleDBEnabled() Then
                _resultset = webObj.CustomSQLExecuteReturn
            Else
                _resultset = GlobalFunctions.OracleGetDataSetBySql(GlobalFunctions.OracleConvertToPLSQL(_sql))
            End If

            Dim _selectedfound As Boolean = False
            If _resultset.Tables(0).Rows.Count > 0 Then
                'now we load the content into the drop down list

                Dim _counter3 As Integer
                For _counter3 = 0 To _resultset.Tables(0).Rows.Count - 1
                    Dim _row As DataRow = _resultset.Tables(0).Rows(_counter3)
                    Dim _value As String = GlobalFunctions.FormatData(_row.Item(_valuefield))

                    Dim _lv As New ListItem(GlobalFunctions.TruncateIt(GlobalFunctions.FormatData(_row.Item(_captionfield))), _value)


                    If StrComp(_value, _selectedvalue, CompareMethod.Text) = 0 Then
                        _selectedfound = True
                        _lv.Selected = True
                    End If

                    _DDList.Items.Add(_lv)
                Next _counter3
            End If

            Dim _lv2 As New ListItem("--Please select a value--", "")
            If _selectedfound = False Then
                Try
                    _lv2.Selected = True
                Catch ex As Exception
                End Try
            End If
            _DDList.Items.Add(_lv2)


        Catch ex As Exception
            logger.Error(ex)
        End Try


errEnd:
        If KeepConnAlive = False Then
            webObj.CloseConnection()
            webObj = Nothing
        End If
        'logger.Error("errEnd: some other thing was wrong")

    End Sub

    Public Function GetDataset(ByVal FormName As String, ByVal FieldName As String, ByVal Value As Object) As DataSet
        GetDataset = Nothing
        Dim webObj As ZukamiLib.WebSession = Nothing
        If KeepConnAlive = True Then
            webObj = GetSession()
        Else
            webObj = New ZukamiLib.WebSession(GetZukamiSettings)
            webObj.OpenConnection()
        End If

        Try
            Dim _FormSet As DataSet = webObj.Forms_GetByFormName(FormName)
            If _FormSet.Tables(0).Rows.Count = 0 Then GoTo errEnd
            Dim _tbs As String = GlobalFunctions.FormatData(_FormSet.Tables(0).Rows(0).Item("TableBindSource"))

            webObj.CustomSQLCommand("SELECT * FROM [" & _tbs & "] WHERE [" & FieldName & "]=" & CStr(Value))
            webObj.CustomClearParameters()
            GetDataset = webObj.CustomSQLExecuteReturn
        Catch ex As Exception
        End Try
errEnd:
        If KeepConnAlive = False Then
            webObj.CloseConnection()
            webObj = Nothing
        End If

    End Function

    Public Function GetDataset2(ByVal TableName As String, ByVal FieldName As String, ByVal Value As Object) As DataSet
        GetDataset2 = Nothing
        Dim webObj As ZukamiLib.WebSession = Nothing
        If KeepConnAlive = True Then
            webObj = GetSession()
        Else
            webObj = New ZukamiLib.WebSession(GetZukamiSettings)
            webObj.OpenConnection()
        End If

        Try
            webObj.CustomSQLCommand("SELECT * FROM [" & TableName & "] WHERE [" & FieldName & "]=" & CStr(Value))
            webObj.CustomClearParameters()
            GetDataset2 = webObj.CustomSQLExecuteReturn
        Catch ex As Exception
        End Try
errEnd:
        If KeepConnAlive = False Then
            webObj.CloseConnection()
            webObj = Nothing
        End If

    End Function

    Public ReadOnly Property PrimaryID() As String
        Get
            Try
                If WebconfigSettings.PrimaryIDIsAlwaysParent = True Then
                    Return CurrentForm.ParentRecordID
                Else
                    Return CurrentForm.RecordID
                End If
            Catch ex As Exception
                logger.Warn(ex)
                Return Guid.Empty.ToString()
            End Try
        End Get
    End Property

    Public ReadOnly Property CurrentFormPrimaryID() As String
        Get
            Return CurrentForm.RecordID
        End Get
    End Property

    Public ReadOnly Property ParentPrimaryID() As String
        Get
            Return CurrentForm.ParentRecordID
        End Get
    End Property

    Public Function Unzip(ByVal Filepath As String, ByVal TargetPath As String) As String
        Try
            Dim _io As New Ionic.Zip.ZipFile(Filepath)
            _io.FlattenFoldersOnExtract = True
            _io.ExtractAll(TargetPath)
            Return ""
        Catch ex As Exception
            Return ex.ToString
        End Try

    End Function

    Public Sub DoSave()
        If _FormAttributes Is Nothing = False Then
            If _FormAttributes.FormObject Is Nothing = False Then
                _FormAttributes.FormObject.DoFormSave()
            End If
        End If
    End Sub

    Public Sub DoAutoSave()
        If _FormAttributes Is Nothing = False Then
            If _FormAttributes.FormObject Is Nothing = False Then
                _FormAttributes.FormObject.DoAutoSave()
            End If
        End If
    End Sub

    Public Function CreateZIP(ByRef FilePathCollection As Collection, ByVal DestinationFileName As String, Optional ByVal DestinationFieldName As String = "") As String

        Dim _newfileguid As String = Guid.NewGuid.ToString
        Dim _fpath As String = GetZukamiSettings.UploadPath.TrimEnd("\") & "\" & _newfileguid

        Try
            System.IO.Directory.CreateDirectory(_fpath)
        Catch ex As Exception

        End Try

        Dim _io As New Ionic.Zip.ZipFile(_fpath.TrimEnd("\") & "\" & DestinationFileName)

        Dim _counter As Integer
        For _counter = 1 To FilePathCollection.Count
            Try
                Dim _fpath2 As String = FilePathCollection.Item(_counter)
                If System.IO.File.Exists(_fpath2) Then
                    _io.AddFile(_fpath2, "")
                End If
            Catch ex As Exception

            End Try

        Next _counter
        _io.Save()
        If Len(DestinationFieldName) > 0 Then
            Data(DestinationFieldName) = _newfileguid & ";" & DestinationFileName
        End If
        Return _newfileguid & ";" & DestinationFileName
    End Function

    Public Sub DisplayCycleMessage()
        If CurrentForm.IsPostBack = False And Len(QueryString("CycMsg")) > 0 Then
            DisplayMessage(QueryString("CycMsg"))
        End If
    End Sub

    Public Sub CycleToNextRow(ByVal TableFieldID As String, Optional ByVal DoForNew As Boolean = True, Optional ByVal DoForEdit As Boolean = False, Optional ByVal InsertSuccessMessage As String = "", Optional ByVal UpdateSuccessMessage As String = "")
        If (CurrentForm.IsEditMode = True) Then
            If DoForEdit = True Then
                Dim _currentID As String = QueryString("ID")
                If Len(_currentID) > 0 Then
                    Dim _set As DataSet = Session(TableFieldID)
                    Dim _idx As Long = GetRowIndex(_set.Tables(0), _currentID)
                    If _idx = -1 Or _idx = _set.Tables(0).Rows.Count - 1 Then
                        'at the last already, we do not do anything (go back to parent form - default behavior)

                    Else
                        'proceed down to next notch
                        Dim _nextID As String = GetText(_set.Tables(0).Rows(_idx + 1).Item("ID"))
                        Redirect("FillForm.aspx?" & IIf(Len(UpdateSuccessMessage) > 0, "CycMsg=" & URLEncode(UpdateSuccessMessage) & "&", "") & "ID=" & _nextID & "&a=" & QueryString("a") & "&ListID=" & QueryString("ListID") & "&ParentRec=" & QueryString("ParentRec"))
                    End If
                End If
            End If
        Else
            If DoForNew = True Then
                Redirect("FillForm.aspx?" & IIf(Len(InsertSuccessMessage) > 0, "CycMsg=" & URLEncode(InsertSuccessMessage) & "&", "") & "ParentRec=" & QueryString("ParentRec") & "&a=" & QueryString("a") & "&ListID=" & QueryString("ListID"))
            End If
        End If
    End Sub

    Public Function GetRowIndex(ByRef TableData As DataTable, ByVal ID As String) As Long
        Dim _counter As Integer
        For _counter = 0 To TableData.Rows.Count - 1
            If StrComp(GlobalFunctions.FormatData(TableData.Rows(_counter).Item("ID")), ID, CompareMethod.Text) = 0 Then
                Return _counter
            End If
        Next _counter
        Return -1
    End Function

    Public Property Session(ByVal ID As String) As Object
        Get
            Return System.Web.HttpContext.Current.Session(ID)
        End Get
        Set(value As Object)
            System.Web.HttpContext.Current.Session(ID) = value
        End Set
    End Property

    Public ReadOnly Property ParentSession() As DataSet
        Get
            Return System.Web.HttpContext.Current.Session(GetParentToken)
        End Get
    End Property


    Public Function GetActionsPerformedByTaskID(TaskID As String) As DataSet


        Dim webObj As New ZukamiLib.WebSession(GetZukamiSettings)
        webObj.OpenConnection()

        webObj.CustomSQLCommand("select a.*,c.*,d.*,c.Remarks AS ActionRemarks,c.[Signature] AS ActionSignature from tasks a inner join instancebubbles b on a.instancenodeid=b.id inner join taskactions c on a.activityID=c.ActionID left join users d ON c.performedby = d.userid where a.taskid='" & TaskID & "' ORDER BY c.DatePerformed ASC")
        webObj.CustomClearParameters()
        GetActionsPerformedByTaskID = webObj.CustomSQLExecuteReturn()

        webObj.CloseConnection()
        webObj = Nothing
    End Function

    Public Function GetActionsPerformedByNodeID(InstanceID As Guid, NodeID As String) As DataSet
        Dim webObj As New ZukamiLib.WebSession(GetZukamiSettings)
        webObj.OpenConnection()

        webObj.CustomSQLCommand("select a.*,c.*,d.*,c.Remarks AS ActionRemarks,c.[Signature] AS ActionSignature from tasks a inner join instancebubbles b on a.instancenodeid=b.id inner join taskactions c on a.activityID=c.ActionID left join users d ON c.performedby = d.userid where a.instanceid='" & InstanceID.ToString & "'  and nodeid LIKE '" & NodeID & "' ORDER BY c.DatePerformed ASC")
        webObj.CustomClearParameters()
        GetActionsPerformedByNodeID = webObj.CustomSQLExecuteReturn()

        webObj.CloseConnection()
        webObj = Nothing

    End Function


    Public Sub SetParentSession(ByRef mSet As DataSet)
        System.Web.HttpContext.Current.Session(GetParentToken) = mSet
    End Sub

    Public Sub PopulateGrid(ByVal TableField As String, ByVal TableName As String, ByVal FieldName As String, ByVal Value As Object, Optional ByVal Append As Boolean = False)
        Dim webObj As ZukamiLib.WebSession = Nothing
        If KeepConnAlive = True Then
            webObj = GetSession()
        Else
            webObj = New ZukamiLib.WebSession(GetZukamiSettings)
            webObj.OpenConnection()
        End If

        Try
            webObj.CustomSQLCommand("SELECT * FROM [" & TableName & "] WHERE [" & FieldName & "]=" & CStr(Value))
            webObj.CustomClearParameters()
            Dim _returnedSet As DataSet = webObj.CustomSQLExecuteReturn

            Dim _currentSet As DataTable = Data(TableField)
            If _currentSet Is Nothing = False Then
                If Append = False Then _currentSet.Rows.Clear()
                Dim _counter As Integer
                For _counter = 0 To _returnedSet.Tables(0).Rows.Count - 1
                    Dim _newrow As DataRow = _currentSet.NewRow
                    _newrow.Item("ID") = Guid.NewGuid

                    For _counter2 = 0 To _returnedSet.Tables(0).Columns.Count - 1
                        Dim _colName As String = _returnedSet.Tables(0).Columns.Item(_counter2).ColumnName
                        If StrComp(_colName, "ID", CompareMethod.Text) <> 0 And StrComp(_colName, "ParentID", CompareMethod.Text) <> 0 Then
                            'find if it exists in currentset
                            If _currentSet.Columns.Contains(_colName) = True Then
                                _newrow.Item(_colName) = _returnedSet.Tables(0).Rows(_counter).Item(_colName)
                            End If
                        End If
                    Next _counter2
                    _currentSet.Rows.Add(_newrow)
                Next
            End If



            Data(TableField) = _currentSet
        Catch ex As Exception
            DisplayMessage(ex.ToString)
        End Try
errEnd:
        If KeepConnAlive = False Then
            webObj.CloseConnection()
            webObj = Nothing
        End If

    End Sub


    Public Sub CommitDataset(ByRef Data As DataSet, ByVal TableName As String)
        Dim webObj As ZukamiLib.WebSession = Nothing
        If KeepConnAlive = True Then
            webObj = GetSession()
        Else
            webObj = New ZukamiLib.WebSession(GetZukamiSettings)
            webObj.OpenConnection()
        End If

        webObj.CommitChanges(Data, TableName)

        If KeepConnAlive = False Then
            webObj.CloseConnection()
            webObj = Nothing
        End If

    End Sub


    Public Function GetData(ByVal FormName As String, ByVal FieldName As String, ByVal Value As Object) As DataRow
        GetData = Nothing
        Dim webObj As ZukamiLib.WebSession = Nothing
        If KeepConnAlive = True Then
            webObj = GetSession()
        Else
            webObj = New ZukamiLib.WebSession(GetZukamiSettings)
            webObj.OpenConnection()
        End If

        Try
            Dim _FormSet As DataSet = webObj.Forms_GetByFormName(FormName)
            If _FormSet.Tables(0).Rows.Count = 0 Then GoTo errEnd
            Dim _tbs As String = GlobalFunctions.FormatData(_FormSet.Tables(0).Rows(0).Item("TableBindSource"))

            webObj.CustomSQLCommand("SELECT * FROM [" & _tbs & "] WHERE [" & FieldName & "]=" & CStr(Value))
            webObj.CustomClearParameters()
            Dim _resultset As DataSet = webObj.CustomSQLExecuteReturn
            If _resultset.Tables(0).Rows.Count > 0 Then
                GetData = _resultset.Tables(0).Rows(0)
            End If
        Catch ex As Exception
        End Try
errEnd:
        If KeepConnAlive = False Then
            webObj.CloseConnection()
            webObj = Nothing
        End If

    End Function

    Public Function Compare(ByVal String1 As String, ByVal string2 As String) As String
        Dim _mainstring As String = String1
        Dim arrsplits() As String = Split(string2, " ")
        Dim mycoll As New Collection

        Dim _counter As Integer
        For _counter = 0 To UBound(arrsplits)
            Dim _tag As String = arrsplits(_counter)
            If Len(_tag) > 0 Then
                If mycoll.Contains(_tag) = False Then
                    _mainstring = Replace(_mainstring, _tag, "<span style=""background-color: #FFFF00"">" & _tag & "</span>", , , CompareMethod.Text)
                    mycoll.Add(_tag, _tag)
                End If
            End If
        Next
        Return _mainstring
    End Function

    'Public Sub RunSQLRemote(ByVal SQL As String)
    '    Dim _zed As New ZEDMain.ZED
    '    _zed.RunSQL(GetZukamiSettings.CurrentUserGUID.ToString, SQL)


    '    Exit Sub
    'End Sub

    'Public Sub RunSQLRemoteAndSubmit(ByVal SQL As String, ByVal FormID As String, ByVal RecordID As String)
    '    Dim _zed As New ZEDMain.ZED
    '    _zed.RunSQLAndSubmit(GetZukamiSettings.CurrentUserGUID.ToString, SQL, FormID, RecordID)
    '    Exit Sub
    'End Sub

    'Public Sub LaunchWorkflowRemote(ByVal FormID As String, ByVal RecordID As String, ByVal Remarks As String)
    '    Dim _zed As New ZEDMain.ZED
    '    _zed.LaunchFormWorkflowNoLogin(GetZukamiSettings.CurrentUserGUID.ToString, FormID, RecordID, Remarks)
    'End Sub

    'Public Function RunSQLReturnRemote(ByVal SQL As String) As DataSet
    '    Dim _zed As New ZEDMain.ZED
    '    Return _zed.RunDynamicSQLReturn(GetZukamiSettings.CurrentUserGUID.ToString, SQL)
    'End Function

    Public Sub RunSQL(ByVal SQL As String)
        Dim webObj As ZukamiLib.WebSession = Nothing
        If KeepConnAlive = True Then
            webObj = GetSession()
        Else
            webObj = New ZukamiLib.WebSession(GetZukamiSettings)
            webObj.OpenConnection()
        End If

        Try
            webObj.CustomSQLCommand(SQL)
            webObj.CustomClearParameters()
            webObj.CustomSQLExecute()
        Catch ex As Exception
        End Try
errEnd:
        If KeepConnAlive = False Then
            webObj.CloseConnection()
            webObj = Nothing
        End If

    End Sub

    Public Sub WriteFile(ByVal filename As String)
        System.Web.HttpContext.Current.Response.WriteFile(filename)
    End Sub

    Public Function RunReturnSQL(ByVal SQL As String) As DataSet
        Dim webObj As ZukamiLib.WebSession = Nothing
        If KeepConnAlive = True Then
            webObj = GetSession()
        Else
            webObj = New ZukamiLib.WebSession(GetZukamiSettings)
            webObj.OpenConnection()
        End If

        Try
            webObj.CustomSQLCommand(SQL)
            webObj.CustomClearParameters()
            RunReturnSQL = webObj.CustomSQLExecuteReturn
        Catch ex As Exception
        End Try
errEnd:
        If KeepConnAlive = False Then
            webObj.CloseConnection()
            webObj = Nothing
        End If

    End Function

    Private Function ValidateJString(Message As String) As String
        Dim _temp As String = Message
        _temp = Replace(_temp, vbCrLf, "\n")
        _temp = Replace(_temp, vbCr, "\n")
        _temp = Replace(_temp, vbLf, "\n")
        _temp = Replace(_temp, "'", "\'")
        Return _temp
    End Function

    Public Sub SendSignalRMessage(Hub As String, From As String, Message As String, Optional DisplayDebug As Boolean = True)
        Dim _contscript As String = "$(function () {" &
                                        "var chat = $.connection.tMChatHub;"

        _contscript += "chat.client.broadcastMessage = function(name, message) {"
        If DisplayDebug = True Then _contscript += "alert(message);"
        _contscript += "__doPostBack('ctl00_UpdatePanel1$DefaultButtonClick',message);};"


        _contscript += "$.connection.hub.start().done(function () {"
        If DisplayDebug = True Then _contscript += "alert('Connected');"
        _contscript += "chat.server.sendToWholeWorld('" & From & "','" & Message & "').done(function () {"
        If DisplayDebug = True Then _contscript += "alert('Sent!');"
        _contscript += "});" &
        "});" &
        "});"

        Dim _script As String = "var url = 'Scripts/jquery.signalR-2.2.0.min.js'; $.getScript(url, function() {  $.getScript('" & Hub & "', function() {" & _contscript & "});      });"
        RegisterStartupScript("sendmsg", _script, True)
    End Sub

    Public Sub RegisterSignalR(Hub As String, Optional DisplayDebug As Boolean = True)
        Dim _contscript As String = "$(function () {" &
                                        "var chat = $.connection.tMChatHub;"

        _contscript += "chat.client.broadcastMessage = function(name, message) {"
        If DisplayDebug = True Then _contscript += "alert(message);"
        _contscript += "__doPostBack('ctl00_UpdatePanel1$DefaultButtonClick',message);};"


        _contscript += "$.connection.hub.start().done(function () {"
        If DisplayDebug = True Then _contscript += "alert('Connected');"
        _contscript += "});" &
        "});"

        Dim _script As String = "var url = 'Scripts/jquery.signalR-2.2.0.min.js'; $.getScript(url, function() {  $.getScript('" & Hub & "', function() {" & _contscript & "});      });"
        RegisterStartupScript("signalr", _script, True)
    End Sub

    Public Sub DeleteData(ByVal FormName As String, ByVal FieldName As String, ByVal Value As Object)
        Dim webObj As ZukamiLib.WebSession = Nothing
        If KeepConnAlive = True Then
            webObj = GetSession()
        Else
            webObj = New ZukamiLib.WebSession(GetZukamiSettings)
            webObj.OpenConnection()
        End If

        Try
            Dim _FormSet As DataSet = webObj.Forms_GetByFormName(FormName)
            If _FormSet.Tables(0).Rows.Count = 0 Then GoTo errEnd
            Dim _tbs As String = GlobalFunctions.FormatData(_FormSet.Tables(0).Rows(0).Item("TableBindSource"))

            webObj.CustomSQLCommand("DELETE FROM [" & _tbs & "] WHERE [" & FieldName & "]=" & CStr(Value))
            webObj.CustomClearParameters()
            webObj.CustomSQLExecute()
        Catch ex As Exception
        End Try
errEnd:
        If KeepConnAlive = False Then
            webObj.CloseConnection()
            webObj = Nothing
        End If

    End Sub

    Public Property IsCompulsory(ByVal FieldName As String) As Boolean
        Get
            Dim result As Boolean = False
            Dim _zfield As ZField = GetZField(FieldName)
            If _zfield Is Nothing = False Then
                result = _zfield.IsCompulsory
            End If
            Return result
        End Get
        Set(value As Boolean)
            SetCompulsory(FieldName, value)
        End Set
    End Property

    Public Property IsReadOnly(ByVal FieldName As String) As Boolean
        Get
            If _FormDetails.IsReadonlyMode = True Then Return True

            If Len(FieldName) = 0 Then Return False
            If _zfields.Contains(FieldName) = False Then Return False
            Dim _temp As ZField = _zfields.Item(FieldName)
            Return Not _temp.Enabled
        End Get
        Set(ByVal value As Boolean)
            If _FormDetails.IsReadonlyMode = True Then Exit Property
            Try
                If Len(FieldName) = 0 Then Return
                If _zfields.Contains(FieldName) = False Then Return
                Dim _temp As ZField = _zfields.Item(FieldName)

                'save the data from the control to the readonly field
                _temp.Enabled = Not value

                'COMMIT THE DATA TO THE SESSION
                Try
                    If _FormDetails.IsSubform = False Then
                        Dim _MainSet As DataSet = System.Web.HttpContext.Current.Session(GetParentToken)
                        _MainSet.EnforceConstraints = False
                        Dim _DataRow As DataRow = _MainSet.Tables(0).Rows(0)
                        GlobalFunctions.CommitZFieldToDataRow(_temp, _DataRow, False)
                        _MainSet.AcceptChanges()
                        System.Web.HttpContext.Current.Session(GetParentToken) = _MainSet
                    End If
                Catch ex As Exception
                    logger.Error(ex)
                End Try

                GlobalFunctions.BindReadonlyLabels(_temp)
                Select Case _temp.FieldType
                    Case GlobalFunctions.FIELDTYPES.FT_FILE
                        CType(_temp.FieldControl, Object).readonlymode = value
                        CType(_temp.FieldControl, Object).RefreshView()
                    Case GlobalFunctions.FIELDTYPES.FT_TABLE
                        CType(_temp.FieldControl, Object).readonlymode = value
                        If GlobalFunctions.IsCurrentRequestMobileVersion() Then
                            CType(_temp.FieldControl, Object).RefreshGrid()
                        End If
                        CType(_temp.FieldControl, Object).RefreshReadOnly()
                    Case GlobalFunctions.FIELDTYPES.FT_FRAME
                    Case GlobalFunctions.FIELDTYPES.FT_HEADER
                    Case GlobalFunctions.FIELDTYPES.FT_HIDDENFIELD
                    Case GlobalFunctions.FIELDTYPES.FT_BUTTON
                    Case GlobalFunctions.FIELDTYPES.FT_CURRENCY
                        GlobalFunctions.CSSShow(_temp.FieldControl, Not value)
                        '_temp.FieldControl.visible = Not value
                        _temp.ReadOnlyControl.visible = value
                    Case GlobalFunctions.FIELDTYPES.FT_LABEL, GlobalFunctions.FIELDTYPES.FT_DBLABEL
                    Case GlobalFunctions.FIELDTYPES.FT_TIFFVIEWER, GlobalFunctions.FIELDTYPES.FT_CAMERA, GlobalFunctions.FIELDTYPES.FT_SIGNATURE
                        CType(_temp.FieldControl, Object).readonlymode = value
                        CType(_temp.FieldControl, Object).RefreshView()
                    Case GlobalFunctions.FIELDTYPES.FT_CHECKLIST

                        Dim _coll As Collection = _temp.FieldControl
                        Dim _tbl As Table = _temp.ReadOnlyControl

                        _tbl.Visible = value
                        If GlobalFunctions.IsCurrentRequestMobileVersion() Then
                            If _temp.EditModeControlsContainer IsNot Nothing Then
                                CType(_temp.EditModeControlsContainer, Panel).Visible = Not value
                            Else
                                logger.Error("EditModeControlsContainer is nothing for field: " + _temp.FieldName)
                            End If
                        Else
                            SetCollectionItemsVisible(_coll, Not value, _tbl, True)
                        End If

                    Case GlobalFunctions.FIELDTYPES.FT_RADIO
                        Dim _coll As Collection = _temp.FieldControl
                        Dim _lbl As Label = _temp.ReadOnlyControl

                        _lbl.Visible = value

                        If GlobalFunctions.IsCurrentRequestMobileVersion() Then
                            If _temp.EditModeControlsContainer IsNot Nothing Then
                                CType(_temp.EditModeControlsContainer, Panel).Visible = Not value
                            Else
                                logger.Error("EditModeControlsContainer is nothing for field: " + _temp.FieldName)
                            End If
                        Else
                            SetCollectionItemsVisible(_coll, Not value, _lbl, False)
                        End If
                    Case GlobalFunctions.FIELDTYPES.FT_SHORTTEXT, GlobalFunctions.FIELDTYPES.FT_INT, GlobalFunctions.FIELDTYPES.FT_FLOAT, GlobalFunctions.FIELDTYPES.FT_LONGTEXT
                        GlobalFunctions.CSSShow(_temp.FieldControl, Not value)
                        _temp.ReadOnlyControl.visible = value
                    Case GlobalFunctions.FIELDTYPES.FT_RATING
                        _temp.ReadOnlyControl.visible = value
                        If Not GlobalFunctions.IsCurrentRequestMobileVersion Then
                            _temp.FieldControl.visible = Not value
                        Else
                            Dim x As DropDownList = CType(_temp.FieldControl, DropDownList)
                            If (x IsNot Nothing) Then
                                x.Attributes.Remove("ShouldHide")
                                If value Then
                                    x.Attributes.Add("ShouldHide", "true")
                                End If
                            End If
                        End If
                    Case Else
                        _temp.FieldControl.visible = Not value
                        _temp.ReadOnlyControl.visible = value

                End Select
            Catch ex As Exception
                logger.Error(ex)
            End Try
        End Set
    End Property

    Private Sub SetCollectionItemsVisible(ByRef Coll As Collection, ByVal Visible As Boolean, ByRef Exception As Object, ByVal IsCheckList As Boolean)
        Dim _counter As Integer
        Dim _tableCell As TableCell = Nothing
        For _counter = 1 To Coll.Count
            Dim obj As Object = Coll.Item(_counter)
            obj.visible = Visible

            'Get the parent table cell
            _tableCell = CType(obj, Control).Parent
        Next _counter
        'We hide the other label controls
        If _tableCell Is Nothing = False Then
            For _counter = 0 To _tableCell.Controls.Count - 1
                If Not Exception Is _tableCell.Controls(_counter) Then
                    If TypeOf _tableCell.Controls(_counter) Is Label Then
                        _tableCell.Controls(_counter).Visible = Visible
                    End If
                    If IsCheckList = True Then
                        'If TypeOf _tableCell.Controls(_counter) Is Table Then
                        ' _tableCell.Controls(_counter).Visible = Visible
                        'End If
                    End If
                End If
            Next _counter
        End If
    End Sub

    Public Function GetFormButton(ByVal ButtonName As String) As Object
        If _FormButtonsList Is Nothing Then Return Nothing

        Dim _counter As Integer
        For _counter = 0 To _FormButtonsList.Controls.Count - 1
            Dim _ctrl As Object = _FormButtonsList.Controls(_counter)
            If TypeOf _ctrl Is Button Or TypeOf _ctrl Is ImageButton Then
                If StrComp(_ctrl.attributes.item("ButtonName"), ButtonName, CompareMethod.Text) = 0 Then
                    Return _ctrl
                End If
            End If
        Next
        Return Nothing
    End Function

    Public Sub HideAllButtons()
        If _FormButtonsList Is Nothing Then Return

        Dim _counter As Integer
        For _counter = 0 To _FormButtonsList.Controls.Count - 1
            Dim _ctrl As Object = _FormButtonsList.Controls(_counter)
            If TypeOf _ctrl Is Button Or TypeOf _ctrl Is ImageButton Then
                Select Case LCase(_ctrl.attributes.item("ButtonName"))
                    Case "save", "cancel", "submit", "confirm", "print", "edit", "resubmit", "checkprevious", "changehistory", "discardchanges", "returntoview", "cancelsave"
                    Case Else
                        Dim _butType As String = _ctrl.attributes.item("ButtonType")
                        If _butType = "1" Then
                            _ctrl.Visible = False
                        End If
                End Select
            End If
        Next
        Return
    End Sub

    Public Property FormButtonCaption(ByVal ButtonName As String) As String
        Get
            Dim _button As Object = GetFormButton(ButtonName)
            If _button Is Nothing Then Return False
            Return _button.Text
        End Get
        Set(ByVal value As String)
            Dim _button As Object = GetFormButton(ButtonName)
            If _button Is Nothing Then Return
            _button.Text = value
        End Set
    End Property

    Public Property FormButtonVisible(ByVal ButtonName As String) As Boolean
        Get
            Dim _button As Object = GetFormButton(ButtonName)
            If _button Is Nothing Then Return False
            Return _button.Visible
        End Get
        Set(ByVal value As Boolean)
            Dim _button As Object = GetFormButton(ButtonName)
            If _button Is Nothing Then Return
            _button.Visible = value
        End Set
    End Property

    Public Property FormButtonEnabled(ByVal ButtonName As String) As Boolean
        Get
            Dim _button As Object = GetFormButton(ButtonName)
            If _button Is Nothing Then Return False
            Return _button.Enabled
        End Get
        Set(ByVal value As Boolean)
            Dim _button As Object = GetFormButton(ButtonName)
            If _button Is Nothing Then Return
            _button.Enabled = value
        End Set
    End Property

    Private Function AddCssClass(current As String, cssClass As String) As String
        Dim result As String = current
        cssClass = cssClass.Trim()
        If result.Contains(cssClass) Then
            ' already contains the css class, do nothing
        Else
            result = (result + " " + cssClass).Trim()
        End If
        Return result
    End Function
    Private Function RemoveCssClass(current As String, cssClass As String) As String
        Dim result As String = current
        cssClass = cssClass.Trim()
        If result.Contains(cssClass) Then
            result = result.Replace(cssClass, "").Trim()
        End If
        Return result
    End Function
    'actionType: Add | Remove
    Private Function ModifyCssClass(current As String, cssClass As String, actionType As String) As String
        If current Is Nothing Then current = ""
        If cssClass Is Nothing Then cssClass = ""
        If actionType = "Add" Then
            Return AddCssClass(current, cssClass)
        ElseIf actionType = "Remove" Then
            Return RemoveCssClass(current, cssClass)
        Else
            Return current ' do nothing
        End If
    End Function
    Public Sub ModifyFieldCSS(ByVal Fieldname As String, cssClass As String, actionType As String)
        ModifyFieldCaptionCSS(Fieldname, cssClass, actionType)
        ModifyFieldControlCSS(Fieldname, cssClass, actionType)
    End Sub
    Public Sub ModifyFieldCaptionCSS(ByVal Fieldname As String, cssClass As String, actionType As String)
        logger.Debug("Fieldname: " + Fieldname + ", cssClass: " + cssClass + ", actionType: " + actionType)
        Dim i As Integer
        For i = 1 To _zfields.Count
            Dim f As ZField = _zfields.Item(i)
            If StrComp(Fieldname, f.FieldName, CompareMethod.Text) = 0 Then
                logger.Debug("found the zfield, FieldType: " + f.FieldType.ToString())
                Try
                    If f.FieldLabelControl Is Nothing = False Then
                        If TypeOf f.FieldLabelControl Is Label Then
                            Dim fieldLabel As Label = CType(f.FieldLabelControl, Label)
                            fieldLabel.CssClass = ModifyCssClass(fieldLabel.CssClass, cssClass, actionType)
                        Else
                            logger.Debug("TypeOf FieldLabelControl is not label, skip")
                        End If
                    Else
                        logger.Debug("FieldLabelControl is nothing, skip")
                    End If
                Catch ex As Exception
                    logger.Error(ex, "Fieldname: " + Fieldname + ", FieldType: " + f.FieldType.ToString())
                End Try
            End If
        Next i
    End Sub
    Public Sub ModifyFieldControlCSS(ByVal Fieldname As String, cssClass As String, actionType As String)
        logger.Debug("Fieldname: " + Fieldname + ", cssClass: " + cssClass + ", actionType: " + actionType)
        Dim i As Integer
        For i = 1 To _zfields.Count
            Dim f As ZField = _zfields.Item(i)
            If StrComp(Fieldname, f.FieldName, CompareMethod.Text) = 0 Then
                logger.Debug("found the zfield, FieldType: " + f.FieldType.ToString())
                Try
                    If f.FieldType = GlobalFunctions.FIELDTYPES.FT_CHECKLIST OrElse
                       f.FieldType = GlobalFunctions.FIELDTYPES.FT_RADIO Then
                        Dim _coll As Collection = f.FieldControl
                        If _coll Is Nothing = False Then
                            For j As Integer = 1 To _coll.Count
                                If _coll.Item(j) Is Nothing = False Then _coll.Item(j).LabelAttributes("class") = ModifyCssClass(_coll.Item(j).LabelAttributes("class"), cssClass, actionType)
                            Next
                        End If
                    Else
                        If f.FieldControl Is Nothing = False Then f.FieldControl.cssclass = ModifyCssClass(f.FieldControl.cssclass, cssClass, actionType)
                    End If
                Catch ex As Exception
                    logger.Error(ex, "Fieldname: " + Fieldname + ", FieldType: " + f.FieldType.ToString())
                End Try
                Try
                    If f.FieldControl2 Is Nothing = False Then f.FieldControl2.cssclass = ModifyCssClass(f.FieldControl2.cssclass, cssClass, actionType)
                Catch ex As Exception
                    logger.Error(ex, "Fieldname: " + Fieldname + ", FieldType: " + f.FieldType.ToString())
                End Try
            End If
        Next i
    End Sub

    Public WriteOnly Property FieldCSS(ByVal Fieldname As String) As Object
        Set(ByVal value As Object)
            Dim _counter As Integer
            For _counter = 1 To _zfields.Count
                Dim _temp As ZField = _zfields.Item(_counter)
                If StrComp(Fieldname, _temp.FieldName, CompareMethod.Text) = 0 Then
                    Try
                        If _temp.FieldControl Is Nothing = False Then _temp.FieldControl.cssclass = value
                    Catch ex As Exception

                    End Try
                    Try
                        If _temp.FieldControl2 Is Nothing = False Then _temp.FieldControl2.cssclass = value
                    Catch ex As Exception

                    End Try
                End If
            Next _counter
        End Set
    End Property

    Public Property Data(ByVal FieldName As String) As Object
        Get
            If _zfields.Contains(FieldName) = False Then Return Nothing
            Dim _temp As ZField = _zfields.Item(FieldName)

            Try

                Select Case _temp.FieldType
                    Case GlobalFunctions.FIELDTYPES.FT_SHORTTEXT
                        Return CStr(CType(_temp.FieldControl, TextBox).Text)
                    Case GlobalFunctions.FIELDTYPES.FT_BARCODE
                        Return CStr(CType(_temp.FieldControl, TextBox).Text)
                    Case GlobalFunctions.FIELDTYPES.FT_LONGTEXT
                        Return CStr(CType(_temp.FieldControl, TextBox).Text)
                    Case GlobalFunctions.FIELDTYPES.FT_HTML
                        Return CStr(_temp.FieldControl.value)
                    Case GlobalFunctions.FIELDTYPES.FT_COUNTRY
                        Return CStr(CType(_temp.FieldControl, DropDownList).SelectedValue)
                    Case GlobalFunctions.FIELDTYPES.FT_CURRENCY, GlobalFunctions.FIELDTYPES.FT_FLOAT
                        Return GlobalFunctions.FormatDouble(CType(_temp.FieldControl, TextBox).Text)
                    Case GlobalFunctions.FIELDTYPES.FT_INT
                        Return GlobalFunctions.FormatInteger(CType(_temp.FieldControl, TextBox).Text)
                    Case GlobalFunctions.FIELDTYPES.FT_CHECKLIST
                        Return CStr(GlobalFunctions.GetCheckListValue(CType(_temp.FieldControl, Collection)))
                    Case GlobalFunctions.FIELDTYPES.FT_RADIO
                        Return CStr(GlobalFunctions.GetRadioListValue(CType(_temp.FieldControl, Collection)))
                    Case GlobalFunctions.FIELDTYPES.FT_DATE, GlobalFunctions.FIELDTYPES.FT_DATETIME
                        If _temp.FieldControl.isempty Then
                            Return ""
                        Else
                            Return GlobalFunctions.GetDateTime(CType(_temp.FieldControl, Object).value)
                        End If
                    Case GlobalFunctions.FIELDTYPES.FT_DROPDOWN
                        Return CType(_temp.FieldControl, Object).text
                    Case GlobalFunctions.FIELDTYPES.FT_GPS
                        Return CType(_temp.FieldControl, Object).text
                    Case GlobalFunctions.FIELDTYPES.FT_YESNO
                        Return GlobalFunctions.FormatBoolean(CType(_temp.FieldControl, Object).SelectedValue)
                    Case GlobalFunctions.FIELDTYPES.FT_AUTOID, GlobalFunctions.FIELDTYPES.FT_LABEL, GlobalFunctions.FIELDTYPES.FT_DBLABEL
                        Return CStr(CType(_temp.FieldControl, Label).Text)
                    Case GlobalFunctions.FIELDTYPES.FT_FILE
                        Return CStr(CType(_temp.FieldControl, Object).Getinternalpath)
                    Case GlobalFunctions.FIELDTYPES.FT_HIDDENFIELD
                        Return CStr(CType(_temp.FieldControl, HiddenField).Value)
                    Case GlobalFunctions.FIELDTYPES.FT_USER
                        Return CStr(CType(_temp.FieldControl, Object).text)
                    Case GlobalFunctions.FIELDTYPES.FT_IMAGE, GlobalFunctions.FIELDTYPES.FT_TIFFVIEWER, GlobalFunctions.FIELDTYPES.FT_SIGNATURE, GlobalFunctions.FIELDTYPES.FT_CAMERA
                        Return CStr(CType(_temp.FieldControl, Object).GetInternalPath)
                    Case GlobalFunctions.FIELDTYPES.FT_TABLE
                        Return CType(_temp.FieldControl, Object).DataSource
                    Case GlobalFunctions.FIELDTYPES.FT_RATING
                        Return CStr(CType(_temp.FieldControl, Object).text)
                End Select
                Return Nothing
            Catch ex As Exception
                DisplayMessage("Could not get data value for [" & _temp.FieldName & "] field<br>" & ex.ToString)
            End Try
            Return Nothing
        End Get
        Set(ByVal value As Object)
            If _zfields.Contains(FieldName) = False Then Return
            Dim _temp As ZField = _zfields.Item(FieldName)


            Try
                Select Case _temp.FieldType
                    Case GlobalFunctions.FIELDTYPES.FT_SHORTTEXT, GlobalFunctions.FIELDTYPES.FT_LONGTEXT
                        CType(_temp.FieldControl, TextBox).Text = value
                    Case GlobalFunctions.FIELDTYPES.FT_BARCODE
                        CType(_temp.FieldControl, TextBox).Text = value
                    Case GlobalFunctions.FIELDTYPES.FT_HTML
                        _temp.FieldControl.value = value
                    Case GlobalFunctions.FIELDTYPES.FT_GPS
                        CType(_temp.FieldControl, Object).text = value
                    Case GlobalFunctions.FIELDTYPES.FT_COUNTRY
                        CType(_temp.FieldControl, DropDownList).SelectedValue = value
                    Case GlobalFunctions.FIELDTYPES.FT_CURRENCY
                        CType(_temp.FieldControl, TextBox).Text = GlobalFunctions.FormatNumberForDisplay(value, GlobalFunctions.FromConfig("MoneyPrecision", 2))
                    Case GlobalFunctions.FIELDTYPES.FT_FLOAT
                        CType(_temp.FieldControl, TextBox).Text = GlobalFunctions.FormatNumberForDisplay(value, _temp.Arguments)
                    Case GlobalFunctions.FIELDTYPES.FT_INT
                        CType(_temp.FieldControl, TextBox).Text = GlobalFunctions.FormatNumberForDisplay(value, "")
                    Case GlobalFunctions.FIELDTYPES.FT_CHECKLIST
                        GlobalFunctions.SetCheckListValue(CType(_temp.FieldControl, Collection), value)
                    Case GlobalFunctions.FIELDTYPES.FT_RADIO
                        GlobalFunctions.SetRadioListValue(CType(_temp.FieldControl, Collection), value)
                    Case GlobalFunctions.FIELDTYPES.FT_DATE
                        CType(_temp.FieldControl, Object).value = GlobalFunctions.FormatDate(value)
                    Case GlobalFunctions.FIELDTYPES.FT_DATETIME
                        CType(_temp.FieldControl, Object).value = GlobalFunctions.FormatDateTime(value)
                    Case GlobalFunctions.FIELDTYPES.FT_DROPDOWN
                        CType(_temp.FieldControl, Object).Text = value
                    Case GlobalFunctions.FIELDTYPES.FT_YESNO
                        CType(_temp.FieldControl, Object).SelectedValue = value
                    Case GlobalFunctions.FIELDTYPES.FT_AUTOID, GlobalFunctions.FIELDTYPES.FT_LABEL, GlobalFunctions.FIELDTYPES.FT_DBLABEL
                        CType(_temp.FieldControl, Label).Text = value
                    Case GlobalFunctions.FIELDTYPES.FT_FILE
                        CType(_temp.FieldControl, Object).setinternalpath(value)
                    Case GlobalFunctions.FIELDTYPES.FT_HIDDENFIELD
                        CType(_temp.FieldControl, HiddenField).Value = value
                    Case GlobalFunctions.FIELDTYPES.FT_USER
                        CType(_temp.FieldControl, Object).Text = value
                    Case GlobalFunctions.FIELDTYPES.FT_IMAGE, GlobalFunctions.FIELDTYPES.FT_TIFFVIEWER, GlobalFunctions.FIELDTYPES.FT_CAMERA, GlobalFunctions.FIELDTYPES.FT_SIGNATURE
                        CType(_temp.FieldControl, Object).SetInternalPath(value)
                    Case GlobalFunctions.FIELDTYPES.FT_TABLE
                        System.Web.HttpContext.Current.Session(_temp.FieldGUID.ToString) = CType(value, DataTable).DataSet
                        CType(_temp.FieldControl, Object).clearrows()
                        CType(_temp.FieldControl, Object).DataSource = value
                    Case GlobalFunctions.FIELDTYPES.FT_RATING
                        CType(_temp.FieldControl, Object).Text = value
                End Select
            Catch ex As Exception
                DisplayMessage("Could not set data value for [" & _temp.FieldName & "] field<br>" & ex.ToString)
            End Try

            GlobalFunctions.BindReadonlyLabels(_temp)
        End Set
    End Property



    Public Function FormatDateTime(ByVal Data As Object) As String
        Return GlobalFunctions.FormatDateTime(Data)
    End Function

    Public Property ControlText(ByVal FieldName As String) As Object
        Get
            Dim _counter As Integer
            For _counter = 1 To _zfields.Count
                Dim _temp As ZField = _zfields.Item(_counter)
                Try
                    If StrComp(FieldName, _temp.FieldName, CompareMethod.Text) = 0 Then
                        If TypeOf _temp.FieldControl Is Label Then
                            Return CStr(CType(_temp.FieldControl, Label).Text)
                        Else
                            Select Case _temp.FieldType
                                Case GlobalFunctions.FIELDTYPES.FT_SHORTTEXT
                                    Return CStr(CType(_temp.FieldControl, TextBox).Text)
                                Case GlobalFunctions.FIELDTYPES.FT_LONGTEXT
                                    Return CStr(CType(_temp.FieldControl, TextBox).Text)
                                Case GlobalFunctions.FIELDTYPES.FT_HTML
                                    Return CStr(_temp.FieldControl.text)
                                Case GlobalFunctions.FIELDTYPES.FT_COUNTRY
                                    Return CStr(CType(_temp.FieldControl, DropDownList).SelectedValue)
                                Case GlobalFunctions.FIELDTYPES.FT_CURRENCY, GlobalFunctions.FIELDTYPES.FT_FLOAT
                                    Return GlobalFunctions.FormatDouble(CType(_temp.FieldControl, TextBox).Text)
                                Case GlobalFunctions.FIELDTYPES.FT_INT
                                    Return GlobalFunctions.FormatInteger(CType(_temp.FieldControl, TextBox).Text)
                                Case GlobalFunctions.FIELDTYPES.FT_CHECKLIST
                                    Return CStr(GlobalFunctions.GetCheckListValue(CType(_temp.FieldControl, Collection)))
                                Case GlobalFunctions.FIELDTYPES.FT_RADIO
                                    Return CStr(GlobalFunctions.GetRadioListValue(CType(_temp.FieldControl, Collection)))
                                Case GlobalFunctions.FIELDTYPES.FT_DATE
                                    Return GlobalFunctions.GetDateTime(CType(_temp.FieldControl, Object).value)
                                Case GlobalFunctions.FIELDTYPES.FT_DATETIME
                                    Return GlobalFunctions.GetDateTime(CType(_temp.FieldControl, Object).value)
                                Case GlobalFunctions.FIELDTYPES.FT_DROPDOWN
                                    If CType(_temp.FieldControl, Object).SelectedItem Is Nothing = False Then
                                        Return CStr(CType(_temp.FieldControl, Object).SelectedItem.Text)
                                    Else
                                        Return ""
                                    End If
                                Case GlobalFunctions.FIELDTYPES.FT_YESNO
                                    Return GlobalFunctions.FormatBoolean(CType(_temp.FieldControl, DropDownList).SelectedValue)
                                Case GlobalFunctions.FIELDTYPES.FT_AUTOID, GlobalFunctions.FIELDTYPES.FT_LABEL, GlobalFunctions.FIELDTYPES.FT_DBLABEL
                                    Return CStr(CType(_temp.FieldControl, Label).Text)
                                Case GlobalFunctions.FIELDTYPES.FT_FILE
                                    Return CStr(CType(_temp.FieldControl, Object).Getinternalpath)
                                Case GlobalFunctions.FIELDTYPES.FT_HIDDENFIELD
                                    Return CStr(CType(_temp.FieldControl, HiddenField).Value)
                                Case GlobalFunctions.FIELDTYPES.FT_USER
                                    Return CStr(CType(_temp.FieldControl, Object).DisplayCaption)
                                Case GlobalFunctions.FIELDTYPES.FT_TABLE
                                    Return CType(_temp.FieldControl, Object).DataSource
                            End Select
                        End If
                        Return Nothing
                    End If
                Catch ex As Exception
                    DisplayMessage("Could not get data value for [" & _temp.FieldName & "] field<br>" & ex.ToString)
                End Try
            Next _counter
            Return Nothing

        End Get
        Set(ByVal value As Object)


            Dim _counter As Integer
            For _counter = 1 To _zfields.Count
                Dim _temp As ZField = _zfields.Item(_counter)
                If StrComp(FieldName, _temp.FieldName, CompareMethod.Text) = 0 Then
                    If _FormDetails.IsReadonlyMode = True Then
                        If _temp.FieldType = GlobalFunctions.FIELDTYPES.FT_DROPDOWN Then
                            CType(_temp.FieldControl, Label).Attributes("Value") = value
                            Dim _webObject As ZukamiLib.WebSession = Nothing
                            If KeepConnAlive = True Then
                                _webObject = GetSession()
                            Else
                                _webObject = New ZukamiLib.WebSession(GetZukamiSettings)
                                _webObject.OpenConnection()
                            End If


                            CType(_temp.FieldControl, Label).Text = GlobalFunctions.GetLookupCaption(_webObject, _temp.Arguments, value)
                            If KeepConnAlive = False Then
                                _webObject.CloseConnection()
                                _webObject = Nothing
                            End If
                        ElseIf _temp.FieldType = GlobalFunctions.FIELDTYPES.FT_HIDDENFIELD Then
                            CType(_temp.FieldControl, HiddenField).Value = value
                        Else
                            CType(_temp.FieldControl, Label).Text = value
                        End If
                    Else
                        Try
                            Select Case _temp.FieldType
                                Case GlobalFunctions.FIELDTYPES.FT_SHORTTEXT, GlobalFunctions.FIELDTYPES.FT_LONGTEXT
                                    CType(_temp.FieldControl, TextBox).Text = value
                                Case GlobalFunctions.FIELDTYPES.FT_COUNTRY
                                    CType(_temp.FieldControl, DropDownList).SelectedValue = value
                                Case GlobalFunctions.FIELDTYPES.FT_CURRENCY, GlobalFunctions.FIELDTYPES.FT_FLOAT, GlobalFunctions.FIELDTYPES.FT_INT
                                    CType(_temp.FieldControl, TextBox).Text = value
                                Case GlobalFunctions.FIELDTYPES.FT_CHECKLIST
                                    GlobalFunctions.SetCheckListValue(CType(_temp.FieldControl, Collection), value)
                                Case GlobalFunctions.FIELDTYPES.FT_RADIO
                                    GlobalFunctions.SetRadioListValue(CType(_temp.FieldControl, Collection), value)
                                Case GlobalFunctions.FIELDTYPES.FT_DATE
                                    CType(_temp.FieldControl, Object).value = GlobalFunctions.FormatDate(value)
                                Case GlobalFunctions.FIELDTYPES.FT_DATETIME
                                    CType(_temp.FieldControl, Object).value = GlobalFunctions.FormatDateTime(value)
                                Case GlobalFunctions.FIELDTYPES.FT_DROPDOWN
                                    CType(_temp.FieldControl, Object).Text = value
                                Case GlobalFunctions.FIELDTYPES.FT_YESNO
                                    CType(_temp.FieldControl, DropDownList).SelectedValue = value
                                Case GlobalFunctions.FIELDTYPES.FT_AUTOID, GlobalFunctions.FIELDTYPES.FT_LABEL, GlobalFunctions.FIELDTYPES.FT_DBLABEL
                                    CType(_temp.FieldControl, Label).Text = value
                                Case GlobalFunctions.FIELDTYPES.FT_FILE
                                    CType(_temp.FieldControl, Object).setinternalpath(value)
                                Case GlobalFunctions.FIELDTYPES.FT_HIDDENFIELD
                                    CType(_temp.FieldControl, HiddenField).Value = value
                                Case GlobalFunctions.FIELDTYPES.FT_HTML
                                    _temp.FieldControl.text = value
                                Case GlobalFunctions.FIELDTYPES.FT_USER
                                    CType(_temp.FieldControl, Object).Text = value
                                Case GlobalFunctions.FIELDTYPES.FT_TABLE
                                    System.Web.HttpContext.Current.Session(_temp.FieldGUID.ToString) = CType(value, DataTable).DataSet
                                    CType(_temp.FieldControl, Object).clearrows()
                                    CType(_temp.FieldControl, Object).DataSource = value
                            End Select
                        Catch ex As Exception
                            DisplayMessage("Could not set data value for [" & _temp.FieldName & "] field<br>" & ex.ToString)
                        End Try
                    End If
                    Exit For
                End If
            Next _counter
        End Set
    End Property

    Public ReadOnly Property Control(ByVal FieldName As String) As Object
        Get
            If _zfields.Contains(FieldName) = False Then Return Nothing
            Dim _temp As ZField = _zfields.Item(FieldName)
            Return _temp.FieldControl
        End Get
    End Property

    Public Sub LaunchWorkflows(ByVal FormID As String, ByVal InList As String, ByVal Remarks As String)
        Dim _Settings As ZukamiLib.ZukamiSettings = GetZukamiSettings()
        Dim webObj As ZukamiLib.WebSession = Nothing
        If KeepConnAlive = True Then
            webObj = GetSession()
        Else
            webObj = New ZukamiLib.WebSession(GetZukamiSettings)
            webObj.OpenConnection()
        End If

        Dim _formset As DataSet = webObj.forms_GetRecord(New Guid(FormID), Nothing)
        If _formset.Tables(0).Rows.Count > 0 Then
            Dim _wflowguid As Guid = GlobalFunctions.GetGUID(_formset.Tables(0).Rows(0).Item("WorkflowID"))
            If _wflowguid <> Guid.Empty Then

                Dim arrsplits() As String = Split(InList, ",")
                Dim _counter As Integer
                For _counter = 0 To UBound(arrsplits)
                    Dim _il As String = Trim(arrsplits(_counter))
                    _il = Replace(_il, "'", "")
                    If Len(_il) > 0 Then
                        If GlobalFunctions.IsGUID(_il) = True Then
                            Dim InitiateWorkflowSession As DataSet = webObj.InitiateFormWorkflowSession(_wflowguid, New Guid(FormID), New Guid(_il), Remarks)
                            Dim _instanceID As Guid = Guid.Empty
                            If InitiateWorkflowSession.Tables(0).Rows.Count > 0 Then
                                _instanceID = GlobalFunctions.GetGUID(InitiateWorkflowSession.Tables(0).Rows(0).Item("InstanceID"))

                                Dim oriItem As New MSMQMessage(_Settings.Queue)
                                oriItem.MessageType = MSMQMessage.MessageTypes.MT_NEWSESSION
                                oriItem.SessionGUID = GlobalFunctions.FormatData(_instanceID)
                                oriItem.SendMessage()
                            End If
                        End If
                    End If
                Next _counter

            End If
        End If
        If KeepConnAlive = False Then
            webObj.CloseConnection()
            webObj = Nothing
        End If

    End Sub

    Public Function GetWorkflowInstance(ByVal recordid As String) As DataSet
        Dim webObj As ZukamiLib.WebSession = Nothing
        If KeepConnAlive = True Then
            webObj = GetSession()
        Else
            webObj = New ZukamiLib.WebSession(GetZukamiSettings)
            webObj.OpenConnection()
        End If
        Dim _set As DataSet = webObj.WorkflowInstance_GetByRecordID(New Guid(recordid))
        If KeepConnAlive = False Then
            webObj.CloseConnection()
            webObj = Nothing
        End If

        Return _set
    End Function

    Public Sub TerminateWorkflow(ByVal recordid As String, Optional ByVal ClearTasks As Boolean = False)
        Dim _webObj As ZukamiLib.WebSession = Nothing
        If KeepConnAlive = True Then
            _webObj = GetSession()
        Else
            _webObj = New ZukamiLib.WebSession(GetZukamiSettings)
            _webObj.OpenConnection()
        End If
        Dim _set As DataSet = _webObj.WorkflowInstance_GetByRecordID(New Guid(recordid))



        If _set.Tables(0).Rows.Count > 0 Then
            Dim _iID As Guid = GlobalFunctions.GetGUID(_set.Tables(0).Rows(0).Item("InstanceID"))
            Dim _curQueue As String = GetZukamiSettings.Queue
            Dim _msgItem As New MSMQMessage(_curQueue)
            _msgItem.MessageType = MSMQMessage.MessageTypes.MT_TERMSESSION
            _msgItem.SessionGUID = _iID.ToString
            _msgItem.SendMessage()


            If ClearTasks = True Then
                _webObj.CustomSQLCommand("Update Tasks set [Status]='Cancelled',[DateCompleted]=GETDATE() where InstanceID='" & _iID.ToString & "'")
                _webObj.CustomClearParameters()
                _webObj.CustomSQLExecute()
            End If

        End If



        If KeepConnAlive = False Then
            _webObj.CloseConnection()
            _webObj = Nothing
        End If


    End Sub

    Public Function LaunchWorkflow(ByVal FormID As String, ByVal RecordID As String, ByVal Remarks As String, Optional ByVal TargetUserID As String = "", Optional ByVal TargetUserFullName As String = "") As Guid
        LaunchWorkflow = Guid.Empty
        Dim _Settings As ZukamiLib.ZukamiSettings = GetZukamiSettings()
        Dim webObj As ZukamiLib.WebSession = Nothing
        If KeepConnAlive = True Then
            webObj = GetSession()
        Else
            webObj = New ZukamiLib.WebSession(_Settings)
            webObj.OpenConnection()
        End If
        Dim _formset As DataSet = webObj.forms_GetRecord(New Guid(FormID), Nothing)
        If _formset.Tables(0).Rows.Count > 0 Then
            Dim _wflowguid As Guid = GlobalFunctions.GetGUID(_formset.Tables(0).Rows(0).Item("WorkflowID"))
            If _wflowguid <> Guid.Empty Then
                Dim InitiateWorkflowSession As DataSet = webObj.InitiateFormWorkflowSession(_wflowguid, New Guid(FormID), New Guid(RecordID), Remarks, TargetUserID, TargetUserFullName)
                Dim _instanceID As Guid = Guid.Empty
                If InitiateWorkflowSession.Tables(0).Rows.Count > 0 Then
                    _instanceID = GlobalFunctions.GetGUID(InitiateWorkflowSession.Tables(0).Rows(0).Item("InstanceID"))

                    Dim oriItem As New MSMQMessage(_Settings.Queue)
                    oriItem.MessageType = MSMQMessage.MessageTypes.MT_NEWSESSION
                    oriItem.SessionGUID = GlobalFunctions.FormatData(_instanceID)
                    oriItem.SendMessage()
                    LaunchWorkflow = _instanceID
                End If
            End If
        End If
        If KeepConnAlive = False Then
            webObj.CloseConnection()
            webObj = Nothing
        End If

    End Function

    Public Property RowVisible(ByVal FieldName As String) As Boolean
        Get
            If Len(FieldName) = 0 Then Return False
            If _zfields.Contains(FieldName) = False Then Return False
            Dim _temp As ZField = _zfields.Item(FieldName)

            If GlobalFunctions.IsCurrentRequestMobileVersion() Then
                Return GetRowVisibleForMobileFields(_temp)
            Else
                If _temp.FieldType = GlobalFunctions.FIELDTYPES.FT_HEADER Then
                    Return _temp.FieldControl.parent.parent.visible
                ElseIf _temp.FieldType = GlobalFunctions.FIELDTYPES.FT_CHECKLIST Or _temp.FieldType = GlobalFunctions.FIELDTYPES.FT_RADIO Then
                    Dim _coll As Collection = _temp.FieldControl
                    If _coll.Count > 0 Then
                        Dim _chklist As CheckBox = _coll.Item(1)
                        Return _chklist.Parent.Parent.Visible
                    End If
                Else
                    Return _temp.FieldControl.parent.parent.visible
                End If
            End If
        End Get
        Set(ByVal value As Boolean)
            If Len(FieldName) = 0 Then Return
            If _zfields.Contains(FieldName) = False Then Return
            Dim _temp As ZField = _zfields.Item(FieldName)

            Try
                If GlobalFunctions.IsCurrentRequestMobileVersion() Then
                    SetRowVisibleForMobileFields(_temp, value)
                Else
                    If _temp.FieldType = GlobalFunctions.FIELDTYPES.FT_HEADER Then
                        _temp.FieldControl.parent.parent.visible = value
                        _temp.FieldControl3.visible = value
                        _temp.FieldControl4.visible = value
                    ElseIf _temp.FieldType = GlobalFunctions.FIELDTYPES.FT_CHECKLIST Or _temp.FieldType = GlobalFunctions.FIELDTYPES.FT_RADIO Then
                        Dim _coll As Collection = _temp.FieldControl
                        If _coll.Count > 0 Then
                            Dim _chklist As CheckBox = _coll.Item(1)
                            _chklist.Parent.Parent.Visible = value
                        End If
                        Try
                            If _temp.ParentMobileLabel Is Nothing = False Then
                                _temp.ParentMobileLabel.Visible = value
                            End If
                        Catch ex As Exception
                            logger.Error(ex)
                        End Try
                    Else
                        _temp.FieldControl.parent.parent.visible = value
                        Try
                            If _temp.ParentMobileLabel Is Nothing = False Then
                                _temp.ParentMobileLabel.Visible = value
                            End If
                        Catch ex As Exception
                            logger.Error(ex)
                        End Try

                    End If

                End If
            Catch ex As Exception
                logger.Error(ex)
            End Try
        End Set
    End Property

    Private Function GetRowVisibleForMobileFields(field As ZField) As Boolean
        Dim result As Boolean = True
        If field.ParentRow IsNot Nothing Then
            Try
                result = CType(field.ParentRow, HtmlGenericControl).Visible
            Catch ex As Exception
                logger.Error(ex, "field name: " + field.FieldName)
            End Try
        Else
            logger.Error("cannot find parent row for field: " + field.FieldName)
        End If
        Return result
    End Function

    Private Sub SetRowVisibleForMobileFields(field As ZField, visible As Boolean)
        If field.ParentRow IsNot Nothing Then
            Try
                CType(field.ParentRow, HtmlGenericControl).Visible = visible
            Catch ex As Exception
                logger.Error(ex, "field name: " + field.FieldName)
            End Try
        Else
            logger.Error("cannot find parent row for field: " + field.FieldName)
        End If
    End Sub

    Public Property SectionVisible(ByVal FieldName As String) As Boolean
        Get
            Try
                If Len(FieldName) = 0 Then Return False
                If _zfields.Contains(FieldName) = False Then Return False
                Dim _temp As ZField = _zfields.Item(FieldName)
                If GlobalFunctions.IsCurrentRequestMobileVersion() Then
                    Return RowVisible(FieldName)
                Else
                    If _temp.FieldType = GlobalFunctions.FIELDTYPES.FT_HEADER Then
                        Return _temp.FieldControl.visible
                    End If
                End If
            Catch ex As Exception
                logger.Error(ex, "FieldName: " + FieldName)
            End Try
        End Get
        Set(ByVal value As Boolean)
            Dim _counter As Integer
            For _counter = 1 To _zfields.Count
                Dim _temp As ZField = _zfields.Item(_counter)
                If StrComp(FieldName, _temp.FieldName, CompareMethod.Text) = 0 Then
                    Try
                        If _temp.FieldType = GlobalFunctions.FIELDTYPES.FT_HEADER Then

                            If GlobalFunctions.IsCurrentRequestMobileVersion() Then
                                RowVisible(FieldName) = value
                            Else
                                _temp.FieldControl.parent.parent.visible = value
                                _temp.FieldControl3.visible = value
                                _temp.FieldControl4.visible = value
                            End If

                            ''we must now set the visibility of all the other rows beneath it
                            Dim _counter2 As Integer = _counter + 1
                            Do While _counter2 <= _zfields.Count
                                Dim _field As ZField = _zfields.Item(_counter2)
                                If _field.FieldType = GlobalFunctions.FIELDTYPES.FT_HEADER Then
                                    Exit Do
                                Else
                                    If _field.FieldType <> GlobalFunctions.FIELDTYPES.FT_HIDDENFIELD Then
                                        RowVisible(_field.FieldName) = value
                                    End If
                                End If
                                _counter2 += 1
                            Loop
                        End If
                    Catch ex As Exception
                        logger.Error(ex, "field name: " + FieldName)
                    End Try
                    Exit For
                End If
            Next _counter
        End Set
    End Property

    Public Property TabVisible(ByVal FieldName As String) As Boolean
        Get
            Dim _tabvisible As String = Me.Session("TV_" & CurrentUser.UserID.ToString & "_" & PrimaryID.ToString & "_" & FieldName)
            If Len(_tabvisible) = 0 Then
                Return True
            Else
                Return GlobalFunctions.FormatBoolean(_tabvisible)
            End If
        End Get
        Set(ByVal value As Boolean)
            Me.Session("TV_" & CurrentUser.UserID.ToString & "_" & PrimaryID.ToString & "_" & FieldName) = IIf(value = True, "true", "false")
            Tabify()
        End Set
    End Property

    Public Property SectionExpand(ByVal FieldName As String) As Boolean
        Get
            Dim _counter As Integer
            For _counter = 1 To _zfields.Count
                Dim _temp As ZField = _zfields.Item(_counter)
                If StrComp(FieldName, _temp.FieldName, CompareMethod.Text) = 0 Then
                    If _temp.FieldType = GlobalFunctions.FIELDTYPES.FT_HEADER Then
                        Return _temp.FieldControl.visible
                    End If
                End If
            Next _counter
            Return False
        End Get
        Set(ByVal value As Boolean)
            Dim _counter As Integer
            For _counter = 1 To _zfields.Count
                Dim _temp As ZField = _zfields.Item(_counter)
                If StrComp(FieldName, _temp.FieldName, CompareMethod.Text) = 0 Then
                    Try
                        If _temp.FieldType = GlobalFunctions.FIELDTYPES.FT_HEADER Then

                            ''we must now set the visibility of all the other rows beneath it
                            Dim _counter2 As Integer = _counter + 1
                            Do While _counter2 <= _zfields.Count
                                Dim _field As ZField = _zfields.Item(_counter2)
                                If _field.FieldType = GlobalFunctions.FIELDTYPES.FT_HEADER Then
                                    Exit Do
                                Else
                                    If _field.FieldType <> GlobalFunctions.FIELDTYPES.FT_HIDDENFIELD Then
                                        RowVisible(_field.FieldName) = value
                                    End If
                                End If
                                _counter2 += 1
                            Loop
                        End If
                    Catch ex As Exception

                    End Try
                    Exit For
                End If
            Next _counter
        End Set
    End Property

    Public WriteOnly Property SectionReadOnly(ByVal FieldName As String) As Boolean
        Set(ByVal value As Boolean)
            Dim _counter As Integer
            For _counter = 1 To _zfields.Count
                Dim _temp As ZField = _zfields.Item(_counter)
                If StrComp(FieldName, _temp.FieldName, CompareMethod.Text) = 0 Then
                    Try
                        If _temp.FieldType = GlobalFunctions.FIELDTYPES.FT_HEADER Then

                            ''we must now set the visibility of all the other rows beneath it
                            Dim _counter2 As Integer = _counter + 1
                            Do While _counter2 <= _zfields.Count
                                Dim _field As ZField = _zfields.Item(_counter2)
                                If _field.FieldType = GlobalFunctions.FIELDTYPES.FT_HEADER Then
                                    Exit Do
                                Else
                                    If _field.FieldType <> GlobalFunctions.FIELDTYPES.FT_HIDDENFIELD Then
                                        IsReadOnly(_field.FieldName) = value
                                    End If
                                End If
                                _counter2 += 1
                            Loop
                        End If
                    Catch ex As Exception

                    End Try
                    Exit For
                End If
            Next _counter
        End Set
    End Property


    Public Function User(ByVal UserID As Guid) As Object
        Return Nothing
    End Function

    Public Property Zfields() As Collection
        Get
            Return _zfields
        End Get
        Set(ByVal value As Collection)
            _zfields = value
        End Set
    End Property

    Public ReadOnly Property CurrentAppID() As Guid
        Get
            If _FormAttributes Is Nothing = False Then
                Return _FormAttributes.AppID
            Else
                Return Guid.Empty
            End If
        End Get
    End Property

    Public Property CurrentForm() As ScribeFormDetails
        Get
            Return _FormDetails
        End Get
        Set(ByVal value As ScribeFormDetails)
            _FormDetails = value
        End Set
    End Property

    Private Function GetZField(ByVal FieldName As String) As ZField
        If _zfields.Contains(FieldName) = False Then Return Nothing
        Return _zfields.Item(FieldName)
    End Function

    Public Function Rawdata(ByVal FieldName As String) As Object
        Try
            If _DataBag Is Nothing = False Then
                If _DataBag.Tables(0).Rows.Count > 0 Then
                    Return _DataBag.Tables(0).Rows(0).Item(FieldName)
                End If
            End If
            Return Nothing
        Catch ex As Exception
            ValidationError.Display("Could not retrieve raw data [" & FieldName & "]:<br>" & GlobalFunctions.MultiTextToHTML(ex.ToString))
            Return Nothing
        End Try

    End Function

    Public Sub AddTableDataBag(ByRef Datasource As DataSet, ByVal fieldName As String)
        If _TableDataBags Is Nothing Then _TableDataBags = New Collection
        If _TableDataBags.Contains(fieldName) = False Then
            _TableDataBags.Add(Datasource, fieldName)
        End If
    End Sub

    Public Function GetSubFormDataset(ByVal FieldName As String) As DataSet
        Try
            If _TableDataBags.Contains(FieldName) = True Then
                Return _TableDataBags.Item(FieldName)
            Else
                Return Nothing
            End If
            Return Nothing
        Catch ex As Exception
            ValidationError.Display("Could not retrieve subform raw data [" & FieldName & "]:<br>" & GlobalFunctions.MultiTextToHTML(ex.ToString))
            Return Nothing
        End Try

    End Function


    Public Function SessionDataExists() As Boolean
        Return Not (System.Web.HttpContext.Current.Session(GetParentToken) Is Nothing)
    End Function

    Public Sub DisplayMessage(ByVal Message As String)
        Try
            If Message Is Nothing = False Then
                ValidationError.Display(Message)
            End If
        Catch ex As Exception
            ValidationError.Display("Could not display message: " & GlobalFunctions.MultiTextToHTML(ex.ToString))
        End Try

    End Sub

    Public Sub CancelWorkflowInstances(ByVal RecordIDs As String)
        If Len(Trim(RecordIDs)) = 0 Then Exit Sub
        MassAction(RecordIDs, MSMQMessage.MessageTypes.MT_TERMSESSION)
    End Sub
    Public Function IsWorkflowCompleted(ByVal RecordID As String) As Boolean

        Dim _settings As ZukamiLib.ZukamiSettings = GetZukamiSettings()
        Dim _web As New ZukamiLib.WebSession(_settings)
        _web.OpenConnection()

        _web.CustomSQLCommand("Select * from WorkflowInstances Where  status = 1 and InstanceCompleted is not null and  [RecordID]='" & RecordID & "'")
        Dim _mw As DataSet = _web.CustomSQLExecuteReturn
        If _mw.Tables(0).Rows.Count > 0 Then
            _web.CloseConnection()
            Return True
        End If
        _web.CloseConnection()
        Return False
    End Function
    Public Sub RestartWorkflowInstances(ByVal RecordIDs As String)

        Dim _settings As ZukamiLib.ZukamiSettings = GetZukamiSettings()
        Dim _web As New ZukamiLib.WebSession(_settings)
        _web.OpenConnection()

        _web.CustomSQLCommand("select InstanceID,WorkflowID from [Workflowinstances] WHERE [RecordID]=" & RecordIDs)
        Dim _set As DataSet = _web.CustomSQLExecuteReturn
        If _set.Tables(0).Rows.Count > 0 Then
            Dim wflowID As String = GlobalFunctions.FormatData(_set.Tables(0).Rows(0).Item("WorkflowID"))
            Dim instanceID As String = GlobalFunctions.FormatData(_set.Tables(0).Rows(0).Item("InstanceID"))


            _web.CustomSQLCommand("select top 1 * from workflowversion where masterworkflowid= (select masterworkflowid from workflowversion where id='" + wflowID + "') Order by Version Desc")
            Dim _mw As DataSet = _web.CustomSQLExecuteReturn
            If _mw.Tables(0).Rows.Count > 0 Then
                Dim _wversion As String = GlobalFunctions.FormatData(_mw.Tables(0).Rows(0).Item("ID"))


                Dim _guid As Guid = Guid.NewGuid
                _web.CustomSQLCommand("UPDATE WorkflowInstances SET WorkflowID='" + _wversion + "',InstanceID='" + _guid.ToString + "',InstanceCompleted=NULL,InstanceStarted=GETDATE(),Status=0 WHERE InstanceID='" + instanceID + "'")
                _web.CustomSQLExecute()

                '_web.CustomSQLCommand("UPDATE InstanceAttachments SET InstanceID='" + _guid.ToString + "' WHERE InstanceID='" + Request.QueryString("ID") + "'")
                '_web.CustomSQLExecute()

                Dim oriItem As New MSMQMessage(_settings.Queue)
                oriItem.MessageType = MSMQMessage.MessageTypes.MT_NEWSESSION
                oriItem.SessionGUID = _guid.ToString
                oriItem.SendMessage()


                _web.CloseConnection()

            End If


        End If

    End Sub

    Public Sub MassApprove(ByVal RecordIDs As String, ByVal NodeToApprove As String, ByVal Action As String, ByVal CustomActionDisplayName As String, ByVal Remarks As String)
        Try

            Dim _webObj As ZukamiLib.WebSession = Nothing
            If KeepConnAlive = True Then
                _webObj = GetSession()
            Else
                _webObj = New ZukamiLib.WebSession(GetZukamiSettings)
                _webObj.OpenConnection()
            End If
            _webObj.CustomSQLCommand("SELECT a.TaskID FROM Tasks a INNER JOIN WorkflowInstances b ON a.InstanceID=b.InstanceID INNER JOIN InstanceBubbles c ON a.InstanceNodeID=c.ID WHERE a.RecipientID='" & GetZukamiSettings.CurrentUserGUID.ToString & "' AND b.RecordID IN (" & RecordIDs & ") AND a.Status='' AND c.NodeID LIKE '%" & GlobalFunctions.FormatSQLData(NodeToApprove) & "%'")


            _webObj.CustomClearParameters()
            Dim _set As DataSet = _webObj.CustomSQLExecuteReturn
            If KeepConnAlive = False Then
                _webObj.CloseConnection()
                _webObj = Nothing
            End If


            Dim _counter As Integer
            For _counter = 0 To _set.Tables(0).Rows.Count - 1
                Dim _taskID As Guid = GlobalFunctions.GetGUID(_set.Tables(0).Rows(_counter).Item("TaskID"))



                GlobalFunctions.PerformAction(_taskID, Action, CustomActionDisplayName, Remarks)
            Next _counter

        Catch ex As Exception
            ValidationError.Display(ex.ToString)

        End Try
    End Sub


    Public Sub MassApprove(ByVal RecordIDs As String, ByVal NodeToApprove As String, ByVal Action As String, ByVal CustomActionDisplayName As String, ByVal Remarks As String, PerformBy As String)
        Try

            Dim _webObj As ZukamiLib.WebSession = Nothing
            If KeepConnAlive = True Then
                _webObj = GetSession()
            Else
                _webObj = New ZukamiLib.WebSession(GetZukamiSettings)
                _webObj.OpenConnection()
            End If
            _webObj.CustomSQLCommand("SELECT a.TaskID FROM Tasks a INNER JOIN WorkflowInstances b ON a.InstanceID=b.InstanceID INNER JOIN InstanceBubbles c ON a.InstanceNodeID=c.ID WHERE a.RecipientID='" & PerformBy & "' AND b.RecordID IN (" & RecordIDs & ") AND a.Status='' AND c.NodeID LIKE '%" & GlobalFunctions.FormatSQLData(NodeToApprove) & "%'")


            _webObj.CustomClearParameters()
            Dim _set As DataSet = _webObj.CustomSQLExecuteReturn
            If KeepConnAlive = False Then
                _webObj.CloseConnection()
                _webObj = Nothing
            End If


            Dim _counter As Integer
            For _counter = 0 To _set.Tables(0).Rows.Count - 1
                Dim _taskID As Guid = GlobalFunctions.GetGUID(_set.Tables(0).Rows(_counter).Item("TaskID"))



                GlobalFunctions.PerformAction(_taskID, Action, CustomActionDisplayName, Remarks)
            Next _counter

        Catch ex As Exception
            ValidationError.Display(ex.ToString)

        End Try
    End Sub

    Private Sub MassAction(ByVal RecordIDs As String, ByVal ActionType As MSMQMessage.MessageTypes)
        Try
            Dim _webObj As ZukamiLib.WebSession = Nothing
            If KeepConnAlive = True Then
                _webObj = GetSession()
            Else
                _webObj = New ZukamiLib.WebSession(GetZukamiSettings)
                _webObj.OpenConnection()
            End If

            _webObj.CustomSQLCommand("SELECT InstanceID FROM WorkflowInstances WHERE RecordID IN (" & RecordIDs & ")")
            _webObj.CustomClearParameters()
            Dim _set As DataSet = _webObj.CustomSQLExecuteReturn
            If KeepConnAlive = False Then
                _webObj.CloseConnection()
                _webObj = Nothing
            End If


            Dim _counter As Integer
            For _counter = 0 To _set.Tables(0).Rows.Count - 1
                Dim _insID As String = GlobalFunctions.FormatData(_set.Tables(0).Rows(_counter).Item("InstanceID"))
                Dim _msgItem As New MSMQMessage(GetZukamiSettings.Queue)
                _msgItem.MessageType = ActionType
                _msgItem.SessionGUID = _insID
                _msgItem.SendMessage()
            Next _counter

        Catch ex As Exception
            ValidationError.Display(ex.ToString)
        End Try
    End Sub


    Public Function GetRoleUsers(ByVal RoleName As String) As Collection
        '=============================================================
        'Open database connection
        '=============================================================
        Dim _webObject As ZukamiLib.WebSession = Nothing
        If KeepConnAlive = True Then
            _webObject = GetSession()
        Else
            _webObject = New ZukamiLib.WebSession(GetZukamiSettings)
            _webObject.OpenConnection()
        End If

        GetRoleUsers = _webObject.GroupUsers_GetByName(RoleName)
        '=============================================================
        'Close database connection
        '=============================================================
        If KeepConnAlive = False Then
            _webObject.CloseConnection()
            _webObject = Nothing
        End If
    End Function

    Public Function GetRoleUsers(ByVal RoleID As Guid) As Collection
        '=============================================================
        'Open database connection
        '=============================================================
        Dim _newcoll As New Collection
        Dim _webObject As ZukamiLib.WebSession = Nothing
        If KeepConnAlive = True Then
            _webObject = GetSession()
        Else
            _webObject = New ZukamiLib.WebSession(GetZukamiSettings)
            _webObject.OpenConnection()
        End If

        Dim _set As DataSet = _webObject.GroupUsers_Get(RoleID)
        Dim _counter As Integer
        For _counter = 0 To _set.Tables(0).Rows.Count - 1
            Dim _userID As String = GlobalFunctions.FormatData(_set.Tables(0).Rows(_counter).Item("UserID"))
            _newcoll.Add(_userID)
        Next _counter
        '=============================================================
        'Close database connection
        '=============================================================
        If KeepConnAlive = False Then
            _webObject.CloseConnection()
            _webObject = Nothing
        End If

        Return _newcoll
    End Function


    Public Function GetWorkflowParticipantEmails(ByVal WorkflowInstanceID As Guid) As String
        '=============================================================
        'Open database connection
        '=============================================================
        Dim _webObject As ZukamiLib.WebSession = Nothing
        If KeepConnAlive = True Then
            _webObject = GetSession()
        Else
            _webObject = New ZukamiLib.WebSession(GetZukamiSettings)
            _webObject.OpenConnection()
        End If
        Dim _counter As Integer
        Dim _subs As String = ""

        Dim _dataset As DataSet = _webObject.NotificationSubscribers_Get(WorkflowInstanceID)
        For _counter = 0 To _dataset.Tables(0).Rows.Count - 1
            Dim strEmail As String = Trim(GlobalFunctions.FormatData(_dataset.Tables(0).Rows(_counter).Item("Email")))
            If Len(strEmail) > 0 Then
                If Len(_subs) > 0 Then _subs += ","
                _subs += strEmail
            End If
        Next _counter

        '=============================================================
        'Close database connection
        '=============================================================
        _dataset.Dispose()
        _dataset = Nothing
        If KeepConnAlive = False Then
            _webObject.CloseConnection()
            _webObject = Nothing
        End If

        Return _subs
    End Function

    Public ReadOnly Property SystemSettings() As ZukamiLib.ZukamiSettings
        Get
            Return GetZukamiSettings()
        End Get
    End Property

    Public Function GetUserProperty(ByVal UserID As String, ByVal PropertyName As String) As Object

        '=============================================================
        'Open database connection
        '=============================================================
        Dim _webObject As ZukamiLib.WebSession = Nothing
        If KeepConnAlive = True Then
            _webObject = GetSession()
        Else
            _webObject = New ZukamiLib.WebSession(GetZukamiSettings)
            _webObject.OpenConnection()
        End If

        GetUserProperty = _webObject.GetUserProperty(UserID, PropertyName)

        '=============================================================
        'Close database connection
        '=============================================================
        If KeepConnAlive = False Then
            _webObject.CloseConnection()
            _webObject = Nothing
        End If
    End Function

    Public Function UserID2Email(ByVal UserIDs As Collection) As String
        If UserIDs.Count = 0 Then Return ""
        Dim _webObject As ZukamiLib.WebSession = Nothing
        If KeepConnAlive = True Then
            _webObject = GetSession()
        Else
            _webObject = New ZukamiLib.WebSession(GetZukamiSettings)
            _webObject.OpenConnection()
        End If
        Dim _Emails As String = ""
        Dim _UserIDs As String = ""
        Dim _counter As Integer


        For _counter = 1 To UserIDs.Count
            Dim _guid As Guid = UserIDs.Item(_counter)
            If Len(_UserIDs) > 0 Then _UserIDs += ","
            _UserIDs += "'" + _guid.ToString + "'"
        Next _counter

        _webObject.CustomSQLCommand("SELECT DISTINCT Email FROM Users WHERE UserID IN (" + _UserIDs + ")")
        Dim _dataset As DataSet = _webObject.CustomSQLExecuteReturn
        For _counter = 0 To _dataset.Tables(0).Rows.Count - 1
            Dim strEmail As String = Trim(GlobalFunctions.FormatData(_dataset.Tables(0).Rows(_counter).Item("Email")))
            If Len(strEmail) > 0 Then
                If Len(_Emails) > 0 Then _Emails += ","
                _Emails += strEmail
            End If
        Next _counter
        _dataset.Dispose()
        _dataset = Nothing
        If KeepConnAlive = False Then
            _webObject.CloseConnection()
            _webObject = Nothing
        End If
        Return _Emails
    End Function

    Public Function GetUserProperty(ByVal UserID As Guid, ByVal PropertyName As String) As Object

        '=============================================================
        'Open database connection
        '=============================================================
        Dim _webObject As ZukamiLib.WebSession = Nothing
        If KeepConnAlive = True Then
            _webObject = GetSession()
        Else
            _webObject = New ZukamiLib.WebSession(GetZukamiSettings)
            _webObject.OpenConnection()
        End If


        GetUserProperty = _webObject.GetUserProperty(UserID, PropertyName)

        '=============================================================
        'Close database connection
        '=============================================================
        If KeepConnAlive = False Then
            _webObject.CloseConnection()
            _webObject = Nothing
        End If
    End Function

    Public Sub New()

    End Sub

    Public Property GlobalProperty(ByVal PropertyName As String, ByVal Scope As Integer) As Object
        Get
            Dim _webObject As ZukamiLib.WebSession = Nothing
            If KeepConnAlive = True Then
                _webObject = GetSession()
            Else
                _webObject = New ZukamiLib.WebSession(GetZukamiSettings)
                _webObject.OpenConnection()
            End If


            If Scope = ZukamiLib.WebSession.CUSTOMVAR_SCOPE.SCOPE_GLOBAL Then
                GlobalProperty = _webObject.CustomVariables_Get(PropertyName, 0, Guid.Empty)
            Else
                GlobalProperty = Nothing
            End If


            '=============================================================
            'Close database connection
            '=============================================================
            If KeepConnAlive = False Then
                _webObject.CloseConnection()
                _webObject = Nothing
            End If
        End Get
        Set(ByVal value As Object)

            Dim _webObject As ZukamiLib.WebSession = Nothing
            If KeepConnAlive = True Then
                _webObject = GetSession()
            Else
                _webObject = New ZukamiLib.WebSession(GetZukamiSettings)
                _webObject.OpenConnection()
            End If


            If Scope = ZukamiLib.WebSession.CUSTOMVAR_SCOPE.SCOPE_GLOBAL Then
                _webObject.CustomVariables_Set(PropertyName, 0, Guid.Empty, value)
            End If

            '=============================================================
            'Close database connection
            '=============================================================
            If KeepConnAlive = False Then
                _webObject.CloseConnection()
                _webObject = Nothing
            End If
        End Set
    End Property

    Public Function UserHasRole(ByVal UserID As Guid, ByVal RoleID As Guid) As Boolean
        '=============================================================
        'Open database connection
        '=============================================================
        Dim _webObject As ZukamiLib.WebSession = Nothing
        If KeepConnAlive = True Then
            _webObject = GetSession()
        Else
            _webObject = New ZukamiLib.WebSession(GetZukamiSettings)
            _webObject.OpenConnection()
        End If


        UserHasRole = _webObject.UserInGroup2_Get(RoleID, UserID)

        '=============================================================
        'Close database connection
        '=============================================================
        If KeepConnAlive = False Then
            _webObject.CloseConnection()
            _webObject = Nothing
        End If
    End Function

    Public Function UserHasRole(ByVal UserID As Guid, ByVal RoleName As String) As Boolean
        '=============================================================
        'Open database connection
        '=============================================================
        Dim _webObject As ZukamiLib.WebSession = Nothing
        If KeepConnAlive = True Then
            _webObject = GetSession()
        Else
            _webObject = New ZukamiLib.WebSession(GetZukamiSettings)
            _webObject.OpenConnection()
        End If


        UserHasRole = _webObject.UserHasRole_Get(UserID, RoleName)

        '=============================================================
        'Close database connection
        '=============================================================
        If KeepConnAlive = False Then
            _webObject.CloseConnection()
            _webObject = Nothing
        End If
    End Function

    Public Function HeadOfDepartment(ByVal Submitter As Guid, ByVal Level As Integer) As Guid
        '=============================================================
        'Open database connection
        '=============================================================
        Dim _webObject As ZukamiLib.WebSession = Nothing
        If KeepConnAlive = True Then
            _webObject = GetSession()
        Else
            _webObject = New ZukamiLib.WebSession(GetZukamiSettings)
            _webObject.OpenConnection()
        End If


        HeadOfDepartment = _webObject.HeadOfDepartment_Get(Submitter, Level)

        '=============================================================
        'Close database connection
        '=============================================================
        If KeepConnAlive = False Then
            _webObject.CloseConnection()
            _webObject = Nothing
        End If
    End Function

    Public Function ValidatePassword(ByVal Username As String, ByVal Password As String, ByRef ValidationError As String) As Boolean
        '=============================================================
        'Open database connection
        '=============================================================
        Dim _webObject As ZukamiLib.WebSession = Nothing
        If KeepConnAlive = True Then
            _webObject = GetSession()
        Else
            _webObject = New ZukamiLib.WebSession(GetZukamiSettings)
            _webObject.OpenConnection()
        End If


        ValidatePassword = GlobalFunctions.ValidatePassword(_webObject, Username, Password, ValidationError)



        '=============================================================
        'Close database connection
        '=============================================================
        If KeepConnAlive = False Then
            _webObject.CloseConnection()
            _webObject = Nothing
        End If
    End Function

    Public Function Supervisor(ByVal Submitter As Guid) As Guid
        Return GlobalFunctions.Supervisor(Submitter)
    End Function

    Public Sub Dispose()
        _fieldList.Dispose()
        _fieldList = Nothing
    End Sub

    Public Function AESEncrypt(ByVal data As String) As String
        Return GlobalFunctions.AESEncrypt(data)
    End Function
    Public Function AESDecrypt(ByVal data As String) As String
        Return GlobalFunctions.AESDecrypt(data)
    End Function
    Public Function GetUploadedFilepath(ByVal FileTag As String) As String
        Return GlobalFunctions.GetUploadedFilePath(FileTag)


    End Function

    Public Function GetFullUploadedFilepath(ByVal FileTag As String) As String
        Return GlobalFunctions.GetFullUploadedFilePath(FileTag)
    End Function

    'Private Function GetBoundField(ByVal Caption As String, ByRef Type As ZukamiLib.BaseNode.GLOBFIELDTYPES) As String
    '    Dim _counter As Integer
    '    For _counter = 0 To _fieldList.Tables(0).Rows.Count - 1
    '        Dim _caption As String = GlobalFunctions.FormatData(_fieldList.Tables(0).Rows(_counter).Item("FieldCaption"))
    '        If StrComp(_caption, Caption, CompareMethod.Text) = 0 Then
    '            Type = GlobalFunctions.FormatInteger(_fieldList.Tables(0).Rows(_counter).Item("FieldType"))
    '            Return GlobalFunctions.FormatData(_fieldList.Tables(0).Rows(_counter).Item("FieldBindSource"))
    '        End If
    '    Next _counter
    '    Return ""
    'End Function

    Private Function GetFormattedData(ByRef Data As Object, ByVal Type As ZukamiLib.BaseNode.GLOBFIELDTYPES) As Object
        Select Case Type
            Case ZukamiLib.BaseNode.GLOBFIELDTYPES.FT_INT
                Return GlobalFunctions.FormatInteger(Data)
            Case ZukamiLib.BaseNode.GLOBFIELDTYPES.FT_FLOAT
                Return GlobalFunctions.FormatDouble(Data)
            Case ZukamiLib.BaseNode.GLOBFIELDTYPES.FT_SHORTTEXT, ZukamiLib.BaseNode.GLOBFIELDTYPES.FT_LONGTEXT, ZukamiLib.BaseNode.GLOBFIELDTYPES.FT_DROPDOWN, ZukamiLib.BaseNode.GLOBFIELDTYPES.FT_USER
                Return GlobalFunctions.FormatData(Data)
            Case ZukamiLib.BaseNode.GLOBFIELDTYPES.FT_DATE, ZukamiLib.BaseNode.GLOBFIELDTYPES.FT_DATETIME
                Return GlobalFunctions.FormatDateTime(Data)
            Case ZukamiLib.BaseNode.GLOBFIELDTYPES.FT_YESNO
                Return GlobalFunctions.FormatBoolean(Data)
            Case ZukamiLib.BaseNode.GLOBFIELDTYPES.FT_FILE
                Return Replace(GlobalFunctions.FormatData(Data), ";", "\")
            Case Else
                Return GlobalFunctions.FormatData(Data)
        End Select
    End Function

    Public Function SimpleCrypt(ByVal Text As String) As String
        Return GlobalFunctions.SimpleCrypt(Text)
    End Function
    Public Function EncryptBase64(Input As String) As String
        Return GlobalFunctions.EncryptBase64(Input)
    End Function
    Public Function DecryptBase64(Input As String) As String
        Return GlobalFunctions.DecryptBase64(Input)
    End Function

    Public Sub DataTableToExcel(dt As DataTable, filePath As String)
        GlobalFunctions.DataTableToExcel(dt, filePath)
    End Sub
    Public Function RunDynamicSQLStoreProc(ByVal AppID As Guid, ByVal Datasource As String, ByVal SQL As String, Optional ByVal Params As Collection = Nothing) As String
        Dim _sql As String = SQL
        Try
            Dim webobj As ZukamiLib.WebSession = Nothing
            RunDynamicSQLStoreProc = ""

            Dim _tempweb As New ZukamiLib.WebSession(GetZukamiSettings)
            _tempweb.OpenConnection()
            Dim _zuksettings As ZukamiLib.ZukamiSettings = GlobalFunctions.GetDatasourceConnectionString(_tempweb, AppID, Datasource)
            _tempweb.CloseConnection()
            _tempweb = Nothing

            webobj = New ZukamiLib.WebSession(_zuksettings)
            webobj.OpenOLEDBConnection(_zuksettings.PrimaryConnectionString)
            If Len(webobj.LastError) > 0 Then
                Return webobj.LastError
            End If

            Dim _counter As Integer
            If _zfields Is Nothing = False Then
                For _counter = 1 To _zfields.Count
                    Dim _temp As ZField = _zfields.Item(_counter)
                    Dim _fieldname As String = "[$" & _temp.FieldName & "]"

                    Dim obj As Object = Data(_temp.FieldName)
                    If Not TypeOf obj Is DataTable Then
                        Try
                            _sql = Replace(_sql, _fieldname, GlobalFunctions.FormatSQLData(obj), , , CompareMethod.Text)
                        Catch ex As Exception

                        End Try

                    End If
                Next _counter
            End If
            _sql = Replace(_sql, "[$ID]", PrimaryID, , , CompareMethod.Text)

            webobj.CustomOLEDBSQLStoreProc(_sql)
            webobj.CustomOLEDBClearParameters()
            'Arby Add paramaters 
            If Not Params Is Nothing Then
                Dim i As Integer
                For i = 1 To Params.Count
                    webobj.CustomOLEDBAddParameter(Params.Item(i), SqlDbType.NVarChar, 255, GetKey(Params, i))
                Next
            End If
            'End Arby

            webobj.CustomOLEDBSQLExecute()
            If Len(webobj.LastError) > 0 Then
                RunDynamicSQLStoreProc = webobj.LastError
                logger.Error(webobj.LastError + ", sql: " + _sql)
            End If
errEnd:
            webobj.CloseOLEDBConnection()
        Catch ex As Exception
            RunDynamicSQLStoreProc = ex.ToString
            logger.Error(ex, "_sql: " + _sql)
        End Try
    End Function

    'Arby 
    'Retrieve the Key from Collection
    Private Function GetKey(ByVal col As Collection, ByVal index As Integer) As String
        Dim listfield As FieldInfo = GetType(Collection).GetField("m_KeyedNodesHash", BindingFlags.NonPublic Or BindingFlags.Instance)
        Dim list As Object = listfield.GetValue(col)
        Dim keylist As IEnumerable(Of String) = list.Keys
        Dim key As String = keylist.ElementAt(index - 1)
        Return key
    End Function

    Public Function ExecuteTask_TryDeleteUnusedFiles() As String
        Return GlobalFunctions.ExecuteTask_DeleteUnusedFiles()
    End Function

    'Public Function UploadFilesToDropBoxAsync(ByVal TokenKey As String, ByVal DropBoxFolder As String, ByVal FilePath As String, ByRef ErrorMsg As String) As Boolean

    '    Dim upload As New AmberSoftDBPlugin.AmberSoftDB

    '    Try
    '        upload.UploadFilesToDropBoxAsync(TokenKey, DropBoxFolder, FilePath).Wait()
    '        Return True
    '    Catch ex As Exception
    '        ErrorMsg = ex.Message
    '    End Try

    '    Return False
    'End Function

End Class

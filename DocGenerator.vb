Imports System.Text
Imports GemBox.Document
Imports GemBox.Document.Tables
Imports System.IO
Imports System.Linq
Imports MindFusion.Diagramming
Imports System.Data
Imports System.Drawing

Public Class DocGenarator

    Dim document As DocumentModel = Nothing
    Dim FormCounter As Integer = 1
    Dim WorkflowCounter As Integer = 1
    Dim heading1 = ""
    Dim heading2 = ""
    Dim Normal1 = ""
    Dim _tempfolder As String = ""
    Dim formtemplate As String = ""
    Dim wtemplate1 As String = ""
    Dim wtemplate2 As String = ""
    Dim outPath As String
    Dim _resourcesFolder As String = ""
    Private _chart As New Diagram

    Public Property ResourcesFolder() As String
        Get
            Return _resourcesFolder
        End Get
        Set(ByVal value As String)
            _resourcesFolder = value
        End Set
    End Property


    Public Function theFieldType(ByVal typeNo As String) As String

        If IsNumeric(typeNo) Then

            Dim Num As Integer = Convert.ToInt32(typeNo)
            Select Case Num
                Case -1
                    Return "Auto"
                Case 0
                    Return "Single Line Text"
                Case 1
                    Return "Multi Line Text"
                Case 2
                    Return "Number"
                Case 3
                    Return "Decimal"
                Case 4
                    Return "Datetime"
                Case 5
                    Return "File Upload"
                Case 6
                    Return "Yes/No"
                Case 7
                    Return "Auto Number"
                Case 8
                    Return "User/Role Picker"
                Case 9
                    Return "Drop Down"
                Case 10
                    Return "Date"
                Case 11
                    Return "File Size"
                Case 12
                    Return "Version No"
                Case 13
                    Return "Submission No"
                Case 14
                    Return "Time in minutes"
                Case 15
                    Return "Money (Currency)"
                Case 20
                    Return "Table"
                Case 21
                    Return "Label"
                Case 22
                    Return "Radio Button"
                Case 23
                    Return "Country Picker"
                Case 24
                    Return "Button"
                Case 25
                    Return "CalField"
                Case 26
                    Return "Image"
                Case 27
                    Return "Auto ID"
                Case 28
                    Return "Header"
                Case 29
                    Return "Frame"
                Case 30
                    Return "Tiff Viewer"
                Case 31
                    Return "Hidden Field"
                Case 32
                    Return "Check List"
                Case 33
                    Return "DB Label"
                Case 34
                    Return "Html Input"
                Case Else
                    Return Nothing


            End Select
        Else
            Return Nothing
        End If

    End Function

    Public Sub WriteSubFormDetailandSchema(ByRef WebObj As ZukamiLib.WebSession, ByVal MasterFormID As Guid)
        Dim _sql As String = "Select [ListID] from [Lists] where [ParentFormID]='" & MasterFormID.ToString & "'"
        WebObj.CustomSQLCommand(_sql)
        WebObj.CustomClearParameters()
        Dim _set As DataSet = WebObj.CustomSQLExecuteReturn()
        If _set Is Nothing = False Then
            For i As Integer = 0 To _set.Tables(0).Rows.Count - 1
                Dim _LID As Guid = (_set.Tables(0).Rows(i).Item("ListID"))
                WriteFormDetailandSchema(WebObj, _LID)
            Next
        End If

    End Sub


    Public Sub InitializeGenerator(ByRef WebObj As ZukamiLib.WebSession, ByVal formdoctemp As String, ByVal workflowdoctemp1 As String, ByVal workflowdoctemp2 As String)
        ComponentInfo.SetLicense("DMPX-J9AT-EL54-2YBC")
        document = New DocumentModel

        heading1 = DirectCast(Style.CreateStyle(StyleTemplateType.Heading1, document), ParagraphStyle)
        heading2 = DirectCast(Style.CreateStyle(StyleTemplateType.Heading2, document), ParagraphStyle)
        document.Styles.Add(heading1)
        document.Styles.Add(heading2)


        document.DefaultCharacterFormat.Size = 11
        document.DefaultCharacterFormat.FontName = "Calibri"
        formtemplate = formdoctemp
        wtemplate1 = workflowdoctemp1
        wtemplate2 = workflowdoctemp2
    End Sub

    Public Sub CreateWord(ByRef WebObj As ZukamiLib.WebSession, ByVal WordFolder As String, ByVal WordFileName As String)
        Try
            'Dim _sql As String = ""
            '_sql = "SELECT [Name] from [lists] where [ListID]='" & FormID.ToString & "'"
            'WebObj.CustomSQLCommand(_sql)
            'WebObj.CustomClearParameters()
            'Dim _set As DataSet = WebObj.CustomSQLExecuteReturn()
            'If _set Is Nothing = False Then

            '    If _set.Tables(0).Rows.Count <> 0 Then
            '        Dim _Name As String = (_set.Tables(0).Rows(0).Item("Name")).ToString
            '        outPath = "C:\Users\user\Desktop\Flare\DocGenerated\" & _Name & "(" & FormID.ToString & ").docx"
            '    End If
            'End If
            outPath = WordFolder + "\" + WordFileName
        Catch ex As Exception
            'System.IO.File.AppendAllText("C:\Users\user\Desktop\errordetail.txt", ex.Message)
        End Try

    End Sub

    Public Sub WriteFormDetailandSchema(ByRef WebObj As ZukamiLib.WebSession, ByVal FormID As Guid)

        'General Details
        Try
            Dim _sql As String = ""
            Dim Sec1 As Section = Nothing

            _sql = "SELECT [Name],[Caption],[Description],[AppName],[listID],[TableBindSource],[Owner],[Datecreated],[WorkflowID],[Published],[AllowAnon],[EnableSync] FROM [Lists] L inner join [Apps] A on L.[AppID]=A.ID where [ListID]='" & FormID.ToString & "'"
            WebObj.CustomSQLCommand(_sql)
            WebObj.CustomClearParameters()
            Dim _set As DataSet = WebObj.CustomSQLExecuteReturn()
            If _set Is Nothing = False Then

                If _set.Tables(0).Rows.Count <> 0 Then

                    Dim _SecID As String = "AP001_F" + FormCounter.ToString("0000")

                    FormCounter = FormCounter + 1
                    Dim _Name As String = (_set.Tables(0).Rows(0).Item("Name")).ToString
                    Dim _Caption As String = (_set.Tables(0).Rows(0).Item("Caption")).ToString
                    Dim _Description As String = (_set.Tables(0).Rows(0).Item("Description")).ToString
                    Dim _AppName As String = (_set.Tables(0).Rows(0).Item("AppName")).ToString
                    Dim _listID As String = (_set.Tables(0).Rows(0).Item("listID")).ToString
                    Dim _TableBindSource As String = (_set.Tables(0).Rows(0).Item("TableBindSource")).ToString
                    Dim _Owner As String = (_set.Tables(0).Rows(0).Item("Owner")).ToString
                    Dim _Datecreated As String = (_set.Tables(0).Rows(0).Item("Datecreated")).ToString
                    Dim _WorkflowID As String = (_set.Tables(0).Rows(0).Item("WorkflowID")).ToString
                    If _WorkflowID.Trim = "" Then
                        _WorkflowID = "No workflow"
                    End If
                    Dim _Published As String = (_set.Tables(0).Rows(0).Item("Published")).ToString
                    If _Published.Trim = "True" Then
                        _Published = "Yes"
                    Else
                        _Published = "No"
                    End If
                    Dim _AllowAnon As String = (_set.Tables(0).Rows(0).Item("AllowAnon")).ToString
                    If _AllowAnon.Trim = "True" Then
                        _AllowAnon = "Yes"
                    Else
                        _AllowAnon = "No"
                    End If
                    Dim _EnableSync As String = (_set.Tables(0).Rows(0).Item("EnableSync")).ToString
                    If _EnableSync.Trim = "True" Then
                        _EnableSync = "Yes"
                    Else
                        _EnableSync = "No"
                    End If


                    Dim _sql2 As String = "SELECT U.Fullname,D.Groupname from [Users] U inner join [GroupMembership] G on U.UserID=G.UserID inner join [Groups] D on D.GroupID=G.GroupID where U.[UserID]='" & _Owner & "'"
                    WebObj.CustomSQLCommand(_sql2)
                    WebObj.CustomClearParameters()
                    Dim _ownerStr As String = ""
                    Dim _set2 As DataSet = WebObj.CustomSQLExecuteReturn()
                    If _set2 Is Nothing = False Then
                        If _set2.Tables(0).Rows.Count <> 0 Then
                            Dim _fullname As String = (_set2.Tables(0).Rows(0).Item("Fullname")).ToString
                            Dim _groupname As String = ""
                            For r As Integer = 0 To _set2.Tables(0).Rows.Count - 1

                                If r = 0 Then
                                    _groupname = (_set2.Tables(0).Rows(r).Item("Groupname")).ToString
                                Else
                                    _groupname = _groupname + ", " + (_set2.Tables(0).Rows(r).Item("Groupname")).ToString
                                End If

                            Next

                            _ownerStr = _fullname + "  ( " + _groupname + " )"
                        End If
                    End If

                    'outPath = "C:\Users\user\Desktop\Flare\DocGenerated\" & _Name & "1.docx"

                    Dim fileName As String = outPath
                    Dim pathToFileDirectory As String = formtemplate.Substring(0, formtemplate.LastIndexOf("\") + 1)
                    Dim sourceFile As String = formtemplate.Substring(formtemplate.LastIndexOf("\") + 1)

                    Dim sourceDocument As DocumentModel = DocumentModel.Load(System.IO.Path.Combine(pathToFileDirectory, sourceFile), DocxLoadOptions.DocxDefault)

                    ' Import all sections from source document. 
                    For Each sourceSection As Section In sourceDocument.Sections
                        Dim destinationSection As Section = document.Import(sourceSection, True, True)
                        document.Sections.Add(destinationSection)
                    Next




                    For Each paragraph As Paragraph In document.GetChildElements(True, ElementType.Paragraph)
                        For Each run As Run In paragraph.GetChildElements(True, ElementType.Run)
                            Dim isHeader As Boolean = False
                            Dim text As String = run.Text
                            If text = "[$APF]" Then
                                isHeader = True
                                run.Text = _SecID + ": " + _Name
                                document.MailMerge.Execute(paragraph)
                            End If

                        Next
                    Next
                    'document.Save(outPath, DocxSaveOptions.DocxDefault)

                    Dim tables As Table() = document.GetChildElements(True, ElementType.Table).Cast(Of Table)().ToArray()

                    '' First table 
                    Dim GeneralTable As Table = tables(tables.Count - 2)
                    'ElementCollection.Cast<TElement>(int index)
                    GeneralTable.Rows(0).Cells(1).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, _SecID))
                    GeneralTable.Rows(1).Cells(1).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, _Name))
                    GeneralTable.Rows(2).Cells(1).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, _Caption))
                    GeneralTable.Rows(3).Cells(1).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, _Description))
                    GeneralTable.Rows(4).Cells(1).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, _AppName))
                    GeneralTable.Rows(5).Cells(1).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, _listID))
                    GeneralTable.Rows(6).Cells(1).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, _TableBindSource))
                    GeneralTable.Rows(7).Cells(1).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, _ownerStr))
                    GeneralTable.Rows(8).Cells(1).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, _Datecreated))
                    GeneralTable.Rows(9).Cells(1).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, _WorkflowID))
                    GeneralTable.Rows(10).Cells(1).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, _Published))
                    GeneralTable.Rows(11).Cells(1).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, _AllowAnon))
                    GeneralTable.Rows(12).Cells(1).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, _EnableSync))



                    'Form Schema
                    _sql = "SELECT [FieldName],[FieldCaption],[FieldType],[MaxChars],[IsCompulsory],[DefValue] from [ListItems] where [ListID]='" & FormID.ToString & "'"
                    WebObj.CustomSQLCommand(_sql)
                    WebObj.CustomClearParameters()
                    _set = WebObj.CustomSQLExecuteReturn()
                    If _set Is Nothing = False Then
                        If _set.Tables(0).Rows.Count <> 0 Then

                            Dim _FieldName As String = ""
                            Dim _FieldCaption As String = ""
                            Dim _MaxChars As String = ""
                            Dim _IsCompulsory As String = ""
                            Dim _Defvalue As String = ""
                            Dim _FieldType As String = ""
                            Dim SchemaTable As Table = tables(tables.Count - 1)

                            For i As Integer = 1 To _set.Tables(0).Rows.Count - 1
                                SchemaTable.Rows.Insert(1, SchemaTable.Rows(1).Clone(True))
                            Next

                            For i As Integer = 0 To _set.Tables(0).Rows.Count - 1

                                _FieldName = ((_set.Tables(0).Rows(i).Item("FieldName")).ToString)
                                _FieldCaption = ((_set.Tables(0).Rows(i).Item("FieldCaption")).ToString)
                                _MaxChars = ((_set.Tables(0).Rows(i).Item("MaxChars")).ToString)
                                _IsCompulsory = ((_set.Tables(0).Rows(i).Item("IsCompulsory")).ToString)
                                If _IsCompulsory.Trim() = "True" Then
                                    _IsCompulsory = "Yes"
                                Else
                                    _IsCompulsory = "No"
                                End If
                                _Defvalue = ((_set.Tables(0).Rows(i).Item("Defvalue")).ToString)
                                _FieldType = theFieldType((_set.Tables(0).Rows(i).Item("FieldType")).ToString)


                                SchemaTable.Rows(i + 1).Cells(0).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, _FieldName))
                                SchemaTable.Rows(i + 1).Cells(1).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, _FieldCaption))
                                SchemaTable.Rows(i + 1).Cells(2).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, _FieldType))
                                SchemaTable.Rows(i + 1).Cells(3).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, _MaxChars))
                                SchemaTable.Rows(i + 1).Cells(4).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, _IsCompulsory))
                                SchemaTable.Rows(i + 1).Cells(5).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, _Defvalue))


                            Next

                        End If
                    End If




                Else
                    'MsgBox("No Details Found!")
                End If
            End If
        Catch ex As Exception
            'System.IO.File.AppendAllText("C:\Users\user\Desktop\errordetail.txt", ex.Message)
        End Try


        document.Save(outPath, DocxSaveOptions.DocxDefault)
    End Sub

    Public Property TempFolder() As String
        Get
            Return _tempfolder
        End Get
        Set(ByVal value As String)
            _tempfolder = value
        End Set
    End Property

    Public Sub WriteWorkFlowDetailandSchema(ByRef WebObj As ZukamiLib.WebSession, ByVal WorkflowID As Guid)
        Try
            'General Details
            Dim _sql As String = ""
            Dim Sec1 As Section = Nothing

            'changes start
            _sql = "select top 1 M.[Name],M.[Owner],M.[description],w.[LastModified],w.[Data],w.[Version] from MasterWorkflows M,WorkflowVersion W where w.Masterworkflowid = m.ID and m.ID='" & WorkflowID.ToString & "' Order By w.version desc"
            'changes end
            WebObj.CustomSQLCommand(_sql)
            WebObj.CustomClearParameters()
            Dim _set As DataSet = WebObj.CustomSQLExecuteReturn()
            If _set Is Nothing = False Then
                If _set.Tables(0).Rows.Count <> 0 Then

                    Dim _SecID As String = "AP001_W" + WorkflowCounter.ToString("0000")

                    WorkflowCounter = WorkflowCounter + 1
                    Dim _Name As String = (_set.Tables(0).Rows(0).Item("Name")).ToString
                    Dim _Owner As String = (_set.Tables(0).Rows(0).Item("Owner")).ToString
                    Dim _desc As String = (_set.Tables(0).Rows(0).Item("description")).ToString
                    Dim _Lastmodified As String = (_set.Tables(0).Rows(0).Item("LastModified")).ToString
                    Dim _mydata As String = (_set.Tables(0).Rows(0).Item("Data")).ToString
                    'changes start
                    Dim _version As String = (_set.Tables(0).Rows(0).Item("Version")).ToString
                    'changes end
                    _chart.LoadFromString(_mydata)
                    Dim image As System.Drawing.Image = _chart.CreateImage()
                    Dim _fpath As String = _tempfolder.TrimEnd("\") & "\" & Guid.NewGuid.ToString
                    Try
                        System.IO.Directory.CreateDirectory(_fpath)
                    Catch ex As Exception
                    End Try

                    Dim _targetpath As String = _fpath.TrimEnd("\") & "\" & _SecID & ".jpg"

                    image.Save(_targetpath)

                    Dim _sql2 As String = "SELECT U.Fullname,D.Groupname from [Users] U inner join [GroupMembership] G on U.UserID=G.UserID inner join [Groups] D on D.GroupID=G.GroupID where U.[UserID]='" & _Owner & "'"
                    WebObj.CustomSQLCommand(_sql2)
                    WebObj.CustomClearParameters()
                    Dim _ownerStr As String = ""
                    Dim _set2 As DataSet = WebObj.CustomSQLExecuteReturn()
                    If _set2 Is Nothing = False Then
                        If _set2.Tables(0).Rows.Count <> 0 Then
                            Dim _fullname As String = (_set2.Tables(0).Rows(0).Item("Fullname")).ToString
                            Dim _groupname As String = ""
                            For r As Integer = 0 To _set2.Tables(0).Rows.Count - 1
                                If r = 0 Then
                                    _groupname = (_set2.Tables(0).Rows(r).Item("Groupname")).ToString
                                Else
                                    _groupname = _groupname + ", " + (_set2.Tables(0).Rows(r).Item("Groupname")).ToString
                                End If
                            Next
                            _ownerStr = _fullname + "  ( " + _groupname + " )"
                        End If
                    End If

                    Dim fileName As String = outPath
                    Dim pathToFileDirectory As String = wtemplate1.Substring(0, wtemplate1.LastIndexOf("\") + 1)
                    Dim sourceFile As String = wtemplate1.Substring(wtemplate1.LastIndexOf("\") + 1)

                    Dim sourceDocument As DocumentModel = DocumentModel.Load(System.IO.Path.Combine(pathToFileDirectory, sourceFile), DocxLoadOptions.DocxDefault)

                    ' Import all sections from source document. 
                    Dim destinationSection1 As Section = Nothing
                    For Each sourceSection As Section In sourceDocument.Sections
                        destinationSection1 = document.Import(sourceSection, True, True)
                        'document.Sections.Add(destinationSection1)
                    Next


                    'Dim picsection As Section = Nothing
                    Dim pic As Picture = Nothing
                    For Each paragraph As Paragraph In sourceDocument.GetChildElements(True, ElementType.Paragraph)
                        For Each run As Run In paragraph.GetChildElements(True, ElementType.Run)

                            Dim text As String = run.Text
                            If text = "[$ScreenShot]" Then
                                pic = New Picture( _
                             document, _
                             _targetpath, _
                             LengthUnitConverter.Convert(500, LengthUnit.Pixel, LengthUnit.Point), _
                             LengthUnitConverter.Convert(400, LengthUnit.Pixel, LengthUnit.Point))
                                'picsection = New Section(document, New Paragraph(document, pic))
                            End If
                        Next
                    Next
                    'document.Sections.Add(picsection)
                    destinationSection1.Blocks.Add(New Paragraph(document, pic))
                    document.Sections.Add(destinationSection1)

                    For Each paragraph As Paragraph In document.GetChildElements(True, ElementType.Paragraph)
                        For Each run As Run In paragraph.GetChildElements(True, ElementType.Run)
                            Dim isHeader As Boolean = False
                            Dim text As String = run.Text
                            If text = "[$APW]" Then
                                isHeader = True
                                run.Text = _SecID + ": " + _Name
                                document.MailMerge.Execute(paragraph)
                            End If
                            If text = "[$ScreenShot]" Then
                                run.Text = ""
                                document.MailMerge.Execute(paragraph)
                            End If
                        Next
                    Next

                    Dim tables As Table() = document.GetChildElements(True, ElementType.Table).Cast(Of Table)().ToArray()

                    Dim GeneralTable As Table = tables(tables.Count - 1)
                    GeneralTable.Rows(0).Cells(1).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, _SecID))
                    GeneralTable.Rows(1).Cells(1).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, _Name))
                    GeneralTable.Rows(2).Cells(1).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, _ownerStr))
                    GeneralTable.Rows(3).Cells(1).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, _desc))
                    GeneralTable.Rows(4).Cells(1).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, _Lastmodified))


                    'changes start
                    _sql = "select N.* from NodeConfig N, WorkflowVersion W where W.ID=N.WorkflowID and W.MasterWorkflowID='" & WorkflowID.ToString & "' and W.Version=" & _version & " order by caption desc"
                    'changes end
                    WebObj.CustomSQLCommand(_sql)
                    WebObj.CustomClearParameters()
                    Dim _set3 As DataSet = WebObj.CustomSQLExecuteReturn()
                    If _set3 Is Nothing = False Then
                        If _set3.Tables(0).Rows.Count <> 0 Then

                            Dim _Caption As String = ""
                            Dim _bubbletype As String = ""
                            Dim _evaluators As String = ""
                            Dim _message As String = ""
                            Dim _Deadline As String = ""
                            Dim _availableAction As String = ""
                            Dim _PermissionAllow As New ArrayList
                            Dim _Notify As String = ""
                            Dim isHeader As Boolean = False

                            pathToFileDirectory = wtemplate2.Substring(0, wtemplate2.LastIndexOf("\") + 1)
                            sourceFile = wtemplate2.Substring(wtemplate2.LastIndexOf("\") + 1)
                            sourceDocument = DocumentModel.Load(System.IO.Path.Combine(pathToFileDirectory, sourceFile), DocxLoadOptions.DocxDefault)
                            Dim destinationSection As New Section(document)
                            For i As Integer = 0 To _set3.Tables(0).Rows.Count - 1
                                _Caption = ((_set3.Tables(0).Rows(_set3.Tables(0).Rows.Count - 1 - i).Item("Caption")).ToString)
                                Dim destinationPara As Element = Nothing
                                For Each sourcePara As Element In sourceDocument.GetChildElements(True)
                                    destinationPara = document.Import(sourcePara, True, True)
                                    If destinationPara.ElementType = ElementType.Table Or (destinationPara.ElementType = ElementType.Run) Then
                                        If destinationPara.ElementType = ElementType.Run Then
                                            Dim run As Run = destinationPara
                                            If run.Text.StartsWith("[$bubble desc") And isHeader = False Then
                                                isHeader = True
                                                destinationSection.Blocks.Add(New Paragraph(document, "Workflow Activities Bubble Description") With _
                                                 {.ParagraphFormat = New ParagraphFormat() With {.Style = heading1}})
                                            ElseIf run.Text.StartsWith("[$Bubble caption") Then
                                                destinationSection.Blocks.Add(New Paragraph(document, _Caption) With _
                                                 {.ParagraphFormat = New ParagraphFormat() With {.Style = heading2}})
                                            End If
                                        Else
                                            destinationSection.Blocks.Add(destinationPara)
                                        End If

                                    End If
                                Next


                                destinationSection.Blocks.Add(New Paragraph(document, ""))
                            Next
                            document.Sections.Add(destinationSection)



                            For i As Integer = 0 To _set3.Tables(0).Rows.Count - 1

                                _Caption = ((_set3.Tables(0).Rows(i).Item("Caption")).ToString)
                                _bubbletype = ((_set3.Tables(0).Rows(i).Item("Type")).ToString)
                                If _bubbletype = "SENDTO" Then
                                    Dim _xmldata As String = ((_set3.Tables(0).Rows(i).Item("NodeConfig")).ToString)

                                    _Deadline = _xmldata.Substring(_xmldata.IndexOf("<ApproveWithin>") + "<ApproveWithin>".Length, _xmldata.IndexOf("</ApproveWithin>") - (_xmldata.IndexOf("<ApproveWithin>") + "<ApproveWithin>".Length))
                                    If _Deadline.Trim() = "0" Then
                                        _Deadline = "No Deadline"
                                    Else
                                        _Deadline = _Deadline + " Days"
                                    End If

                                    _message = _xmldata.Substring(_xmldata.IndexOf("<Description>") + "<Description>".Length, _xmldata.IndexOf("</Description>") - (_xmldata.IndexOf("<Description>") + "<Description>".Length))
                                    If _message.Trim() = "" Then
                                        _message = " - "
                                    End If

                                    _Notify = _xmldata.Substring(_xmldata.IndexOf("<PostActionSubscribers>") + "<PostActionSubscribers>".Length, _xmldata.IndexOf("</PostActionSubscribers>") - (_xmldata.IndexOf("<PostActionSubscribers>") + "<PostActionSubscribers>".Length))
                                    If _Notify.Trim() = "" Then
                                        _Notify = _xmldata.Substring(_xmldata.IndexOf("<ExpirySubscribers>") + "<ExpirySubscribers>".Length, _xmldata.IndexOf("</ExpirySubscribers>") - (_xmldata.IndexOf("<ExpirySubscribers>") + "<ExpirySubscribers>".Length))
                                    Else
                                        _Notify = _Notify + ", " + _xmldata.Substring(_xmldata.IndexOf("<ExpirySubscribers>") + "<ExpirySubscribers>".Length, _xmldata.IndexOf("</ExpirySubscribers>") - (_xmldata.IndexOf("<ExpirySubscribers>") + "<ExpirySubscribers>".Length))
                                    End If
                                    If _Notify.Trim() = "" Then
                                        _Notify = " - "
                                    End If


                                    _availableAction = _xmldata.Substring(_xmldata.IndexOf("<ActionBarID>") + "<ActionBarID>".Length, _xmldata.IndexOf("</ActionBarID>") - (_xmldata.IndexOf("<ActionBarID>") + "<ActionBarID>".Length))
                                    If _availableAction.Trim() = "" Then
                                        'default
                                        _availableAction = "Approve, Reject"
                                    Else
                                        _sql = "SELECT [ActionBarID],[ButtonName],[ButtonOrder] FROM [ActionBarItems] where ActionBarID='" & _availableAction.Trim & "' order by ButtonOrder"
                                        WebObj.CustomSQLCommand(_sql)
                                        WebObj.CustomClearParameters()
                                        _availableAction = ""
                                        Dim _set4 As DataSet = WebObj.CustomSQLExecuteReturn()
                                        If _set4 Is Nothing = False Then
                                            If _set4.Tables(0).Rows.Count <> 0 Then
                                                For a As Integer = 0 To _set4.Tables(0).Rows.Count - 1
                                                    If _availableAction = "" Then
                                                        _availableAction = _set4.Tables(0).Rows(a).Item("ButtonName").ToString
                                                    Else
                                                        _availableAction = _availableAction + ", " + _set4.Tables(0).Rows(a).Item("ButtonName").ToString
                                                    End If
                                                Next
                                            End If
                                        End If
                                    End If

                                    _PermissionAllow.Clear()
                                    If _xmldata.Substring(_xmldata.IndexOf("<AllowChanges>") + "<AllowChanges>".Length, _xmldata.IndexOf("</AllowChanges>") - (_xmldata.IndexOf("<AllowChanges>") + "<AllowChanges>".Length)) = "True" Then
                                        _PermissionAllow.Add("Allow changes to attached data")
                                    End If
                                    If _xmldata.Substring(_xmldata.IndexOf("<AllowPostingOfMessages>") + "<AllowPostingOfMessages>".Length, _xmldata.IndexOf("</AllowPostingOfMessages>") - (_xmldata.IndexOf("<AllowPostingOfMessages>") + "<AllowPostingOfMessages>".Length)) = "True" Then
                                        _PermissionAllow.Add("Allow posting of messages")
                                    End If
                                    If _xmldata.Substring(_xmldata.IndexOf("<AllowTaskReassignment>") + "<AllowTaskReassignment>".Length, _xmldata.IndexOf("</AllowTaskReassignment>") - (_xmldata.IndexOf("<AllowTaskReassignment>") + "<AllowTaskReassignment>".Length)) = "True" Then
                                        _PermissionAllow.Add("Allow reassignment of task")
                                    End If
                                    If _xmldata.Substring(_xmldata.IndexOf("<AllowReviewRequest>") + "<AllowReviewRequest>".Length, _xmldata.IndexOf("</AllowReviewRequest>") - (_xmldata.IndexOf("<AllowReviewRequest>") + "<AllowReviewRequest>".Length)) = "True" Then
                                        _PermissionAllow.Add("Allow review request")
                                    End If
                                    If _xmldata.Substring(_xmldata.IndexOf("<ApproveWithoutConfirmation>") + "<ApproveWithoutConfirmation>".Length, _xmldata.IndexOf("</ApproveWithoutConfirmation>") - (_xmldata.IndexOf("<ApproveWithoutConfirmation>") + "<ApproveWithoutConfirmation>".Length)) = "True" Then
                                        _PermissionAllow.Add("Approve without confirmation")
                                    End If
                                    If _xmldata.Substring(_xmldata.IndexOf("<RejectWithoutRemarks>") + "<RejectWithoutRemarks>".Length, _xmldata.IndexOf("</RejectWithoutRemarks>") - (_xmldata.IndexOf("<RejectWithoutRemarks>") + "<RejectWithoutRemarks>".Length)) = "True" Then
                                        _PermissionAllow.Add("Reject without remarks")
                                    End If
                                    If _xmldata.Substring(_xmldata.IndexOf("<ContinueWaitingUponExpiry>") + "<ContinueWaitingUponExpiry>".Length, _xmldata.IndexOf("</ContinueWaitingUponExpiry>") - (_xmldata.IndexOf("<ContinueWaitingUponExpiry>") + "<ContinueWaitingUponExpiry>".Length)) = "True" Then
                                        _PermissionAllow.Add("Continue waiting upon expiry")
                                    End If

                                    _evaluators = _xmldata.Substring(_xmldata.IndexOf("<Evaluators>") + "<Evaluators>".Length, _xmldata.IndexOf("</Evaluators>") - (_xmldata.IndexOf("<Evaluators>") + "<Evaluators>".Length))
                                    If _evaluators.Trim() = "" Then
                                        _evaluators = " - "
                                    Else
                                        _sql = "select [Groupname] from groups where GroupID='" & _evaluators & "'"
                                        WebObj.CustomSQLCommand(_sql)
                                        WebObj.CustomClearParameters()
                                        _evaluators = ""
                                        Dim _set5 As DataSet = WebObj.CustomSQLExecuteReturn()
                                        If _set5 Is Nothing = False Then
                                            If _set5.Tables(0).Rows.Count <> 0 Then
                                                _evaluators = _set5.Tables(0).Rows(0).Item("Groupname").ToString
                                            End If
                                        End If
                                    End If


                                Else


                                    _Deadline = " - "
                                    _message = " - "
                                    _Notify = " - "
                                    _evaluators = " - "
                                    _PermissionAllow.Clear()
                                    _availableAction = " - "
                                End If

                                If _bubbletype = "SENDTO" Then
                                    _bubbletype = "User Action Request"
                                ElseIf _bubbletype = "SENDEMAIL" Then
                                    _bubbletype = "Generate Email"
                                ElseIf _bubbletype = "RUNCODE" Then
                                    _bubbletype = "Run Custom Code"
                                Else
                                    _bubbletype = "Others"
                                End If


                                Dim tables2 As Table() = document.GetChildElements(True, ElementType.Table).Cast(Of Table)().ToArray()

                                Dim bubbletable As Table = tables2(tables2.Count - 1 - i)
                                bubbletable.Rows(0).Cells(1).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, _Caption.Trim))
                                bubbletable.Rows(1).Cells(1).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, _bubbletype.Trim))
                                bubbletable.Rows(2).Cells(1).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, _evaluators.Trim))
                                bubbletable.Rows(3).Cells(1).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, _message.Trim))
                                bubbletable.Rows(4).Cells(1).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, _Deadline.Trim))
                                bubbletable.Rows(5).Cells(1).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, _availableAction.Trim))
                                For p As Integer = 0 To _PermissionAllow.Count - 1

                                    bubbletable.Rows(6).Cells(1).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Picture(document, _resourcesFolder.TrimEnd("\") & "\template\tick.png"))
                                    bubbletable.Rows(6).Cells(1).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, _PermissionAllow(p)))
                                    bubbletable.Rows(6).Cells(1).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New SpecialCharacter(document, SpecialCharacterType.LineBreak))
                                Next
                                If _PermissionAllow.Count = 0 Then
                                    bubbletable.Rows(6).Cells(1).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, " - "))
                                End If
                                bubbletable.Rows(7).Cells(1).Blocks.Cast(Of Paragraph)(0).Inlines.Add(New Run(document, _Notify.Trim))
                            Next
                        End If
                    End If

                End If
            End If

        Catch ex As Exception

            'System.IO.File.WriteAllText("d:\vbtest\dang.txt", ex.ToString)
        End Try
        document.Save(outPath, DocxSaveOptions.DocxDefault)
    End Sub


End Class

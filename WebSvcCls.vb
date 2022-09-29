Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Linq
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Net
Imports System.IO
Imports System.Reflection
Imports System.CodeDom.Compiler
Imports System.Web.Services.Description
Imports System.Xml.Serialization
Imports System.CodeDom
Imports System.Collections
Imports System.Text
Imports System.Text.RegularExpressions
Imports NLog

Public Class WebSvcCls

    Private Shared logger As Logger = LogManager.GetCurrentClassLogger()
    Private Sub RegExp(strHTML As String)
        Dim ht As New Hashtable()
        Dim qariRegex As New Regex("<Table>\s*<Country>([" + "\s\S]*?)</Country>\s*<City>([\s\S]*?)</City>\s*</Table>")
        Dim mc As MatchCollection = qariRegex.Matches(strHTML)
        'CaptureCollection cc = mc[0].Captures;
        ht.Add(mc(0).Captures, mc(1).Captures)
    End Sub

    Private Shared Function RegExpForCountryCity(strHTML As String) As String
        Dim qariRegex As New Regex("<Table>\s*<Country>(?<Country>" + "[\s\S]*?)</Country>\s*<City>(?<City>[\s\S]*?)</City>\s*</Table>", RegexOptions.IgnoreCase Or RegexOptions.Multiline)
        Dim mc As MatchCollection = qariRegex.Matches(strHTML)
        Dim strCountryCity As String = ""
        For i As Integer = 0 To mc.Count - 1
            If String.IsNullOrEmpty(strCountryCity) Then
                strCountryCity = "Country: " + "<b>" + mc(i).Groups("Country").Value + "</b>" + " " + "City: " + "<b>" + mc(i).Groups("City").Value + "</b>" + "</br>"
            Else
                strCountryCity += "</br>" + "Country: " + "<b>" + mc(i).Groups("Country").Value + "</b>" + " " + "City: " + "<b>" + mc(i).Groups("City").Value + "</b>" + "</br>"
            End If
        Next
        Return strCountryCity
    End Function

    Public Shared Function GetWebServiceServices(ByRef ddServices As DropDownList, URL As String, Optional DLLs As String = "", Optional username As String = "", Optional password As String = "", Optional domain As String = "", Optional ProxyURL As String = "", Optional ProxyPort As Integer = -1) As Object
        logger.Debug("start")
        Try
            ddServices.Items.Clear()
            Dim li As New ListItem("Select a service", "")
            li.Selected = True
            ddServices.Items.Add(li)

            Dim client As New System.Net.WebClient()
            If Len(username) > 0 Then
                If Len(domain) > 0 Then
                    client.Credentials = New NetworkCredential(username, password, domain)
                Else
                    client.Credentials = New NetworkCredential(username, password)
                End If

            End If
            If Len(ProxyURL) > 0 Then
                client.Proxy = New WebProxy(ProxyURL, ProxyPort)
            End If

            logger.Debug("check point 1")
            Dim stream As System.IO.Stream = client.OpenRead(URL)
            Dim description As ServiceDescription = ServiceDescription.Read(stream)
            Dim importer As New ServiceDescriptionImporter()
            importer.ProtocolName = "Soap12"
            importer.AddServiceDescription(description, Nothing, Nothing)
            importer.Style = ServiceDescriptionImportStyle.Client
            importer.CodeGenerationOptions = System.Xml.Serialization.CodeGenerationOptions.GenerateProperties
            Dim nmspace As New CodeNamespace()
            Dim unit1 As New CodeCompileUnit()
            unit1.Namespaces.Add(nmspace)
            Dim warning As ServiceDescriptionImportWarnings = importer.Import(nmspace, unit1)
            logger.Debug("check point 2, warning: [" + warning.ToString() + "] " + Newtonsoft.Json.JsonConvert.SerializeObject(warning))
            If warning = 0 Or (GlobalFunctions.FromConfig("IgnoreWarningWhenGenerateCallWebServiceCode_2_OptionalExtensionsIgnored") = "true" And warning = 2) Then
                logger.Debug("check point 2.1")

                Dim arrDLL() As String
                If Len(DLLs) > 0 Then
                    arrDLL = Split(DLLs, ",")
                Else
                    ReDim arrDLL(5)
                    arrDLL(0) = "System.Web.Services.dll"
                    arrDLL(1) = "System.Xml.dll"
                    arrDLL(2) = "System.data.dll"
                    arrDLL(3) = "System.dll"
                    arrDLL(4) = "System.drawing.dll"
                    arrDLL(5) = "System.security.dll"
                End If

                Dim provider1 As CodeDomProvider = CodeDomProvider.CreateProvider("CSharp")
                Dim parms As New CompilerParameters(arrDLL)
                Dim results As CompilerResults = provider1.CompileAssemblyFromDom(parms, unit1)
                logger.Debug("check point 2.5")
                If results.Errors.Count > 0 Then
                    Dim _all As String = ""
                    For Each erroritem As CompilerError In results.Errors
                        _all += erroritem.ErrorText & "<br>"
                    Next erroritem
                    Return _all
                End If

                logger.Debug("check point 3")
                Dim types() As Type = results.CompiledAssembly.GetTypes
                Dim _foundType As Type = Nothing
                For Each t As Type In types
                    Dim listitem As New ListItem(t.ToString, t.ToString)
                    ddServices.Items.Add(listitem)
                Next
            Else
                logger.Debug("check point 2.2")
                Return warning
            End If
        Catch ex As Exception
            logger.Error(ex)
        End Try
    End Function

    Public Shared Function GetWebServiceMethods(ByRef ddMethods As DropDownList, Servicename As String, URL As String, Optional DLLs As String = "", Optional username As String = "", Optional password As String = "", Optional domain As String = "", Optional ProxyURL As String = "", Optional ProxyPort As Integer = -1) As Object
        logger.Debug("start")
        Try
            ddMethods.Items.Clear()
            Dim li As New ListItem("Select a webmethod", "")
            li.Selected = True
            ddMethods.Items.Add(li)

            Dim client As New System.Net.WebClient()
            If Len(username) > 0 Then
                If Len(domain) > 0 Then
                    client.Credentials = New NetworkCredential(username, password, domain)
                Else
                    client.Credentials = New NetworkCredential(username, password)
                End If

            End If
            If Len(ProxyURL) > 0 Then
                client.Proxy = New WebProxy(ProxyURL, ProxyPort)
            End If


            Dim stream As System.IO.Stream = client.OpenRead(URL)
            Dim description As ServiceDescription = ServiceDescription.Read(stream)
            Dim importer As New ServiceDescriptionImporter()
            importer.ProtocolName = "Soap12"
            importer.AddServiceDescription(description, Nothing, Nothing)
            importer.Style = ServiceDescriptionImportStyle.Client
            importer.CodeGenerationOptions = System.Xml.Serialization.CodeGenerationOptions.GenerateProperties
            Dim nmspace As New CodeNamespace()
            Dim unit1 As New CodeCompileUnit()
            unit1.Namespaces.Add(nmspace)
            Dim warning As ServiceDescriptionImportWarnings = importer.Import(nmspace, unit1)
            logger.Debug("check point 1, warning: [" + warning.ToString() + "] " + Newtonsoft.Json.JsonConvert.SerializeObject(warning))
            If warning = 0 Or (GlobalFunctions.FromConfig("IgnoreWarningWhenGenerateCallWebServiceCode_2_OptionalExtensionsIgnored") = "true" And warning = 2) Then

                Dim arrDLL() As String
                If Len(DLLs) > 0 Then
                    arrDLL = Split(DLLs, ",")
                Else
                    ReDim arrDLL(5)
                    arrDLL(0) = "System.Web.Services.dll"
                    arrDLL(1) = "System.Xml.dll"
                    arrDLL(2) = "System.data.dll"
                    arrDLL(3) = "System.dll"
                    arrDLL(4) = "System.drawing.dll"
                    arrDLL(5) = "System.security.dll"
                End If

                Dim provider1 As CodeDomProvider = CodeDomProvider.CreateProvider("CSharp")
                Dim parms As New CompilerParameters(arrDLL)
                Dim results As CompilerResults = provider1.CompileAssemblyFromDom(parms, unit1)
                If results.Errors.Count > 0 Then
                    Dim _all As String = ""
                    For Each erroritem As CompilerError In results.Errors
                        _all += erroritem.ErrorText & "<br>"
                    Next erroritem
                    Return _all
                End If

                Dim MyService As Object = results.CompiledAssembly.CreateInstance(Servicename)
                Dim methodinfos() As MethodInfo = MyService.GetType().GetMethods()

                For Each m As MethodInfo In methodinfos

                    Dim listitem As New ListItem(m.Name, m.Name & ";" & m.GetParameters.Count)
                    ddMethods.Items.Add(listitem)
                Next
            Else
                Return warning
            End If
        Catch ex As Exception
            logger.Error(ex)
        End Try
    End Function


    ' Private Shared parms_default As New CompilerParameters(New String() {
    '     "System.Web.Services.dll",
    '     "System.Xml.dll",
    '     "System.data.dll",
    '     "System.dll",
    '     "System.drawing.dll",
    '     "System.security.dll"
    ' })
    'Private Shared CompilerResults_default As CompilerResults = Nothing
    Private Shared CompilerResults_default As New Dictionary(Of String, CompilerResults)
    Public Shared Function CallingWebService(URL As String, ServiceName As String, MethodName As String, Arguments() As Object, Optional DLLs As String = "", Optional username As String = "", Optional password As String = "", Optional domain As String = "", Optional ProxyURL As String = "", Optional ProxyPort As Integer = -1) As Object
        logger.Trace("start")
        Dim client As New System.Net.WebClient()
        If Len(username) > 0 Then
            If Len(domain) > 0 Then
                client.Credentials = New NetworkCredential(username, password, domain)
            Else
                client.Credentials = New NetworkCredential(username, password)
            End If

        End If
        If Len(ProxyURL) > 0 Then
            client.Proxy = New WebProxy(ProxyURL, ProxyPort)
        End If
        logger.Trace("step 1")


        Dim stream As System.IO.Stream = client.OpenRead(URL)
        Dim description As ServiceDescription = ServiceDescription.Read(stream)
        Dim importer As New ServiceDescriptionImporter()
        importer.ProtocolName = "Soap12"
        importer.AddServiceDescription(description, Nothing, Nothing)
        importer.Style = ServiceDescriptionImportStyle.Client
        importer.CodeGenerationOptions = System.Xml.Serialization.CodeGenerationOptions.GenerateProperties
        Dim nmspace As New CodeNamespace()
        Dim unit1 As New CodeCompileUnit()
        unit1.Namespaces.Add(nmspace)
        Dim warning As ServiceDescriptionImportWarnings = importer.Import(nmspace, unit1)
        Dim usingDefault As Boolean = False
        logger.Trace("step 2")
        logger.Debug("check point 1, warning: [" + warning.ToString() + "] " + Newtonsoft.Json.JsonConvert.SerializeObject(warning))
        If warning = 0 Or (GlobalFunctions.FromConfig("IgnoreWarningWhenGenerateCallWebServiceCode_2_OptionalExtensionsIgnored") = "true" And warning = 2) Then
            Dim parms As CompilerParameters
            Dim arrDLL() As String
            If Len(DLLs) > 0 Then
                arrDLL = Split(DLLs, ",")
                parms = New CompilerParameters(arrDLL)
            Else
                ReDim arrDLL(5)
                arrDLL(0) = "System.Web.Services.dll"
                arrDLL(1) = "System.Xml.dll"
                arrDLL(2) = "System.data.dll"
                arrDLL(3) = "System.dll"
                arrDLL(4) = "System.drawing.dll"
                arrDLL(5) = "System.security.dll"
                'parms = parms_default
                parms = New CompilerParameters(arrDLL)
                usingDefault = True
            End If
            logger.Trace("step 3")

            Dim provider1 As CodeDomProvider = CodeDomProvider.CreateProvider("CSharp")
            'Dim parms As New CompilerParameters(arrDLL)
            Dim results As CompilerResults = Nothing
            If usingDefault Then
                If CompilerResults_default.ContainsKey(URL.ToLower()) Then
                    logger.Trace("Compile: no need, load from default for URL: " + URL.ToLower())
                    results = CompilerResults_default.Item(URL.ToLower())
                Else
                    logger.Trace("Compile: load from default 1st time for URL: " + URL.ToLower())
                    results = provider1.CompileAssemblyFromDom(parms, unit1)
                    CompilerResults_default.Add(URL.ToLower(), results)
                End If
            Else
                logger.Trace("Compile: required for URL: " + URL.ToLower())
                results = provider1.CompileAssemblyFromDom(parms, unit1)
            End If
            logger.Trace("Compile done")

            If results.Errors.Count > 0 Then
                If CompilerResults_default.ContainsKey(URL.ToLower()) Then
                    CompilerResults_default.Remove(URL.ToLower())
                End If
                Dim _all As String = ""
                For Each erroritem As CompilerError In results.Errors
                    _all += erroritem.ErrorText & "<br>"
                Next erroritem
                logger.Trace("got error, _all: " + _all)
                Return _all
            End If
            logger.Trace("before create instance")
            Dim wsvcClass As Object = results.CompiledAssembly.CreateInstance(ServiceName)
            logger.Trace("instance created")
            Dim mi As MethodInfo = wsvcClass.[GetType]().GetMethod(MethodName)

            logger.Trace("before call method")
            Return mi.Invoke(wsvcClass, Arguments).ToString()
        Else
            Return warning
        End If
    End Function

    Public Shared Function DynamicInvocation() As String
        Dim uri As New Uri("http://www.webservicex.net/globalweather.asmx?WSDL")
        Dim webRequest__1 As WebRequest = WebRequest.Create(uri)
        Dim requestStream As System.IO.Stream = webRequest__1.GetResponse().GetResponseStream()
        Dim sd As ServiceDescription = ServiceDescription.Read(requestStream)
        Dim sdName As String = sd.Services(0).Name
        Dim servImport As New ServiceDescriptionImporter()
        servImport.AddServiceDescription(sd, [String].Empty, [String].Empty)
        servImport.ProtocolName = "Soap"
        servImport.CodeGenerationOptions = CodeGenerationOptions.GenerateProperties
        Dim [nameSpace] As New CodeNamespace()
        Dim codeCompileUnit As New CodeCompileUnit()
        codeCompileUnit.Namespaces.Add([nameSpace])
        Dim warnings As ServiceDescriptionImportWarnings = servImport.Import([nameSpace], codeCompileUnit)
        logger.Debug("check point 1, warning: [" + warnings.ToString() + "] " + Newtonsoft.Json.JsonConvert.SerializeObject(warnings))
        If warnings = 0 Or (GlobalFunctions.FromConfig("IgnoreWarningWhenGenerateCallWebServiceCode_2_OptionalExtensionsIgnored") = "true" And warnings = 2) Then
            Dim stringWriter As New StringWriter(System.Globalization.CultureInfo.CurrentCulture)
            Dim prov As New Microsoft.CSharp.CSharpCodeProvider()
            prov.GenerateCodeFromNamespace([nameSpace], stringWriter, New CodeGeneratorOptions())
            Dim assemblyReferences As String() = New String(1) {"System.Web.Services.dll", "System.Xml.dll"}
            Dim param As New CompilerParameters(assemblyReferences)
            param.GenerateExecutable = False
            param.GenerateInMemory = True
            param.TreatWarningsAsErrors = False
            param.WarningLevel = 4
            Dim results As New CompilerResults(New TempFileCollection())
            results = prov.CompileAssemblyFromDom(param, codeCompileUnit)
            Dim assembly As Assembly = results.CompiledAssembly
            Dim strClassName As String = String.Empty
            Dim strMethodName As String = String.Empty
            Dim types As Type() = assembly.GetTypes()
            Dim arrl As New ArrayList()
            For Each cls As Type In types
                arrl.Add("Namespace: " + cls.FullName)
                If cls.IsAbstract Then
                    arrl.Add("Abstract Class Name: " + cls.Name.ToString())
                ElseIf cls.IsPublic Then
                    arrl.Add("Public Class Name: " + cls.Name.ToString())
                ElseIf cls.IsSealed Then
                    arrl.Add("Sealed Class Name: " + cls.Name.ToString())
                End If
            Next

            For i As Integer = 0 To arrl.Count - 1
                If String.IsNullOrEmpty(strClassName) Then
                    strClassName = arrl(i).ToString()
                Else
                    strClassName += "</br>" + arrl(i).ToString()
                End If
            Next

            strMethodName = strClassName & Convert.ToString("</br>")
            For Each t As MethodInfo In assembly.[GetType](sdName).GetMethods()
                If t.Name = "Discover" Then
                    Exit For
                End If

                If String.IsNullOrEmpty(strMethodName) Then
                    strMethodName = "MethodName : " + t.Name + "</br>" + t.ToString()
                Else
                    strMethodName += System.Environment.NewLine + "</br></br>" + "MethodName : " + t.Name + "</br>" + t.ToString()
                End If
            Next

            Return strMethodName
        End If
    End Function
End Class

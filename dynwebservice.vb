Imports Microsoft.VisualBasic

Imports System.Collections.Generic
Imports System.Text
Imports System.Reflection
Imports System.CodeDom
Imports System.CodeDom.Compiler
Imports System.Web.Services
Imports System.Web.Services.Description
Imports System.Web.Services.Discovery
Imports System.Xml

Public Class WebServiceInvoker2
    Private availableTypes As Dictionary(Of String, Type)

    ''' <summary>
    ''' Text description of the available services within this web service.
    ''' </summary>
    Public ReadOnly Property AvailableServices() As List(Of String)
        Get
            Return Me.services
        End Get
    End Property

    ''' <summary>
    ''' Creates the service invoker using the specified web service.
    ''' </summary>
    ''' <param name="webServiceUri"></param>
    Public Sub New(ByVal webServiceUri As Uri)
        Me.services = New List(Of String)()
        ' available services
        Me.availableTypes = New Dictionary(Of String, Type)()
        ' available types
        ' create an assembly from the web service description
        Me.webServiceAssembly = BuildAssemblyFromWSDL(webServiceUri)

        ' see what service types are available
        Dim types As Type() = Me.webServiceAssembly.GetExportedTypes()

        ' and save them
        For Each type As Type In types
            services.Add(type.FullName)
            availableTypes.Add(type.FullName, type)
        Next
    End Sub

    ''' <summary>
    ''' Gets a list of all methods available for the specified service.
    ''' </summary>
    ''' <param name="serviceName"></param>
    ''' <returns></returns>
    Public Function EnumerateServiceMethods(ByVal serviceName As String) As List(Of String)
        Dim methods As New List(Of String)()

        If Not Me.availableTypes.ContainsKey(serviceName) Then
            Throw New Exception("Service Not Available")
        Else
            Dim type As Type = Me.availableTypes(serviceName)

            ' only find methods of this object type (the one we generated)
            ' we don't want inherited members (this type inherited from SoapHttpClientProtocol)
            For Each minfo As MethodInfo In type.GetMethods(BindingFlags.Instance Or BindingFlags.[Public] Or BindingFlags.DeclaredOnly)
                methods.Add(minfo.Name)
            Next

            Return methods
        End If
    End Function

    ''' <summary>
    ''' Invokes the specified method of the named service.
    ''' </summary>
    ''' <typeparam name="T">The expected return type.</typeparam>
    ''' <param name="serviceName">The name of the service to use.</param>
    ''' <param name="methodName">The name of the method to call.</param>
    ''' <param name="args">The arguments to the method.</param>
    ''' <returns>The return value from the web service method.</returns>
    Public Function InvokeMethod(Of T)(ByVal serviceName As String, ByVal methodName As String, ByVal ParamArray args As Object()) As T
        ' create an instance of the specified service
        ' and invoke the method
        Dim obj As Object = Me.webServiceAssembly.CreateInstance(serviceName)

        Dim type As Type = obj.[GetType]()

        Return DirectCast(type.InvokeMember(methodName, BindingFlags.InvokeMethod, Nothing, obj, args), T)
    End Function

    ''' <summary>
    ''' Builds the web service description importer, which allows us to generate a proxy class based on the 
    ''' content of the WSDL described by the XmlTextReader.
    ''' </summary>
    ''' <param name="xmlreader">The WSDL content, described by XML.</param>
    ''' <returns>A ServiceDescriptionImporter that can be used to create a proxy class.</returns>
    Private Function BuildServiceDescriptionImporter(ByVal xmlreader As XmlTextReader) As ServiceDescriptionImporter
        ' make sure xml describes a valid wsdl
        If Not ServiceDescription.CanRead(xmlreader) Then
            Throw New Exception("Invalid Web Service Description")
        End If

        ' parse wsdl
        Dim serviceDescription__1 As ServiceDescription = ServiceDescription.Read(xmlreader)

        ' build an importer, that assumes the SOAP protocol, client binding, and generates properties
        Dim descriptionImporter As New ServiceDescriptionImporter()
        descriptionImporter.ProtocolName = "Soap"
        descriptionImporter.AddServiceDescription(serviceDescription__1, Nothing, Nothing)
        descriptionImporter.Style = ServiceDescriptionImportStyle.Client
        descriptionImporter.CodeGenerationOptions = System.Xml.Serialization.CodeGenerationOptions.GenerateProperties

        Return descriptionImporter
    End Function

    ''' <summary>
    ''' Compiles an assembly from the proxy class provided by the ServiceDescriptionImporter.
    ''' </summary>
    ''' <param name="descriptionImporter"></param>
    ''' <returns>An assembly that can be used to execute the web service methods.</returns>
    Private Function CompileAssembly(ByVal descriptionImporter As ServiceDescriptionImporter) As Assembly
        ' a namespace and compile unit are needed by importer
        Dim codeNamespace As New CodeNamespace()
        Dim codeUnit As New CodeCompileUnit()

        codeUnit.Namespaces.Add(codeNamespace)

        Dim importWarnings As ServiceDescriptionImportWarnings = descriptionImporter.Import(codeNamespace, codeUnit)

        If importWarnings = 0 Then
            ' no warnings
            ' create a c# compiler
            Dim compiler As CodeDomProvider = CodeDomProvider.CreateProvider("VisualBasic")

            ' include the assembly references needed to compile
            Dim references As String() = New String(2) {"System.Web.Services.dll", "System.Xml.dll", "System.data.dll"}

            Dim parameters As New CompilerParameters(references)

            ' compile into assembly
            Dim results As CompilerResults = compiler.CompileAssemblyFromDom(parameters, codeUnit)

            For Each oops As CompilerError In results.Errors
                ' trap these errors and make them available to exception object
                Throw New Exception("Compilation Error Creating Assembly")
            Next

            ' all done....
            Return results.CompiledAssembly
        Else
            ' warnings issued from importers, something wrong with WSDL
            Throw New Exception("Invalid WSDL")
        End If
    End Function

    ''' <summary>
    ''' Builds an assembly from a web service description.
    ''' The assembly can be used to execute the web service methods.
    ''' </summary>
    ''' <param name="webServiceUri">Location of WSDL.</param>
    ''' <returns>A web service assembly.</returns>
    Private Function BuildAssemblyFromWSDL(ByVal webServiceUri As Uri) As Assembly
        If [String].IsNullOrEmpty(webServiceUri.ToString()) Then
            Throw New Exception("Web Service Not Found")
        End If

        Dim xmlreader As New XmlTextReader(webServiceUri.ToString() & "?wsdl")

        Dim descriptionImporter As ServiceDescriptionImporter = BuildServiceDescriptionImporter(xmlreader)

        Return CompileAssembly(descriptionImporter)
    End Function

    Private webServiceAssembly As Assembly
    Private services As List(Of String)
End Class

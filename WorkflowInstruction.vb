Imports Microsoft.VisualBasic
Imports System.Xml
Imports System.Text
Imports System.IO


Public Class WorkflowInstruction
    Public Enum ActionTypes
        ACTIONTYPE_ADDBUBBLE
        ACTIONTYPE_EDITBUBBLE
        ACTIONTYPE_SETARROW
    End Enum
    Private _ActionType As ActionTypes
    Private _BubbleID As Guid
    Private _Type As String
    Private _AttachedToID As String
    Private _Caption As String

    Public Property Caption() As String
        Get
            Return _Caption
        End Get
        Set(ByVal value As String)
            _Caption = value
        End Set
    End Property

    Public Function Save() As String
        Dim _xml As New XmlDocument
        Dim _xmlChild As XmlElement = _xml.CreateElement("WI")
        _xml.AppendChild(_xmlChild)

        GlobalFunctions.AddElement(_xml, _xmlChild, "ActionType", _ActionType)
        GlobalFunctions.AddElement(_xml, _xmlChild, "BubbleID", _BubbleID.ToString)
        GlobalFunctions.AddElement(_xml, _xmlChild, "Type", _Type)
        GlobalFunctions.AddElement(_xml, _xmlChild, "Caption", _Caption)
        GlobalFunctions.AddElement(_xml, _xmlChild, "AttachedToID", _AttachedToID)

        Dim _sb As New StringBuilder
        Dim _stringwriter As New StringWriter(_sb)
        _xml.Save(_stringwriter)

        Dim arrBytes() As Byte = ASCIIEncoding.Unicode.GetBytes(_sb.ToString)
        Return Convert.ToBase64String(arrBytes, Base64FormattingOptions.None)
    End Function

    Public Shared Function Load(ByVal MIMEString As String) As WorkflowInstruction
        If Len(MIMEString) = 0 Then Return Nothing

        Dim XMLData As String = ASCIIEncoding.Unicode.GetString(Convert.FromBase64String(MIMEString))
        If Len(XMLData) = 0 Then Return Nothing


        Dim _xml As New XmlDocument
        _xml.LoadXml(XMLData)
        Dim _xmlNL As XmlNodeList = _xml.GetElementsByTagName("WI")
        If _xmlNL.Count = 0 Then Return Nothing

        Dim _instruction As New WorkflowInstruction
        Dim _xmlActivity As XmlElement = _xmlNL.Item(0)
        _instruction.ActionType = CInt(GlobalFunctions.GetElement(_xmlActivity, "ActionType"))
        _instruction.BubbleID = New Guid(GlobalFunctions.GetElement(_xmlActivity, "BubbleID"))
        _instruction.Type = GlobalFunctions.GetElement(_xmlActivity, "Type")
        _instruction.Caption = GlobalFunctions.GetElement(_xmlActivity, "Caption")
        _instruction.AttachedToID = GlobalFunctions.GetElement(_xmlActivity, "AttachedToID")
        Return _instruction
    End Function


    Public Property Type() As String
        Get
            Return _Type
        End Get
        Set(ByVal value As String)
            _Type = value
        End Set
    End Property

    Public Property AttachedToID() As String
        Get
            Return _AttachedToID
        End Get
        Set(ByVal value As String)
            _AttachedToID = value
        End Set
    End Property

    Public Property BubbleID() As Guid
        Get
            Return _BubbleID
        End Get
        Set(ByVal value As Guid)
            _BubbleID = value
        End Set
    End Property

    Public Property ActionType() As ActionTypes
        Get
            Return _ActionType
        End Get
        Set(ByVal value As ActionTypes)
            _ActionType = value
        End Set
    End Property
End Class

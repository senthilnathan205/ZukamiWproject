Imports Microsoft.VisualBasic

Public Class SubstitutionFilter
    Inherits System.IO.Stream

    Private Base As System.IO.Stream

    Public Sub New(ByVal ResponseStream As System.IO.Stream)
        If ResponseStream Is Nothing Then Throw New ArgumentNullException("ResponseStream")
        Me.Base = ResponseStream
    End Sub

    Public Overrides ReadOnly Property CanRead() As Boolean
        Get

        End Get
    End Property

    Public Overrides ReadOnly Property CanSeek() As Boolean
        Get

        End Get
    End Property

    Public Overrides ReadOnly Property CanWrite() As Boolean
        Get

        End Get
    End Property

    Public Overrides Sub Flush()

    End Sub

    Public Overrides ReadOnly Property Length() As Long
        Get

        End Get
    End Property

    Public Overrides Property Position() As Long
        Get

        End Get
        Set(ByVal value As Long)

        End Set
    End Property

    Public Overrides Function Read(ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer) As Integer
        Return Me.Base.Read(buffer, offset, count)
    End Function

    Public Overrides Function Seek(ByVal offset As Long, ByVal origin As System.IO.SeekOrigin) As Long

    End Function

    Public Overrides Sub SetLength(ByVal value As Long)

    End Sub

    Public Overrides Sub Write(ByVal buffer() As Byte, ByVal offset As Integer, ByVal count As Integer)
        ' Get HTML code  
        Dim HTML As String = System.Text.Encoding.UTF8.GetString(buffer, offset, count)
        ' Replace the text with something else  
        HTML = HTML.Replace("Hello World!", "I've replaced the Hello World example!")

        ' Send output  
        buffer = System.Text.Encoding.UTF8.GetBytes(HTML)
        Me.Base.Write(buffer, 0, buffer.Length)
    End Sub

End Class

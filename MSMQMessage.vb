Imports Microsoft.VisualBasic
Imports System.Messaging
Imports NLog

Public Class MSMQMessage

    Private Shared logger As Logger = LogManager.GetCurrentClassLogger()

    Public Enum MessageTypes
        MT_UNKNOWN = -1
        MT_NEWSESSION = 0
        MT_TERMSESSION = 1
        MT_RESTSESSION = 2
        MT_NODEACTION = 3
        MT_NEWTIMERFLOW = 4
        MT_REMTIMERFLOW = 5
    End Enum
    Private _MessageType As MessageTypes = MessageTypes.MT_UNKNOWN
    Private _SessionGUID As String = ""
    Private _SessionNodeGUID As String = ""
    Private _TimerWorkflowGUID As String = ""
    Private _ExecutionThreadGUID As String = ""
    Private _w3mq As MessageQueue = Nothing
    Private _SilentMode As Boolean = False
    Private _lasterrormsg As String = ""
    Public SEPARATOR1 As String = "\\$$\\"

    Public Property SessionNodeGUID() As String
        Get
            Return _SessionNodeGUID
        End Get
        Set(ByVal value As String)
            _SessionNodeGUID = value
        End Set
    End Property

    Public Property TimerWorkflowGUID() As String
        Get
            Return _TimerWorkflowGUID
        End Get
        Set(ByVal value As String)
            _TimerWorkflowGUID = value
        End Set
    End Property

    Public Property ExecutionThreadGUID() As String
        Get
            Return _ExecutionThreadGUID
        End Get
        Set(ByVal value As String)
            _ExecutionThreadGUID = value
        End Set
    End Property

    Public Property SessionGUID() As String
        Get
            Return _SessionGUID
        End Get
        Set(ByVal value As String)
            _SessionGUID = value
        End Set
    End Property

    Public Function Serialize() As String
        Select Case _MessageType
            Case MessageTypes.MT_NEWSESSION
                Return CStr(_MessageType) + SEPARATOR1 + _SessionGUID
            Case MessageTypes.MT_NEWTIMERFLOW, MessageTypes.MT_REMTIMERFLOW
                Return CStr(_MessageType) + SEPARATOR1 + _TimerWorkflowGUID
            Case MessageTypes.MT_NODEACTION
                Return CStr(_MessageType) + SEPARATOR1 + _ExecutionThreadGUID + SEPARATOR1 + _SessionGUID + SEPARATOR1 + _SessionNodeGUID
            Case MessageTypes.MT_TERMSESSION, MessageTypes.MT_RESTSESSION
                Return CStr(_MessageType) + SEPARATOR1 + _SessionGUID
        End Select

    End Function

    Public Sub Deserialize(ByVal Item As String)
        _MessageType = MessageTypes.MT_UNKNOWN
        Dim arrSplits() As String = Split(Item, SEPARATOR1)
        If UBound(arrSplits) >= 0 Then
            _MessageType = GlobalFunctions.FormatInteger(arrSplits(0), MessageTypes.MT_UNKNOWN)
            If _MessageType = MessageTypes.MT_UNKNOWN Then Exit Sub
            Select Case _MessageType
                Case MessageTypes.MT_NEWSESSION, MessageTypes.MT_TERMSESSION, MessageTypes.MT_RESTSESSION
                    If UBound(arrSplits) >= 1 Then
                        _SessionGUID = arrSplits(1)
                    Else
                        Exit Sub
                    End If
                Case MessageTypes.MT_NEWTIMERFLOW, MessageTypes.MT_REMTIMERFLOW
                    If UBound(arrSplits) >= 1 Then
                        _TimerWorkflowGUID = arrSplits(1)
                    Else
                        Exit Sub
                    End If
                Case MessageTypes.MT_NODEACTION
                    If UBound(arrSplits) >= 3 Then
                        _ExecutionThreadGUID = arrSplits(1)
                        _SessionGUID = arrSplits(2)
                        _SessionNodeGUID = arrSplits(3)
                    Else
                        Exit Sub
                    End If
            End Select
        End If
    End Sub

    Public Sub SendMessage()
        Dim strText As String = Serialize()
        Try
            _w3mq.Send(strText)
            _w3mq.Close()
        Catch ex As Exception
            _lasterrormsg = "Send MSMQ message error : " & ex.ToString
        End Try

    End Sub

    Public Property MessageType() As MessageTypes
        Get
            Return _MessageType
        End Get
        Set(ByVal value As MessageTypes)
            _MessageType = value

        End Set
    End Property

    Public ReadOnly Property W3MQ() As MessageQueue
        Get
            Return _w3mq
        End Get
    End Property

    Public Function GetLastErrorMsg() As String
        Return _lasterrormsg
    End Function

    Private Sub SetupMSMQ(ByVal QueueLocation As String)
        Try
            ' mqPath assigning

            Dim W3mqPath As String = QueueLocation

            'Checking if that queue is there in private queue or not.
            'If it is not there, then creates that queue
            Try
                If (Not MessageQueue.Exists(W3mqPath)) Then
                    MessageQueue.Create(W3mqPath)
                    logger.Debug("MessageQueue.Create(W3mqPath: " + GlobalFunctions.FormatData(W3mqPath) + ")")
                End If
            Catch ex As Exception
                logger.Error(ex, "QueueLocation: " + GlobalFunctions.FormatData(W3mqPath))
            End Try

            'Use that queue and send the entered text into that queue using send method
            _w3mq = New MessageQueue(W3mqPath)
        Catch ex As Exception
            _lasterrormsg = "Setup MSMQ Error : " & ex.ToString
            logger.Error(ex, "Setup MSMQ Error")
        End Try

    End Sub

    Public Sub New(ByVal QueueLocation As String)
        SetupMSMQ(QueueLocation)
    End Sub
End Class

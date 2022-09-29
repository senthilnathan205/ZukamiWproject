Imports Microsoft.VisualBasic
Imports System.Data

Public Class AggregateClass
    Public Class AggregateRow
        Public Class AggregateColumn
            Private _Caption As String = ""
            Private _Value As String = ""
            Private _RetrievedValue As String = ""


            Public Property RetrievedValue() As String
                Get
                    Return _RetrievedValue
                End Get
                Set(ByVal value As String)
                    _RetrievedValue = value
                End Set
            End Property

            Public Property Caption() As String
                Get
                    Return _caption
                End Get
                Set(ByVal value As String)
                    _caption = value
                End Set
            End Property

            Public Property Value() As String
                Get
                    Return _Value
                End Get
                Set(ByVal value As String)
                    _Value = value
                End Set
            End Property
        End Class

        Private _rowHeader As String
        Private _rowID As String
        Private _Columns As New Collection
        Private _associatedgroup As String = ""

        Public Property AssociatedGroup() As String
            Get
                Return _associatedgroup
            End Get
            Set(ByVal value As String)
                _associatedgroup = value
            End Set
        End Property

        Public ReadOnly Property Columns() As Collection
            Get
                Return _Columns
            End Get
        End Property

        Public Property RowHeader() As String
            Get
                Return _rowHeader
            End Get
            Set(ByVal value As String)
                _rowHeader = value
            End Set
        End Property

        Public Property RowID() As String
            Get
                Return _rowID
            End Get
            Set(ByVal value As String)
                _rowID = value
            End Set
        End Property

        Public Sub New(ByRef Row As DataRow)
            Dim _rowfields As String = GlobalFunctions.FormatData(Row.Item("RowFields"))
            _rowHeader = GlobalFunctions.FormatData(Row.Item("RowHeaderText"))
            _rowID = GlobalFunctions.FormatData(Row.Item("RowID"))
            Try
                _associatedgroup = GlobalFunctions.FormatData(Row.Item("AssociatedGroup"))
            Catch ex As Exception

            End Try

            Dim arrsplits() As String = Split(_rowfields, GlobalFunctions.GLOBSEPARATOR_AGGREGATEFIELDS)
            Dim _counter2 As Integer
            For _counter2 = 0 To UBound(arrsplits)
                If Len(arrsplits(_counter2)) > 0 Then
                    Dim arrsplits2() As String = Split(arrsplits(_counter2), GlobalFunctions.GLOBSEPARATOR_AGGREGATEFIELDS2)
                    If UBound(arrsplits2) = 1 Then
                        Dim _ColCaption As String = arrsplits2(0)
                        Dim _Value As String = arrsplits2(1)

                        Dim _col As New AggregateColumn
                        _col.Caption = _ColCaption
                        _col.Value = _Value
                        If Len(_col.Caption) > 0 Then
                            If _Columns.Contains(_col.Caption) = False Then
                                _Columns.Add(_col, _col.Caption)
                            End If
                        End If
                    End If
                End If
            Next _counter2

        End Sub
    End Class

    Private _rows As New Collection

    Public ReadOnly Property Rows() As Collection
        Get
            Return _rows
        End Get
    End Property

    Public Sub New(ByVal Value As String)
        Dim _Data As New DataSet
        _Data.Tables.Add(New DataTable)
        _Data.Tables(0).Columns.Add(New DataColumn("RowID", GetType(String)))
        _Data.Tables(0).Columns.Add(New DataColumn("RowHeaderText", GetType(String)))
        _Data.Tables(0).Columns.Add(New DataColumn("RowFields", GetType(String)))
        _Data.Tables(0).Columns.Add(New DataColumn("AssociatedGroup", GetType(String)))
        GlobalFunctions.DeparseAggregate(_Data, Value)
        Dim _counter As Integer
        For _counter = 0 To _Data.Tables(0).Rows.Count - 1
            Dim _ar As New AggregateRow(_Data.Tables(0).Rows(_counter))
            _rows.Add(_ar)
        Next _counter
    End Sub
End Class

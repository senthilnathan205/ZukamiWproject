Imports Microsoft.VisualBasic

Public Class ChartDataMobile
    Public numberSuffix As String = ""
    Public xAxisName As String = ""
    Public yAxisName As String = ""

    Public dataList As List(Of Double)
    Public dataList2 As List(Of Double)
    Public dataList3 As List(Of Double)
    Public dataLabel As String = ""
    Public dataLabel2 As String = ""
    Public dataLabel3 As String = ""
    Public colorList As List(Of String)
    Public colorList2 As List(Of String)
    Public colorList3 As List(Of String)
    Public labelList As List(Of String)
    Public Sub New()
        dataList = New List(Of Double)
        dataList2 = New List(Of Double)
        dataList3 = New List(Of Double)
        colorList = New List(Of String)
        colorList2 = New List(Of String)
        colorList3 = New List(Of String)
        labelList = New List(Of String)
    End Sub

    Public Sub AddItem(name As String, value As String, color As String)
        labelList.Add(name)
        dataList.Add(GlobalFunctions.FormatDouble(value))
        colorList.Add("#" + color)
    End Sub
End Class

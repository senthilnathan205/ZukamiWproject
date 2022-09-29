Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Drawing

Public Class ColorPalette
    Public Shared Sub CreateColors(ByRef webobj As ZukamiLib.WebSession, ByRef arrColors() As String, ByVal ColorCount As Integer)
        Try

        
            Dim _set As DataSet = webobj.ColorPalettes_Get("Default Palette")
            If _set.Tables(0).Rows.Count = 0 Then Exit Sub

            Dim _color As String = GlobalFunctions.FormatData(_set.Tables(0).Rows(0).Item("Colors"))
            Dim _index As Long = 0
            Dim _max As Long = 0
            Dim _ceiling As Long = 0
            Dim arrsplits() As String = Split(_color, ",")
            _max = UBound(arrsplits)
            _ceiling = Math.Min(_max, ColorCount - 1)

            ReDim arrColors(ColorCount - 1)
            _index = 0
            Dim _counter As Integer
            For _counter = 0 To _ceiling
                arrColors(_counter) = arrsplits(_counter)
            Next

            For _counter = _ceiling + 1 To ColorCount - 1
                arrColors(_counter) = generateRandomColor(Color.White)
            Next
        Catch ex As Exception

        End Try
    End Sub

    Public Shared Function generateRandomColor(ByVal mix As Color) As String
        Dim red As Integer = Int((255 - 0 + 1) * Rnd()) + 0
        Dim green As Integer = Int((255 - 0 + 1) * Rnd()) + 0
        Dim blue As Integer = Int((255 - 0 + 1) * Rnd()) + 0

        ' mix the color
        red = (red + mix.R) \ 2
        green = (green + mix.G) \ 2
        blue = (blue + mix.B) \ 2

        Return Hex(red) & Hex(green) & Hex(blue)
    End Function

    Public Shared Sub Initialize()
        Randomize()
    End Sub
End Class

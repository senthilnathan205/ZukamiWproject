Imports Microsoft.VisualBasic
Imports System.Data

Public Class UserDetails
    Private _userSet As DataRow = Nothing
    Private _roleSet As DataSet = Nothing

    Public ReadOnly Property Username() As String
        Get
            Return GlobalFunctions.FormatData(_userSet.Item("Username"))
        End Get
    End Property

    Public ReadOnly Property Fullname() As String
        Get
            Return GlobalFunctions.FormatData(_userSet.Item("Fullname"))
        End Get
    End Property

    Public ReadOnly Property ADCredential() As String
        Get
            Return GlobalFunctions.FormatData(_userSet.Item("ADCredential"))
        End Get
    End Property

    Public ReadOnly Property Email() As String
        Get
            Return GlobalFunctions.FormatData(_userSet.Item("Email"))
        End Get
    End Property

    Public ReadOnly Property Title() As String
        Get
            Return GlobalFunctions.FormatData(_userSet.Item("Title"))
        End Get
    End Property

    Public ReadOnly Property PhoneNumber() As String
        Get
            Return GlobalFunctions.FormatData(_userSet.Item("PhoneNumber"))
        End Get
    End Property

    Public ReadOnly Property Description() As String
        Get
            Return GlobalFunctions.FormatData(_userSet.Item("Description"))
        End Get
    End Property

    Public ReadOnly Property Department() As String
        Get
            Return GlobalFunctions.FormatData(_userSet.Item("DepartmentID"))
        End Get
    End Property

    Public ReadOnly Property DepartmentName() As String
        Get
            Return GlobalFunctions.FormatData(_userSet.Item("Department"))
        End Get
    End Property

    Public ReadOnly Property Initials() As String
        Get
            Return GlobalFunctions.FormatData(_userSet.Item("Initials"))
        End Get
    End Property

    Public ReadOnly Property UserID() As String
        Get
            Return GlobalFunctions.FormatData(_userSet.Item("UserID"))
        End Get
    End Property

    Public ReadOnly Property Locality() As String
        Get
            Return GlobalFunctions.FormatData(_userSet.Item("Locality"))
        End Get
    End Property

    Public ReadOnly Property Language() As String
        Get
            Return GlobalFunctions.FormatData(_userSet.Item("Language"))
        End Get
    End Property

    Public ReadOnly Property Supervisor() As String
        Get
            Return GlobalFunctions.FormatData(_userSet.Item("Supervisor"))
        End Get
    End Property

    Public ReadOnly Property Locked() As Boolean
        Get
            Return GlobalFunctions.FormatBoolean(_userSet.Item("Locked"))
        End Get
    End Property


    Public ReadOnly Property Expired() As Boolean
        Get
            Return GlobalFunctions.FormatBoolean(_userSet.Item("Expired"))
        End Get
    End Property

    Public ReadOnly Property Disabled() As Boolean
        Get
            Return GlobalFunctions.FormatBoolean(_userSet.Item("Disabled"))
        End Get
    End Property

    Public ReadOnly Property Roles() As String
        Get
            If _roleSet Is Nothing = False Then
                Dim _mainstring As String = ""
                Dim _counter As Integer
                For _counter = 0 To _roleSet.Tables(0).Rows.Count - 1
                    If Len(_mainstring) > 0 Then _mainstring += ", "
                    _mainstring += GlobalFunctions.FormatData(_roleSet.Tables(0).Rows(_counter).Item("Groupname"))
                Next _counter
                Return _mainstring
            Else
                Return ""
            End If
        End Get
    End Property

    Public ReadOnly Property RoleSet() As DataSet
        Get
            Return _roleSet
        End Get
    End Property

    Public Sub New(ByRef USet As DataRow, ByVal RoleSet As DataSet)
        _userSet = USet
        _roleSet = RoleSet
    End Sub
End Class

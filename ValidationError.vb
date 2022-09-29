Imports Microsoft.VisualBasic
Public Class ValidationError
    Implements IValidator
    Private Sub New(ByVal message As String)
        ErrorMessage = message
        IsValid = False
    End Sub

    Private _ErrorMessage As String
    Private _IsValid As Boolean


    Public Sub Validate() Implements System.Web.UI.IValidator.Validate

    End Sub

    Public Shared Sub Display2(ByVal message As String)
        Display(GlobalFunctions.DbResT(message))
    End Sub
    Public Shared Sub Display(ByVal message As String)
        Dim currentPage As Page = TryCast(HttpContext.Current.Handler, Page)
        currentPage.Validators.Add(New ValidationError(message))
    End Sub

    Public Shared Function ErrorsCount() As Integer
        Dim currentPage As Page = TryCast(HttpContext.Current.Handler, Page)
        Return currentPage.Validators.Count
    End Function

    Public Property ErrorMessage() As String Implements System.Web.UI.IValidator.ErrorMessage
        Get
            Return _ErrorMessage
        End Get
        Set(ByVal value As String)
            _ErrorMessage = value
        End Set
    End Property

    Public Property IsValid() As Boolean Implements System.Web.UI.IValidator.IsValid
        Get
            Return _IsValid
        End Get
        Set(ByVal value As Boolean)
            _IsValid = value
        End Set
    End Property

  
End Class
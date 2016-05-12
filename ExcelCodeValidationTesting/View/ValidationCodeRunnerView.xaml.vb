Imports ExcelCodeValidationTesting.ViewModel

Class ValidationCodeRunnerView
    Dim ReadOnly _viewModel As ValidationCodeRunnerViewModel

    Public Sub New()
        _viewModel = New ValidationCodeRunnerViewModel()
        DataContext = _viewModel
        ' This call is required by the designer.
        InitializeComponent()


    End Sub
End Class
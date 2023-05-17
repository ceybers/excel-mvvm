Attribute VB_Name = "SampleViewModel"
'@Folder("ViewModel")
Option Explicit

Public Function GetSampleViewModel(ByVal Context As AppContext) As ViewModel
    With New ViewModel
        Set .Context = Context
        .FirstName = "John"
        .LastName = "Doe"
        .DateOfBirth = DateSerial(1984, 1, 1)
        .Foo = "Lorem ipsum"
        .Bar = 42
        .Size = "Small"
        Set GetSampleViewModel = .Self
    End With
End Function

Public Sub TestViewModel()
    Dim ctx As AppContext
    Set ctx = New AppContext
    
    Dim vm As ViewModel
    Set vm = GetSampleViewModel(ctx)
    
    vm.Foo = "TestViewModel"
    
    Stop
End Sub

Attribute VB_Name = "SampleViewModel"
'@Folder("ViewModel")
Option Explicit

Public Function GetSampleViewModel() As ViewModel
    With New ViewModel
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
    Dim vm As ViewModel
    Set vm = GetSampleViewModel
    vm.Foo = "TestViewModel"
    Stop
End Sub

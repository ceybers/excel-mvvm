Attribute VB_Name = "TestViewModel"
'@Folder("ViewModel")
Option Explicit

Public Function GetTestViewModel() As ExampleViewModel
    With New ExampleViewModel
        .FirstName = "John"
        .LastName = "Doe"
        .DateOfBirth = DateSerial(1984, 1, 1)
        .Foo = "Lorem ipsum"
        .Bar = 42
        Set GetTestViewModel = .Self
    End With
End Function

Public Sub TestViewModel()
    Dim vm As ExampleViewModel
    Set vm = GetTestViewModel
    vm.Foo = "TestViewModel"
    Stop
End Sub

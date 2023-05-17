Attribute VB_Name = "SampleViewModel"
'@Folder "MVVM.ViewModels"
Option Explicit

Public Function GetSampleViewModel(ByVal Context As IAppContext) As SomeViewModel
    With New SomeViewModel
        Set .Context = Context
        .FirstName = "John"
        .LastName = "Doe"
        .DateOfBirth = DateSerial(1984, 1, 1)
        .Foo = "Lorem ipsum"
        .Bar = 42
        .Size = "Small"
        Set .SizeOptions = GetSampleSizeOptions
        .IsFoobar = False
        .IsFoobarCaption = "Is Foo Bar Caption"
        Set GetSampleViewModel = .Self
        Set .TestMsgboxCommand = New TestMsgboxCommand
        
        .InitializeCommands
    End With
End Function

Private Function GetSampleSizeOptions() As Scripting.Dictionary
    Dim Result As Scripting.Dictionary
    Set Result = New Scripting.Dictionary
    With Result
        .Add Key:="S", Item:="Small"
        .Add Key:="M", Item:="Medium"
        .Add Key:="L", Item:="Large"
    End With
    Set GetSampleSizeOptions = Result
End Function

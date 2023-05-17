Attribute VB_Name = "SampleViewModel"
'@Folder("ViewModel")
Option Explicit

Public Function GetSampleViewModel(ByVal Context As AppContext) As ViewModel
    Dim mSizeOptions As Scripting.Dictionary
    Set mSizeOptions = New Scripting.Dictionary
    mSizeOptions.Add Key:="S", Item:="Small"
    mSizeOptions.Add Key:="M", Item:="Medium"
    mSizeOptions.Add Key:="L", Item:="Large"
    
    With New ViewModel
        Set .Context = Context
        .FirstName = "John"
        .LastName = "Doe"
        .DateOfBirth = DateSerial(1984, 1, 1)
        .Foo = "Lorem ipsum"
        .Bar = 42
        .Size = "Small"
        Set .SizeOptions = mSizeOptions
        .IsFoobar = False
        .IsFoobarCaption = "Is Foo Bar Caption"
        Set GetSampleViewModel = .Self
        Set .TestMsgboxCommand = New TestMsgboxCommand
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

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestVMSubCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A Command specific to a type of ViewModel."
'@ModuleDescription "A Command specific to a type of ViewModel."
'@Folder "MVVM.Commands"
Option Explicit
Implements ICommand

Private mViewModel As SomeViewModel ' This is ViewModel specific
Private mAppContext As IAppContext

Public Property Get ViewModel() As SomeViewModel
    Set ViewModel = mViewModel
End Property

Public Property Set ViewModel(ByVal vNewValue As SomeViewModel)
    Set mViewModel = vNewValue
End Property

Public Property Get Context() As IAppContext
    Set Context = mAppContext
End Property

Public Property Set Context(ByVal vNewValue As IAppContext)
    Set mAppContext = vNewValue
End Property

Private Property Get ICommand_Caption() As String
    ICommand_Caption = "Test VM.DoCmd"
End Property

Private Function ICommand_CanExecute() As Boolean
    ICommand_CanExecute = True
End Function

Private Sub ICommand_Execute()
    mViewModel.DoSomethingByCommand
    'RaiseEvent Executed
End Sub

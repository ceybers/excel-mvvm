VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestVMSubCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Commands")
Option Explicit
Implements ICommand

Public Event CanExecuteChanged()
Public Event Executed()

Private mViewModel As ViewModel
Private mAppContext As AppContext

Private Property Get ICommand_Caption() As String
    ICommand_Caption = "Test VM.DoCmd"
End Property

Private Function ICommand_CanExecute() As Boolean
    ICommand_CanExecute = True
End Function

Private Sub ICommand_Execute()
    mViewModel.DoSomethingByCommand
    RaiseEvent Executed
End Sub

Public Property Get ViewModel() As ViewModel
    Set ViewModel = mViewModel
End Property

Public Property Get Context() As AppContext
    Set Context = mAppContext
End Property

Public Property Set ViewModel(ByVal vNewValue As ViewModel)
    Set mViewModel = vNewValue
End Property

Public Property Set Context(ByVal vNewValue As AppContext)
    Set mAppContext = vNewValue
End Property

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestMsgboxCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Basic test command that does not reference an AppContext, View, nor ViewModel."
'@ModuleDescription "Basic test command that does not reference an AppContext, View, nor ViewModel."
'@Folder "Commands"
Option Explicit
Implements ICommand

Private Property Get ICommand_Caption() As String
    ICommand_Caption = "Test MsgBox"
End Property

Private Function ICommand_CanExecute() As Boolean
    ICommand_CanExecute = True
End Function

Private Sub ICommand_Execute()
    MsgBox "Test Command"
    'RaiseEvent Executed
End Sub

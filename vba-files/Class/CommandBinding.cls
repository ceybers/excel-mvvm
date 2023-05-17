VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder MVVM.Infrastructure.Bindings.CommandBindings
Option Explicit
Implements ICommandBinding

Private Const DO_DEBUG As Boolean = False

Private mAppContext As IAppContext
Private mSource As IViewModel
'@MemberAttribute VB_VarHelpID, -1
Private mCommand As ICommand
Attribute mCommand.VB_VarHelpID = -1
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents mTarget As MSForms.CommandButton
Attribute mTarget.VB_VarHelpID = -1

Private Property Get ICommandBinding_Context() As IAppContext
    Set ICommandBinding_Context = mAppContext
End Property

Private Property Get ICommandBinding_Source() As IViewModel
    Set ICommandBinding_Source = mSource
End Property

Private Property Get ICommandBinding_Command() As ICommand
    Set ICommandBinding_Command = mCommand
End Property

Private Property Get ICommandBinding_Target() As Object
    Set ICommandBinding_Target = mTarget
End Property

Public Sub Create(ByVal AppContext As IAppContext, ByVal Source As IViewModel, ByVal Command As ICommand, ByVal Target As MSForms.CommandButton)
    Set mAppContext = AppContext
    Set mSource = Source
    Set mCommand = Command
    Set mTarget = Target
    
    mTarget.Caption = mCommand.Caption
    mTarget.Enabled = mCommand.CanExecute
End Sub

Private Sub mTarget_Click()
    If DO_DEBUG Then Debug.Print "CommandBinding mTarget_Click()"
    mCommand.Execute
End Sub

' Cannot RaiseEvents at ICommand interface from the concrete class TestMsgBoxCommand etc.
'Private Sub mCommand_CanExecuteChanged()
'    If DO_DEBUG Then Debug.Print "CommandBinding cmd_CanExecuteChanged()"
'    mTarget.Enabled = mCommand.CanExecute
'End Sub

'Private Sub mCommand_Executed()
'    If DO_DEBUG Then Debug.Print "CommandBinding cmd_Executed()"
'End Sub

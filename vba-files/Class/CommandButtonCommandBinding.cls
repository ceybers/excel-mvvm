VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandButtonCommandBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Commands")
Option Explicit
Implements ICommandBinding

Private Const DO_DEBUG As Boolean = False

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents mSource As ViewModel
Attribute mSource.VB_VarHelpID = -1
Private mSourcePropertyPath As String
'@MemberAttribute VB_VarHelpID, -1
Private mCommand As ICommand
Attribute mCommand.VB_VarHelpID = -1
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents mTarget As MSForms.CommandButton
Attribute mTarget.VB_VarHelpID = -1

Public Sub Create(ByVal Source As ViewModel, ByVal SourcePropertyPath As String, ByVal Target As MSForms.CommandButton)
    Set mSource = Source
    mSourcePropertyPath = SourcePropertyPath
    
    Set mCommand = CallByName(mSource, SourcePropertyPath, VbGet)
    
    Set mTarget = Target
    mTarget.Caption = mCommand.Caption
    mTarget.Enabled = mCommand.CanExecute
End Sub

Private Sub mTarget_Click()
    If DO_DEBUG Then Debug.Print "CommandButtonBinding mTarget_Click()"
    mCommand.Execute
End Sub

Private Sub mCommand_CanExecuteChanged()
    If DO_DEBUG Then Debug.Print "CommandButtonBinding cmd_CanExecuteChanged()"
    mTarget.Enabled = mCommand.CanExecute
End Sub

Private Sub mCommand_Executed()
    If DO_DEBUG Then Debug.Print "CommandButtonBinding cmd_Executed()"
End Sub

Private Property Get ICommandBinding_Source() As Object
    Set ICommandBinding_Source = mSource
End Property

Private Property Get ICommandBinding_SourcePropertyPath() As String
    ICommandBinding_SourcePropertyPath = mSourcePropertyPath
End Property

Private Property Get ICommandBinding_Target() As Object
    Set ICommandBinding_Target = mTarget
End Property

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestViewCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Commands")
Option Explicit
Implements ICommand

Public Event CanExecuteChanged()
Public Event Executed()

Private mView As ExampleView
Private mAppContext As AppContext

Private Property Get ICommand_Caption() As String
    ICommand_Caption = "Test VM.DoCmd"
End Property

Private Function ICommand_CanExecute() As Boolean
    ICommand_CanExecute = True
End Function

Private Sub ICommand_Execute()
    mView.DoSomething
    RaiseEvent Executed
End Sub

Public Property Get View() As IView
    Set View = mView
End Property

Public Property Set View(ByVal vNewValue As IView)
    Set mView = vNewValue
End Property

Public Property Get Context() As AppContext
    Set Context = mAppContext
End Property

Public Property Set Context(ByVal vNewValue As AppContext)
    Set mAppContext = vNewValue
End Property


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "AppContext"
Option Explicit

Private Type TBindingManager
    Commands As Collection
End Type
Private This As TBindingManager

Private Sub Class_Initialize()
    Set This.Commands = New Collection
End Sub

Public Sub BindCommand(ByVal Context As IAppContext, ByVal ViewModel As IViewModel, ByVal Command As ICommand, ByVal Control As Object)
    Dim NewBinding As CommandBinding
    Set NewBinding = New CommandBinding
    NewBinding.Create Context, ViewModel, Command, Control
    This.Commands.Add Item:=NewBinding
End Sub

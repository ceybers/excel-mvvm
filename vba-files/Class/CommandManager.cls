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
    'Strategy As CommandBindingStrategy
End Type
Private This As TBindingManager

Private Sub class_initialize()
    Set This.Commands = New Collection
    'Set This.Strategy = New CommandBindingStrategy
End Sub

'Public Sub BindCommand(ByVal ViewModel As ViewModel, ByVal CommandName As String, ByVal Control As MSForms.CommandButton)
'    Dim NewBinding As CommandButtonCommandBinding
'    Set NewBinding = This.Strategy.CreateBinding(ViewModel, CommandName, Control)
'    This.Commands.Add NewBinding
'End Sub

Public Sub BindCommand2(ByVal Context As AppContext, ByVal ViewModel As Object, ByVal Command As ICommand, ByVal Control As MSForms.CommandButton)
    Dim NewBinding As CommandBinding
    Set NewBinding = New CommandBinding
    NewBinding.Create Context, ViewModel, Command, Control
    This.Commands.Add Item:=NewBinding
End Sub

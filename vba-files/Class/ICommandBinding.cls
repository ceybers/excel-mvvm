VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICommandBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Interface
'@Folder "Abstractions"
Option Explicit

Public Property Get Context() As AppContext
End Property

Public Property Get Source() As Object ' ViewModel
End Property

Public Property Get Command() As ICommand ' Command
End Property

Public Property Get Target() As Object ' UserForm Control
End Property

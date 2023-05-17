VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandBindingStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Commands")
Option Explicit

Public Function CreateBinding(ByVal Source As IViewModel, ByVal SourcePropertyPath As String, ByVal Target As Control) As CommandButtonCommandBinding
    Dim NewBinding As CommandButtonCommandBinding
    Set NewBinding = New CommandButtonCommandBinding
    NewBinding.Create Source, SourcePropertyPath, Target
    Set CreateBinding = NewBinding
End Function


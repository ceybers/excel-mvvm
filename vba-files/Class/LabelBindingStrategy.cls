VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LabelBindingStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PropertyBindings"
Option Explicit
Implements IBindingStrategy

Private Function IBindingStrategy_CreateBinding(ByVal Source As IViewModel, ByVal SourcePropertyPath As String, ByVal Target As Control, Optional ByVal TargetProperty As String) As IPropertyBinding
    Dim NewBinding As LabelPropertyBinding
    Set NewBinding = New LabelPropertyBinding
    NewBinding.Create Source, SourcePropertyPath, Target, TargetProperty
    Set IBindingStrategy_CreateBinding = NewBinding
End Function
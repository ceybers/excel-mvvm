VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BindingManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Bindings")
Option Explicit

Private Type TBindingManager
    PropertyBindings As Collection
    BindingStrategies As Scripting.Dictionary
End Type
Private This As TBindingManager

Private Sub Class_Initialize()
    Set This.PropertyBindings = New Collection
    Set This.BindingStrategies = New Scripting.Dictionary
    InitializeBindingStrategies
End Sub

Private Sub InitializeBindingStrategies()
    With This.BindingStrategies
        .Add Item:=New TextBoxBindingStrategy, Key:="MSForms.TextBox"
    End With
End Sub

Public Sub BindPropertyPath(ByVal Source As IViewModel, ByVal SourcePropertyPath As String, ByVal Target As Control, Optional ByVal TargetProperty As String)
    Dim BindingStrategy As IBindingStrategy
    Set BindingStrategy = This.BindingStrategies.Item(GetFullTypeName(Target))
    
    Dim NewBinding As IPropertyBinding
    Set NewBinding = BindingStrategy.CreateBinding(Source, SourcePropertyPath, Target, TargetProperty)
    This.PropertyBindings.Add Item:=NewBinding
End Sub

Private Function GetFullTypeName(ByVal Target As Control) As String
    GetFullTypeName = TypeName(Target)
    
    ' TODO Check how to handle MScomctllib
    Debug.Assert InStr(GetFullTypeName, "ListView") = 0
    
    If TypeOf Target Is MSForms.Control Then
        GetFullTypeName = "MSForms." & GetFullTypeName
    End If
End Function

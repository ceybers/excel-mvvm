VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BindingManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder MVVM.Infrastructure.Bindings
Option Explicit

Private Type TBindingManager
    PropertyBindings As Collection
    BindingStrategies As Scripting.Dictionary
End Type
Private This As TBindingManager

Private Sub class_initialize()
    Set This.PropertyBindings = New Collection
    
    Set This.BindingStrategies = New Scripting.Dictionary
    InitializeBindingStrategies
End Sub

Private Sub InitializeBindingStrategies()
    With This.BindingStrategies
        .Add Item:=New TextBoxBindingStrategy, Key:="MSForms.TextBox"
        .Add Item:=New LabelBindingStrategy, Key:="MSForms.Label"
        .Add Item:=New CheckBoxBindingStrategy, Key:="MSForms.CheckBox"
        .Add Item:=New OptionButtonBindingStrategy, Key:="MSForms.OptionButton"
        .Add Item:=New ComboBoxBindingStrategy, Key:="MSForms.ComboBox"
        .Add Item:=New ListViewBindingStrategy, Key:="MSComctlLib.ListView"
        .Add Item:=New TreeViewBindingStrategy, Key:="MSComctlLib.TreeView"
        .Add Item:=New CommandButtonBindingStrategy, Key:="MSForms.CommandButton"
    End With
End Sub

Public Sub BindPropertyPath(ByVal Source As IViewModel, ByVal SourcePropertyPath As String, ByVal Target As Object, Optional ByVal TargetProperty As String)
    Dim BindingStrategy As IBindingStrategy
    
    On Error GoTo NoStrategyFound
    Set BindingStrategy = This.BindingStrategies.Item(GetFullTypeName(Target))
    On Error GoTo 0
    
    Dim NewBinding As IPropertyBinding
    Set NewBinding = BindingStrategy.CreateBinding(Source, SourcePropertyPath, Target, TargetProperty)
    This.PropertyBindings.Add Item:=NewBinding
    
    Dim SourceNotifier As INotifyPropertyChanged
    Set SourceNotifier = Source
    If Not SourceNotifier Is Nothing And TypeOf NewBinding Is IHandlePropertyChanged Then
        SourceNotifier.RegisterHandler NewBinding
    End If
    
    Exit Sub
    
NoStrategyFound:
    Err.Clear
    MsgBox "This control hasn't been implemented yet." & vbCrLf & GetFullTypeName(Target)
End Sub

Private Function GetFullTypeName(ByVal Target As Control) As String
    GetFullTypeName = TypeName(Target)
   
    Select Case GetFullTypeName
        Case "ListView4":
            GetFullTypeName = "MSComctlLib.ListView"
        Case "TreeView4":
            GetFullTypeName = "MSComctlLib.TreeView"
        Case Else:
            If TypeOf Target Is MSForms.Control Then
                GetFullTypeName = "MSForms." & GetFullTypeName
            End If
    End Select
End Function

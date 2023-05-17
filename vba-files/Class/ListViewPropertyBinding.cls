VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListViewPropertyBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule SetAssignmentWithIncompatibleObjectType
'@Folder "PropertyBindings"
Option Explicit
Implements IPropertyBinding
Implements IHandlePropertyChanged

Private Const DO_DEBUG As Boolean = False

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents lv As MSComctlLib.ListView
Attribute lv.VB_VarHelpID = -1

Private Type TState
    BindingMode As BindingMode
    ViewModel As SomeViewModel
    SourcePropertyPath As String
    TargetProperty As String
    OneTimeBoundCounter As Long
End Type
Private This As TState

Private Property Get IPropertyBinding_Mode() As BindingMode
    IPropertyBinding_Mode = This.BindingMode
End Property

Private Property Get IPropertyBinding_Source() As IViewModel
    Set IPropertyBinding_Source = vm
End Property

Private Property Get IPropertyBinding_SourcePropertyPath() As String
    IPropertyBinding_SourcePropertyPath = This.SourcePropertyPath
End Property

Private Property Get IPropertyBinding_Target() As Object
    Set IPropertyBinding_Target = lv
End Property

Private Property Get IPropertyBinding_TargetProperty() As String
    IPropertyBinding_TargetProperty = This.TargetProperty
End Property

Public Sub Create(ByVal Source As IViewModel, ByVal SourcePropertyPath As String, ByVal Target As Control, Optional ByVal TargetProperty As String)
    If DO_DEBUG Then Debug.Print "ListViewPropertyBinding Create(vm, "; SourcePropertyPath; ", "; Target.Name; ")"
    Set This.ViewModel = Source
    This.SourcePropertyPath = SourcePropertyPath
    Set lv = Target
    This.TargetProperty = TargetProperty
    
    If TargetProperty = "SelectedItem" Then
        This.BindingMode = TwoWayBinding
    Else
        This.BindingMode = OneWayBinding
    End If
    
    UpdateTarget
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If DO_DEBUG Then Debug.Print "ListViewPropertyBinding lv_ItemClick()"
    UpdateSource
End Sub

Private Sub IHandlePropertyChanged_HandlePropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    If PropertyName <> This.SourcePropertyPath Then Exit Sub
    If DO_DEBUG Then Debug.Print "ListViewPropertyBinding vm_PropertyChanged()"
    UpdateTarget
End Sub

Private Sub UpdateTarget()
    If CanUpdateTarget Then
        If This.TargetProperty = "SelectedItem" Then
            UpdateListViewSelectedItem
        ElseIf This.TargetProperty = "ListItems" Then
            UpdateListViewListItems
        Else
            Debug.Print "Not yet implemented"
            Debug.Assert False
        End If
    End If
    
    If This.BindingMode = OneTimeBinding Then This.OneTimeBoundCounter = This.OneTimeBoundCounter + 1
End Sub

Private Sub UpdateSource()
    If This.BindingMode = TwoWayBinding Or This.BindingMode = OneWayToSource Then
        If This.TargetProperty = "SelectedItem" Then
            Dim vNewValue As Variant
            vNewValue = CallByName(lv, This.TargetProperty, VbGet)
            CallByName This.ViewModel, This.SourcePropertyPath, VbLet, vNewValue
        Else
            Debug.Assert False
        End If
    End If
End Sub

Private Function CanUpdateTarget() As Boolean
    If This.BindingMode = TwoWayBinding Then CanUpdateTarget = True
    If This.BindingMode = OneWayBinding Then CanUpdateTarget = True
    If This.BindingMode = OneTimeBinding And This.OneTimeBoundCounter = 0 Then CanUpdateTarget = True
End Function

Private Sub UpdateListViewSelectedItem()
    Dim vNewValue As String
    vNewValue = CallByName(This.ViewModel, This.SourcePropertyPath, VbGet)
    
    Dim ListItem As ListItem
    For Each ListItem In lv.ListItems
        If ListItem.Text = vNewValue Then
            Set lv.SelectedItem = ListItem
            Exit Sub
        End If
    Next ListItem
End Sub

Private Sub UpdateListViewListItems()
    If DO_DEBUG Then Debug.Print "ListViewPropertyBinding UpdateListViewListItems()"
    
    lv.ListItems.Clear
    
    Dim NewList As Scripting.Dictionary
    Set NewList = CallByName(This.ViewModel, This.SourcePropertyPath, VbGet)
    
    Dim Key As Variant
    For Each Key In NewList.Keys
        lv.ListItems.Add Key:=Key, Text:=NewList.Item(Key)
    Next Key
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TreeViewPropertyBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule SetAssignmentWithIncompatibleObjectType
'@Folder "PropertyBindings"
Option Explicit
Implements IPropertyBinding

Private Const DO_DEBUG As Boolean = True

Private mBindingMode As BindingMode
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents vm As ViewModel
Attribute vm.VB_VarHelpID = -1
Private mSourcePropertyPath As String
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents tv As MSComctlLib.TreeView
Attribute tv.VB_VarHelpID = -1
Private mTargetProperty As String
Private mOneTimeBoundCounter As Long

Private Property Get IPropertyBinding_Mode() As BindingMode
    IPropertyBinding_Mode = mBindingMode
End Property

Private Property Get IPropertyBinding_Source() As Object
    Set IPropertyBinding_Source = vm
End Property

Private Property Get IPropertyBinding_SourcePropertyPath() As String
    IPropertyBinding_SourcePropertyPath = mSourcePropertyPath
End Property

Private Property Get IPropertyBinding_Target() As Object
    Set IPropertyBinding_Target = tv
End Property

Private Property Get IPropertyBinding_TargetProperty() As String
    IPropertyBinding_TargetProperty = mTargetProperty
End Property

Public Sub Create(ByVal Source As IViewModel, ByVal SourcePropertyPath As String, ByVal Target As Control, Optional ByVal TargetProperty As String)
    If DO_DEBUG Then Debug.Print "TreeViewPropertyBinding Create(vm, "; SourcePropertyPath; ", "; Target.Name; ")"
    Set vm = Source
    mSourcePropertyPath = SourcePropertyPath
    Set tv = Target
    mTargetProperty = TargetProperty
    
    If TargetProperty = "SelectedItem" Then
        mBindingMode = TwoWayBinding
    Else
        mBindingMode = OneWayBinding
    End If
    
    UpdateTarget
End Sub

Private Sub tv_NodeClick(ByVal Node As MSComctlLib.Node)
    If DO_DEBUG Then Debug.Print "TreeViewPropertyBinding tv_NodeClick()"
    UpdateSource
End Sub

Private Sub vm_PropertyChanged(ByVal PropertyName As String)
    If PropertyName <> mSourcePropertyPath Then Exit Sub
    If DO_DEBUG Then Debug.Print "TreeViewPropertyBinding vm_PropertyChanged()"
    UpdateTarget
End Sub

Private Sub UpdateTarget()
    If CanUpdateTarget Then
        If mTargetProperty = "SelectedItem" Then
            UpdateTreeViewSelectedItem
        ElseIf mTargetProperty = "Nodes" Then
            UpdateTreeViewNodes
        Else
            Debug.Print "Not yet implemented"
            Debug.Assert False
        End If
    End If
    
    If mBindingMode = OneTimeBinding Then mOneTimeBoundCounter = mOneTimeBoundCounter + 1
End Sub

Private Sub UpdateSource()
    If mBindingMode = TwoWayBinding Or mBindingMode = OneWayToSource Then
        If mTargetProperty = "SelectedItem" Then
            Dim vnewValue As Node
            Set vnewValue = CallByName(tv, mTargetProperty, VbGet)
            CallByName vm, mSourcePropertyPath, VbLet, vnewValue
        Else
            Debug.Assert False
        End If
    End If
End Sub

Private Function CanUpdateTarget() As Boolean
    If mBindingMode = TwoWayBinding Then CanUpdateTarget = True
    If mBindingMode = OneWayBinding Then CanUpdateTarget = True
    If mBindingMode = OneTimeBinding And mOneTimeBoundCounter = 0 Then CanUpdateTarget = True
End Function

Private Sub UpdateTreeViewSelectedItem()
    Dim vnewValue As String
    vnewValue = CallByName(vm, mSourcePropertyPath, VbGet)
    
    Dim Node As Node
    For Each Node In tv.Nodes
        If Node.Text = vnewValue Then
            Set tv.SelectedItem = Node
            Exit Sub
        End If
    Next Node
End Sub

Private Sub UpdateTreeViewNodes()
    If DO_DEBUG Then Debug.Print "TreeViewPropertyBinding UpdateTreeViewNodes()"
    
    tv.Nodes.Clear
    
    Dim NewList As Scripting.Dictionary
    Set NewList = CallByName(vm, mSourcePropertyPath, VbGet)
    
    Dim Key As Variant
    For Each Key In NewList.Keys
        tv.Nodes.Add Key:=Key, Text:=NewList.Item(Key)
    Next Key
End Sub


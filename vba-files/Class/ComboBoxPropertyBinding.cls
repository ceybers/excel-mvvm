VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComboBoxPropertyBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule SetAssignmentWithIncompatibleObjectType
'@Folder "PropertyBindings"
Option Explicit
Implements IPropertyBinding

Private Const DO_DEBUG As Boolean = False

Private mBindingMode As BindingMode
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents vm As ViewModel
Attribute vm.VB_VarHelpID = -1
Private mSourcePropertyPath As String
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents cbo As MSForms.ComboBox
Attribute cbo.VB_VarHelpID = -1
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
    Set IPropertyBinding_Target = cbo
End Property

Private Property Get IPropertyBinding_TargetProperty() As String
    IPropertyBinding_TargetProperty = mTargetProperty
End Property

Public Sub Create(ByVal Source As IViewModel, ByVal SourcePropertyPath As String, ByVal Target As Control, Optional ByVal TargetProperty As String)
    If DO_DEBUG Then Debug.Print "ComboBoxPropertyBinding Create(vm, "; SourcePropertyPath; ", "; Target.Name; ")"
    Set vm = Source
    mSourcePropertyPath = SourcePropertyPath
    Set cbo = Target
    mTargetProperty = TargetProperty
    
    If TargetProperty = "Value" Then
        mBindingMode = TwoWayBinding
    Else
        mBindingMode = OneWayBinding
    End If
    
    UpdateTarget
End Sub

Private Sub cbo_Change()
    If DO_DEBUG Then Debug.Print "ComboBoxPropertyBinding cbo_Change()"
    UpdateSource
End Sub

Private Sub vm_PropertyChanged(ByVal PropertyName As String)
    If PropertyName <> mSourcePropertyPath Then Exit Sub
    If DO_DEBUG Then Debug.Print "ComboBoxPropertyBinding vm_PropertyChanged()"
    UpdateTarget
End Sub

Private Sub UpdateTarget()
    If CanUpdateTarget Then
        If mTargetProperty = "Value" Then
            UpdateComboBoxSelected
        ElseIf mTargetProperty = "List" Then
            UpdateComboBoxList
        Else
            Debug.Print "Not yet implemented"
            Debug.Assert False
        End If
    End If
    
    If mBindingMode = OneTimeBinding Then mOneTimeBoundCounter = mOneTimeBoundCounter + 1
End Sub

Private Sub UpdateSource()
    If mBindingMode = TwoWayBinding Or mBindingMode = OneWayToSource Then
        Dim vNewValue As Variant
        vNewValue = CallByName(cbo, mTargetProperty, VbGet)
        CallByName vm, mSourcePropertyPath, VbLet, vNewValue
    End If
End Sub

Private Function CanUpdateTarget() As Boolean
    If mBindingMode = TwoWayBinding Then CanUpdateTarget = True
    If mBindingMode = OneWayBinding Then CanUpdateTarget = True
    If mBindingMode = OneTimeBinding And mOneTimeBoundCounter = 0 Then CanUpdateTarget = True
End Function

Private Sub UpdateComboBoxSelected()
    Dim vNewValue As String
    vNewValue = CallByName(vm, mSourcePropertyPath, VbGet)
    CallByName cbo, mTargetProperty, VbLet, vNewValue
End Sub

Private Sub UpdateComboBoxList()
    If DO_DEBUG Then Debug.Print "ComboBoxPropertyBinding UpdateComboBoxList()"
    
    cbo.Clear
    
    Dim NewList As Scripting.Dictionary
    Set NewList = CallByName(vm, mSourcePropertyPath, VbGet)
    
    Dim Key As Variant
    For Each Key In NewList.Keys
        cbo.AddItem NewList.Item(Key)
    Next Key
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CheckBoxPropertyBinding"
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
Private WithEvents vm As SomeViewModel
Attribute vm.VB_VarHelpID = -1
Private mSourcePropertyPath As String
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents cb As MSForms.CheckBox
Attribute cb.VB_VarHelpID = -1
Private mTargetProperty As String
Private mOneTimeBoundCounter As Long

Private Property Get IPropertyBinding_Mode() As BindingMode
    IPropertyBinding_Mode = mBindingMode
End Property

Private Property Get IPropertyBinding_Source() As IViewModel
    Set IPropertyBinding_Source = vm
End Property

Private Property Get IPropertyBinding_SourcePropertyPath() As String
    IPropertyBinding_SourcePropertyPath = mSourcePropertyPath
End Property

Private Property Get IPropertyBinding_Target() As Object
    Set IPropertyBinding_Target = cb
End Property

Private Property Get IPropertyBinding_TargetProperty() As String
    IPropertyBinding_TargetProperty = mTargetProperty
End Property

Public Sub Create(ByVal Source As IViewModel, ByVal SourcePropertyPath As String, ByVal Target As Control, Optional ByVal TargetProperty As String)
    If DO_DEBUG Then Debug.Print "CheckBoxPropertyBinding Create(vm, "; SourcePropertyPath; ", "; Target.Name; ")"
    Set vm = Source
    mSourcePropertyPath = SourcePropertyPath
    Set cb = Target
    mTargetProperty = TargetProperty
    
    If TargetProperty = "Value" Then
        mBindingMode = TwoWayBinding
    Else
        mBindingMode = OneWayBinding
    End If
    
    UpdateTarget
End Sub

Private Sub cb_Change()
    If DO_DEBUG Then Debug.Print "CheckBoxPropertyBinding cb_Change()"
    UpdateSource
End Sub

Private Sub vm_PropertyChanged(ByVal PropertyName As String)
    If PropertyName <> mSourcePropertyPath Then Exit Sub
    If DO_DEBUG Then Debug.Print "CheckBoxPropertyBinding vm_PropertyChanged()"
    UpdateTarget
End Sub

Private Sub UpdateTarget()
    If CanUpdateTarget Then
        Dim vNewValue As String
        vNewValue = CallByName(vm, mSourcePropertyPath, VbGet)
        CallByName cb, mTargetProperty, VbLet, vNewValue
    End If
    
    If mBindingMode = OneTimeBinding Then mOneTimeBoundCounter = mOneTimeBoundCounter + 1
End Sub

Private Sub UpdateSource()
    If mBindingMode = TwoWayBinding Or mBindingMode = OneWayToSource Then
        Dim vNewValue As Variant
        vNewValue = CallByName(cb, mTargetProperty, VbGet)
        CallByName vm, mSourcePropertyPath, VbLet, vNewValue
    End If
End Sub

Private Function CanUpdateTarget() As Boolean
    If mBindingMode = TwoWayBinding Then CanUpdateTarget = True
    If mBindingMode = OneWayBinding Then CanUpdateTarget = True
    If mBindingMode = OneTimeBinding And mOneTimeBoundCounter = 0 Then CanUpdateTarget = True
End Function

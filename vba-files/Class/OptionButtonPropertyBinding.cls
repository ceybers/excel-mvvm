VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OptionButtonPropertyBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule SetAssignmentWithIncompatibleObjectType
'@Folder("Bindings")
Option Explicit
Implements IPropertyBinding

Private Const DO_DEBUG As Boolean = False

Private mBindingMode As BindingMode
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents vm As ViewModel
Attribute vm.VB_VarHelpID = -1
Private mSourcePropertyPath As String
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents opt As MSForms.OptionButton
Attribute opt.VB_VarHelpID = -1
Private mTargetProperty As String

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
    Set IPropertyBinding_Target = opt
End Property

Private Property Get IPropertyBinding_TargetProperty() As String
    IPropertyBinding_TargetProperty = mTargetProperty
End Property

Public Sub Create(ByVal Source As IViewModel, ByVal SourcePropertyPath As String, ByVal Target As Control, Optional ByVal TargetProperty As String)
    If DO_DEBUG Then Debug.Print "OptionButtonPropertyBinding Create(vm, "; SourcePropertyPath; ", "; Target.Name; ")"
    Set vm = Source
    mSourcePropertyPath = SourcePropertyPath
    Set opt = Target
    mTargetProperty = TargetProperty
    
    If TargetProperty = "Value" Then
        mBindingMode = TwoWayBinding
    Else
        mBindingMode = OneWayBinding
    End If
    
    UpdateTarget
End Sub

Private Sub opt_Change()
    If DO_DEBUG Then Debug.Print "OptionButtonPropertyBinding opt_Change()"
    UpdateSource
End Sub

Private Sub vm_PropertyChanged(ByVal PropertyName As String)
    If PropertyName <> mSourcePropertyPath Then Exit Sub
    If DO_DEBUG Then Debug.Print "OptionButtonPropertyBinding vm_PropertyChanged()"
    UpdateTarget
End Sub

Private Sub UpdateTarget()
    If mBindingMode = TwoWayBinding Or mBindingMode = OneWayBinding Then
        Dim vNewValue As String
        vNewValue = CallByName(vm, mSourcePropertyPath, VbGet)
        CallByName opt, mTargetProperty, VbLet, vNewValue
    End If
End Sub

Private Sub UpdateSource()
    If mBindingMode = TwoWayBinding Or mBindingMode = OneWayToSource Then
        Dim vNewValue As Variant
        vNewValue = CallByName(opt, mTargetProperty, VbGet)
        CallByName vm, mSourcePropertyPath, VbLet, vNewValue
    End If
End Sub
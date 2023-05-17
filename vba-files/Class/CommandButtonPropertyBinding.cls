VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandButtonPropertyBinding"
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
Private WithEvents cmb As MSForms.CommandButton
Attribute cmb.VB_VarHelpID = -1
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
    Set IPropertyBinding_Target = cmb
End Property

Private Property Get IPropertyBinding_TargetProperty() As String
    IPropertyBinding_TargetProperty = mTargetProperty
End Property

Public Sub Create(ByVal Source As IViewModel, ByVal SourcePropertyPath As String, ByVal Target As Control, Optional ByVal TargetProperty As String)
    If DO_DEBUG Then Debug.Print "CommandButtonPropertyBinding Create(vm, "; SourcePropertyPath; ", "; Target.Name; ")"
    Set vm = Source
    mSourcePropertyPath = SourcePropertyPath
    Set cmb = Target
    mTargetProperty = TargetProperty
    
    mBindingMode = OneWayBinding
        
    UpdateTarget
End Sub

'Private Sub lbl_Click()
'    If DO_DEBUG Then Debug.Print "TextBoxPropertyBinding lbl_Click()"
    'UpdateSource
'End Sub

Private Sub vm_PropertyChanged(ByVal PropertyName As String)
    If PropertyName <> mSourcePropertyPath Then Exit Sub
    If DO_DEBUG Then Debug.Print "CommandButtonPropertyBinding vm_PropertyChanged()"
    UpdateTarget
End Sub

Private Sub UpdateTarget()
    If CanUpdateTarget Then
        If DO_DEBUG Then Debug.Print "CommandButtonPropertyBinding cmb_PropertyChanged()"
        Dim vnewValue As String
        vnewValue = CallByName(vm, mSourcePropertyPath, VbGet)
        CallByName cmb, mTargetProperty, VbLet, vnewValue
    End If
    
    If mBindingMode = OneTimeBinding Then mOneTimeBoundCounter = mOneTimeBoundCounter + 1
End Sub

'Private Sub UpdateSource()
'    If mBindingMode = TwoWayBinding Or mBindingMode = OneWayToSource Then
    ' NOOP
'    End If
'End Sub

Private Function CanUpdateTarget() As Boolean
    If mBindingMode = TwoWayBinding Then CanUpdateTarget = True
    If mBindingMode = OneWayBinding Then CanUpdateTarget = True
    If mBindingMode = OneTimeBinding And mOneTimeBoundCounter = 0 Then CanUpdateTarget = True
End Function


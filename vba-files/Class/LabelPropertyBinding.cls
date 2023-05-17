VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LabelPropertyBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Bindings")
Option Explicit
Implements IPropertyBinding

Private Const DO_DEBUG As Boolean = False

Private mBindingMode As BindingMode
Private WithEvents vm As ViewModel
Attribute vm.VB_VarHelpID = -1
Private mSourcePropertyPath As String
Private WithEvents lbl As MSForms.Label
Attribute lbl.VB_VarHelpID = -1
Private mTargetProperty As String

Private Property Get IPropertyBinding_Mode() As BindingMode
    IPropertyBinding_Source = mBindingMode
End Property

Private Property Get IPropertyBinding_Source() As Object
    Set IPropertyBinding_Source = vm
End Property

Private Property Get IPropertyBinding_SourcePropertyPath() As String
    IPropertyBinding_SourcePropertyPath = mSourcePropertyPath
End Property

Private Property Get IPropertyBinding_Target() As Object
    Set IPropertyBinding_Target = lbl
End Property

Private Property Get IPropertyBinding_TargetProperty() As String
    IPropertyBinding_TargetProperty = mTargetProperty
End Property

Public Sub Create(ByVal Source As IViewModel, ByVal SourcePropertyPath As String, ByVal Target As Control, Optional ByVal TargetProperty As String)
    If DO_DEBUG Then Debug.Print "LabelPropertyBinding Create(vm, "; SourcePropertyPath; ", "; Target.Name; ")"
    Set vm = Source
    mSourcePropertyPath = SourcePropertyPath
    Set lbl = Target
    mTargetProperty = TargetProperty
    
    mBindingMode = OneWayBinding
        
    UpdateTarget
End Sub

Private Sub lbl_Click()
    If DO_DEBUG Then Debug.Print "TextBoxPropertyBinding lbl_Click()"
    'UpdateSource
End Sub

Private Sub vm_PropertyChanged(ByVal PropertyName As String)
    If PropertyName <> mSourcePropertyPath Then Exit Sub
    If DO_DEBUG Then Debug.Print "LabelPropertyBinding vm_PropertyChanged()"
    UpdateTarget
End Sub

Private Sub UpdateTarget()
    If mBindingMode = TwoWayBinding Or mBindingMode = OneWayBinding Then
        If DO_DEBUG Then Debug.Print "LabelPropertyBinding vm_PropertyChanged()"
        Dim vNewValue As String
        vNewValue = CallByName(vm, mSourcePropertyPath, VbGet)
        CallByName lbl, mTargetProperty, VbLet, vNewValue
    End If
End Sub

Private Sub UpdateSource()
    If mBindingMode = TwoWayBinding Or mBindingMode = OneWayToSource Then
    ' NOOP
    End If
End Sub

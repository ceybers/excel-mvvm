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
'@Folder MVVM.Infrastructure.Bindings.PropertyBindings
Option Explicit
Implements IPropertyBinding
Implements IHandlePropertyChanged

Private Const DO_DEBUG As Boolean = False

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents opt As MSForms.OptionButton
Attribute opt.VB_VarHelpID = -1

Private Type TState
    BindingMode As BindingMode
    ViewModel As IViewModel
    SourcePropertyPath As String
    TargetProperty As String
    OneTimeBoundCounter As Long
End Type
Private This As TState

Private Property Get IPropertyBinding_Mode() As BindingMode
    IPropertyBinding_Mode = This.BindingMode
End Property

Private Property Get IPropertyBinding_Source() As IViewModel
    Set IPropertyBinding_Source = This.ViewModel
End Property

Private Property Get IPropertyBinding_SourcePropertyPath() As String
    IPropertyBinding_SourcePropertyPath = This.SourcePropertyPath
End Property

Private Property Get IPropertyBinding_Target() As Object
    Set IPropertyBinding_Target = opt
End Property

Private Property Get IPropertyBinding_TargetProperty() As String
    IPropertyBinding_TargetProperty = This.TargetProperty
End Property

Public Sub Create(ByVal Source As IViewModel, ByVal SourcePropertyPath As String, ByVal Target As Control, Optional ByVal TargetProperty As String)
    If DO_DEBUG Then Debug.Print "OptionButtonPropertyBinding Create(vm, "; SourcePropertyPath; ", "; Target.Name; ")"
    Set This.ViewModel = Source
    This.SourcePropertyPath = SourcePropertyPath
    Set opt = Target
    This.TargetProperty = TargetProperty
    
    If TargetProperty = "Value" Then
        This.BindingMode = TwoWayBinding
    Else
        This.BindingMode = OneWayBinding
    End If
    
    UpdateTarget
End Sub

Private Sub opt_Change()
    If DO_DEBUG Then Debug.Print "OptionButtonPropertyBinding opt_Change()"
    UpdateSource
End Sub

Private Sub IHandlePropertyChanged_HandlePropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    If PropertyName <> This.SourcePropertyPath Then Exit Sub
    If DO_DEBUG Then Debug.Print "OptionButtonPropertyBinding vm_PropertyChanged()"
    UpdateTarget
End Sub

Private Sub UpdateTarget()
    If CanUpdateTarget Then
        Dim vNewValue As String
        vNewValue = CallByName(This.ViewModel, This.SourcePropertyPath, VbGet)
        CallByName opt, This.TargetProperty, VbLet, vNewValue
    End If
    
    If This.BindingMode = OneTimeBinding Then This.OneTimeBoundCounter = This.OneTimeBoundCounter + 1
End Sub

Private Sub UpdateSource()
    If This.BindingMode = TwoWayBinding Or This.BindingMode = OneWayToSource Then
        Dim vNewValue As Variant
        vNewValue = CallByName(opt, This.TargetProperty, VbGet)
        CallByName This.ViewModel, This.SourcePropertyPath, VbLet, vNewValue
    End If
End Sub

Private Function CanUpdateTarget() As Boolean
    If This.BindingMode = TwoWayBinding Then CanUpdateTarget = True
    If This.BindingMode = OneWayBinding Then CanUpdateTarget = True
    If This.BindingMode = OneTimeBinding And This.OneTimeBoundCounter = 0 Then CanUpdateTarget = True
End Function

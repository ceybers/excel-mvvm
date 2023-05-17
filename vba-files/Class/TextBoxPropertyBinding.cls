VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextBoxPropertyBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Bindings")
Option Explicit
Implements IPropertyBinding

Private WithEvents vm As ViewModel
Attribute vm.VB_VarHelpID = -1
Private mSourcePropertyPath As String
Private WithEvents tb As MSForms.TextBox
Attribute tb.VB_VarHelpID = -1
Private mTargetProperty As String

Private Property Get IPropertyBinding_Source() As Object
    Set IPropertyBinding_Source = vm
End Property

Private Property Get IPropertyBinding_SourcePropertyPath() As String
    IPropertyBinding_SourcePropertyPath = mSourcePropertyPath
End Property

Private Property Get IPropertyBinding_Target() As Object
    Set IPropertyBinding_Target = tb
End Property

Private Property Get IPropertyBinding_TargetProperty() As String
    IPropertyBinding_TargetProperty = mTargetProperty
End Property

Public Sub Create(ByVal Source As IViewModel, ByVal SourcePropertyPath As String, ByVal Target As Control, Optional ByVal TargetProperty As String)
    Set vm = Source
    mSourcePropertyPath = SourcePropertyPath
    Set tb = Target
    mTargetProperty = TargetProperty
End Sub

Private Sub tb_Change()
    Debug.Print "TextBoxPropertyBinding tb_Change()"
    Dim vNewValue As String
    vNewValue = CallByName(tb, mTargetProperty, VbGet)
    CallByName vm, mSourcePropertyPath, VbLet, vNewValue
End Sub

Private Sub vm_PropertyChanged(ByVal PropertyName As String)
    Debug.Print "TextBoxPropertyBinding vm_PropertyChanged()"
    Dim vNewValue As String
    vNewValue = CallByName(vm, mSourcePropertyPath, VbGet)
    CallByName tb, mTargetProperty, VbLet, vNewValue
End Sub

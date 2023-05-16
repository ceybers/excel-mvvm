VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("VBAProject")
Option Explicit

Private WithEvents cb As MSForms.CommandButton
Attribute cb.VB_VarHelpID = -1
Private WithEvents tb As MSForms.TextBox
Attribute tb.VB_VarHelpID = -1
Private WithEvents lb As MSForms.Label
Attribute lb.VB_VarHelpID = -1

Private Type TState
    'ViewModel As IViewModel
    ViewModel As ExampleViewModel
    Name As String
End Type
Private This As TState

Public Property Get Name() As String
    Name = This.Name
End Property

Public Property Let Name(ByVal RHS As String)
    This.Name = RHS
End Property

Public Property Get ViewModel() As ExampleViewModel
    Set ViewModel = This.ViewModel
End Property

Public Property Set ViewModel(ByVal RHS As ExampleViewModel)
    Set This.ViewModel = RHS
End Property

Public Sub Update(ByVal ctrl As MSForms.Control, ByVal PropertyName As String, ByVal vm As IViewModel)
    'Set Create = New EventHandler
    Name = PropertyName
    Set This.ViewModel = vm
    
    If TypeOf ctrl Is MSForms.CommandButton Then
        Set cb = ctrl
    ElseIf TypeOf ctrl Is MSForms.TextBox Then
        Set tb = ctrl
    ElseIf TypeOf ctrl Is MSForms.Label Then
        Set lb = ctrl
    End If
End Sub

Public Sub HandlePropertyChanged(ByVal PropertyName As String)
    If This.ViewModel Is Nothing Then Exit Sub
    
    If PropertyName = Name Then
        Repaint
    End If
End Sub

Public Sub Repaint()
    Debug.Assert Not This.ViewModel Is Nothing
    
    If Not cb Is Nothing Then
        cb.Caption = CallByName(This.ViewModel, Name, VbGet)
    ElseIf Not tb Is Nothing Then
        tb.Value = CallByName(This.ViewModel, Name, VbGet)
    ElseIf Not lb Is Nothing Then
        lb.Caption = CallByName(This.ViewModel, Name, VbGet)
    End If
End Sub

Private Sub cb_Click()
    MsgBox "cb_Click(" & This.Name & ")"
End Sub

Private Sub tb_Change()
    Debug.Print "tb_Change(" & This.Name & ")"
End Sub

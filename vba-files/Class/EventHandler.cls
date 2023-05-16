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
Private WithEvents cbo As MSForms.ComboBox
Attribute cbo.VB_VarHelpID = -1
Private WithEvents lv As MSComctlLib.ListView
Attribute lv.VB_VarHelpID = -1

Private Type TState
    'ViewModel As IViewModel
    ViewModel As ViewModel
    Name As String
End Type
Private This As TState

Public Property Get Name() As String
    Name = This.Name
End Property

Public Property Let Name(ByVal RHS As String)
    This.Name = RHS
End Property

Public Property Get ViewModel() As ViewModel
    Set ViewModel = This.ViewModel
End Property

Public Property Set ViewModel(ByVal RHS As ViewModel)
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
    ElseIf TypeOf ctrl Is MSForms.ComboBox Then
        Set cbo = ctrl
    ElseIf TypeOf ctrl Is MSComctlLib.ListView Then
        Set lv = ctrl
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
    ' TODO Get 'CallByName(This.ViewModel, Name, VbGet)' once
    If Not cb Is Nothing Then
        cb.Caption = CallByName(This.ViewModel, Name, VbGet)
    ElseIf Not tb Is Nothing Then
        tb.Value = CallByName(This.ViewModel, Name, VbGet)
    ElseIf Not lb Is Nothing Then
        lb.Caption = CallByName(This.ViewModel, Name, VbGet)
    ElseIf Not cbo Is Nothing Then
        cbo.Clear
        cbo.Value = CallByName(This.ViewModel, Name, VbGet)
        ' TODO Implement this in VM
        cbo.AddItem "Small"
        cbo.AddItem "Medium"
        cbo.AddItem "Large"
    ElseIf Not lv Is Nothing Then
        lv.ListItems.Clear
        lv.ListItems.Add Text:="Small"
        lv.ListItems.Add Text:="Medium"
        lv.ListItems.Add Text:="Large"
        Dim li As ListItem
        For Each li In lv.ListItems
            li.Selected = li.Text = CallByName(This.ViewModel, Name, VbGet)
        Next li
    End If
End Sub

Private Sub cb_Click()
    MsgBox "cb_Click(" & This.Name & ")"
End Sub

Private Sub tb_Change()
    Debug.Print "tb_Change(" & This.Name & ")"
End Sub

Private Sub cbo_Change()
    Debug.Print
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    MsgBox "lv_Change(" & This.Name & ")"
    CallByName This.ViewModel, Name, VbLet, Item.Text
End Sub

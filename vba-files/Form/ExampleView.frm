VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExampleView 
   Caption         =   "Test View Form"
   ClientHeight    =   4935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6810
   OleObjectBlob   =   "ExampleView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExampleView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "View"
Option Explicit
Implements IView
 
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents vm As ViewModel
Attribute vm.VB_VarHelpID = -1
Private Type TView
    IsCancelled As Boolean
    EventHandlers As Collection
    'ViewModel As IViewModel
End Type
Private This As TView

Private Sub cmbTest_Click()
    vm.Foo = "zzz"
End Sub

Private Sub CommandButton1_Click()
    vm.FirstName = "Bob"
End Sub

Private Sub OkButton_Click()
    Me.Hide
End Sub
 
Private Sub CancelButton_Click()
    OnCancel
End Sub
 
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub
 
Private Sub OnCancel()
    This.IsCancelled = True
    Me.Hide
End Sub
 
Private Function IView_ShowDialog(ByVal ViewModel As Object) As Boolean
    'Set This.ViewModel = ViewModel
    Set vm = ViewModel
    vm.Foo = "test"
    
    InitializeListViews
    BindControls
    
    Me.Show
    IView_ShowDialog = Not This.IsCancelled
End Function

Private Sub InitializeListViews()
    With Me.lvSize
        .view = lvwReport
        '.Arrange = lvwAutoLeft
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .FullRowSelect = True
        .ColumnHeaders.Clear
        .ColumnHeaders.Add Text:="Value"
    End With
End Sub

Private Sub BindControls()
    Set This.EventHandlers = New Collection
    
    BindControl Me.cmbTest, "Foo"
    BindControl Me.txtFoobar, "Bar"
    BindControl Me.lblFoobar, "foo"
    BindControl Me.txtFirstname, "FirstName"
    BindControl Me.txtLastName, "LastName"
    BindControl Me.cboSize, "Size"
    BindControl Me.lvSize, "Size"
End Sub

Private Sub BindControl(ByVal ctrl As Control, ByVal PropertyName As String)
    Dim eh As EventHandler
    Set eh = New EventHandler
    eh.Update ctrl, PropertyName, vm
    This.EventHandlers.Add Item:=eh
    eh.Repaint
End Sub

Private Sub vm_PropertyChanged(ByVal PropertyName As String)
    If This.EventHandlers Is Nothing Then Exit Sub
    
    Debug.Print "PropertyChanged("; PropertyName; ")"
    
    Dim eh As EventHandler
    For Each eh In This.EventHandlers
        eh.HandlePropertyChanged PropertyName
    Next eh
End Sub

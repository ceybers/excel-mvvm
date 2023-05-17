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
    'ViewModel As IViewModel
End Type
Private This As TView

Private Sub cmbTest_Click()
    vm.DebugPrint
End Sub

Private Sub cmdSetNameToBob_Click()
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
    Set vm = ViewModel
    
    InitializeControls
    BindControls
    BindCommands
    
    Me.Show vbModal
    
    IView_ShowDialog = Not This.IsCancelled
End Function

Private Sub InitializeControls()
    vm.InitializeListViewSize Me.lvSize
    vm.InitializeTreeViewSize Me.tvSize
End Sub

Private Sub BindControls()
    With vm.Context.BindingManager
        .BindPropertyPath vm, "FirstName", Me.txtFirstname, "Value"
        
        .BindPropertyPath vm, "LastName", Me.lblFirstName, "Caption"
        
        .BindPropertyPath vm, "IsFoobar", Me.chkIsFoobar, "Value"
        .BindPropertyPath vm, "IsFoobarCaption", Me.chkIsFoobar, "Caption"
        
        .BindPropertyPath vm, "IsFoobar", Me.optIsFooBar, "Value"
        .BindPropertyPath vm, "IsFoobarCaption", Me.optIsFooBar, "Caption"
        
        .BindPropertyPath vm, "Size", Me.cboSize, "Value"
        .BindPropertyPath vm, "SizeOptions", Me.cboSize, "List"
        
        .BindPropertyPath vm, "Size", Me.lvSize, "SelectedItem"
        .BindPropertyPath vm, "SizeOptions", Me.lvSize, "ListItems"
        
        .BindPropertyPath vm, "Size", Me.tvSize, "SelectedItem"
        .BindPropertyPath vm, "SizeOptions", Me.tvSize, "Nodes"
        
        .BindPropertyPath vm, "FirstName", Me.cmbTestMsgbox, "Caption"
    End With
End Sub

Private Sub BindCommands()
    With vm.Context.CommandManager
        .BindCommand vm, "TestMsgboxCommand", Me.cmbTestMsgbox
    End With
End Sub

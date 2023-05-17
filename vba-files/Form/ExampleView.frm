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

'@Folder "MVVM.Views"
Option Explicit
Implements IView
 
Private Type TView
    IsCancelled As Boolean
    ViewModel As SomeViewModel
End Type
Private This As TView

Private Sub cmbTest_Click()
    This.ViewModel.DebugPrint
End Sub

Private Sub cmdSetNameToBob_Click()
    This.ViewModel.FirstName = "Bob"
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
    Set This.ViewModel = ViewModel
    
    InitializeControls
    BindControls
    BindCommands
    
    Me.Show vbModal
    
    IView_ShowDialog = Not This.IsCancelled
End Function

Private Sub InitializeControls()
    '@Ignore ArgumentWithIncompatibleObjectType
    This.ViewModel.InitializeListViewSize Me.lvSize
    '@Ignore ArgumentWithIncompatibleObjectType
    This.ViewModel.InitializeTreeViewSize Me.tvSize
End Sub

Private Sub BindControls()
    With This.ViewModel.Context.BindingManager
        .BindPropertyPath This.ViewModel, "FirstName", Me.txtFirstname, "Value"
        
        .BindPropertyPath This.ViewModel, "LastName", Me.lblFirstName, "Caption"
        
        .BindPropertyPath This.ViewModel, "IsFoobar", Me.chkIsFoobar, "Value"
        .BindPropertyPath This.ViewModel, "IsFoobarCaption", Me.chkIsFoobar, "Caption"
        
        .BindPropertyPath This.ViewModel, "IsFoobar", Me.optIsFooBar, "Value"
        .BindPropertyPath This.ViewModel, "IsFoobarCaption", Me.optIsFooBar, "Caption"
        
        .BindPropertyPath This.ViewModel, "Size", Me.cboSize, "Value"
        .BindPropertyPath This.ViewModel, "SizeOptions", Me.cboSize, "List"
        
        .BindPropertyPath This.ViewModel, "FavColorViewModel.FavoriteColors", Me.lvSize, "ListItems"
        .BindPropertyPath This.ViewModel, "FavColorViewModel.FavoriteColor", Me.lvSize, "SelectedItem"
        
        '.BindPropertyPath This.ViewModel, "FavColorViewModel.FavFoodViewModel.FavoriteFoods", Me.lvSize, "ListItems"
        '.BindPropertyPath This.ViewModel, "FavColorViewModel.FavFoodViewModel.FavoriteFood", Me.lvSize, "SelectedItem"
        
        .BindPropertyPath This.ViewModel, "Size", Me.tvSize, "SelectedItem"
        .BindPropertyPath This.ViewModel, "SizeOptions", Me.tvSize, "Nodes"
        
        .BindPropertyPath This.ViewModel, "FirstName", Me.cmbTestMsgbox, "Caption"
    End With
End Sub

Private Sub BindCommands()
    Dim SomeViewCommand As TestViewCommand
    Set SomeViewCommand = New TestViewCommand
    Set SomeViewCommand.Context = This.ViewModel.Context
    Set SomeViewCommand.View = Me
    
    With This.ViewModel.Context.CommandManager
        .BindCommand This.ViewModel.Context, This.ViewModel, This.ViewModel.TestMsgboxCommand, Me.cmbTestMsgbox
        .BindCommand This.ViewModel.Context, This.ViewModel, This.ViewModel.TestDoCmdCommand, Me.cmdDoVMCmd
        .BindCommand This.ViewModel.Context, This.ViewModel, SomeViewCommand, Me.cmdDoVCmd
    End With
End Sub

Public Sub DoSomething()
    MsgBox "ExampleView.DoSomething()"
End Sub

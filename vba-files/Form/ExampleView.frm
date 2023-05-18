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

'@Folder "MVVM.Example1"
Option Explicit
Implements IView
Implements ICancellable
 
Private Type TView
    Context As IAppContext
    IsCancelled As Boolean
    ViewModel As SomeViewModel
End Type
Private This As TView

Private Property Get IView_ViewModel() As Object
    Set IView_ViewModel = This.ViewModel
End Property

Public Property Get ViewModel() As SomeViewModel
    Set ViewModel = This.ViewModel
End Property

Public Property Set ViewModel(ByVal vNewValue As SomeViewModel)
    Set This.ViewModel = vNewValue
End Property

Public Property Get Context() As IAppContext
    Set Context = This.Context
End Property

Public Property Set Context(ByVal vNewValue As IAppContext)
    Set This.Context = vNewValue
End Property
 
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

Private Property Get ICancellable_IsCancelled() As Boolean
    ICancellable_IsCancelled = This.IsCancelled
End Property

Private Sub ICancellable_OnCancel()
    OnCancel
End Sub

Private Sub IView_Show()
    IView_ShowDialog
End Sub
 
Private Sub IView_Hide()
    Me.Hide
End Sub

Public Function Create(ByVal Context As IAppContext, ByVal ViewModel As SomeViewModel) As IView
    Dim Result As ExampleView
    Set Result = New ExampleView
    
    Set Result.Context = Context
    Set Result.ViewModel = ViewModel

    Set Create = Result
End Function

Private Function IView_ShowDialog() As Boolean
    InitializeControls
    BindControls
    BindCommands
    
    Me.Show vbModal
    
    IView_ShowDialog = Not This.IsCancelled
End Function

Private Sub InitializeControls()
    '@Ignore ArgumentWithIncompatibleObjectType
    ViewModel.InitializeListViewSize Me.lvSize
    '@Ignore ArgumentWithIncompatibleObjectType
    ViewModel.InitializeTreeViewSize Me.tvSize
End Sub

Private Sub BindControls()
    With Context.BindingManager
        .BindPropertyPath ViewModel, "FirstName", Me.txtFirstname, "Value"
        
        .BindPropertyPath ViewModel, "LastName", Me.lblFirstName, "Caption"
        
        .BindPropertyPath ViewModel, "IsFoobar", Me.chkIsFoobar, "Value"
        .BindPropertyPath ViewModel, "IsFoobarCaption", Me.chkIsFoobar, "Caption"
        
        .BindPropertyPath ViewModel, "IsFoobar", Me.optIsFooBar, "Value"
        .BindPropertyPath ViewModel, "IsFoobarCaption", Me.optIsFooBar, "Caption"
        
        .BindPropertyPath ViewModel, "Size", Me.cboSize, "Value"
        .BindPropertyPath ViewModel, "SizeOptions", Me.cboSize, "List"
        
        .BindPropertyPath ViewModel, "FavColorViewModel.FavoriteColors", Me.lvSize, "ListItems"
        .BindPropertyPath ViewModel, "FavColorViewModel.FavoriteColor", Me.lvSize, "SelectedItem"
        
        '.BindPropertyPath ViewModel, "FavColorViewModel.FavFoodViewModel.FavoriteFoods", Me.lvSize, "ListItems"
        '.BindPropertyPath ViewModel, "FavColorViewModel.FavFoodViewModel.FavoriteFood", Me.lvSize, "SelectedItem"
        
        .BindPropertyPath ViewModel, "Size", Me.tvSize, "SelectedItem"
        .BindPropertyPath ViewModel, "SizeOptions", Me.tvSize, "Nodes"
        
        .BindPropertyPath ViewModel, "FirstName", Me.cmbTestMsgbox, "Caption"
    End With
End Sub

Private Sub BindCommands()
    Dim SomeViewCommand As TestViewCommand
    Set SomeViewCommand = New TestViewCommand
    Set SomeViewCommand.Context = Context
    Set SomeViewCommand.View = Me
    
    Dim OKView As ICommand
    Set OKView = OKViewCommand.Create(Context, Me, ViewModel)
    
    With This.Context.CommandManager
        .BindCommand Context, ViewModel, ViewModel.TestMsgboxCommand, Me.cmbTestMsgbox
        .BindCommand Context, ViewModel, SomeViewCommand, Me.cmdDoVCmd
        .BindCommand Context, ViewModel, ViewModel.TestDoCmdCommand, Me.cmdDoVMCmd
        .BindCommand Context, ViewModel, OKView, Me.OkButton
    End With
End Sub

Private Sub cmbTest_Click()
    This.ViewModel.DebugPrint
End Sub

Private Sub cmdSetNameToBob_Click()
    This.ViewModel.FirstName = "Bob"
End Sub

Private Sub CancelButton_Click()
    OnCancel
End Sub

Public Sub DoSomething()
    MsgBox "ExampleView.DoSomething()"
End Sub

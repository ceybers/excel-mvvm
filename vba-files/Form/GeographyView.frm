VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GeographyView 
   Caption         =   "GeographyView"
   ClientHeight    =   6600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7485
   OleObjectBlob   =   "GeographyView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GeographyView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "MVVM.Example2"
Option Explicit
Implements IView
Implements ICancellable
 
Private Type TView
    Context As IAppContext
    IsCancelled As Boolean
    ViewModel As CountryViewModel
End Type
Private This As TView

Private Property Get IView_ViewModel() As Object
    Set IView_ViewModel = This.ViewModel
End Property

Public Property Get ViewModel() As CountryViewModel
    Set ViewModel = This.ViewModel
End Property

Public Property Set ViewModel(ByVal vNewValue As CountryViewModel)
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

Public Function Create(ByVal Context As IAppContext, ByVal ViewModel As CountryViewModel) As IView
    Dim Result As GeographyView
    Set Result = New GeographyView
    
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
    'ViewModel.InitializeListViewSize Me.lvSize
    '@Ignore ArgumentWithIncompatibleObjectType
    'ViewModel.InitializeTreeViewSize Me.tvSize
End Sub

Private Sub BindControls()
    With Context.BindingManager
        .BindPropertyPath ViewModel, "Countries", Me.ListView1, "ListItems"
        .BindPropertyPath ViewModel, "Country", Me.ListView1, "SelectedItem"
        .BindPropertyPath ViewModel, "Country", Me.TextBox1, "Value"
        
        .BindPropertyPath ViewModel, "CityViewModel.Cities", Me.ListView2, "ListItems"
        .BindPropertyPath ViewModel, "CityViewModel.City", Me.ListView2, "SelectedItem"
        .BindPropertyPath ViewModel, "CityViewModel.Cities", Me.ComboBox1, "List"
        .BindPropertyPath ViewModel, "CityViewModel.City", Me.ComboBox1, "Value"
    End With
End Sub

Private Sub BindCommands()
    Dim OKView As ICommand
    Set OKView = OKViewCommand.Create(Context, Me, ViewModel)
    
    With This.Context.CommandManager
        .BindCommand Context, ViewModel, OKView, Me.cmbOK
    End With
End Sub

Private Sub cmbCancel_Click()
    OnCancel
End Sub
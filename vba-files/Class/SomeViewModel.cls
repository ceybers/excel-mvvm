VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SomeViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule ProcedureNotUsed
'@Folder "MVVM.ViewModels"
Option Explicit
Implements IViewModel
Implements INotifyPropertyChanged

Private Type TState
    AppContext As AppContext
    Handlers As Collection
    
    FirstName As String
    LastName As String
    DateOfBirth As Date
    Foo As String
    Bar As Double
    
    Size As String
    SizeOptions As Scripting.Dictionary
    
    IsFoobar As Boolean
    IsFoobarCaption As String
    
    TestMsgboxCommand As ICommand
    TestDoCmdCommand As ICommand
    
    FavColorViewModel As FavColorViewModel
End Type
Private This As TState

Public Property Get FavColorViewModel() As FavColorViewModel
    Set FavColorViewModel = This.FavColorViewModel
End Property

Public Property Get TestDoCmdCommand() As ICommand
    Set TestDoCmdCommand = This.TestDoCmdCommand
End Property

Public Property Set TestDoCmdCommand(ByVal vNewValue As ICommand)
    Set This.TestDoCmdCommand = vNewValue
End Property

Public Property Get TestMsgboxCommand() As ICommand
    Set TestMsgboxCommand = This.TestMsgboxCommand
End Property

Public Property Set TestMsgboxCommand(ByVal vNewValue As ICommand)
    Set This.TestMsgboxCommand = vNewValue
End Property

Public Property Get Context() As IAppContext
    Set Context = This.AppContext
End Property

Public Property Set Context(ByVal vNewValue As IAppContext)
    Set This.AppContext = vNewValue
    Set This.FavColorViewModel.Context = vNewValue
End Property

Public Property Get IsFoobar() As Boolean
    IsFoobar = This.IsFoobar
End Property

Public Property Let IsFoobar(ByVal RHS As Boolean)
    If This.IsFoobar <> RHS Then
        This.IsFoobar = RHS
        OnPropertyChanged "IsFoobar"
    End If
End Property

Public Property Get IsFoobarCaption() As String
    IsFoobarCaption = This.IsFoobarCaption
End Property

Public Property Let IsFoobarCaption(ByVal vNewValue As String)
    This.IsFoobarCaption = vNewValue
End Property

Public Property Get FirstName() As String
    FirstName = This.FirstName
End Property

Public Property Let FirstName(ByVal RHS As String)
    If This.FirstName <> RHS Then
        This.FirstName = RHS
        OnPropertyChanged "FirstName"
    End If
End Property

Public Property Get LastName() As String
    LastName = This.LastName
End Property

Public Property Let LastName(ByVal RHS As String)
    If This.LastName <> RHS Then
        This.LastName = RHS
        OnPropertyChanged "LastName"
    End If
End Property

Public Property Get DateOfBirth() As Date
    DateOfBirth = This.DateOfBirth
End Property

Public Property Let DateOfBirth(ByVal RHS As Date)
    If This.DateOfBirth <> RHS Then
        This.DateOfBirth = RHS
        OnPropertyChanged "DateOfBirth"
    End If
End Property

Public Property Get Foo() As String
    Foo = This.Foo
End Property

Public Property Let Foo(ByVal RHS As String)
    If This.Foo <> RHS Then
        This.Foo = RHS
        OnPropertyChanged "Foo"
    End If
End Property

Public Property Get Size() As String
    Size = This.Size
End Property

Public Property Let Size(ByVal RHS As String)
    If This.Size <> RHS Then
        This.Size = RHS
        OnPropertyChanged "Size"
    End If
End Property

Public Property Get SizeOptions() As Scripting.Dictionary
    Set SizeOptions = This.SizeOptions
End Property

Public Property Set SizeOptions(ByVal RHS As Scripting.Dictionary)
    Set This.SizeOptions = RHS
    OnPropertyChanged "SizeOptions"
End Property

Public Property Get Bar() As Double
    Bar = This.Bar
End Property

Public Property Let Bar(ByVal RHS As Double)
    If This.Bar <> RHS Then
        This.Bar = RHS
        OnPropertyChanged "Bar"
    End If
End Property

Private Sub class_initialize()
    Set This.TestMsgboxCommand = New TestMsgboxCommand
    Set This.Handlers = New Collection
    
    Set This.FavColorViewModel = New FavColorViewModel
End Sub

Private Function IViewModel_IsValid() As Boolean
    IViewModel_IsValid = True
End Function

Private Function IViewModel_Self() As IViewModel
    Set IViewModel_Self = Me
End Function

Public Function Self() As SomeViewModel
    Set Self = Me
End Function

Private Sub INotifyPropertyChanged_RegisterHandler(ByVal Handler As IHandlePropertyChanged)
    This.Handlers.Add Item:=Handler
End Sub

Private Sub INotifyPropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    Dim Handler As IHandlePropertyChanged
    For Each Handler In This.Handlers
        Handler.HandlePropertyChanged Me, PropertyName
    Next Handler
End Sub

Private Sub OnPropertyChanged(ByVal PropertyName As String)
    INotifyPropertyChanged_OnPropertyChanged Me, PropertyName
End Sub

Public Sub DebugPrint()
    Debug.Print "ViewModel.DebugPrint()"
    Debug.Print " FirstName: "; FirstName
    'Debug.Print " LastName: "; LastName
    'Debug.Print " DateOfBirth: "; DateOfBirth
    'Debug.Print " Foo: "; Foo
    'Debug.Print " Bar: "; Bar
    Debug.Print " Size: "; Size
    Debug.Print " IsFoobar: "; IsFoobar
    Debug.Print "---"
End Sub

Public Sub InitializeListViewSize(ByVal lv As ListView4)
    With lv
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add Text:="Value"
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .FullRowSelect = True
        .ListItems.Add Text:="Blue"
        .ListItems.Add Text:="Red"
        .ListItems.Add Text:="Green"
    End With
End Sub

Public Sub InitializeTreeViewSize(ByVal tv As TreeView4)
    With tv
        .Appearance = cc3D
        .BorderStyle = ccNone
        .FullRowSelect = True
        .HideSelection = False
        .SingleSel = True
        .Style = tvwTreelinesPlusMinusText
    End With
End Sub

Public Sub DoSomethingByCommand()
    FirstName = "Something"
End Sub

Public Sub InitializeCommands()
    Dim TestVMCmd As TestVMSubCommand
    Set TestVMCmd = New TestVMSubCommand
    Set TestVMCmd.Context = This.AppContext
    Set TestVMCmd.ViewModel = Me
    Set This.TestDoCmdCommand = TestVMCmd
End Sub
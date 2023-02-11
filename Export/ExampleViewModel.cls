VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExampleViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("ViewModel")
Option Explicit
Implements IViewModel
Implements INotifyPropertyChanged

Event PropertyChanged(ByVal PropertyName As String)

Private Type TState
    FirstName As String
    LastName As String
    DateOfBirth As Date
    Foo As String
    Bar As Double
End Type
Private This As TState

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

Public Property Get Bar() As Double
    Bar = This.Bar
End Property

Public Property Let Bar(ByVal RHS As Double)
    If This.Bar <> RHS Then
        This.Bar = RHS
        OnPropertyChanged "Bar"
    End If
End Property

Private Function IViewModel_IsValid() As Boolean
    IViewModel_IsValid = True
End Function

Private Function IViewModel_Self() As IViewModel
    Set IViewModel_Self = Me
End Function

Public Function Self() As ExampleViewModel
    Set Self = Me
End Function

Private Sub INotifyPropertyChanged_OnPropertyChanged(ByVal PropertyName As String)
    RaiseEvent PropertyChanged(PropertyName)
End Sub

Private Sub OnPropertyChanged(ByVal PropertyName As String)
    INotifyPropertyChanged_OnPropertyChanged PropertyName
End Sub

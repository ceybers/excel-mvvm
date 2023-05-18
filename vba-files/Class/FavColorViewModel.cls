VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FavColorViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.Example1"
Option Explicit
Implements IViewModel
Implements INotifyPropertyChanged

Private Type TState
    AppContext As AppContext
    Notifier As INotifyPropertyChanged
    
    FavoriteColor As String
    FavoriteColors As Scripting.Dictionary
    
    FavFoodViewModel As FavFoodViewModel
End Type
Private This As TState

Public Property Get FavFoodViewModel() As FavFoodViewModel
    Set FavFoodViewModel = This.FavFoodViewModel
End Property

Public Property Get FavoriteColor() As String
    FavoriteColor = This.FavoriteColor
End Property

Public Property Let FavoriteColor(ByVal RHS As String)
    If This.FavoriteColor = RHS Then Exit Property
    This.FavoriteColor = RHS
    OnPropertyChanged "FavoriteColor"
End Property

Public Property Get FavoriteColors() As Scripting.Dictionary
    Set FavoriteColors = This.FavoriteColors
End Property

'Public Property Set FavoriteColors(ByVal RHS As Collection)
'    Set This.FavoriteColors = RHS
'End Property

Public Property Get Context() As IAppContext
    Set Context = This.AppContext
End Property

Public Property Set Context(ByVal vNewValue As IAppContext)
    Set This.AppContext = vNewValue
    
    Set This.FavFoodViewModel.Context = vNewValue
End Property

Private Sub Class_Initialize()
    Set This.Notifier = New PropertyChangeNotifier
    Set This.FavoriteColors = New Scripting.Dictionary
    With This.FavoriteColors
        .Add Key:="Red", Item:="Red"
        .Add Key:="Green", Item:="Green"
        .Add Key:="Blue", Item:="Blue"
    End With
    This.FavoriteColor = "Green"
    
    Set This.FavFoodViewModel = New FavFoodViewModel
End Sub

Private Function IViewModel_IsValid() As Boolean
    IViewModel_IsValid = True
End Function

Private Function IViewModel_Self() As IViewModel
    Set IViewModel_Self = Me
End Function

Private Sub INotifyPropertyChanged_RegisterHandler(ByVal Handler As IHandlePropertyChanged)
    This.Notifier.RegisterHandler Handler
End Sub

Private Sub INotifyPropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    This.Notifier.OnPropertyChanged Source, PropertyName
End Sub

Private Sub OnPropertyChanged(ByVal PropertyName As String)
    INotifyPropertyChanged_OnPropertyChanged Me, PropertyName
End Sub

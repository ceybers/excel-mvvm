VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FavFoodViewModel"
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
    Handlers As Collection
    
    FavoriteFood As String
    FavoriteFoods As Scripting.Dictionary
End Type
Private This As TState

Public Property Get FavoriteFood() As String
    FavoriteFood = This.FavoriteFood
End Property

Public Property Let FavoriteFood(ByVal RHS As String)
    If This.FavoriteFood = RHS Then Exit Property
    This.FavoriteFood = RHS
    OnPropertyChanged "FavoriteFood"
End Property

Public Property Get FavoriteFoods() As Scripting.Dictionary
    Set FavoriteFoods = This.FavoriteFoods
End Property

'Public Property Set FavoriteFoods(ByVal RHS As Collection)
'    Set This.FavoriteFoods = RHS
'End Property

Public Property Get Context() As IAppContext
    Set Context = This.AppContext
End Property

Public Property Set Context(ByVal vNewValue As IAppContext)
    Set This.AppContext = vNewValue
End Property

Private Sub class_initialize()
    Set This.Handlers = New Collection
    Set This.FavoriteFoods = New Scripting.Dictionary
    With This.FavoriteFoods
        .Add Key:="Beef", Item:="Beef"
        .Add Key:="Chicken", Item:="Chicken"
        .Add Key:="Fish", Item:="Fish"
    End With
    This.FavoriteFood = "Chicken"
End Sub

Private Function IViewModel_IsValid() As Boolean
    IViewModel_IsValid = True
End Function

Private Function IViewModel_Self() As IViewModel
    Set IViewModel_Self = Me
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


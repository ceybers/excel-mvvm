VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CountryViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.Example2"
Option Explicit
Implements IViewModel
Implements INotifyPropertyChanged

Private Type TState
    Notifier As INotifyPropertyChanged
    
    Country As String
    Countries As Scripting.Dictionary
    
    CityViewModel As CityViewModel
End Type
Private This As TState

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

Private Sub Class_Initialize()
    Set This.Notifier = New PropertyChangeNotifier
    
    PopulateDefaultData
End Sub

'---
Public Property Get Country() As String
    Country = This.Country
End Property

Public Property Let Country(ByVal RHS As String)
    If This.Country = RHS Then Exit Property
    This.Country = RHS
    OnPropertyChanged "Country"
End Property

Public Property Get Countries() As Scripting.Dictionary
    Set Countries = This.Countries
End Property

Public Property Get CityViewModel() As CityViewModel
    Set CityViewModel = This.CityViewModel
End Property

Private Sub PopulateDefaultData()
    Set This.Countries = New Scripting.Dictionary
    
    Dim WorldModel As WorldModel
    Set WorldModel = New WorldModel
    WorldModel.Load
    
    Dim CountryName As Variant
    For Each CountryName In WorldModel.Countries.Keys
        With This.Countries
            .Add Key:=CountryName, Item:=WorldModel.Countries.Item(CountryName)
        End With
    Next CountryName

    Set This.CityViewModel = New CityViewModel
    This.CityViewModel.Load WorldModel
End Sub

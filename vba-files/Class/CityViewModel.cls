VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CityViewModel"
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
    
    City As String
    Cities As Scripting.Dictionary
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
Public Property Get City() As String
    City = This.City
End Property

Public Property Let City(ByVal RHS As String)
    If This.City = RHS Then Exit Property
    This.City = RHS
    OnPropertyChanged "City"
End Property

Public Property Get Cities() As Scripting.Dictionary
    Set Cities = This.Cities
End Property

Private Sub PopulateDefaultData()
    Set This.Cities = New Scripting.Dictionary
    'With This.Cities
    '    .Add Key:="Atlanta", Item:="Atlanta"
    '    .Add Key:="Dallas", Item:="Dallas"
    '    .Add Key:="Denver", Item:="Denver"
    'End With
    'This.City = "Atlanta"
End Sub

Public Sub Load(ByVal WorldModel As WorldModel)
    Dim Cities() As Variant
    Cities = WorldModel.Cities.Items
    
    Dim i As Long
    For i = LBound(Cities) To UBound(Cities)
        This.Cities.Add Key:=Cities(i).Name, Item:=Cities(i)
    Next i
End Sub

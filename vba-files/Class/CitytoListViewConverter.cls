VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CitytoListViewConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("MVVM.Example2")
Option Explicit
Implements IValueConverter

Private Sub IValueConverter_Convert(ByVal Source As Object, ByVal Target As Object)
    Dim Cities As Scripting.Dictionary
    Set Cities = Source
    Dim ListView As ListView
    Set ListView = Target
    
    Dim Key As Variant
    For Each Key In Cities
        Dim ThisCity As City
        Set ThisCity = Cities.Item(Key)
        Debug.Assert TypeOf ThisCity Is City
        
        Dim ListItem As ListItem
        Set ListItem = ListView.ListItems.Add(Key:=ThisCity.Name, Text:=ThisCity.Name)
        ListItem.ListSubItems.Add Text:=ThisCity.Country.Name
        ListItem.ListSubItems.Add Text:=ThisCity.Airport
    Next Key
End Sub

Private Sub IValueConverter_ConvertBack(ByVal Source As Object, ByVal Target As Object)
    'noop
End Sub

Public Sub InitializeListView(ByVal ListView As ListView)
    With ListView
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add Text:="City Name"
        .ColumnHeaders.Add Text:="Country Name"
        .ColumnHeaders.Add Text:="Airport Code"
        .Appearance = cc3D
        .BorderStyle = ccNone
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .FullRowSelect = True
    End With
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CountryToListViewConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("MVVM.Example2")
Option Explicit
Implements IValueConverter

Private Sub IValueConverter_Convert(ByVal Source As Object, ByVal Target As Object)
    Dim Countries As Scripting.Dictionary
    Set Countries = Source
    Dim ListView As ListView
    Set ListView = Target
    
    Dim Key As Variant
    For Each Key In Countries
        Dim ThisCountry As Country
        Set ThisCountry = Countries.Item(Key)
        Debug.Assert TypeOf ThisCountry Is Country
        
        Dim ListItem As ListItem
        Set ListItem = ListView.ListItems.Add(Key:=ThisCountry.Name, Text:=ThisCountry.Name)
    Next Key
End Sub

Private Sub IValueConverter_ConvertBack(ByVal Source As Object, ByVal Target As Object)
    'noop
End Sub

Public Sub InitializeListView(ByVal ListView As ListView)
    With ListView
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add Text:="Country Name"
        .View = lvwReport
        .Gridlines = True
        .LabelEdit = lvwManual
        .HideSelection = False
        .FullRowSelect = True
    End With
End Sub

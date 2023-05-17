Attribute VB_Name = "TestView"
'@Folder("View")
Option Explicit

Public Sub TestView()
    Dim ctx As AppContext
    Set ctx = New AppContext
    
    Dim view As IView
    Set view = New ExampleView
    
    Dim vm As IViewModel
    Set vm = GetSampleViewModel(ctx)
    
    With view
        If .ShowDialog(vm) Then
            'Debug.Print "OK"
        Else
            'Debug.Print "Cancelled"
        End If
    End With
End Sub

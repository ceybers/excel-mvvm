Attribute VB_Name = "modDoTest"
'@Folder "Worksheets"
Option Explicit

Private Const DO_DEBUG As Boolean = False

'@EntryPoint "DoTest"
Public Sub DoTest()
    Dim ctx As IAppContext
    Set ctx = New AppContext
    
    Dim View As IView
    Set View = New ExampleView
    
    Dim vm As SomeViewModel 'IViewModel
    Set vm = GetSampleViewModel(ctx)

    With View
        If .ShowDialog(vm) Then
            If DO_DEBUG Then Debug.Print "View.ShowDialog(vm) returned True"
        Else
            If DO_DEBUG Then Debug.Print "View.ShowDialog(vm) returned False"
        End If
    End With
End Sub

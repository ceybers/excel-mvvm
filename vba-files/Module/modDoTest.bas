Attribute VB_Name = "modDoTest"
'@Folder "Worksheets"
Option Explicit

Private Const DO_DEBUG As Boolean = False

'@EntryPoint "DoTest"
Public Sub DoTest()
    Dim ctx As IAppContext
    Set ctx = New AppContext
    
    Dim vm As CountryViewModel
    Set vm = New CountryViewModel
    
    Dim View As IView
    Set View = GeographyView.Create(ctx, vm)
    
    With View
        If .ShowDialog() Then
            If DO_DEBUG Then Debug.Print "GeographyView.ShowDialog(vm) returned True"
        Else
            If DO_DEBUG Then Debug.Print "GeographyView.ShowDialog(vm) returned False"
        End If
    End With
End Sub

Private Sub DoTestExample1()
    Dim ctx As IAppContext
    Set ctx = New AppContext
    
    Dim vm As SomeViewModel 'IViewModel
    Set vm = GetSampleViewModel(ctx)
    
    Dim inpc As INotifyPropertyChanged
    Set inpc = vm
    inpc.RegisterHandler ctx.BindingManager
    
    Dim View As IView
    Set View = ExampleView.Create(ctx, vm)
    
    'View.Hide
    'View.Show
    'View.ShowDialog
    'View.ViewModel
    
    'Dim ic As ICancellable
    'Set ic = View
    'ic.IsCancelled
    'ic.OnCancel
    
    With View
        If .ShowDialog() Then
            If DO_DEBUG Then Debug.Print "View.ShowDialog(vm) returned True"
        Else
            If DO_DEBUG Then Debug.Print "View.ShowDialog(vm) returned False"
        End If
    End With
End Sub

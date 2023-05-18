Attribute VB_Name = "RunExample1"
'@Folder "MVVM.Example1"
Option Explicit

Private Const DO_DEBUG As Boolean = False

Public Sub DoTestExample1()
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
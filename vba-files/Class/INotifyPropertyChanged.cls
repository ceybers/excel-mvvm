VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "INotifyPropertyChanged"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Abstractions"
Option Explicit

'@Description "Registers the specified handler."
Public Sub RegisterHandler(ByVal Handler As IHandlePropertyChanged)
End Sub

'@Description "Notifies all registered handlers of a property value change."
Public Sub OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
End Sub

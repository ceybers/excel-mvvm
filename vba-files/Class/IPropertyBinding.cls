VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPropertyBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Abstractions"
Option Explicit

Public Enum BindingMode
    TwoWayBinding
    OneWayBinding
    OneWayToSource
    OneTimeBinding 'NYI
End Enum

Public Property Get Mode() As BindingMode
End Property

Public Property Get Source() As Object
End Property

Public Property Get SourcePropertyPath() As String
End Property

Public Property Get Target() As Object
End Property

Public Property Get TargetProperty() As String
End Property
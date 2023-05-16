VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("ExcelMVVM")
Option Explicit

Public Type KeyColumn
    Selected As String
    Columns As Collection
End Type

Property Get Dirty() As Boolean

End Property

Property Let Dirty(ByVal RHS As Boolean)

End Property

Public Function LoadFromSettings()

End Function

Public Function LoadFromSomewhere()
    
End Function

Public Sub SaveToSettings()

End Sub

Public Function IsValid()

End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "AppContext"
Option Explicit

Private Type TAppContext
    BindingManager As BindingManager
End Type
Private This As TAppContext

Public Property Get BindingManager() As BindingManager
    Set BindingManager = This.BindingManager
End Property

Private Sub Class_Initialize()
    Set This.BindingManager = New BindingManager
End Sub
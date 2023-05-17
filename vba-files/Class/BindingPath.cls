VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BindingPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder MVVM.Infrastructure.Bindings
Option Explicit
Implements IBindingPath

Private Type TBindingPath
    Context As Object
    Path As String
    Object As Object
    PropertyName As String
End Type
Private This As TBindingPath

Public Property Get Context() As IViewModel
    Set Context = This.Context
End Property

Public Property Set Context(ByVal vNewValue As Object)
    Set This.Context = vNewValue
End Property

Public Property Get Path() As String
    Path = This.Path
End Property

Public Property Let Path(ByVal vNewValue As String)
    This.Path = vNewValue
End Property

Public Function Create(ByVal vContext As Object, ByVal vPath As String) As BindingPath
    Dim NewBindingPath As BindingPath
    Set NewBindingPath = New BindingPath
    With NewBindingPath
        Set .Context = vContext
        .Path = vPath
        .Resolve
    End With
    Set Create = NewBindingPath
End Function

Private Property Get IBindingPath_Context() As Object
    Set IBindingPath_Context = This.Context
End Property

Private Property Get IBindingPath_Path() As String
    IBindingPath_Path = This.Path
End Property

Private Property Get IBindingPath_Object() As Object
    Set IBindingPath_Object = This.Object
End Property

Private Property Get IBindingPath_PropertyName() As String
    IBindingPath_PropertyName = This.PropertyName
End Property

Public Sub Resolve()
    Set This.Object = ResolvePropertyPath(This.Context, This.Path)
    This.PropertyName = ResolvePropertyName(This.Path)
End Sub

Private Function ResolvePropertyName(ByVal PropertyPath As String) As String
    Dim Tokens() As String
    Tokens = Split(PropertyPath, ".")
    ResolvePropertyName = Tokens(UBound(Tokens))
End Function

Private Function ResolvePropertyPath(ByVal Context As Object, ByVal PropertyPath As String) As Object
    Dim Tokens() As String
    Tokens = Split(PropertyPath, ".")
    
    If UBound(Tokens) = 1 Then
        Dim actx As SomeViewModel
        Set actx = Context
        Set ResolvePropertyPath = CallByName(actx, Tokens(0), VbGet)
        Exit Function
    Else
        Dim ctx As SomeViewModel
        Set ctx = This.Context
        Dim NextObject As Object
        Set NextObject = CallByName(ctx, Tokens(0), VbGet)
        Debug.Assert Not NextObject Is Nothing
        
        Dim NextPath As String
        NextPath = Mid$(This.Path, Len(Tokens(0)) + 2)
        NextPath = Left$(NextPath, Len(NextPath) - Len(Tokens(UBound(Tokens))) - 1)
        Set This.Object = ResolvePropertyPath(NextObject, NextPath)
    End If
End Function

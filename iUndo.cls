VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iUndoObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'=================================================================
' Paul Leyten (C) March, 2001
' This code may be used freely, but please include my name in it.
'=================================================================
' This object is only a placeholder for the var's used in the undo collection
' This code isn't that hard..... so look for explanation in Form1 and Undobuffer.cls

Private mvarPrivateVariable As Variant 'local copy
Private mvarScreenObject As Object
Private mvarScreenObjectFunction As String
Private mvarOldValue As Variant 'local copy
Private mvarNewValue As Variant 'local copy
Public Property Let NewValue(ByVal vData As Variant)
    mvarNewValue = vData
End Property
Public Property Set NewValue(ByVal vData As Variant)
    Set mvarNewValue = vData
End Property
Public Property Get NewValue() As Variant
    If IsObject(mvarNewValue) Then
        Set NewValue = mvarNewValue
    Else
        NewValue = mvarNewValue
    End If
End Property
Public Property Let OldValue(ByVal vData As Variant)
    mvarOldValue = vData
    '=================================================================
    ' Paul Leyten (c) March. 2001
    ' This code may be used freely, but please include my name in it.
    '=================================================================
End Property
Public Property Set OldValue(ByVal vData As Variant)
    Set mvarOldValue = vData
End Property
Public Property Get OldValue() As Variant
    If IsObject(mvarOldValue) Then
        Set OldValue = mvarOldValue
    Else
        OldValue = mvarOldValue
    End If
End Property
Public Property Set ScreenObject(ByVal vData As Object)
    Set mvarScreenObject = vData
End Property
Public Property Get ScreenObject() As Object
    Set ScreenObject = mvarScreenObject
End Property
Public Property Let ScreenObjectFunction(ByVal vData As String)
    mvarScreenObjectFunction = vData
End Property
Public Property Get ScreenObjectFunction() As String
    ScreenObjectFunction = mvarScreenObjectFunction
End Property
Public Property Let PrivateVariable(ByVal vData As Variant)
    mvarPrivateVariable = vData
End Property
Public Property Set PrivateVariable(ByVal vData As Variant)
    Set mvarPrivateVariable = vData
End Property
Public Property Get PrivateVariable() As Variant
    If IsObject(mvarPrivateVariable) Then
        Set PrivateVariable = mvarPrivateVariable
    Else
        PrivateVariable = mvarPrivateVariable
    End If
End Property
'=================================================================
' Paul Leyten (c) March, 2001
' This code may be used freely,but please include my name in it.
'=================================================================
Private Sub Class_Initialize()
  mvarOldValue = Null
End Sub

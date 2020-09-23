VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UndoBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"iUndo"
Attribute VB_Ext_KEY = "Member0" ,"iUndo"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'=================================================================
' Paul Leyten (c) March, 2001
' This code may be used freely, but please include my name in it..
'=================================================================


' Events for this class
Public Event Undo(ByVal ScreenObject As Object, ByVal ScreenObjectFunction As String, OldValue As Variant)
Public Event Redo(ByVal ScreenObject As Object, ByVal ScreenObjectFunction As String, NewValue As Variant)
Public Event UndoEnabled(Enabled As Boolean)
Public Event RedoEnabled(Enabled As Boolean)

' Private Variables for this class
Private mCol As Collection          ' Contains the collection of iUndoObject's
Private lngPointerID As Long        ' the internal counter for the te collection

Public Function UndoChanges() As Long
    Dim ColItem As New iUndoObject
    
    'Set temportary object with the item of the internal collection
    Set ColItem = mCol(lngPointerID)
    
    'Raise the Undo Event with the screenobject, Functionname and the oldvalue
    RaiseEvent Undo(ColItem.ScreenObject, ColItem.ScreenObjectFunction, ColItem.OldValue)
    
    'decrease the internal pointer
    lngPointerID = lngPointerID - 1
    
    'raise the EnabledEvents
    'so the userinterface can enable or disable some buttons, menues etc..
    RaiseEvent UndoEnabled(lngPointerID > 0)
    RaiseEvent RedoEnabled(mCol.Count > 0)
    
    'Return  the internalpointer
    UndoChanges = lngPointerID
    
    'Set temporary object to nothing
    Set ColItem = Nothing
End Function
Public Function RedoChanges()
    Dim ColItem As New iUndoObject
    
    'Increase internal pointer
    lngPointerID = lngPointerID + 1
    
    'Set temportary object with the item of the internal collection
    Set ColItem = mCol(lngPointerID)
    
    'Raise the Redo Event with the screenobject, Functionname and the oldvalue
    RaiseEvent Redo(ColItem.ScreenObject, ColItem.ScreenObjectFunction, ColItem.NewValue)
    
    '=================================================================
    ' Paul Leyten (c) March, 2001
    ' This code may be used freely, but please include my name in it.
    '=================================================================
    
    
    'raise the EnabledEvents
    'so the userinterface can enable or disable some buttons, menues etc..
    RaiseEvent RedoEnabled(mCol.Count > lngPointerID)
    RaiseEvent UndoEnabled(lngPointerID > 0)
    
    'Return  the internalpointer
    RedoChanges = lngPointerID
    
    'Set temporary object to nothing
    Set ColItem = Nothing

End Function
Public Property Get UndoCollection() As Collection
    ' Can be used to dispalay the collection of changes
    ' In combination with the "Property Get Pointer()" it can be used to
    ' display the current undo/redo stuff...
    Set UndoCollection = mCol
End Property

Public Property Get Pointer() As Long
  ' returns the pointer of the current position of the redo/undo buffer
  Pointer = lngPointerID
End Property

' This function lets you add new undo-objects to the undobuffer.
Public Function Add(OldValue As Variant, NewValue As Variant, Optional PrivateVariable As Variant, Optional ByVal ScreenObject As Object, Optional ByVal ScreenObjectFunction As String) As Long
    'create a new object
    Dim objNewMember As iUndoObject
    Set objNewMember = New iUndoObject

    'set the properties passed into the method
    If IsObject(PrivateVariable) Then
        Set objNewMember.PrivateVariable = PrivateVariable
    Else
        objNewMember.PrivateVariable = PrivateVariable
    End If
    
    Set objNewMember.ScreenObject = ScreenObject
    objNewMember.ScreenObjectFunction = ScreenObjectFunction
    
    '=================================================================
    ' Paul Leyten (c) March, 2001
    ' This code may be used freely, but please include my name in it
    '=================================================================
    
    If IsObject(NewValue) Then
        Set objNewMember.NewValue = NewValue
    Else
        objNewMember.NewValue = NewValue
    End If
    
    If IsObject(OldValue) Then
      Set objNewMember.OldValue = OldValue
    Else
      objNewMember.OldValue = OldValue
    End If
        
       
    ' If we're not at the end of the collection -which means we've had done some undo's-
    ' we'll have to remove the remaining redo's at the end of the buffer.
    ' We'll keep the history of undo, but must destroy the redo buffer.
    If lngPointerID < mCol.Count Then
      While lngPointerID < mCol.Count
        mCol.Remove mCol.Count
      Wend
      lngPointerID = mCol.Count
      RaiseEvent RedoEnabled(False)
    End If
    
    'Add the new object to the buffer (collection thus...).
    mCol.Add objNewMember
    
    'Set the temporary object to nothing.
    Set objNewMember = Nothing
    
    'Increase the internal pointer
    lngPointerID = mCol.Count
    Add = mCol.Count
    If Add > 0 Then
        RaiseEvent UndoEnabled(True)
    End If
End Function
Public Property Get Count() As Long
    Count = mCol.Count
End Property
Private Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub
Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub
Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

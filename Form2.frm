VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=================================================================
' Paul Leyten (c) March, 2001
' This code may be used freely, but please include my name in it.
'=================================================================



' The code in thsi form is just for fun and debuggin.... so have fun. i won't explain it..
Public Property Set coll(prvCollection As Collection)
    List1.Clear
    Dim i As Integer
    For i = 1 To prvCollection.Count
        If IsArray(prvCollection(i).ScreenObject) Then
          List1.AddItem "Newvalue:" & prvCollection(i).NewValue & "; OldValue: " & prvCollection(i).OldValue & "; Screnobject: " & prvCollection(i).ScreenObject.Name & "(" & prvCollection(i).ScreenObject.Index & ")"
        Else
          List1.AddItem "Newvalue:" & prvCollection(i).NewValue & "; OldValue: " & prvCollection(i).OldValue & "; Screnobject: " & prvCollection(i).ScreenObject.Name
        End If
    Next
End Property
Private Sub Form_Activate()
  MDIForm1.UndoRedoBarVisible = False
End Sub
Private Sub Form_Resize()
  List1.Width = Me.ScaleWidth - List1.Left
  List1.Height = Me.ScaleHeight - List1.Top
End Sub

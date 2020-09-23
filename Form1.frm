VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3750
   ScaleWidth      =   7845
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtField 
      Height          =   345
      Index           =   1
      Left            =   1260
      TabIndex        =   0
      Top             =   630
      Width           =   6180
   End
   Begin VB.TextBox txtField 
      Height          =   345
      Index           =   2
      Left            =   1260
      TabIndex        =   1
      Top             =   1035
      Width           =   3150
   End
   Begin VB.TextBox txtField 
      Height          =   345
      Index           =   5
      Left            =   1260
      TabIndex        =   4
      Top             =   2355
      Width           =   2340
   End
   Begin VB.TextBox txtField 
      Height          =   345
      Index           =   7
      Left            =   1260
      TabIndex        =   5
      Top             =   2805
      Width           =   2340
   End
   Begin VB.TextBox txtField 
      Height          =   345
      Index           =   8
      Left            =   1260
      TabIndex        =   6
      Top             =   3270
      Width           =   2340
   End
   Begin VB.TextBox txtField 
      Height          =   345
      Index           =   6
      Left            =   5100
      TabIndex        =   7
      Top             =   2355
      Width           =   1800
   End
   Begin VB.ComboBox cboCountry 
      Height          =   315
      Left            =   5100
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2820
      Width           =   1800
   End
   Begin VB.TextBox txtField 
      Height          =   345
      Index           =   9
      Left            =   5100
      TabIndex        =   9
      Top             =   3270
      Width           =   1800
   End
   Begin VB.TextBox txtField 
      Height          =   720
      Index           =   4
      Left            =   1260
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1455
      Width           =   6180
   End
   Begin VB.TextBox txtField 
      Height          =   345
      Index           =   3
      Left            =   5100
      TabIndex        =   2
      Top             =   1035
      Width           =   2340
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   120
      Picture         =   "Form1.frx":0000
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Company Name:"
      Height          =   195
      Index           =   1
      Left            =   15
      TabIndex        =   20
      Top             =   675
      Width           =   1185
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Contact Name:"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   19
      Top             =   1110
      Width           =   1080
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "City:"
      Height          =   195
      Index           =   5
      Left            =   870
      TabIndex        =   18
      Top             =   2430
      Width           =   345
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Region:"
      Height          =   195
      Index           =   6
      Left            =   4485
      TabIndex        =   17
      Top             =   2430
      Width           =   555
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Postal Code:"
      Height          =   195
      Index           =   7
      Left            =   300
      TabIndex        =   16
      Top             =   2880
      Width           =   915
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Country:"
      Height          =   195
      Index           =   8
      Left            =   4395
      TabIndex        =   15
      Top             =   2880
      Width           =   645
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Phone:"
      Height          =   195
      Index           =   9
      Left            =   705
      TabIndex        =   14
      Top             =   3345
      Width           =   510
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Fax:"
      Height          =   195
      Index           =   10
      Left            =   4710
      TabIndex        =   13
      Top             =   3345
      Width           =   330
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Address:"
      Height          =   195
      Index           =   4
      Left            =   555
      TabIndex        =   12
      Top             =   1515
      Width           =   645
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Title:"
      Height          =   195
      Index           =   2
      Left            =   4680
      TabIndex        =   11
      Top             =   1110
      Width           =   360
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000003&
      Caption         =   "Customers   "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   525
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7710
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "Form1.frx":0442
      Top             =   30
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=================================================================
' Paul Leyten (c) March, 2001
' This code may be used freely, but please include my name in it.
'=================================================================

' Declare withevents the UndoBuffer
Dim WithEvents UndoBuffer As prjUndo.UndoBuffer
Attribute UndoBuffer.VB_VarHelpID = -1

'To hold the original values of the indexed textfields, use two arrays
Private mvarChanged() As Boolean
Private mvarFields() As Variant

'To hold the original values of the ComboBox, use two Vars
Private mvarcboChanged As Boolean
Private mvarCboFields As Variant

' Just an function to show the collection. Not needed in any program
Private Sub ShowColl()
    Set Form2.coll = UndoBuffer.UndoCollection
End Sub

' Just to let it look a little bit more profi...
' Not needed to demonstrate the Undo/Redo buffer
Private Sub Form_Resize()
    lblName.Width = Me.Width
End Sub

' This MDIChild has undo/redo facilities, so show the buttons and menu's
' you can do this by calling:
' mdiform1.mnuEditUndo.Visible = true.........mdiform1.toolbar1.buttons("...").enabled.... etc
' it's nicer to do this with a property let (or function or sub or whatever you like..)
Private Sub Form_Activate()
  MDIForm1.UndoRedoBarVisible = True
End Sub
Private Sub Form_Load()

'Fill in the combo just to test...
cboCountry.AddItem "NL"
cboCountry.AddItem "DE"
cboCountry.AddItem "BE"
cboCountry.AddItem "FR"
cboCountry.AddItem "GB"
cboCountry.AddItem "USA"

  Dim i As Integer
  Dim ctl As Control
  
  'Loop through the form and define the changed boolean and set it to false
  ' and set the initial values in the mvarFields(index) fields, so the
  ' original value can be passed to the undo buffer.
  Set UndoBuffer = New prjUndo.UndoBuffer
  For Each ctl In Me.Controls
    If TypeOf ctl Is TextBox Then
      ReDim Preserve mvarFields(i)
      ReDim Preserve mvarChanged(i)
      mvarFields(i) = ""
      mvarChanged(i) = False
      i = i + 1
    End If
  Next
  
  'This can also be done in above loop....
  mvarCboFields = cboCountry.ListIndex
  mvarcboChanged = False
End Sub

' If a change event is fired, the textbox is changed
' You can use this to detect it with the Validate() event
Private Sub txtField_Change(Index As Integer)
    mvarChanged(Index) = True
End Sub

' When a user tabs through the fields and the CausesValidation property is set to true
' this function will be passed when a user leaves the field
' I think it's a better function then LostFocus, because of the cancel boolean :-)
Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
  If mvarChanged(Index) Then
    ' If Changed, then add a new row in the undo/redo buffer....
    Me.Caption = "Count in UNDO-Buffer: " & UndoBuffer.Add(mvarFields(Index), txtField(Index).Text, , txtField(Index), "Text")
    
    ' Just for test purposes show the undo collection in form2.
    ShowColl
    
    ' pass the new value in the mvarFields array
    mvarFields(Index) = txtField(Index)
  End If
  
  'The field isn't changed anymore...yet..
  mvarChanged(Index) = False
End Sub
' If a click event is fired, the Combobox is changed
' You can use this to detect it with the Validate() event
Private Sub cboCountry_Click()
  mvarcboChanged = True
End Sub

' When a user tabs through the fields and the CausesValidation property is set to true
' this function will be passed when a user leaves the field
' I think it's a better function then LostFocus, because of the cancel boolean :-)
Private Sub cboCountry_Validate(Cancel As Boolean)
  If mvarcboChanged Then
   ' If Changed, then add a new row in the undo/redo buffer....
   ' In this case we use the Listindex property instead of the text property which we use
   ' in the textbox functions.
    Me.Caption = "Count in UNDO-Buffer: " & UndoBuffer.Add(mvarCboFields, cboCountry.ListIndex, , cboCountry, "ListIndex")
    
    ' Just for test purposes show the undo collection in form2.
    ShowColl
    
  ' pass the new value in the mvarFields array
    mvarCboFields = cboCountry.ListIndex
  End If
  mvarcboChanged = False
End Sub

' Because we're working in a MDI-Child modus and the Undo/Redo buttons resist on
' the MDIForm's Toolbar, we have to have public subs which can be called by the toolbar's events
Public Sub Undo()
    ' Set the caption of the form with the returning pointer of the buffer.
    ' Of course you don't have to use the return value :-)
    Me.Caption = "Count in UNDO-Buffer: " & UndoBuffer.UndoChanges
End Sub

' When the .UndoChanges sub is called, the UndoBuffer will raise an event
' This is the place to catch the event.
Private Sub UndoBuffer_Undo(ByVal ScreenObject As Object, ByVal ScreenObjectFunction As String, OldValue As Variant)
    ' Here we will call the CallByName function which can be used to call functions by Name
    ' In the Add function (see validate()) we'll put the Screenobject (e.g. TextBox..), the Oldvalue, Newvalue
    ' and the Property of the object (e.g. "Text", or "ListIndex")
    ' These are all "vbLet" functions
    CallByName ScreenObject, ScreenObjectFunction, VbLet, OldValue
    ScreenObject.SetFocus
    'For testing purposes, let's see the collection in form2
    Call ShowColl
End Sub

' An other event is fired so we can trigger of the buttons have to be shown or not.
' So this we will pass to the MDIParent object, in which resides or menu structure and
' Toolbar buttons.
Private Sub UndoBuffer_UndoEnabled(Enabled As Boolean)
    MDIForm1.UndoEnabled = Enabled
End Sub

' Because we're working in a MDI-Child modus and the Undo/Redo buttons resist on
' the MDIForm's Toolbar, we have to have public subs which can be called by the toolbar's events
Public Sub Redo()
    ' Set the caption of the form with the returning pointer of the buffer.
    ' Of course you don't have to use the return value :-)
    Me.Caption = "Count in UNDO-Buffer: " & UndoBuffer.RedoChanges
End Sub

' When the .RedoChanges sub is called, the UndoBuffer will raise an event
' This is the place to catch the event.
Private Sub UndoBuffer_Redo(ByVal ScreenObject As Object, ByVal ScreenObjectFunction As String, NewValue As Variant)
    ' Here we will call the CallByName function which can be used to call functions by Name
    ' In the Add function (see validate()) we'll put the Screenobject (e.g. TextBox..), the Oldvalue, Newvalue
    ' and the Property of the object (e.g. "Text", or "ListIndex")
    ' These are all "vbLet" functions
    CallByName ScreenObject, ScreenObjectFunction, VbLet, NewValue
    ScreenObject.SetFocus
    'For testing purposes, let's see the collection in form2
    Call ShowColl
End Sub

' An other event is fired so we can trigger of the buttons have to be shown or not.
' So this we will pass to the MDIParent object, in which resides or menu structure and
' Toolbar buttons.
Private Sub UndoBuffer_RedoEnabled(Enabled As Boolean)
    MDIForm1.RedoEnabled = Enabled
End Sub

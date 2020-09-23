VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   5580
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   12120
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12120
      _ExtentX        =   21378
      _ExtentY        =   794
      BandCount       =   2
      _CBWidth        =   12120
      _CBHeight       =   450
      _Version        =   "6.7.8862"
      Child1          =   "Toolbar1"
      MinHeight1      =   390
      Width1          =   1005
      NewRow1         =   0   'False
      MinHeight2      =   360
      Width2          =   1440
      NewRow2         =   0   'False
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   5760
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":0000
               Key             =   "undo"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIForm1.frx":0114
               Key             =   "redo"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   390
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         Appearance      =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "undo"
               Object.ToolTipText     =   "Undo"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "redo"
               Object.ToolTipText     =   "Redo"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "&Redo"
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
      Begin VB.Menu mnuWindowBar1 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About MyApp..."
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form is just a simple test form.
' The most of the code happens to be in Form1 and in the UndoBuffer Class.
' They are (i think) well documented, so please look there for help.
' Have fun and succes..!

'=================================================================
' Paul Leyten (c) March, 2001
' This code may be used freely, but please include my name in it.
'=================================================================

Private Sub mnuHelpAbout_Click()
  MsgBox "This demo demonstrates the Undo/Redo buffer created by JP Leyten"
End Sub
Private Sub mnuWindowArrangeIcons_Click()
  Me.Arrange vbArrangeIcons
End Sub
Private Sub mnuWindowCascade_Click()
  Me.Arrange vbCascade
End Sub
Private Sub mnuWindowTileHorizontal_Click()
  Me.Arrange vbTileHorizontal
End Sub
Private Sub mnuWindowTileVertical_Click()
  Me.Arrange vbTileVertical
End Sub
Private Sub mnuEditCopy_Click()
  MsgBox "Place Copy Code here!"
End Sub
Private Sub mnuEditCut_Click()
  MsgBox "Place Cut Code here!"
End Sub
Private Sub mnuEditPaste_Click()
  MsgBox "Place Paste Code here!"
End Sub
Private Sub mnuEditUndo_Click()
  MsgBox "Place Undo Code here!"
End Sub
Private Sub mnuFileExit_Click()
  Unload Me
End Sub
Private Sub MDIForm_Load()
    Me.WindowState = vbMaximized
    Load Form2
    Load Form1
    Me.Arrange vbTileHorizontal
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "undo"
        ActiveForm.Undo
    Case "redo"
        ActiveForm.Redo
End Select
End Sub
Property Let RedoEnabled(vdata As Boolean)
  Toolbar1.Buttons("redo").Enabled = vdata
  mnuRedo.Enabled = vdata
End Property
Property Let UndoEnabled(vdata As Boolean)
  Toolbar1.Buttons("undo").Enabled = vdata
  mnuEditUndo.Enabled = vdata
End Property
Property Let UndoRedoBarVisible(vdata As Boolean)
  CoolBar1.Bands(1).Visible = vdata
End Property


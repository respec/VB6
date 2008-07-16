VERSION 5.00
Begin VB.UserControl ATCoSelectListSorted 
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6675
   ScaleHeight     =   2400
   ScaleWidth      =   6675
   ToolboxBitmap   =   "ATCoSelectListSorted.ctx":0000
   Begin VB.CommandButton cmdMove 
      Height          =   615
      Index           =   1
      Left            =   6120
      Picture         =   "ATCoSelectListSorted.ctx":0312
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Move Item Down In List"
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cmdMove 
      Height          =   615
      Index           =   0
      Left            =   6120
      Picture         =   "ATCoSelectListSorted.ctx":0754
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Move Item Up In List"
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdMoveAllLeft 
      Caption         =   "<<- Remove All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdMoveAllRight 
      Caption         =   "Add All ->>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdMoveLeft 
      Caption         =   "<- Remove"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdMoveRight 
      Caption         =   "Add ->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.ListBox lstRight 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      Left            =   3960
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   360
      Width           =   2055
   End
   Begin VB.ListBox lstLeft 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      Left            =   0
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label lblRight 
      Caption         =   "Selected:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   9
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblLeft 
      Caption         =   "Available:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "ATCoSelectListSorted"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Copyright 2000 by AQUA TERRA Consultants

  '##MODULE_NAME AtcoSelectListSorted
  '##MODULE_DATE December 14, 2000
  '##MODULE_AUTHOR Mark Gray and Robert Dusenbury or AQUA TERRA CONSULTANTS
  '##MODULE_SUMMARY <P>This control is a dual-paned selection tool with a list box in each _
          pane. The items in each list box are sorted by ascending alphanumeric order of their name.
' ##MODULE_REMARKS This control was made from a combination of standard VB _
          tools.&nbsp;The left pane is titled <EM>Available</EM> and&nbsp;contains a list _
          of items that may be, but are not as of yet, selected. The right pane is titled _
          <EM>Selected</EM> and&nbsp;contains a list of such items. _
          <P></P> _
          <P>There are 4 tools located between the left and right panes that allow the _
          user to transfer items between list boxes: a single right arrow labled _
          <EM>Add</EM> and a single left arrow labled&nbsp;<EM>Remove</EM>&nbsp;that _
          transfer only selected items to or from the <EM>Selected</EM> list box, and _
          double right arrows labled <EM>Add All</EM>&nbsp;and double left _
          arrows&nbsp;labled <EM>Remove All</EM>&nbsp;that tranfer all items to or from _
          the <EM>Selected</EM> list box.</P> _
          <P>Additionally, there is an up and a down arrow on the right side of the _
          control that allows the user to move <EM>Selected</EM> items&nbsp;up or down in _
          the list rank.&nbsp;</P>

Public Event Change()

Public Property Let ButtonVisible(ByVal butt As Variant, ByVal NewValue As Boolean)
Attribute ButtonVisible.VB_Description = "Boolean determining whether specified button is visible or not."
  '##SUMMARY Boolean determining whether specified button is visible or not.
  '##PARAM butt (I) Index or Name identifying one of 6 buttons on control.
  '##PARAM NewValue (I) Boolean setting assigned to Visible property of button.
  Dim label As String 'text identifying one of the 6 buttons on the control
  label = LCase(butt)
  Select Case butt
    Case "1", "add", "move right":           cmdMoveRight.Visible = NewValue
    Case "2", "remove", "move left":         cmdMoveLeft.Visible = NewValue
    Case "3", "add all", "move all right":   cmdMoveAllRight.Visible = NewValue
    Case "4", "remove all", "move all left": cmdMoveAllLeft.Visible = NewValue
    Case "5", "move up":                     cmdMove(0).Visible = NewValue
    Case "6", "move down":                   cmdMove(1).Visible = NewValue
  End Select
End Property
Public Property Get ButtonVisible(ByVal butt As Variant) As Boolean
  Dim label As String 'text identifying one of the 6 buttons on the control
  label = LCase(butt)
  Select Case butt
    Case "1", "add":        ButtonVisible = cmdMoveRight.Visible
    Case "2", "remove":     ButtonVisible = cmdMoveLeft.Visible
    Case "3", "add all":    ButtonVisible = cmdMoveAllRight.Visible
    Case "4", "remove all": ButtonVisible = cmdMoveAllLeft.Visible
    Case "5", "move up":    ButtonVisible = cmdMove(0).Visible
    Case "6", "move down":  ButtonVisible = cmdMove(1).Visible
  End Select
End Property

Public Property Let RightItem(ByVal i As Long, ByVal NewValue As String)
Attribute RightItem.VB_Description = "Name of ith item in right Selected ListBox."
  '##SUMMARY Name of <EM>i</EM>th item in right <EM>Selected</EM> ListBox.
  '##PARAM i (I) Index identifying sequential member in right <EM>Selected</EM> ListBox.
  '##PARAM NewValue (I) Name assigned to <EM>i</EM>th item in right <EM>Selected</EM> ListBox.
  If Len(NewValue) > 0 And Not InRightList(NewValue) Then
    If i = lstRight.ListCount Then lstRight.AddItem NewValue
    If i >= 0 And i < lstRight.ListCount Then lstRight.List(i) = NewValue
  End If
End Property
Public Property Get RightItem(ByVal i As Long) As String
  If i >= 0 And i < lstRight.ListCount Then
    RightItem = lstRight.List(i)
  Else
    RightItem = ""
  End If
End Property

Public Property Let LeftItem(ByVal i As Long, ByVal NewValue As String)
Attribute LeftItem.VB_Description = "Name of ith item in left Available ListBox."
  '##SUMMARY Name of <EM>i</EM>th item in left <EM>Available</EM> ListBox.
  '##PARAM i (I) Index identifying sequential member in left <EM>Available</EM> ListBox.
  '##PARAM NewValue (I) Name assigned to <EM>i</EM>th item in left <EM>Available</EM> ListBox.
  If Not InLeftList(NewValue) Then
    If i = lstLeft.ListCount Then lstLeft.AddItem NewValue
    If i >= 0 And i < lstLeft.ListCount Then lstLeft.List(i) = NewValue
  End If
End Property
Public Property Get LeftItem(ByVal i As Long) As String
  If i >= 0 And i < lstLeft.ListCount Then
    LeftItem = lstLeft.List(i)
  Else
    LeftItem = ""
  End If
End Property

Public Sub LeftItemFastAdd(ByVal NewValue As String)
Attribute LeftItemFastAdd.VB_Description = "Used in place of Let LeftItem to speed addition of hundreds of items. If you are adding fewer than 100 items, use LeftItem property instead."
  '##SUMMARY Used in place of Let LeftItem to speed addition of hundreds of items. _
    If you are adding fewer than 100 items, use LeftItem property instead.
  '##PARAM NewValue (I) Name assigned to new item in left <EM>Available</EM> ListBox.
  lstLeft.AddItem NewValue
End Sub

Public Property Let RightItemData(ByVal i As Long, ByVal NewValue As Long)
Attribute RightItemData.VB_Description = "Integer set as property of ith item in right Selected ListBox."
  '##SUMMARY Integer set as property of <EM>i</EM>th item in right <EM>Selected</EM> ListBox.
  '##PARAM i (I) Index identifying sequential member in right <EM>Selected</EM> ListBox.
  '##PARAM NewValue (I) Integer assigned to ItemData property of <EM>i</EM>th item in right <EM>Selected</EM> ListBox.
  lstRight.ItemData(i) = NewValue
End Property
Public Property Get RightItemData(ByVal i As Long) As Long
  RightItemData = lstRight.ItemData(i)
End Property

Public Property Let LeftItemData(ByVal i As Long, ByVal NewValue As Long)
Attribute LeftItemData.VB_Description = "Integer set as property of ith item in left Available ListBox."
  '##SUMMARY Integer set as property of <EM>i</EM>th item in left <EM>Available</EM> ListBox.
  '##PARAM i (I) Index identifying sequential member in left <EM>Available</EM> ListBox.
  '##PARAM NewValue (I) Integer assigned to ItemData property of <EM>i</EM>th item in left <EM>Available</EM> ListBox.
  lstLeft.ItemData(i) = NewValue
End Property
Public Property Get LeftItemData(ByVal i As Long) As Long
  LeftItemData = lstLeft.ItemData(i)
End Property

Public Property Get RightCount() As Long
Attribute RightCount.VB_Description = "Number of items in right Selected ListBox."
  '##SUMMARY Number of items in right <EM>Selected</EM> ListBox.
  RightCount = lstRight.ListCount
End Property

Public Property Get LeftCount() As Long
Attribute LeftCount.VB_Description = "Number of items in left Available ListBox."
  '##SUMMARY Number of items in left <EM>Available</EM> ListBox.
  LeftCount = lstLeft.ListCount
End Property

Public Property Let RightLabel(ByVal NewValue As String)
Attribute RightLabel.VB_Description = "Name of right ListBox; Selected is default."
  '##SUMMARY Name of right ListBox; <EM>Selected</EM> is default.
  '##PARAM NewValue (I) Name assigned to Caption property of right ListBox.
  lblRight.Caption = NewValue
End Property
Public Property Get RightLabel() As String
  RightLabel = lblRight.Caption
End Property

Public Property Let LeftLabel(ByVal NewValue As String)
Attribute LeftLabel.VB_Description = "Name of left ListBox; Available is default."
  '##SUMMARY Name of left ListBox; <EM>Available</EM> is default.
  '##PARAM NewValue (I) Name assigned to Caption property of left ListBox.
  lblLeft.Caption = NewValue
End Property
Public Property Get LeftLabel() As String
  LeftLabel = lblLeft.Caption
End Property

Public Property Let MoveUpTip(ByVal NewValue As String)
Attribute MoveUpTip.VB_Description = "Pop-up advice when cursor held over 'up arrow' on right side of control; "
  '##SUMMARY Pop-up advice when cursor held over 'up arrow' on right side of _
    control; "Move Item Up In List" is default.
  '##PARAM NewValue (I) Text assigned to ToolTipText property of 'up arrow' button.
  cmdMove(0).ToolTipText = NewValue
End Property
Public Property Get MoveUpTip() As String
  MoveUpTip = cmdMove(0).ToolTipText
End Property

Public Property Let MoveDownTip(ByVal NewValue As String)
Attribute MoveDownTip.VB_Description = "Pop-up advice when cursor held over 'down arrow' on right side of control; "
  '##SUMMARY Pop-up advice when cursor held over 'down arrow' on right side of _
    control; "Move Item Down In List" is default.
  '##PARAM NewValue (I) Text assigned to ToolTipText property of 'down arrow' button.
  cmdMove(1).ToolTipText = NewValue
End Property
Public Property Get MoveDownTip() As String
  MoveDownTip = cmdMove(1).ToolTipText
End Property

Public Sub MoveRight(ByVal i As Long)
Attribute MoveRight.VB_Description = "Moves selected item from left Available to right Selected ListBox."
  '##SUMMARY Moves selected item from left <EM>Available</EM> to right <EM>Selected</EM> ListBox.
  '##PARAM i (I) Integer identifying <EM>i</EM>th item in left <EM>Available</EM> ListBox.
  If i >= 0 And i < lstLeft.ListCount Then
    lstRight.AddItem lstLeft.List(i)
    lstRight.ItemData(lstRight.ListCount - 1) = lstLeft.ItemData(i)
    lstLeft.RemoveItem i
  End If
  RaiseEvent Change
End Sub

Public Sub MoveLeft(ByVal i As Long)
Attribute MoveLeft.VB_Description = "Moves selected item from right Selected to left Available ListBox."
  '##SUMMARY Moves selected item from right <EM>Selected</EM> to left <EM>Available</EM> ListBox.
  '##PARAM i (I) Integer identifying <EM>i</EM>th item in right <EM>Selected</EM> ListBox.
  If i >= 0 And i < lstRight.ListCount Then
    lstLeft.AddItem lstRight.List(i)
    lstLeft.ItemData(lstLeft.ListCount - 1) = lstRight.ItemData(i)
    lstRight.RemoveItem i
  End If
  RaiseEvent Change
End Sub

Public Sub MoveAllRight()
Attribute MoveAllRight.VB_Description = "Moves all items from left Available to right Selected ListBox."
  '##SUMMARY Moves all items from left <EM>Available</EM> to right <EM>Selected</EM> ListBox.
  Dim i As Long 'Index used to loop thru items in left ListBox
  For i = 0 To lstLeft.ListCount - 1
    lstRight.AddItem lstLeft.List(i)
    lstRight.ItemData(lstRight.ListCount - 1) = lstLeft.ItemData(i)
  Next i
  lstLeft.Clear
  RaiseEvent Change
End Sub

Public Sub MoveAllLeft()
Attribute MoveAllLeft.VB_Description = "Moves all items from right Selected to left Available ListBox."
  '##SUMMARY Moves all items from right <EM>Selected</EM> to left <EM>Available</EM> ListBox.
  Dim i As Long 'Index used to loop thru items in right ListBox
  For i = 0 To lstRight.ListCount - 1
    lstLeft.AddItem lstRight.List(i)
    lstLeft.ItemData(lstLeft.ListCount - 1) = lstRight.ItemData(i)
  Next i
  lstRight.Clear
  RaiseEvent Change
End Sub

Public Sub ClearRight()
Attribute ClearRight.VB_Description = "Removes all items from right Selected ListBox."
  '##SUMMARY Removes all items from right <EM>Selected</EM> ListBox.
  lstRight.Clear
  RaiseEvent Change
End Sub

Public Sub ClearLeft()
Attribute ClearLeft.VB_Description = "Removes all items from left Available ListBox."
  '##SUMMARY Removes all items from left <EM>Available</EM> ListBox.
  lstLeft.Clear
  RaiseEvent Change
End Sub

Public Function InRightList(ByVal search As String) As Boolean
Attribute InRightList.VB_Description = "Boolean check whether specified item is in right Selected ListBox."
  '##SUMMARY Boolean check whether specified item is in right <EM>Selected</EM> ListBox.
  '##PARAM search (I) Name of item to search for in right <EM>Selected</EM> ListBox.
  '##RETURNS True if incoming argument appears as item in right ListBox.
  InRightList = InList(search, lstRight)
End Function

Public Function InLeftList(ByVal search As String) As Boolean
Attribute InLeftList.VB_Description = "Boolean check whether specified item is in left Available ListBox."
  '##SUMMARY Boolean check whether specified item is in left <EM>Available</EM> ListBox.
  '##PARAM search (I) Name of item to search for in left <EM>Available</EM> ListBox.
  '##RETURNS True if incoming argument appears as item in left ListBox.
  InLeftList = InList(search, lstLeft)
End Function

Private Function InList(ByVal s As String, Lst As Object) As Boolean
Attribute InList.VB_Description = "Boolean check whether specified string occurs as item in ListBox."
  '##SUMMARY Boolean check whether specified string occurs as item in ListBox.
  '##PARAM s (I) String item for which lst object will be searched.
  '##PARAM lst (I) List object to be searched thru for specified string item.
  '##RETURNS True if 'S' appears as item in 'Lst'.
    Dim i As Long 'Index used to loop thru items in lst object
    Dim found As Boolean 'True if item is found in list
    
    i = 0
    found = False
    Do While Not found
      If s = Lst.List(i) Then
        found = True
      ElseIf i < Lst.ListCount - 1 Then
        i = i + 1
      Else
        Exit Do
      End If
    Loop
    
    InList = found
    
End Function

Private Sub cmdMove_Click(index As Integer)
Attribute cmdMove_Click.VB_Description = "Moves selected item in right Selected ListBox either up or down one notch in list."
  '##SUMMARY Moves selected item in right <EM>Selected</EM> ListBox either up or down one notch in list.
  '##PARAM Index (I) Integer identifying whether to move selected item up (0) or down (1).
  Dim i As Long 'Index used to loop thru items in right <EM>Selected</EM> ListBox.
  Dim tmp As String 'Used to hold names of list members when reordering.
  Dim tmpData As Long 'Used to hold value of ItemData property of list members when reordering.
  
  If index = 0 Then 'Move Up
    For i = 1 To lstRight.ListCount - 1
      If lstRight.Selected(i) And Not lstRight.Selected(i - 1) Then
        
        tmp = lstRight.List(i - 1)
        tmpData = lstRight.ItemData(i - 1)
        
        lstRight.List(i - 1) = lstRight.List(i)
        lstRight.ItemData(i - 1) = lstRight.ItemData(i)
        
        lstRight.List(i) = tmp
        lstRight.ItemData(i) = tmpData
        
        lstRight.Selected(i) = False
        lstRight.Selected(i - 1) = True
        RaiseEvent Change
      End If
    Next i
  ElseIf index = 1 Then 'Move Down
    For i = lstRight.ListCount - 2 To 0 Step -1
      If lstRight.Selected(i) And Not lstRight.Selected(i + 1) Then
        tmp = lstRight.List(i + 1)
        tmpData = lstRight.ItemData(i + 1)
        
        lstRight.List(i + 1) = lstRight.List(i)
        lstRight.ItemData(i + 1) = lstRight.ItemData(i)
        
        lstRight.List(i) = tmp
        lstRight.ItemData(i) = tmpData
        
        lstRight.Selected(i) = False
        lstRight.Selected(i + 1) = True
        RaiseEvent Change
      End If
    Next i
  End If
End Sub

Private Sub cmdMoveAllLeft_Click()
Attribute cmdMoveAllLeft_Click.VB_Description = "Moves all items from right Selected to left Available ListBox."
  '##SUMMARY Moves all items from right <EM>Selected</EM> to left <EM>Available</EM> ListBox.
  MoveAllLeft
End Sub

Private Sub cmdMoveAllRight_Click()
Attribute cmdMoveAllRight_Click.VB_Description = "Moves all items from left Available to right Selected ListBox."
  '##SUMMARY Moves all items from left <EM>Available</EM> to right <EM>Selected</EM> ListBox.
  MoveAllRight
End Sub

Private Sub cmdMoveLeft_Click()
Attribute cmdMoveLeft_Click.VB_Description = "Moves selected item from right Selected to left Available ListBox."
  '##SUMMARY Moves selected item from right <EM>Selected</EM> to left <EM>Available</EM> ListBox.
  Dim i As Long 'Index used to loop thru items in right <EM>Selected</EM> ListBox.
  i = 0
  While i < lstRight.ListCount
    If lstRight.Selected(i) Then MoveLeft i Else i = i + 1
  Wend
End Sub

Private Sub cmdMoveRight_Click()
Attribute cmdMoveRight_Click.VB_Description = "Moves selected item from left Available to right Selected ListBox."
  '##SUMMARY Moves selected item from left <EM>Available</EM> to right <EM>Selected</EM> ListBox.
  Dim i As Long 'Index used to loop thru items in left <EM>Available</EM> ListBox.
  i = 0
  While i < lstLeft.ListCount
    If lstLeft.Selected(i) Then MoveRight i Else i = i + 1
  Wend
End Sub

Private Sub lstRight_DblClick()
Attribute lstRight_DblClick.VB_Description = "Moves double-clicked item from right Selected to left Available ListBox."
  '##SUMMARY Moves double-clicked item from right <EM>Selected</EM> to left <EM>Available</EM> ListBox.
  MoveLeft lstRight.ListIndex
End Sub

Private Sub lstLeft_DblClick()
Attribute lstLeft_DblClick.VB_Description = "Moves double-clicked item from left Available to right Selected ListBox."
  '##SUMMARY Moves double-clicked item from left <EM>Available</EM> to right <EM>Selected</EM> ListBox.
  MoveRight lstLeft.ListIndex
End Sub

Private Sub UserControl_Resize()
Attribute UserControl_Resize.VB_Description = "Resizes components depending on how user sizes control."
  '##SUMMARY Resizes components depending on how user sizes control.
  Dim UsedWidth As Long 'Quantifies width of control being used by components
  Dim margin As Long 'Quantifies size of margin
  margin = 225
  UsedWidth = cmdMoveRight.Width + cmdMove(0).Width + margin * 3
  'Adjust Width of components
  If Width - UsedWidth > 1000 Then
    lstLeft.Width = (Width - UsedWidth) / 2
    lblLeft.Width = lstLeft.Width - (lblLeft.Left - lstLeft.Left)
    lstRight.Width = lstLeft.Width
    cmdMoveRight.Left = lstLeft.Left + lstLeft.Width + margin
    cmdMoveLeft.Left = cmdMoveRight.Left
    cmdMoveAllRight.Left = cmdMoveRight.Left
    cmdMoveAllLeft.Left = cmdMoveRight.Left
    lstRight.Left = cmdMoveRight.Left + cmdMoveRight.Width + margin
    lblRight.Left = lstRight.Left + (lblLeft.Left - lstLeft.Left)
    lblRight.Width = lstRight.Width - (lblRight.Left - lstRight.Left)
    cmdMove(0).Left = lstRight.Left + lstRight.Width + margin
    cmdMove(1).Left = cmdMove(0).Left
  End If
  'Adjust Height of components
  If Height > cmdMoveAllLeft.Top + cmdMoveAllLeft.Height Then
    lstLeft.Height = Height - 400
    lstRight.Height = lstLeft.Height
    cmdMove(1).Top = lstRight.Top + lstRight.Height - cmdMove(1).Height - (cmdMove(0).Top - lstRight.Top)
  End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Attribute UserControl_ReadProperties.VB_Description = "Reads LeftLabel and RightLabel properties of left and right ListBoxes, repectively, and sets label caption of each ListBox to that property. Default for right ListBox is Selected:, and left is Available:."
  '##SUMMARY Reads LeftLabel and RightLabel properties of left _
    and right ListBoxes, repectively, and sets label caption _
    of each ListBox to that property. Default for right _
    ListBox is <EM>Selected:</EM>, and left is <EM>Available:</EM>.
  '##PARAM PropBag (I) Intrinsic VB object containing properties _
    of controls within main control.
  RightLabel = PropBag.ReadProperty("RightLabel", "Selected:")
  LeftLabel = PropBag.ReadProperty("LeftLabel", "Available:")
End Sub

Public Property Let Enabled(ByVal NewValue As Boolean)
Attribute Enabled.VB_Description = "Boolean determining whether control is enabled or not."
  '##SUMMARY Boolean determining whether control is enabled or not.
  '##PARAM NewValue (I) Boolean used to set Enabled property of all _
    components within main object.
  lstRight.Enabled = NewValue
  lstLeft.Enabled = NewValue
  cmdMoveRight.Enabled = NewValue
  cmdMoveLeft.Enabled = NewValue
  cmdMoveAllRight.Enabled = NewValue
  cmdMoveAllLeft.Enabled = NewValue
  cmdMove(0).Enabled = NewValue
  cmdMove(1).Enabled = NewValue
End Property
Public Property Get Enabled() As Boolean
  Enabled = lstRight.Enabled
End Property


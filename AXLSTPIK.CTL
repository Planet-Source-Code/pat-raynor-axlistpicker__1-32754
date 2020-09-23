VERSION 5.00
Begin VB.UserControl AXListPicker 
   ClientHeight    =   1785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5790
   LockControls    =   -1  'True
   PropertyPages   =   "AXLstPik.ctx":0000
   ScaleHeight     =   1785
   ScaleWidth      =   5790
   ToolboxBitmap   =   "AXLstPik.ctx":002D
   Begin VB.ListBox lstSource 
      Height          =   1455
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   255
      Width           =   2220
   End
   Begin VB.ListBox lstDestination 
      Height          =   1455
      Index           =   0
      Left            =   2970
      TabIndex        =   8
      Top             =   255
      Width           =   2220
   End
   Begin VB.CommandButton cmdRightOne 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   2340
      Picture         =   "AXLstPik.ctx":01AF
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      UseMaskColor    =   -1  'True
      Width           =   576
   End
   Begin VB.CommandButton cmdRightAll 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   2340
      Picture         =   "AXLstPik.ctx":02B1
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   615
      UseMaskColor    =   -1  'True
      Width           =   576
   End
   Begin VB.CommandButton cmdLeftOne 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   2340
      Picture         =   "AXLstPik.ctx":03B3
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   990
      UseMaskColor    =   -1  'True
      Width           =   576
   End
   Begin VB.CommandButton cmdLeftAll 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   2340
      Picture         =   "AXLstPik.ctx":04B5
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1365
      UseMaskColor    =   -1  'True
      Width           =   576
   End
   Begin VB.CommandButton cmdUp 
      Enabled         =   0   'False
      Height          =   435
      Left            =   5265
      Picture         =   "AXLstPik.ctx":05B7
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   465
      UseMaskColor    =   -1  'True
      Width           =   435
   End
   Begin VB.CommandButton cmdDown 
      Enabled         =   0   'False
      Height          =   435
      Left            =   5280
      Picture         =   "AXLstPik.ctx":06B9
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1050
      UseMaskColor    =   -1  'True
      Width           =   435
   End
   Begin VB.ListBox lstDestination 
      Height          =   1455
      Index           =   1
      Left            =   2970
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   255
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.ListBox lstSource 
      Height          =   1455
      Index           =   1
      Left            =   60
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   255
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.Label lblSource 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Source:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Tag             =   "2406"
      Top             =   15
      Width           =   555
   End
   Begin VB.Label lblDestination 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Destination:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2970
      TabIndex        =   7
      Tag             =   "2407"
      Top             =   15
      Width           =   840
   End
End
Attribute VB_Name = "AXListPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "ActiveX List Picker Control"
Dim msglWidth As Single
Dim mnSourceList As Integer
Dim mnDestinationList As Integer
Dim mnSourceSortType As Integer
Dim mnDestinationSortType As Integer
Dim mbReadingProperties As Boolean

Const CTL_WIDTH1 = 5265
Const CTL_WIDTH2 = 5820

'================================================================
' Enums
'================================================================
Public Enum lspSourceSortType
  lspSTSourceUnsorted
  lspSTSourceSorted
End Enum

Public Enum lspDestinationSortType
  lspSTDestinationUnsorted
  lspSTDestinationSorted
  lspSTOrderButtons
End Enum

Public Enum lspListCountConstants
  lspSource
  lspDesintation
  lspAll
End Enum

Public Enum lspControlsType
  lspCTSource
  lspCTDestination
  lspCTAddAll
  lspCTAddOne
  lspCTRemoveAll
  lspCTRemoveOne
  lspCTOrderUp
  lspCTOrderDown
  lspCTNone
End Enum

Public Enum lspListIndex
  lspSourceList
  lspDestinationList
End Enum

'================================================================
' Event Declarations
'================================================================
Event MoveToDestination(Cancel As Boolean, Count As Integer)
Event MoveToSource(Cancel As Boolean, Count As Integer)

'================================================================
' Design Time/Runtime Public Properties
'================================================================

Public Property Let SourceCaption(ByVal sVal As String)
  lblSource.Caption = sVal
  PropertyChanged "SourceCaption"
End Property

Public Property Get SourceCaption() As String
  SourceCaption = lblSource.Caption
End Property

Public Property Let DestinationCaption(ByVal sVal As String)
  lblDestination.Caption = sVal
  PropertyChanged "DestinationCaption"
End Property

Public Property Get DestinationCaption() As String
  DestinationCaption = lblDestination.Caption
End Property

Public Property Get SourceSortType() As lspSourceSortType
  SourceSortType = mnSourceSortType
End Property

Public Property Let SourceSortType(ByVal nVal As lspSourceSortType)
  'disallow at run time
  If Ambient.UserMode And (Not mbReadingProperties) Then
    MsgBox Error(382)
    Exit Property
  End If
  
  If nVal = lspSTSourceUnsorted Then
    lstSource(1).Visible = False
    lstSource(0).Visible = True
  ElseIf nVal = lspSTSourceSorted Then
    lstSource(0).Visible = False
    lstSource(1).Visible = True
  Else
    MsgBox Error(380)
    Exit Property
  End If
  mnSourceSortType = nVal
  mnSourceList = nVal
  PropertyChanged "SourceSortType"
End Property

Public Property Get DestinationSortType() As lspDestinationSortType
  DestinationSortType = mnDestinationSortType
End Property

Public Property Let DestinationSortType(ByVal nVal As lspDestinationSortType)
  On Error GoTo DestinationSortTypeErr
  'disallow at run time
  If Ambient.UserMode And (Not mbReadingProperties) Then
    MsgBox Error(382)
    Exit Property
  End If
  
  If nVal = lspSTDestinationSorted Then   'sorted
    lstDestination(0).Visible = False
    lstDestination(1).Visible = True
    mnDestinationList = 1
    msglWidth = CTL_WIDTH1
  ElseIf nVal = lspSTOrderButtons Or nVal = lspSTDestinationUnsorted Then
    lstDestination(1).Visible = False
    lstDestination(0).Visible = True
    mnDestinationList = 0
    If nVal = lspSTDestinationUnsorted Then
      msglWidth = CTL_WIDTH1
      cmdUp.Enabled = False
      cmdDown.Enabled = False
    Else
      'show the order buttons
      msglWidth = CTL_WIDTH2
      cmdUp.Enabled = True
      cmdDown.Enabled = True
    End If
  Else
    MsgBox Error(380)
  End If
  
  mnDestinationSortType = nVal
  UserControl_Resize
  PropertyChanged "DestinationSortType"
  
  Exit Property
DestinationSortTypeErr:
  MsgBox Err.Description
End Property

Public Property Get BackColor() As OLE_COLOR
  BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
  UserControl.BackColor = vNewValue
  PropertyChanged "BackColor"
End Property

Public Property Get ListColor() As OLE_COLOR
  ListColor = lstSource(0).BackColor
End Property

Public Property Let ListColor(ByVal vNewValue As OLE_COLOR)
  lstSource(0).BackColor = vNewValue
  lstSource(1).BackColor = vNewValue
  lstDestination(0).BackColor = vNewValue
  lstDestination(1).BackColor = vNewValue
  PropertyChanged "ListColor"
End Property

Public Property Get ListBoxFont() As Font
  Set ListBoxFont = lstSource(0).Font
End Property

Public Property Set ListBoxFont(ByVal vNewValue As Font)
  Set lstSource(0).Font = vNewValue
  Set lstSource(1).Font = vNewValue
  Set lstDestination(0).Font = vNewValue
  Set lstDestination(1).Font = vNewValue
  PropertyChanged "ListBoxFont"
End Property

Public Property Get LabelFont() As Font
  Set LabelFont = lblSource.Font
End Property

Public Property Set LabelFont(ByVal vNewValue As Font)
  Set lblSource.Font = vNewValue
  Set lblDestination.Font = vNewValue
  PropertyChanged "LabelFont"
End Property

Public Property Get Enabled() As Boolean
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
  Dim ctl As Control
  For Each ctl In UserControl.Controls
    ctl.Enabled = vNewValue
  Next ctl
  UserControl.Enabled = vNewValue
  PropertyChanged "Enabled"
End Property

'================================================================
' Property Bag Functions
'================================================================
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  On Error Resume Next
  'Load property values from storage
  mbReadingProperties = True
  SourceCaption = PropBag.ReadProperty("SourceCaption")
  DestinationCaption = PropBag.ReadProperty("DestinationCaption")
  SourceSortType = PropBag.ReadProperty("SourceSortType", mnSourceSortType)
  DestinationSortType = PropBag.ReadProperty("DestinationSortType", mnDestinationSortType)
  BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
  ListColor = PropBag.ReadProperty("ListColor", &H80000005)
  Set ListBoxFont = PropBag.ReadProperty("ListBoxFont")
  Set LabelFont = PropBag.ReadProperty("LabelFont")
  Enabled = PropBag.ReadProperty("Enabled", True)
'   = PropBag.ReadProperty("", )
  mbReadingProperties = False
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  On Error Resume Next
  'Write property values to storage
  Call PropBag.WriteProperty("SourceCaption", SourceCaption, "")
  Call PropBag.WriteProperty("DestinationCaption", DestinationCaption, "")
  Call PropBag.WriteProperty("SourceSortType", SourceSortType, 0)
  Call PropBag.WriteProperty("DestinationSortType", DestinationSortType, 0)
  Call PropBag.WriteProperty("BackColor", BackColor, &H8000000F)
  Call PropBag.WriteProperty("ListColor", ListColor, &H80000005)
  Call PropBag.WriteProperty("ListBoxFont", ListBoxFont)
  Call PropBag.WriteProperty("LabelFont", LabelFont)
  Call PropBag.WriteProperty("Enabled", Enabled, True)
'  Call PropBag.WriteProperty("", , )
End Sub

'================================================================
' Runtime Only Public Properties
'================================================================
Public Property Let SourceListIndex(nIndex As Integer)
Attribute SourceListIndex.VB_MemberFlags = "400"
  On Error Resume Next
  lstSource(mnSourceList).ListIndex = nIndex
  If Err Then
    MsgBox Error(9)
  End If
End Property

Public Property Get SourceListIndex() As Integer
  SourceListIndex = lstSource(mnSourceList).ListIndex
End Property

Public Property Let DestinationListIndex(nIndex As Integer)
Attribute DestinationListIndex.VB_MemberFlags = "400"
  On Error Resume Next
  lstDestination(mnDestinationList).ListIndex = nIndex
  If Err Then
    MsgBox Error(9)
  End If
End Property

Public Property Get DestinationListTopIndex() As Integer
Attribute DestinationListTopIndex.VB_MemberFlags = "400"
  DestinationListTopIndex = lstDestination(mnDestinationList).TopIndex
End Property

Public Property Let SourceListTopIndex(nTopIndex As Integer)
Attribute SourceListTopIndex.VB_MemberFlags = "400"
  On Error Resume Next
  lstSource(mnSourceList).TopIndex = nTopIndex
  If Err Then
    MsgBox Error(9)
  End If
End Property

Public Property Get SourceListTopIndex() As Integer
  SourceListTopIndex = lstSource(mnSourceList).TopIndex
End Property

Public Property Let DestinationListTopIndex(nTopIndex As Integer)
  On Error Resume Next
  lstDestination(mnDestinationList).ListIndex = nTopIndex
  If Err Then
    MsgBox Error(9)
  End If
End Property

Public Property Get DestinationListIndex() As Integer
  DestinationListIndex = lstDestination(mnDestinationList).ListIndex
End Property

Public Property Get SourceList(nIndex As Integer) As String
Attribute SourceList.VB_MemberFlags = "400"
  On Error GoTo SourceListErr
  SourceList = lstSource(mnSourceList).List(nIndex)
  Exit Property
SourceListErr:
  MsgBox Error
End Property

Public Property Get DestinationList(nIndex As Integer) As String
Attribute DestinationList.VB_MemberFlags = "400"
  On Error GoTo DestinationListErr
  DestinationList = lstDestination(mnDestinationList).List(nIndex)
  Exit Property
DestinationListErr:
  MsgBox Error
End Property

Public Property Get ListCount(nList As lspListCountConstants) As Long
Attribute ListCount.VB_MemberFlags = "400"
  If nList = lspSource Then
    ListCount = lstSource(mnSourceList).ListCount
  ElseIf nList = lspDesintation Then
    ListCount = lstDestination(mnDestinationList).ListCount
  ElseIf nList = lspAll Then
    ListCount = lstSource(mnSourceList).ListCount + lstDestination(mnDestinationList).ListCount
  Else
    MsgBox Error(380)
  End If
End Property

'================================================================
' Public methods
'================================================================
Public Sub AddItem(sItem As String, lCtrl As lspListIndex, Optional nIndex As Variant)
  On Error GoTo AddSourceItemErr

  If lCtrl = lspSourceList Then
    If TypeName(nIndex) = "Integer" Then
      lstSource(mnSourceList).AddItem sItem, nIndex
    Else
      lstSource(mnSourceList).AddItem sItem
    End If
  '  If lstSource(mnSourceList).ListIndex < 0 And lstSource(mnSourceList).ListCount > 0 Then
      'select the first item added
      lstSource(mnSourceList).ListIndex = 0
  '  End If
  ElseIf lCtrl = lspDestinationList Then
    If TypeName(nIndex) = "Integer" Then
      lstDestination(mnDestinationList).AddItem sItem, nIndex
    Else
      lstDestination(mnDestinationList).AddItem sItem
    End If
  Else
    MsgBox Error(5)
  End If

  Exit Sub
  
AddSourceItemErr:
  MsgBox Error
End Sub

Public Sub RemoveItem(lCtrl As lspListIndex, nIndex As Integer)
  On Error GoTo RemoveSourceItemErr

  If lCtrl = lspSourceList Then
    lstSource(mnSourceList).RemoveItem nIndex
  ElseIf lCtrl = lspDestinationList Then
    lstDestination(mnDestinationList).RemoveItem nIndex
  Else
    MsgBox Error(5)
  End If

  Exit Sub
  
RemoveSourceItemErr:
  MsgBox Error
End Sub

Public Sub ClearSource()
  lstSource(mnSourceList).Clear
End Sub

Public Sub ClearDestination()
  lstDestination(mnDestinationList).Clear
End Sub

Public Sub ClearAll()
  lstSource(mnSourceList).Clear
  lstDestination(mnDestinationList).Clear
End Sub

'================================================================
' Private events and support routines
'================================================================
Private Sub cmdUp_Click()
  On Error Resume Next
  Dim nItem As Integer
  
  With lstDestination(mnDestinationList)
    If .ListIndex < 0 Then Exit Sub
    nItem = .ListIndex
    If nItem = 0 Then Exit Sub  'can't move 1st item up
    'move item up
    .AddItem .Text, nItem - 1
    'remove old item
    .RemoveItem nItem + 1
    'select the item that was just moved
    .Selected(nItem - 1) = True
  End With
End Sub

Private Sub cmdDown_Click()
  On Error Resume Next
  Dim nItem As Integer
  
  With lstDestination(mnDestinationList)
    If .ListIndex < 0 Then Exit Sub
    nItem = .ListIndex
    If nItem = .ListCount - 1 Then Exit Sub 'can't move last item down
    'move item down
    .AddItem .Text, nItem + 2
    'remove old item
    .RemoveItem nItem
    'select the item that was just moved
    .Selected(nItem + 1) = True
  End With
End Sub

Private Sub cmdRightOne_Click()
  On Error GoTo cmdRightOne_ClickErr
  Dim i As Integer
  Dim bCancel As Boolean
  
  If lstSource(mnSourceList).ListIndex < 0 Then Exit Sub
  i = lstSource(mnSourceList).ListIndex
  
  'raise the event before doing the move
  RaiseEvent MoveToDestination(bCancel, 1)
  
  If bCancel Then Exit Sub
  
  'go ahead and move it
  lstDestination(mnDestinationList).AddItem lstSource(mnSourceList).Text
  lstSource(mnSourceList).RemoveItem lstSource(mnSourceList).ListIndex
  If lstSource(mnSourceList).ListCount > 0 Then
    If i > lstSource(mnSourceList).ListCount - 1 Then
      lstSource(mnSourceList).ListIndex = i - 1
    Else
      lstSource(mnSourceList).ListIndex = i
    End If
  End If
  lstDestination(mnDestinationList).ListIndex = lstDestination(mnDestinationList).NewIndex
  
  Exit Sub
cmdRightOne_ClickErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRightAll_Click()
  On Error GoTo cmdRightAll_ClickErr
  Dim i As Integer
  Dim bCancel As Boolean
  Dim nCount As Integer
  
  nCount = lstSource(mnSourceList).ListCount
  If nCount = 0 Then Exit Sub
  
  'raise the event before doing the move
  RaiseEvent MoveToDestination(bCancel, nCount)
  
  If bCancel Then Exit Sub
  
  'go ahead and move then
  For i = 0 To lstSource(mnSourceList).ListCount - 1
    lstDestination(mnDestinationList).AddItem lstSource(mnSourceList).List(i)
  Next
  lstSource(mnSourceList).Clear
  lstDestination(mnDestinationList).ListIndex = 0
  
  Exit Sub
cmdRightAll_ClickErr:
  MsgBox Err.Description
End Sub

Private Sub cmdLeftOne_Click()
  On Error GoTo cmdLeftOne_ClickErr
  Dim i As Integer
  Dim bCancel As Boolean
  
  If lstDestination(mnDestinationList).ListIndex < 0 Then Exit Sub
  i = lstDestination(mnDestinationList).ListIndex
  
  'raise the event before doing the move
  RaiseEvent MoveToSource(bCancel, 1)
  
  If bCancel Then Exit Sub
  
  'go ahead and move it
  lstSource(mnSourceList).AddItem lstDestination(mnDestinationList).Text
  lstDestination(mnDestinationList).RemoveItem lstDestination(mnDestinationList).ListIndex
  
  lstSource(mnSourceList).ListIndex = lstSource(mnSourceList).NewIndex
  If lstDestination(mnDestinationList).ListCount > 0 Then
    If i > lstDestination(mnDestinationList).ListCount - 1 Then
      lstDestination(mnDestinationList).ListIndex = i - 1
    Else
      lstDestination(mnDestinationList).ListIndex = i
    End If
  End If
  
  Exit Sub
cmdLeftOne_ClickErr:
  MsgBox Err.Description
End Sub

Private Sub cmdLeftAll_Click()
  On Error GoTo cmdLeftAll_ClickErr
  Dim i As Integer
  Dim bCancel As Boolean
  Dim nCount As Integer
  
  nCount = lstDestination(mnDestinationList).ListCount
  If nCount = 0 Then Exit Sub
  
  'raise the event before doing the move
  RaiseEvent MoveToSource(bCancel, nCount)
  
  If bCancel Then Exit Sub
  
  'go ahead and move then
  For i = 0 To lstDestination(mnDestinationList).ListCount - 1
    lstSource(mnSourceList).AddItem lstDestination(mnDestinationList).List(i)
  Next
  lstDestination(mnDestinationList).Clear
  lstSource(mnSourceList).ListIndex = lstSource(mnSourceList).NewIndex

  Exit Sub
cmdLeftAll_ClickErr:
  MsgBox Err.Description
End Sub

Private Sub lstSource_DblClick(Index As Integer)
  cmdRightOne_Click
End Sub

Private Sub lstDestination_DblClick(Index As Integer)
  cmdLeftOne_Click
End Sub

Private Sub UserControl_Initialize()
  mnDestinationSortType = DestinationSortType
  msglWidth = CTL_WIDTH1
End Sub

Private Sub UserControl_Resize()
  'force the size
  UserControl.Size msglWidth, 1785
End Sub


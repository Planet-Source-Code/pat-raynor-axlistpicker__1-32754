VERSION 5.00
Object = "*\AAXLstpik.vbp"
Begin VB.Form Test 
   Caption         =   "Form1"
   ClientHeight    =   3870
   ClientLeft      =   4320
   ClientTop       =   3285
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3870
   ScaleWidth      =   6090
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Destination Item"
      Height          =   465
      Index           =   1
      Left            =   3135
      TabIndex        =   2
      Top             =   2040
      Width           =   2160
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Source Item"
      Height          =   465
      Index           =   0
      Left            =   210
      TabIndex        =   1
      Top             =   2040
      Width           =   2160
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear All"
      Height          =   480
      Left            =   210
      TabIndex        =   5
      Top             =   3240
      Width           =   2160
   End
   Begin VB.CommandButton cmdShowAll 
      Caption         =   "Show All Items Picked"
      Height          =   480
      Left            =   3135
      TabIndex        =   6
      Top             =   3210
      Width           =   2160
   End
   Begin VB.CommandButton cmdCurrent 
      Caption         =   "Show Current Dest. Item"
      Height          =   465
      Index           =   1
      Left            =   3135
      TabIndex        =   4
      Top             =   2640
      Width           =   2160
   End
   Begin VB.CommandButton cmdCurrent 
      Caption         =   "Show Current Source Item"
      Height          =   465
      Index           =   0
      Left            =   210
      TabIndex        =   3
      Top             =   2640
      Width           =   2160
   End
   Begin AXLstPik.AXListPicker AXListPicker1 
      Height          =   1785
      Left            =   120
      TabIndex        =   0
      Top             =   195
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   3149
      SourceCaption   =   "&Source:"
      DestinationCaption=   "&Destination:"
      SourceSortType  =   1
      DestinationSortType=   2
   End
End
Attribute VB_Name = "Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AXListPicker1_MoveToDestination(Cancel As Boolean, Count As Integer)
  MsgBox "Moving " & Count & " items to the Destination!"
End Sub

Private Sub AXListPicker1_MoveToSource(Cancel As Boolean, Count As Integer)
  MsgBox "Moving " & Count & " items to the Source!"
End Sub

Private Sub cmdAdd_Click(Index As Integer)
  Dim sTmp As String
  sTmp = InputBox("Add Item:")
  If Len(sTmp) = 0 Then Exit Sub
  If Index = 0 Then
    AXListPicker1.AddItem sTmp, lspSource
  Else
    AXListPicker1.AddItem sTmp, lspDesintation
  End If
End Sub

Private Sub cmdClearAll_Click()
  AXListPicker1.ClearAll
End Sub

Private Sub cmdCurrent_Click(Index As Integer)
  If Index = 0 Then
    If AXListPicker1.SourceListIndex < 0 Then Exit Sub
    MsgBox AXListPicker1.SourceList(AXListPicker1.SourceListIndex)
  Else
    If AXListPicker1.DestinationListIndex < 0 Then Exit Sub
    MsgBox AXListPicker1.DestinationList(AXListPicker1.DestinationListIndex)
  End If
End Sub

Private Sub cmdShowAll_Click()
  Dim sTmp As String
  Dim i As Integer
  For i = 0 To AXListPicker1.ListCount(lspDesintation) - 1
    sTmp = sTmp & ":" & AXListPicker1.DestinationList(i)
  Next
  MsgBox sTmp
End Sub

Private Sub Form_Load()
  AXListPicker1.AddItem "AAA", lspSourceList
  AXListPicker1.AddItem "BBB", lspSourceList
  AXListPicker1.AddItem "DDD", lspSourceList
  AXListPicker1.AddItem "CCC", lspSourceList
  AXListPicker1.AddItem "ABC", lspSourceList
End Sub

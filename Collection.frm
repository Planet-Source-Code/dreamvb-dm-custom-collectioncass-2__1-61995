VERSION 5.00
Begin VB.Form frmmain 
   Caption         =   "DM Collection Class V2"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSort2 
      Caption         =   "Sort Desending"
      Height          =   375
      Left            =   5130
      TabIndex        =   21
      Top             =   3165
      Width           =   1485
   End
   Begin VB.CommandButton cmdSort1 
      Caption         =   "Sort Asending"
      Height          =   375
      Left            =   3510
      TabIndex        =   20
      Top             =   3165
      Width           =   1485
   End
   Begin VB.CommandButton cmdRemoveEX 
      Caption         =   "RemoveEX Items 1,3,5"
      Height          =   375
      Left            =   1800
      TabIndex        =   19
      Top             =   3765
      Width           =   2700
   End
   Begin VB.CommandButton Command2 
      Caption         =   "RemoveAll"
      Height          =   375
      Left            =   90
      TabIndex        =   18
      Top             =   3765
      Width           =   1515
   End
   Begin VB.CommandButton cmdmove 
      Caption         =   "Move Item 1 to Item 5"
      Height          =   375
      Left            =   4605
      TabIndex        =   17
      Top             =   3735
      Width           =   2700
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   150
      TabIndex        =   16
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton cmdabout 
      Caption         =   "About"
      Height          =   375
      Left            =   1110
      TabIndex        =   15
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton cmdaFind 
      Caption         =   "Find Item"
      Height          =   375
      Left            =   1785
      TabIndex        =   14
      Top             =   3180
      Width           =   1515
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Collection form File"
      Height          =   435
      Left            =   4425
      TabIndex        =   13
      Top             =   2010
      Width           =   2880
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save Collection to File"
      Height          =   435
      Left            =   4425
      TabIndex        =   12
      Top             =   1455
      Width           =   2880
   End
   Begin VB.TextBox txtkey 
      Height          =   315
      Left            =   2550
      TabIndex        =   9
      Text            =   "Key"
      Top             =   2610
      Width           =   1140
   End
   Begin VB.TextBox txtItem 
      Height          =   315
      Left            =   615
      TabIndex        =   8
      Text            =   "ItemName"
      Top             =   2610
      Width           =   1380
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   315
      Left            =   3825
      TabIndex        =   6
      Top             =   2610
      Width           =   600
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove a Item"
      Height          =   375
      Left            =   90
      TabIndex        =   1
      Top             =   3180
      Width           =   1515
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   75
      TabIndex        =   0
      Top             =   240
      Width           =   4125
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Add new item to collection"
      Height          =   195
      Left            =   135
      TabIndex        =   11
      Top             =   2295
      Width           =   1860
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Key:"
      Height          =   195
      Left            =   2130
      TabIndex        =   10
      Top             =   2670
      Width           =   315
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   2670
      Width           =   345
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Collection Count:"
      Height          =   195
      Left            =   4410
      TabIndex        =   5
      Top             =   1050
      Width           =   1200
   End
   Begin VB.Label lblIndex 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Key:"
      Height          =   195
      Left            =   4410
      TabIndex        =   4
      Top             =   795
      Width           =   315
   End
   Begin VB.Label lblitem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item:"
      Height          =   195
      Left            =   4410
      TabIndex        =   3
      Top             =   270
      Width           =   345
   End
   Begin VB.Label lblkey 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Key:"
      Height          =   195
      Left            =   4410
      TabIndex        =   2
      Top             =   555
      Width           =   315
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DMCollection As New CollectionFX
Dim Lst_Index As Long

Sub ReFillList()
Dim x As Long
    'Clears the listbox and refills with the collection data
    List1.Clear
    For x = DMCollection.Lower To DMCollection.Count
        List1.AddItem DMCollection.Item(x)
    Next
End Sub

Private Sub cmdabout_Click()
    MsgBox "Collection Class by DreamVb V2" _
    & vbCrLf & "Please vote if you find this usfull.", vbInformation
    
End Sub

Private Sub cmdAdd_Click()
    DMCollection.Add txtItem.Text, txtkey.Text
    'Update and show new collection
    Call ReFillList
End Sub

Private Sub cmdaFind_Click()
Dim sFind As String
    'Find an item in the collection
    sFind = InputBox("Enter a item name to find", "Find Item")
    
    If Not DMCollection.ItemExists(sFind) Then
        MsgBox "Item '" & sFind & "' Not found"
        'Not found
    Else
        'Found item and return the index number
        MsgBox "Item '" & sFind & "' found at Index " & DMCollection.ItemIndex(sFind)
    End If
    
End Sub

Private Sub cmdexit_Click()
    List1.Clear
    DMCollection.RemoveAll
    Set DMCollection = Nothing
    Unload frmmain
End Sub

Private Sub cmdLoad_Click()
    DMCollection.LoadFromFile FixPath(App.Path) & "Test.txt" 'Load a collection form a file
    Call ReFillList ' Do update
End Sub

Private Sub cmdmove_Click()
    DMCollection.MoveTo 1, 5
    Call ReFillList
End Sub

Private Sub cmdRemove_Click()
    DMCollection.RemoveItem Lst_Index 'Remove an item from the collection
    Call ReFillList ' Do update
End Sub

Function FixPath(lpPath As String) As String
    'Fix path
    If Right(lpPath, 1) = "\" Then
        FixPath = lpPath
        Exit Function
    Else
        FixPath = lpPath & "\"
    End If
End Function

Private Sub cmdRemoveEX_Click()
    DMCollection.RemoveItemEx 1, 3, 5
    Call ReFillList
End Sub

Private Sub cmdsave_Click()
    DMCollection.SaveToFile FixPath(App.Path) & "Test.txt" 'Save a collecion to a given file
End Sub

Private Sub cmdSort1_Click()
    DMCollection.Sort Ascending
    Call ReFillList
End Sub

Private Sub cmdSort2_Click()
    DMCollection.Sort Descending
    Call ReFillList
End Sub

Private Sub Form_Load()
Dim x As Long
    List1.Clear
    'Make a test collection
    For x = DMCollection.Lower To 26
        DMCollection.Add Chr(64 + x), "Key_" & Chr(64 + x)
    Next
    
    DMCollection.Compare = vbBinaryCompare
    Call ReFillList ' Do update
    List1.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmmain = Nothing
End Sub

Private Sub List1_Click()
    Lst_Index = List1.ListIndex + 1
    'Show some information
    lblkey.Caption = "Key: " & DMCollection.key(Lst_Index)
    lblitem.Caption = "Item: " & DMCollection.Item(Lst_Index)
    lblIndex.Caption = "Item Index: " & Lst_Index
    lblCount.Caption = "Collection Count: " & DMCollection.Count
End Sub

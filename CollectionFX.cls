VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CollectionFX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'DM Custom Collection 1

' The project was started
' Support to:
' Add, Remove and Find Items in a collection
' Feature To add and Save Collections to a given file
' Example project make to show the class in action.

'DM Custom Collection 2
'Added KeyIndex function this now fixed the bug form version 1
'Added key checking support so now you can;t add two items with the same Key
'Added compare options for finding items and keys supports vbtextcompare and vb binarycompare
'Added sort feature to sort items in a collection Ascending or Descending
'Added new RemoveEx Feature to allow you to delete items in a parmarray this mean you chould delete about 10 items in one snap
'Added moveto feature so you can now move items from different location in the collection
'updated example showing the new features in action


'DM Custom Collection
'was a class i made to replace the statdare one in VB and to add extra features
'as of now all the data is stored as Variant. But i may update it to support objects sometime
'if you want to use this class in your own projects by all means do so, all i ask in return is you add my name somewere

'Thanks
' Ben Jones

Private Type mCollection
    Item As Variant
    key As Variant
End Type

Enum ExistsType
    ItemName = 1
    KeyName = 2
End Enum

Enum SortOrder
    Ascending = 1
    Descending = 2
End Enum

Private m_Collection_Count As Long 'Hold collection count
Private m_CollectionType() As mCollection 'Collection
Private m_CompareMethod As VbCompareMethod 'Hold compare method

Function Sort(sOrder As SortOrder)
Dim x As Long, y As Long, nSize As Long
    'Simple Quick sort this may not be perfect. and may need some little fixes to inprove on it
    For x = 1 To m_Collection_Count
        For y = x + 1 To m_Collection_Count
            If sOrder = Ascending Then 'Sort Ascending
                If Mid(m_CollectionType(x).Item, 1, 1) > Mid(m_CollectionType(y).Item, 1, 1) Then
                    Swap m_CollectionType(x).Item, m_CollectionType(y).Item
                    Swap m_CollectionType(x).key, m_CollectionType(y).key
                End If
            Else 'Sort Descending
                If Mid(m_CollectionType(x).Item, 1, 1) < Mid(m_CollectionType(y).Item, 1, 1) Then
                    Swap m_CollectionType(x).Item, m_CollectionType(y).Item
                    Swap m_CollectionType(x).key, m_CollectionType(y).key
                End If
            End If
        Next y
    Next x
    x = 0
    y = 0
    nSize = 0
    
End Function

Private Function IsEqual(lpStr1 As Variant, lpStr2 As Variant, lCompare As VbCompareMethod) As Boolean
    'Used for Compareing items in our collection
    IsEqual = False
    If Not Len(lpStr1) = Len(lpStr2) Then
        Exit Function
    Else
        Select Case lCompare
            Case vbTextCompare:
                IsEqual = (StrComp(lpStr1, lpStr2, lCompare) = 0) 'Compare both values
            Case vbBinaryCompare
                If Len(lpStr1) = 0 Then 'If first value is null return true
                    IsEqual = True
                    Exit Function
                Else
                    IsEqual = (InStr(1, lpStr1, lpStr2, lCompare) <> 0) 'Compare both values
                End If
        End Select
     End If
     
End Function

Sub Add(Item As Variant, Optional key)
    'Add a new item and key to the collection
    m_Collection_Count = m_Collection_Count + Lower 'Keep a count of the items in the collection
    ReDim Preserve m_CollectionType(Lower To m_Collection_Count) As mCollection 'Resize the collection
    m_CollectionType(m_Collection_Count).Item = Item 'Assign the item value
    If IsMissing(key) Then key = vbNullString 'If key is missing fill with a nullchar
    
    If KeyIndex(key) <> -1 Then
        'This will inform the user if they have added to the collection with a key that already been added
        Err.Raise 475, "Add", "The key is already associated with an element of this collection"
        Exit Sub
    Else
        'Assign the key value
        m_CollectionType(m_Collection_Count).key = key
    End If
    
End Sub

Public Function Count() As Long
    'Returns the number of items in the collection
    Count = m_Collection_Count
End Function

Public Property Get Item(Index As Long) As Variant
'On Error Resume Next
    'Return a items value in the collection
    If Not RangeOk(Index) Then
        Err.Raise 9
        Exit Property
    Else
        Item = m_CollectionType(Index).Item
    End If
  End Property

Public Property Let Item(Index As Long, vNewValue As Variant)
    'Chnage a items value in the collection
    If Not RangeOk(Index) Then
        Err.Raise 9
        Exit Property
    Else
        m_CollectionType(Index).Item = vNewValue
    End If
End Property

Public Property Get key(Index As Long) As Variant
    'Return the value of a key in the collection
    If Not RangeOk(Index) Then
        Err.Raise 9
        Exit Property
    Else
        key = m_CollectionType(Index).key
    End If
  End Property

Public Property Let key(Index As Long, vNewValue As Variant)
    'Chnage a keys value in the collection
    If Not RangeOk(Index) Then
        Err.Raise 9
        Exit Property
    ElseIf KeyIndex(vNewValue) <> -1 Then
        'This will inform the user if they have added to the collection with a key that already been added
        Err.Raise 475, "Key", "The key is already associated with an element of this collection"
    Else
        m_CollectionType(Index).key = vNewValue 'Update the collection key with the new value
    End If
End Property

Sub RemoveAll()
    'Resets our collection
    m_Collection_Count = 0
    Erase m_CollectionType
End Sub

Private Sub Class_Terminate()
   Call RemoveAll 'Clear up
End Sub

Private Function RangeOk(Index As Long) As Boolean
    'This just checks the range of an index the user has entered
    RangeOk = CBool((Index = 0) Or (Index > m_Collection_Count)) <> True
End Function

Function ItemExists(Item) As Boolean
    'Return true if an item is found otherwise false is returned
    ItemExists = ItemIndex(Item) <> -1
End Function

Function KeyExists(key) As Boolean
    'Return true if an key is found otherwise false is returned
    KeyExists = KeyIndex(key) <> -1
End Function

Public Property Get ItemIndex(Item) As Long
Dim x As Long
    'Function used to return the index number of an item
    If m_Collection_Count = 0 Then Err.Raise 9: Exit Function 'exit if we have no items in the collection
    ItemIndex = -1 'Default errror index
    
    For x = 1 To m_Collection_Count 'Loop tho the collection
        If IsEqual(m_CollectionType(x).Item, Item, m_CompareMethod) Then  'Check if item is here
            ItemIndex = x 'Yes so return the index
            Exit For 'exit out
        End If
    Next
End Property

Public Property Get KeyIndex(key As Variant) As Long
Dim x As Long
    'Function used to return the index number of a key
    If m_Collection_Count = 0 Then Err.Raise 9: Exit Property 'exit if we have no items in the collection
    KeyIndex = -1 'Default errror index
    
    For x = 1 To m_Collection_Count 'Loop tho the collection
        If IsEqual(m_CollectionType(x).key, key, m_CompareMethod) Then 'Check if item is here
            KeyIndex = x 'Yes so return the index
            Exit For 'exit out
        End If
    Next
End Property

Public Sub SaveToFile(lpFileName As String)
Dim nFile As Long, iSize As Long
    iSize = UBound(m_CollectionType)
    'Used to save the collection contents to a file
    nFile = FreeFile
    Open lpFileName For Binary As #nFile
        Put #nFile, , iSize
        Put #nFile, , m_CollectionType
    Close #nFile
End Sub

Public Sub LoadFromFile(lpFileName As String)
Dim nFile As Long, iSize As Long
    'Used to load a file and fill the collection
    Call RemoveAll
    nFile = FreeFile
    Open lpFileName For Binary As #nFile
        Get #nFile, , iSize
        ReDim m_CollectionType(1 To iSize)
        Get #nFile, , m_CollectionType
    Close #nFile
    
    m_Collection_Count = iSize
End Sub

Public Sub RemoveItem(Index As Long)
Dim TempArr() As mCollection, lngDestIndex As Long, x As Long
On Error Resume Next
    'Remove an item form the collection
    lngDestIndex = Lower
    
    If (m_Collection_Count = Lower) Then
        'Clear the collection if we hit the lower bound
        Call RemoveAll
        Exit Sub
    End If
    
    ReDim TempArr(Lower To m_Collection_Count - Lower) 'Resize the temp array
    
    For x = Lower To m_Collection_Count
        If x <> Index Then
            'X not equal to delete index then make a copy of the collection
            TempArr(lngDestIndex).Item = m_CollectionType(x).Item
            TempArr(lngDestIndex).key = m_CollectionType(x).key
            lngDestIndex = lngDestIndex + 1
        End If
    Next
    
    m_CollectionType = TempArr 'Pass temp array over to our collection
    m_Collection_Count = UBound(m_CollectionType) 'Update our collection count
    Erase TempArr 'Destroy the temp array
    
End Sub

Function Lower()
    Lower = 1 'Our collections lower bound
End Function

Sub MoveTo(Index As Long, ToIndex As Long)
Dim TempA As Variant, TempB As Variant
    'This sub is used to move an item in the list
    ' for example if we wanted to move item 1 to item 5 position
    'item 1 will be moved to item 5s position
    ' were item 5 will be moved to item 1s position
    
    If Not RangeOk(Index) Then
        Err.Raise 9
        Exit Sub
    ElseIf Not RangeOk(ToIndex) Then
        Err.Raise 9
        Exit Sub
    ElseIf Index = ToIndex Then
        Exit Sub
    Else
        TempA = m_CollectionType(ToIndex).Item
        TempB = m_CollectionType(ToIndex).key
        
        m_CollectionType(ToIndex).Item = m_CollectionType(Index).Item
        m_CollectionType(ToIndex).key = m_CollectionType(Index).key
        
        m_CollectionType(Index).Item = TempA
        m_CollectionType(Index).key = TempB
        
        TempA = ""
        TempB = ""
        
    End If

End Sub

Public Sub RemoveItemEx(ParamArray ArgList() As Variant)
Dim x As Integer, nIndex As Long
Dim tmpA() As Variant
    tmpA = ArgList

    'Sort out the array first.
    Call SortArray(tmpA)
    
    For x = 0 To UBound(ArgList) 'Loop tho the sorted list
        nIndex = CLng(tmpA(x)) 'Extract the index
        If Not RangeOk(nIndex) Then
            'Check if we are within our collection range
            Err.Raise 9, "RemoveItemEx" ' opps send error
            Exit For 'get out ta here
        Else
            Call RemoveItem(nIndex) 'Remove the item from the collection
        End If
    Next x
    
    nIndex = 0
    Erase tmpA
    
End Sub


Public Property Get Compare() As VbCompareMethod
    Compare = m_CompareMethod
End Property

Public Property Let Compare(ByVal New_Compare As VbCompareMethod)
    m_CompareMethod = New_Compare
End Property

Private Sub SortArray(lpArr As Variant)
Dim bSize As Long, x As Long, y As Long
    'Quick sort
    ' Used to sort nunmber form low level to high level eg 1,2,3 to 3,2,1
    If IsEmpty(lpArr) Then Err.Raise 13, "SortArray": Exit Sub 'Check if the array is empty
    bSize = UBound(lpArr)
    If bSize = 0 Then Exit Sub 'if only one item is found in the array there no need to carray on
    
    For x = 0 To bSize
        For y = x + 1 To bSize
            If lpArr(x) < lpArr(y) Then
                Swap lpArr(x), lpArr(y)
            End If
        Next
    Next
    x = 0
    y = 0
    bSize = 0
End Sub

Private Sub Swap(a, b)
Dim t
    'Swaps two values
    t = b 'Temp holder for b
    b = a 'Assign B with a's value
    a = t 'Assign a with the temp holder value for b
End Sub


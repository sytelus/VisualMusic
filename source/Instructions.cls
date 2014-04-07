VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Instructions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"MIDIEngine"
Attribute VB_Ext_KEY = "Member0" ,"MIDIEngine"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Private variable to hold collection
Private mCol As Collection
Public Lables As Collection

Public Property Get LOWER_BOUND() As Long

    'Tell the caller that index value for this collection starts from 1
    LOWER_BOUND = 1

End Property

Public Property Get Count() As Long

    'Simply return the count stored in private collection
    Count = mCol.Count

End Property

Public Property Get Item(ByVal Index As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    
    'Error could be generated because
    'specified index doesn't exist.
    On Error GoTo ERR_Get_Item

    'Store the refrence stored in private collection
    Item = mCol.Item(Index)
    
Exit Property
ERR_Get_Item:

    HandleColError

End Property

'This function gives the ability to the collection
'for using For Each statement with it

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    'Return the private collection's
    'hidden enum member
    Set NewEnum = mCol.[_NewEnum]

End Property

Public Function Clear()

    Call ClearCollection(mCol)

End Function

Public Function IsExists(ByVal Name As String) As Boolean

    'Use delayed error handling
    On Error Resume Next
    
    Dim vItem As Variant
    
    'Try to access specified item
    vItem = mCol.Item(Name)
    
    'If there is no error then specified item exist
    If Err.Number = 0 Then
    
        'then return True
        IsExists = True
    
    'If error is "Invalid procedure call or argument", it
    'means specified item does not exist in collection
    ElseIf Err.Number = 5 Then
        
        'else return false
        IsExists = False
    
    'Some other unexpected error has occured
    Else
    
        'So reraise it
        ReRaiseError
    
    End If

End Function

Public Sub Add(ByVal vvaInstructionInfo As Variant, Optional ByVal vsInstructionName As String = "", Optional ByVal vlInstructionNumber As Long = 0)

    'Error could be generated because
    'specified field name may exist already
    On Error GoTo ERR_Add

    If vlInstructionNumber <> 0 Then Err.Raise 1000, , "Not implemented"

    If vsInstructionName <> vbNullString Then

        'Add the field refrence in the collection with the field name as the key
        Call mCol.Add(vvaInstructionInfo, vsInstructionName)
        
    Else
        
        Call mCol.Add(vvaInstructionInfo)
    
    End If
    
Exit Sub
ERR_Add:

    'Some error has occured while adding the
    'item to the collection. The HandleColError
    'function converts this error to more
    'specific error for this custom collection
    HandleColError

End Sub

Public Sub Remove(ByVal Index As Variant)

    'Error could be generated because
    'specified index doesn't exist.
    On Error GoTo ERR_Remove

    'First remove the refrence for the field stored in that element
    mCol.Item(Index) = Empty

    'Remove the specified item from the collection
    Call mCol.Remove(Index)
    
Exit Sub
ERR_Remove:

    'Some error has occured while getting the
    'removing the item from the collection.
    'The HandleColError function converts this
    'error to more specific error for this custom
    'collection.
    HandleColError

End Sub

Private Sub Class_Initialize()
    
    'create the private collection
    Set mCol = New Collection
    Set Lables = New Collection

End Sub

Private Sub Class_Terminate()

    'Remove all items from the collection
    Call ClearCollection(mCol)

    'Destroy the private collection
    Set mCol = Nothing
    Set Lables = Nothing
    
End Sub

Private Sub HandleColError()

    'Check the error number
    Select Case Err.Number

        'Err 5: Invalid Procedure call or argument - occures when wrong key is supplied to collection
        Case 5
        
            Err.Raise 1000, , "Specified instruction name does not exist."
            
        'Err 9: Subscript out of range - occures when invalid index number is specified
        Case 9
            
            Err.Raise 1000, , "There is no instruction for the specified index"
            
        'Err 457: This Key is already associtaed with an element of this collection
        Case 457
        
            Err.Raise 1000, , "A instruction with the same name already present in the collection."
        
        'Unexpected error has occured
        Case Else

            'So just reraise it
            ReRaiseError
            
    End Select

End Sub


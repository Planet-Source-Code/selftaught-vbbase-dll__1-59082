VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSubclasses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
'cSubclasses.cls                        7/5/04
'
'           PURPOSE:
'               Adds and removes subclasses for a single object through a familiar collection interface.
'
'           MODULES CALLED FROM THIS MODULE:
'               mSubclass.cls
'
'==================================================================================================

'1.  iEnumerable Interface    - For ... Each enumeration
'2.  Friendly Interface       - Used internally to initialize the modular variable.
'3.  Collection Interface     - Methods which manage the collection of cSubclass objects.

Implements iEnumerable

Private miPtr As Long       'Stores a pointer to the iSubclass object to be called back for messages

Private Function iEnumerable_Clone() As stdole.IUnknown

End Function

'<Private Interface>
'</Private Interface>

'<iEnumerable Interface>
Private Sub iEnumerable_GetNextItem(tEnum As tEnum, vNextItem As Variant, bNoMoreItems As Boolean)
    mSubclass.Subclasses_NextItem miPtr, tEnum, vNextItem, bNoMoreItems
End Sub

Private Sub iEnumerable_Reset(tEnum As tEnum)
    tEnum.iIndex = -1&
End Sub

Private Sub iEnumerable_Skip(tEnum As tEnum, ByVal iSkipCount As Long, bSkippedAll As Boolean)
    mSubclass.Subclasses_Skip miPtr, tEnum, iSkipCount, bSkippedAll
End Sub

Private Sub iEnumerable_Cleanup(tEnum As tEnum)
    '
End Sub
'</iEnumerable Interface>

'<Friendly Interface>
Friend Sub fInit( _
            iPtr As Long)
    miPtr = iPtr
End Sub
'<Friendly Interface>

'<Public Interface>
'<Collection Interface>
Public Function Add( _
            ByVal hWnd As Long) _
                As cSubclass
    Set Add = mSubclass.Subclasses_Add(miPtr, hWnd)
End Function

Public Sub Remove( _
            ByVal hWnd As Long)
    mSubclass.Subclasses_Remove miPtr, hWnd
End Sub

Public Function Item( _
            ByVal hWnd As Long) _
                As cSubclass
Attribute Item.VB_UserMemId = 0
    Set Item = mSubclass.Subclasses_Item(hWnd, miPtr)
End Function

Public Function Exists( _
            ByVal hWnd As Long) _
                As Boolean
    Exists = mSubclass.Subclasses_Exists(hWnd, miPtr)
End Function

Public Function Count() As Long
    Count = mSubclass.Subclasses_Count(miPtr)
End Function

Public Function Clear() As Long
    Clear = mSubclass.Subclasses_Clear(miPtr)
End Function

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Dim loEnum As cEnumeration
    Set loEnum = New cEnumeration
    Set NewEnum = loEnum.GetEnum(Me, Subclasses_Control(miPtr))
End Property
'</Collection Interface>
'</Public Interface>


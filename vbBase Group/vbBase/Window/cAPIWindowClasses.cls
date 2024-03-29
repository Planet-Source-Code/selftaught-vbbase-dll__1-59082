VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cApiWindowClasses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
'cApiWindowClasses.cls                  7/4/04
'
'           PURPOSE:
'               Exposes functionality to register api window classes through a familiar collection
'               interface.  Classes are created and destroyed by adding and removing from the
'               collection, and accessed and enumerated as items.
'
'           MODULES CALLED FROM THIS MODULE:
'               mWindow
'
'           CLASSES CREATED BY THIS MODULE:
'               cEnumerator
'
'==================================================================================================

'1.  iEnumerable Interface  - allow for .. each enumeration
'2.  Collection Interface   - Delegates to methods in mWindow to maintain registered the class collection

Implements iEnumerable



'<iEnumerable Interface>
Private Sub iEnumerable_GetNextItem(tEnum As tEnum, vNextItem As Variant, bNoMoreItems As Boolean)
    mWindow.ApiWindowClasses_NextItem tEnum, vNextItem, bNoMoreItems
End Sub

Private Sub iEnumerable_Reset(tEnum As tEnum)
    tEnum.iIndex = -1&
End Sub

Private Sub iEnumerable_Skip(tEnum As tEnum, ByVal iSkipCount As Long, bSkippedAll As Boolean)
    tEnum.iIndex = tEnum.iIndex + iSkipCount
    bSkippedAll = CBool(tEnum.iIndex < (mWindow.ApiWindowClasses_Count() - 1&))
End Sub

Private Sub iEnumerable_Cleanup(tEnum As tEnum)
'
End Sub

Private Function iEnumerable_Clone() As stdole.IUnknown
'
End Function

'</iEnumerable Interface>

'<Public Interface>
'<Collection Interface>
Public Function Add( _
                    ByRef sClass As String, _
           Optional ByVal iBackColor As Long = &HFFFFFF, _
           Optional ByVal iStyle As eClassStyle = 0, _
           Optional ByVal hCursor As Long = 0, _
           Optional ByVal hIcon As Long = 0, _
           Optional ByVal hIconSm As Long = 0, _
           Optional ByVal cbClassExtra As Long = 0, _
           Optional ByVal cbWndExtra As Long = 0) _
                As cApiWindowClass

    Set Add = mWindow.ApiWindowClasses_Register(sClass, iBackColor, iStyle, hCursor, hIcon, hIconSm, cbClassExtra, cbWndExtra)
    
End Function

Public Sub Remove( _
            ByVal sClass As String)
    mWindow.ApiWindowClasses_Unregister sClass
End Sub

Public Function Item( _
            ByRef sClass As String) _
                As cApiWindowClass
Attribute Item.VB_UserMemId = 0
    Set Item = mWindow.ApiWindowClasses_Item(sClass)
End Function

Public Function Exists( _
            ByRef sClass As String) _
                As Boolean
    Exists = mWindow.ApiWindowClasses_Exists(sClass)
End Function

Public Function Count() As Long
    Count = mWindow.ApiWindowClasses_Count()
End Function

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Dim loEnum As cEnumeration
    Set loEnum = New cEnumeration
    Set NewEnum = loEnum.GetEnum(Me, mWindow.ApiWindowClasses_GetControl)
End Property
'</Collection Interface>
'</Public Interface>

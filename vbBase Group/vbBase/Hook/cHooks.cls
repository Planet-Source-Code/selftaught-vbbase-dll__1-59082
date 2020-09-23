VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
'cHooks.cls                            7/5/04
'
'           PURPOSE:
'               Adds and removes windows hooks though mHook.bas.
'
'           MODULES CALLED FROM THIS MODULE:
'               mHook.bas
'               mVbBaseGeneral.bas
'
'==================================================================================================


'1.  Friendly Interface     - Initializes the modular variable
'2.  Collection Interface   - Delegates to methods in mWindow
'3.  CurrentHook Interface  - this is a single method that is only valid during a hook callback. it
'                             returns the value identifying the hook type. see eHookType
'4.  lParam Cast Interface  - Contains methods to return appropriate structures from the lParam
'                             memory addresses passed in the lParam argument of a hook callback

Private miPtr As Long

'<Private Interface>
'</Private Interface>

'<Friendly Interface>
Friend Sub fInit( _
            ByVal iPtr As Long)
    miPtr = iPtr
End Sub
'</Friendly Interface>

'<Public Interface>
'<Collection Interface>
Public Sub Add( _
            ByVal iType As eHookType, _
   Optional ByVal bThread As Boolean = True)
    
    mHook.AddHook miPtr, iType, bThread
End Sub

Public Sub Remove( _
            ByVal iType As eHookType, _
   Optional ByVal bThread As Boolean = True)
    mHook.RemoveHook miPtr, iType, bThread
End Sub

Public Function Exists( _
            ByVal iType As eHookType, _
   Optional ByVal bThread As Boolean = True) _
                As Boolean
    Exists = mHook.HookExists(miPtr, iType, bThread)
End Function

Public Function Count( _
   Optional ByVal bThread As Boolean = True) _
                As Long
    Count = mHook.HookCount(miPtr, bThread)
End Function

Public Function Clear( _
   Optional ByVal bThread As Boolean = True) _
                As Long
    Clear = mHook.HookClear(miPtr, bThread)
End Function
'</Collection Interface>

'<CurrentHook Interface>
Public Function CurrentHook() _
                As eHookType
    If iCurrentHook = -100 Then gErr vbbInvalidProcedureCall, "cHooks.CurrentHook"
    CurrentHook = iCurrentHook
End Function
'</CurrentHook Interface>

'<lParam Cast Interface>
Public Function xCBT_CREATEWND(ByVal lParam As Long) As tCBTCreateWndStruct
  Call CopyMemory(xCBT_CREATEWND, ByVal lParam, LenB(xCBT_CREATEWND))
End Function

Public Function xCBT_ACTIVATE(ByVal lParam As Long) As tCBTActivateStruct
  Call CopyMemory(xCBT_ACTIVATE, ByVal lParam, LenB(xCBT_ACTIVATE))
End Function

Public Function xCREATESTRUCT(ByVal lParam As Long) As tCreateStruct
  Call CopyMemory(xCREATESTRUCT, ByVal lParam, LenB(xCREATESTRUCT))
End Function

Public Function xCWPSTRUCT(ByVal lParam As Long) As tCWPStruct
  Call CopyMemory(xCWPSTRUCT, ByVal lParam, LenB(xCWPSTRUCT))
End Function

Public Function xCWPRETSTRUCT(ByVal lParam As Long) As tCWPRetStruct
  Call CopyMemory(xCWPRETSTRUCT, ByVal lParam, LenB(xCWPRETSTRUCT))
End Function

Public Function xDEBUGSTRUCT(ByVal lParam As Long) As tDebugHookInfo
  Call CopyMemory(xDEBUGSTRUCT, ByVal lParam, LenB(xDEBUGSTRUCT))
End Function

Public Function xEVENTMSG(ByVal lParam As Long) As tEventMsg
  Call CopyMemory(xEVENTMSG, ByVal lParam, LenB(xEVENTMSG))
End Function

Public Function xKBDLLHOOKSTRUCT(ByVal lParam As Long) As tKBDLLHookStruct
  Call CopyMemory(xKBDLLHOOKSTRUCT, ByVal lParam, LenB(xKBDLLHOOKSTRUCT))
End Function

Public Function xMOUSEHOOKSTRUCT(ByVal lParam As Long) As tMouseHookStruct
  Call CopyMemory(xMOUSEHOOKSTRUCT, ByVal lParam, LenB(xMOUSEHOOKSTRUCT))
End Function

Public Function xMSG(ByVal lParam As Long) As tMsg
  Call CopyMemory(xMSG, ByVal lParam, LenB(xMSG))
End Function

Public Function xMSLLHOOKSTRUCT(ByVal lParam As Long) As tMSLLHookStruct
  Call CopyMemory(xMSLLHOOKSTRUCT, ByVal lParam, LenB(xMSLLHOOKSTRUCT))
End Function

Public Function xRECT(ByVal lParam As Long) As tRect
  Call CopyMemory(xRECT, ByVal lParam, LenB(xRECT))
End Function
'</lParam Cast Interface>
'</Public Interface>

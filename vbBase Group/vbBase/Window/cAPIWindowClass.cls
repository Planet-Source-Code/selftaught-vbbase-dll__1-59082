VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cApiWindowClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
'cApiWindowClass.cls                  7/4/04
'
'           PURPOSE:
'               Exposes functionality to register api window classes through a familiar collection
'               interface.  Classes are created and destroyed by adding and removing from the
'               collection, and accessed and enumerated as items.
'
'           MODULES CALLED FROM THIS MODULE:
'               mWindow
'
'==================================================================================================

'1.  Friendly Interface     -  Used internally to initialize the modular variable.
'2.  Message Interface      -  Methods which manage the messages that are delived to the owner of the window.
'3.  Window Interface       -  Function that returns a collection of api class window objects

Private msClass As String

'<Private Interface>
'</Private Interface>

'<Friendly Interface>
Friend Sub fInit(ByRef sClass As String)
    msClass = sClass
End Sub
'</Friendly Interface>

'<Public Interface>
Public Function Name() As String
    Name = msClass
End Function

Public Function TotalWindowCount() As Long
    TotalWindowCount = mWindow.ApiWindowClass_WindowCount(msClass)
End Function

'<Message Interface>
Public Function AddDefMsg( _
            ByVal iMsg As eMsg) _
                As Boolean
    AddDefMsg = mWindow.ApiWindowClass_AddDefMsg(iMsg, msClass)
End Function

Public Function DelDefMsg( _
            ByVal iMsg As eMsg) _
        As Boolean
    DelDefMsg = mWindow.ApiWindowClass_DelDefMsg(iMsg, msClass)
End Function

Public Function DefMsgExists( _
            ByVal iMsg As eMsg) _
        As Boolean
    DefMsgExists = mWindow.ApiWindowClass_DefMsgExists(msClass, iMsg)
End Function

Public Function DefMsgCount() As Long
    DefMsgCount = mWindow.ApiWindowClass_DefMsgCount(msClass)
End Function

Public Function GetDefMessages( _
            ByRef iOutArray() As Long) _
                As Long
    GetDefMessages = mWindow.ApiWindowClass_GetDefMessages(iOutArray, msClass)
End Function
'</Message Interface>

'</Window Interface>
Public Function OwnedWindows( _
            ByVal oWho As iWindow) _
                As cApiClassWindows
    Set OwnedWindows = mWindow.ApiWindowClass_OwnedWindows(msClass, oWho)
End Function
'</Window Interface>
'</Public Interface>

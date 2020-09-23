VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cApiClassWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
'cApiClassWindow.cls                                7/4/04
'
'           PURPOSE:
'               Exposes common api calls made to act on a window as well as procedures for dll clients
'               to tell this compont which messages are to be delivered to the owner on the iWindow
'               callback interface.  Custom messages on a per-window basis or default messages defined
'               by the cApiWindowClass object can be received.
'
'           MODULES CALLED FROM THIS MODULE:
'               mWindow
'
'==================================================================================================


'1.  Friendly Interface -  Used internally to initialize the modular variables.
'2.  Message Interface  -  Use these methods to specify which messages you with to have processed.
'3.  Win Api Interface  -  Methods which delegate to Api calls you may find useful.


Private msClass  As String      'stores the class name of this window
Private mtPointer As tItemPointer

'<Private Interface>
'</Private Interface>

'<Friendly Interface>
Friend Sub fInit(ByRef tItemPointer As tItemPointer, _
                 ByRef sClass As String)
    LSet mtPointer = tItemPointer
    msClass = sClass
End Sub
'</Friendly Interface>

'<Public Interface>
Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = 0
    hWnd = mtPointer.iId
End Property
Public Property Get ClassName() As String
    ClassName = msClass
End Property

'<Message Interface>
Public Function AddMsg( _
            ByVal iMsg As eMsg) _
                As Boolean
    AddMsg = mWindow.ApiClassWindow_AddMsg(msClass, mtPointer, iMsg)
End Function

Public Function DelMsg( _
            ByVal iMsg As eMsg) _
                As Boolean
    DelMsg = mWindow.ApiClassWindow_DelMsg(msClass, mtPointer, iMsg)
End Function

Public Function MsgExists( _
            ByVal iMsg As eMsg) _
                As Boolean
    MsgExists = mWindow.ApiClassWindow_MsgExists(msClass, mtPointer, iMsg)
End Function

Public Function MsgCount() As Long
    MsgCount = mWindow.ApiClassWindow_MsgCount(msClass, mtPointer)
End Function

Public Function GetMessages( _
            ByRef iArray() As Long) _
                As Long
    GetMessages = mWindow.ApiClassWindow_GetMessages(msClass, iArray, mtPointer)
End Function

Public Property Get DefaultMessages() _
                    As Boolean
    DefaultMessages = mWindow.ApiClassWindow_DefMessages(msClass, mtPointer)
End Property

Public Property Let DefaultMessages( _
             ByVal bVal As Boolean)
    mWindow.ApiClassWindow_DefMessages(msClass, mtPointer) = bVal
End Property
'</Message Interface>

'<Win Api Interface>
Public Function CallDefWindowProc( _
            ByVal iMsg As eMsg, _
            ByVal wParam As Long, _
            ByVal lParam As Long) _
                As Long
    CallDefWindowProc = DefWindowProc(mtPointer.iId, iMsg, wParam, lParam)
End Function

Public Property Get Enabled() As Boolean
    Enabled = WindowIsEnabled(mtPointer.iId)
End Property

Public Property Let Enabled( _
            ByVal bVal As Boolean)
    WindowEnable mtPointer.iId, bVal
End Property

Public Property Get Text() As String
    Text = WindowText(mtPointer.iId)
End Property

Public Property Let Text( _
            ByVal sVal As String)
    WindowText(mtPointer.iId) = sVal
End Property

Public Function GetRect(ByRef tRect As tRect) As Long
    GetRect = mWindow.WindowGetRect(mtPointer.iId, tRect)
End Function

Public Function Move( _
            ByVal iLeft As Long, _
            ByVal iTop As Long, _
            ByVal iWidth As Long, _
            ByVal iHeight As Long, _
   Optional ByVal bRepaint As Boolean = True) _
                As Boolean
    
    Dim ltRectSize As tRectSize
    With ltRectSize
        
        .Left = iLeft
        .Top = iTop
        .Width = iWidth
        .Height = iHeight
        
        Move = WindowMove(mtPointer.iId, ltRectSize, bRepaint)
        
    End With
    
End Function

Public Function MoveByRect( _
            ByRef tRectSize As tRectSize, _
   Optional ByVal bRepaint As Boolean = True) _
                As Boolean
                
    MoveByRect = mWindow.WindowMove(mtPointer.iId, tRectSize, bRepaint)
                
End Function

Public Function SetPos( _
   Optional ByVal iLeft As Long, _
   Optional ByVal iTop As Long, _
   Optional ByVal iWidth As Long, _
   Optional ByVal iHeight As Long, _
   Optional ByVal hWndInsertAfter As eSWPhWnd, _
   Optional ByVal iFlags As eSWPFlags) _
                As Long
    
    Dim ltRectSize As tRectSize
    
    With ltRectSize
        .Left = iLeft
        .Top = iTop
        .Width = iWidth
        .Height = iHeight
    End With
    
    SetPos = WindowSetPos(mtPointer.iId, hWndInsertAfter, ltRectSize, iFlags)
End Function

Public Function SetPosByRect( _
            ByRef tRectSize As tRectSize, _
   Optional ByVal hWndInsertAfter As eSWPhWnd, _
   Optional ByVal iFlags As eSWPFlags) _
                As Long
    SetPosByRect = WindowSetPos(mtPointer.iId, hWndInsertAfter, tRectSize, iFlags)
End Function

Public Function GetPos( _
   Optional ByRef iLeft As Long, _
   Optional ByRef iTop As Long, _
   Optional ByRef iWidth As Long, _
   Optional ByRef iHeight As Long) _
                As Boolean
    Dim ltRectSize As tRectSize
    With ltRectSize
        .Left = iLeft
        .Top = iTop
        .Width = iWidth
        .Height = iHeight
    End With
    GetPos = WindowGetPos(mtPointer.iId, ltRectSize)
End Function

Public Function GetPosByRect( _
            ByRef tRectSize As tRectSize) _
                As Boolean
        
    GetPosByRect = WindowGetPos(mtPointer.iId, tRectSize)
End Function

Public Function Size( _
            ByVal iWidth As Long, _
            ByVal iHeight As Long, _
   Optional ByVal bRepaint As Boolean) _
                As Boolean
    Dim ltRectSize As tRectSize
    WindowGetPos mtPointer.iId, ltRectSize
    ltRectSize.Width = iWidth
    ltRectSize.Height = iHeight
    WindowMove mtPointer.iId, ltRectSize, bRepaint
End Function

Public Function GetClientDimensions( _
   Optional ByRef iWidth As Long, _
   Optional ByRef iHeight As Long) _
                As Boolean
    GetClientDimensions = mWindow.WindowGetClientDimensions(mtPointer.iId, iWidth, iHeight)
End Function

Public Function GetDc( _
   Optional ByVal bIncludeNonClient As Boolean) _
                As Long
    GetDc = mWindow.WindowGetDC(mtPointer.iId, bIncludeNonClient)
End Function

Public Function ReleaseDc(ByVal hDc As Long) As Long
    ReleaseDc = mWindow.WindowReleaseDC(mtPointer.iId, hDc)
End Function

Public Function Owner() As Long
    Owner = mWindow.WindowGetOwner(mtPointer.iId)
End Function

Public Property Get Parent() As Long
    Parent = mWindow.WindowParent(mtPointer.iId)
End Property

Public Property Let Parent( _
            ByVal iNew As Long)
    mWindow.WindowParent(mtPointer.iId) = iNew
End Property

Public Property Get UserData() As Long
    UserData = mWindow.WindowGetLong(mtPointer.iId, GWL_USERDATA)
End Property

Public Property Let UserData( _
            ByVal iNew As Long)
    mWindow.WindowSetLongDirect mtPointer.iId, GWL_USERDATA, iNew
End Property

Public Property Get ID() As Long
    ID = mWindow.WindowGetLong(mtPointer.iId, GWL_ID)
End Property

Public Property Let ID( _
            ByVal iNew As Long)
    mWindow.WindowSetLongDirect mtPointer.iId, GWL_ID, iNew
End Property

Public Function SetStyle( _
   Optional ByVal iStyleOr As eWindowStyle, _
   Optional ByVal iStyleAndNot As eWindowStyle) _
                As Boolean
    SetStyle = mWindow.WindowSetLong(mtPointer.iId, GWL_STYLE, iStyleOr, iStyleAndNot)
End Function

Public Property Get Style() As Long
    Style = mWindow.WindowGetLong(mtPointer.iId, GWL_STYLE)
End Property

Public Property Let Style(ByVal iNew As Long)
    mWindow.WindowSetLongDirect mtPointer.iId, GWL_STYLE, iNew
End Property

Public Function SetStyleEx( _
   Optional ByVal iStyleExOr As eWindowStyleEx, _
   Optional ByVal iStyleExAndNot As eWindowStyleEx) _
                As Boolean
    SetStyleEx = mWindow.WindowSetLong(mtPointer.iId, GWL_EXSTYLE, iStyleExOr, iStyleExAndNot)
End Function

Public Property Get StyleEx() As Long
    StyleEx = mWindow.WindowGetLong(mtPointer.iId, GWL_EXSTYLE)
End Property

Public Property Let StyleEx(ByVal iNew As Long)
    mWindow.WindowSetLongDirect mtPointer.iId, GWL_EXSTYLE, iNew
End Property

Public Property Get Property( _
            ByRef sPropName As String) _
                As Long
    Property = mWindow.WindowProp(mtPointer.iId, sPropName)
End Property

Public Property Let Property( _
            ByRef sPropName As String, _
            ByVal iNew As Long)
    mWindow.WindowProp(mtPointer.iId, sPropName) = iNew
End Property

Public Function PropertyRemove( _
            ByRef sPropName As String) _
                As Boolean
    PropertyRemove = mWindow.WindowRemoveProp(mtPointer.iId, sPropName)
End Function

Public Function SetFocus() As Long
    SetFocus = mWindow.WindowSetFocus(mtPointer.iId)
End Function

Public Function PeekMsg( _
            ByRef iMsg As tMsg, _
   Optional ByVal iFilterMin As Long, _
   Optional ByVal iFilterMax As Long, _
   Optional ByVal bRemove As Boolean) _
                As Long
    PeekMsg = WindowPeekMsg(mtPointer.iId, iMsg, iFilterMin, iFilterMax, bRemove)
End Function

Public Function GetMsg( _
            ByRef iMsg As tMsg, _
   Optional ByVal iFilterMin As Long, _
   Optional ByVal iFilterMax As Long) _
                As Long
    GetMsg = WindowGetMsg(mtPointer.iId, iMsg, iFilterMin, iFilterMax)
End Function

Public Function SendMsg( _
            ByVal iMsg As eMsg, _
            ByVal wParam As Long, _
            ByVal lParam As Long) _
                As Long
    SendMsg = WindowSendMsg(mtPointer.iId, iMsg, wParam, lParam)
End Function

Public Function SendMsgStr( _
            ByVal iMsg As eMsg, _
            ByVal wParam As Long, _
            ByRef lParam As String) _
                As Long
    SendMsgStr = WindowSendMsgStr(mtPointer.iId, iMsg, wParam, lParam)
End Function

Public Function PostMsg( _
            ByVal iMsg As eMsg, _
            ByVal wParam As Long, _
            ByVal lParam As Long) _
                As Long
    PostMsg = WindowPostMsg(mtPointer.iId, iMsg, wParam, lParam)
End Function

Public Function Redraw( _
            ByVal iFlags As eRedrawFlags) _
                As Long
    Redraw = mWindow.WindowRedraw(mtPointer.iId, iFlags)
                
End Function

Public Sub Invalidate(Optional ByVal bErase As Boolean)
    WindowInvalidate mtPointer.iId, bErase
End Sub

Public Sub InvalidateRect(ByRef tRect As tRect, Optional ByVal bErase As Boolean)
    WindowInvalidateRect mtPointer.iId, tRect, bErase
End Sub

Public Sub ZOrder()
    WindowZOrder mtPointer.iId
End Sub

Public Sub Show(ByVal iShow As eSWCmd)
    WindowShow mtPointer.iId, iShow
End Sub

Public Sub Update()
    WindowUpdate mtPointer.iId
End Sub

'</Win Api Interface>
'</Public Interface>

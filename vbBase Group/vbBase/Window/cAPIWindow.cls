VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cApiWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
'cApiWindow.cls                     7/4/04
'
'           PURPOSE:
'               Represents an api window, whether created using the cApiWindows collection or not,
'               and delegates to many common api calls.
'
'           MODULES CALLED FROM THIS MODULE:
'               mWindow
'
'==================================================================================================

'1.  Friendly Interface -  Used internally to initialize the modular variables.
'2.  Win Api Interface  -  Methods which delegate to Api calls you may find useful.

Private msClass  As String      'stores the class name of this window
Private mhWnd    As Long        'store the window's handle

'<Private Interface>
'</Private Interface>

'<Friendly Interface>
Friend Sub fInit(ByVal ihWnd As Long, ByRef sClass As String)
    mhWnd = ihWnd
    msClass = sClass
End Sub
'</Friendly Interface>

'<Public Interface>
Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = 0
    hWnd = mhWnd
End Property

Public Property Get ClassName() As String
    ClassName = msClass
End Property

'<Win Api Interface>
Public Function CallDefWndProc( _
            ByVal iMsg As eMsg, _
            ByVal wParam As Long, _
            ByVal lParam As Long) _
                As Long
    CallDefWndProc = DefWindowProc(mhWnd, iMsg, wParam, lParam)
End Function

Public Property Get Enabled() As Boolean
    Enabled = WindowIsEnabled(mhWnd)
End Property

Public Property Let Enabled( _
            ByVal bVal As Boolean)
    WindowEnable mhWnd, bVal
End Property

Public Property Get Text() As String
    Text = WindowText(mhWnd)
End Property

Public Property Let Text( _
            ByVal sVal As String)
    WindowText(mhWnd) = sVal
End Property

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
        
        Move = WindowMove(mhWnd, ltRectSize, bRepaint)
        
    End With
    
End Function

Public Function MoveByRect( _
            ByRef tRectSize As tRectSize, _
   Optional ByVal bRepaint As Boolean = True) _
                As Boolean
                
    MoveByRect = mWindow.WindowMove(mhWnd, tRectSize, bRepaint)
                
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
    
    SetPos = WindowSetPos(mhWnd, hWndInsertAfter, ltRectSize, iFlags)
End Function

Public Function SetPosByRect( _
            ByRef tRectSize As tRectSize, _
   Optional ByVal hWndInsertAfter As eSWPhWnd, _
   Optional ByVal iFlags As eSWPFlags) _
                As Long
    SetPosByRect = WindowSetPos(mhWnd, hWndInsertAfter, tRectSize, iFlags)
End Function

Public Function GetPos( _
   Optional ByRef iLeft As Long, _
   Optional ByRef iTop As Long, _
   Optional ByRef iWidth As Long, _
   Optional ByRef iHeight As Long) _
                As Boolean
    Dim ltRectSize As tRectSize
    
    GetPos = WindowGetPos(mhWnd, ltRectSize)
    
    With ltRectSize
        iLeft = .Left
        iTop = .Top
        iWidth = .Width
        iHeight = .Height
    End With
End Function

Public Function GetPosByRect( _
            ByRef tRectSize As tRectSize) _
                As Boolean
        
    GetPosByRect = WindowGetPos(mhWnd, tRectSize)
End Function

Public Function Size( _
            ByVal iWidth As Long, _
            ByVal iHeight As Long, _
   Optional ByVal bRepaint As Boolean) _
                As Boolean
    Dim ltRectSize As tRectSize
    WindowGetPos mhWnd, ltRectSize
    ltRectSize.Width = iWidth
    ltRectSize.Height = iHeight
    WindowMove mhWnd, ltRectSize, bRepaint
End Function

Public Function GetClientDimensions( _
   Optional ByRef iWidth As Long, _
   Optional ByRef iHeight As Long) _
                As Boolean
    GetClientDimensions = mWindow.WindowGetClientDimensions(mhWnd, iWidth, iHeight)
End Function

Public Function GetDc( _
   Optional ByVal bIncludeNonClient As Boolean) _
                As Long
    GetDc = mWindow.WindowGetDC(mhWnd, bIncludeNonClient)
End Function

Public Function ReleaseDc(ByVal hDc As Long) As Long
    ReleaseDc = mWindow.WindowReleaseDC(mhWnd, hDc)
End Function

Public Function Owner() As Long
    Owner = mWindow.WindowGetOwner(mhWnd)
End Function

Public Property Get Parent() As Long
    Parent = mWindow.WindowParent(mhWnd)
End Property

Public Property Let Parent( _
            ByVal iNew As Long)
    mWindow.WindowParent(mhWnd) = iNew
End Property

Public Property Get UserData() As Long
    UserData = mWindow.WindowGetLong(mhWnd, GWL_USERDATA)
End Property

Public Property Let UserData( _
            ByVal iNew As Long)
    mWindow.WindowSetLongDirect mhWnd, GWL_USERDATA, iNew
End Property

Public Property Get ID() As Long
    ID = mWindow.WindowGetLong(mhWnd, GWL_ID)
End Property

Public Property Let ID( _
            ByVal iNew As Long)
    mWindow.WindowSetLongDirect mhWnd, GWL_ID, iNew
End Property

Public Function SetStyle( _
   Optional ByVal iStyleOr As eWindowStyle, _
   Optional ByVal iStyleAndNot As eWindowStyle) _
                As Boolean
    SetStyle = mWindow.WindowSetLong(mhWnd, GWL_STYLE, iStyleOr, iStyleAndNot)
End Function

Public Property Get Style() As Long
    Style = mWindow.WindowGetLong(mhWnd, GWL_STYLE)
End Property

Public Property Let Style(ByVal iNew As Long)
    mWindow.WindowSetLongDirect mhWnd, GWL_STYLE, iNew
End Property

Public Function SetStyleEx( _
   Optional ByVal iStyleExOr As eWindowStyleEx, _
   Optional ByVal iStyleExAndNot As eWindowStyleEx) _
                As Boolean
    SetStyleEx = mWindow.WindowSetLong(mhWnd, GWL_EXSTYLE, iStyleExOr, iStyleExAndNot)
End Function

Public Property Get StyleEx() As Long
    StyleEx = mWindow.WindowGetLong(mhWnd, GWL_EXSTYLE)
End Property

Public Property Let StyleEx(ByVal iNew As Long)
    mWindow.WindowSetLongDirect mhWnd, GWL_EXSTYLE, iNew
End Property

Public Property Get Property( _
            ByRef sPropName As String) _
                As Long
    Property = mWindow.WindowProp(mhWnd, sPropName)
End Property

Public Property Let Property( _
            ByRef sPropName As String, _
            ByVal iNew As Long)
    mWindow.WindowProp(mhWnd, sPropName) = iNew
End Property

Public Function PropertyRemove( _
            ByRef sPropName As String) _
                As Boolean
    PropertyRemove = mWindow.WindowRemoveProp(mhWnd, sPropName)
End Function

Public Function SetFocus() As Long
    SetFocus = mWindow.WindowSetFocus(mhWnd)
End Function

Public Function PeekMsg( _
            ByRef iMsg As tMsg, _
   Optional ByVal iFilterMin As Long, _
   Optional ByVal iFilterMax As Long, _
   Optional ByVal bRemove As Boolean) _
                As Long
    PeekMsg = WindowPeekMsg(mhWnd, iMsg, iFilterMin, iFilterMax, bRemove)
End Function

Public Function GetMsg( _
            ByRef iMsg As tMsg, _
   Optional ByVal iFilterMin As Long, _
   Optional ByVal iFilterMax As Long) _
                As Long
    GetMsg = WindowGetMsg(mhWnd, iMsg, iFilterMin, iFilterMax)
End Function

Public Function SendMsg( _
            ByVal iMsg As eMsg, _
            ByVal wParam As Long, _
            ByVal lParam As Long) _
                As Long
    SendMsg = WindowSendMsg(mhWnd, iMsg, wParam, lParam)
End Function

Public Function SendMsgStr( _
            ByVal iMsg As eMsg, _
            ByVal wParam As Long, _
            ByRef lParam As String) _
                As Long
    SendMsgStr = WindowSendMsgStr(mhWnd, iMsg, wParam, lParam)
End Function

Public Function PostMsg( _
            ByVal iMsg As eMsg, _
            ByVal wParam As Long, _
            ByVal lParam As Long) _
                As Long
    PostMsg = WindowPostMsg(mhWnd, iMsg, wParam, lParam)
End Function

Public Function Redraw( _
            ByVal iFlags As eRedrawFlags) _
                As Long
    Redraw = mWindow.WindowRedraw(mhWnd, iFlags)
                
End Function

Public Sub Invalidate(Optional ByVal bErase As Boolean)
    WindowInvalidate mhWnd, bErase
End Sub

Public Sub InvalidateRect(ByRef tRect As tRect, Optional ByVal bErase As Boolean)
    WindowInvalidateRect mhWnd, tRect, bErase
End Sub

Public Sub ZOrder()
    WindowZOrder mhWnd
End Sub

Public Sub Show(ByVal iShow As eSWCmd)
    WindowShow mhWnd, iShow
End Sub

Public Sub Update()
    WindowUpdate mhWnd
End Sub

'</Win Api Interface>
'</Public Interface>

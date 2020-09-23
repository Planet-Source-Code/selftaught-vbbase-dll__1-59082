VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcWindowClassHub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==================================================================================================
'pcWindowClassHub.cls                   7/4/04
'
'           PURPOSE:
'               Creates windows from a registered class and manages the messages delivered to each window.
'
'           MODULES CALLED FROM THIS MODULE:
'               mVbBaseGeneral.bas
'
'           CLASSES CREATED BY THIS MODULE:
'               cApiClassWindow
'
'==================================================================================================

'1.  Private Interface           - General procedures
'2.  Public Interface            - procedures to add and remove messages for specific clients and to create/destroy windows
'3.  iWindow Interface           - callback from the asm thunk

#Const bVBVMTypeLib = False

Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Implements iWindow

'<Utility Api's>
Private Type tWndClassEx
    cbSize          As Long
    Style           As Long
    lpfnWndProc     As Long
    cbClsExtra      As Long
    cbWndExtra      As Long
    hInstance       As Long
    hIcon           As Long
    hCursor         As Long
    hbrBackground   As Long
    lpszMenuName    As Long
    lpszClassName   As Long
    hIconSm         As Long
End Type

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function RegisterClassEx Lib "user32" Alias "RegisterClassExA" (ByRef ltWNDCLASSEX As tWndClassEx) As Integer
Private Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long

'</Utility Api's>


Private Const DWordBits                 As Long = 32    'don't adjust this one - 32 bits in a long

Private Const MsgMaskCount              As Long = 4     'Adjust this constant up/down to change the limit on messages
                                                        'normally it would be an even power of 2 and not larger than 16
                                                        
Private Const MaxMsgCount               As Long _
                     = MsgMaskCount * DWordBits         'no more than this number of messages will be delivered collectively to
                                                        'all windows created from this class.
                                                        
Private Const AllMessages               As Long = 1
Private Const CustomMessages            As Long = 2

Private Const ClientIndexPropName       As String = "WndCIProp070404" 'store this property against each hWnd created to
                                                                      'eliminate the extra loop in the window procedure.
                                                                      'this also changes the performance degredation from
                                                                      'linear relative to total number of objects creating
                                                                      'windows to linear relative to the number of windows
                                                                      'created by each individual object.

Private Type tClassWindow
    hWnd                                As Long
    Flags                               As Long         'bit 1 is for all messages, and bit 2 indicates that the window receives the messages based on it's own msgmask
    iMsgMask(0 To MsgMaskCount - 1&)    As Long         'masks indicating which messages are received
End Type

Private Type tClassClient
    iPtr                                As Long
    iControl                            As Long
    iWindowCount                        As Long
    tWindows()                          As tClassWindow
End Type

Private miDefMsgMask(0 To MsgMaskCount - 1&) As Long
Private mbDefAllMsgs                    As Boolean      'if default is all messages

Private miMsgTable()                    As Long         'this array will never grow beyond MaxMsgCount # of elements
Private miMsgCount                      As Long
Private miAllMsgCount                   As Long         'count the number of windows that have requested all messages

Private mtClients()                     As tClassClient 'client data store
Private miClientCount                   As Long         'number of clients defined in the array

Private miRegisterCount                 As Long         'Outstanding calls to Register

Private msClass                         As String       'store the class name

Private moWindow                        As pcWindowClass 'worker - asm thunk

Private mbCreating                      As Boolean      'use these to allow default messages to be sent while creating a window.
Private miCreatingPtr                   As Long

Private Const cApiClsWnds = "cApiClassWindows"
Private Const cApiClsWnd = "cApiClassWindow"

'<Private Interface>
Private Sub Class_Initialize()
    Set moWindow = New pcWindowClass
    Set moWindow.Owner = Me
End Sub

Private Sub Class_Terminate()
    Set moWindow = Nothing
    If Len(msClass) > 0 Then                                 'If a class was registered
      Call UnregisterClass(msClass, App.hInstance)           'Unregister the window class
    End If
End Sub

Private Sub PatchTable()
    
    Dim iPatchVal As Long
    
    If miAllMsgCount = 0& Then                  'If no one wants all messages
        iPatchVal = miMsgCount                  'patch the real count
    Else                                        'else (Someone wants all messages)
        'ALL_MESSAGES
        iPatchVal = -1&                         'patch -1, telling the asm thunk we want all messages
    End If
    
    moWindow.PatchTable miMsgTable, iPatchVal   'patch the values into the thunk
End Sub

Private Function pFindClient( _
            ByVal iPtr As Long, _
   Optional ByRef iFirstAvailable As Long) _
            As Long
    'find the index in the main array for this object
    
    Dim liTemp As Long
    
    iFirstAvailable = Undefined
    
    For pFindClient = 0& To miClientCount - 1&                  'loop through each object
        liTemp = mtClients(pFindClient).iPtr                    'store the pointer
        If liTemp <> 0& And liTemp <> Undefined Then            'validate the pointer
            If liTemp = iPtr Then Exit Function                 'if the pointer matches, then bail
        Else                                                    'if the pointer is invalid, it may be the first index
            If iFirstAvailable = Undefined Then iFirstAvailable = pFindClient
        End If
    Next
    
    pFindClient = Undefined                                     'client was not found
    'Debug.Assert False
End Function

Private Function pFindWindow( _
            ByVal iClientIndex As Long, _
            ByVal hWnd As Long, _
   Optional ByRef iFirstAvailable As Long) _
            As Long
    
    'find the window in the main client data store
    
    Dim liTemp As Long
    
    iFirstAvailable = Undefined
    
    With mtClients(iClientIndex)
        For pFindWindow = 0 To .iWindowCount - 1&               'loop through each window
            liTemp = .tWindows(pFindWindow).hWnd                'store the hWnd
            If liTemp <> 0& Then                                'if it is valid
                If liTemp = hWnd Then Exit Function             'if it matches then bail
            Else
                If iFirstAvailable = Undefined Then iFirstAvailable = pFindWindow   'if the index is invalid, it may be the first available
            End If
        Next
    End With
    
    pFindWindow = Undefined                                     'window not found

End Function

Private Function pValidatePointer( _
            ByRef tPointer As tPointer) _
                As Boolean
    'validate a pointer to a client object to ensure that the cApiClassWindows object
    'has not been held past it's lifetime.
    
    If tPointer.iIndex = Undefined Then tPointer.iIndex = pFindClient(tPointer.iId) '
    If tPointer.iIndex > Undefined Then
        If tPointer.iIndex < miClientCount Then _
            pValidatePointer = (mtClients(tPointer.iIndex).iPtr = tPointer.iId)
    End If
    'Debug.Assert pValidatePointer
End Function

Private Function pValidateItemPointer( _
            ByRef tItemPointer As tItemPointer) _
                As Boolean
    
    'validate a pointer to a client object to ensure that the cApiClassWindow object
    'has not been held past it's lifetime.
    
    If tItemPointer.iIndex < miClientCount Then
        With mtClients(tItemPointer.iIndex)
            If tItemPointer.iItemIndex < .iWindowCount Then
                pValidateItemPointer = (.tWindows(tItemPointer.iItemIndex).hWnd = tItemPointer.iId)
            End If
        End With
    End If
    
    Debug.Assert pValidateItemPointer
            
End Function
            

'private implementation of ArrRedim to allow string typing
Private Sub ArrRedimT( _
            ByRef tArray() As tClassWindow, _
            ByVal iElements As Long, _
   Optional ByVal bPreserve As Boolean = True)
    
    'Adjust from elements to zero-based upper bound
    'iElements is now a zero-based array bound
    iElements = iElements - 1&

    Dim liNewUbound As Long: liNewUbound = ArrAdjustUbound(iElements)

    'If we don't have enough room already, then redim the array
    If liNewUbound > ArrUboundT(tArray) Then
        If bPreserve Then _
            ReDim Preserve tArray(0 To liNewUbound) _
        Else _
            ReDim tArray(0 To liNewUbound)
    End If
End Sub

Private Function ArrUboundT( _
            ByRef tArray() As tClassWindow _
        ) As Long
    On Error Resume Next
    ArrUboundT = UBound(tArray)
    If Err.Number <> 0& Then ArrUboundT = Undefined
End Function

Private Sub ArrRedimT2( _
            ByRef tArray() As tClassClient, _
            ByVal iElements As Long, _
   Optional ByVal bPreserve As Boolean = True)
    
    'Adjust from elements to zero-based upper bound
    'iElements is now a zero-based array bound
    iElements = iElements - 1&

    Dim liNewUbound As Long: liNewUbound = ArrAdjustUbound(iElements)

    'If we don't have enough room already, then redim the array
    If liNewUbound > ArrUboundT2(tArray) Then
        If bPreserve Then _
            ReDim Preserve tArray(0 To liNewUbound) _
        Else _
            ReDim tArray(0 To liNewUbound)
    End If
End Sub

Private Function ArrUboundT2( _
            ByRef tArray() As tClassClient _
        ) As Long
    On Error Resume Next
    ArrUboundT2 = UBound(tArray)
    If Err.Number <> 0& Then ArrUboundT2 = Undefined
End Function

Private Function pMsgIsDead( _
                ByVal iTableIndex As Long, _
                ByVal iMask As Long) _
                    As Boolean
    
    'determine if a message is being used by any window.
    
    Dim liEachClient As Long
    Dim liEachWindow As Long
    
    For liEachClient = 0& To miClientCount - 1&                                     'loop through each client
        With mtClients(liEachClient)
            For liEachWindow = 0& To mtClients(liEachClient).iWindowCount - 1&
                If mtClients(liEachClient).iPtr <> Undefined Then                   'if the client is defined
                    If CBool(.tWindows(iTableIndex).iMsgMask(iTableIndex) And iMask) _
                        Then Exit Function                                          'if the client wants this message, bail
                End If
            Next
        End With
    Next
    
    If CBool(miDefMsgMask(iTableIndex) And iMask) _
        Then Exit Function                                                          'Msg is still alive it it is a default message.
    
    pMsgIsDead = True                                                               'if we made it here, nobody cares about this message
End Function

Private Sub pTranslateMsgIndex( _
            ByVal iMsgTableIndex As Long, _
            ByRef iClientTableIndex As Long, _
            ByRef iMask As Long)

    iClientTableIndex = iMsgTableIndex \ DWordBits                      'get the index in the tSubclassClient msg array
    iMask = BitMask(iMsgTableIndex Mod DWordBits)                       'get the mask that identifies this message
    
    'internal error
    Debug.Assert iClientTableIndex < MsgMaskCount                       'alert developer in case we've exceeded the max
                                                                        'This will cause a subscript out of range error,
                                                                        'which is particulary bad because of the 'remove
                                                                        'array bounds checks' compiler option that is used.
End Sub

Private Function pDelAllMessages( _
                ByRef iMsgMask() As Long) _
                    As Long

    Dim liMask     As Long      'a mask containing each bit one at a time
    Dim liTempMask As Long      'the current msgmask
    Dim iEachBit   As Long      'counter
    Dim iEachIndex As Long      'counter
    
    For iEachBit = 0 To BitMax                              'loop through each bit
        liMask = BitMask(iEachBit)                          'store the bitmask for this bit
        For iEachIndex = 0 To MsgMaskCount - 1&             'for each bit, loop through each index
            If ((iEachBit * DWordBits) + iEachIndex) = miMsgCount Then Exit Function
            liTempMask = iMsgMask(iEachIndex)              'store the msgmask at this index
            If CBool(liTempMask And liMask) Then            'if the client has requested this message
                pDelAllMessages = pDelAllMessages + 1&      'inc the count of messages removed
                iMsgMask(iEachIndex) = (liTempMask And Not liMask) 'remove the message from the client's table
                If pMsgIsDead(iEachIndex, liMask) _
                    Then miMsgTable(iEachIndex * DWordBits + iEachBit) = Undefined 'If nobody wants the message anymore, remove it from the table
            End If
        Next
    Next

End Function

Private Function pAddMsg( _
            ByVal iClientIndex As Long, _
            ByVal iWindowIndex As Long, _
            ByVal iMsg As eMsg) _
                As Boolean
    
    Dim liFirst As Long
    Dim liIndex As Long
    
    If iMsg <> ALL_MESSAGES Then                                                            'if we're adding a specific message
        
        liIndex = ArrFindInt(miMsgTable, miMsgCount, iMsg, liFirst)                         'try to find the message in the main table
        If liIndex = Undefined Then                                                         'if the message is not found
            If liFirst = Undefined Then                                                     'if there is not an available slot
                If miMsgCount = MaxMsgCount Then Exit Function                              'don't exceed the maximum
                liIndex = miMsgCount
                miMsgCount = miMsgCount + 1&
                ArrRedim miMsgTable, miMsgCount, True                                       'resize the array
            Else
                liIndex = liFirst                                                           'just use the availble slot
            End If
            miMsgTable(liIndex) = iMsg                                                      'store the message
        End If
        
        Dim liClientTableIndex As Long
        Dim liMask As Long
        Dim liCurrent As Long
            
        pTranslateMsgIndex liIndex, liClientTableIndex, liMask                              'get the index and bit for this message
        
        If iClientIndex <> Undefined Then 'custom message                                   'if the message belongs in a liclient's table
            With mtClients(iClientIndex).tWindows(iWindowIndex)
                If Not (.Flags And CustomMessages) Then .Flags = .Flags Or CustomMessages   'ensure that the client is receiving the custom, not default messages
                liCurrent = .iMsgMask(liClientTableIndex)
                pAddMsg = Not CBool(liCurrent And liMask)                                   'return whether the bit was already set
                If pAddMsg Then .iMsgMask(liClientTableIndex) = (liCurrent Or liMask)       'ensure that the bit is now set
            End With
        Else 'default message                                                               'if the message belongs in the default table
            liCurrent = miDefMsgMask(liClientTableIndex)
            pAddMsg = Not CBool(liCurrent And liMask)                                       'return whether the bit was already set
            If pAddMsg Then miDefMsgMask(liClientTableIndex) = (liCurrent Or liMask)        'ensure that the bit is set
        End If
        
    Else 'ALL_MESSAGES
        
        If iClientIndex <> Undefined Then 'custom message                                   'if the message belongs in a client's table
            With mtClients(iClientIndex).tWindows(iWindowIndex)
                If Not CBool(.Flags And AllMessages) Then                                   'if not already receiving all messages
                    pAddMsg = True                                                          'return true
                    .Flags = .Flags Or AllMessages                                          'receive all messages
                End If
            End With
        Else 'default message
            If Not mbDefAllMsgs Then                                                        'if not already receiving all messages
                pAddMsg = True                                                              'return true
                mbDefAllMsgs = True                                                         'receive all messages
            End If
        End If
        
        If pAddMsg Then miAllMsgCount = miAllMsgCount + 1&                                  'Inc the ref count for all messages
        
    End If
    
    PatchTable                                                                              'Ensure the correct table locations
    
End Function

Private Function pDelMsg( _
            ByVal iClientIndex As Long, _
            ByVal iWindowIndex As Long, _
            ByVal iMsg As eMsg) _
                As Long
                
    Dim liMsgIndex          As Long
    Dim liClientMsgIndex    As Long
    Dim liClientMsgMask     As Long
    Dim liTempMask          As Long
    
    If iMsg <> ALL_MESSAGES Then                                                'If we're adding a single message
        liMsgIndex = ArrFindInt(miMsgTable, miMsgCount, iMsg)                   'See if the message exists in the table
            
        If liMsgIndex <> Undefined Then                                         'if the message exists
            pTranslateMsgIndex liMsgIndex, liClientMsgIndex, liClientMsgMask    'get the array index and bit mask for our MsgMask
            
            If iClientIndex <> Undefined Then 'custom message
                With mtClients(iClientIndex).tWindows(iWindowIndex)
                    If CBool(.Flags And CustomMessages) Then
                        liTempMask = .iMsgMask(liClientMsgIndex)                'Store the current mask
                        pDelMsg = CBool(liTempMask And liClientMsgMask)         'succeed only if subscribed to this msg
                        .iMsgMask(liClientMsgIndex) = _
                            liTempMask And Not liClientMsgMask                  'remove the bit from the MsgMask
                    End If
                End With
            Else 'default message
                liTempMask = miDefMsgMask(liClientMsgIndex)                     'Store the current mask
                pDelMsg = CBool(liTempMask And liClientMsgMask)                 'succeed only if subscribed to this msg
                miDefMsgMask(liClientMsgIndex) = _
                    liTempMask And Not liClientMsgMask                          'remove the bit from the MsgMask
            End If
            
            If pDelMsg Then                                                     'if we succeeded then we may have a message that's not delivered to any object
                If pMsgIsDead(liClientMsgIndex, liClientMsgMask) Then _
                    miMsgTable(liMsgIndex) = Undefined
            End If
            
        End If
    Else                                                                        'deleting all messages
        If iClientIndex <> Undefined Then
            With mtClients(iClientIndex).tWindows(iWindowIndex)
                pDelMsg = CBool(.Flags And AllMessages)                         'succeed if receiving all messages
                If pDelMsg Then                                                 'if succeeded
                    miAllMsgCount = miAllMsgCount = 1&                          'decrement the refcount
                    .Flags = .Flags And Not AllMessages                         'remove the bit
                End If
                pDelMsg = pDelMsg Or CBool(pDelAllMessages(.iMsgMask))          'also succeed if any custom messages to remove
                
            End With
        Else
            If mbDefAllMsgs Then miAllMsgCount = miAllMsgCount - 1&             'decrement the count if necessary
            pDelMsg = mbDefAllMsgs Or CBool(pDelAllMessages(miDefMsgMask))      'succeed if receiving all messages or any default messages
            mbDefAllMsgs = False                                                'dont receive all messages by default.
        End If
    End If
    
    For liMsgIndex = miMsgCount - 1& To 0& Step -1&
        If miMsgTable(liMsgIndex) <> Undefined Then Exit For                    'find the lowest necessary count
    Next
    
    If liMsgIndex < miMsgCount - 1& Then                                        'if the count is lower than it is already
        miMsgCount = liMsgIndex + 1&                                            'set the new count
        PatchTable                                                              'patch the new table.  This is necessary if ALL_MESSAGES was deleted.
    End If
    
End Function

'</Private Interface>

'<Public Interface>
'declared friend b/c VB won't let udts declared in std modules to be passed in public procedures, even
'if the objects are private!
Friend Property Get Active() As Boolean
    Active = (miRegisterCount > 0&)
End Property

Friend Function AddMsg( _
            ByRef tPointer As tItemPointer, _
            ByVal iMsg As eMsg) _
                As Boolean
 
    If pValidateItemPointer(tPointer) _
        Then AddMsg = pAddMsg(tPointer.iIndex, tPointer.iItemIndex, iMsg) _
        Else gErr vbbItemDetached, cApiClsWnd
    
    'client or window not found or message already there
    Debug.Assert AddMsg
    
    
End Function

Friend Function DelMsg( _
            ByRef tPointer As tItemPointer, _
            ByVal iMsg As eMsg) _
                As Boolean
                
    If pValidateItemPointer(tPointer) Then _
        DelMsg = pDelMsg(tPointer.iIndex, tPointer.iItemIndex, iMsg) _
        Else gErr vbbItemDetached, cApiClsWnd
    
    'client, window or message not found
    Debug.Assert DelMsg
End Function

Friend Function AddDefMsg( _
                    ByVal iMsg As eMsg _
                ) As Boolean
 
    AddDefMsg = pAddMsg(-1, -1, iMsg)
 
    'message already there?
    Debug.Assert AddDefMsg
  
End Function

Friend Function DelDefMsg( _
                       ByVal iMsg As eMsg _
                ) As Boolean
                
    DelDefMsg = pDelMsg(-1, -1, iMsg)
    
    'message not there?
    Debug.Assert DelDefMsg
  
End Function

Friend Property Get DefMessages( _
            ByRef tPointer As tItemPointer) _
                As Boolean
    'return whether the object receives default messages
    If pValidateItemPointer(tPointer) _
        Then DefMessages = Not CBool(mtClients(tPointer.iIndex).tWindows(tPointer.iItemIndex).Flags And CustomMessages) _
        Else gErr vbbItemDetached, cApiClsWnd
    
End Property

Friend Property Let DefMessages( _
            ByRef tPointer As tItemPointer, _
            ByVal bVal As Boolean)
    'set whether the object receives default messages
    If pValidateItemPointer(tPointer) Then
        With mtClients(tPointer.iIndex).tWindows(tPointer.iItemIndex)
            If bVal _
            Then .Flags = (.Flags And Not CustomMessages) _
            Else .Flags = (.Flags Or CustomMessages)
        End With
    Else
        gErr vbbItemDetached, cApiClsWnd
    End If

End Property

Friend Function Register( _
                     ByRef sClass As String, _
            Optional ByVal colBackground As Long = &HFFFFFF, _
            Optional ByVal Style As eClassStyle, _
            Optional ByVal hCursor As Long, _
            Optional ByVal hIcon As Long, _
            Optional ByVal hIconSm As Long, _
            Optional ByVal cbClassExtra As Long, _
            Optional ByVal cbWndExtra As Long) _
                 As Boolean

    If miRegisterCount = 0& Then
        Dim wc      As tWndClassEx
        Dim sAscii  As String
        
        msClass = sClass                                            'Store the class name
        sAscii = StrConv(msClass & vbNullChar, vbFromUnicode)       'Get the ascii (rather than unicode) version of the class name
          
        With wc
            .cbSize = Len(wc)                                       'Size of the window class type
            .cbClsExtra = cbClassExtra                              'Number of class extra bytes
            .cbWndExtra = cbWndExtra                                'Number of window extra bytes
            .hbrBackground = CreateSolidBrush(colBackground)        'Class background
            .hCursor = hCursor                                      'Class cursor
            .hIcon = hIcon                                          'Class icon
            .hIconSm = hIconSm                                      'Class small icon
            .hInstance = App.hInstance                              'Application instance handle
            .lpfnWndProc = moWindow.AddrWndProc                     'Class WndProc address
            .Style = Style                                          'Class style
            .lpszClassName = StrPtr(sAscii)                         'Class name
        End With

        Register = (RegisterClassEx(wc) <> 0)                       'Register the window class
        
        moWindow.ClassName = msClass
        miRegisterCount = 1&
        
        'Windows Registration failed!
        Debug.Assert Register
    Else
        Debug.Assert False
        If msClass = sClass Then _
            Register = True: miRegisterCount = miRegisterCount + 1&
        'can't register different classes on one pcWindowClassHub object
        Debug.Assert Register
    End If
End Function

Friend Function Unregister() As Boolean
    
    miRegisterCount = miRegisterCount - 1&                  'decrement the count
    
    If miRegisterCount <= 0& Then                           'If we've unregistered the last register
        miRegisterCount = 0&                                'make sure we don't go too low
        Unregister = _
            UnregisterClass(msClass, App.hInstance) <> 0&   'call the api
        If Unregister Then msClass = vbNullString           'delete the stored class name
    Else
        Debug.Assert False
        Unregister = True                                   'wait until the last one is released
    End If

    'unregister api failed
    Debug.Assert Unregister
    
End Function

Friend Function Create( _
                    ByRef tPointer As tPointer, _
           Optional ByVal dwExStyle As eWindowStyleEx, _
           Optional ByVal dwStyle As eWindowStyle, _
           Optional ByVal x As Long = 0, _
           Optional ByVal y As Long = 0, _
           Optional ByVal nWidth As Long = 0, _
           Optional ByVal nHeight As Long = 0, _
           Optional ByRef sCaption As String = "", _
           Optional ByVal hWndParent As Long = 0, _
           Optional ByVal hMenu As Long = 0, _
           Optional ByVal lParam As Long = 0 _
                ) As cApiClassWindow
  
    Debug.Assert (Len(msClass) > 0)                             'LOGIC ERROR! Class name not defined
 
    Dim liClientIndex As Long
    Dim liWindowIndex As Long
    Dim liFirst As Long
 
    If pValidatePointer(tPointer) Then                          'if the client object pointer is valid
        liClientIndex = tPointer.iIndex                          'use the index
    Else
        liClientIndex = pFindClient(tPointer.iId, liFirst)      'search for the object
        
        If liClientIndex = Undefined Then                       'if the object was not found
            If liFirst = Undefined Then                         'if there is an open slot
                liClientIndex = miClientCount                   'inc the count
                miClientCount = miClientCount + 1&
                ArrRedimT2 mtClients, miClientCount, True       'resize the array
            Else
                liClientIndex = liFirst                         'use the open slot
            End If
            With mtClients(liClientIndex)
               .iPtr = tPointer.iId                             'store the pointer
               .iWindowCount = 0&                               'init the window count
            End With
            tPointer.iIndex = liClientIndex                     'point the pointer to the current index
        Else
            tPointer.iIndex = liClientIndex                     'point to the current index
                                                                'this can only happen if all windows are released and then
                                                                'some are added, and the client is put in a different place.
        End If
    End If
        
   'Create the window
   Dim ltPointer As tItemPointer
                                                                'create the window, setting the flags to cause the
   mbCreating = True                                            'default messages to be delivered
   miCreatingPtr = tPointer.iId
   ltPointer.iId = CreateWindowEx(dwExStyle, msClass, sCaption, dwStyle, x, y, nWidth, nHeight, hWndParent, hMenu, App.hInstance, lParam)
   miCreatingPtr = 0&
   mbCreating = False
   
   Debug.Assert ltPointer.iId   'ensure that the window could be created
   
   If ltPointer.iId Then
       SetProp ltPointer.iId, ClientIndexPropName, liClientIndex    'set the property for quick access to the client index
       
       With mtClients(liClientIndex)
           For liWindowIndex = 0& To .iWindowCount - 1&
               If .tWindows(liWindowIndex).hWnd = 0& Then
                   .tWindows(liWindowIndex).hWnd = ltPointer.iId    'find the first available index, and store the handle
                   Exit For
               End If
           Next
           
           If liWindowIndex = .iWindowCount Then                    'if there is no available index then increase the size
               .iWindowCount = liWindowIndex + 1&                   'of the array
               ArrRedimT .tWindows, .iWindowCount, True
               .tWindows(liWindowIndex).hWnd = ltPointer.iId
           End If
       
           ltPointer.iIndex = liClientIndex                         'pointer points to this client and this window
           ltPointer.iItemIndex = liWindowIndex
       
           With .tWindows(liWindowIndex)
               .Flags = 0&                                          'init the flags
               ZeroMemory .iMsgMask(0), (4& * MsgMaskCount)         'init the msgmask
           End With
            
            Set Create = New cApiClassWindow
            Create.fInit ltPointer, msClass                         'initialize the return object
       End With
   Else
        gErr vbbApiError, cApiClsWnds                               'window creation failed
   End If
End Function

Friend Sub Destroy(ByRef tPointer As tPointer, ByVal hWnd As Long)
    
    Dim liWindowIndex As Long
    

    If pValidatePointer(tPointer) Then
        With mtClients(tPointer.iIndex)
            liWindowIndex = pFindWindow(tPointer.iIndex, hWnd)                      'Find the index of the window
            
            If liWindowIndex = Undefined Then gErr vbbItemDetached, cApiClsWnds     'if the window wasn't found, error
            
            With .tWindows(liWindowIndex)
                RemoveProp .hWnd, ClientIndexPropName                               'remove the property
                If IsWindow(.hWnd) Then                                             'if the window is valid
                    If DestroyWindow(.hWnd) = 0& Then gErr vbbApiError, cApiClsWnds 'destroy the window, erring on failure
                Else
                    .hWnd = 0&                                                      'did you destroy my window??
                End If
                pDelAllMessages .iMsgMask                                           'ensure all messages are deleted if necessary
                .hWnd = 0&
                .Flags = 0&
            End With
            
            If .iWindowCount = liWindowIndex + 1& Then
                For liWindowIndex = liWindowIndex - 1& To 0& Step -1&
                    If .tWindows(liWindowIndex).hWnd <> 0& Then Exit For            'find the lowest valid index
                Next
                .iWindowCount = liWindowIndex + 1&                                  'set the new count
            End If
            
        End With
    Else
        gErr vbbItemDetached, cApiClsWnds                                           'holding an object past it's lifetime
    End If
End Sub

Friend Function Item( _
            ByRef tPointer As tPointer, _
            ByVal hWnd As Long) _
                As cApiClassWindow
    If pValidatePointer(tPointer) Then
        Dim ltItemPointer As tItemPointer
        ltItemPointer.iItemIndex = pFindWindow(tPointer.iIndex, hWnd)               'find the window index
                                                                                    'if the window wasn't found, err
        If ltItemPointer.iItemIndex = Undefined Then gErr vbbKeyNotFound, cApiClsWnds
        
        ltItemPointer.iId = hWnd                                                    'point to the window
        ltItemPointer.iIndex = tPointer.iIndex
        
        Set Item = New cApiClassWindow
        Item.fInit ltItemPointer, msClass                                           'init the item
        
    Else
        gErr vbbItemDetached, cApiClsWnds                                           'holding the object past it's lifetime
    End If
End Function

Friend Function Exists( _
                    ByRef tPointer As tPointer, _
                    ByVal hWnd As Long _
                ) As Boolean
    
    If pValidatePointer(tPointer) _
        Then Exists = (pFindWindow(tPointer.iIndex, hWnd) <> Undefined)
    
End Function

Friend Function Count( _
                    ByRef tPointer As tPointer _
                ) As Long
    
    If pValidatePointer(tPointer) Then              'validate the pointer
        
        Dim i As Long
        
        With mtClients(tPointer.iIndex)
            Count = .iWindowCount                   'init count to maximum value
            For i = 0 To .iWindowCount - 1&         'loop through all the windows
                If .tWindows(i).hWnd = 0& Then      'if it is not valid, decrement the count
                    Count = Count - 1&
                End If
            Next
        End With
        
    Else
        gErr vbbItemDetached, cApiClsWnds
        
    End If
    
End Function

Friend Function GetControl(ByRef tPointer As tPointer) As Long
    If pValidatePointer(tPointer) Then GetControl = mtClients(tPointer.iIndex).iControl
End Function

Friend Sub Enum_NextItem( _
                    ByRef tPointer As tPointer, _
                    ByRef tEnum As tEnum, _
                    ByRef vNextItem As Variant, _
                    ByRef bNoMore As Boolean)
    
    If pValidatePointer(tPointer) Then                              'validate the pointer
        tEnum.iIndex = tEnum.iIndex + 1&                            'inc the index
        
        With mtClients(tPointer.iIndex)
            
            If .iControl <> tEnum.iControl Then gErr vbbCollChangedDuringEnum, cApiClsWnds
            
            Do While tEnum.iIndex < .iWindowCount                   'get the next valid index
                If .tWindows(tEnum.iIndex).hWnd <> 0& Then Exit Do
                tEnum.iIndex = tEnum.iIndex + 1&
            Loop
            
            If tEnum.iIndex < .iWindowCount Then                    'if there is a valid index
                Dim tItemPointer As tItemPointer
                tItemPointer.iId = .tWindows(tEnum.iIndex).hWnd     'point to this window
                tItemPointer.iIndex = tPointer.iIndex
                tItemPointer.iItemIndex = tEnum.iIndex
                
                Dim oTemp As cApiClassWindow
                Set oTemp = New cApiClassWindow                     'create an object
                oTemp.fInit tItemPointer, msClass                   'initialize it
                Set vNextItem = oTemp
            Else
                bNoMore = True                                      'no more items
            End If
        
        End With
        
    Else
        'error will be raised in the NewEnum property if the client has detached,
        'if all windows are released between calls to this procedure, then simply stop iterating
        bNoMore = True
    End If
    
    
End Sub

Friend Sub Enum_Skip( _
                    ByRef tPointer As tPointer, _
                    ByRef tEnum As tEnum, _
                    ByVal iSkipCount As Long, _
                    ByRef bSkippedAll As Boolean)

    If pValidatePointer(tPointer) Then                              'validate the pointer
        
        Dim liSkipped As Long
       
        With mtClients(tPointer.iIndex)
            
            If .iControl <> tEnum.iControl Then gErr vbbCollChangedDuringEnum, cApiClsWnds
            
            For tEnum.iIndex = tEnum.iIndex + 1& To .iWindowCount - 1&
                If .tWindows(tEnum.iIndex).hWnd <> 0& Then liSkipped = liSkipped + 1&
                If liSkipped = iSkipCount Then Exit For
            Next
            
            bSkippedAll = CBool(liSkipped = iSkipCount)
       
        End With
        
    Else
        bSkippedAll = False
    End If

End Sub

Friend Function AbsCount() As Long
    Dim liCount As Long
    Dim iEachClient As Long
    Dim iEachWindow As Long
    
    For iEachClient = 0 To miClientCount - 1&                                   'loop through each client
        With mtClients(iEachClient)
            liCount = .iWindowCount                                             'loop through each window
            For iEachWindow = 0& To .iWindowCount - 1&
                If .tWindows(iEachWindow).hWnd = 0& Then liCount = liCount - 1& 'if it's valid, inc the count
            Next
        End With
        AbsCount = AbsCount + liCount
    Next
    
End Function

Friend Function Clear( _
                    ByRef tPointer As tPointer _
                ) As Long
    Dim i As Long
    
    If pValidatePointer(tPointer) Then                                              'validate the pointer
        
        With mtClients(tPointer.iIndex)
            For i = 0& To .iWindowCount - 1&                                        'loop through each window
                If .tWindows(i).hWnd Then
                    If IsWindow(.tWindows(i).hWnd) Then                             'if the window is valid
                        If DestroyWindow(.tWindows(i).hWnd) Then                    'destroy it
                            .tWindows(i).hWnd = 0&                                  'reset the flags and MsgMask
                            .tWindows(i).Flags = 0&
                            ZeroMemory .tWindows(i).iMsgMask(0), MsgMaskCount * 4&
                            Clear = Clear + 1&                                      'inc the number cleared
                        End If
                    Else
                        .tWindows(i).hWnd = 0&                                      'reset the flags and MsgMask
                        .tWindows(i).Flags = 0&
                        ZeroMemory .tWindows(i).iMsgMask(0), MsgMaskCount * 4&
                    End If
                End If
            Next
                
            Incr .iControl                                                          'inc the enumeration control number
                
            For .iWindowCount = .iWindowCount - 1& To 0& Step -1&
                If .tWindows(.iWindowCount).hWnd Then                               'find the lowest valid count
                    .iWindowCount = .iWindowCount + 1&                              'if a window wasn't deleted, err
                    gErr vbbApiError, cApiClsWnds
                End If
            Next
            .iWindowCount = .iWindowCount + 1&
        End With
    Else
        gErr vbbItemDetached, cApiClsWnds
    
    End If
    
End Function

Friend Function MsgExists( _
                    ByRef tPointer As tItemPointer, _
                    ByVal iMsg As eMsg) _
                        As Boolean
    If pValidateItemPointer(tPointer) Then                                  'validate the pointer
        Dim liMsgIndex As Long
        Dim liClientTableIndex As Long
        Dim liMask As Long
        
        liMsgIndex = ArrFindInt(miMsgTable, miMsgCount, iMsg)               'find the message in the main table
        
        If liMsgIndex <> Undefined Then                                     'if the message was found
            pTranslateMsgIndex liMsgIndex, liClientTableIndex, liMask       'get the index in the client table
            MsgExists = (mtClients(tPointer.iIndex).tWindows(tPointer.iItemIndex).iMsgMask(liClientTableIndex) And liMask)
        End If
        
    Else
        gErr vbbItemDetached, cApiClsWnd                                    'holding the object past it's lifetime
    End If

End Function

Friend Function MsgCount( _
                    ByRef tPointer As tItemPointer) _
                        As Long
    
    If pValidateItemPointer(tPointer) Then
        Dim liIndex As Long
        Dim liBit As Long
        
        With mtClients(tPointer.iIndex).tWindows(tPointer.iItemIndex)
            For liIndex = 0& To MsgMaskCount - 1&                           'loop through each possible message
                For liBit = 0& To BitMax                                    'for each valid message, inc the count
                    If (liIndex * DWordBits) + liBit = miMsgCount Then Exit Function
                    If CBool(.iMsgMask(liIndex) And BitMask(liBit)) Then MsgCount = MsgCount + 1&
                Next
            Next
        End With
    Else
        gErr vbbItemDetached, cApiClsWnd                                    'holding the object past it's lifetime
    End If
    
End Function

Friend Function GetMessages( _
                    ByRef iOutArray() As Long, _
                    ByRef tPointer As tItemPointer) _
                        As Long
                        
    If pValidateItemPointer(tPointer) Then
        If miMsgCount = 0& Then                                             'if there are no messages
            GetMessages = CLng(CBool(mtClients(tPointer.iIndex).tWindows(tPointer.iItemIndex).Flags And AllMessages))
            Erase iOutArray                                                 'return -1 for ALL_MESSAGES or 0 for none
            Exit Function
        End If
        
        Dim liIndex As Long
        Dim liBit As Long
        ReDim iOutArray(0 To miMsgCount - 1&)
        With mtClients(tPointer.iIndex).tWindows(tPointer.iItemIndex)
            For liIndex = 0& To MsgMaskCount - 1&                                       'loop through each mask
                For liBit = 0& To BitMax                                                'loop through each bit
                    If (liIndex * DWordBits) + liBit = miMsgCount Then Exit Function    'if past the max, exit
                    If CBool(.iMsgMask(liIndex) And BitMask(liBit)) Then                'if the bit is set, add it to the array
                        iOutArray(GetMessages) = miMsgTable((liIndex * DWordBits) + liBit)
                        GetMessages = GetMessages + 1&
                    End If
                Next
            Next
            If GetMessages = 0& Then                                                    'if no array, exit
                Erase iOutArray
                Exit Function
            End If
            ReDim Preserve iOutArray(0 To GetMessages - 1&)                             'trim the array
        End With
    Else
        gErr vbbItemDetached, cApiClsWnd                                                'holding the object past it's lifetime
        
    End If
    
End Function

Friend Function DefMsgExists( _
                    ByVal iMsg As eMsg _
                ) As Boolean
    Dim liIndex As Long
    liIndex = ArrFindInt(miMsgTable, miMsgCount, iMsg)                      'find the message
    
    If liIndex <> Undefined Then
        Dim liClientTableIndex As Long
        Dim liMask As Long
        pTranslateMsgIndex liIndex, liClientTableIndex, liMask              'get the client table index
        DefMsgExists = CBool(miDefMsgMask(liClientTableIndex) And liMask)   'return whether the message was found
    End If

End Function

Friend Function DefMsgCount() As Long
    
    Dim liIndex As Long
    Dim liBit As Long
    
    For liIndex = 0& To MsgMaskCount - 1&                                   'loop through each mask
        For liBit = 0& To BitMax                                            'Loop through each bit
            If (liIndex * DWordBits) + liBit = miMsgCount Then Exit Function 'if past the max, exit
            If CBool(miDefMsgMask(liIndex) And BitMask(liBit)) Then DefMsgCount = DefMsgCount + 1&
        Next
    Next
    
End Function

Friend Function GetDefMessages( _
            ByRef iOutArray() As Long) _
                As Long
    Dim liIndex As Long
    Dim liBit As Long
    If mbDefAllMsgs Then                                                        'if ALL_MESSAGES, return -1
        GetDefMessages = -1&
        Erase iOutArray
        Exit Function
    End If
    
    If miMsgCount = 0& Then                                                     'if no messages, return 0
        Erase iOutArray
        Exit Function
    End If
    ReDim iOutArray(0 To miMsgCount - 1&)                                       'resize the array to the max
        
    For liIndex = 0& To MsgMaskCount - 1&
        For liBit = 0& To BitMax                                                'loop through each mask and bit
            If (liIndex * DWordBits) + liBit = miMsgCount Then Exit For         'if past the max, exit
            If CBool(miDefMsgMask(liIndex) And BitMask(liBit)) Then             'if the bit is set, store the message
                iOutArray(GetDefMessages) = miMsgTable((liIndex * DWordBits) + liBit)
                GetDefMessages = GetDefMessages + 1&                            'inc the count
            End If
        Next
    Next
    
    If GetDefMessages = 0& Then                                                 'if no messages, return 0
        Erase iOutArray
        Exit Function
    End If
    
    ReDim Preserve iOutArray(0 To GetDefMessages - 1&)                          'trim the array
    
    
End Function
Friend Property Get ClassName() As String
    ClassName = msClass
End Property
'</Public Interface>

'<iWindow Interface>
Private Sub iWindow_Proc( _
                    ByRef bHandled As Boolean, _
                    ByRef lReturn As Long, _
                    ByRef hWnd As Long, _
                    ByRef iMsg As eMsg, _
                    ByRef wParam As Long, _
                    ByRef lParam As Long)
            
    Dim liClientIndex As Long
    Dim liWindowIndex As Long
    Dim liMsgIndex As Long
    Dim liClientTableIndex As Long
    Dim liMask As Long
    Dim iPtr As Long
    
    liMsgIndex = ArrFindInt(miMsgTable, miMsgCount, iMsg)                                   'find the message in the main table
    
    If Not mbCreating Then                                                                  'if not creating a window
        liClientIndex = GetProp(hWnd, ClientIndexPropName)                                  'get the client index prop
        
        If liClientIndex < miClientCount Then                                               'if it is a valid index
            
            liWindowIndex = pFindWindow(liClientIndex, hWnd)                                'try to find the hWnd
            
            If liWindowIndex <> Undefined Then                                              'if the hWnd was found
                If liMsgIndex <> Undefined Then                                             'if the message was found
                    pTranslateMsgIndex liMsgIndex, liClientTableIndex, liMask               'get the client table index
                    iPtr = mtClients(liClientIndex).iPtr                                    'store the objptr
                    With mtClients(liClientIndex).tWindows(liWindowIndex)
                        If Not CBool(.Flags And AllMessages) Then                           'if not receiving all messages
                            If CBool(.Flags And CustomMessages) Then                        'zero pointer if the message is not wanted
                                If Not CBool(.iMsgMask(liClientTableIndex) And liMask) Then iPtr = 0&
                            Else
                                If Not mbDefAllMsgs Then                                    'zero pointer if the message is not wanted
                                    If Not CBool(miDefMsgMask(liClientTableIndex) And liMask) Then iPtr = 0&
                                End If
                            End If
                        End If
                    End With
                Else
                    With mtClients(liClientIndex)
                        iPtr = .iPtr                                                        'store the pointer
                        With .tWindows(liWindowIndex)
                            If Not CBool(.Flags And AllMessages) Then                       'if not getting ALL_MESSAGES
                                If CBool(.Flags And CustomMessages) Then                    'if we're getting custom messages then the
                                    Debug.Assert False                                      'message should have been found
                                    iPtr = 0&
                                Else
                                    If Not mbDefAllMsgs Then                                'if not defaulting to all messages then
                                        Debug.Assert False                                  'message should have been found
                                        iPtr = 0&
                                    End If
                                End If
                            End If
                        End With
                    End With
                End If
            Else
                Debug.Assert False
            End If
        Else
            Debug.Assert False
        End If
    Else                                                                                    'if creating a window
        If mbDefAllMsgs Then                                                                'if defaulting to all messages, store the pointer
            iPtr = miCreatingPtr
        Else
            If liMsgIndex <> Undefined Then                                                 'if not defaulting to all messages
                pTranslateMsgIndex liMsgIndex, liClientTableIndex, liMask                   'get the index to the client table
                If CBool(miDefMsgMask(liClientTableIndex) And liMask) Then iPtr = miCreatingPtr 'if the message is requested, store the pointer
            End If
        End If
    End If
    
    Dim oThis As iWindow

    If iPtr <> 0& And iPtr <> Undefined Then        'If an object was found
        #If bVBVMTypeLib Then                       'If using the VBVM Type lib
            ObjectPtr(oThis) = iPtr                 'Set the object pointer directly
            oThis.Proc bHandled, lReturn, hWnd, _
                       iMsg, wParam, lParam         'call the object's procedure
            ObjectPtr(oThis) = 0&                   'remove the illegal reference
        #Else
            CopyMemory oThis, iPtr, 4&              'use the standard copymem method
            oThis.Proc bHandled, lReturn, hWnd, _
                       iMsg, wParam, lParam         'call the object's procedure
            CopyMemory oThis, 0&, 4&                'remove the illegal reference
        #End If
    Else
        'nobody wants this message
        Debug.Assert False
    End If

End Sub
'</iWindow Interface>


Attribute VB_Name = "mOleControl"
'==================================================================================================
'mOleControl.bas                                8/25/04
'
'           LINEAGE:
'               Paul Wilde's vbACOM.dll from www.vbaccelerator.com
'
'           PURPOSE:
'               Provides VTable subclassing for the IOleControl interface.
'
'           CLASSES CREATED BY THIS MODULE:
'               pcVTableSubclass
'
'==================================================================================================

Option Explicit

Private Enum eVTable
    ' Ignore item 1: QueryInterface
    ' Ignore item 2: AddRef
    ' Ignore item 3: Release
    vtblGetControlInfo = 4
    vtblOnMnemonic
    vtblOnAmbientPropertyChange
    vtblFreezeEvents
    vtblCount
End Enum

Private moSubclass As pcSubclassVTable

Public Sub ReplaceIOleControl(ByVal pObject As vbBaseTlb.IOleControl)
'replace vtable for IOleControl interface

    If moSubclass Is Nothing Then Set moSubclass = New pcSubclassVTable
    
    If moSubclass.RefCount = 0& Then
        moSubclass.Subclass ObjPtr(pObject), vtblCount, vtblGetControlInfo, _
                            AddressOf IOleControl_GetControlInfo, _
                            AddressOf IOleControl_OnMnemonic, _
                            AddressOf IOleControl_OnAmbientPropertyChange, _
                            AddressOf IOleControl_FreezeEvents
        
        'Debug.Print "Replaced vtable methods IOleControl"
        
    End If
    
    moSubclass.AddRef
    
End Sub
Public Sub RestoreIOleControl(ByVal pObject As vbBaseTlb.IOleControl)
'restore vtable for IOleControl interface

    If Not moSubclass Is Nothing Then
    
        moSubclass.Release
        
        If moSubclass.RefCount = 0& Then
            moSubclass.UnSubclass
            'Debug.Print "Restored vtable methods IOleControl"
        End If
        
    End If

End Sub
Private Function IOleControl_OnAmbientPropertyChange(ByVal oThis As Object, ByVal DispID As Long) As Long
'new vtable method for IOleControl::OnAmbientPropertyChange

    Dim oIOleControlVB As iOleControlVB
    Dim bNoDefault As Boolean
    
    'handle all errors so we don't crash caller
    On Error GoTo CATCH_EXCEPTION
    
    'get ref to custom interface
    Set oIOleControlVB = oThis
    
    'call custom implementation of 'OnAmbientPropertyChange'
    oIOleControlVB.OnAmbientPropertyChange bNoDefault, DispID
    
    'if control is not overriding default method
    If Not bNoDefault Then
        'call method from original vtable
        IOleControl_OnAmbientPropertyChange = Original_IOleControl_OnAmbientPropertyChange(oThis, DispID)
        
    Else
        'return 'OK'
        IOleControl_OnAmbientPropertyChange = S_OK
        
    End If
    Exit Function
    
CATCH_EXCEPTION:
    'call method from original vtable
    IOleControl_OnAmbientPropertyChange = Original_IOleControl_OnAmbientPropertyChange(oThis, DispID)
    
End Function
Private Function IOleControl_FreezeEvents(ByVal oThis As Object, ByVal fFreeze As Long) As Long
'new vtable method for IOleControl::FreezeEvents

    Dim oIOleControlVB As iOleControlVB
    Dim bNoDefault As Boolean
    
    'handle all errors so we don't crash caller
    On Error GoTo CATCH_EXCEPTION
    
    'get ref to custom interface
    Set oIOleControlVB = oThis
    
    'call custom implementation of 'FreezeEvents'
    oIOleControlVB.FreezeEvents bNoDefault, IIf(fFreeze = 0, False, True)
    
    'if control is not overriding default method
    If Not bNoDefault Then
        'call method from original vtable
        IOleControl_FreezeEvents = Original_IOleControl_FreezeEvents(oThis, fFreeze)
        
    Else
        'return 'OK'
        IOleControl_FreezeEvents = S_OK
        
    End If
    Exit Function
    
CATCH_EXCEPTION:
    'call method from original vtable
    IOleControl_FreezeEvents = Original_IOleControl_FreezeEvents(oThis, fFreeze)
    
End Function
Private Function IOleControl_OnMnemonic(ByVal oThis As Object, pMsg As tMsg) As Long
'new vtable method for IOleControl::OnMnemonic

    Dim oIOleControlVB As iOleControlVB
    Dim nShift As Integer
    Dim bNoDefault As Boolean
    
    'handle all errors so we don't crash caller
    On Error GoTo CATCH_EXCEPTION
    
    'validate params
    If VarPtr(pMsg) = 0& Then
        IOleControl_OnMnemonic = E_POINTER
        Exit Function
        
    End If
    
    'get ref to custom interface
    Set oIOleControlVB = oThis
    
    'get status of modifier keys
    nShift = GetKeyModifiers()
    
    'call custom implementation of 'OnMnemonic'
    oIOleControlVB.OnMnemonic bNoDefault, pMsg, nShift
    
    'if control is not overriding default method
    If Not bNoDefault Then
        'call method from original vtable
        IOleControl_OnMnemonic = Original_IOleControl_OnMnemonic(oThis, pMsg)
        
    Else
        'return 'OK'
        IOleControl_OnMnemonic = S_OK
        
    End If
    Exit Function
    
CATCH_EXCEPTION:
    'call method from original vtable
    IOleControl_OnMnemonic = Original_IOleControl_OnMnemonic(oThis, pMsg)
    
End Function
Private Function IOleControl_GetControlInfo(ByVal oThis As Object, pCI As CONTROLINFO) As Long
'new vtable method for IOleControl::GetControlInfo

    Dim oIOleControlVB As iOleControlVB
    Dim bNoDefault As Boolean
    
    'handle all errors so we don't crash caller
    On Error GoTo CATCH_EXCEPTION
    
    'validate params
    If VarPtr(pCI) = 0& Then
        IOleControl_GetControlInfo = E_POINTER
        Exit Function
        
    End If
    
    'get ref to custom interface
    Set oIOleControlVB = oThis
    
    'call custom implementation of 'GetControlInfo'
    pCI.cb = LenB(pCI)
    
    oIOleControlVB.GetControlInfo bNoDefault, pCI.cAccel, pCI.hAccel, pCI.dwFlags
    
    'if control is not overriding default method
    If Not bNoDefault Then
        'call method from original vtable
        IOleControl_GetControlInfo = Original_IOleControl_GetControlInfo(oThis, pCI)
        
    Else
        
        'if array contains items but mem handle is 0 then problem
        If pCI.cAccel > 0& And pCI.hAccel = 0& Then
            'return 'out of memory' error
            IOleControl_GetControlInfo = E_OUTOFMEMORY
            
        Else
            'return 'OK'
            IOleControl_GetControlInfo = S_OK
            
        End If
    End If
    Exit Function
    
CATCH_EXCEPTION:
    'call method from original vtable
    IOleControl_GetControlInfo = Original_IOleControl_GetControlInfo(oThis, pCI)
    
End Function
Private Function Original_IOleControl_OnAmbientPropertyChange(ByVal oThis As vbBaseTlb.IOleControl, ByVal DispID As Long) As Long
'call original 'OnAmbientPropertyChange' method
   
    'temporarily unhook method so we can call the original
    moSubclass.SubclassEntry(vtblOnAmbientPropertyChange) = False
    
    'call the original method
    Original_IOleControl_OnAmbientPropertyChange = oThis.OnAmbientPropertyChange(DispID)
    
    're-hook the method
    moSubclass.SubclassEntry(vtblOnAmbientPropertyChange) = True
    
End Function
Private Function Original_IOleControl_FreezeEvents(ByVal oThis As vbBaseTlb.IOleControl, ByVal fFreeze As Long) As Long
'call original 'FreezeEvents' method
    
    'temporarily unhook method so we can call the original
    moSubclass.SubclassEntry(vtblFreezeEvents) = False

    Original_IOleControl_FreezeEvents = oThis.FreezeEvents(fFreeze)
    
    're-hook the method
    moSubclass.SubclassEntry(vtblFreezeEvents) = True
    
End Function
Private Function Original_IOleControl_GetControlInfo(ByVal oThis As vbBaseTlb.IOleControl, pCI As CONTROLINFO) As Long
'call original 'GetControlInfo' method
    
    'temporarily unhook method so we can call the original
    moSubclass.SubclassEntry(vtblGetControlInfo) = False
    
    Original_IOleControl_GetControlInfo = oThis.GetControlInfo(pCI)
       
    're-hook the method
    moSubclass.SubclassEntry(vtblGetControlInfo) = True
End Function
Private Function Original_IOleControl_OnMnemonic(ByVal oThis As vbBaseTlb.IOleControl, pMsg As tMsg) As Long
'call original 'OnMnemonic' method
    
    'temporarily unhook method so we can call the original
    moSubclass.SubclassEntry(vtblOnMnemonic) = False

    Original_IOleControl_OnMnemonic = oThis.OnMnemonic(ByVal VarPtr(pMsg))

    're-hook the method
    moSubclass.SubclassEntry(vtblOnMnemonic) = True
End Function

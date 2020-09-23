Attribute VB_Name = "mOleInPlaceActiveObject"
'==================================================================================================
'mOleInPlaceActiveObject.bas            8/25/04
'
'           LINEAGE:
'               Based on modIOleInPlaceActiveObject.bas in vbACOM.dll from vbaccelerator.com
'
'           PURPOSE:
'               Subclassed inplementation of IOleInPlaceActiveObject
'
'           MODULES CALLED FROM THIS MODULE:
'               NONE
'
'           CLASSES CREATED BY THIS MODULE:
'               pcSubclassVTable
'
'==================================================================================================


Option Explicit

Private Enum eVTable
    ' Ignore item 1: QueryInterface
    ' Ignore item 2: AddRef
    ' Ignore item 3: Release
    ' Ignore item 4: GetWindow
    ' Ignore item 5: ContextSensitiveHelp
    vtblTranslateAccelerator = 6
    vtblOnFrameWindowActivate
    vtblOnDocWindowActivate
    vtblResizeBorder
    vtblEnableModeless
    vtblCount
End Enum

Private moSubclass As pcSubclassVTable

Public Function IOleInPlaceActiveObject_OnDocWindowActivate(ByVal oThis As Object, ByVal fActive As Long) As Long
'new vtable method for IOleInPlaceActiveObject::OnDocWindowActivate

    Dim oIOleInPlaceActiveObjectVB As iOleInPlaceActiveObjectVB
    Dim lbHandled As Boolean
    
    'handle all errors so we don't crash caller
    On Error GoTo CATCH_EXCEPTION
    
    'get ref to custom interface
    Set oIOleInPlaceActiveObjectVB = oThis
    
    'call custom implementation of 'OnDocWindowActivate'
    oIOleInPlaceActiveObjectVB.OnDocWindowActivate lbHandled, IIf(fActive = 0, False, True)
    
    'if control is not overriding default method
    If Not lbHandled Then
        'call method from original vtable
        IOleInPlaceActiveObject_OnDocWindowActivate = Original_IOleInPlaceActiveObject_OnDocWindowActivate(oThis, fActive)
        
    Else
        IOleInPlaceActiveObject_OnDocWindowActivate = S_OK
        
    End If
    
    Exit Function
    
CATCH_EXCEPTION:
    'call method from original vtable
    IOleInPlaceActiveObject_OnDocWindowActivate = Original_IOleInPlaceActiveObject_OnDocWindowActivate(oThis, fActive)
    
End Function
Public Function IOleInPlaceActiveObject_EnableModeless(ByVal oThis As Object, ByVal fActive As Long) As Long
'new vtable method for IOleInPlaceActiveObject::EnableModeless

    Dim oIOleInPlaceActiveObjectVB As iOleInPlaceActiveObjectVB
    Dim lbHandled As Boolean
    
    'handle all errors so we don't crash caller
    On Error GoTo CATCH_EXCEPTION
    
    'get ref to custom interface
    Set oIOleInPlaceActiveObjectVB = oThis
    
    'call custom implementation of 'EnableModeless'
    oIOleInPlaceActiveObjectVB.EnableModeless lbHandled, IIf(fActive = 0, False, True)
    
    'if control is not overriding default method
    If Not lbHandled Then
        'call method from original vtable
        IOleInPlaceActiveObject_EnableModeless = Original_IOleInPlaceActiveObject_EnableModeless(oThis, fActive)
    Else
        IOleInPlaceActiveObject_EnableModeless = S_OK
    End If
    
    Exit Function
    
CATCH_EXCEPTION:
    'call method from original vtable
    IOleInPlaceActiveObject_EnableModeless = Original_IOleInPlaceActiveObject_EnableModeless(oThis, fActive)
    
End Function
Public Function IOleInPlaceActiveObject_ResizeBorder(ByVal oThis As Object, prcBorder As tRect, ByVal oUIWindow As vbBaseTlb.IOleInPlaceUIWindow, ByVal fFrameWindow As Long) As Long
'new vtable method for IOleInPlaceActiveObject::ResizeBorder

    Dim oIOleInPlaceActiveObjectVB As iOleInPlaceActiveObjectVB
    Dim lbHandled As Boolean
   
    'handle all errors so we don't crash caller
    On Error GoTo CATCH_EXCEPTION
    
    'get ref to custom interface
    Set oIOleInPlaceActiveObjectVB = oThis
    
    'call custom implementation of 'ResizeBorder'
    oIOleInPlaceActiveObjectVB.ResizeBorder lbHandled, prcBorder, oUIWindow, IIf(fFrameWindow = 0, False, True)
    
    'if control is not overriding default method
    If Not lbHandled Then
        'call method from original vtable
        IOleInPlaceActiveObject_ResizeBorder = Original_IOleInPlaceActiveObject_ResizeBorder(oThis, prcBorder, oUIWindow, fFrameWindow)
        
    Else
        IOleInPlaceActiveObject_ResizeBorder = S_OK
        
    End If
    Exit Function
    
CATCH_EXCEPTION:
    'call method from original vtable
    IOleInPlaceActiveObject_ResizeBorder = Original_IOleInPlaceActiveObject_ResizeBorder(oThis, prcBorder, oUIWindow, fFrameWindow)
    
End Function
Public Function IOleInPlaceActiveObject_OnFrameWindowActivate(ByVal oThis As Object, ByVal fActive As Long) As Long
'new vtable method for IOleInPlaceActiveObject::OnFrameWindowActivate

    Dim oIOleInPlaceActiveObjectVB As iOleInPlaceActiveObjectVB
    Dim lbHandled As Boolean
    
    'handle all errors so we don't crash caller
    On Error GoTo CATCH_EXCEPTION
    
    'get ref to custom interface
    Set oIOleInPlaceActiveObjectVB = oThis

    'call custom implementation of 'OnFrameWindowActivate'
    oIOleInPlaceActiveObjectVB.OnFrameWindowActivate lbHandled, IIf(fActive = 0, False, True)
    
    'if control is not overriding default method
    If Not lbHandled Then
        'call method from original vtable
        IOleInPlaceActiveObject_OnFrameWindowActivate = Original_IOleInPlaceActiveObject_OnFrameWindowActivate(oThis, fActive)
    Else
        IOleInPlaceActiveObject_OnFrameWindowActivate = S_OK
    End If
    Exit Function
    
CATCH_EXCEPTION:
    'call method from original vtable
    IOleInPlaceActiveObject_OnFrameWindowActivate = Original_IOleInPlaceActiveObject_OnFrameWindowActivate(oThis, fActive)
    
End Function
Public Function IOleInPlaceActiveObject_TranslateAccelerator(ByVal oThis As Object, pMsg As tMsg) As Long
'new vtable method for IOleInPlaceActiveObject::TranslateAccelerator
    
    Dim oIOleInPlaceActiveObjectVB As iOleInPlaceActiveObjectVB
    Dim lbHandled As Boolean
    
    'handle all errors so we don't crash caller
    On Error GoTo CATCH_EXCEPTION
    
    'validate params
    If VarPtr(pMsg) = 0& Then
        IOleInPlaceActiveObject_TranslateAccelerator = E_POINTER
        Exit Function
        
    End If
    
    'get ref to custom interface
    Set oIOleInPlaceActiveObjectVB = oThis
    
    'default to S_OK
    IOleInPlaceActiveObject_TranslateAccelerator = S_OK
    
    'call custom implementation of 'TranslateAccelerator'
    oIOleInPlaceActiveObjectVB.TranslateAccelerator lbHandled, IOleInPlaceActiveObject_TranslateAccelerator, pMsg, GetKeyModifiers()
    
    'if control is not overriding default method
    If Not lbHandled Then
        'call method from original vtable
        IOleInPlaceActiveObject_TranslateAccelerator = Original_IOleInPlaceActiveObject_TranslateAccelerator(oThis, pMsg)
        
    End If
        
    Exit Function
    
CATCH_EXCEPTION:
    'call method from original vtable
    IOleInPlaceActiveObject_TranslateAccelerator = Original_IOleInPlaceActiveObject_TranslateAccelerator(oThis, pMsg)
    
End Function
Private Function Original_IOleInPlaceActiveObject_TranslateAccelerator(ByVal oThis As vbBaseTlb.IOleInPlaceActiveObject, pMsg As tMsg) As Long
'call original 'TranslateAccelerator' method
    
    'temporarily unhook method so we can call the original
    moSubclass.SubclassEntry(vtblTranslateAccelerator) = False
    
    'call the original method
    Original_IOleInPlaceActiveObject_TranslateAccelerator = oThis.TranslateAccelerator(ByVal VarPtr(pMsg))

    're-hook the method
    moSubclass.SubclassEntry(vtblTranslateAccelerator) = True
    
End Function
Private Function Original_IOleInPlaceActiveObject_OnDocWindowActivate(ByVal oThis As IOleInPlaceActiveObject, ByVal fActive As Long) As Long
'call original 'OnDocWindowActivate' method
    
    'temporarily unhook method so we can call the original
    moSubclass.SubclassEntry(vtblOnDocWindowActivate) = False

    'call the original method
    Original_IOleInPlaceActiveObject_OnDocWindowActivate = oThis.OnDocWindowActivate(fActive)
    
    're-hook the method
    moSubclass.SubclassEntry(vtblOnDocWindowActivate) = True

End Function
Private Function Original_IOleInPlaceActiveObject_EnableModeless(ByVal oThis As IOleInPlaceActiveObject, ByVal fActive As Long) As Long
'call original 'EnableModeless' method
    
    'temporarily unhook method so we can call the original
    moSubclass.SubclassEntry(vtblEnableModeless) = False
    
    'call the original method
    Original_IOleInPlaceActiveObject_EnableModeless = oThis.EnableModeless(fActive)
    
    're-hook the method
    moSubclass.SubclassEntry(vtblEnableModeless) = True
End Function
Private Function Original_IOleInPlaceActiveObject_ResizeBorder(ByVal oThis As IOleInPlaceActiveObject, prcBorder As tRect, ByVal oUIWindow As vbBaseTlb.IOleInPlaceUIWindow, ByVal fFrameWindow As Long) As Long
'call original 'ResizeBorder' method
    
    'temporarily unhook method so we can call the original
    moSubclass.SubclassEntry(vtblResizeBorder) = False

    'call the original method
    Original_IOleInPlaceActiveObject_ResizeBorder = oThis.ResizeBorder(ByVal VarPtr(prcBorder), oUIWindow, fFrameWindow)
   
    're-hook the method
    moSubclass.SubclassEntry(vtblResizeBorder) = True
    
End Function
Private Function Original_IOleInPlaceActiveObject_OnFrameWindowActivate(ByVal oThis As IOleInPlaceActiveObject, ByVal fActive As Long) As Long
'call original 'OnFrameWindowActivate' method
    
    'temporarily unhook method so we can call the original
    moSubclass.SubclassEntry(vtblOnFrameWindowActivate) = False

    'call the original method
    Original_IOleInPlaceActiveObject_OnFrameWindowActivate = oThis.OnFrameWindowActivate(fActive)
    
    're-hook the method
    moSubclass.SubclassEntry(vtblOnFrameWindowActivate) = True
    
End Function
Public Sub ReplaceIOleInPlaceActiveObject(ByVal oThis As vbBaseTlb.IOleInPlaceActiveObject)
'replace vtable for IOleInPlaceActiveObject interface

    If moSubclass Is Nothing Then Set moSubclass = New pcSubclassVTable
    
    If moSubclass.RefCount = 0 Then
        moSubclass.Subclass ObjPtr(oThis), vtblCount, vtblTranslateAccelerator, _
                                AddressOf IOleInPlaceActiveObject_TranslateAccelerator, _
                                AddressOf IOleInPlaceActiveObject_OnFrameWindowActivate, _
                                AddressOf IOleInPlaceActiveObject_OnDocWindowActivate, _
                                AddressOf IOleInPlaceActiveObject_ResizeBorder, _
                                AddressOf IOleInPlaceActiveObject_EnableModeless
        'Debug.Print "Replaced vtable methods IOleInPlaceActiveObject"
    End If
    
    moSubclass.AddRef
End Sub
Public Sub RestoreIOleInPlaceActiveObject(ByVal pObject As vbBaseTlb.IOleInPlaceActiveObject)
'restore vtable for IOleInPlaceActiveObject interface

    If Not moSubclass Is Nothing Then

        moSubclass.Release
        
        If moSubclass.RefCount = 0& Then
            moSubclass.UnSubclass
            'Debug.Print "Restored vtable methods IOleInPlaceActiveObject"
        End If
        
    End If
    
End Sub

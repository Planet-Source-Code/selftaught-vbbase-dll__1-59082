Attribute VB_Name = "mVTableSubclass"
'==================================================================================================
'mVTableSubclass.bas            8/25/04
'
'           LINEAGE:
'               vbACOM.dll from vbaccelerator.com
'
'           PURPOSE:
'               Manages VTable subclassing for all the interfaces subclassed by this component other
'               than IEnumVARIANT.
'
'           MODULES CALLED FROM THIS MODULE:
'               mOleInPlaceActiveObject.bas
'               mPerPropertyBrowsing.bas
'               mOleControl.bas
'
'           CLASSES CREATED BY THIS MODULE:
'               pcSubclassVTable
'
'==================================================================================================

Option Explicit

#Const bVBVMTypeLib = False

''accelerator flags (used with ACCEL structure)
'Public Const FVIRTKEY As Long = 1 '/* Assumed to be as long =as long = TRUE */
'Public Const FNOINVERT As Long = &H2&
'Public Const FSHIFT As Long = &H4&
'Public Const FCONTROL As Long = &H8&
'Public Const FALT As Long = &H10&

Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer

Private Enum eInterfaces
    intOleControl
    intOleIPAO
    intOlePerPropertyBrowsing
End Enum

Public Function GetKeyModifiers() As Integer
'get pressed status of [SHIFT],[CONTROL], and [ALT] keys

    Dim nResult As Integer
    
    nResult = nResult Or (-1 * KeyIsPressed(vbKeyShift))
    nResult = nResult Or (-2 * KeyIsPressed(vbKeyMenu))
    nResult = nResult Or (-4 * KeyIsPressed(vbKeyControl))
    GetKeyModifiers = nResult
End Function

Private Function KeyIsPressed(ByVal VirtKeyCode As KeyCodeConstants) As Boolean
'poll windows to see if specified key is pressed

    Dim lngResult As Long
    
    lngResult = GetAsyncKeyState(VirtKeyCode)
    If (lngResult And &H8000&) = &H8000& Then
        KeyIsPressed = True
            
    End If
End Function

Public Function Attach(ByVal oObject As Object) As Boolean
    'Attach vTable Subclassing
    
    If pSupports(oObject, intOlePerPropertyBrowsing) Then
        ReplaceIPerPropertyBrowsing oObject
        Attach = True
    End If

    If pSupports(oObject, intOleControl) Then
        ReplaceIOleControl oObject
        Attach = True
    End If

    If pSupports(oObject, intOleIPAO) Then
        ReplaceIOleInPlaceActiveObject oObject
        Attach = True
    End If
    
End Function

Public Function Detach(ByVal oObject As Object) As Boolean
    'this function must be called on a 1:1 basis with Attach
    
    If pSupports(oObject, intOlePerPropertyBrowsing) Then
        RestoreIPerPropertyBrowsing oObject
        Detach = True
    End If
    
    If pSupports(oObject, intOleControl) Then
        RestoreIOleControl oObject
        Detach = True
    End If
    
    If pSupports(oObject, intOleIPAO) Then
        RestoreIOleInPlaceActiveObject oObject
        Detach = True
    End If
    
End Function

Private Function pSupports(oUserControl As Object, iInterface As eInterfaces) As Boolean

    'determine whether the object supports both the ole interface and the subclass interface
    'for any of the ole interfacing that we are interested in.

    On Error GoTo handler

    Select Case iInterface
        Case intOlePerPropertyBrowsing
            
            Dim oIPerPropertyBrowsing As vbBaseTlb.IPerPropertyBrowsing
            Dim oIPerPropertyBrowsingVB As iPerPropertyBrowsingVB
            
            Set oIPerPropertyBrowsing = oUserControl: Set oIPerPropertyBrowsingVB = oUserControl
            pSupports = Not (oIPerPropertyBrowsing Is Nothing Or oIPerPropertyBrowsingVB Is Nothing)
            
        Case intOleIPAO
            
            Dim oIOleInPlaceActiveObject As vbBaseTlb.IOleInPlaceActiveObject
            Dim oIOleInPlaceActiveObjectVB As iOleInPlaceActiveObjectVB
            
            Set oIOleInPlaceActiveObject = oUserControl: Set oIOleInPlaceActiveObjectVB = oUserControl
            pSupports = Not (oIOleInPlaceActiveObject Is Nothing Or oIOleInPlaceActiveObjectVB Is Nothing)
            
        Case intOleControl
        
            Dim oIOleControl As vbBaseTlb.IOleControl
            Dim oIOleControlVB As iOleControlVB
        
            Set oIOleControl = oUserControl: Set oIOleControlVB = oUserControl
            pSupports = Not (oIOleControl Is Nothing Or oIOleControlVB Is Nothing)
            
    End Select

    Exit Function
    
handler:
    
End Function

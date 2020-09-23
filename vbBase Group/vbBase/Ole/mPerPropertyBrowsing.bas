Attribute VB_Name = "mPerPropertyBrowsing"
'==================================================================================================
'mPerPropertyBrowsing.bas               8/25/04
'
'           LINEAGE:
'               Based on modIPerPropertyBrowsing.bas in vbACOM.dll from vbaccelerator.com
'
'           PURPOSE:
'               Subclassed implementation of IPerPropertyBrowsing
'
'           MODULES CALLED FROM THIS MODULE:
'               NONE
'
'           CLASSES CREATED BY THIS MODULE:
'               pcSubclassVTable
'
'==================================================================================================

Option Explicit

'registry key flags
Private Const HKEY_CLASSES_ROOT As Long = &H80000000
Private Const ERROR_SUCCESS As Long = 0&
Private Const ERROR_NO_MORE_ITEMS As Long = 259&
Private Const REG_SZ As Long = 1

'OLEAUT32
Private Declare Function SysAllocString Lib "oleaut32.dll" (ByVal lpString As Long) As Long

'OLE32
Private Declare Function IIDFromString Lib "ole32.dll" (ByVal lpsz As String, lpGuid As CLSID) As Long
Private Declare Function CoTaskMemAlloc Lib "ole32.dll" (ByVal cBytes As Long) As Long

'ADVAPI32
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpszSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpszValueName As Any, ByVal dwReserved As Long, lpdwType As Long, lpbData As Any, cbData As Long) As Long

'private members
Private moSubclass As pcSubclassVTable
Private mbStringsNotImpl As Boolean

Private Enum eVTable
    ' Ignore item 1: QueryInterface
    ' Ignore item 2: AddRef
    ' Ignore item 3: Release
    vtblGetDisplayString = 4
    vtblMapPropertyToPage
    vtblGetPredefinedStrings
    vtblGetPredefinedValue
    vtblCount
End Enum

Private Function FindGUIDForProgID(ByVal ProgID As String) As String
'find specified prog ID in registry & return GUID

    Dim hKey As Long
    Dim strCLSID As String
    Dim lngNullPos As Long
    Dim lngValueType As Long, lngStrLen As Long
    
    'open ProgID\CLSID registry key
    If RegOpenKey(HKEY_CLASSES_ROOT, ProgID & "\CLSID", hKey) <> ERROR_SUCCESS Then
        'attempt to open new version key for progid
        If RegOpenKey(HKEY_CLASSES_ROOT, ProgID & "\CurVer", hKey) <> ERROR_SUCCESS Then
            Exit Function
            
        Else
            'get ProgID string from key
            'get data type & size
            If RegQueryValueEx(hKey, 0&, 0&, lngValueType, ByVal 0&, lngStrLen) = ERROR_SUCCESS Then
                'if data type is string & size is > 0
                If lngValueType = REG_SZ And lngStrLen > 0 Then
                    ProgID = String$(lngStrLen, vbNullChar)
                    'get default value
                    If RegQueryValueEx(hKey, 0&, 0&, 0&, ByVal ProgID, lngStrLen) = ERROR_SUCCESS Then
                        'strip null terminator
                        lngNullPos = InStr(ProgID, vbNullChar)
                        If lngNullPos > 0 Then
                            ProgID = Left$(ProgID, lngNullPos - 1)
                            
                        End If
                        
                    End If
                
                End If
            
            End If
            'close ProgID\CurVer registry key
            RegCloseKey hKey
            'open ProgID\CLSID registry key
            If RegOpenKey(HKEY_CLASSES_ROOT, ProgID & "\CLSID", hKey) <> ERROR_SUCCESS Then
                Exit Function
                
            End If
            
        End If
    
    End If
    
    'get CLSID string from key
    'get data type & size
    If RegQueryValueEx(hKey, 0&, 0&, lngValueType, ByVal 0&, lngStrLen) = ERROR_SUCCESS Then
        'if data type is string & size is > 0
        If lngValueType = REG_SZ And lngStrLen > 0 Then
            strCLSID = String$(lngStrLen, vbNullChar)
            'get default value
            If RegQueryValueEx(hKey, 0&, 0&, 0&, ByVal strCLSID, lngStrLen) = ERROR_SUCCESS Then
                'strip null terminator
                lngNullPos = InStr(strCLSID, vbNullChar)
                If lngNullPos > 0 Then
                    strCLSID = Left$(strCLSID, lngNullPos - 1)
                    
                End If
                'store CLSID
                FindGUIDForProgID = strCLSID
                
            End If
        
        End If
    
    End If
    
    'close ProgID\CLSID registry key
    RegCloseKey hKey
End Function

Private Function GUIDFromString(ByVal Guid As String) As CLSID
'convert string GUID to *real* GUID
    
    Dim lpGuid As CLSID
    
    'convert string to unicode
    Guid = StrConv(Guid, vbUnicode)
    'convert string to GUID
    IIDFromString Guid, lpGuid
    'return *real* GUID
    GUIDFromString = lpGuid
End Function


Private Function IPerPropertyBrowsing_GetDisplayString(ByVal oThis As Object, ByVal DispID As Long, ByVal lpDisplayName As Long) As Long
'new vtable method for IPerPropertyBrowsing::GetDisplayString
    
    Dim oIPerPropertyBrowsingVB As iPerPropertyBrowsingVB
    Dim bNoDefault As Boolean
    Dim strDisplayName As String
    Dim lpString As Long
    
    'handle all errors so we don't crash caller
    On Error GoTo CATCH_EXCEPTION
    
    'Debug.Print "GetDisplayString"
    
    'validate passed pointer
    If VarPtr(lpDisplayName) = 0 Then
        IPerPropertyBrowsing_GetDisplayString = E_POINTER
        Exit Function
        
    End If
    
    'cast method to source interface
    Set oIPerPropertyBrowsingVB = oThis
    oIPerPropertyBrowsingVB.GetDisplayString bNoDefault, DispID, strDisplayName
        
    'if no param set by user
    If bNoDefault Then

        'copy display string to passed ptr (caller should free the memory allocated)
        lpString = SysAllocString(StrPtr(strDisplayName))
        
        CopyMemory ByVal lpDisplayName, lpString, 4
        
    Else
CATCH_EXCEPTION:
        
        IPerPropertyBrowsing_GetDisplayString = Original_IPerPropertyBrowsing_GetDisplayString(oThis, DispID, lpDisplayName)
        
    End If

End Function
Private Function IPerPropertyBrowsing_MapPropertyToPage(ByVal oThis As Object, ByVal DispID As Long, lpCLSID As CLSID) As Long
'new vtable method for IPerPropertyBrowsing::MapPropertyToPage

    'Debug.Print "MapPropertyPage"

    Dim oIPerPropertyBrowsingVB As iPerPropertyBrowsingVB
    Dim bNoDefault As Boolean
    Dim strGUID As String
    
    'handle all errors so we don't crash caller
    On Error GoTo CATCH_EXCEPTION
    
    'validate passed pointer
    If VarPtr(lpCLSID) = 0 Then
        IPerPropertyBrowsing_MapPropertyToPage = E_POINTER
        Exit Function
        
    End If
    
    'cast method to source interface
    Set oIPerPropertyBrowsingVB = oThis
    oIPerPropertyBrowsingVB.MapPropertyToPage bNoDefault, DispID, strGUID
        
    'if no param set by user
    If bNoDefault Then
        'if valid string
        If Len(strGUID) > 2 Then
            'if not a GUID
            If Not (Left$(strGUID, 1) = "{" And Right$(strGUID, 1) = "}") Then
                'get CLSID from ProgID
                strGUID = FindGUIDForProgID(strGUID)
                
            End If
            'convert string CLSID to *real* CLSID
            lpCLSID = GUIDFromString(strGUID)
            
        End If
    Else
    
CATCH_EXCEPTION:
    
        IPerPropertyBrowsing_MapPropertyToPage = Original_IPerPropertyBrowsing_MapPropertyToPage(oThis, DispID, lpCLSID)
    
    End If

    
End Function
Private Function IPerPropertyBrowsing_GetPredefinedStrings(ByVal oThis As Object, ByVal DispID As Long, pCaStringsOut As CALPOLESTR, pCaCookiesOut As CADWORD) As Long
'new vtable method for IPerPropertyBrowsing::GetPredefinedStrings

    Dim oIPerPropertyBrowsingVB As iPerPropertyBrowsingVB
    Dim bNoDefault As Boolean
    
    Dim cElems As Long
    Dim pElems As Long
    Dim lpString As Long
    
    'Debug.Print "GetPredefinedStrings"
    
    'handle all errors so we don't crash caller
    On Error GoTo CATCH_EXCEPTION
    mbStringsNotImpl = False
    
    'validate passed pointers
    If VarPtr(pCaStringsOut) = 0 Or VarPtr(pCaCookiesOut) = 0 Then
        IPerPropertyBrowsing_GetPredefinedStrings = E_POINTER
        Exit Function
        
    End If
    
    'create & initialise cPropertyListItems collection
    Dim loProps As cPropertyListItems
    Set loProps = New cPropertyListItems
    
    'cast method to source interface
    Set oIPerPropertyBrowsingVB = oThis
    oIPerPropertyBrowsingVB.GetPredefinedStrings bNoDefault, DispID, loProps
    
    'if no param set by user
    If bNoDefault And loProps.Count > 0& Then
        'initialise CALPOLESTR struct
        cElems = loProps.Count
        pElems = CoTaskMemAlloc(cElems * 4)
        
        pCaStringsOut.cElems = cElems
        pCaStringsOut.pElems = pElems
        
        
        Dim lsTemp As String
        Dim i As Long
        
        For i = 1 To loProps.Count
            lpString = loProps.Item(i).lpDisplayName
            CopyMemory ByVal pElems, lpString, 4&
            'incr the element count
            pElems = UnsignedAdd(pElems, 4&)
        Next
        
        
        'initialise CADWORD struct
        pElems = CoTaskMemAlloc(cElems * 4)
        pCaCookiesOut.cElems = cElems
        pCaCookiesOut.pElems = pElems
        
        'copy dwords to CADWORD struct
        For i = 1 To loProps.Count
            CopyMemory ByVal pElems, loProps(i).Cookie, 4
            pElems = UnsignedAdd(pElems, 4&)
        Next
        
    Else

CATCH_EXCEPTION:
        
        IPerPropertyBrowsing_GetPredefinedStrings = Original_IPerPropertyBrowsing_GetPredefinedStrings(oThis, DispID, pCaStringsOut, pCaCookiesOut)
        mbStringsNotImpl = True
    End If

    
End Function
Private Function IPerPropertyBrowsing_GetPredefinedValue(ByVal oThis As Object, ByVal DispID As Long, ByVal dwCookie As Long, pVarOut As Variant) As Long
'new vtable method for IPerPropertyBrowsing::GetPredefinedValue

    'Debug.Print "GetPredefinedValue"
    
    If mbStringsNotImpl Then
        Debug.Print "Strings Not Implemented"
        IPerPropertyBrowsing_GetPredefinedValue = E_NOTIMPL
        Exit Function
    End If
    
    Dim oIPerPropertyBrowsingVB As iPerPropertyBrowsingVB
    Dim bNoDefault As Boolean
    
    'handle all errors so we don't crash caller
    On Error GoTo CATCH_EXCEPTION

    'validate passed pointers
    If VarPtr(dwCookie) = 0 Or VarPtr(pVarOut) = 0 Then
        IPerPropertyBrowsing_GetPredefinedValue = E_POINTER
        Exit Function
        
    End If
    
    'cast method to source interface
    Set oIPerPropertyBrowsingVB = oThis
    oIPerPropertyBrowsingVB.GetPredefinedValue bNoDefault, DispID, dwCookie, pVarOut
    
    'if no param set by user
    If bNoDefault Then
        
        IPerPropertyBrowsing_GetPredefinedValue = S_OK
        
    Else
    
CATCH_EXCEPTION:
        
        IPerPropertyBrowsing_GetPredefinedValue = Original_IPerPropertyBrowsing_GetPredefinedValue(oThis, DispID, dwCookie, pVarOut)
        
    End If
    
End Function

Private Function Original_IPerPropertyBrowsing_GetDisplayString(ByVal oThis As IPerPropertyBrowsing, ByVal DispID As Long, ByVal lpDisplayName As Long) As Long
    Original_IPerPropertyBrowsing_GetDisplayString = E_NOTIMPL
    'Exit Function
    
    'moSubclass.SubclassEntry(vtblGetDisplayString) = False
    'Dim ls As String
    'CopyMemory ls, lpDisplayName, 4&
    'Original_IPerPropertyBrowsing_GetDisplayString = oThis.GetDisplayString(DispID, ls)
    'moSubclass.SubclassEntry(vtblGetDisplayString) = True
End Function
Private Function Original_IPerPropertyBrowsing_MapPropertyToPage(ByVal oThis As IPerPropertyBrowsing, ByVal DispID As Long, lpCLSID As CLSID) As Long
    Original_IPerPropertyBrowsing_MapPropertyToPage = E_NOTIMPL
    'Exit Function
    
    'moSubclass.SubclassEntry(vtblMapPropertyToPage) = False
    'Original_IPerPropertyBrowsing_MapPropertyToPage = oThis.MapPropertyToPage(DispID, lpCLSID)
    'moSubclass.SubclassEntry(vtblMapPropertyToPage) = True
End Function
Private Function Original_IPerPropertyBrowsing_GetPredefinedStrings(ByVal oThis As IPerPropertyBrowsing, ByVal DispID As Long, pCaStringsOut As CALPOLESTR, pCaCookiesOut As CADWORD) As Long
    Original_IPerPropertyBrowsing_GetPredefinedStrings = E_NOTIMPL
    'Exit Function
    
    'moSubclass.SubclassEntry(vtblGetPredefinedStrings) = False
    'Original_IPerPropertyBrowsing_GetPredefinedStrings = oThis.GetPredefinedStrings(DispID, pCaStringsOut, pCaCookiesOut)
    'moSubclass.SubclassEntry(vtblGetPredefinedStrings) = True
End Function
Private Function Original_IPerPropertyBrowsing_GetPredefinedValue(ByVal oThis As IPerPropertyBrowsing, ByVal DispID As Long, ByVal dwCookie As Long, pVarOut As Variant) As Long
    If VarPtr(pVarOut) Then pVarOut = Empty
    Original_IPerPropertyBrowsing_GetPredefinedValue = E_NOTIMPL
    'Exit Function
    
    'moSubclass.SubclassEntry(vtblGetPredefinedValue) = False
    'Original_IPerPropertyBrowsing_GetPredefinedValue = oThis.GetPredefinedValue(DispID, dwCookie, pVarOut)
    'moSubclass.SubclassEntry(vtblGetPredefinedValue) = True
End Function

Public Sub ReplaceIPerPropertyBrowsing(ByVal pObject As vbBaseTlb.IPerPropertyBrowsing)
'replace vtable for IPerPropertyBrowsing interface

    If moSubclass Is Nothing Then Set moSubclass = New pcSubclassVTable
    
    If moSubclass.RefCount = 0& Then
        
        moSubclass.Subclass ObjPtr(pObject), vtblCount, vtblGetDisplayString, _
                            AddressOf IPerPropertyBrowsing_GetDisplayString, _
                            AddressOf IPerPropertyBrowsing_MapPropertyToPage, _
                            AddressOf IPerPropertyBrowsing_GetPredefinedStrings, _
                            AddressOf IPerPropertyBrowsing_GetPredefinedValue
    
        'Debug.Print "Replaced vtable methods IPerPropertyBrowsing"
    
    End If
    
    moSubclass.AddRef

End Sub
Public Sub RestoreIPerPropertyBrowsing(ByVal pObject As vbBaseTlb.IPerPropertyBrowsing)
'restore vtable for IPerPropertyBrowsing interface
        
    If Not moSubclass Is Nothing Then
        moSubclass.Release
        
        If moSubclass.RefCount = 0& Then
            moSubclass.UnSubclass
            'Debug.Print "Restored vtable methods IPerPropertyBrowsing"
        End If
    End If
End Sub

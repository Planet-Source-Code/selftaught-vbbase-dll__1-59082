Attribute VB_Name = "mEnumVariant"
'==================================================================================================
'mOleControl.bas                                8/25/04
'
'           LINEAGE:
'               Paul Wilde's vbACOM.dll from www.vbaccelerator.com
'
'           PURPOSE:
'               Provides VTable subclassing for the IEnumVARIANT interface.
'
'           CLASSES CREATED BY THIS MODULE:
'               pcVTableSubclass
'
'==================================================================================================

Option Explicit

Private Declare Function ArrPtr Lib "msvbvm60.dll" Alias "VarPtr" (Var() As Any) As Long

Private Type SafeArray1D
  cDims       As Integer
  fFeatures   As Integer
  cbElements  As Long
  cLocks      As Long
  pvData      As Long
  cElements   As Long
  lLBound     As Long
End Type

Private mtSAHeader      As SafeArray1D
Private mvArray()       As Variant 'never dimensioned, accesses memory already allocated

Private moSubclass      As pcSubclassVTable

Private Enum eVTable
'      Ignore item 1: QueryInterface
'      Ignore item 2: AddRef
'      Ignore item 3: Release
    vtblNext = 4
    vtblSkip
    vtblReset
    vtblClone
    vtblCount
End Enum

Public Sub ReplaceIEnumVARIANT(ByVal oObject As vbBaseTlb.IEnumVARIANT)
'replace vtable for IEnumVARIANT interface
    
    If moSubclass Is Nothing Then Set moSubclass = New pcSubclassVTable
    
    If moSubclass.RefCount = 0& Then
        moSubclass.Subclass ObjPtr(oObject), vtblCount, vtblNext, _
                            AddressOf IEnumVARIANT_Next, _
                            AddressOf IEnumVARIANT_Skip, _
                            AddressOf IEnumVARIANT_Reset, _
                            AddressOf IEnumVARIANT_Clone
            
        
        'Debug.Print "Replaced vtable methods IEnumVARIANT"
        
    End If
    
    moSubclass.AddRef
    
End Sub
Public Sub RestoreIEnumVARIANT(ByVal oObject As vbBaseTlb.IEnumVARIANT)
'restore vtable for IEnumVARIANT interface

    If Not moSubclass Is Nothing Then
        moSubclass.Release
        
        If moSubclass.RefCount = 0& Then
            moSubclass.UnSubclass
            'Debug.Print "Restored vtable methods IEnumVARIANT"
            
            pInitArray 0, 0
            
        End If
        
    End If

End Sub
Private Function IEnumVARIANT_Next(ByVal oThis As Object, ByVal lngVntCount As Long, vntArray As Variant, ByVal pcvFetched As Long) As Long
'new vtable method for IEnumVARIANT::Next
    
    On Error GoTo CATCH_EXCEPTION

    Dim oEnumVARIANT As cEnumeration
    Dim liFetched As Long, lbNoMore As Boolean
    Dim i As Integer
    
    pInitArray VarPtr(vntArray), lngVntCount
    
    'cast method to source interface
    Set oEnumVARIANT = oThis
    
    'loop through each requested variant
    For i = 0 To lngVntCount - 1&
        'call the class method
        oEnumVARIANT.GetNextItem mvArray(i), lbNoMore
        
        'if nothing fetched, we're done
        If lbNoMore Then Exit For
        
        ' Count the item fetched
        liFetched = liFetched + 1&
    Next
    
    'Return success if we got all requested items
    If liFetched = lngVntCount Then
        IEnumVARIANT_Next = S_OK
        
    Else
        IEnumVARIANT_Next = S_FALSE
        
    End If
        
    'copy the actual number fetched to the pointer to fetched count
    If pcvFetched Then
        #If bVBVMTypeLib Then
            MemLong(ByVal pcvFetched) = liFetched
        #Else
            CopyMemory ByVal pcvFetched, liFetched, 4&
        #End If
    End If
    
    Exit Function
    
CATCH_EXCEPTION:
        
    'convert error to COM format
    IEnumVARIANT_Next = MapCOMErr(Err.Number)
    
    'iterate back, emptying the invalid fetched variants
    For i = i To 0& Step -1&
        mvArray(i) = Empty
    Next

    'return 0 as the number fetched after error
    If pcvFetched Then
        #If bVBVMTypeLib Then
            MemLong(ByVal pcvFetched) = 0&
        #Else
            CopyMemory ByVal pcvFetched, 0&, 4&
        #End If
    End If
    
End Function
Private Function IEnumVARIANT_Skip(ByVal oThis As Object, ByVal cV As Long) As Long
'new vtable method for IEnumVARIANT::Skip

    Dim oEnumVARIANT As cEnumeration
    Dim bSkippedAll As Boolean
    
    'handle all errors so we don't crash caller
    On Error GoTo CATCH_EXCEPTION
    
    'cast method to source interface
    Set oEnumVARIANT = oThis
    
    'call the class method
    oEnumVARIANT.Skip cV, bSkippedAll
   
    If bSkippedAll Then
        IEnumVARIANT_Skip = S_OK
    Else
        IEnumVARIANT_Skip = S_FALSE
    End If
    
    Exit Function
    
CATCH_EXCEPTION:
    
    IEnumVARIANT_Skip = MapCOMErr(Err.Number)
    
End Function

Private Function IEnumVARIANT_Reset(ByVal oThis As Object) As Long
    Dim oEnumVARIANT As cEnumeration
    
    On Error GoTo CATCH_EXCEPTION
    
    Set oEnumVARIANT = oThis
    oEnumVARIANT.Reset
    IEnumVARIANT_Reset = S_OK
    
    Exit Function
    
CATCH_EXCEPTION:
    
    IEnumVARIANT_Reset = MapCOMErr(Err.Number)
        
End Function

Private Function IEnumVARIANT_Clone(ByVal oThis As Object, ByRef ppEnum As vbBaseTlb.IEnumVARIANT) As Long
    
    Dim oEnumVARIANT As cEnumeration
    
    On Error GoTo CATCH_EXCEPTION
    
    Set oEnumVARIANT = oThis
    Set ppEnum = oEnumVARIANT.Clone
    
    If ppEnum Is Nothing _
        Then IEnumVARIANT_Clone = E_NOTIMPL _
        Else IEnumVARIANT_Clone = S_OK
    
    Exit Function
    
CATCH_EXCEPTION:
    
    IEnumVARIANT_Clone = MapCOMErr(Err.Number)
    
End Function


Private Sub pInitArray(ByVal iAddr As Long, icEl As Long)
    Const FADF_STATIC = &H2&      '// Array is statically allocated.
    Const FADF_FIXEDSIZE = &H10&  '// Array may not be resized or reallocated.
    Const FADF_VARIANT = &H800&   '// An array of VARIANTs.
    
    Const FADF_Flags = FADF_STATIC Or FADF_FIXEDSIZE Or FADF_VARIANT
    
    With mtSAHeader
        If .cDims = 0& Then
            .cbElements = 16
            .cDims = 1
            .fFeatures = FADF_Flags
            CopyMemory ByVal ArrPtr(mvArray), VarPtr(mtSAHeader), 4&
        End If
        .cElements = icEl + 1&
        .pvData = iAddr
    End With

End Sub

Private Function MapCOMErr(ByVal ErrNumber As Long) As Long
'map vb error to COM error

    If ErrNumber <> 0& Then
        If (ErrNumber And &H80000000) Or (ErrNumber = 1&) Then
            'Error HRESULT already set
            MapCOMErr = ErrNumber
            
        Else
            'Map back to a basic error number
            MapCOMErr = &H800A0000 Or ErrNumber
            
        End If
        
    End If
End Function

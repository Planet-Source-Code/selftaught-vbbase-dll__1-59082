VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcSubclassVTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==================================================================================================
'pcSubclassVTable.cls                   8/24/04
'
'           LINEAGE:
'               Based on vbACOM.dll from vbaccelerator.com written by Paul Wilde, which was in turn
'               based on code from Matt Curland's book about something or other.
'
'           PURPOSE:
'               Maintain pointers to old and new vtable functions, subclassing and unsubclassing each
'               as requested.
'
'           MODULES CALLED FROM THIS MODULE:
'               NONE
'
'           CLASSES CREATED BY THIS MODULE:
'               NONE
'
'==================================================================================================

#Const bVBVMTypeLib = False

Private Declare Function VirtualProtect Lib "kernel32.dll" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Private Const PAGE_EXECUTE_READWRITE As Long = &H40&

Private miVTable()      As Long         'Pointers to new VTable entries
Private miOldVTable()   As Long         'Pointers to original VTable entries

Private miPtrVTable     As Long         'pointer to the head of the VTable
Private miRefCount      As Long         'count the number of objects requiring subclass services
Private miVTableCount   As Long         'number of entries in the VTable

Public Property Get RefCount() As Long
    RefCount = miRefCount               'return the refcount
End Property

Public Sub AddRef()
    miRefCount = miRefCount + 1&        'increment the refcount
End Sub

Public Sub Release()
    miRefCount = miRefCount - 1&        'decrement the refcount
End Sub

Public Sub Subclass(ByVal iObjPtr As Long, ByVal iVTableCount As Long, ByVal iFirstEntry As Long, ParamArray iNewEntries() As Variant)
    
    iFirstEntry = iFirstEntry - 1&      'adjust from 1-based to 0-based index
    
    'If you get this assertion, you're either passing some wacky parameters
    'or you are already subclassed!
    Debug.Assert Not ( _
       iFirstEntry < 0& Or _
       iFirstEntry >= iVTableCount Or _
       iVTableCount < 0& Or _
       miPtrVTable <> 0& Or _
       iObjPtr = 0&)
    
    #If bVBVMTypeLib Then
        miPtrVTable = MemLong(ByVal iObjPtr)        'use the illegal type library
    #Else
        CopyMemory miPtrVTable, ByVal iObjPtr, 4&   'use the block-moving sledgehammer
    #End If
    
    miVTableCount = iVTableCount                    'store the VTable count
    
    ReDim miVTable(0 To miVTableCount - 1&)         'redim the arrays
    ReDim miOldVTable(0 To miVTableCount - 1&)
    
    Dim liEntry As Long
    Dim liPtrEntry As Long
    
    liEntry = UBound(iNewEntries) + iFirstEntry                             'get the number of entries that we will subclass
    If liEntry >= miVTableCount Then liEntry = miVTableCount - 1&           'make sure it's not bigger than the count
    liPtrEntry = UnsignedAdd(miPtrVTable, iFirstEntry * 4&)                 'get the pointer to the first entry we care about
    
    For liEntry = iFirstEntry To liEntry
        miVTable(liEntry) = iNewEntries(liEntry - iFirstEntry)              'store the new vtable address
        If miVTable(liEntry) <> 0& Then                                     'if we are subclassing this entry
            pSubclass liPtrEntry, miVTable(liEntry), miOldVTable(liEntry)   'set the subclass
        End If
        liPtrEntry = UnsignedAdd(liPtrEntry, 4&)                            'increment the VTable entry pointer
    Next
    
End Sub

Public Property Let SubclassEntry(ByVal iEntry As Long, ByVal bNew As Boolean)
    
    iEntry = iEntry - 1&                                                    'change from 1-based to 0-based
    
    Dim liPtrEntry As Long
    
    'if you get this assertion, you're either passing a bad parameter or have not yet set the subclass
    Debug.Assert iEntry > -1& And iEntry < miVTableCount And miPtrVTable <> 0&
    
    If SubclassEntry(iEntry + 1&) Xor bNew Then
    
        ' Calculate pointer to specified entry
        liPtrEntry = UnsignedAdd(miPtrVTable, iEntry * 4&)
        
        If bNew Then
            pSubclass liPtrEntry, miVTable(iEntry), miOldVTable(iEntry)     'set the new entry
            
        Else
            pSubclass liPtrEntry, miOldVTable(iEntry), 0&                   'restore the old entry
            miOldVTable(iEntry) = 0&                                        'release the pointer to the old entry
            
        End If
    
    Else
        'if you get this assertion, the subclass state is the same as you are requesting
        Debug.Assert False
        
    End If
    
End Property

Public Property Get SubclassEntry(ByVal iEntry As Long) As Boolean
    
    iEntry = iEntry - 1&        '1-based to 0-based
    
    'bad parameter
    Debug.Assert iEntry > -1& And iEntry < miVTableCount And miPtrVTable <> 0&
    
    SubclassEntry = CBool(miOldVTable(iEntry))  'we are subclassed if we have a pointer to the old entry
    
End Property

Public Sub UnSubclass()
    
    If miPtrVTable <> 0& Then                       'if we are subclassed
        Dim i As Long
        
        For i = 0& To miVTableCount - 1&            'loop through each entry
            If miOldVTable(i) Then                  'If we have an old entry
                pSubclass UnsignedAdd(miPtrVTable, i * 4&), miOldVTable(i), 0&  'restore the old entry
                miOldVTable(i) = 0&                 'release the pointer
            End If
        Next
        
        miPtrVTable = 0&                            'release the pointer to the VTable
        
    Else
        Debug.Assert False                          'you can't unsubclass if you're not subclassed!
        
    End If
    
End Sub

Private Sub pSubclass(ByVal iPtrEntry As Long, ByVal iPtrNew As Long, ByRef iPtrOld As Long)

    ' Save address of previous method for return
    #If bVBVMTypeLib Then
        iPtrOld = MemLong(ByVal iPtrEntry)
    #Else
        CopyMemory iPtrOld, ByVal iPtrEntry, 4&
    #End If
    
    ' Ignore if they're already the same
    If iPtrOld <> iPtrNew Then
        ' Need to change page protection to write to code
        VirtualProtect iPtrEntry, 4&, PAGE_EXECUTE_READWRITE, 0&
        ' Write the new function address into the v-table
        #If bVBVMTypeLib Then
            MemLong(ByVal iPtrEntry) = iPtrNew
        #Else
            CopyMemory ByVal iPtrEntry, iPtrNew, 4&
        #End If
        ' Restore the previous page protection
        'VirtualProtect liPtrEntry, 4&,  OldProtect, lOldProtect 'Optional
    Else
        Debug.Assert False
        
    End If
    
End Sub

Private Sub Class_Terminate()
    If miPtrVTable Then UnSubclass
End Sub

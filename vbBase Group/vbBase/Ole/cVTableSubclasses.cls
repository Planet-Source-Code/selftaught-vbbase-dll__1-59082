VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cVTableSubclasses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
'cVTableSubclasses.cls                            7/5/04
'
'           PURPOSE:
'               Adds and removes objects from ole vtable subclasses.
'               This may not work as you expect.  Due to the nature of v-table
'               subclassing, all objects that implement both the ole and the
'               subclass interface will receive callbacks as long as a single
'               object has attached by calling this function.

'               In spite of this, it is the recommended practice to attach and
'               detach all objects using this subclassing in their initialize and
'               terminate events, to ensure that the subclassing starts and
'               stops correctly.  Improper use of v-table subclassing can cause
'               some spectacular crashes.
'
'           MODULES CALLED FROM THIS MODULE:
'               mVTableSubclass.bas
'
'==================================================================================================


Public Function Add(ByVal oObject As Object) As Boolean
    Add = mVTableSubclass.Attach(oObject)
End Function

Public Function Remove(ByVal oObject As Object) As Boolean
    Remove = mVTableSubclass.Detach(oObject)
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEnumeration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
'cEnumeration.cls                       7/4/04
'
'           PURPOSE:
'               Holds a placemarker and other client-specific data for an enumeration.
'
'==================================================================================================

Implements vbBaseTlb.IEnumVARIANT

Public Type tEnum
    iIndex      As Long
    iControl    As Long
    iData       As Long
End Type

Private mtEnum As tEnum
Private moOwner As iEnumerable

Private Sub IEnumVARIANT_Clone(ppEnum As vbBaseTlb.IEnumVARIANT)
'dummy - see mEnumVariant.bas
End Sub

Private Sub IEnumVARIANT_Next(ByVal Celt As Long, rgVar As Variant, ByVal pCeltFetched As Long)
'dummy - see mEnumVariant.bas
End Sub

Private Sub IEnumVARIANT_Reset()
'dummy - see mEnumVariant.bas
End Sub

Private Sub IEnumVARIANT_Skip(ByVal Celt As Long)
'dummy - see mEnumVariant.bas
End Sub


Public Function GetEnum( _
            ByVal oWho As iEnumerable, _
            ByVal iControl As Long, _
   Optional ByVal iData As Long, _
   Optional ByVal iIndex As Long = -1&) _
                As IUnknown
                
    With mtEnum
        .iIndex = iIndex
        .iControl = iControl
        .iData = iData
    End With

    Set GetEnum = Me
    Set moOwner = oWho
End Function

Private Sub Class_Initialize()
    ReplaceIEnumVARIANT Me
End Sub

Private Sub Class_Terminate()
    If Not moOwner Is Nothing Then moOwner.Cleanup mtEnum
    Set moOwner = Nothing
    RestoreIEnumVARIANT Me
End Sub

Friend Sub GetNextItem( _
            ByRef vNextItem As Variant, _
            ByRef bNoMoreItems As Boolean)
    Debug.Assert Not moOwner Is Nothing
    If Not moOwner Is Nothing Then moOwner.GetNextItem mtEnum, vNextItem, bNoMoreItems Else bNoMoreItems = True
End Sub

Friend Sub Reset()
    Debug.Assert Not moOwner Is Nothing
    If Not moOwner Is Nothing Then moOwner.Reset mtEnum
End Sub

Friend Sub Skip( _
            ByVal iSkipCount As Long, _
            ByRef bSkippedAll As Boolean)
    Debug.Assert Not moOwner Is Nothing
    If Not moOwner Is Nothing Then moOwner.Skip mtEnum, iSkipCount, bSkippedAll
End Sub

Friend Function Clone() As IUnknown
    Debug.Assert Not moOwner Is Nothing
    If Not moOwner Is Nothing Then Set Clone = moOwner.Clone()
End Function


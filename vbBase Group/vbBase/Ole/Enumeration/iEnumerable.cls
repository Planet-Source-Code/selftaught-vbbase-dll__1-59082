VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iEnumerable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
'iEnumerable.cls                        8/19/04
'
'           PURPOSE:
'               High level callback procedure for iEnumVARIANT
'
'==================================================================================================
 

Public Sub GetNextItem(tEnum As tEnum, vNextItem As Variant, bNoMoreItems As Boolean)
End Sub

Public Sub Reset(tEnum As tEnum)
End Sub

Public Sub Skip(tEnum As tEnum, ByVal iSkipCount As Long, ByRef bSkippedAll As Boolean)
End Sub

Public Sub Cleanup(tEnum As tEnum)
End Sub

Public Function Clone() As IUnknown
End Function

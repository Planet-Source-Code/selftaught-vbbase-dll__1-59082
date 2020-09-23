VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iPerPropertyBrowsingVB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2650
'==================================================================================================
'iPerPropertyBrowsingVB.cls             8/25/04
'
'           LINEAGE:
'               vbACOM.dll from vbaccelerator.com
'
'           PURPOSE:
'               Subclass interface for IPerPropertyBrowsing.
'
'==================================================================================================

Option Explicit

Public Sub GetPredefinedValue(ByRef bHandled As Boolean, ByVal iDispID As Long, ByVal iCookie As Long, vValue As Variant)
End Sub
Public Sub GetDisplayString(ByRef bHandled As Boolean, ByVal iDispID As Long, sDisplayName As String)
End Sub
Public Sub MapPropertyToPage(ByRef bHandled As Boolean, ByVal iDispID As Long, sClassID As String)
End Sub
Public Sub GetPredefinedStrings(ByRef bHandled As Boolean, ByVal iDispID As Long, ByVal oProperties As cPropertyListItems)
End Sub


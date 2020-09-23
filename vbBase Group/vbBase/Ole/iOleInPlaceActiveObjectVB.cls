VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iOleInPlaceActiveObjectVB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2530
'==================================================================================================
'iOleInPlaceActiveObject.cls            8/25/04
'
'           LINEAGE:
'               vbACOM.dll from vbaccelerator.com
'
'           PURPOSE:
'               Subclass interface for IOleInPlaceActiveObject.
'
'==================================================================================================

Option Explicit



Public Sub OnFrameWindowActivate(ByRef bHandled As Boolean, ByVal bActive As Boolean)
End Sub
Public Sub OnDocWindowActivate(ByRef bHandled As Boolean, ByVal bActive As Boolean)
End Sub
Public Sub EnableModeless(ByRef bHandled As Boolean, ByVal bEnable As Boolean)
End Sub
Public Sub ResizeBorder(ByRef bHandled As Boolean, ByRef tBorder As tRect, ByVal oUIWindow As Object, ByVal bFrameWindow As Boolean)
End Sub
Public Sub TranslateAccelerator(ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef tMsg As tMsg, ByVal iShift As ShiftConstants)
End Sub


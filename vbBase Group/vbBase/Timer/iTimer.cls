VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "iTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
'iTimer.cls                             8/19/04
'
'           LINEAGE:
'               Paul Caton's WinSubHook.tlb at www.planet-source-code.com
'
'           PURPOSE:
'               Callback procedure for api timers.
'
'==================================================================================================

Public Sub Proc( _
            ByVal lElapsedMS As Long, _
            ByVal lTimerID As Long)

End Sub

VERSION 5.00
Begin VB.UserControl ucTest 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Label Label1 
      Caption         =   $"ucTest.ctx":0000
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "ucTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements iOleInPlaceActiveObjectVB
Implements iPerPropertyBrowsingVB

Private miDispIdTest As Long

Public Property Get Test() As String

End Property

Public Property Let Test(ByRef sNew As String)

End Property

Private Sub iOleInPlaceActiveObjectVB_EnableModeless(bHandled As Boolean, ByVal bEnable As Boolean)

End Sub

Private Sub iOleInPlaceActiveObjectVB_OnDocWindowActivate(bHandled As Boolean, ByVal bActive As Boolean)

End Sub

Private Sub iOleInPlaceActiveObjectVB_OnFrameWindowActivate(bHandled As Boolean, ByVal bActive As Boolean)

End Sub

Private Sub iOleInPlaceActiveObjectVB_ResizeBorder(bHandled As Boolean, tBorder As vbBase.tRect, ByVal oUIWindow As Object, ByVal bFrameWindow As Boolean)

End Sub

Private Sub iOleInPlaceActiveObjectVB_TranslateAccelerator(bHandled As Boolean, lReturn As Long, tMsg As vbBase.tMsg, ByVal iShift As ShiftConstants)
        Select Case tMsg.Message
        Case WM_KEYDOWN
            Select Case tMsg.wParam And &HFFFF&
            Case vbKeyTab
                If MsgBox("You pressed Tab!  Do you want to intercept this keypress?", vbYesNo) = vbYes Then
                    bHandled = True
                End If
            Case vbKeyReturn
                If MsgBox("You pressed Enter!  Do you want to intercept this keypress?", vbYesNo) = vbYes Then
                    bHandled = True
                End If
         End Select
      End Select
End Sub

Private Sub iPerPropertyBrowsingVB_GetDisplayString(bHandled As Boolean, ByVal iDispID As Long, sDisplayName As String)

End Sub

Private Sub iPerPropertyBrowsingVB_GetPredefinedStrings(bHandled As Boolean, ByVal iDispID As Long, ByVal oProperties As vbBase.cPropertyListItems)
    If iDispID = miDispIdTest Then
        bHandled = True
        Dim i As Long
        For i = 1 To 10
            oProperties.Add "List Item " & i, 0
        Next
    End If
End Sub

Private Sub iPerPropertyBrowsingVB_GetPredefinedValue(bHandled As Boolean, ByVal iDispID As Long, ByVal iCookie As Long, vValue As Variant)
    If iDispID = miDispIdTest Then
        bHandled = True
        vValue = ""
    End If
End Sub

Private Sub iPerPropertyBrowsingVB_MapPropertyToPage(bHandled As Boolean, ByVal iDispID As Long, sClassID As String)
    miDispIdTest = xGetDispId(Me, "Test")
End Sub

Private Sub UserControl_Initialize()
    VTableSubclasses.Add Me
End Sub

Private Sub UserControl_Terminate()
    VTableSubclasses.Remove Me
End Sub

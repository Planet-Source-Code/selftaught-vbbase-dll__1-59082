VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   4860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin Project1.ucTest ucTest1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   3201
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    MsgBox "Default button clicked!"
End Sub

Private Sub Form_Load()
    ForceWindowToShowAllUIStates hWnd
End Sub

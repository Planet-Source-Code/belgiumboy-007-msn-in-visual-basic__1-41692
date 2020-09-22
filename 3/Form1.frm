VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change your status."
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   1950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   1950
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optOnThePhone 
      Caption         =   "On The Phone"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin VB.OptionButton optBeRightBack 
      Caption         =   "Be Right Back"
      Height          =   255
      Left            =   248
      TabIndex        =   5
      Top             =   600
      Width           =   1455
   End
   Begin VB.OptionButton optBusy 
      Caption         =   "Busy"
      Height          =   255
      Left            =   248
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
   Begin VB.OptionButton optOnline 
      Caption         =   "Online"
      Height          =   255
      Left            =   248
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.OptionButton optAppearOffline 
      Caption         =   "Appear Offline"
      Height          =   255
      Left            =   248
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.OptionButton optOutToLunch 
      Caption         =   "Out To Lunch"
      Height          =   255
      Left            =   248
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.OptionButton optAway 
      Caption         =   "Away"
      Height          =   255
      Left            =   248
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MSN As New MsgrObject

Private Sub Form_Load()
    Select Case MSN.LocalState
        Case MSTATE_ONLINE
            optOnline.Value = True
        Case MSTATE_BUSY
            optBusy.Value = True
        Case MSTATE_BE_RIGHT_BACK
            optBeRightBack.Value = True
        Case MSTATE_AWAY
            optAway.Value = True
        Case MSTATE_ON_THE_PHONE
            optOnThePhone.Value = True
        Case MSTATE_OUT_TO_LUNCH
            optOutToLunch.Value = True
        Case MSTATE_INVISIBLE
            optAppearOffline.Value = True
    End Select
End Sub

Private Sub optAppearOffline_Click()
    MSN.LocalState = MSTATE_INVISIBLE
End Sub

Private Sub optAway_Click()
    MSN.LocalState = MSTATE_AWAY
End Sub

Private Sub optBeRightBack_Click()
    MSN.LocalState = MSTATE_BE_RIGHT_BACK
End Sub

Private Sub optBusy_Click()
    MSN.LocalState = MSTATE_BUSY
End Sub

Private Sub optOnline_Click()
    MSN.LocalState = MSTATE_ONLINE
End Sub

Private Sub optOnThePhone_Click()
    MSN.LocalState = MSTATE_ON_THE_PHONE
End Sub

Private Sub optOutToLunch_Click()
    MSN.LocalState = MSTATE_OUT_TO_LUNCH
End Sub

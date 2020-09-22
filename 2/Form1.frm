VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change your nickname."
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdChangeNickName 
      Caption         =   "Change NickName"
      Height          =   375
      Left            =   2933
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox txtNewNickName 
      Height          =   285
      Left            =   113
      TabIndex        =   1
      Top             =   360
      Width           =   7335
   End
   Begin VB.Label lblNewNickName 
      Caption         =   "New NickName:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MSN As New MsgrObject

Private Sub cmdChangeNickName_Click()
    If MSN.LocalState = MSTATE_OFFLINE Then
        MsgBox "You are not Signed In"
    Else
        MSN.Services.PrimaryService.FriendlyName = txtNewNickName.Text
        txtNewNickName.Text = ""
    End If
End Sub

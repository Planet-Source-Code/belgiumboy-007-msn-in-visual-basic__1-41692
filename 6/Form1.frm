VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send an Instant Message."
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4905
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   4905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRefreshList 
      Caption         =   "Refresh List"
      Height          =   375
      Left            =   3345
      TabIndex        =   3
      Top             =   1725
      Width           =   1455
   End
   Begin VB.CommandButton cmdSendIM 
      Caption         =   "Send IM"
      Height          =   375
      Left            =   3345
      TabIndex        =   2
      Top             =   2325
      Width           =   1455
   End
   Begin VB.ListBox lstOnlineContacts 
      Height          =   3960
      Left            =   105
      TabIndex        =   0
      Top             =   352
      Width           =   3015
   End
   Begin VB.Label lblOnlineContacts 
      Caption         =   "Online Contacts:"
      Height          =   255
      Left            =   105
      TabIndex        =   1
      Top             =   112
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MSN As New MsgrObject

Private Sub RefreshList()
    lstOnlineContacts.Visible = False

    Dim User As IMsgrUser
    
    lstOnlineContacts.Clear
    
    For Each User In MSN.List(MLIST_CONTACT)
        If User.State <> MSTATE_OFFLINE Then lstOnlineContacts.AddItem (User.EmailAddress)
    Next
    
    lstOnlineContacts.Visible = True
End Sub

Private Sub cmdRefreshList_Click()
    If MSN.LocalState <> MSTATE_OFFLINE Then RefreshList
End Sub

Private Sub cmdSendIM_Click()
    Dim User As IMsgrUser
    Dim bstrMsgHeader As String
    Dim bstrMsgText As String
    
    If MSN.LocalState = MSTATE_OFFLINE Then
        MsgBox "You are not Signed In"
    Else
        If MSN.LocalState = MSTATE_INVISIBLE Then
            MsgBox "Change you status first !"
        Else
            Set User = MSN.CreateUser(lstOnlineContacts.Text, MSN.Services.PrimaryService)
            bstrMsgText = InputBox("Enter text to send", "Send What ?", "Howdy", Me.Left, Me.Top)
            User.SendText bstrMsgHeader, bstrMsgText, MMSGTYPE_NO_RESULT
            MsgBox "The following message was sent to " & User.EmailAddress & " : " & bstrMsgText
        End If
    End If
End Sub

Private Sub Form_Load()
    cmdRefreshList_Click
End Sub

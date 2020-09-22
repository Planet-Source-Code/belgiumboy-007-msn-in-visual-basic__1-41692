VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send your contacts list to a listbox."
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRefreshList 
      Caption         =   "Refresh List"
      Height          =   375
      Left            =   2460
      TabIndex        =   4
      Top             =   4440
      Width           =   1455
   End
   Begin VB.ListBox lstOfflineContacts 
      Height          =   3960
      Left            =   3240
      TabIndex        =   3
      Top             =   360
      Width           =   3015
   End
   Begin VB.ListBox lstOnlineContacts 
      Height          =   3960
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label lblOfflineContacts 
      Caption         =   "Offline Contacts:"
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label lblOnlineContacts 
      Caption         =   "Online Contacts:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MSN As New MsgrObject

Private Sub RefreshList()
    lstOfflineContacts.Visible = False
    lstOnlineContacts.Visible = False

    Dim User As IMsgrUser
    
    lstOnlineContacts.Clear
    lstOfflineContacts.Clear
    
    For Each User In MSN.List(MLIST_CONTACT)
        If User.State = MSTATE_OFFLINE Then
            lstOfflineContacts.AddItem (User.EmailAddress)
        Else
            lstOnlineContacts.AddItem (User.EmailAddress)
        End If
    Next
    
    lstOfflineContacts.Visible = True
    lstOnlineContacts.Visible = True
End Sub

Private Sub cmdRefreshList_Click()
    If MSN.LocalState <> MSTATE_OFFLINE Then RefreshList
End Sub

Private Sub Form_Load()
    cmdRefreshList_Click
End Sub

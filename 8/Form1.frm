VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Block/unblock contacts."
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
   Begin VB.CommandButton cmdRefreshLists 
      Caption         =   "Refresh Lists"
      Height          =   375
      Left            =   2460
      TabIndex        =   6
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdBlock 
      Caption         =   "<< Block"
      Height          =   375
      Left            =   4260
      TabIndex        =   5
      Top             =   4418
      Width           =   1455
   End
   Begin VB.CommandButton cmdAllow 
      Caption         =   "Allow >>"
      Height          =   375
      Left            =   660
      TabIndex        =   4
      Top             =   4418
      Width           =   1455
   End
   Begin VB.ListBox lstAllow 
      Height          =   3960
      Left            =   3240
      TabIndex        =   3
      Top             =   338
      Width           =   3015
   End
   Begin VB.ListBox lstBlock 
      Height          =   3960
      Left            =   120
      TabIndex        =   2
      Top             =   338
      Width           =   3015
   End
   Begin VB.Label lblAllow 
      Caption         =   "Allow List:"
      Height          =   255
      Left            =   3240
      TabIndex        =   1
      Top             =   98
      Width           =   1815
   End
   Begin VB.Label lblBlock 
      Caption         =   "Block List:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   98
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MSN As New MsgrObject

Private Sub RefreshLists()
    lstBlock.Visible = False
    lstAllow.Visible = False
    
    lstBlock.Clear
    lstAllow.Clear

    Dim User As IMsgrUser
    
    For Each User In MSN.List(MLIST_BLOCK)
        lstBlock.AddItem (User.EmailAddress)
    Next
    
    For Each User In MSN.List(MLIST_ALLOW)
        lstAllow.AddItem (User.EmailAddress)
    Next
    
    lstBlock.Visible = True
    lstAllow.Visible = True
End Sub

Private Sub cmdAllow_Click()
    Dim User As IMsgrUser
    
    Set User = MSN.CreateUser(lstBlock.Text, MSN.Services.PrimaryService)

    MSN.List(MLIST_ALLOW).Add User
    MSN.List(MLIST_BLOCK).Remove User
    
    lstAllow.AddItem (lstBlock.Text)
    lstBlock.RemoveItem (lstBlock.ListIndex)
End Sub

Private Sub cmdBlock_Click()
    Dim User As IMsgrUser
    
    Set User = MSN.CreateUser(lstAllow.Text, MSN.Services.PrimaryService)

    MSN.List(MLIST_ALLOW).Remove User
    MSN.List(MLIST_BLOCK).Add User
    
    lstBlock.AddItem (lstAllow.Text)
    lstAllow.RemoveItem (lstAllow.ListIndex)
End Sub

Private Sub cmdRefreshLists_Click()
    RefreshLists
End Sub

Private Sub Form_Load()
    RefreshLists
End Sub

VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add/remove contacts."
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRefreshList 
      Caption         =   "Refresh List"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   2558
      Width           =   1455
   End
   Begin VB.CommandButton cmdRemoveContact 
      Caption         =   "Remove Contact"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   1958
      Width           =   1455
   End
   Begin VB.CommandButton cmdAddContact 
      Caption         =   "Add Contact"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   1478
      Width           =   1455
   End
   Begin VB.ListBox lstContacts 
      Height          =   3960
      Left            =   120
      TabIndex        =   0
      Top             =   345
      Width           =   3015
   End
   Begin VB.Label lblContacts 
      Caption         =   "Contacts:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   105
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MSN As New MsgrObject

Private Sub RefreshList()
    lstContacts.Visible = False
    
    lstContacts.Clear
    
    Dim User As IMsgrUser
    
    For Each User In MSN.List(MLIST_CONTACT)
        lstContacts.AddItem (User.EmailAddress)
    Next
    
    lstContacts.Visible = True
End Sub

Private Sub cmdAddContact_Click()
    Dim User As IMsgrUser
    
    Set User = MSN.CreateUser(InputBox("Enter users email address.", "Add Contact", "", Me.Left, Me.Top), MSN.Services.PrimaryService)
    
    MSN.List(MLIST_CONTACT).Add User
    MSN.List(MLIST_ALLOW).Add User
    
    lstContacts.AddItem User.EmailAddress
End Sub

Private Sub cmdRefreshList_Click()
    RefreshList
End Sub

Private Sub cmdRemoveContact_Click()
On Error Resume Next
    Dim User As IMsgrUser
    
    Set User = MSN.CreateUser(lstContacts.Text, MSN.Services.PrimaryService)
    
    MSN.List(MLIST_ALLOW).Remove User
    MSN.List(MLIST_CONTACT).Remove User
    MSN.List(MLIST_BLOCK).Remove User
    
    lstContacts.RemoveItem (lstContacts.ListIndex)
End Sub

Private Sub Form_Load()
    RefreshList
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send your contacts list to a treeview."
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRefreshList 
      Caption         =   "Refresh List"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   5145
      Width           =   1455
   End
   Begin MSComctlLib.ImageList ilsIcons 
      Left            =   2242
      Top             =   2512
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   "Away"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0352
            Key             =   "AwaySelected"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":06A4
            Key             =   "Blocked"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":09F6
            Key             =   "BlockedSelected"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0D48
            Key             =   "Busy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":109A
            Key             =   "BusySelected"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":13EC
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":173E
            Key             =   "DownSelected"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1A90
            Key             =   "Offline"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1DE2
            Key             =   "OfflineSelected"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2134
            Key             =   "Online"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2486
            Key             =   "OnlineSelected"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":27D8
            Key             =   "Up"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2B2A
            Key             =   "UpSelected"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tContacts 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   345
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   8281
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      Style           =   1
      ImageList       =   "ilsIcons"
      Appearance      =   0
   End
   Begin VB.Label lblContacts 
      BackStyle       =   0  'Transparent
      Caption         =   "Contacts:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   105
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MSN As New MsgrObject

Private Sub RefreshList()
    tContacts.Visible = False
    
    tContacts.Nodes.Clear
    
    Dim User As IMsgrUser
    Dim UsersOnline As Integer
    Dim UsersOffline As Integer
    
    For Each User In MSN.List(MLIST_CONTACT)
        If User.State = MSTATE_OFFLINE Then
            UsersOffline = UsersOffline + 1
        Else
            UsersOnline = UsersOnline + 1
        End If
    Next
    
    tContacts.Nodes.Add , , "Online", "Online (" & UsersOnline & ")", "Up", "UpSelected"
    With tContacts.Nodes(1)
        .Selected = False
        .Expanded = True
        .Bold = True
        .ForeColor = &H8000000D
        .Sorted = True
    End With
    
    tContacts.Nodes.Add , , "Offline", "Offline (" & UsersOffline & ")", "Down", "DownSelected"
    With tContacts.Nodes(2)
        .Selected = True
        .Expanded = False
        .Bold = True
        .ForeColor = &H8000000D
        .Sorted = True
    End With
    
    For Each User In MSN.List(MLIST_ALLOW)
        Select Case User.State
            Case MSTATE_AWAY
                tContacts.Nodes.Add "Online", tvwChild, User.EmailAddress, User.FriendlyName & " (Away)", "Away", "AwaySelected"
            Case MSTATE_BE_RIGHT_BACK
                tContacts.Nodes.Add "Online", tvwChild, User.EmailAddress, User.FriendlyName & " (Be Right Back)", "Away", "AwaySelected"
            Case MSTATE_BUSY
                tContacts.Nodes.Add "Online", tvwChild, User.EmailAddress, User.FriendlyName & " (Busy)", "Busy", "BusySelected"
            Case MSTATE_OFFLINE
                tContacts.Nodes.Add "Offline", tvwChild, User.EmailAddress, User.FriendlyName, "Offline", "OfflineSelected"
            Case MSTATE_ON_THE_PHONE
                tContacts.Nodes.Add "Online", tvwChild, User.EmailAddress, User.FriendlyName & " (On The Phone)", "Busy", "BusySelected"
            Case MSTATE_ONLINE
                tContacts.Nodes.Add "Online", tvwChild, User.EmailAddress, User.FriendlyName, "Online", "OnlineSelected"
            Case MSTATE_OUT_TO_LUNCH
                tContacts.Nodes.Add "Online", tvwChild, User.EmailAddress, User.FriendlyName & " (Out To Lunch)", "Away", "AwaySelected"
        End Select
    Next
    
    For Each User In MSN.List(MLIST_BLOCK)
        If User.State = MSTATE_OFFLINE Then
            tContacts.Nodes.Add "Offline", tvwChild, User.EmailAddress, User.FriendlyName & " (Blocked)", "Blocked", "BlockedSelected"
        Else
            tContacts.Nodes.Add "Offline", tvwChild, User.EmailAddress, User.FriendlyName & " (Blocked)", "Blocked", "BlockedSelected"
        End If
    Next
    
    tContacts.Visible = True
End Sub

Private Sub cmdRefreshList_Click()
    RefreshList
End Sub

Private Sub Form_Load()
    RefreshList
End Sub


Private Sub tContacts_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Key = "Offline" Or Node.Key = "Online" Then
        If Node.Expanded = True Then
            Node.Expanded = False
            Node.Image = "Down"
            Node.SelectedImage = "DownSelected"
        Else
            Node.Expanded = True
            Node.Image = "Up"
            Node.SelectedImage = "UpSelected"
        End If
    Else

    End If
End Sub

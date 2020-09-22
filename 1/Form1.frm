VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Signing in to/out of the .NET Messenger Service."
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1500
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox txtUserEmail 
      Height          =   285
      Left            =   1500
      TabIndex        =   2
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton cmdSignOut 
      Caption         =   "Sign &Out"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdSignIn 
      Caption         =   "Sign &In"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password:"
      Height          =   255
      Left            =   300
      TabIndex        =   5
      Top             =   515
      Width           =   1215
   End
   Begin VB.Label lblUserEmail 
      Caption         =   "E-mail address:"
      Height          =   255
      Left            =   300
      TabIndex        =   4
      Top             =   155
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MSN As New MsgrObject
Attribute MSN.VB_VarHelpID = -1

Private Sub cmdSignIn_Click()
On Error Resume Next
    MSN.Logon txtUserEmail.Text, txtPassword.Text, MSN.Services.PrimaryService
End Sub

Private Sub cmdSignOut_Click()
On Error Resume Next
    MSN.Logoff
End Sub

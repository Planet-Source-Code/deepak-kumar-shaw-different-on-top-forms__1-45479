VERSION 5.00
Begin VB.Form frmConfig 
   Caption         =   "Set form Loading Configuration"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2805
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1665
      TabIndex        =   4
      Top             =   2310
      Width           =   1185
   End
   Begin VB.Frame fraCnf 
      Caption         =   "Configuration"
      Height          =   2040
      Left            =   210
      TabIndex        =   0
      Top             =   90
      Width           =   4230
      Begin VB.OptionButton Option1 
         Caption         =   "&Load Inactive form"
         Height          =   240
         Index           =   3
         Left            =   285
         TabIndex        =   5
         Top             =   1635
         Value           =   -1  'True
         Width           =   3630
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&On Top Form (On all other Application)"
         Height          =   240
         Index           =   2
         Left            =   285
         TabIndex        =   3
         Top             =   1215
         Width           =   3630
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&Form On Top Within Application"
         Height          =   240
         Index           =   1
         Left            =   285
         TabIndex        =   2
         Top             =   780
         Width           =   3630
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&Normal Form Loading"
         Height          =   240
         Index           =   0
         Left            =   285
         TabIndex        =   1
         Top             =   345
         Width           =   3630
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private OptionValue As Byte
Private Sub cmdOK_Click()
    frmBase.MyOption OptionValue
    Me.Hide
End Sub

Private Sub Option1_Click(Index As Integer)
    If Option1(0).Value = True Then OptionValue = 0
    If Option1(1).Value = True Then OptionValue = 1
    If Option1(2).Value = True Then OptionValue = 2
    If Option1(3).Value = True Then OptionValue = 3
End Sub

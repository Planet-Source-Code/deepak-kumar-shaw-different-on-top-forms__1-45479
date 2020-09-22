VERSION 5.00
Begin VB.Form frmBase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form On top Type's"
   ClientHeight    =   4275
   ClientLeft      =   4365
   ClientTop       =   3015
   ClientWidth     =   7740
   Icon            =   "frmBase.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   7740
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Close"
      Height          =   450
      Index           =   2
      Left            =   5490
      TabIndex        =   10
      Top             =   3450
      Width           =   1605
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Set Configuration"
      Height          =   450
      Index           =   1
      Left            =   2565
      TabIndex        =   8
      Top             =   3450
      Width           =   1605
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Show My Details"
      Height          =   450
      Index           =   0
      Left            =   600
      TabIndex        =   7
      Top             =   3450
      Width           =   1605
   End
   Begin VB.Frame fraPers 
      Caption         =   "Personal Information"
      Height          =   3135
      Left            =   555
      TabIndex        =   0
      Top             =   150
      Width           =   6555
      Begin VB.TextBox txtPInfo 
         Height          =   780
         Index           =   2
         Left            =   2235
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1695
         Width           =   3090
      End
      Begin VB.TextBox txtPInfo 
         Height          =   315
         Index           =   1
         Left            =   2220
         TabIndex        =   4
         Top             =   1080
         Width           =   3090
      End
      Begin VB.TextBox txtPInfo 
         Height          =   315
         Index           =   0
         Left            =   2220
         TabIndex        =   2
         Top             =   480
         Width           =   3090
      End
      Begin VB.Label lblTitles 
         Alignment       =   1  'Right Justify
         Caption         =   "Address"
         Height          =   225
         Index           =   2
         Left            =   420
         TabIndex        =   5
         Top             =   1770
         Width           =   1260
      End
      Begin VB.Label lblTitles 
         Alignment       =   1  'Right Justify
         Caption         =   "SurName"
         Height          =   225
         Index           =   1
         Left            =   405
         TabIndex        =   3
         Top             =   1155
         Width           =   1260
      End
      Begin VB.Label lblTitles 
         Alignment       =   1  'Right Justify
         Caption         =   "Name"
         Height          =   225
         Index           =   0
         Left            =   390
         TabIndex        =   1
         Top             =   555
         Width           =   1260
      End
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   6975
      TabIndex        =   9
      Top             =   3945
      Width           =   705
   End
End
Attribute VB_Name = "frmBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mMyOption As Byte
Private Sub cmdBtn_Click(Index As Integer)

Select Case Index
    Case 0 'Action
        frmInfo.ShowME mMyOption
    Case 1 'Config Window
        frmConfig.Show vbModal
    Case 2 'Close
        End
End Select

End Sub
Friend Sub MyOption(vValue As Byte)
    mMyOption = vValue
End Sub

Private Sub Form_Load()
    mMyOption = 1
    frmConfig.Option1(mMyOption).Value = True
End Sub

Private Sub lblInfo_Click()
MsgBox "Please feel free to write your Comments/Suggestions. Thnx!" & vbCrLf & "-Deepakk_2k@yahoo.com", vbInformation, "Thanks for your feeback"
End Sub


Private Sub lblInfo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblInfo.FontUnderline = True
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblInfo.FontUnderline = False
End Sub

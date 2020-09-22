VERSION 5.00
Begin VB.Form frmInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Your Personal Information"
   ClientHeight    =   3435
   ClientLeft      =   6465
   ClientTop       =   5580
   ClientWidth     =   6810
   Icon            =   "frmInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   6810
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   435
      Left            =   2340
      TabIndex        =   7
      Top             =   2820
      Width           =   1800
   End
   Begin VB.Frame fraPers 
      Caption         =   "Personal Information"
      Height          =   2625
      Left            =   135
      TabIndex        =   0
      Top             =   60
      Width           =   6555
      Begin VB.TextBox txtPInfo 
         Height          =   315
         Index           =   0
         Left            =   2220
         TabIndex        =   3
         Top             =   480
         Width           =   3090
      End
      Begin VB.TextBox txtPInfo 
         Height          =   315
         Index           =   1
         Left            =   2220
         TabIndex        =   2
         Top             =   1080
         Width           =   3090
      End
      Begin VB.TextBox txtPInfo 
         Height          =   780
         Index           =   2
         Left            =   2235
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   1695
         Width           =   3090
      End
      Begin VB.Label lblTitles 
         Alignment       =   1  'Right Justify
         Caption         =   "Name"
         Height          =   225
         Index           =   0
         Left            =   390
         TabIndex        =   6
         Top             =   555
         Width           =   1260
      End
      Begin VB.Label lblTitles 
         Alignment       =   1  'Right Justify
         Caption         =   "SurName"
         Height          =   225
         Index           =   1
         Left            =   405
         TabIndex        =   5
         Top             =   1155
         Width           =   1260
      End
      Begin VB.Label lblTitles 
         Alignment       =   1  'Right Justify
         Caption         =   "Address"
         Height          =   225
         Index           =   2
         Left            =   420
         TabIndex        =   4
         Top             =   1770
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'APIs for showing the window without activating it:
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal _
   nCmdShow As Long) As Long
Private Const SW_SHOWNOACTIVATE = 4

'APIs for making the window a top-most window:
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal _
   hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
   ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Public Sub ShowME(ByVal vValue As Byte)  'ShowStatus(someinfo As String)
   
   Select Case vValue
    Case 0 'Normal form Loading
        '*** No thing ***
        Show
    Case 1 '
        Show , frmBase
    Case 2
        '*** Calling SetWindowPos API ***
        Show
        SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE _
       Or SWP_NOSIZE Or SWP_NOACTIVATE
    Case 3
        '*** Calling ShowWindow  API ***
        'This code demonstrates loading and showing a form while keeping the focus on the main form.
        ShowWindow hwnd, SW_SHOWNOACTIVATE 'Show the form, but don't activate it.
   End Select
   
 '*** Check this for the Testing with 4th Option ***
'Print "Name : " & frmBase.txtPInfo(0).Text
'Print "Sur Name : " & frmBase.txtPInfo(1).Text
'Print "Address : " & frmBase.txtPInfo(2).Text

txtPInfo(0).Text = frmBase.txtPInfo(0).Text
txtPInfo(1).Text = frmBase.txtPInfo(1).Text
txtPInfo(2).Text = frmBase.txtPInfo(2).Text
   
   
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

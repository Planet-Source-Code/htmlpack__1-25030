VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   4560
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   4335
      Left            =   60
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   4275
      ScaleWidth      =   6795
      TabIndex        =   0
      Top             =   60
      Width           =   6855
      Begin VB.TextBox WWW 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   4340
         Locked          =   -1  'True
         MousePointer    =   10  'Up Arrow
         TabIndex        =   3
         Text            =   "http://rocky.how.to"
         Top             =   3840
         Width           =   1815
      End
      Begin VB.TextBox Mail 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   3370
         Locked          =   -1  'True
         MousePointer    =   10  'Up Arrow
         TabIndex        =   2
         Text            =   "rocky.fff@usa.net"
         Top             =   3600
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   2175
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   1
         Text            =   "Form2.frx":59FE
         Top             =   1920
         Width           =   6735
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function HideCaret Lib "user32" (ByVal hwnd As Long) As Long

Private Sub Command1_Click()
   Unload Me
   Form1.Show
End Sub

Private Sub Mail_Click()
   HideCaret Mail.hwnd
   ShellExecute 0, vbNullString, "mailto:rocky.fff@usa.net?subject=HTMLpack!", vbNullString, "", 1
End Sub

Private Sub Mail_GotFocus()
   HideCaret Mail.hwnd
End Sub

Private Sub WWW_Click()
   HideCaret WWW.hwnd
   ShellExecute 0, vbNullString, "http://rocky.how.to", vbNullString, "", 1
End Sub

Private Sub WWW_GotFocus()
   HideCaret WWW.hwnd
End Sub

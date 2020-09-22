VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HTMLpack!"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3480
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   3480
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1305
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   3255
      Begin VB.CheckBox Blockquotes 
         Caption         =   "Replace <BLOCKQUOTE> tags"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox Spaces2Tab 
         Caption         =   "Replace spaces sequences with TAB"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   200
         Value           =   1  'Checked
         Width           =   3100
      End
      Begin VB.CheckBox Returns 
         Caption         =   "Remove superfluous returns"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   450
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox Strongs 
         Caption         =   "Replace <STRONG> tags"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   710
         Value           =   1  'Checked
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Height          =   450
      Left            =   1800
      TabIndex        =   3
      Top             =   520
      Width           =   1575
      Begin VB.Label HTMLsize 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0 bytes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   900
         TabIndex        =   8
         Top             =   165
         Width           =   600
      End
   End
   Begin VB.CommandButton About 
      Caption         =   "&About"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog openHTML 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton OpenFile 
      Caption         =   "&Open"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Compress 
      Caption         =   "&Compress"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------
'   HTMLpack!
'----------------
' Ultimate HTML Packer
'
' Copyright 2001 by Sangaletti Federico

Private Declare Function CharLower Lib "user32" Alias "CharLowerA" (ByVal lpsz As String) As String
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim HTMLbuffer As String
Dim OriginalSize As Long
Dim CompressedSize As Long

Private Sub About_Click()
   Me.Hide
   Form2.Show
End Sub

Private Sub Compress_Click()
   Dim tmp As String
   
   If HTMLbuffer = vbNullString Then
      MsgBox "Open HTML code first!", vbCritical
      Exit Sub
   End If
   
   'Hold the original size to calculate compression ratio
   OriginalSize = Len(HTMLbuffer)
   
   If Returns.Value = Checked Then
      'Find and replace superfluous vbCrLf
      Do While InStr(1, HTMLbuffer, ">" & vbCrLf)
         HTMLbuffer = Replace(HTMLbuffer, ">" & vbCrLf, ">")
      Loop
   End If
   
   If Spaces2Tab.Value = Checked Then
      'Find and replace with TAB superfluous spaces
      Do While InStr(1, HTMLbuffer, Space(3))
         HTMLbuffer = Replace(HTMLbuffer, Space(3), Chr$(9))
      Loop
   End If
   
   If Strongs.Value = Checked Then
      'Find and replace with <b> all <strong> tags
      tmp = HTMLbuffer
      Do While InStr(1, CharLower(tmp), "<strong>")
         HTMLbuffer = Replace(HTMLbuffer, "<strong>", "<b>")
         HTMLbuffer = Replace(HTMLbuffer, "<STRONG>", "<b>")
         HTMLbuffer = Replace(HTMLbuffer, "<Strong>", "<b>")
         HTMLbuffer = Replace(HTMLbuffer, "</strong>", "</b>")
         HTMLbuffer = Replace(HTMLbuffer, "</STRONG>", "</b>")
         HTMLbuffer = Replace(HTMLbuffer, "</Strong>", "</b>")
         tmp = HTMLbuffer
      Loop
   End If
   
   If Blockquotes.Value = Checked Then
      'Find and replace with <ul> all <blockquote> tags
      tmp = HTMLbuffer
      Do While InStr(1, CharLower(tmp), "<blockquote>")
         HTMLbuffer = Replace(HTMLbuffer, "<blockquote>", "<ul>")
         HTMLbuffer = Replace(HTMLbuffer, "<BLOCKQUOTE>", "<ul>")
         HTMLbuffer = Replace(HTMLbuffer, "<Blockquote>", "<ul>")
         HTMLbuffer = Replace(HTMLbuffer, "</blockquote>", "</ul>")
         HTMLbuffer = Replace(HTMLbuffer, "</BLOCKQUOTE>", "</ul>")
         HTMLbuffer = Replace(HTMLbuffer, "</Blockquote>", "</ul>")
         tmp = HTMLbuffer
      Loop
   End If
   
   CompressedSize = Len(HTMLbuffer)
   HTMLsize.Caption = CompressedSize & " bytes"
   
   If MsgBox(OriginalSize & " => " & CompressedSize & " bytes" & vbCrLf & "Compression ratio: " & 100 - Int((CompressedSize * 100) / OriginalSize) & "%" & vbCrLf & vbCrLf & "Save compressed HTML code?", vbYesNo + vbExclamation, "Done") = vbYes Then
      openHTML.FileName = vbNullString
      openHTML.Filter = "HTML pages|*.htm"
      openHTML.ShowSave
      If openHTML.FileName <> vbNullString Then
         Open openHTML.FileName For Output As #1
            Print #1, HTMLbuffer
         Close #1
         If MsgBox("Test current HTML code?", vbQuestion + vbYesNo) = vbYes Then ShellExecute 0, vbNullString, openHTML.FileName, vbNullString, "", 1
      End If
   End If
   HTMLbuffer = vbNullString
End Sub

Private Sub OpenFile_Click()
   openHTML.Filter = "HTML pages|*.htm"
   openHTML.ShowOpen
   If openHTML.FileName <> vbNullString Then
      Open openHTML.FileName For Binary As #1
         HTMLbuffer = Input(LOF(1), #1)
      Close #1
      HTMLsize.Caption = Len(HTMLbuffer) & " bytes"
   End If
End Sub

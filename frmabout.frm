VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About FTP To"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4095
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4095
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Information"
      Height          =   2415
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3855
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   720
         Left            =   360
         Picture         =   "frmAbout.frx":038A
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   0
         Top             =   480
         Width           =   720
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label1 
         Height          =   1335
         Left            =   2040
         TabIndex        =   4
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Change Log"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      ToolTipText     =   "View change log file"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      ToolTipText     =   "Return to main screen"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Visit website"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      MouseIcon       =   "frmAbout.frx":3084
      MousePointer    =   99  'Custom
      TabIndex        =   3
      ToolTipText     =   "Visit the T-RonX Modding website"
      Top             =   2760
      Width           =   975
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Used for website link on About form.
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    'Update.txt not found.
    If Not Dir$(App.Path & "\changelog.txt") <> "" Then
        MsgBox "Could not find Change Log.", vbInformation
        Exit Sub
    End If
    
    'Else, execute it with notepad.
    Shell ("notepad.exe " & App.Path & "\changelog.txt"), vbNormalFocus
End Sub

Private Sub Form_Load()
    Label1.Caption = "FTP To" & vbCrLf & vbCrLf & "Copyright © 2005" & vbCrLf & "T-RonX Modding" & vbCrLf & "All rights reserved"
    Label2.Caption = "Version: " & FULL_VERSION
End Sub

Private Sub Label3_Click()
   ShellExecute 0&, vbNullString, "http://home.deds.nl/~t-ronx/trm", vbNullString, vbNullString, vbNormalFocus
End Sub

VERSION 5.00
Begin VB.Form frmTransferOK 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transfer Completed"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7710
   Icon            =   "frmTransferOK.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   200
      Picture         =   "frmTransferOK.frx":000C
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   200
      Width           =   480
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   1200
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK (3)"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Caption         =   "Label1"
      Height          =   975
      Left            =   960
      TabIndex        =   1
      Top             =   195
      Width           =   6615
   End
End
Attribute VB_Name = "frmTransferOK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long

Dim Counter As Integer

Private Sub btnOK_Click()
    EndAll
End Sub

Private Sub Form_Load()
    'Set variables.
    lblInfo.Caption = "File transfer successful." & vbCrLf & vbCrLf & "Destination server: " & FTPAddress & vbCrLf & "Local file: " & Chr(39) & LocalFile & Chr(39)
    Counter = 10 '10 seconds before quit.
    btnOK.Caption = "OK (" & Counter & ")"
    Timer1.Enabled = True
    MessageBeep vbInformation
    'PlaySound "info.wav", 0, 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Red cross.
    EndAll
End Sub

Private Sub Timer1_Timer()
    'Count down.
    Counter = Counter - 1
    btnOK.Caption = "OK (" & Counter & ")"
    If Counter <= 0 Then EndAll
End Sub

Private Sub EndAll()
    'Quits the program.
    Unload Me
    End
End Sub

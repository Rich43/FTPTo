VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTransferState 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Starting transfer..."
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4485
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Form1.frx":038A
   ScaleHeight     =   87
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   299
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3360
      TabIndex        =   1
      ToolTipText     =   "Stop transferring"
      Top             =   960
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   180
      Left            =   45
      TabIndex        =   0
      ToolTipText     =   "Transfer progress"
      Top             =   345
      Width           =   4400
      _ExtentX        =   7752
      _ExtentY        =   318
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblTotalSent 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Sent:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblTimeLeft 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label lblSpeed 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Time Left:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Caption         =   "File: "
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   75
      Width           =   4215
   End
   Begin VB.Label lblTransferRate 
      BackStyle       =   0  'Transparent
      Caption         =   "Transfer Rate:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   585
      Width           =   1215
   End
End
Attribute VB_Name = "frmTransferState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    CancelTransfer = True
End Sub

Private Sub Form_Load()
    Me.lblTimeLeft.ForeColor = vbRed
    Me.lblSpeed.ForeColor = vbRed
    Me.lblTotalSent.ForeColor = vbRed
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1

Select Case MsgBox("Are you sure you want to stop uploading " & RemoteFile & "?", vbYesNo + vbQuestion)
    Case vbYes
        CancelTransfer = True
End Select

End Sub

Public Function ShowTransferOk()
    frmTransferOK.Show vbModal, Me
End Function

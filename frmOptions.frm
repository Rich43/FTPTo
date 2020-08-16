VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4095
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   208
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   273
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3855
      Begin VB.TextBox txtMinKB 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   1190
         Width           =   855
      End
      Begin VB.CheckBox chkCheckForUpdates 
         Caption         =   "Check for updates automatically"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Check for updates when the manager tool starts"
         Top             =   600
         Width           =   3255
      End
      Begin VB.CheckBox Integrate 
         Caption         =   "Integrate FTP To in context menu"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "This will add 'FTP To' to the context menu of any file, this allows you to easlie send files to the FTP's you configured"
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "Kilo Byte."
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Do not show transfer progress if file is smaller than:"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Makes the progress bar only show up while transferring larger files"
         Top             =   960
         Width           =   3615
      End
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      ToolTipText     =   "Return to main screen"
      Top             =   2640
      Width           =   975
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnOK_Click()
    If IsNumeric(txtMinKB.Text) = True Then
        Me.Hide
    Else
        MsgBox "Enter a number in the Kilo Byte text box.", vbExclamation
    End If
End Sub


VERSION 5.00
Begin VB.Form frmServerList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FTP To"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4335
   Icon            =   "frmServerList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkCopyURL 
      Caption         =   "Copy URL to clipboard"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Copies the text you specified to the clipboard"
      Top             =   4200
      Width           =   3615
   End
   Begin VB.CommandButton btnManage 
      Caption         =   "Manage servers"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      ToolTipText     =   "Start FTP To Manager"
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      ToolTipText     =   "Send the file to the selected destination server"
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      ToolTipText     =   "Cancel file sending"
      Top             =   4440
      Width           =   975
   End
   Begin VB.ListBox ServerList 
      Height          =   3765
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "List of available servers"
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Select the destination server:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmServerList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
    End
End Sub

Private Sub btnManage_Click()
    'Shows config tool.
    frmConfig.Show
    ManageFromSrvList = True
    Unload Me
End Sub

Private Sub btnOK_Click()
    ContinueFileSend
End Sub

Private Sub Form_Load()
    'loads server list.
    For i = 1 To UBound(Server)
        ServerList.AddItem Server(i).Name
    Next
    
    If ServerList.ListCount > 0 Then ServerList.ListIndex = 0
End Sub

Private Sub ServerList_DblClick()
    'Continue if something is selected.
    If ServerList.ListIndex >= 0 Then ContinueFileSend
End Sub

Public Sub ContinueFileSend()

'Checks value of checkbox.
Select Case chkCopyURL.Value
    Case 1: bCopyURL = True
    Case 0: bCopyURL = False
End Select
    
Me.Hide

    'Loops through Server array.
    For i = 1 To UBound(Server)
        'Matching name found.
        If Server(i).Name = ServerList.List(ServerList.ListIndex) Then
            'Updates variables.
            FTPAddress = Server(i).Address
            FTPServerPort = Server(i).Port
            FTPUserName = Server(i).Username
            FTPDirectory = Server(i).Directory
            FTPPassword = Server(i).Password
            sCopyURL = Server(i).CopyURL
            If Dir$(FileToSend) <> "" Then SendFile
        End If
    Next
    
    Unload Me
    Exit Sub
    
End Sub


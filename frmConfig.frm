VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FTP To Manager"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   5775
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDefCopyURL 
      Caption         =   "Copy URL to clipboard"
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      ToolTipText     =   "Copy URL function for default server"
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox txtCopyURL 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CheckBox chkDefServer 
      Caption         =   "Always use this server"
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      ToolTipText     =   "Check this option if you always want to use this server"
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox txtFTPDir 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   3360
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox FTPAddress 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox Username 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox Password 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      ToolTipText     =   "Save settings and exit"
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   10
      ToolTipText     =   "Quit without saving changes"
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox ServerPort 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.ListBox ServerList 
      Height          =   2595
      Left            =   3600
      Sorted          =   -1  'True
      TabIndex        =   8
      ToolTipText     =   "List of configured servers"
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label defSrv 
      Caption         =   "Default Server:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      ToolTipText     =   "This shows the current default server (server with the 'Always use this server' check enabled)"
      Top             =   3120
      Width           =   3375
   End
   Begin VB.Label Label5 
      Caption         =   "Clipboard URL:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      ToolTipText     =   "Text to be copied to the clipboard after uploading a file"
      Top             =   1950
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "FTP Directory:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      ToolTipText     =   "The file you upload will be stored in this directory"
      Top             =   510
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "FTP Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "Location of the server (example: ftp.mywebsite.com)"
      Top             =   150
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      ToolTipText     =   "Enter your password here"
      Top             =   1590
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Username:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "Enter you username here"
      Top             =   1230
      Width           =   1215
   End
   Begin VB.Label FTPPort 
      Caption         =   "Server Port:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      ToolTipText     =   "This is the port the server is listening on, usually 21"
      Top             =   870
      Width           =   1215
   End
   Begin VB.Menu FileMenu 
      Caption         =   "File"
      Begin VB.Menu FileExit 
         Caption         =   "Save and Exit"
      End
   End
   Begin VB.Menu MenuEdit 
      Caption         =   "Edit"
      Begin VB.Menu EditAddServer 
         Caption         =   "Add server..."
      End
      Begin VB.Menu EditEditServername 
         Caption         =   "Rename server..."
      End
      Begin VB.Menu EditDeleteServer 
         Caption         =   "Delete server..."
      End
      Begin VB.Menu EditDuplicateServer 
         Caption         =   "Duplicate server..."
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu EditOptions 
         Caption         =   "Options..."
      End
   End
   Begin VB.Menu MenuHelp 
      Caption         =   "Help"
      Begin VB.Menu HelpFile 
         Caption         =   "Help file..."
      End
      Begin VB.Menu HelpAbout 
         Caption         =   "About..."
      End
      Begin VB.Menu HelpUpdate 
         Caption         =   "Check for updates..."
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Used for website link on About form.
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Used for autoupdate.
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

'Init Reg edit class.
Dim Reg As New RegistryFunctions

'Variable for 'default' check box, needed here else it will be reset.
Dim Def As Integer

Private Sub btnCancel_Click()
    ResumeServerList Command$ 'Show server list if Manage Servers has been clicked
    End 'No saving on Cancel button.
End Sub

Private Sub btnOK_Click()
    SaveConfig
    ResumeServerList Command$ 'Show server list if Manage Servers has been clicked
    End
End Sub

Private Sub EditDuplicateServer_Click()
Dim NewName As String

If ServerList.ListIndex = -1 Then MsgBox "Please select the server you want to duplicate first.", vbExclamation: Exit Sub
    'Finds right item from list.
    For i = 1 To UBound(Server)
        'Found the selected one from list.
        If Server(i).Name = ServerList.List(ServerList.ListIndex) Then
            'Enter new name.
            NewName = InputBox("Enter a new server name...", , Server(i).Name)
            
            'Blank name not allowed.
            If Trim(NewName) = "" Then Exit Sub
            
                'check if new name already exits.
                For n = 1 To UBound(Server)
                    If Server(n).Name = Trim(NewName) Then
                    MsgBox "This item already exists. Please enter an other name.", vbExclamation
                    EditDuplicateServer_Click
                    Exit Sub
                    End If
                Next n

            'Assign values to new array item.
            ReDim Preserve Server(0 To (UBound(Server) + 1))
            Server(UBound(Server)).Address = Server(i).Address
            Server(UBound(Server)).CopyURL = Server(i).CopyURL
            Server(UBound(Server)).Directory = Server(i).Directory
            Server(UBound(Server)).Password = Server(i).Password
            Server(UBound(Server)).Port = Server(i).Port
            Server(UBound(Server)).Username = Server(i).Username
            Server(UBound(Server)).Name = NewName
            ServerList.AddItem NewName
        End If
    Next i


End Sub

Private Sub EditEditServername_Click()
Dim NewName As String
'Checks of something is selected.
If ServerList.ListIndex = -1 Then MsgBox "Select an item from the list first.", vbExclamation: Exit Sub

    'Finds right item from list.
    For i = 1 To UBound(Server)
        'Found the selected one from list.
        If Server(i).Name = ServerList.List(ServerList.ListIndex) Then
            'Enter new name.
            NewName = InputBox("Enter a new server name...", , Server(i).Name)

            'Blank name not allowed.
            If Trim(NewName) = "" Then Exit Sub
            
                'check if new name already exits.
                For n = 1 To UBound(Server)
                    If Server(n).Name = Trim(NewName) Then
                    MsgBox "This item already exists. Please enter an other name.", vbExclamation
                    EditEditServername_Click
                    Exit Sub
                    End If
                Next n
            
            'Update variables with new name.
            ServerList.AddItem NewName
            ServerList.RemoveItem ServerList.ListIndex
            Server(i).Name = NewName
            ServerList.ListIndex = 0
            Exit Sub
        End If
    Next i
    
End Sub

Private Sub EditOptions_Click()
    frmOptions.Show vbModal
End Sub

Private Sub FileExit_Click()
    SaveConfig
    ResumeServerList Command$ 'Show server list if Manage Servers has been clicked
    End
End Sub

Private Sub EditAddServer_Click()
'Box to enter the name.
Dim NewItem As String
NewItem = InputBox("Enter the new item name...")

'Duplicated names are not allowed.
For i = 1 To UBound(Server)
    If Server(i).Name = Trim(NewItem) Then
    MsgBox "This item already exists. Please enter an other name.", vbExclamation
    EditAddServer_Click
    Exit Sub
    End If
Next

    If Not Trim(NewItem) = "" Then 'Blank name is not allowed.
        ReDim Preserve Server(UBound(Server) + 1) 'Increase the array to make room for the new server info.
        Server(UBound(Server)).Name = NewItem
        ServerList.AddItem NewItem
        If ServerList.ListCount > 0 Then ServerList.ListIndex = 0 'Selects first item in the list.
    End If

'Enables or disables the textboxes.
CheckListBox

End Sub

Private Sub EditDeleteServer_Click()

'Prompt.
If ServerList.ListIndex = -1 Then
    MsgBox "Select an item from the list first.", vbExclamation
Else
    Dim YN As Integer
    YN = MsgBox("Are you sure you want to delete " & ServerList.List(ServerList.ListIndex) & "?", vbYesNo + vbQuestion)
End If

If YN = vbYes Then
    'Loops through array.
    For i = 1 To UBound(Server)
        If Server(i).Name = ServerList.List(ServerList.ListIndex) Then 'Found the selected one.
            Server(i).Name = vbNullString 'Blank names will not be written to output file.
            ServerList.RemoveItem ServerList.ListIndex
            
            If ServerList.ListCount > 0 Then
                ServerList.ListIndex = 0 'Selects first item in the list.
            Else
                'Clean text boxes of no item exists.
                Password.Text = vbNullString
                FTPAddress.Text = vbNullString
                txtFTPDir.Text = vbNullString
                Username.Text = vbNullString
                ServerPort.Text = vbNullString
                txtCopyURL.Text = vbNullString
                chkDefServer.Value = 0
            End If
            
            Exit For
        End If
    Next
End If

'Enables or disables the textboxes.
CheckListBox
chkCheck
End Sub

Private Sub Form_Load()
Me.Caption = "FTP To Manager - " & APP_VERSION
Load frmOptions

    'Sub Main has already loaded the config file, adds item here.
    For i = 1 To UBound(Server)
        ServerList.AddItem Server(i).Name
    Next

'Enables or disables the textboxes.
CheckListBox
chkCheck

    frmOptions.Integrate.Value = IntegrateInContextMenu
    frmOptions.chkCheckForUpdates.Value = CheckForUpdatesAutomatically
    frmOptions.txtMinKB.Text = TransferProgressMinKB
    If ServerList.ListCount > 0 Then ServerList.ListIndex = 0

UpdateDefServerLabel
Me.Show
If CheckForUpdatesAutomatically = 1 Then CheckForUpdates False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload frmConfig
    Unload frmOptions
    Unload frmAbout
    Unload frmServerList
End Sub

Private Sub HelpAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub HelpFile_Click()
    'Help.rtf not found.
    If Not Dir$(App.Path & "\help.rtf") <> "" Then
        MsgBox "Could not find help file.", vbInformation
        Exit Sub
    End If

    ShellExecute 0&, vbNullString, App.Path & "\help.rtf", vbNullString, vbNullString, vbNormalFocus
End Sub

Private Sub HelpUpdate_Click()
    CheckForUpdates True
End Sub

Private Sub ServerList_Click()
   
    For i = 1 To UBound(Server)
        If Server(i).Name = ServerList.List(ServerList.ListIndex) Then 'Finds matching name in the array.
            'Updates GUI.
            FTPAddress.Text = Server(i).Address
            txtFTPDir.Text = Server(i).Directory
            ServerPort.Text = Server(i).Port: If Server(i).Port = vbNullString Then ServerPort.Text = "21"
            Username.Text = Server(i).Username
            Password.Text = Server(i).Password
            txtCopyURL.Text = Server(i).CopyURL
            chkDefServer.Value = Server(i).Default
            chkCheck
            If chkDefServer.Value = 1 Then chkDefCopyURL.Value = Server(i).DefCopyURL
            If chkDefServer.Value = 0 Then chkDefCopyURL.Value = 0
            Exit For
        End If
    Next
    
UpdateDefServerLabel

End Sub

Private Sub chkDefServer_Click()

    'Loop through server array.
    For i = 1 To UBound(Server)
        'Name of selected item is i.
        If Server(i).Name = ServerList.List(ServerList.ListIndex) Then
            'Selected item now has value of checkbox.
            Server(i).Default = chkDefServer.Value
            'Igonre this item in next loop.
            If chkDefServer.Value Then Def = i
            
            'Unchecks URL checkbox
            If chkDefServer.Value = 0 Then chkDefCopyURL.Value = 0
        End If
    Next
    
    For i = 1 To UBound(Server)
        'We can only select 1 default server, so set the rest to 0 except latest selected selected one.
        If Not i = Def Then Server(i).Default = 0
    Next
    
'Enables or disables the 'copy url' checkbox
chkCheck
UpdateDefServerLabel
End Sub

Private Sub chkDefCopyURL_Click()
    'Loop through server array.
    For i = 1 To UBound(Server)
        'Name of selected item is i.
        If Server(i).Name = ServerList.List(ServerList.ListIndex) Then
            'Selected item now has value of checkbox.
            Server(i).DefCopyURL = chkDefCopyURL.Value
            Exit For
        End If
    Next
End Sub

Private Sub ServerPort_Change()
    'Updates memory immediately after after typing. Easier for handing OK button.
    For i = 1 To UBound(Server)
        If Server(i).Name = ServerList.List(ServerList.ListIndex) Then 'Compares array .Name with selected list item.
            Server(i).Port = ServerPort.Text
            Exit For
        End If
    Next
End Sub

Private Sub txtCopyURL_LostFocus()
    'Send warning if ther is no %filename% in textbox.
    If Len(Trim(txtCopyURL.Text)) > 0 And InStr(1, txtCopyURL.Text, "%filename%", vbTextCompare) = 0 Then MsgBox "The 'Clipboard URL' text box does not contain %filename%, the filename will not be copied to the clipboard.", vbExclamation
    ServerList_Click
End Sub

Private Sub txtFTPDir_Change()
    'Updates memory immediately after after typing. Easyer for handing OK button.
    For i = 1 To UBound(Server)
        If Server(i).Name = ServerList.List(ServerList.ListIndex) Then 'Compares array .Name with selected list item.
            Server(i).Directory = txtFTPDir.Text
            Exit For
        End If
    Next
End Sub

Private Sub txtCopyURL_Change()
    'Updates memory immediately after after typing. Easyer for handing OK button.
    For i = 1 To UBound(Server)
        If Server(i).Name = ServerList.List(ServerList.ListIndex) Then 'Compares array .Name with selected list item.
            Server(i).CopyURL = txtCopyURL.Text
            Exit For
        End If
    Next
End Sub

Private Sub Username_Change()
    'Updates memory immediately after after typing. Easier for handing OK button.
    For i = 1 To UBound(Server)
        If Server(i).Name = ServerList.List(ServerList.ListIndex) Then 'Compares array .Name with selected list item.
            Server(i).Username = Username.Text
            Exit For
        End If
    Next
End Sub

Private Sub FTPAddress_Change()
    'Updates memory immediately after after typing. Easier for handing OK button.
    For i = 1 To UBound(Server)
        If Server(i).Name = ServerList.List(ServerList.ListIndex) Then 'Compares array .Name with selected list item.
            Server(i).Address = FTPAddress.Text
            Exit For
        End If
    Next
End Sub

Private Sub Password_Change()
    'Updates memory immediately after after typing. Easyer for handing OK button.
    For i = 1 To UBound(Server)
        If Server(i).Name = ServerList.List(ServerList.ListIndex) Then 'Compares array .Name with selected list item.
            Server(i).Password = Password.Text
            Exit For
        End If
    Next
End Sub

Public Sub SaveConfig()
On Error GoTo ErrHand

    Open App.Path & "\config.ini" For Output As #1
    
        For i = 1 To UBound(Server)
            If Not Server(i).Name = vbNullString Then 'Deleted items will be skipped.
                'Writes config file.
                Print #1, "ServerName = " & Server(i).Name
                Print #1, "FTPAddress = " & Server(i).Address
                Print #1, "FTPDirectory = " & Server(i).Directory
                Print #1, "FTPServerPort = " & Server(i).Port
                Print #1, "FTPUserName = " & Server(i).Username
                Print #1, "FTPPassword = " & Encryption.StrEncrypt(Server(i).Password)
                Print #1, "CopyURL = " & Server(i).CopyURL
                Print #1, "DefCopyURL = " & Server(i).DefCopyURL
                Print #1, "IsDefault = " & Server(i).Default
                Print #1, vbCr
            End If
        Next
        
        Print #1, "IntegrateInContextMenu = " & frmOptions.Integrate.Value
        Print #1, "CheckForUpdatesAutomatically = " & frmOptions.chkCheckForUpdates.Value
        Print #1, "TransferProgressMinKB = " & frmOptions.txtMinKB.Text
        
    Close #1

'Add or remove reg key, depending on Integrate value.
Select Case frmOptions.Integrate.Value
    Case 0
        'Remove reg keys.
        Reg.DeleteKey "HKEY_CLASSES_ROOT\*\shell\FTP To...\command"
        Reg.DeleteKey "HKEY_CLASSES_ROOT\*\shell\FTP To..."
    Case 1
        'Add reg keys.
        Reg.CreateKey "HKEY_CLASSES_ROOT\*\shell\FTP To...\command"
        Reg.SetStringValue "HKEY_CLASSES_ROOT\*\shell\FTP To...\command", "", Chr(34) & App.Path & "\ftpto.exe" & Chr(34) & " %1"
End Select

Exit Sub

ErrHand:
    'File already open.
    If Err.Number = 55 Then
        Close #1
        SaveConfig
    End If
        
End Sub

Public Sub CheckListBox()

    'This checks if there are objects in the list box.
    Select Case ServerList.ListCount
        Case 0 'No objects.
            'Makes you unable to type in the textboxes.
            Password.Enabled = False
            ServerPort.Enabled = False
            txtFTPDir.Enabled = False
            FTPAddress.Enabled = False
            Username.Enabled = False
            chkDefServer.Enabled = False
            txtCopyURL.Enabled = False
            
            'Changes colors of textboxes.
            Password.BackColor = &HF0F0F0
            ServerPort.BackColor = &HF0F0F0
            txtFTPDir.BackColor = &HF0F0F0
            FTPAddress.BackColor = &HF0F0F0
            Username.BackColor = &HF0F0F0
            txtCopyURL.BackColor = &HF0F0F0

        Case Else 'Objects in listbox.
            'Allows typing.
            Password.Enabled = True
            ServerPort.Enabled = True
            txtFTPDir.Enabled = True
            FTPAddress.Enabled = True
            Username.Enabled = True
            chkDefServer.Enabled = True
            txtCopyURL.Enabled = True
            
            'Change color to white.
            Password.BackColor = &H80000005
            ServerPort.BackColor = &H80000005
            txtFTPDir.BackColor = &H80000005
            FTPAddress.BackColor = &H80000005
            Username.BackColor = &H80000005
            txtCopyURL.BackColor = &H80000005
    End Select

End Sub

Public Function chkCheck()

'Enables or disables the 'copy url' checkbox
Select Case chkDefServer.Value
    Case 1
        chkDefCopyURL.Enabled = True
        
    Case 0
        chkDefCopyURL.Enabled = False
End Select
                
End Function

'Resume file sending.
Public Function ResumeServerList(Path As String)
    If ManageFromSrvList = True Then Shell (App.Path & "\ftpto.exe " & Path), vbNormalFocus
End Function

Public Function CheckForUpdates(Prompt As Boolean)
Dim Version As String 'Holds new version number.
Dim YesNo As Integer

'Downlaod the version file.
DownloadFile "http://home.deds.nl/~t-ronx/trm/version/ftpto_version.txt", App.Path & "\ftpto_version.txt"

'File could not be dowloaded.
If Not Dir(App.Path & "\ftpto_version.txt") <> "" Then
    If Prompt = True Then MsgBox "FTP To was unable to check for a newer version.", vbCritical: Exit Function
    Exit Function
End If

'Read first line of the just downloaded file.
Open App.Path & "\ftpto_version.txt" For Input As #1
    'Load Version variable with new version number.
    Line Input #1, Version
Close #1
    
'Remove the version file.
Kill (App.Path & "\ftpto_version.txt")

'What to do next.
Select Case CLng(Version)

    'Web version is newer.
    Case Is > UPDATE_VERSION
        'Prompt if you wanna download and install.
        YesNo = MsgBox("A newer version of FTP To is available, do you want to download and install it now?", vbInformation + vbYesNo)
        
            'Yes we want to install.
            If YesNo = vbYes Then
            
                'Update GUI.
                frmDownloadingNewVersion.Show
                Me.Refresh
                frmDownloadingNewVersion.Refresh
                DoEvents
                
                'Download the new file.
                DownloadFile "http://home.deds.nl/~t-ronx/trm/files/ftpto" & Version & ".exe", App.Path & "\ftpto" & Version & ".exe"
                
                Unload frmDownloadingNewVersion
                
                    'Run the file if it exists.
                    If Dir(App.Path & "\ftpto" & Version & ".exe") <> "" Then
                        Shell (App.Path & "\ftpto" & Version & ".exe"), vbNormalFocus
                        End 'Quit, so I can be overwritten.
                    Else
                        'The file has not been downloaded.
                        MsgBox "File could not be downloaded successfully.", vbCritical
                    End If
            End If
    
    'Web version is same as current app version.
    Case Is = UPDATE_VERSION
        If Prompt = True Then MsgBox "No update available. You already have the latest version.", vbInformation
    
    'Your version is newer, hmm.
    Case Is < UPDATE_VERSION
        If Prompt = True Then MsgBox "You already have a newer version than available on the internet.", vbInformation

End Select
    
End Function

'Download the file. 'API Guide'
Public Function DownloadFile(URL As String, LocalFilename As String) As Boolean
    Dim lngRetVal As Long
    lngRetVal = URLDownloadToFile(0, URL, LocalFilename, 0, 0)
    If lngRetVal = 0 Then DownloadFile = True
End Function

Public Function UpdateDefServerLabel()
Dim NewValue As String

    'Checks for a default server.
    For i = 1 To UBound(Server)
        If Server(i).Default = 1 Then
            NewValue = Server(i).Name
            Exit For
        Else
            NewValue = "none"
        End If
    Next

'Updates label.
defSrv.Caption = "Default Server: " & NewValue
End Function








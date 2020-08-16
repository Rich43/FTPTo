Attribute VB_Name = "modMain"
'Thanks API Guide
'Constants of APIs.
Const FTP_TRANSFER_TYPE_UNKNOWN = &H0
Const FTP_TRANSFER_TYPE_ASCII = &H1
Const FTP_TRANSFER_TYPE_BINARY = &H2
Const INTERNET_DEFAULT_FTP_PORT = 21               ' default for FTP servers
Const INTERNET_SERVICE_FTP = 1
Const INTERNET_FLAG_PASSIVE = &H8000000            ' used for FTP connections
Const INTERNET_OPEN_TYPE_PRECONFIG = 0                    ' use registry configuration
Const INTERNET_OPEN_TYPE_DIRECT = 1                        ' direct to net
Const INTERNET_OPEN_TYPE_PROXY = 3                         ' via named proxy
Const INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY = 4   ' prevent using java/script/INS
Const MAX_PATH = 260
Const GENERIC_WRITE = &H40000000
Const BUFFERSIZE = 255

Public Const APP_VERSION = "1.0.9"
Public Const UPDATE_VERSION = "10903"
Public Const FULL_VERSION = "1.0.9.03"

'API declarations.
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hConnect As Long, ByVal lpszLocalFile As String, ByVal lpszNewRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Private Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Private Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Boolean
Private Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (lpdwError As Long, ByVal lpszBuffer As String, lpdwBufferLength As Long) As Boolean
Private Declare Function FtpOpenFile Lib "wininet.dll" Alias "FtpOpenFileA" (ByVal hFtpSession As Long, ByVal sBuff As String, ByVal Access As Long, ByVal Flags As Long, ByVal Context As Long) As Long
Private Declare Function InternetWriteFile Lib "wininet.dll" (ByVal hFile As Long, ByRef sBuffer As Byte, ByVal lNumBytesToWrite As Long, dwNumberOfBytesWritten As Long) As Integer

'For XP Style.
Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long

'Key values.
Public FTPAddress As String
Public FTPServerPort As String
Public FTPUserName As String
Public FTPPassword As String
Public FTPDirectory As String
Public sCopyURL As String
Public IntegrateInContextMenu As Integer
Public CheckForUpdatesAutomatically As Integer
Public TransferProgressMinKB As Long
'This is the type that holds all server information.

Public Type ServerType
    Name As String
    Address As String
    Directory As String
    Port As String
    Username As String
    Password As String
    Default As Integer
    DefCopyURL As Integer
    CopyURL As String
End Type
Public Server() As ServerType

Private Enum Com
    InetOpen
    InetConnect
    CreateDirectory
    SetCurrentDirectory
    DeleteFile
    PutFile
    InetCloseHandle
End Enum

Public FileToSend As String 'Holds Command$.
Public Encryption As New clsEncryption

'Holds True of the Manage Servers button is clicked
Public ManageFromSrvList As Boolean

Dim hConnection As Long 'Here because of ShowError()
Dim hOpen As Long, hFiles As Long

'User For ShowError, declared here because it moves in and outside the procedure.
Dim CreateErr As Integer
Dim TransferOK As Integer
Dim DirCreated As Boolean
Dim NewTimer As Date
Dim TotalSent As Double
Dim TotalSentString As String
Dim Written As Long
Dim FileSize As Long
Dim Sum As Long
Dim BeginTransfer As Single
Dim TransferRate As Single
Dim ShowProgress As Boolean
Public CancelTransfer As Boolean
Public LocalFile As String
Public RemoteFile As String
Public bCopyURL As Boolean

Private Sub Main()
'Initializes XP Style.
Call InitCommonControls

'This is when a file has been dragged over the exe file, removes the "..." of %1.
If Left(Command$, 1) = Chr(34) And Right(Command$, 1) = Chr(34) Then
    FileToSend = Mid(Command$, 2, (Len(Command$) - 2))
Else
    FileToSend = Command$  '"D:\15 - Kill Ride Medley.mp3" '
End If

ReDim Preserve Server(0) 'Array must be set!

'Checks if cfg file exists.
If Not Dir$(App.Path & "\config.ini") <> "" Then
    'Show config tool if cfg file doesn't exist.
    IntegrateInContextMenu = 1
    CheckForUpdatesAutomatically = 1
    TransferProgressMinKB = 64
    frmConfig.Show
Else
    'Load it if it exists.
    LoadConfig
    
    'This happends when config file doesn't contain any server info.
    If UBound(Server) = 0 Then
        IntegrateInContextMenu = 1
        CheckForUpdatesAutomatically = 1
        TransferProgressMinKB = 64
        frmConfig.Show
        Exit Sub
    End If
    
    'If Command is "" then the exe file has been executed manually so we show the config tool.
    If FileToSend = vbNullString Then
        frmConfig.Show
    Else
            
            'Loads main variables.
            For i = 1 To UBound(Server)
                If Server(i).Default = 1 Then
                    FTPAddress = Server(i).Address
                    FTPServerPort = Server(i).Port
                    FTPDirectory = Replace(Server(i).Directory, "\", "/")
                    FTPUserName = Server(i).Username
                    FTPPassword = Server(i).Password
                    sCopyURL = Server(i).CopyURL
                    bCopyURL = Server(i).DefCopyURL
                    
                    If Dir$(FileToSend) <> "" Then SendFile  'Else we just send the file, if it exists.
                    Exit For
                Else
                    'Handles multiple servers without one set ot default.
                    If i = UBound(Server) Then
                        ManageFromSrvList = True 'Makes server select list show up again afer config tool coloses.
                        frmServerList.Show 'Only show form if the entire array has been checked.
                    End If
                End If
             Next

    End If
End If

End Sub

Public Sub SendFile()
On Error GoTo ErrHand

Dim sOrgPath  As String
Dim RemoteFileSplit() As String
Dim SizeKB As Long

'Get file name from Command$.
LocalFile = FileToSend
RemoteFileSplit = Split(LocalFile, "\")
RemoteFile = Right(LocalFile, Len(RemoteFileSplit(UBound(RemoteFileSplit))))

    'Open the internet connection.
    hOpen = InternetOpen("FTP To", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    ShowError InetOpen

    'Connect to the FTP server.
    hConnection = InternetConnect(hOpen, FTPAddress, FTPServerPort, FTPUserName, FTPPassword, INTERNET_SERVICE_FTP, IIf(PassiveConnection, INTERNET_FLAG_PASSIVE, 0), 0)
    ShowError InetConnect
    
    'Open the FTP directory.
    FtpSetCurrentDirectory hConnection, FTPDirectory
    ShowError SetCurrentDirectory
    
    'Determs if progress bar will be shown or no.
    Open LocalFile For Binary Access Read As #1
        SizeKB = LOF(1) / 1024
        Select Case SizeKB
            Case Is >= TransferProgressMinKB
                ShowProgress = True
                
            Case Is < TransferProgressMinKB
                ShowProgress = False
        End Select
    Close #1

    'Uploads the file...
    'If CreateErr = 0 Then FtpPutFile hConnection, LocalFile, RemoteFile, FTP_TRANSFER_TYPE_UNKNOWN, 0
    If CreateErr = 0 Then If UploadFile = True Then TransferOK = 1
    'If CreateErr = 0 Then ShowError PutFile
    
    'Close the internet connection.
    InternetCloseHandle hOpen
    ShowError InetCloseHandle

        'Sends message to user.
    If TransferOK Then
        
        'Checks if url has to be copied.
        If bCopyURL = True Then

            If Len(Trim(sCopyURL)) > 0 Then
                Clipboard.Clear
                Clipboard.SetText Replace(Trim(Replace(sCopyURL, "%filename%", RemoteFile, , , vbTextCompare)), " ", "%20")
            Else
                If Len(Trim(FTPDirectory)) > 0 Then
                    'Replace \ by /.
                    FTPDirectory = Replace(FTPDirectory, "\", "/")
                    'Adds /../ to sting.
                    If Not Left(FTPDirectory, 1) = "/" Then FTPDirectory = "/" & FTPDirectory
                    If Not Right(FTPDirectory, 1) = "/" Then FTPDirectory = FTPDirectory & "/"
                Else
                    FTPDirectory = "/"
                End If
                    'Sets to clipboard.
                    Clipboard.Clear
                    Clipboard.SetText "ftp://" & FTPAddress & Replace(FTPDirectory, " ", "%20") & Replace(RemoteFile, " ", "%20")
            End If
            
        End If
        
        'Success message.
        frmTransferState.ShowTransferOk
       ' frmTransferOK.Show vbModal
        'MsgBox "File transfer successful." & vbCrLf & vbCrLf & "Destination server: " & FTPAddress & vbCrLf & "Local file: " & Chr(39) & LocalFile & Chr(39), vbInformation, "Transfer Completed"
    Else
        'Error.
        MsgBox "File transfer could not be completed successfully. Error while uploading file." & vbCrLf & "File: " & Chr(39) & LocalFile & Chr(39) & "." & vbCrLf & "Please try again.", vbExclamation
        End
    End If
    
    CreateErr = 0 'Reset folder create error
    TransferOK = 0
    CancelTransfer = False

Exit Sub

ErrHand:
    MsgBox "An error occured, the server settngs may be wrong or incomplete. Check the settings and try again.", vbExclamation
    frmConfig.Show

End Sub

Public Function LoadConfig()
On Error GoTo ErrHand

Dim LineString As String
Dim LineSplit() As String

    'Opens the config file.
    Open App.Path & "\config.ini" For Input As #1
        'Loop through the end of file.
        Do Until EOF(1)
            Line Input #1, LineString

                'This line is a flag.
                If InStr(1, LineString, "=") Then
                    LineSplit = Split(LineString, "=") 'Gets left and right side of the = sign.
    
                        'If a new server block is found, increase the array.
                        If Trim(LineSplit(0)) = "ServerName" Then
                             ReDim Preserve Server(0 To (UBound(Server) + 1))
                             Server(UBound(Server)).Name = Trim(LineSplit(1))
                        End If

                        'Automaticly sets right values to the right array.
                        If Trim(LineSplit(0)) = "FTPAddress" Then Server(UBound(Server)).Address = Trim(LineSplit(1))
                        If Trim(LineSplit(0)) = "FTPServerPort" Then Server(UBound(Server)).Port = Trim(LineSplit(1))
                        If Trim(LineSplit(0)) = "FTPUserName" Then Server(UBound(Server)).Username = Trim(LineSplit(1))
                        If Trim(LineSplit(0)) = "FTPDirectory" Then Server(UBound(Server)).Directory = Trim(LineSplit(1))
                        If Trim(LineSplit(0)) = "FTPPassword" Then Server(UBound(Server)).Password = Encryption.StrDecrypt(Trim(LineSplit(1)))
                        If Trim(LineSplit(0)) = "IntegrateInContextMenu" Then IntegrateInContextMenu = CInt(Trim(LineSplit(1)))
                        If Trim(LineSplit(0)) = "CheckForUpdatesAutomatically" Then CheckForUpdatesAutomatically = CInt(Trim(LineSplit(1)))
                        If Trim(LineSplit(0)) = "IsDefault" Then Server(UBound(Server)).Default = CInt(Trim(LineSplit(1)))
                        If Trim(LineSplit(0)) = "CopyURL" Then Server(UBound(Server)).CopyURL = Trim(LineSplit(1))
                        If Trim(LineSplit(0)) = "DefCopyURL" Then Server(UBound(Server)).DefCopyURL = Trim(LineSplit(1))
                        If Trim(LineSplit(0)) = "TransferProgressMinKB" Then TransferProgressMinKB = Trim(LineSplit(1))
                
                End If
        Loop
    Close #1
    
Exit Function
ErrHand:
    'File already open.
    If Err.Number = 55 Then
        Close #1
        LoadConfig
    Else
        MsgBox "An error occured while loading the config file. You may lose data if you continue.", vbExclamation
    End If
       
End Function

Private Sub ShowError(ConnectionType As Com)
Dim lErr As Long, sErr As String, lenBuf As Long

    'Get the required buffer size.
    InternetGetLastResponseInfo lErr, sErr, lenBuf
    
    'Create a buffer.
    sErr = String(lenBuf, 0)
    
    'Retrieve the last respons info.
    InternetGetLastResponseInfo lErr, sErr, lenBuf
    
    Select Case ConnectionType
        Case Com.InetOpen
            If sErr <> "" Then MsgBox sErr, vbExclamation, "Connection Error"

        Case Com.InetConnect
            'No server status info receaved, means not connected.
            If sErr = vbNullString Then
                MsgBox "Unable to connect to " & FTPAddress & ":" & FTPServerPort & ". Make sure the server information is correct.", vbExclamation, "Connection Error"
            End If
            
            'Login or pwd wrong.
            If InStr(1, sErr, "530") Then
                MsgBox "Incorrect password or username for server " & FTPAddress & ".", vbExclamation, "Login Failure"
            End If

        Case Com.SetCurrentDirectory
            If InStr(1, sErr, "550") Then  'Directory does not exist.
                'Create the ftp directory.
                Dim DirSplitCreate() As String
                
                    'Create dir per dir.
                    DirSplitCreate = Split(FTPDirectory, "/")
                        For i = 0 To UBound(DirSplitCreate)
                            If Len(Trim(DirSplitCreate(i))) > 0 Then
                                FtpCreateDirectory hConnection, DirSplitCreate(i)
                                FtpSetCurrentDirectory hConnection, DirSplitCreate(i)
                            End If
                        Next
                
                ShowError CreateDirectory
                
                If CreateErr = 0 Then 'Kills the loop.
                    'Open the FTP directory after creation of it.
                    ShowError SetCurrentDirectory
                End If
                'No resetting of CreateErr here, else the file will be uploaded to the root folder.
            End If
            
        Case Com.CreateDirectory
                If InStr(1, sErr, "550") Then
                    MsgBox "You do not have premission to create directories on this server.", vbExclamation, "Create Directory Error"
                    CreateErr = 1 'Need to set this, else it keeps looping and error wont get away.
                End If
                If InStr(1, sErr, "257") Then DirCreated = True

        Case Com.DeleteFile
            If InStr(1, sErr, "550") Then MsgBox "You do not have premission to delete the files.", vbExclamation, "Delete File"
        
        Case Com.PutFile
            '266 means transfer ok.
            If InStr(1, sErr, "226") Then
                TransferOK = 1 'Used for final success message
            Else
                'No premission.
                If InStr(1, sErr, "550") Then
                    MsgBox "You do not have premission to put files on this server.", vbExclamation, "Upload File Error"
                End If
                TransferOK = 0 'Used for final success message
            End If
        
        Case Com.InetCloseHandle
            If sErr <> "" Then MsgBox sErr, vbExclamation, "Connection Error"
        
    End Select

End Sub

Public Function UploadFile() As Boolean
On Local Error GoTo ErrHand

If ShowProgress = True Then frmTransferState.Show
BeginTransfer = Timer
frmTransferState.lblFile = "File: " & RemoteFile

Dim Data(BUFFERSIZE - 1) As Byte
Dim lBlock As Long
    
    hFile = FtpOpenFile(hConnection, RemoteFile, GENERIC_WRITE, 0, 0)
    
    'Open file.
    Open LocalFile For Binary Access Read As #1
    
    'Set PB.
    FileSize = LOF(1)
    If FileSize > 0 Then frmTransferState.PB.Max = FileSize
    If FileSize = 0 Then TransferOK = 1
    frmTransferState.PB.Min = 0
    
    'loop trough file.
    For lBlock = 1 To FileSize \ BUFFERSIZE
        Get #1, , Data

        'Send chunk of data.
        If (InternetWriteFile(hFile, Data(0), BUFFERSIZE, Written) = 0) Then
            UploadFile = False
            Exit Function
        End If
        
        DoEvents
        Sum = Sum + BUFFERSIZE
        
        'Updates speed variables.
        If Sum / 1024 >= 1024 Then
            TotalSent = Sum / 1048576
            TotalSentString = Format(TotalSent, "########.00") & " MB"
        Else
            TotalSent = Sum / 1024
            TotalSentString = Format(TotalSent, "######0") & " KB"
        End If

        If ShowProgress = True Then TransferProgress Sum

            If CancelTransfer Then
            
                Select Case MsgBox("File transfer aborted. Do you want to delete the partly uploaded file from the server?", vbQuestion + vbYesNo)
                    Case vbYes
                        InternetCloseHandle hFile
                        FtpDeleteFile hConnection, RemoteFile
                        InternetCloseHandle hConnect
                        ShowError DeleteFile
                End Select
                
                End
            End If
    Next
    
        'get and send the last chunk of data.
        Get #1, , Data
        If (InternetWriteFile(hFile, Data(0), FileSize Mod BUFFERSIZE, Written) = 0) Then
            UploadFile = False
            Exit Function
        End If
        
        Sum = Sum + BUFFERSIZE
        
        'Updates speed variables.
        If Sum / 1024 >= 1024 Then
            'Show MB.
            TotalSent = Sum / 1048576 '1024 x 1024 = 1048576
            TotalSentString = Format(TotalSent, "########.00") & " MB"
        Else
            'Show KB.
            TotalSent = Sum / 1024
            TotalSentString = Format(TotalSent, "######0") & " KB"
        End If

        If ShowProgress = True Then TransferProgress Sum
      
    Close #1
      
        'Close file and connection.
        InternetCloseHandle hFile
        InternetCloseHandle hOpen
        UploadFile = True
        frmTransferState.PB.Value = frmTransferState.PB.Max
        frmTransferState.Caption = "100% - " & RemoteFile
        
Exit Function

ErrHand:
    If Err.Number = 55 Then
        Close #1
    End If

End Function

Public Sub TransferProgress(lCurrentBytes As Long)
Dim Percent
Dim TransferRate As Single

'Set general variables.
TransferRate = (Sum / (Timer - BeginTransfer)) / 1024
tmp = FileSize - Sum
Percent = (Sum / FileSize) * 100

    'Updates every 0.3 seconds.
    If Timer >= NewTimer Then
        'Update labels.
        frmTransferState.lblSpeed.Caption = Format(Int(Sum / (Timer - BeginTransfer)) / 1024, "####0.00") & " KB/s"
        frmTransferState.lblTotalSent.Caption = TotalSentString
        If frmTransferState.PB.Max >= Sum Then frmTransferState.PB.Value = Sum
        frmTransferState.lblTimeLeft.Caption = ConvertTime(Int(((frmTransferState.PB.Max - frmTransferState.PB.Value) / 1024) / TransferRate))
        frmTransferState.Caption = Format(Percent, "#0") & "% - " & RemoteFile
        NewTimer = Timer + 0.3
    End If
    
End Sub

'This converts TheTime to download time left, thx to the guy who made it.
Public Function ConvertTime(ByVal TheTime As Single) As String
    Dim NewTime                         As String
    Dim Sec                             As Single
    Dim Min                             As Single
    Dim H                               As Single
    
    If TheTime > 60 Then
        Sec = TheTime
        Min = Sec / 60
        Min = Int(Min)
        Sec = Sec - Min * 60
        H = Int(Min / 60)
        Min = Min - H * 60
        NewTime = H & ":" & Min & ":" & Sec
        If H < 0 Then H = 0
        If Min < 0 Then Min = 0
        If Sec < 0 Then Sec = 0
        NewTime = Format(NewTime, "HH:MM:SS")
        ConvertTime = NewTime
    End If
    
    If TheTime < 60 Then
        NewTime = "00:00:" & TheTime
        NewTime = Format(NewTime, "HH:MM:SS")
        ConvertTime = NewTime
    End If
    
End Function

Public Function DoQueue(File As String)
Dim Line As String
Dim Content As String

    Open App.Path & "\Queue.txt" For Input As #1
        Line Input #1, Line
            Content = Content & vbCrLf & Line
    Close #1
    
    Open App.Path & "\Queue.txt" For Output As #1
        Print #1, Content
    Close #1

End Function














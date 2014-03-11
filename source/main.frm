VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Begin VB.Form frmMain 
   Caption         =   "PortTalk for Advanced TCP/IP Users - By Shital"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   10170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "File Size"
      Height          =   495
      Left            =   8760
      TabIndex        =   22
      Top             =   5940
      Width           =   675
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Send File"
      Height          =   555
      Left            =   7860
      TabIndex        =   21
      Top             =   5940
      Width           =   795
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Undo File Recv"
      Height          =   495
      Left            =   6960
      TabIndex        =   20
      Top             =   6000
      Width           =   795
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   5220
      TabIndex        =   18
      Top             =   6000
      Width           =   1575
   End
   Begin VB.TextBox txtSize 
      Height          =   285
      Left            =   2880
      TabIndex        =   16
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CheckBox chkIsChunk 
      Caption         =   "Receive BIN Chunk"
      Height          =   255
      Left            =   180
      TabIndex        =   15
      Top             =   6000
      Width           =   1755
   End
   Begin VB.Frame fraBreak2 
      Height          =   135
      Left            =   0
      TabIndex        =   14
      Top             =   5760
      Visible         =   0   'False
      Width           =   7755
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Disconnect"
      Height          =   315
      Left            =   6360
      TabIndex        =   11
      Top             =   240
      Width           =   1155
   End
   Begin VB.CheckBox cbAppendCR 
      Caption         =   "Append CR+LF"
      Height          =   195
      Left            =   5040
      TabIndex        =   10
      Top             =   5040
      Value           =   1  'Checked
      Width           =   1515
   End
   Begin VB.CommandButton btnSend 
      Caption         =   "S&end"
      Height          =   375
      Left            =   6660
      TabIndex        =   9
      Top             =   5280
      Width           =   1035
   End
   Begin VB.TextBox txtUserLine 
      Height          =   375
      Left            =   180
      TabIndex        =   7
      Top             =   5280
      Width           =   6375
   End
   Begin VB.TextBox txtPortDlg 
      Height          =   3555
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Top             =   1260
      Width           =   6435
   End
   Begin VB.Frame fraBreak1 
      Height          =   135
      Left            =   -240
      TabIndex        =   6
      Top             =   720
      Width           =   7755
   End
   Begin VB.CommandButton btnConnect 
      Caption         =   "&Connect Now!"
      Height          =   315
      Left            =   4680
      TabIndex        =   5
      Top             =   240
      Width           =   1575
   End
   Begin ComCtl2.UpDown UpDown1 
      Height          =   315
      Left            =   3960
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   240
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   327681
      BuddyControl    =   "txtPort"
      BuddyDispid     =   196623
      OrigLeft        =   3720
      OrigTop         =   240
      OrigRight       =   3960
      OrigBottom      =   555
      Max             =   65000
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   7080
      Top             =   1380
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   327681
   End
   Begin VB.TextBox txtPort 
      Height          =   315
      Left            =   3300
      TabIndex        =   3
      Text            =   "23145"
      Top             =   240
      Width           =   675
   End
   Begin VB.TextBox txtServer 
      Height          =   315
      Left            =   780
      TabIndex        =   1
      Text            =   "192.168.1.160"
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lblBytesSent 
      AutoSize        =   -1  'True
      Caption         =   "(File Trasfer Progress)"
      Height          =   195
      Left            =   3660
      TabIndex        =   23
      Top             =   960
      Width           =   1530
   End
   Begin VB.Label lblFile 
      Caption         =   "File:"
      Height          =   255
      Left            =   4860
      TabIndex        =   19
      Top             =   6060
      Width           =   375
   End
   Begin VB.Label lblSize 
      Caption         =   "Size:"
      Height          =   255
      Left            =   2400
      TabIndex        =   17
      Top             =   6060
      Width           =   375
   End
   Begin VB.Label lblUserLine 
      AutoSize        =   -1  'True
      Caption         =   "Type here you want to send to server:"
      Height          =   195
      Left            =   180
      TabIndex        =   13
      Top             =   5040
      Width           =   2685
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Here is your interaction with server:"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   1020
      Width           =   2475
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Port:"
      Height          =   195
      Left            =   2940
      TabIndex        =   2
      Top             =   300
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Server:"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   300
      Width           =   510
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const nEND = 0
Const nRESUME_NEXT = 1
Const nRESUME = 2

Private lInBytes As Long
Private lOutBytes As Long
Private nFileNum As Integer

Private Sub btnSend_Click()
On Error GoTo err_Send
Dim sSendStr As String
    sSendStr = txtUserLine.Text
    If cbAppendCR.Value = vbChecked Then sSendStr = sSendStr & vbCrLf
    Winsock1.SendData sSendStr
    If cbAppendCR.Value = vbChecked Then
        DoLog "You: " & txtUserLine.Text & " <CR-LF>"
    Else
        DoLog "You: " & txtUserLine.Text
    End If
    
    txtUserLine.SelStart = 0
    txtUserLine.SelLength = Len(txtUserLine.Text)
    
Exit Sub
err_Send:
    Select Case HandleErr
    Case nEND
        Unload Me
    Case nRESUME_NEXT
        Resume Next
    Case nRESUME
        Resume
    End Select
    
    Exit Sub
End Sub

Private Sub btnConnect_Click()
On Error GoTo err_Connect

    If (Winsock1.State = sckConnected) Or (Winsock1.State = sckError) Or (Winsock1.State = sckOpen) Then
        Winsock1.Close
    End If

    Winsock1.RemoteHost = txtServer.Text
    Winsock1.RemotePort = txtPort.Text
    
    DoLog "// Trying to connect " & Winsock1.RemoteHost & ":" & Winsock1.RemotePort & " ..."
    
    Winsock1.Connect
    
    txtUserLine.SelStart = 0
    txtUserLine.SelLength = Len(txtUserLine.Text)
    txtUserLine.SetFocus
    
Exit Sub
err_Connect:
    Select Case HandleErr
    Case nEND
        Unload Me
    Case nRESUME_NEXT
        Resume Next
    Case nRESUME
        Resume
    End Select
    
    Exit Sub

End Sub


Private Sub Command1_Click()
    
    On Error Resume Next
    lInBytes = -1
    lOutBytes = 0
    
    Close #nFileNum
    
End Sub

Private Sub Command2_Click()
On Error GoTo err_Close
    Winsock1.Close
    
Exit Sub
err_Close:
    Select Case HandleErr
    Case nEND
        Unload Me
    Case nRESUME_NEXT
        Resume Next
    Case nRESUME
        Resume
    End Select
    
    Exit Sub
End Sub

Private Sub Command3_Click()
    On Error GoTo Err_SendFile
 
    Dim baFileData() As Byte
    Dim nFreeFileNum As Integer
    Dim lFileLen As Long
    
    lFileLen = FileLen(txtFileName.Text)
    
    ReDim baFileData(lFileLen)
    
    nFreeFileNum = FreeFile
    
    Open txtFileName.Text For Binary Access Read As #nFreeFileNum
    
        Get #nFreeFileNum, , baFileData()
    
    Close #nFreeFileNum
    
    lOutBytes = 0
    
    Winsock1.SendData baFileData()
    
Exit Sub
Err_SendFile:
    Select Case HandleErr
    Case nEND
        Unload Me
    Case nRESUME_NEXT
        Resume Next
    Case nRESUME
        Resume
    End Select
    
    Exit Sub
End Sub

Private Sub Command4_Click()
    On Error Resume Next
    txtSize.Text = CStr(FileLen(txtFileName.Text))
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    txtPortDlg.Width = CInt(frmMain.Width * 0.94)
    txtUserLine.Width = txtPortDlg.Width
    
    txtSize.Top = CInt(frmMain.Height - txtSize.Height - frmMain.Height * 0.07)
    chkIsChunk.Top = txtSize.Top
    lblSize.Top = txtSize.Top
    lblFile.Top = txtSize.Top
    txtFileName.Top = txtSize.Top
    Command1.Top = txtSize.Top
    Command3.Top = txtSize.Top
    Command4.Top = txtSize.Top
    
    btnSend.Left = txtPortDlg.Left + txtPortDlg.Width - btnSend.Width
    btnSend.Top = CInt(frmMain.Height - btnSend.Height - txtSize.Height - frmMain.Height * 0.07)
    txtUserLine.Top = CInt(btnSend.Top - txtUserLine.Height - frmMain.Height * 0.02)
    cbAppendCR.Top = btnSend.Top
    txtPortDlg.Height = CInt(txtUserLine.Top - lblUserLine.Height - 20 - frmMain.Height * 0.02) - txtPortDlg.Top
    fraBreak1.Width = frmMain.Width * 2
    lblUserLine.Top = txtUserLine.Top - lblUserLine.Height - 20
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Winsock1.State = sckConnected Then Winsock1.Close
End Sub

Private Sub txtPort_GotFocus()
    txtPort.SelStart = 0
    txtPort.SelLength = Len(txtPort.Text)
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then btnConnect_Click
End Sub


Private Sub txtServer_GotFocus()
    txtServer.SelStart = 0
    txtServer.SelLength = Len(txtServer.Text)
End Sub

Private Sub txtServer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then btnConnect_Click
End Sub

Private Sub txtUserLine_GotFocus()
    txtUserLine.SelStart = 0
    txtUserLine.SelLength = Len(txtUserLine.Text)
End Sub

Private Sub txtUserLine_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then btnSend_Click
End Sub

Private Sub Winsock1_Close()
    On Error Resume Next
    
    DoLog "// Connection closed!"
    Winsock1.Close
End Sub

Private Sub Winsock1_Connect()
    DoLog "// Connected!"
    lInBytes = -1
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error GoTo err_DataArrival
Dim vntDataArrived As Variant
Dim avntData() As Byte

    If chkIsChunk.Value = vbChecked Then
    
        'If file not created yet
        If lInBytes = -1 Then
            
            nFileNum = FreeFile
            
            Open txtFileName.Text For Binary Access Write As #nFileNum
            
            lInBytes = 0
                        
            DoLog "Srv: down load started"
            
        End If
            
        Winsock1.GetData avntData(), vbByte + vbArray
        
        Put #nFileNum, , avntData()
        
        lInBytes = lInBytes + bytesTotal
        
        lblBytesSent.Caption = lInBytes & " bytes arrived of total " & bytesTotal
        
        If lInBytes >= CLng(txtSize.Text) Then
            
            lInBytes = -1
            
            Close #nFileNum
            
            DoLog "Srv: download complete"
            
        End If
        
    Else
        Winsock1.GetData vntDataArrived, vbString
        DoLog "Srv: " & vntDataArrived
    End If
    
Exit Sub
err_DataArrival:
    Select Case HandleErr
    Case nEND
        Unload Me
    Case nRESUME_NEXT
        Resume Next
    Case nRESUME
        Resume
    End Select
    
    Exit Sub

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    DoLog "// Error " & Number & " : " & Description & " (" & Scode & ") " & " Source : " & Source
End Sub

Private Sub DoLog(sLogStr As String)
    txtPortDlg.Text = txtPortDlg.Text & sLogStr & vbCrLf
    txtPortDlg.SelStart = Len(txtPortDlg.Text)
End Sub

Private Function HandleErr() As Integer
    Select Case MsgBox("Error " & Err.Number & " : " & Err.Description, vbAbortRetryIgnore)
    Case vbAbort
        HandleErr = nEND
    Case vbRetry
        HandleErr = nRESUME
    Case vbIgnore
        HandleErr = nRESUME_NEXT
    End Select
End Function

Private Sub Winsock1_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
        lOutBytes = lOutBytes + bytesSent
        lblBytesSent.Caption = lOutBytes & " bytes sent out of " & txtSize.Text
End Sub

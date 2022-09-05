VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Autopoll 1.03"
   ClientHeight    =   10440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16155
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10440
   ScaleWidth      =   16155
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picVotes 
      Height          =   1575
      Left            =   1680
      ScaleHeight     =   1515
      ScaleWidth      =   5595
      TabIndex        =   35
      Top             =   2280
      Visible         =   0   'False
      Width           =   5655
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   0
         TabIndex        =   37
         Top             =   960
         Width           =   5535
      End
      Begin VB.Label lblVotes 
         Alignment       =   2  'Center
         Caption         =   "Votes casted 0/0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   0
         TabIndex        =   36
         Top             =   360
         Width           =   5535
      End
   End
   Begin VB.Frame frmReady 
      Caption         =   "I am ready!. Auto vote now......."
      Enabled         =   0   'False
      Height          =   4935
      Left            =   120
      TabIndex        =   27
      Top             =   5400
      Width           =   8775
      Begin VB.CheckBox chkUseHistory 
         Caption         =   "Use history to load th page for faster operation "
         Height          =   255
         Left            =   4560
         TabIndex        =   40
         Top             =   3720
         Width           =   3855
      End
      Begin VB.CheckBox chkReset 
         Caption         =   "Reset IP after each vote"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Frame fraIPReset 
         Height          =   2175
         Left            =   120
         TabIndex        =   32
         Top             =   1320
         Width           =   8535
         Begin VB.TextBox txtRenewTime 
            Height          =   375
            Left            =   2400
            TabIndex        =   21
            Text            =   "120"
            Top             =   1440
            Width           =   975
         End
         Begin VB.CheckBox chkWaitTerminate 
            Caption         =   "Wait for application to terminate"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   960
            Width           =   2535
         End
         Begin VB.TextBox txtAppPath 
            Height          =   375
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Width           =   7695
         End
         Begin VB.CommandButton cmdPick 
            Caption         =   "..."
            Height          =   375
            Left            =   7920
            TabIndex        =   19
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label11 
            Caption         =   "secs"
            Height          =   375
            Left            =   3480
            TabIndex        =   34
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "Wait time for IP renewal (secs)"
            Height          =   375
            Left            =   240
            TabIndex        =   33
            Top             =   1560
            Width           =   2295
         End
      End
      Begin VB.TextBox txtWaitTime 
         Height          =   375
         Left            =   2520
         TabIndex        =   22
         Text            =   "60"
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton cmdLastStep 
         Caption         =   "Autopoll Now"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2400
         TabIndex        =   23
         Top             =   4200
         Width           =   3015
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   7800
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.ComboBox cboVoteType 
         Height          =   315
         ItemData        =   "Form1.frx":1B9A
         Left            =   2520
         List            =   "Form1.frx":1B9C
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtTotalVotes 
         Height          =   375
         Left            =   2520
         TabIndex        =   15
         Text            =   "10"
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "secs"
         Height          =   375
         Left            =   3600
         TabIndex        =   31
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Wait time between votes"
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Top             =   3720
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "Option"
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "No of votes to cast"
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   480
         Width           =   2295
      End
   End
   Begin VB.Frame frmTest 
      Caption         =   "Test a vote"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   120
      TabIndex        =   26
      Top             =   4320
      Width           =   8775
      Begin VB.CommandButton Command1 
         Caption         =   "3. Vote now && Test"
         Height          =   375
         Left            =   4200
         TabIndex        =   14
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtButton 
         Height          =   375
         Left            =   2520
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "If the voting did not work and if the page was navigated to another webpage, please load the page again and try"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   720
         Visible         =   0   'False
         Width           =   8295
      End
      Begin VB.Label Label4 
         Caption         =   "Vote Button Index"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdSource 
      Caption         =   "View Source"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin VB.Frame frmProp 
      Caption         =   "Tune"
      Enabled         =   0   'False
      Height          =   2775
      Left            =   120
      TabIndex        =   24
      Top             =   1440
      Width           =   8775
      Begin VB.ComboBox cboFormIndex 
         Height          =   315
         Left            =   2520
         TabIndex        =   38
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton cmdAutotune 
         Caption         =   "&Auto Tune!"
         Height          =   375
         Left            =   600
         TabIndex        =   4
         Top             =   0
         Width           =   1815
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "2. Fill && Verify"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5040
         TabIndex        =   10
         Top             =   720
         Width           =   1575
      End
      Begin VB.ListBox lstOptions 
         Height          =   1230
         Left            =   5040
         TabIndex        =   11
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox txtOptionStart 
         Height          =   375
         Left            =   2520
         TabIndex        =   9
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtOptionCount 
         Height          =   375
         Left            =   2520
         TabIndex        =   7
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "No of options"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Poll Form Index"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Starting option index"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   2160
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "1. Go"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8160
      TabIndex        =   2
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox txtURL 
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   240
      Width           =   5415
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   10215
      Left            =   9000
      TabIndex        =   25
      Top             =   120
      Width           =   7095
      ExtentX         =   12515
      ExtentY         =   18018
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label1 
      Caption         =   "Website Address"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nFormNo As Integer
Dim nOptionStart As Integer, noptionCount As Integer
Dim nCmdNo As Integer
Private Sub cboFormIndex_Click()
    cboFormIndex_Validate False
End Sub

Private Sub cboFormIndex_Validate(Cancel As Boolean)
    Dim i As Integer
    Dim nRadioCount As Integer
    Dim lFormCount As Integer
    Dim sObjType As String
    
    nRadioCount = 0
    txtOptionStart.Text = ""
    txtOptionCount.Text = ""
    nFormNo = Val("0" + cboFormIndex.Text)
    cboFormIndex.Text = nFormNo
    
    On Error GoTo ERR_HAND:
    
    lFormCount = WebBrowser1.Document.Forms(nFormNo).length
    
    
    For i = 0 To lFormCount - 1
        On Error Resume Next
        sObjType = WebBrowser1.Document.Forms(nFormNo).elements(i).Type
        If Err Then
            sObjType = ""
        End If
        On Error GoTo ERR_HAND:
        Debug.Print sObjType
        If UCase(sObjType) = "RADIO" Then
            If txtOptionStart = "" Then
                txtOptionStart.Text = i
            End If
            nRadioCount = nRadioCount + 1
        End If
        If (UCase(sObjType) = "SUBMIT" Or _
            UCase(sObjType) = "BUTTON") And _
            txtOptionStart <> "" Then
                txtButton.Text = i
        End If
    Next i
    txtOptionCount.Text = nRadioCount
    EnableDisableFill
    Exit Sub
ERR_HAND:
    MsgBox Err.Description + vbCrLf + "Invalid Form Index. It must be between 0 and " + Str(WebBrowser1.Document.Forms.length - 1), vbCritical
    Cancel = True
    Exit Sub
End Sub

Sub EnableDisableFill()
    If Trim(txtOptionCount) <> "" And Trim(txtOptionStart) <> "" And Trim(cboFormIndex.Text) <> "" Then
        cmdTest.Enabled = True
    Else
        cmdTest.Enabled = False
    End If
End Sub
Private Sub chkReset_Click()
    If chkReset.Value = False Then
        fraIPReset.Enabled = False
    Else
        fraIPReset.Enabled = True
    End If
End Sub

Private Sub cmdAutotune_Click()
    Dim i As Integer
    cboFormIndex.Clear
    For i = 0 To WebBrowser1.Document.Forms.length - 1
        cboFormIndex.AddItem i
    Next i
    For i = 0 To WebBrowser1.Document.Forms.length - 1
        cboFormIndex.ListIndex = i
        If txtOptionCount.Text <> "0" And txtOptionCount.Text <> "" Then
            Exit For
        End If
    Next
End Sub

Private Sub cmdLastStep_Click()
    
    lblVotes.Caption = "Votes casted 0/" + txtTotalVotes
    lblStatus.Caption = ""
    picVotes.Visible = True
    
    Dim IPRenewTime As Integer
    Dim hWnd As Long
    Dim nWaitTime As Integer
    Dim nTotalVotes As Integer
    Dim i As Integer
    
    frmProp.Enabled = False
    frmTest.Enabled = False
    frmReady.Enabled = False
    
    nTotalVotes = 0
    
    For i = 1 To CInt("0" + txtTotalVotes.Text)
    
        'Start the application
        If chkReset.Value = 1 Then
            lblStatus.Caption = "Resetting IP..."
            lblStatus.Refresh
            If chkWaitTerminate.Value = 1 Then
                ShellExecuteWait hWnd, "open", txtAppPath, "", "", vbNormalFocus
            Else
                Shell txtAppPath, vbMinimizedFocus
            End If
            IPRenewTime = CInt("0" + txtRenewTime.Text)
            lblStatus.Caption = "Waiting for new IP..."
            lblStatus.Refresh
            Pause IPRenewTime
        End If
        
        'Process vote now
        lblStatus.Caption = "Loading webpage..."
        lblStatus.Refresh
        
        If chkUseHistory.Value = 1 Then
            WebBrowser1.GoBack
            Do While Trim(LCase(WebBrowser1.LocationURL)) <> Trim(LCase(txtURL.Text))
                DoEvents
            Loop
        Else
            WebBrowser1.Navigate txtURL
        End If
        
        Do While WebBrowser1.ReadyState <> READYSTATE_COMPLETE
            DoEvents
        Loop
        
        Do While WebBrowser1.Busy = True
            DoEvents
        Loop
        
        
        
        'MsgBox WebBrowser1.LocationURL
        
        Dim nOptionSelected As Integer
        
        If cboVoteType.ListIndex = 1 Then
            nOptionSelected = Random(0, lstOptions.ListCount)
            lstOptions.ListIndex = nOptionSelected
        End If
        Debug.Print nOptionSelected
        WebBrowser1.Document.Forms(nFormNo).Item(lstOptions.ListIndex + nOptionStart).Checked = True
        
        Command1_Click
        nTotalVotes = nTotalVotes + 1
        lblVotes.Caption = "Votes casted " + Trim(Str(nTotalVotes)) + "/" + txtTotalVotes
        lblVotes.Refresh
        lblStatus.Caption = "Voting done. Waiting..."
        lblStatus.Refresh
        nWaitTime = CInt("0" + txtWaitTime.Text)
        Pause nWaitTime
    Next
    frmProp.Enabled = True
    frmTest.Enabled = True
    frmReady.Enabled = True
    MsgBox "Voting finsihed", vbInformation
End Sub

Private Sub cmdLoad_Click()
    Dim i As Integer
    On Error GoTo ERR_HAND
    WebBrowser1.Navigate txtURL
    Do While WebBrowser1.ReadyState <> READYSTATE_COMPLETE
        DoEvents
    Loop
    frmProp.Enabled = True
    cmdSource.Enabled = True
    MsgBox "Site loaded", vbInformation
    Exit Sub
ERR_HAND:
    MsgBox "Error loading website. Please try again", vbCritical
    frmProp.Enabled = False
    frmTest.Enabled = False
    frmReady.Enabled = False
    cmdSource.Enabled = False
    Exit Sub
End Sub

Private Sub cmdPick_Click()
    CommonDialog1.Filter = "Executables(*.exe;*.bat;*.cmd;*.com)|*.exe;*.bat;*.cmd;*.com"
    CommonDialog1.CancelError = False
    CommonDialog1.ShowOpen
    txtAppPath.Text = CommonDialog1.FileName
End Sub

Private Sub cmdSource_Click()
    On Error GoTo ErrHand
    WriteStrToFile "c:\temp.html", WebBrowser1.Document.body.innerHTML
    Shell "notepad c:\temp.html", vbNormalFocus
    Exit Sub
ErrHand:
    MsgBox "Unable to view the source", vbCritical
    Exit Sub
End Sub

Private Sub cmdTest_Click()
    Dim i As Integer
    frmTest.Enabled = False
    nOptionStart = Val("0" + txtOptionStart.Text)
    noptionCount = Val("0" + txtOptionCount.Text)
    lstOptions.Clear
    On Error GoTo ERR_HAND
    For i = nOptionStart To nOptionStart + noptionCount - 1
        lstOptions.AddItem WebBrowser1.Document.Forms(nFormNo).Item(i).Name
    Next
    If lstOptions.ListCount > 0 Then
        frmTest.Enabled = True
    End If
    Exit Sub
ERR_HAND:
    MsgBox Err.Description + vbCrLf + "Invalid tuning values. Try auto tune", vbCritical
    Exit Sub
End Sub

Private Sub Command1_Click()
    On Error GoTo ERR_HAND
    nCmdNo = Val("0" + txtButton.Text)
    txtButton.Text = nCmdNo
    WebBrowser1.Document.Forms(nFormNo).Item(nCmdNo).Click
    frmReady.Enabled = True
    Label12.Visible = True
    Exit Sub
ERR_HAND:
    MsgBox Err.Description + vbCrLf + "Invalid button Index", vbCritical
    Exit Sub
End Sub


Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
    
    WebBrowser1.Navigate "about:blank"
    WebBrowser1.Silent = True
    cboVoteType.AddItem "Fixed option (Selected Above)"
    cboVoteType.AddItem "Random option"
    cboVoteType.ListIndex = 0
End Sub

Private Sub lstOptions_Click()
    On Error Resume Next
    WebBrowser1.Document.Forms(nFormNo).Item(lstOptions.ListIndex + nOptionStart).Checked = True
    If Err Then
        MsgBox Err.Description + vbCrLf + "Not a radio button. Retune", vbCritical
    End If
    On Error GoTo 0
End Sub

Private Sub txtFormNo_Validate(Cancel As Boolean)

End Sub

Private Sub txtOptionCount_Change()
    EnableDisableFill
End Sub

Private Sub txtOptionStart_Change()
    EnableDisableFill
End Sub

Private Sub txtURL_Change()
    If txtURL.Text <> "" Then
        cmdLoad.Enabled = True
    Else
        cmdLoad.Enabled = False
    End If
End Sub

Sub WriteStrToFile(ByVal sFileName As String, ByVal iStr As String)
    Dim nFreeFile As Integer
    nFreeFile = FreeFile
    Open sFileName For Output As #nFreeFile
    Print #nFreeFile, iStr
    Close #nFreeFile
End Sub

Public Sub Pause(iSecs As Integer)

    
    Dim i As Integer
    

    For i = 1 To iSecs * 10
        Sleep 100


        DoEvents
    Next

End Sub
Function Random(Lowerbound As Long, Upperbound As Long)
    Randomize Timer
    Random = Int(Rnd * Upperbound) + Lowerbound
End Function

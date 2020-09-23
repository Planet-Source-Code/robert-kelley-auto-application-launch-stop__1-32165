VERSION 5.00
Object = "{E30A83A7-F955-11D1-9AA0-400060046636}#1.0#0"; "SysTray.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AutoLaunch"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   4485
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   2715
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   4789
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   0
      TabCaption(0)   =   "Application"
      TabPicture(0)   =   "Form1.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdGo"
      Tab(0).Control(1)=   "txtFile"
      Tab(0).Control(2)=   "cmdApp"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Schedule"
      TabPicture(1)   =   "Form1.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkWeek(6)"
      Tab(1).Control(1)=   "chkWeek(5)"
      Tab(1).Control(2)=   "chkWeek(4)"
      Tab(1).Control(3)=   "chkWeek(3)"
      Tab(1).Control(4)=   "chkWeek(2)"
      Tab(1).Control(5)=   "chkWeek(1)"
      Tab(1).Control(6)=   "chkWeek(0)"
      Tab(1).Control(7)=   "optSpecific"
      Tab(1).Control(8)=   "optEvery"
      Tab(1).Control(9)=   "txtStart"
      Tab(1).Control(10)=   "cboMins"
      Tab(1).Control(11)=   "Line1(1)"
      Tab(1).Control(12)=   "Line1(0)"
      Tab(1).Control(13)=   "Label3"
      Tab(1).Control(14)=   "Label1(0)"
      Tab(1).Control(15)=   "Label2"
      Tab(1).ControlCount=   16
      TabCaption(2)   =   "Help"
      TabPicture(2)   =   "Form1.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtHelp"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "About"
      TabPicture(3)   =   "Form1.frx":0060
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label4"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Line2"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "lblVersion"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "lblComments"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Picture1(2)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Picture1(0)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Timer2"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Picture1(1)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).ControlCount=   8
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         Height          =   480
         Index           =   1
         Left            =   825
         Picture         =   "Form1.frx":007C
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   25
         Top             =   1800
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Timer Timer2 
         Interval        =   100
         Left            =   5250
         Top             =   1830
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   0
         Left            =   360
         Negotiate       =   -1  'True
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   24
         Tag             =   "0"
         Top             =   1800
         Width           =   480
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Index           =   2
         Left            =   1320
         Picture         =   "Form1.frx":0CBE
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   256
         TabIndex        =   23
         Top             =   1800
         Visible         =   0   'False
         Width           =   3840
      End
      Begin VB.TextBox txtHelp 
         BackColor       =   &H00C0C0C0&
         Height          =   2175
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   420
         Width           =   4455
      End
      Begin VB.CommandButton cmdGo 
         Caption         =   "Go"
         Default         =   -1  'True
         Height          =   315
         Left            =   -71520
         Picture         =   "Form1.frx":6D00
         TabIndex        =   18
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txtFile 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -74820
         TabIndex        =   17
         Top             =   480
         Width           =   3795
      End
      Begin VB.CommandButton cmdApp 
         Caption         =   "Choose Application"
         Height          =   285
         Left            =   -74820
         TabIndex        =   16
         Top             =   780
         Width           =   1995
      End
      Begin VB.CheckBox chkWeek 
         Caption         =   "Saturday"
         Height          =   195
         Index           =   6
         Left            =   -71640
         TabIndex        =   15
         Top             =   1560
         Width           =   1035
      End
      Begin VB.CheckBox chkWeek 
         Caption         =   "Friday"
         Height          =   195
         Index           =   5
         Left            =   -71640
         TabIndex        =   14
         Top             =   1320
         Width           =   1035
      End
      Begin VB.CheckBox chkWeek 
         Caption         =   "Thursday"
         Height          =   195
         Index           =   4
         Left            =   -73020
         TabIndex        =   13
         Top             =   2280
         Width           =   1035
      End
      Begin VB.CheckBox chkWeek 
         Caption         =   "Wednesday"
         Height          =   195
         Index           =   3
         Left            =   -73020
         TabIndex        =   12
         Top             =   2040
         Width           =   1155
      End
      Begin VB.CheckBox chkWeek 
         Caption         =   "Tuesday"
         Height          =   195
         Index           =   2
         Left            =   -73020
         TabIndex        =   11
         Top             =   1800
         Width           =   1035
      End
      Begin VB.CheckBox chkWeek 
         Caption         =   "Monday"
         Height          =   195
         Index           =   1
         Left            =   -73020
         TabIndex        =   10
         Top             =   1560
         Width           =   1035
      End
      Begin VB.CheckBox chkWeek 
         Caption         =   "Sunday"
         Height          =   195
         Index           =   0
         Left            =   -73020
         TabIndex        =   9
         Top             =   1320
         Width           =   1035
      End
      Begin VB.OptionButton optSpecific 
         Caption         =   "On Specific Days"
         Height          =   255
         Left            =   -73020
         TabIndex        =   8
         Top             =   1020
         Value           =   -1  'True
         Width           =   2355
      End
      Begin VB.OptionButton optEvery 
         Caption         =   "Every Day"
         Height          =   195
         Left            =   -74580
         TabIndex        =   7
         Top             =   1020
         Width           =   1815
      End
      Begin VB.TextBox txtStart 
         Height          =   315
         Left            =   -73980
         TabIndex        =   3
         Text            =   "12:00"
         Top             =   420
         Width           =   555
      End
      Begin VB.ComboBox cboMins 
         Height          =   315
         Left            =   -72240
         TabIndex        =   2
         Text            =   "1"
         Top             =   420
         Width           =   555
      End
      Begin VB.Label lblComments 
         Caption         =   "Comments"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   1200
         Width           =   4275
      End
      Begin VB.Label lblVersion 
         Caption         =   "Version :"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   900
         Width           =   1095
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   4440
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Label Label4 
         Caption         =   "AutoLaunch"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   480
         Width           =   1455
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   -74880
         X2              =   -70560
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         Index           =   0
         X1              =   -74880
         X2              =   -70560
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label3 
         Caption         =   "Minute(s)"
         Height          =   195
         Left            =   -71580
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Start Time :"
         Height          =   195
         Index           =   0
         Left            =   -74820
         TabIndex        =   5
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Duration :"
         Height          =   255
         Left            =   -72960
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
   End
   Begin ComctlLib.StatusBar StatBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2715
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   6174
            MinWidth        =   6174
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "4:16 PM"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   2460
      Top             =   2160
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   120
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin SysTrayCtl.cSysTray SysTray 
      Left            =   240
      Top             =   1800
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   0   'False
      TrayIcon        =   "Form1.frx":75CA
      TrayTip         =   "DataMate"
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   0
      Top             =   0
      Width           =   4665
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStart 
         Caption         =   "Start Application"
      End
      Begin VB.Menu mnuKill 
         Caption         =   "Kill Application"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Properties"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strApplication As String    ' String to Hold Path of Application
Dim strStart As String          ' String holding Start Time
Dim RUNNING As Boolean          ' Boolean active during Applications Run

 
Private Sub chkWeek_Click(Index As Integer)
    Weekdays(Index) = chkWeek(Index).Value
End Sub

Private Sub cmdApp_Click()
Dim res As Long                 ' Return Value
Dim Buffer As String * 255      ' Empty String to Hold New Short Path

' SEt Dialog Boxes Properties and Open
    cd1.DialogTitle = "Choose Application"
    cd1.ShowOpen
    
    On Error Resume Next
    strApplication = cd1.FileName
        
' Make Short Path Name from App Chosen
        res = GetShortPathName(strApplication, Buffer, Len(Buffer))
        
' If no error then set Local String to New Short path
        If Err.LastDllError <> 0 Then
            Err.Raise vbObjectError + 1234, "ShortFileName", "ShortFileName: Error calling GetShortPathName: "
        Else
            strApplication = Left(Buffer, res)
        End If

' Set TextBox Value to Path, update status bar
        txtFile.Text = strApplication
        StatBar.Panels(1).Text = strApplication
End Sub

Private Sub cmdGo_Click()
' Minimize the Form, which calls Form_Resize
    Me.WindowState = vbMinimized
End Sub


Private Sub Command1_Click()
Dim I As Integer
Dim tString As String

    For I = 0 To 6
      tString = tString & frmMain.chkWeek(I).Caption & ":" & Weekdays(I)
    Next I
    
    MsgBox tString
End Sub

Private Sub Form_Load()
Dim I As Integer
Timer2_Timer
' Set The labels on the about tab
    SetAbout
    
' Load Icon from Resource File
 '   Me.Icon = LoadResPicture(101, vbResIcon)

' Fill the Duration List with 60 minutes
    For I = 1 To 120
        cboMins.AddItem I
    Next I
    
' Set StatusBar Caption Initially
    StatBar.Panels(1).Text = "Choose and Application"
    
' Call Function to Fill The HelpBox with Text
    FillHelp txtHelp
    
' Are we Running?
    RUNNING = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    PostMessage Me.hwnd, WM_CLOSE, 0&, 0&
    Unload Me
    End
End Sub

Private Sub Form_Resize()
' If the Form is minimized, hide and put into taskbar, enable Timer
    If Me.WindowState = vbMinimized Then
        SysTray.InTray = True
        Me.Visible = False
        strStart = txtStart.Text
        Timer1.Enabled = True
    Else
        SysTray.InTray = False
        Me.Visible = True
        Timer1.Enabled = False
    End If
End Sub

Private Sub mnuExit_Click()
    Unload Me
    End
End Sub

Private Sub mnuKill_Click()
    KillProcess
End Sub

Private Sub mnuOpen_Click()
' make State Normal again, calling Form_Resize
    Me.WindowState = vbNormal
    Me.Visible = True
    
' Set Z Order so we are on top
    SetForegroundWindow Me.hwnd
End Sub

Private Sub mnuStart_Click()
    StartProcess strApplication
End Sub

Private Sub optEvery_Click()
Dim I As Integer

' Loop Through the Check Array and Disable
    For I = 0 To 6
        chkWeek(I).Enabled = False
    Next I
End Sub

Private Sub optSpecific_Click()
Dim I As Integer

' Loop Through the Check array and enable
    If optSpecific.Value = True Then
        For I = 0 To 6
            chkWeek(I).Enabled = True
        Next I
    End If
        
End Sub

Private Sub SysTray_MouseDblClick(Button As Integer, Id As Long)
    mnuOpen_Click
End Sub

Private Sub SysTray_MouseUp(Button As Integer, Id As Long)
'------------------------------------------------------------
    ' SetForegroundWindow and PostMessage (WM_USER) must wrap all popup menu's _
      in order to work correctely with the Notification Icons... _
      (* see KB article Q1357888 for more info *)
    SetForegroundWindow Me.hwnd                     ' Set current window as ForegroundWindow
    
    Select Case Button                              ' Track mouse clicks...
    Case vbRightButton
        Me.PopupMenu mnuMain, vbPopupMenuRightButton  ' Popup memu...
    Case vbLeftButton
        
    End Select
    
    PostMessage Me.hwnd, WM_USER, 0&, 0&
End Sub



Private Sub Timer1_Timer()
Dim MinStartDiff As Integer     ' Int to Hold # Mins away from Starting
Dim MinStopDiff As Integer      ' Int to Hold # Mins away from Stoppping
Dim EndTime As String           ' Str That holds Stopping Time (Calculated with DateAdd)

' Check If We Run Specific Day, IF so see if we run ,false then just exit the sub
    If optSpecific.Value = True Then
        If CheckWeekDay = False Then Exit Sub
    End If

' Calculate When to Stop App, add cbomins.text to the start time
    EndTime = DateAdd("n", cboMins.Text, strStart)
    
' Set Vlaues for Minutes away from Starting and Stopping
    MinStartDiff = DateDiff("N", Format(Time, "Long Time"), Format(strStart, "Long Time"))
    MinStopDiff = DateDiff("N", Format(Time, "Long Time"), Format(EndTime, "Long Time"))
    
' Check if it's time to Start the App. I we started the app (RUNNING) then check to stop
    If RUNNING = False Then
        If MinStartDiff = 0 Then
            RUNNING = True
            StartProcess strApplication
        End If
    ElseIf RUNNING = True Then
        If MinStopDiff = 0 Then
            RUNNING = False
            KillProcess
        End If
    End If
End Sub





Private Sub txtStart_LostFocus()
' Check to make sure Time is a valid time
    CheckTime txtStart.Text
End Sub

Private Function SetAbout()
    lblVersion.Caption = "Version : " & App.Major & "." & App.Minor & "." & App.Revision
    lblComments.Caption = "Comments : " & App.Comments
    'lblCopy.Caption = "Copy Right : " & App.LegalCopyright
    'lblCompany.Caption = "Company : " & App.CompanyName
End Function
Private Sub Timer2_Timer()
With Picture1
.Item(0).Cls
.Item(0).PaintPicture .Item(1).Picture, 0, 0, , , , , , , vbSrcPaint
.Item(0).PaintPicture .Item(2).Picture, 0, 0, 32, 32, .Item(0).Tag * 32, 0, 32, 32, vbSrcAnd
.Item(0).Tag = .Item(0).Tag + 1
If .Item(0).Tag > 7 Then .Item(0).Tag = 0
End With
End Sub

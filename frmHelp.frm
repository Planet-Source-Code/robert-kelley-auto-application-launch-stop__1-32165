VERSION 5.00
Begin VB.Form frmHelp 
   Caption         =   "Help/About"
   ClientHeight    =   2085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   ScaleHeight     =   2085
   ScaleWidth      =   4845
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   3180
      TabIndex        =   1
      Top             =   1740
      Width           =   1575
   End
   Begin VB.TextBox txtHelp 
      BackColor       =   &H00C0C0C0&
      Height          =   1635
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   60
      Width           =   4695
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFile As String

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmMain.Icon
    
    Dim HelpText As String
    Open App.Path & "\help.txt" For Input As #1
        Do While Not EOF(1)
            Line Input #1, strFile
            HelpText = HelpText & strFile & vbCrLf
        Loop
    Close #1
        
    txtHelp.Text = HelpText
End Sub



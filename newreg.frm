VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Reg 
   Caption         =   "Registry Hacker"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "newreg.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3765
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Apply 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   255
      Left            =   2520
      TabIndex        =   0
      Top             =   3120
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2775
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4895
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Windows Registration"
      TabPicture(0)   =   "newreg.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "System"
      TabPicture(1)   =   "newreg.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "About"
      TabPicture(2)   =   "newreg.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Line1(1)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblTitle"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Line1(0)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "lblVersion"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "war"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "picIcon"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cmdOK"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "cmdSysInfo"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Text9"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).ControlCount=   9
      Begin VB.TextBox Text9 
         BackColor       =   &H80000018&
         Height          =   1335
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   22
         Text            =   "newreg.frx":0496
         Top             =   1320
         Width           =   4215
      End
      Begin VB.Frame Frame4 
         Caption         =   "Installation Details"
         Height          =   975
         Left            =   120
         TabIndex        =   19
         Top             =   1560
         Width           =   4215
         Begin VB.TextBox Text8 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
            Height          =   285
            Left            =   120
            MaxLength       =   400
            TabIndex        =   21
            ToolTipText     =   "Installation Path"
            Top             =   600
            Width           =   3975
         End
         Begin VB.TextBox Text7 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
            Height          =   285
            Left            =   120
            MaxLength       =   400
            TabIndex        =   20
            ToolTipText     =   "Installation Key"
            Top             =   240
            Width           =   3975
         End
      End
      Begin VB.CommandButton cmdSysInfo 
         Caption         =   "&System Info..."
         Height          =   345
         Left            =   -70830
         TabIndex        =   13
         Top             =   3315
         Width           =   1245
      End
      Begin VB.CommandButton cmdOK 
         Cancel          =   -1  'True
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   345
         Left            =   -70845
         TabIndex        =   12
         Top             =   2865
         Width           =   1260
      End
      Begin VB.PictureBox picIcon 
         AutoSize        =   -1  'True
         ClipControls    =   0   'False
         Height          =   540
         Left            =   -74850
         Picture         =   "newreg.frx":05CF
         ScaleHeight     =   337.12
         ScaleMode       =   0  'User
         ScaleWidth      =   337.12
         TabIndex        =   11
         Top             =   480
         Width           =   540
      End
      Begin VB.Frame Frame1 
         Caption         =   "Windows Registration Info"
         Height          =   1095
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   4215
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
            Height          =   285
            Left            =   120
            MaxLength       =   400
            TabIndex        =   10
            ToolTipText     =   "Change Your Name"
            Top             =   240
            Width           =   3975
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
            Height          =   285
            Left            =   120
            MaxLength       =   400
            TabIndex        =   9
            ToolTipText     =   "Change your Business/Organization"
            Top             =   600
            Width           =   3975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Recycle Bin"
         Height          =   1095
         Left            =   -74880
         TabIndex        =   5
         Top             =   360
         Width           =   4215
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
            Height          =   285
            Left            =   120
            MaxLength       =   400
            TabIndex        =   7
            ToolTipText     =   "Recycle Bin Name"
            Top             =   240
            Width           =   3975
         End
         Begin VB.TextBox Text4 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
            Height          =   285
            Left            =   120
            MaxLength       =   400
            TabIndex        =   6
            ToolTipText     =   "Recycle Bin Description"
            Top             =   600
            Width           =   3975
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Contol Panel"
         Height          =   1095
         Left            =   -74880
         TabIndex        =   4
         Top             =   1560
         Width           =   4215
         Begin VB.TextBox Text6 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
            Height          =   285
            Left            =   120
            MaxLength       =   400
            TabIndex        =   18
            ToolTipText     =   "Control Panel Description"
            Top             =   600
            Width           =   3975
         End
         Begin VB.TextBox Text5 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
            Height          =   285
            Left            =   120
            MaxLength       =   400
            TabIndex        =   17
            ToolTipText     =   "Control Panel Name"
            Top             =   240
            Width           =   3975
         End
      End
      Begin VB.Label war 
         Caption         =   "Warning: ..."
         ForeColor       =   &H00000000&
         Height          =   825
         Left            =   -74835
         TabIndex        =   16
         Top             =   2865
         Width           =   3870
      End
      Begin VB.Label lblVersion 
         Caption         =   "Version"
         Height          =   225
         Left            =   -74160
         TabIndex        =   15
         Top             =   840
         Width           =   2085
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   -74985
         X2              =   -69436
         Y1              =   2820
         Y2              =   2820
      End
      Begin VB.Label lblTitle 
         Caption         =   "Application Title"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   -74160
         TabIndex        =   14
         Top             =   480
         Width           =   2805
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         Index           =   1
         X1              =   -75000
         X2              =   -69436
         Y1              =   2805
         Y2              =   2805
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Copyright ©2001 By John McKee and Tretyakov Konstantin"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   4335
   End
End
Attribute VB_Name = "Reg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright John McKee (johnpub@yahoo.com)
'You may use this code for free, if you give me some credit
'At least remember, thet it is not fair to put your name on what you didn't do

'And I would surely appreciate, if you mail me the program (or link to it)
'you created, using this code, (or if you somehow modified this one)

Private Sub Apply_Click()
'When clicked it will re-write back to the registry
SetStringKey HKEY_LOCAL_MACHINE, MainRoot & RecycleBin, , Text1.Text
SetStringKey HKEY_LOCAL_MACHINE, MainRoot & RecycleBin, "InfoTip", Text4.Text
SetStringKey HKEY_LOCAL_MACHINE, WinInfo, "RegisteredOwner", Text2.Text
SetStringKey HKEY_LOCAL_MACHINE, WinInfo, "RegisteredOrganization", Text3.Text
SetStringKey HKEY_LOCAL_MACHINE, MainRoot & ControlPanel, , Text5.Text
SetStringKey HKEY_LOCAL_MACHINE, MainRoot & ControlPanel, "InfoTip", Text6.Text
SetStringKey HKEY_LOCAL_MACHINE, Key, "ProductKey", Text7.Text
SetStringKey HKEY_LOCAL_MACHINE, Setup, "SourcePath", Text8.Text
End Sub

Private Sub Command2_Click()
'Ends the program
End
End Sub

Private Sub Form_Load()
    'Gets all the required information
    txtOwner = GetStringKey(HKEY_LOCAL_MACHINE, WinInfo, "RegisteredOwner")
    txtOrg = GetStringKey(HKEY_LOCAL_MACHINE, WinInfo, "RegisteredOrganization")
    txtRecycle = GetStringKey(HKEY_LOCAL_MACHINE, MainRoot & RecycleBin)
    txtRecycleTip = GetStringKey(HKEY_LOCAL_MACHINE, MainRoot & RecycleBin, "InfoTip")
    txtControlPanel = GetStringKey(HKEY_LOCAL_MACHINE, MainRoot & ControlPanel)
    txtControlTip = GetStringKey(HKEY_LOCAL_MACHINE, MainRoot & ControlPanel, "InfoTip")
    txtInstallKey = GetStringKey(HKEY_LOCAL_MACHINE, Key, "ProductKey")
    txtInstallPath = GetStringKey(HKEY_LOCAL_MACHINE, Setup, "SourcePath")

    'Puts the required information to a text field
    Text1.Text = txtRecycle
    Text2.Text = txtOwner
    Text3.Text = txtOrg
    Text4.Text = txtRecycleTip
    Text5.Text = txtControlPanel
    Text6.Text = txtControlTip
    Text7.Text = txtInstallKey
    Text8.Text = txtInstallPath
    
    'The about text fields
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    
End Sub


'All of these are used when they click on a text field
Private Sub Text1_Click()
Apply.Enabled = True
End Sub

Private Sub Text2_Click()
Apply.Enabled = True
End Sub

Private Sub Text3_Click()
Apply.Enabled = True
End Sub

Private Sub Text4_Click()
Apply.Enabled = True
End Sub

Private Sub Text5_Click()
Apply.Enabled = True
End Sub

Private Sub Text6_Click()
Apply.Enabled = True
End Sub

Private Sub Text7_Click()
Apply.Enabled = True
End Sub

Private Sub Text8_Click()
Apply.Enabled = True
End Sub




